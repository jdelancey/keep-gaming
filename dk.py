#!/usr/bin/python3 -u

import argparse
from bs4 import BeautifulSoup as bs
import datetime
from pymongo import MongoClient
import requests
from typing import List

import event
import google_sheets_utils as gsu
import kenpom

DK_STR_BASE_URL = 'https://sportsbook.draftkings.com'
DK_STR_EVENTS_URL = f'{DK_STR_BASE_URL}/event'

DK_STR_GAME_TABLE_TAG = 'tbody'
DK_STR_GAME_TABLE_CLASS = 'sportsbook-table__body'

DK_STR_GAME_TABLE_ROW_TAG = 'tr'
DK_STR_GAME_TABLE_ROW_CLASS = '' # unused?

DK_STR_TEAMS_TAG = 'div'
DK_STR_TEAMS_CLASS = 'event-cell__name-text'

DK_STR_GAME_TABLE_ROW_COLUMN_TAG = 'div'
DK_STR_GAME_TABLE_ROW_COLUMN_CLASS = 'sportsbook-outcome-body-wrapper'
DK_STR_GAME_TABLE_EMPTY_CELL = 'sportsbook-empty-cell body'

DK_STR_GAME_TABLE_SPREAD_TAG = 'div'
DK_STR_GAME_TABLE_SPREAD_CLASS = 'sportsbook-outcome-cell__label-line-container'

DK_STR_GAME_TABLE_ODDS_TAG = 'span'
DK_STR_GAME_TABLE_ODDS_CLASS = 'sportsbook-odds'

DK_STR_GAME_TABLE_OVER_UNDER_TAG = 'span'
DK_STR_GAME_TABLE_OVER_UNDER_CLASS = 'sportsbook-outcome-cell__line'

DK_STR_DAILY_CARD_TAG = 'div'
DK_STR_DAILY_CARD_CLASS = 'parlay-card-10-a'

DK_STR_DAILY_CARD_DATE_TAG = 'div'
DK_STR_DAILY_CARD_DATE_CLASS = 'sportsbook-table-header__title'

DK_STR_SINGLE_GAME_START_TIME_TAG = 'span'
DK_STR_SINGLE_GAME_START_TIME_CLASS = 'event-cell__start-time'

DK_STR_SINGLE_GAME_STATUS_TAG = 'div'
DK_STR_SINGLE_GAME_STATUS_CLASS = 'event-cell__status'
DK_STR_SINGLE_GAME_TIME_TAG = 'span'
DK_STR_SINGLE_GAME_TIME_CLASS = 'event-cell__time'
DK_STR_SINGLE_GAME_PERIOD_TAG = 'span'
DK_STR_SINGLE_GAME_PERIOD_CLASS = 'event-cell__period'

DK_STR_SINGLE_GAME_EVENT_LINK_TAG = 'a'
DK_STR_SINGLE_GAME_EVENT_LINK_CLASS = 'event-cell-link'

FULL_GAME_PARAMS = {'category': 'game-lines', 'subcategory': 'game'}

CFB_URL = 'https://sportsbook.draftkings.com/leagues/football/88670775' #?category=game-lines&subcategory=game'
CFB_SHEET_NAME = 'CFB: DraftKings (Full Game)'

NCAAM_URL = 'https://sportsbook.draftkings.com/leagues/basketball/88670771' #?category=game-lines&subcategory=game
NCAAM_SHEET_NAME = 'NCAAM: DraftKings (Full Game)'

NFL_URL = 'https://sportsbook.draftkings.com/leagues/football/88670561' #?category=game-lines&subcategory=game
NBA_URL = 'https://sportsbook.draftkings.com/leagues/basketball/88670846' #?category=game-lines&subcategory=game

CFB_SHEET_INDEX = 0
NCAAM_SHEET_INDEX = 1

REQUEST_TIMEOUT = 5

class DraftKingsSingleEvent(event.SingleEvent):
    '''A single event (game) including basic gambling information.'''

    __slots__ = [
        'last_updated',
        'event_id',
        'game_date',
        'game_time',
        'in_progress',
        'away_team',
        'home_team',
        'betting_lines',
        'betting_choices',
        'outcome',
        'sheet_name',
        'database'
    ]

    def __init__(
        self,
        database):

        for slot in self.__slots__:
            self.__setattr__(slot, '')

        self.betting_lines = []
        self.betting_choices = event.BettingChoices()
        self.outcome = None
        self.database = database

        return

    def create_event_url(
        self) -> str:

        return f'{DK_STR_EVENTS_URL}/{self.event_id}' if self.event_id else ''

    def update_from_rows(
        self,
        rows: List[bs],
        **kwargs) -> None:

        # we do not want to update in-progress events. this allows us to
        # retain the last update as the closing lines for the event.
        self.in_progress = kwargs['in_progress'] if 'in_progress' in kwargs else False
        if self.in_progress:
            return

        # this event_id is unique to this event and is used to generate
        # the link back to the event page on draftkings. we only need to
        # set it once, as it will never change.
        if not self.event_id:
            if 'event_id' in kwargs:
                self.event_id = kwargs['event_id']

        # we only need to update the team names once.
        if not self.away_team or not self.home_team:
            teams = list(map(lambda t: t.find_all([DK_STR_TEAMS_TAG], class_ = DK_STR_TEAMS_CLASS), rows))
            if len(teams) != 2:
                return

            self.away_team = teams[0][0].text
            self.home_team = teams[1][0].text

        # the top row is the away team, and the bottom row is the home
        # team. the top row also holds the over data, and the bottom row
        # also holds the under data.
        top_row = rows[0]
        bottom_row = rows[1]

        top_row_columns = top_row.find_all([DK_STR_GAME_TABLE_ROW_COLUMN_TAG], class_ = [DK_STR_GAME_TABLE_ROW_COLUMN_CLASS, DK_STR_GAME_TABLE_EMPTY_CELL])
        bottom_row_columns = bottom_row.find_all([DK_STR_GAME_TABLE_ROW_COLUMN_TAG], class_ = [DK_STR_GAME_TABLE_ROW_COLUMN_CLASS, DK_STR_GAME_TABLE_EMPTY_CELL])

        new_betting_lines = event.EventLines()
        away_spread = 0
        if len(top_row_columns) > 0:
            spread = top_row_columns[0].find([DK_STR_GAME_TABLE_SPREAD_TAG], class_ = DK_STR_GAME_TABLE_SPREAD_CLASS)
            away_spread = spread.text if spread else 0
        new_betting_lines.away_team_spread = float(away_spread) if away_spread != 'pk' else 0

        odds = 0
        if len(top_row_columns) > 0:
            odds = top_row_columns[0].find([DK_STR_GAME_TABLE_ODDS_TAG], class_ = DK_STR_GAME_TABLE_ODDS_CLASS)
            new_betting_lines.away_team_odds = float(odds.text) if odds else 0
        new_betting_lines.away_team_moneyline = 0
        if len(top_row_columns) > 2:
            new_betting_lines.away_team_moneyline = float(top_row_columns[2].text) if top_row_columns[2].text else 0

        home_spread = 0
        if len(bottom_row_columns) > 0:
            spread = bottom_row_columns[0].find([DK_STR_GAME_TABLE_SPREAD_TAG], class_ = DK_STR_GAME_TABLE_SPREAD_CLASS)
            home_spread = spread.text if spread else 0
        new_betting_lines.home_team_spread = float(home_spread) if home_spread != 'pk' else 0

        odds = 0
        if len(bottom_row_columns) > 0:
            odds = bottom_row_columns[0].find([DK_STR_GAME_TABLE_ODDS_TAG], class_ = DK_STR_GAME_TABLE_ODDS_CLASS)
        new_betting_lines.home_team_odds = float(odds.text) if odds else 0
        new_betting_lines.home_team_moneyline = 0
        if len(bottom_row_columns) > 2:
            new_betting_lines.home_team_moneyline = float(bottom_row_columns[2].text) if bottom_row_columns[2].text else 0

        new_betting_lines.over_under = 0
        new_betting_lines.over_odds = 0
        new_betting_lines.under_odds = 0

        if len(top_row_columns) > 1:
            new_betting_lines.over_under = top_row_columns[1].find([DK_STR_GAME_TABLE_OVER_UNDER_TAG], class_ = DK_STR_GAME_TABLE_OVER_UNDER_CLASS)
            if new_betting_lines.over_under:
                new_betting_lines.over_under = float(new_betting_lines.over_under.text)

        if len(top_row_columns) > 1:
            over_odds = top_row_columns[1].find([DK_STR_GAME_TABLE_ODDS_TAG], class_ = DK_STR_GAME_TABLE_ODDS_CLASS)
            new_betting_lines.over_odds = float(over_odds.text) if over_odds else 0

        if len(bottom_row_columns) > 1:
            under_odds = bottom_row_columns[1].find([DK_STR_GAME_TABLE_ODDS_TAG], class_ = DK_STR_GAME_TABLE_ODDS_CLASS )
            new_betting_lines.under_odds = float(under_odds.text) if under_odds else 0

        # if we have not set the game date yet, do so now. we want to
        # convert text like 'today' and 'tomorrow' into absolute dates.
        # we also need to account for some bullshit time zone issues, as
        # it appears that we get times back in zulu/utc. one day i'll
        # fix this properly, but it seems to work with this approach for
        # now.
        if not self.game_date:
            self.game_date = kwargs['date'].strip() if 'date' in kwargs else ''
            if self.game_date.lower() == 'today':
                # check if it's past 7 pm eastern. if so, 'today' is really 'tomorrow'
                if datetime.datetime.now().time() > datetime.time(19, 0, 0):
                    self.game_date = (datetime.datetime.today() + datetime.timedelta(days = 1)).strftime('%a %b %-d')
                else:
                    # probably need a similar check here, too
                    self.game_date = datetime.date.today().strftime('%a %b %-d')
            elif self.game_date.lower() == 'tomorrow':
                self.game_date = (datetime.date.today() + datetime.timedelta(days = 1)).strftime('%a %b %-d')
            elif self.game_date:
                self.game_date = self.game_date[0:-2]

        if not self.game_time:
            if 'time' in kwargs:
                self.game_time = kwargs['time'].strip()

        last_updated = f'{datetime.date.today()} {datetime.datetime.now().strftime("%H:%M:%S")}'
        new_betting_lines.last_updated = last_updated
        self.add_update(new_betting_lines)

        return


class DraftKingsEventGroup:
    '''A single category of events for a single sport/event type.'''

    __slots__ = [
        'last_updated',
        'url',
        'url_params',
        'sheet_id',
        'sheet_name',
        'events',
        'skip_missing_moneyline',
        'include_kenpom',
        'names_to_update',
        'database_name',
        'database'
    ]

    def __init__(
        self,
        url: str,
        url_params: str,
        sheet_id: int,
        sheet_name: str,
        skip_missing_moneyline: bool,
        include_kenpom: bool,
        database_client,
        database_name: str):

        self.url = url
        self.url_params = url_params
        self.sheet_id = sheet_id
        self.sheet_name = sheet_name
        self.skip_missing_moneyline = skip_missing_moneyline
        self.include_kenpom = include_kenpom
        self.names_to_update = []
        self.database_name = database_name
        self.database = database_client.get_database(database_name)

        return

    def load_from_url(
        self,
        **kwargs) -> bool:

        if not self.url:
            return False

        kenpom_events = kenpom.get_kenpom_events() if self.include_kenpom else []
        self.names_to_update = []

        cookies = kwargs['cookies'] if 'cookies' in kwargs else {}
        headers = kwargs['headers'] if 'headers' in kwargs else {}
        response = requests.get(self.url, cookies = cookies, headers = headers, params = self.url_params, timeout = REQUEST_TIMEOUT)
        print(f'Retrieved {self.sheet_name} data from: {response.url}')

        doc = bs(response.text, 'html.parser')
        daily_cards = doc.find_all([DK_STR_DAILY_CARD_TAG], class_ = DK_STR_DAILY_CARD_CLASS)

        self.events = []
        for day in daily_cards:
            date = day.find([DK_STR_DAILY_CARD_DATE_TAG], class_ = DK_STR_DAILY_CARD_DATE_CLASS)
            if not date:
                continue

            date = date.text
            day_table = day.find([DK_STR_GAME_TABLE_TAG], class_ = DK_STR_GAME_TABLE_CLASS)

            day_rows = day_table.find_all([DK_STR_GAME_TABLE_ROW_TAG])
            #jmd: only run a few games
            #day_rows = day_rows[0:6]
            for row in range(0, len(day_rows), 2):
                start_time = ''
                in_progress = False

                label = day_rows[row].find([DK_STR_SINGLE_GAME_START_TIME_TAG], class_ = DK_STR_SINGLE_GAME_START_TIME_CLASS)
                if label:
                    start_time = label.text
                else:
                    label = day_rows[row].find([DK_STR_SINGLE_GAME_STATUS_TAG], class_ = DK_STR_SINGLE_GAME_STATUS_CLASS)
                    if label:
                        start_time = f'{label.find([DK_STR_SINGLE_GAME_TIME_TAG], class_ = DK_STR_SINGLE_GAME_TIME_CLASS).text} | {label.find([DK_STR_SINGLE_GAME_PERIOD_TAG], class_ = DK_STR_SINGLE_GAME_PERIOD_CLASS).text}'
                        in_progress = True

                event_id = day_rows[row].find([DK_STR_SINGLE_GAME_EVENT_LINK_TAG], class_ = DK_STR_SINGLE_GAME_EVENT_LINK_CLASS).attrs['href'].split('/', -1)[-1]

                new_event = DraftKingsSingleEvent(self.database)
                _ = event.populate_event_from_database(
                    self.database,
                    new_event,
                    event_id)

                new_event.update_from_rows([day_rows[row], day_rows[row + 1]], date = date, time = start_time, event_id = event_id, in_progress = in_progress)
                new_event.sheet_name = self.sheet_name

                skip = False
                if not new_event.game_date:
                    skip = True

                # jmd: temporary hack: only accept events that have a valid moneyline
                if self.skip_missing_moneyline:
                    if new_event.betting_lines:
                        if new_event.betting_lines[0].home_team_moneyline == 0 and new_event.betting_lines[0].away_team_moneyline == 0:
                            skip = True

                if not skip:
                    self.events.append(new_event)

                    for kenpom_event in kenpom_events:
                        if kenpom_event.contains_team(new_event.home_team) and kenpom_event.contains_team(new_event.away_team):
                            new_event.betting_lines[-1].kenpom_event = kenpom_event if new_event.betting_lines else None
                        elif kenpom_event.contains_team(new_event.home_team) or kenpom_event.contains_team(new_event.away_team):
                            if new_event.home_team != kenpom_event.home_team:
                                self.names_to_update.append((kenpom_event.home_team, new_event.home_team))
                            if new_event.away_team != kenpom_event.away_team:
                                self.names_to_update.append((kenpom_event.away_team, new_event.away_team))

                    new_event.update_database()
                else:
                    game_time_string = f' ({new_event.game_date}, {new_event.game_time})' if (new_event.game_date and new_event.game_time) else ''
                    event_url = f' - {new_event.create_event_url()}'
                    print(f'Skipping incomplete event: {new_event.away_team} @ {new_event.home_team}{game_time_string}{event_url}')
                    continue

        self.last_updated = f'{datetime.date.today()} {datetime.datetime.now().strftime("%H:%M:%S")}'
        return True


def main(
    args: argparse.Namespace) -> None:

    db_client = MongoClient('localhost')

    event_groups = []
    if args.cfb:
        skip_missing_moneyline = False
        include_kenpom = False
        event_groups.append(DraftKingsEventGroup(
            CFB_URL,
            FULL_GAME_PARAMS,
            CFB_SHEET_INDEX,
            CFB_SHEET_NAME,
            skip_missing_moneyline,
            include_kenpom,
            db_client,
            'cfb')
        )
    if args.ncaam:
        skip_missing_moneyline = True
        include_kenpom = True
        event_groups.append(DraftKingsEventGroup(
            NCAAM_URL,
            FULL_GAME_PARAMS,
            NCAAM_SHEET_INDEX,
            NCAAM_SHEET_NAME,
            skip_missing_moneyline,
            include_kenpom,
            db_client,
            'ncaam')
        )

    if not event_groups:
        print('ERROR: At least one event group (CFB, NCAAM, etc.) must be specified; exiting.')
        return

    service = gsu.get_spreadsheet_service()
    service._http.timeout = REQUEST_TIMEOUT

    if not args.new_spreadsheet:
        print(f'Updating spreadsheet ({args.existing_spreadsheet}): {gsu.create_spreadsheet_url(args.existing_spreadsheet)}')

    # cookies = dict(clientDateOffset = '240') # when DST is active
    cookies = dict(clientDateOffset = '300') # when DST is inactive
    headers = {
        "User-Agent": "Mozilla/5.0 (X11; CrOS x86_64 12871.102.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.141 Safari/537.36"
    }

    for event_group in event_groups:
        event_group.load_from_url(cookies = cookies, headers = headers)

        if event_group.names_to_update:
            print('The following name mismatches were detected. Add these entries to team_index.py:')
            for name in event_group.names_to_update:
                print(f'\'{name[0]}\': \'{name[1]}\',')

        print(f'  Retrieved data for {len(event_group.events)} qualifying events')

    if args.new_spreadsheet:
        gsu.create_new_spreadsheet_from_events(
            'KEEP GAMING',
            event_groups)
    else:
        gsu.update_spreadsheet_from_events(
            args.existing_spreadsheet,
            event_groups)

    return

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description =
        'Keep Gaming! is a command line tool for automatically downloading \
        DraftKings betting lines and creating a Google Sheets spreadsheet for \
        managing bets and tracking line changes throughout the week.')

    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        '--new',
        action = 'store_true',
        dest = 'new_spreadsheet',
        default = False,
        help = 'Create and populate a new spreadsheet.')
    group.add_argument(
        '--update',
        dest = 'existing_spreadsheet',
        default = '',
        help = 'Update the specified spreadsheet.')
    parser.add_argument(
        '--cfb',
        action = 'store_true',
        dest = 'cfb',
        default = False,
        help = 'Request CFB data'
    )
    parser.add_argument(
        '--ncaam',
        action = 'store_true',
        dest = 'ncaam',
        default = False,
        help = 'Request NCAAM data'
    )

    try:
        args = parser.parse_args()
    except Exception as e:
        print(f'An error occurred: {str(e)}')
        exit(1)

    main(args)
    exit(0)
