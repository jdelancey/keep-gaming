#!/usr/bin/python3

import argparse
from bs4 import BeautifulSoup as bs
from datetime import date, datetime, timedelta
import requests
from typing import List

import google_sheets_utils as gsu

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

CFB_URL = 'https://sportsbook.draftkings.com/leagues/football/88670775' #?category=game-lines&subcategory=game'
CFB_FIRST_HALF_URL = 'https://sportsbook.draftkings.com/leagues/football/88670775' #?category=halves'
NCAAM_URL = 'https://sportsbook.draftkings.com/leagues/basketball/88670771' #?category=game-lines&subcategory=game

NFL_URL = 'https://sportsbook.draftkings.com/leagues/football/88670561' #?category=game-lines&subcategory=game

NBA_URL = 'https://sportsbook.draftkings.com/leagues/basketball/88670846' #?category=game-lines&subcategory=game

class SingleEvent:
    '''A single event (game) including basic gambling information.'''

    __slots__ = [
        'last_updated',
        'event_id',
        'game_date',
        'game_time',
        'away_team',
        'away_team_spread',
        'away_team_odds',
        'away_team_moneyline',
        'home_team',
        'home_team_spread',
        'home_team_odds',
        'home_team_moneyline',
        'over_under',
        'over_odds',
        'under_odds',
        'sheet_name'
    ]

    def __init__(
        self):

        for slot in self.__slots__:
            self.__setattr__(slot, '')

        return

    def create_event_url(
        self) -> str:

        return f'{DK_STR_EVENTS_URL}/{self.event_id}' if self.event_id else ''

    def load_from_rows(
        self,
        rows: List[bs],
        **kwargs):

        teams = list(map(lambda t: t.find_all([DK_STR_TEAMS_TAG], class_ = DK_STR_TEAMS_CLASS), rows))
        if len(teams) != 2:
            return

        self.away_team = teams[0][0].text
        self.home_team = teams[1][0].text

        top_row = rows[0]
        bottom_row = rows[1]

        top_row_columns = top_row.find_all([DK_STR_GAME_TABLE_ROW_COLUMN_TAG], class_ = [DK_STR_GAME_TABLE_ROW_COLUMN_CLASS, DK_STR_GAME_TABLE_EMPTY_CELL])
        bottom_row_columns = bottom_row.find_all([DK_STR_GAME_TABLE_ROW_COLUMN_TAG], class_ = [DK_STR_GAME_TABLE_ROW_COLUMN_CLASS, DK_STR_GAME_TABLE_EMPTY_CELL])

        away_spread = 0
        if len(top_row_columns) > 0:
            spread = top_row_columns[0].find([DK_STR_GAME_TABLE_SPREAD_TAG], class_ = DK_STR_GAME_TABLE_SPREAD_CLASS)
            away_spread = spread.text if spread else 0
        self.away_team_spread = float(away_spread) if away_spread != 'pk' else 0

        odds = 0
        if len(top_row_columns) > 0:
            odds = top_row_columns[0].find([DK_STR_GAME_TABLE_ODDS_TAG], class_ = DK_STR_GAME_TABLE_ODDS_CLASS)
            self.away_team_odds = float(odds.text) if odds else 0
        self.away_team_moneyline = 0
        if len(top_row_columns) > 2:
            self.away_team_moneyline = float(top_row_columns[2].text) if top_row_columns[2].text else 0

        home_spread = 0
        if len(bottom_row_columns) > 0:
            spread = bottom_row_columns[0].find([DK_STR_GAME_TABLE_SPREAD_TAG], class_ = DK_STR_GAME_TABLE_SPREAD_CLASS)
            home_spread = spread.text if spread else 0
        self.home_team_spread = float(home_spread) if home_spread != 'pk' else 0

        odds = 0
        if len(bottom_row_columns) > 0:
            odds = bottom_row_columns[0].find([DK_STR_GAME_TABLE_ODDS_TAG], class_ = DK_STR_GAME_TABLE_ODDS_CLASS)
        self.home_team_odds = float(odds.text) if odds else 0
        self.home_team_moneyline = 0
        if len(bottom_row_columns) > 2:
            self.home_team_moneyline = float(bottom_row_columns[2].text) if bottom_row_columns[2].text else 0

        self.over_under = 0
        self.over_odds = 0
        self.under_odds = 0

        if len(top_row_columns) > 1:
            self.over_under = top_row_columns[1].find([DK_STR_GAME_TABLE_OVER_UNDER_TAG], class_ = DK_STR_GAME_TABLE_OVER_UNDER_CLASS)
            if self.over_under:
                self.over_under = float(self.over_under.text)

        if len(top_row_columns) > 1:
            over_odds = top_row_columns[1].find([DK_STR_GAME_TABLE_ODDS_TAG], class_ = DK_STR_GAME_TABLE_ODDS_CLASS)
            self.over_odds = float(over_odds.text) if over_odds else 0

        if len(bottom_row_columns) > 1:
            under_odds = bottom_row_columns[1].find([DK_STR_GAME_TABLE_ODDS_TAG], class_ = DK_STR_GAME_TABLE_ODDS_CLASS )
            self.under_odds = float(under_odds.text) if under_odds else 0

        self.game_date = kwargs['date'].strip() if 'date' in kwargs else ''
        if self.game_date.lower() == 'today':
            self.game_date = datetime.today().strftime('%a %b %-d').upper()
        elif self.game_date.lower() == 'tomorrow':
            self.game_date = (datetime.today() + timedelta(days = 1)).strftime('%a %b %-d').upper()
        elif self.game_date:
            self.game_date = self.game_date[0:-2]

        self.game_time = kwargs['time'].strip() if 'time' in kwargs else ''
        self.event_id = kwargs['event_id'] if 'event_id' in kwargs else ''
        self.last_updated = f'{date.today()} {datetime.now().strftime("%H:%M:%S")}'

        return

    def load_from_table(
        self,
        table: bs,
        **kwargs):

        table_rows = table.find_all([DK_STR_GAME_TABLE_ROW_TAG])
        if len(table_rows) != 2:
            return

        self.load_from_rows([table_rows[0], table_rows[1]], kwargs)
        return

    def print(
        self) -> None:

        game_time_string = f' ({self.game_date}, {self.game_time})' if (self.game_date and self.game_time) else ''
        event_url = f' - {self.create_event_url()}'

        print()
        print(f'Summary of {self.away_team} @ {self.home_team}{game_time_string}{event_url}')
        print(f'  {str("Last Updated: ").ljust(15)}\t {date.today()} {datetime.now().strftime("%H:%M:%S")}')
        print(f'  {self.away_team.ljust(15)}\t {self.away_team_spread} ({self.away_team_odds})\t Moneyline: {self.away_team_moneyline}')
        print(f'  {self.home_team.ljust(15)}\t {self.home_team_spread} ({self.home_team_odds})\t Moneyline: {self.home_team_moneyline}')
        print(f'  {str("Over:").ljust(15)}\t {self.over_under} ({self.over_odds})')
        print(f'  {str("Under:").ljust(15)}\t {self.over_under} ({self.under_odds})')

        return

def main(args: argparse.Namespace) -> None:

    service = gsu.get_spreadsheet_service()
    service._http.timeout = 5

    if not args.new_spreadsheet:
        print(f'Updating spreadsheet ({args.existing_spreadsheet}): {gsu.create_spreadsheet_url(args.existing_spreadsheet)}')

    # cookies = dict(clientDateOffset = '240') # when DST is active
    cookies = dict(clientDateOffset = '300') # when DST is inactive
    headers = {
        "User-Agent": "Mozilla/5.0 (X11; CrOS x86_64 12871.102.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/81.0.4044.141 Safari/537.36"
    }

    urls = (
        (CFB_URL, {'category': 'game-lines', 'subcategory': 'game'}, 'CFB: DraftKings (Full Game)'),
        (NCAAM_URL, {'category': 'game-lines', 'subcategory': 'game'}, 'NCAAM: DraftKings (Full Game)'),
        #(NFL_URL, {'category': 'game-lines', 'subcategory': 'game'}, 'NFL: DraftKings (Full Game)'),
        #(NBA_URL, {'category': 'game-lines', 'subcategory': 'game'}, 'NBA: DraftKings (Full Game)')
    )

    all_events = []
    for url in urls:
        response = requests.get(url[0], cookies = cookies, headers = headers, params = url[1], timeout = 5)
        print(f'Retrieved {url[2]} data from: {response.url}')
        doc = bs(response.text, 'html.parser')

        daily_cards = doc.find_all([DK_STR_DAILY_CARD_TAG], class_ = DK_STR_DAILY_CARD_CLASS)

        events = []
        for day in daily_cards:
            date = day.find([DK_STR_DAILY_CARD_DATE_TAG], class_ = DK_STR_DAILY_CARD_DATE_CLASS)
            if not date:
                continue

            date = date.text
            day_table = day.find([DK_STR_GAME_TABLE_TAG], class_ = DK_STR_GAME_TABLE_CLASS)

            day_rows = day_table.find_all([DK_STR_GAME_TABLE_ROW_TAG])
            for row in range(0, len(day_rows), 2):
                start_time = ''

                label = day_rows[row].find([DK_STR_SINGLE_GAME_START_TIME_TAG], class_ = DK_STR_SINGLE_GAME_START_TIME_CLASS)
                if label:
                    start_time = label.text
                else:
                    label = day_rows[row].find([DK_STR_SINGLE_GAME_STATUS_TAG], class_ = DK_STR_SINGLE_GAME_STATUS_CLASS)
                    if label:
                        start_time = f'Event in progress when spreadsheet was built ({label.find([DK_STR_SINGLE_GAME_TIME_TAG], class_ = DK_STR_SINGLE_GAME_TIME_CLASS).text} | {label.find([DK_STR_SINGLE_GAME_PERIOD_TAG], class_ = DK_STR_SINGLE_GAME_PERIOD_CLASS).text})'

                event_id = day_rows[row].find([DK_STR_SINGLE_GAME_EVENT_LINK_TAG], class_ = DK_STR_SINGLE_GAME_EVENT_LINK_CLASS).attrs['href'].split('/', -1)[-1]
                new_event = SingleEvent()
                new_event.load_from_rows([day_rows[row], day_rows[row + 1]], date = date, time = start_time, event_id = event_id)
                if new_event.game_date:
                    new_event.sheet_name = url[2]
                    events.append(new_event)

        # jmd testing: only work on a few events when testing
        #events = events[0:2]
        # jmd testing: pop off a couple of the early games as though they have already been played
        #events = events[3:]
        # jmd testing: add a new fake event in update mode
        # if not args.new_spreadsheet:
        #     new_event = SingleEvent()
        #     new_event.event_id = '12345'
        #     new_event.home_team = 'test new home team'
        #     new_event.away_team = 'test new away team'
        #     events.append(new_event)

        all_events.append(events)

        print(f'  Retrieved data for {len(events)} qualifying events')

    if args.new_spreadsheet:
        gsu.create_new_spreadsheet_from_events(
            'KEEP GAMING',
            all_events)
    else:
        gsu.update_spreadsheet_from_events(
            args.existing_spreadsheet,
            all_events)

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

    try:
        args = parser.parse_args()
    except Exception as e:
        print(f'An error occurred: {str(e)}')
        exit(1)

    main(args)
    exit(0)
