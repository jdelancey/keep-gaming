import plotly.express as px
import plotly.graph_objects as go
from pymongo import database, MongoClient
from typing import List

from kenpom import KenPomEvent

class BettingChoices:
    '''A grouping of betting choices for a single event.'''

    __slots__ = [
        'bet_away_spread',
        'bet_away_moneyline',
        'bet_home_spread',
        'bet_home_moneyline',
        'bet_over',
        'bet_under'
    ]

    def __init__(self):

        for slot in self.__slots__:
            self.__setattr__(slot, False)

        return

    def load_from_database(
        self,
        db_entry: dict):

        self.bet_away_spread = db_entry['bet_away_spread'] if 'bet_away_spread' in db_entry else ''
        self.bet_away_moneyline = db_entry['bet_away_moneyline'] if 'bet_away_moneyline' in db_entry else ''
        self.bet_home_spread = db_entry['bet_home_spread'] if 'bet_home_spread' in db_entry else ''
        self.bet_home_moneyline = db_entry['bet_home_moneyline'] if 'bet_home_moneyline' in db_entry else ''
        self.bet_over = db_entry['bet_over'] if 'bet_over' in db_entry else ''
        self.bet_under = db_entry['bet_under'] if 'bet_under' in db_entry else ''

        return

    def create_mongodb_dict(
        self) -> dict:

        d = {
            'bet_away_spread': self.bet_away_spread,
            'bet_away_moneyline': self.bet_away_moneyline,
            'bet_home_spread': self.bet_home_spread,
            'bet_home_moneyline': self.bet_home_moneyline,
            'bet_over': self.bet_over,
            'bet_under': self.bet_under
        }

        return d

    def print(
        self):

        print()
        print(f'  {str("Bet away spread:").ljust(22)} {self.bet_away_spread}')
        print(f'  {str("Bet away moneyline:").ljust(22)} {self.bet_away_moneyline}')
        print(f'  {str("Bet home spread:").ljust(22)} {self.bet_home_spread}')
        print(f'  {str("Bet home moneyline:").ljust(22)} {self.bet_home_moneyline}')
        print(f'  {str("Bet over:").ljust(22)} {self.bet_over}')
        print(f'  {str("Bet under:").ljust(22)} {self.bet_under}')

        return


class EventLines:
    '''A grouping of betting lines related to a single event.'''

    __slots__ = [
        'last_updated',
        'away_team_spread',
        'away_team_odds',
        'away_team_moneyline',
        'home_team_spread',
        'home_team_odds',
        'home_team_moneyline',
        'over_under',
        'over_odds',
        'under_odds',
        'kenpom_event'
    ]

    def __init__(
        self):

        for slot in self.__slots__:
            self.__setattr__(slot, '')

        self.kenpom_event = None

        return

    def load_from_database(
        self,
        db_entry: dict):

        self.last_updated = db_entry['last_updated'] if 'last_updated' in db_entry else ''
        self.away_team_spread = db_entry['away_team_spread'] if 'away_team_spread' in db_entry else ''
        self.away_team_odds = db_entry['away_team_odds'] if 'away_team_odds' in db_entry else ''
        self.away_team_moneyline = db_entry['away_team_moneyline'] if 'away_team_moneyline' in db_entry else ''
        self.home_team_spread = db_entry['home_team_spread'] if 'home_team_spread' in db_entry else ''
        self.home_team_odds = db_entry['home_team_odds'] if 'home_team_odds' in db_entry else ''
        self.home_team_moneyline = db_entry['home_team_moneyline'] if 'home_team_moneyline' in db_entry else ''
        self.over_under = db_entry['over_under'] if 'over_under' in db_entry else ''
        self.over_odds = db_entry['over_odds'] if 'over_odds' in db_entry else ''
        self.under_odds = db_entry['under_odds'] if 'under_odds' in db_entry else ''

        if 'kenpom_event' in db_entry:
            self.kenpom_event = KenPomEvent()
            self.kenpom_event.load_from_database(db_entry['kenpom_event'])

        return

    def create_mongodb_dict(
        self) -> dict:

        d = {
            'last_updated': self.last_updated,
            'away_team_spread': self.away_team_spread,
            'away_team_odds': self.away_team_odds,
            'away_team_moneyline': self.away_team_moneyline,
            'home_team_spread': self.home_team_spread,
            'home_team_odds': self.home_team_odds,
            'home_team_moneyline': self.home_team_moneyline,
            'over_under': self.over_under,
            'over_odds': self.over_odds,
            'under_odds': self.under_odds
        }

        if self.kenpom_event:
            d['kenpom_event'] = self.kenpom_event.create_mongodb_dict()

        return d

    def calculate_kelly_criterion(
        self,
        team: str,
        is_home_team: bool) -> float:

        if not self.kenpom_event:
            return 0

        moneyline = self.home_team_moneyline if is_home_team else self.away_team_moneyline
        decimal_odds = 0
        if moneyline >= 0:
            decimal_odds = (moneyline / 100) + 1
        else:
            decimal_odds = (100 / -moneyline) + 1


        b = decimal_odds - 1
        p = self.kenpom_event.confidence if self.kenpom_event.winning_team == team else 1 - self.kenpom_event.confidence
        q = 1 - p
        k = (b * p - q) / b if b != 0 else 0

        f = '{:.2f}'.format(k)
        return float(f)

    def print(
        self,
        away_team: str,
        home_team: str) -> None:

        print()
        print(f'    {str("Updated: ").ljust(15)}\t {self.last_updated}')
        print(f'      {away_team.ljust(15)}\t {self.away_team_spread} ({self.away_team_odds})\t Moneyline: {self.away_team_moneyline}')
        print(f'      {home_team.ljust(15)}\t {self.home_team_spread} ({self.home_team_odds})\t Moneyline: {self.home_team_moneyline}')
        print(f'      {str("Over:").ljust(15)}\t {self.over_under} ({self.over_odds})')
        print(f'      {str("Under:").ljust(15)}\t {self.over_under} ({self.under_odds})')

        if self.kenpom_event:
            print(f'      KenPom:')
            print(f'        Winner: {self.kenpom_event.winning_team} {self.kenpom_event.score} ({self.kenpom_event.confidence * 100}%)')

            k_home = self.calculate_kelly_criterion(home_team, True)
            k_away = self.calculate_kelly_criterion(away_team, False)

            if k_home > 0:
                print(f'    Kelly: {home_team}: {k_home * 100}%')
            if k_away > 0:
                print(f'    Kelly: {away_team}: {k_away * 100}%')

        return


class EventOutcome:
    '''The outcome of a single event.'''

    __slots__ = [
        'winning_team',
        'winning_team_score',
        'losing_team',
        'losing_team_score'
    ]

    def __init__(self):

        for slot in self.__slots__:
            self.__setattr__(slot, '')

        return


class SingleEvent:
    '''A single event (game) with betting information.'''

    __slots__ = [
        'last_updated',
        'event_id',
        'game_date',
        'game_time',
        'away_team',
        'home_team',
        'betting_lines',
        'betting_choices',
        'outcome',
        'database'
    ]

    def __init__(
        self):

        for slot in self.__slots__:
            self.__setattr__(slot, '')

        self.betting_lines = []
        self.betting_choices = BettingChoices()
        self.outcome = None
        self.database = None

        return

    def __init__(
        self,
        database):

        for slot in self.__slots__:
            self.__setattr__(slot, '')

        self.betting_lines = []
        self.betting_choices = BettingChoices()
        self.outcome = None
        self.database = database

        return

    def create_event_url(
        self) -> str:

        return 'Link not implemented'

    def add_update(
        self,
        update: EventLines) -> None:

        self.betting_lines.append(update)
        self.last_updated = update.last_updated

        return

    def set_outcome(
        self,
        outcome: EventOutcome) -> None:

        self.outcome = outcome

        return

    def create_empty_database_document(
        self) -> bool:

        if self.database is None:
            return False

        found_event = self.database.events.find({
            'event_id': self.event_id
        })

        try:
            _ = found_event[0]
        except:
            event = {
                'event_id': self.event_id,
                'last_updated': self.last_updated,
                'game_date': self.game_date,
                'game_time': self.game_time,
                'away_team': self.away_team,
                'home_team': self.home_team,
                'betting_lines': [],
                'betting_choices': {},
                'outcome': {}
            }

            return self.database.events.insert_one(event)

        return False

    def append_latest_lines_to_database(
        self) -> None:

        if self.database is None or not self.betting_lines:
            return

        results = self.database.events.find({'event_id': self.event_id})
        try:
            result = results[0]
        except:
            return

        self.database.events.update_one(
            {
                '_id': result['_id']
            },
            {
                '$push': {
                    'betting_lines': self.betting_lines[-1].create_mongodb_dict()
                }
            }
        )

        return

    def update_database(
        self) -> None:

        # ensure we have a document and create one if we don't
        self.create_empty_database_document()

        # update the document whether it existed or we just created it
        self.append_latest_lines_to_database()

        return

    def update_betting_choices_in_database(
        self) -> None:

        if self.database is None or not self.betting_choices:
            return False

        results = self.database.events.find({'event_id': self.event_id})
        try:
            result = results[0]
        except:
            return False

        self.database.events.update_one(
            {
                '_id': result['_id']
            },
            {
                '$set': {
                    'betting_choices': self.betting_choices.create_mongodb_dict()
                }
            }
        )

        return

    def print(
        self) -> None:

        if not self.betting_lines:
            return

        game_time_string = f' ({self.game_date}, {self.game_time})' if (self.game_date and self.game_time) else ''
        event_url = f' - {self.create_event_url()}'

        print()
        print(f'Summary of {self.away_team} @ {self.home_team}{game_time_string}{event_url}')
        self.betting_choices.print()

        print()
        print('  Starting lines:')
        self.betting_lines[0].print(
            self.away_team,
            self.home_team
        )

        print()
        print('  Latest lines:')
        self.betting_lines[-1].print(
            self.away_team,
            self.home_team
        )

        return

    def write_html(
        self,
        filename) -> None:

        plots = [
            self.plot_away_spread(),
            self.plot_away_moneyline(),
            self.plot_away_kelly(),
            self.plot_home_spread(),
            self.plot_home_moneyline(),
            self.plot_home_kelly(),
            self.plot_over_under()
        ]

        write = False
        for plot in plots:
            if plot:
                write = True
                break

        if not write:
            return

        with open('/home/delancey/projects/dk/template.html', 'r') as f:
            template = f.readlines()
        f.close()

        game_time_string = f' ({self.game_date}, {self.game_time})' if (self.game_date and self.game_time) else ''

        template = '\n'.join(template)
        template = template.replace('REPLACE_TITLE', f'Summary of {self.away_team} @ {self.home_team}')
        template = template.replace('REPLACE_H1', f'Summary of <a href="{self.create_event_url()}">{self.away_team} @ {self.home_team}{game_time_string}</a>')
        template = template.replace('REPLACE_BODY', '\n'.join(plots))

        with open(filename, 'w') as f:
            f.write(template)
        f.close()

        return

    def plot_away_kelly(
        self) -> str:

        if not self.betting_lines or not self.betting_lines[0].kenpom_event:
            return ''

        kellys = list(map(lambda x: x.calculate_kelly_criterion(self.away_team, False) * 100, self.betting_lines))
        updates = list(map(lambda x: x.last_updated, self.betting_lines))
        return self.create_plot_html(
            updates,
            kellys,
            'Date/Time',
            f'{self.away_team} Kelly Criterion (%)',
            f'{self.away_team} Kelly Criterion vs. Time',
            fixed_y = 0,
            fixed_y_label = 'Safe Bet Line'
        )

    def plot_home_kelly(
        self) -> str:

        if not self.betting_lines or not self.betting_lines[0].kenpom_event:
            return ''

        kellys = list(map(lambda x: x.calculate_kelly_criterion(self.home_team, True) * 100, self.betting_lines))
        updates = list(map(lambda x: x.last_updated, self.betting_lines))
        return self.create_plot_html(
            updates,
            kellys,
            'Date/Time',
            f'{self.home_team} Kelly Criterion (%)',
            f'{self.home_team} Kelly Criterion vs. Time',
            fixed_y = 0,
            fixed_y_label = 'Safe Bet Line'
        )

    def plot_away_spread(
        self) -> str:

        # if not self.betting_choices.bet_away_spread or not self.betting_lines:
        #     return ''

        spreads = list(map(lambda x: x.away_team_spread, self.betting_lines))
        updates = list(map(lambda x: x.last_updated, self.betting_lines))
        return self.create_plot_html(
            updates,
            spreads,
            'Date/Time',
            f'{self.away_team} Spread',
            f'{self.away_team} Spread vs. Time',
            fixed_y = self.betting_lines[0].away_team_spread if self.betting_choices.bet_away_spread else None
        )

    def plot_home_spread(
        self) -> str:

        # if not self.betting_choices.bet_home_spread or not self.betting_lines:
        #     return ''

        spreads = list(map(lambda x: x.home_team_spread, self.betting_lines))
        updates = list(map(lambda x: x.last_updated, self.betting_lines))
        return self.create_plot_html(
            updates,
            spreads,
            'Date/Time',
            f'{self.home_team} Spread',
            f'{self.home_team} Spread vs. Time',
            fixed_y = self.betting_lines[0].home_team_spread if self.betting_choices.bet_home_spread else None
        )

    def plot_over_under(
        self) -> str:

        # if (not self.betting_choices.bet_over and not self.betting_choices.bet_under) or not self.betting_lines:
        #     return ''

        over_unders = list(map(lambda x: x.over_under, self.betting_lines))
        updates = list(map(lambda x: x.last_updated, self.betting_lines))
        return self.create_plot_html(
            updates,
            over_unders,
            'Date/Time',
            f'Over/Under',
            f'{self.home_team} vs {self.away_team} Over/Under vs. Time',
            fixed_y = self.betting_lines[0].over_under if self.betting_choices.bet_over or self.betting_choices.bet_under else None
        )

    def plot_away_moneyline(
        self) -> str:

        # if not self.betting_choices.bet_away_moneyline or not self.betting_lines:
        #     return ''

        moneylines = list(map(lambda x: x.away_team_moneyline, self.betting_lines))
        updates = list(map(lambda x: x.last_updated, self.betting_lines))
        return self.create_plot_html(
            updates,
            moneylines,
            'Date/Time',
            f'{self.away_team} Moneyline',
            f'{self.away_team} Moneyline vs. Time',
            fixed_y = self.betting_lines[0].away_team_moneyline if self.betting_choices.bet_away_moneyline else None
        )

    def plot_home_moneyline(
        self) -> str:

        # if not self.betting_choices.bet_home_moneyline or not self.betting_lines:
        #     return ''

        moneylines = list(map(lambda x: x.home_team_moneyline, self.betting_lines))
        updates = list(map(lambda x: x.last_updated, self.betting_lines))
        return self.create_plot_html(
            updates,
            moneylines,
            'Date/Time',
            f'{self.home_team} Moneyline',
            f'{self.home_team} Moneyline vs. Time',
            fixed_y = self.betting_lines[0].home_team_moneyline if self.betting_choices.bet_home_moneyline else None
        )

    def create_plot_html(
        self,
        x_data: List[float],
        y_data: List[float],
        x_label: str,
        y_label: str,
        title: str,
        **kwargs) -> str:

        layout = dict(
            title = title,
            xaxis = dict(
                title = x_label
            ),
            yaxis = dict(
                title = y_label
            )
        )
        fig = go.Figure(layout = layout)

        fig.add_trace(go.Scatter(
            x = x_data,
            y = y_data,
            mode = 'lines+markers',
            line_shape = 'hv',
            name = y_label
        ))

        fixed_y = kwargs['fixed_y'] if 'fixed_y' in kwargs else None
        if fixed_y is not None:
            fixed_y_label = kwargs['fixed_y_label'] if 'fixed_y_label' in kwargs else (f'Bet Placed: {"+" if fixed_y > 0 else ""}{fixed_y}')
            fig.add_trace(go.Scatter(
                x = x_data,
                y = [fixed_y for _ in range(len(x_data))],
                mode = 'lines',
                name = fixed_y_label
            ))

        #fig.show()
        return fig.to_html(full_html = False, include_plotlyjs = 'cdn')

def populate_event_from_database(
    database,
    new_event: SingleEvent,
    event_id: str) -> bool:

    if database is None:
        return False

    found = database.events.find_one({
        'event_id': event_id
    })

    if not found:
        return False

    #new_event = SingleEvent(database)
    new_event.event_id = event_id
    new_event.game_date = found['game_date'] if 'game_date' in found else ''
    new_event.game_time = found['game_time'] if 'game_time' in found else ''
    new_event.away_team = found['away_team'] if 'away_team' in found else ''
    new_event.home_team = found['home_team'] if 'home_team' in found else ''
    new_event.game_date = found['game_date'] if 'game_date' in found else ''

    betting_lines = found['betting_lines'] if 'betting_lines' in found else []
    for lines in betting_lines:
        new_lines = EventLines()
        new_lines.load_from_database(lines)
        new_event.add_update(new_lines)

    betting_choices = found['betting_choices'] if 'betting_choices' in found else None
    if betting_choices:
        choices = BettingChoices()
        choices.load_from_database(betting_choices)
        new_event.betting_choices = choices
        new_event.update_betting_choices_in_database()

    return True
