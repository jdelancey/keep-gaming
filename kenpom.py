from bs4 import BeautifulSoup as bs
from datetime import date, datetime
import re
from selenium import webdriver
import time
from typing import List

import team_index
import kenpom_credentials as kpc

KP_STR_FANMATCH_TABLE_TAG = 'table'
KP_STR_FANMATCH_TABLE_ID = 'fanmatch-table'

KP_STR_FANMATCH_ROW_TAG = 'tr'
KP_STR_FANMATCH_COLUMN_TAG = 'td'

def standardize_team_name(
    name: str,
    all_names: dict) -> str:

    if 'St.' in name:
        name = name.replace('St.', 'State')

    if name in all_names:
        return all_names[name]

    return name


class KenPomEvent:

    __slots__ = [
        'last_updated',
        'away_team',
        'home_team',
        'winning_team',
        'score',
        'confidence'
    ]

    def __init__(
        self):

        for slot in self.__slots__:
            self.__setattr__(slot, '')

        return

    def load_from_row(
        self,
        row) -> None:

        columns = row.find_all([KP_STR_FANMATCH_COLUMN_TAG])
        if not columns or len(columns) < 2:
            return

        teams = columns[0].text.strip()
        prediction = columns[1].text.strip()

        team_a = ''
        team_b = ''
        split = ' at '
        if ' vs. ' in teams:
            split = ' vs. '

        result = re.split(split, teams)
        if not result or len(result) < 2:
            return

        team_a = result[0].strip()
        team_b = result[1].strip()

        if not 'NR ' in team_a:
            team_a = re.split('\d* ', team_a, 1)[1]
        else:
            team_a = re.split('NR ', team_a, 1)[1]

        if not 'NR ' in team_b:
            team_b = re.split('\d* ', team_b, 1)[1]
        else:
            team_b = re.split('NR ', team_b, 1)[1]

        team_a = standardize_team_name(team_a, team_index.NAME_REPLACEMENT_DICT)
        team_b = standardize_team_name(team_b, team_index.NAME_REPLACEMENT_DICT)
        self.home_team = team_b
        self.away_team = team_a
        if split == ' vs. ':
            self.home_team = team_b
            self.away_team = team_a

        score = re.search('(\d*-\d*)', prediction)
        percent = re.search('(\(.*%)', prediction)
        winning_team = re.search('([^\d]*)', prediction)

        if score and percent and winning_team:
            self.score = score.group(0).strip()
            self.winning_team = standardize_team_name(winning_team.group(0).strip(), team_index.NAME_REPLACEMENT_DICT)
            self.confidence = float(percent.group(0).strip()[1:-1])
            self.confidence = self.confidence / 100.0

        self.last_updated = f'{date.today()} {datetime.now().strftime("%H:%M:%S")}'

        return

    def contains_team(
        self,
        team: str) -> bool:

        if team == self.home_team or team == self.away_team:
            return True

        return False


def get_kenpom_events() -> List[KenPomEvent]:

    browser_options = webdriver.ChromeOptions()
    browser_options.add_argument('--no-sandbox')
    browser_options.add_argument('--headless')
    browser_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options = browser_options)
    browser.get('https://kenpom.com/fanmatch.php')

    email = browser.find_element_by_name('email')
    email.send_keys(kpc.email)
    password = browser.find_element_by_name('password')
    password.send_keys(kpc.password)
    login = browser.find_element_by_name('submit')
    login.click()

    doc = bs(browser.page_source, 'html.parser')
    time.sleep(1)
    browser.close()

    table = doc.find([KP_STR_FANMATCH_TABLE_TAG], id = KP_STR_FANMATCH_TABLE_ID)
    rows = table.find_all([KP_STR_FANMATCH_ROW_TAG])

    events = []
    for row in rows:
        event = KenPomEvent()
        event.load_from_row(row)
        if event.home_team and event.away_team:
            events.append(event)

    return events
