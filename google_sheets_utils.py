from datetime import date, datetime
from os import path
import string
import time
from typing import List, Tuple

from googleapiclient.discovery import Resource, build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from dk import SingleEvent, DK_STR_EVENTS_URL
import team_colors as tc

GOOGLE_CLIENT_SECRETS_FILE = './keys/app_secret.json'
GOOGLE_CREDENTIALS_FILE_LOCAL = './keys/client_secret_588444115345-rncj0gehs30silqrafl8cbtoo3p5j70f.apps.googleusercontent.com.json'

GOOGLE_API_SERVICE_NAME = 'sheets'
GOOGLE_API_SERVICE_VERSION = 'v4'
GOOGLE_API_SCOPES = ['https://www.googleapis.com/auth/drive.file']

GOOGLE_SHEETS_BASE_URL = 'https://docs.google.com/spreadsheets/d'

SLEEP_TIME = 3

GAME_LINK = 'Game Link'
GAME_DATE = 'Date'
START_TIME = 'Start Time'
MATCHUP = 'Matchup'
SPREAD = 'Spread'
BET_SPREAD = 'Bet Spread?'
LAST_BET_SPREAD = 'LAST BET (Spread)'
OVER_UNDER = 'O/U'
OVER_UNDER_SPACER = ''
BET_OVER_UNDER = 'Bet O/U?'
LAST_BET_OVER_UNDER = 'LAST BET (O/U)'
MONEYLINE = 'Moneyline'
BET_MONEYLINE = 'Bet Moneyline?'
UPDATED = 'Updated ->'
SPREAD_LATEST = 'Spread (Latest)'
SPREAD_MOVEMENT = 'Spread Movement'
OVER_UNDER_LATEST = 'O/U (Latest)'
OVER_UNDER_MOVEMENT = 'O/U Movement'
MONEYLINE_LATEST = 'Moneyline (Latest)'
MONEYLINE_MOVEMENT = 'Moneyline Movement)'

SHEET_COLUMNS = list(string.ascii_uppercase)
SHEET_HEADER_COLUMN_ORDER = [
    GAME_LINK,
    GAME_DATE,
    START_TIME,
    MATCHUP,
    SPREAD,
    BET_SPREAD,
    LAST_BET_SPREAD,
    OVER_UNDER,
    OVER_UNDER_SPACER,
    BET_OVER_UNDER,
    LAST_BET_OVER_UNDER,
    MONEYLINE,
    BET_MONEYLINE,
    UPDATED,
    SPREAD_LATEST,
    SPREAD_MOVEMENT,
    OVER_UNDER_LATEST,
    OVER_UNDER_MOVEMENT,
    MONEYLINE_LATEST,
    MONEYLINE_MOVEMENT
]

def hex_to_rgb(
    hex: str) -> Tuple[float]:

    hex = hex.lstrip('#')
    rgb_temp = list(int(hex[i:i+2], 16) for i in (0, 2, 4))
    rgb = []
    for val in rgb_temp:
        rgb.append(val / 255)

    return tuple(rgb)

def create_spreadsheet_url(
    spreadsheet_id: str) -> str:

    return f'{GOOGLE_SHEETS_BASE_URL}/{spreadsheet_id}'

def get_credentials(
    credentials_file: str,
    scopes: str) -> str:

    credentials = None
    if path.exists(credentials_file):
        credentials = Credentials.from_authorized_user_file(credentials_file, scopes)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                GOOGLE_CLIENT_SECRETS_FILE, scopes)
            credentials = flow.run_local_server(port=0)

        with open(credentials_file, 'w') as token:
            token.write(credentials.to_json())

    return credentials

def get_spreadsheet_service() -> Resource:

    credentials = get_credentials(
        GOOGLE_CREDENTIALS_FILE_LOCAL,
        GOOGLE_API_SCOPES)

    if not credentials or not credentials.valid:
        return None

    service = build(
        GOOGLE_API_SERVICE_NAME,
        GOOGLE_API_SERVICE_VERSION,
        credentials = credentials)

    return service

def create_new_spreadsheet(
    service: Resource,
    spreadsheet_name: str) -> str:

    spreadsheet_id = ''

    new_spreadsheet = {
        'properties': {
            'title': spreadsheet_name
        }
    }

    result = service.spreadsheets().create(
        body = new_spreadsheet,
        fields = 'spreadsheetId').execute()

    spreadsheet_id = result.get('spreadsheetId')
    print(f'Created new spreadsheet ({spreadsheet_id}): {create_spreadsheet_url(spreadsheet_id)}')

    return spreadsheet_id

def add_header_row(
    service: Resource,
    spreadsheet_id: str,
    sheet_id: int,
    sheet_name: str ) -> bool:

    sheets = service.spreadsheets()

    # create cells and add text
    values = [SHEET_HEADER_COLUMN_ORDER]
    body = {
        'values': values,
    }

    start_column_int = 0
    end_column_int = len(SHEET_HEADER_COLUMN_ORDER) - 1
    start_column = SHEET_COLUMNS[start_column_int]
    stop_column = SHEET_COLUMNS[end_column_int]
    row = 1 # 1-based
    sheets.values().update(
        spreadsheetId = spreadsheet_id,
        body = body,
        range = f'{sheet_name}!{start_column}{row}:{stop_column}{row + 1}',
        valueInputOption = 'RAW').execute()

    # format cells
    body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': row - 1,
                        'endRowIndex': row
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': 'CENTER',
                            'verticalAlignment': 'MIDDLE',
                            'wrapStrategy': 'WRAP',
                            'textFormat': {
                                'bold': True
                            }
                        }
                    },
                    'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)'
                }
            },
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': sheet_id,
                        'dimension': 'ROWS',
                        'startIndex': 0,
                        'endIndex': 1
                    },
                    'properties': {
                        'pixelSize': 80
                    },
                    'fields': 'pixelSize'
                }
            },
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': sheet_id,
                        'gridProperties': {
                            'frozenRowCount': 2,
                            'frozenColumnCount': 4
                        }
                    },
                    'fields': 'gridProperties(frozenRowCount,frozenColumnCount)'
                }
            },
            {
                'mergeCells': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': 0,
                        'endRowIndex': 1,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER) + 2
                    },
                    'mergeType': 'MERGE_ROWS'
                }
            }
        ]
    }

    sheets.batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = body).execute()

    # initial population information
    first_populated = f'Built: {date.today()} {datetime.now().strftime("%H:%M:%S")}'

    # create cells and add text
    values = [[first_populated]]
    body = {
        'values': values,
    }

    start_column_int = 0
    end_column_int = 1
    start_column = SHEET_COLUMNS[start_column_int]
    stop_column = SHEET_COLUMNS[end_column_int]
    row = 2 # 1-based
    sheets.values().update(
        spreadsheetId = spreadsheet_id,
        body = body,
        range = f'{sheet_name}!{start_column}{row}:{stop_column}{row + 1}',
        valueInputOption = 'RAW').execute()

    # format cells
    body = {
        'requests': [
            {
                'repeatCell': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': row - 1,
                        'endRowIndex': row,
                        'startColumnIndex': 0,
                        'endColumnIndex': 1
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'horizontalAlignment': 'CENTER',
                            'verticalAlignment': 'MIDDLE',
                            'wrapStrategy': 'WRAP',
                            'textFormat': {
                                'bold': True
                            }
                        }
                    },
                    'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)'
                }
            },
            {
                'mergeCells': {
                    'range': {
                        'sheetId': sheet_id,
                        'startRowIndex': row - 1,
                        'endRowIndex': row,
                        'startColumnIndex': 0,
                        'endColumnIndex': 4
                    },
                    'mergeType': 'MERGE_ROWS'
                }
            }
        ]
    }

    sheets.batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = body).execute()

    return True

def add_event_rows(
    service: Resource,
    spreadsheet_id: str,
    sheet_name: str,
    event: SingleEvent,
    row: int,
    update: bool) -> bool:

    sheets = service.spreadsheets()

    if update:
        start_column_int = SHEET_HEADER_COLUMN_ORDER.index(UPDATED)
        end_column_int = SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE_MOVEMENT)
        start_column = SHEET_COLUMNS[start_column_int]
        stop_column = SHEET_COLUMNS[end_column_int]

        starting_spread_column = SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(SPREAD)]
        starting_over_under_column = SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_SPACER)]
        starting_moneyline_column = SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE)]
        latest_spread_column = SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(SPREAD_LATEST)]
        latest_over_under_column = SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_LATEST)]
        latest_moneyline_column = SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE_LATEST)]

        # away team row
        values = [
            [
                event.last_updated,
                event.away_team_spread,
                f'=MINUS({starting_spread_column}{row}, {latest_spread_column}{row})',
                event.over_under,
                f'=MINUS({starting_over_under_column}{row},{latest_over_under_column}{row})',
                event.away_team_moneyline,
                f'=MINUS({starting_moneyline_column}{row},{latest_moneyline_column}{row})'
            ]
        ]

        body = {
            'values': values,
        }

        sheets.values().update(
            spreadsheetId = spreadsheet_id,
            body = body,
            range = f'{sheet_name}!{start_column}{row}:{stop_column}{row}',
            valueInputOption = 'USER_ENTERED').execute()

        # home team row
        values = [
            [
                event.last_updated,
                event.home_team_spread,
                f'=MINUS({starting_spread_column}{row + 1},{latest_spread_column}{row + 1})',
                event.over_under,
                f'=MINUS({starting_over_under_column}{row + 1},{latest_over_under_column}{row + 1})',
                event.home_team_moneyline,
                f'=MINUS({starting_moneyline_column}{row + 1},{latest_moneyline_column}{row + 1})'
            ]
        ]

        body = {
            'values': values,
        }

        sheets.values().update(
            spreadsheetId = spreadsheet_id,
            body = body,
            range = f'{sheet_name}!{start_column}{row + 1}:{stop_column}{row + 1}',
            valueInputOption = 'USER_ENTERED').execute()
    else:
        start_column_int = 0
        end_column_int = SHEET_HEADER_COLUMN_ORDER.index(UPDATED)
        start_column = SHEET_COLUMNS[start_column_int]
        stop_column = SHEET_COLUMNS[end_column_int]

        # away team row
        values = [
            [
                f'{DK_STR_EVENTS_URL}/{event.event_id}',
                event.game_date,
                event.game_time,
                event.away_team,
                event.away_team_spread,
                '', # checkbox for bet spread
                '', # leave blank for last bet spread
                f'O',
                event.over_under,
                '', # checkbox for bet o/u
                '', # leave blank for last bet /u
                event.away_team_moneyline,
                '', # checkbox for bet moneyline
                event.last_updated
            ]
        ]

        body = {
            'values': values,
        }

        sheets.values().update(
            spreadsheetId = spreadsheet_id,
            body = body,
            range = f'{sheet_name}!{start_column}{row}:{stop_column}{row}',
            valueInputOption = 'RAW').execute()

        # home team row
        values = [
            [
                f'{DK_STR_EVENTS_URL}/{event.event_id}',
                event.game_date,
                event.game_time,
                event.home_team,
                event.home_team_spread,
                '', # checkbox for bet spread
                '', # leave blank for last bet spread
                f'U',
                event.over_under,
                '', # checkbox for bet o/u
                '', # leave blank for last bet o/u
                event.home_team_moneyline,
                '', # checkbox for bet moneyline
                event.last_updated
            ]
        ]

        body = {
            'values': values,
        }

        sheets.values().update(
            spreadsheetId = spreadsheet_id,
            body = body,
            range = f'{sheet_name}!{start_column}{row + 1}:{stop_column}{row + 1}',
            valueInputOption = 'RAW').execute()

    return True

def create_merge_cells_request(
    sheet_id: int,
    starting_row: int,
    ending_row: int,
    starting_column: int,
    ending_column: int,
    merge_type: str,
    hyperlink: str) -> List[dict]:

    # rows/columns are assumed to have 0-based indices. ending rows and
    # columns are provided as inclusive and must be incremented when
    # creating the request.
    merge_request = {
        'mergeCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row,
                'endRowIndex': ending_row + 1,
                'startColumnIndex': starting_column,
                'endColumnIndex': ending_column + 1
            },
            'mergeType': merge_type
        }
    }

    format_request = {
        'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row,
                'endRowIndex': ending_row + 1,
                'startColumnIndex': starting_column,
                'endColumnIndex': ending_column + 1
            },
            'cell': {
                'userEnteredFormat': {
                    'horizontalAlignment': 'CENTER',
                    'verticalAlignment': 'MIDDLE',
                    'wrapStrategy': 'WRAP',
                    'textFormat': {
                        'bold': True
                    }
                }
            },
            'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment,wrapStrategy,textFormat)'
        }
    }

    hyperlink_request = {
        'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row,
                'endRowIndex': ending_row + 1,
                'startColumnIndex': starting_column,
                'endColumnIndex': starting_column + 1
            },
            'cell': {
                'userEnteredFormat': {
                    'textFormat': {
                        'link': {
                            'uri': hyperlink
                        }
                    },
                    'hyperlinkDisplayType': 'LINKED'
                }
            },
            'fields': 'userEnteredFormat(hyperlinkDisplayType,textFormat.link)'
        }
    }

    resize_request = {
        'updateDimensionProperties': {
            'range': {
                'sheetId': sheet_id,
                'dimension': 'ROWS',
                'startIndex': starting_row,
                'endIndex': starting_row + 2
            },
            'properties': {
                'pixelSize': 50
            },
            'fields': 'pixelSize'
        }
    }

    return [merge_request, format_request, hyperlink_request, resize_request]

def create_checkbox_request(
    sheet_id: int,
    starting_row: int,
    ending_row: int,
    starting_column: int,
    ending_column: int ) -> List[dict]:

    checkbox_request = {
        'repeatCell': {
            'cell': {
                'dataValidation': {
                    'condition': {
                        'type': 'BOOLEAN'
                    }
                },
                'userEnteredValue': {
                    'boolValue': 'FALSE'
                }
            },
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row,
                'endRowIndex': ending_row,
                'startColumnIndex': starting_column,
                'endColumnIndex': ending_column
            },
            'fields': 'dataValidation,userEnteredValue'
        }
    }

    return [checkbox_request]

def get_text_format_string(
    bold: bool,
    italic: bool) -> dict:

    if not bold and not italic:
        return {}

    val = {}
    if bold:
        val['bold'] = True
    if italic:
        val['italic'] = True

    return val

def create_format_value_column_request(
    sheet_id: int,
    starting_row: int,
    starting_column: int,
    ending_column: int,
    bold: bool,
    italic: bool,
    show_plus_minus: bool ) -> List[dict]:

    request = {
        'repeatCell': {
            'cell': {
                'userEnteredFormat': {
                    'textFormat':
                        get_text_format_string(bold, italic)
                    ,
                    'numberFormat': {
                        'type': 'NUMBER'
                    }
                }
            },
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row,
                'startColumnIndex': starting_column,
                'endColumnIndex': ending_column
            },
            'fields': 'userEnteredFormat(textFormat,numberFormat)'
        }
    }

    if show_plus_minus:
        request['repeatCell']['cell']['userEnteredFormat']['numberFormat']['pattern'] = '"+"0.0;"-"0.0;0'

    return [request]

def create_format_starting_value_column_request(
    sheet_id: int,
    starting_row: int,
    starting_column: int,
    ending_column: int,
    show_plus_minus: bool ) -> List[dict]:

    bold = True
    italic = False

    return create_format_value_column_request(
        sheet_id,
        starting_row,
        starting_column,
        ending_column,
        bold,
        italic,
        show_plus_minus)

def create_format_movement_column_request(
    sheet_id: int,
    starting_row: int,
    starting_column: int,
    ending_column: int ) -> List[dict]:

    bold = False
    italic = True
    show_plus_minus = True

    return create_format_value_column_request(
        sheet_id,
        starting_row,
        starting_column,
        ending_column,
        bold,
        italic,
        show_plus_minus)

def create_format_team_colors_request(
    sheet_id: int,
    event: SingleEvent,
    starting_row: int) -> List[dict]:

    # away colors
    away_bg_hex = tc.team_colors[event.away_team][0] if event.away_team in tc.team_colors else '#FFFFFF'
    if not away_bg_hex:
        away_bg_hex = '#FFFFFF'
    away_bg_rgb = hex_to_rgb(away_bg_hex)
    away_fg_hex = tc.team_colors[event.away_team][1] if event.away_team in tc.team_colors else '#000000'
    if not away_fg_hex:
        away_fg_hex = '#000000'
    away_fg_rgb = hex_to_rgb(away_fg_hex)

    # home colors
    home_bg_hex = tc.team_colors[event.home_team][0] if event.home_team in tc.team_colors else '#FFFFFF'
    if not home_bg_hex:
        home_bg_hex = '#FFFFFF'
    home_bg_rgb = hex_to_rgb(home_bg_hex)
    home_fg_hex = tc.team_colors[event.home_team][1] if event.home_team in tc.team_colors else '#000000'
    if not home_fg_hex:
        home_fg_hex = '#000000'
    home_fg_rgb = hex_to_rgb(home_fg_hex)

    home_team_request = {
        'updateCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row - 1,
                'endRowIndex': starting_row,
                'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(MATCHUP),
                'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(MATCHUP) + 1
            },
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredFormat': {
                                'backgroundColor': {
                                    'red': away_bg_rgb[0],
                                    'green': away_bg_rgb[1],
                                    'blue': away_bg_rgb[2]
                                },
                                'textFormat': {
                                    'foregroundColor': {
                                        'red': away_fg_rgb[0],
                                        'green': away_fg_rgb[1],
                                        'blue': away_fg_rgb[2],
                                    },
                                    'bold': True
                                }
                            }
                        }
                    ]
                }
            ],
            'fields': 'userEnteredFormat(backgroundColor,textFormat)'
        }
    }

    away_team_request = {
        'updateCells': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row,
                'endRowIndex': starting_row + 1,
                'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(MATCHUP),
                'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(MATCHUP) + 1
            },
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredFormat': {
                                'backgroundColor': {
                                    'red': home_bg_rgb[0],
                                    'green': home_bg_rgb[1],
                                    'blue': home_bg_rgb[2]
                                },
                                'textFormat': {
                                    'foregroundColor': {
                                        'red': home_fg_rgb[0],
                                        'green': home_fg_rgb[1],
                                        'blue': home_fg_rgb[2],
                                    },
                                    'bold': True
                                }
                            }
                        }
                    ]
                }
            ],
            'fields': 'userEnteredFormat(backgroundColor,textFormat)'
        }
    }

    return [away_team_request, home_team_request]

def create_format_event_alignment_request(
    sheet_id: int,
    starting_row: int) -> List[dict]:

    format_request = {
        'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row - 1,
                'endRowIndex': starting_row + 1,
            },
            'cell': {
                'userEnteredFormat': {
                    'horizontalAlignment': 'CENTER',
                    'verticalAlignment': 'MIDDLE'
                }
            },
            'fields': 'userEnteredFormat(horizontalAlignment,verticalAlignment)'
        }
    }

    return [format_request]

def create_format_borders_request(
    sheet_id: int,
    starting_row: int) -> List[dict]:

    thick_border = {
        'style': 'SOLID_THICK',
        'color': {
            'red': 0,
            'green': 0,
            'blue': 0
        }
    }

    spread_request = {
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row - 1,
                'endRowIndex': starting_row + 1,
                'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(SPREAD),
                'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(LAST_BET_SPREAD) + 1
            },
            'top': thick_border,
            'bottom': thick_border,
            'left': thick_border,
            'right': thick_border
        }
    }

    over_under_request = {
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row - 1,
                'endRowIndex': starting_row + 1,
                'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER),
                'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(LAST_BET_OVER_UNDER) + 1
            },
            'top': thick_border,
            'bottom': thick_border,
            'left': thick_border,
            'right': thick_border
        }
    }

    moneyline_request = {
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row - 1,
                'endRowIndex': starting_row + 1,
                'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE),
                'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(BET_MONEYLINE) + 1
            },
            'top': thick_border,
            'bottom': thick_border,
            'left': thick_border,
            'right': thick_border
        }
    }

    latest_spread_request = {
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row - 1,
                'endRowIndex': starting_row + 1,
                'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(SPREAD_LATEST),
                'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(SPREAD_MOVEMENT) + 1
            },
            'top': thick_border,
            'bottom': thick_border,
            'left': thick_border,
            'right': thick_border
        }
    }

    latest_over_under_request = {
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row - 1,
                'endRowIndex': starting_row + 1,
                'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_LATEST),
                'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT) + 1
            },
            'top': thick_border,
            'bottom': thick_border,
            'left': thick_border,
            'right': thick_border
        }
    }

    latest_moneyline_request = {
        'updateBorders': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': starting_row - 1,
                'endRowIndex': starting_row + 1,
                'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE_LATEST),
                'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE_MOVEMENT) + 1
            },
            'top': thick_border,
            'bottom': thick_border,
            'left': thick_border,
            'right': thick_border
        }
    }

    return [spread_request, over_under_request, moneyline_request, latest_spread_request, latest_over_under_request, latest_moneyline_request]

def create_format_obsolete_event_request(
    sheet_id: int,
    row: int) -> List[dict]:

    background_color = hex_to_rgb('#888888')

    obsolete_request = {
        'repeatCell': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': row,
                'endRowIndex': row + 2,
                'startColumnIndex': 0,
                'endColumnIndex': len(SHEET_HEADER_COLUMN_ORDER)
            },
            'cell': {
                'userEnteredFormat': {
                    'backgroundColor': {
                        'red': background_color[0],
                        'green': background_color[1],
                        'blue': background_color[2]
                    },
                    'textFormat': {
                        'strikethrough': True
                    }
                }
            },
            'fields': 'userEnteredFormat(backgroundColor,textFormat)'
        }
    }

    return [obsolete_request]

def create_format_auto_column_width_request(
    sheet_id: int) -> List[dict]:

    auto_width_request = {
        'autoResizeDimensions': {
            'dimensions': {
                'sheetId': sheet_id,
                'dimension': 'COLUMNS',
                'startIndex': 0,
                'endIndex': len(SHEET_HEADER_COLUMN_ORDER) + 1
            }
        }
    }

    return [auto_width_request]

def create_conditional_formatting_rules_request(
    sheet_id: int) -> List[dict]:

    good_color = hex_to_rgb('#00FF00')
    bad_color = hex_to_rgb('#FF0000')
    index = 0

    format_spread_request = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(SPREAD),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(SPREAD) + 1
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_SPREAD)]}3)"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': good_color[0],
                            'green': good_color[1],
                            'blue': good_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    format_over_under_request_1 = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER) + 1
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_OVER_UNDER)]}3)"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': good_color[0],
                            'green': good_color[1],
                            'blue': good_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    format_over_under_request_2 = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER) + 1,
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER) + 2
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_OVER_UNDER)]}3)"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': good_color[0],
                            'green': good_color[1],
                            'blue': good_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    format_moneyline_request = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE) + 1
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_MONEYLINE)]}3)"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': good_color[0],
                            'green': good_color[1],
                            'blue': good_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    format_spread_movement_request_good = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(SPREAD_MOVEMENT),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(SPREAD_MOVEMENT) + 1
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                #'=AND(F3,GT(N3,0),NOT(ISBLANK(M3)))'
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_SPREAD)]}3,NOT(ISBLANK({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(SPREAD_LATEST)]}3)),GT({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(SPREAD_MOVEMENT)]}3,0))"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': good_color[0],
                            'green': good_color[1],
                            'blue': good_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    format_spread_movement_request_bad = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(SPREAD_MOVEMENT),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(SPREAD_MOVEMENT) + 1
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_SPREAD)]}3,NOT(ISBLANK({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(SPREAD_LATEST)]}3)),LT({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(SPREAD_MOVEMENT)]}3,0))"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': bad_color[0],
                            'green': bad_color[1],
                            'blue': bad_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    format_over_movement_request_good = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT) + 1
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_OVER_UNDER)]}3,EQ({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER)]}3,\"O\"),NOT(ISBLANK({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_LATEST)]}3)),LT({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT)]}3,0))"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': good_color[0],
                            'green': good_color[1],
                            'blue': good_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    format_over_movement_request_bad = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT) + 1
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_OVER_UNDER)]}3,EQ({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER)]}3,\"O\"),NOT(ISBLANK({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_LATEST)]}3)),GT({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT)]}3,0))"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': bad_color[0],
                            'green': bad_color[1],
                            'blue': bad_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    format_under_movement_request_good = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT) + 1
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_OVER_UNDER)]}3,EQ({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER)]}3,\"U\"),NOT(ISBLANK({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_LATEST)]}3)),GT({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT)]}3,0))"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': good_color[0],
                            'green': good_color[1],
                            'blue': good_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    format_under_movement_request_bad = {
        'addConditionalFormatRule': {
            'rule': {
                'ranges': [
                    {
                        'sheetId': sheet_id,
                        'startRowIndex': 2,
                        'startColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT),
                        'endColumnIndex': SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT) + 1
                    }
                ],
                'booleanRule': {
                    'condition': {
                        'type': 'CUSTOM_FORMULA',
                        'values': [
                            {
                                'userEnteredValue': f"=AND({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(BET_OVER_UNDER)]}3,EQ({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER)]}3,\"U\"),NOT(ISBLANK({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_LATEST)]}3)),LT({SHEET_COLUMNS[SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT)]}3,0))"
                            }
                        ]
                    },
                    'format': {
                        'backgroundColor': {
                            'red': bad_color[0],
                            'green': bad_color[1],
                            'blue': bad_color[2],
                        }
                    }
                }
            },
            'index': index
        }
    }
    index +=1

    return [
        format_spread_request,
        format_over_under_request_1,
        format_over_under_request_2,
        format_moneyline_request,
        format_spread_movement_request_good,
        format_spread_movement_request_bad,
        format_over_movement_request_good,
        format_over_movement_request_bad,
        format_under_movement_request_good,
        format_under_movement_request_bad]

def create_add_sheet_request(
    sheet_id: int,
    sheet_name: str ) -> List[dict]:

    add_sheet_request = {
        'addSheet': {
            'properties': {
                'sheetId': sheet_id,
                'title': sheet_name
            }
        }
    }

    return [add_sheet_request]

def create_rename_sheet_request(
    sheet_id: int,
    sheet_name: str) -> List[dict]:

    rename_sheet_request = {
        'updateSheetProperties': {
            'properties': {
                'sheetId': sheet_id,
                'title': sheet_name
            },
            'fields': 'title'
        }
    }

    return [rename_sheet_request]

def format_event_rows(
    service: Resource,
    spreadsheet_id: str,
    sheet_id: int,
    event: SingleEvent,
    starting_row: int,
    update: bool) -> bool:

    sheets = service.spreadsheets()

    requests = []

    # merge rows for columns A-C
    if not update:
        merge_request = create_merge_cells_request(
            sheet_id,
            starting_row - 1,
            starting_row,
            0,
            2,
            'MERGE_COLUMNS',
            event.create_event_url())

        for request in merge_request:
            requests.append(request)

    # create checkboxes for betting choices in columns 5,8,10
    if not update:
        indices = [
            SHEET_HEADER_COLUMN_ORDER.index(BET_SPREAD),
            SHEET_HEADER_COLUMN_ORDER.index(BET_OVER_UNDER),
            SHEET_HEADER_COLUMN_ORDER.index(BET_MONEYLINE)
        ]

        for i in indices:
            checkbox_request = create_checkbox_request(
                sheet_id,
                starting_row - 1,
                starting_row + 1,
                i,
                i + 1)

            for request in checkbox_request:
                requests.append(request)

    # add foreground/background colors to team name cells
    if not update:
        format_team_colors_request = create_format_team_colors_request(
            sheet_id,
            event,
            starting_row)

        for request in format_team_colors_request:
            requests.append(request)

    # middle and center align all cells
    if not update:
        alignment_request = create_format_event_alignment_request(
            sheet_id,
            starting_row)

        for request in alignment_request:
            requests.append(request)

    # add borders around spread, over/under, and moneyline cells and
    # their corresponding bet checkboxes
    #if not update:
    format_borders_request = create_format_borders_request(
    sheet_id,
    starting_row)
    #jmd: end if not update

    for request in format_borders_request:
        requests.append(request)

    body = {
        'requests': requests
    }

    if requests:
        sheets.batchUpdate(
            spreadsheetId = spreadsheet_id,
            body = body).execute()

    return True

def format_all_value_columns(
    service: Resource,
    spreadsheet_id: str,
    sheet_id: int) -> bool:

    sheets = service.spreadsheets()

    requests = []
    starting_row = 3 # maybe?

    # format starting spread, moneyline
    # also apply this for the updated spread and moneyline columns
    indices = [
        SHEET_HEADER_COLUMN_ORDER.index(SPREAD),
        SHEET_HEADER_COLUMN_ORDER.index(LAST_BET_SPREAD),
        SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE),
        SHEET_HEADER_COLUMN_ORDER.index(SPREAD_LATEST),
        SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE_LATEST)
    ]
    for i in indices:
        format_values = create_format_starting_value_column_request(
            sheet_id,
            starting_row - 1,
            i,
            i + 1,
            True)

        for request in format_values:
            requests.append(request)

    # format o/u f
    # also apply this for the updated o/u column
    indices = [
        SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER),
        SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_SPACER),
        SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_LATEST)
    ]
    for i in indices:
        format_values = create_format_starting_value_column_request(
            sheet_id,
            starting_row - 1,
            i,
            i + 1,
            False)

        for request in format_values:
            requests.append(request)

    # format line movement columns 13,15,17
    indices = [
        SHEET_HEADER_COLUMN_ORDER.index(SPREAD_MOVEMENT),
        SHEET_HEADER_COLUMN_ORDER.index(OVER_UNDER_MOVEMENT),
        SHEET_HEADER_COLUMN_ORDER.index(MONEYLINE_MOVEMENT)
    ]
    for i in indices:
        format_movement = create_format_movement_column_request(
            sheet_id,
            starting_row - 1,
            i,
            i + 1)

        for request in format_movement:
            requests.append(request)

    body = {
        'requests': requests
    }

    sheets.batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = body).execute()

    return True

def format_auto_column_width(
    service: Resource,
    spreadsheet_id: str,
    sheet_id: int) -> bool:

    sheets = service.spreadsheets()

    requests = []

    auto_width_request = create_format_auto_column_width_request(
        sheet_id)

    for request in auto_width_request:
        requests.append(request)

    body = {
        'requests': auto_width_request
    }

    sheets.batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = body).execute()

    return

def format_conditional_formatting(
    service: Resource,
    spreadsheet_id: str,
    sheet_id: int) -> bool:

    sheets = service.spreadsheets()

    requests = []
    format_conditional = create_conditional_formatting_rules_request(
        sheet_id)

    for request in format_conditional:
        requests.append(request)

    body = {
        'requests': requests
    }

    sheets.batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = body).execute()

    return

def format_obsolete_event(
    service: Resource,
    spreadsheet_id: str,
    sheet_id: int,
    row: int) -> bool:

    sheets = service.spreadsheets()

    requests = []
    format_obsolete = create_format_obsolete_event_request(
        sheet_id,
        row)

    for request in format_obsolete:
        requests.append(request)

    body = {
        'requests': requests
    }

    sheets.batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = body).execute()

    return

def format_sheet(
    service: Resource,
    spreadsheet_id: str,
    sheet_id: int,
    update: bool) -> bool:

    if not update:
        format_conditional_formatting(
            service,
            spreadsheet_id,
            sheet_id)

        format_all_value_columns(
            service,
            spreadsheet_id,
            sheet_id)
    #jmd: end if not update block

    format_auto_column_width(
        service,
        spreadsheet_id,
        sheet_id)

    return True

def create_sheet(
    service: Resource,
    spreadsheet_id: str,
    sheet_id: int,
    sheet_name: str) -> None:

    sheets = service.spreadsheets()

    requests = []

    add = create_add_sheet_request(
        sheet_id,
        sheet_name)

    for request in add:
        requests.append(request)

    body = {
        'requests': requests
    }

    sheets.batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = body).execute()

    return

def rename_sheet(
    service: Resource,
    spreadsheet_id: str,
    sheet_id: int,
    sheet_name: str) -> None:

    sheets = service.spreadsheets()

    requests = []

    rename = create_rename_sheet_request(
        sheet_id,
        sheet_name)

    for request in rename:
        requests.append(request)

    body = {
        'requests': requests
    }

    sheets.batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = body).execute()

    return

def create_new_spreadsheet_from_events(
    title: str,
    all_events: List[List[SingleEvent]]) -> str:

    service = get_spreadsheet_service()
    if not service or len(all_events) == 0:
        return ''

    spreadsheet_id = create_new_spreadsheet(
        service,
        title)

    sheet_id = 0
    for events in all_events:
        if not len(events):
            sheet_id += 1
            continue

        sheet_name = events[0].sheet_name

        if sheet_id == 0:
            rename_sheet(
                service,
                spreadsheet_id,
                sheet_id,
                sheet_name)
        else:
            create_sheet(
                service,
                spreadsheet_id,
                sheet_id,
                sheet_name)

        add_header_row(
            service,
            spreadsheet_id,
            sheet_id,
            sheet_name)

        row = 3
        update = False

        for event in events:
            event.print()
            add_event_rows(
                service,
                spreadsheet_id,
                sheet_name,
                event,
                row,
                update)

            format_event_rows(
                service,
                spreadsheet_id,
                sheet_id,
                event,
                row,
                update)

            print(f'Added new entry for {event.away_team} @ {event.home_team}')
            row += 3

            time.sleep(SLEEP_TIME)

        format_sheet(
            service,
            spreadsheet_id,
            sheet_id,
            update
        )

        sheet_id += 1

    print(f'Spreadsheet has been created and is available at: {create_spreadsheet_url(spreadsheet_id)}')
    print(f'To update: dk.py --update {spreadsheet_id}')

    return spreadsheet_id

def update_spreadsheet_from_events(
    spreadsheet_id: str,
    all_events: List[List[SingleEvent]]) -> bool:

    service = get_spreadsheet_service()
    if not service or len(all_events) == 0:
        return False

    sheet_id = 0
    for events in all_events:
        if not len(events):
            sheet_id += 1
            continue

        sheet_name = events[0].sheet_name

        event_ids = get_event_ids_from_sheet(
            service,
            spreadsheet_id,
            sheet_name)

        num_rows = get_number_of_rows(
            service,
            spreadsheet_id,
            sheet_name) + 2

        update = True

        for event in events:
            event.print()

            if event.event_id in event_ids:
                row = event_ids[event.event_id] + 1 # i hate everything

                add_event_rows(
                    service,
                    spreadsheet_id,
                    sheet_name,
                    event,
                    row,
                    update)

                format_event_rows(
                    service,
                    spreadsheet_id,
                    sheet_id,
                    event,
                    row,
                    update)

                print(f'Updated data for {event.away_team} @ {event.home_team}')
                del event_ids[event.event_id]

            else:
                row = num_rows
                num_rows += 3

                add_event_rows(
                    service,
                    spreadsheet_id,
                    sheet_name,
                    event,
                    row,
                    False)

                format_event_rows(
                    service,
                    spreadsheet_id,
                    sheet_id,
                    event,
                    row)

                print(f'Updated data for {event.away_team} @ {event.home_team}')

            time.sleep(SLEEP_TIME)

        for event_id in event_ids:
            format_obsolete_event(
                service,
                spreadsheet_id,
                sheet_id,
                event_ids[event_id])

        format_sheet(
            service,
            spreadsheet_id,
            sheet_id,
            update
        )

    print(f'Spreadsheet has been updated and is available at: {create_spreadsheet_url(spreadsheet_id)}')
    print(f'To update again: dk.py --update {spreadsheet_id}')

    return True

def get_event_ids_from_sheet(
    service: Resource,
    spreadsheet_id: str,
    sheet_name: str) -> dict:

    requested_range = f'{sheet_name}!A:A'
    request = service.spreadsheets().values().get(
        spreadsheetId = spreadsheet_id,
        range = requested_range)
    response = request.execute()

    event_ids = {}
    row_index = 2
    values = response['values'][2:]
    for row in values:
        if len(row) > 0:
            event_id = row[0].rsplit('/', 1)[1]
            event_ids[event_id] = row_index
        row_index += 1

    return event_ids

def get_number_of_rows(
    service: Resource,
    spreadsheet_id: str,
    sheet_name: str) -> int:

    requested_range = f'{sheet_name}!A:A'
    request = service.spreadsheets().values().get(
        spreadsheetId = spreadsheet_id,
        range = requested_range)
    response = request.execute()

    return len(response['values'])
