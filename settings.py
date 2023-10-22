
import sys
import configparser
from pathlib import Path


IS_POSIX = sys.platform.startswith(("darwin", "cygwin", "linux", "linux2"))

TRUTHY = [
    'TRUE', 'True', 'true', 'T', 't', True,
    '1', 1,
    'ON', 'On', 'on',
    'YES', 'Yes', 'yes', 'Y', 'y'
]


FALSY = [
    'FALSE', 'False', 'false', 'F', 'f', False,
    '0', 0,
    'OFF', 'Off', 'off',
    'NO', 'No', 'no', 'N', 'n'
]

BASE_DIR = Path("./")
CONFIGURATION_FILE = BASE_DIR / 'settings.config'
CHROME_DIR = BASE_DIR / "chrome"
USER_DATA_DIR = CHROME_DIR / "user-data"
CONTACT_FOLDER_PATH = BASE_DIR / "contacts"

CHROME_DIR.mkdir(exist_ok=True)
USER_DATA_DIR.mkdir(mode=777, exist_ok=True)
CONTACT_FOLDER_PATH.mkdir(exist_ok=True)

config = configparser.RawConfigParser()
config.read(CONFIGURATION_FILE)
SETTINGS = dict(config.items('settings'))

PROJECT_NAME = SETTINGS.get('project_name').strip()
PROJECT_DESCRIPTION = SETTINGS.get('project_description').strip()
VERSION = SETTINGS.get('version').strip()
SITE_DOMAIN = SETTINGS.get('site_domain').strip()
LOGIN_URL = SETTINGS.get('login_url').strip()
LOGIN_TITLE = SETTINGS.get('login_title').strip()
LOGIN_REDIRECT_TITLE = SETTINGS.get('login_redirect_title').strip()
SEND_URL = SETTINGS.get('send_url').strip()
CONTACT_NUMBER_COLUMN_NAME = SETTINGS.get('contact_number_column_name').strip()
CONTACT_NAME_COLUMN_NAME = SETTINGS.get('contact_name_column_name').strip()
IMAGE_NAME_COLUMN_NAME = SETTINGS.get('image_name_column_name').strip()
STATUS_COLUMN_NAME = SETTINGS.get('status_column_name').strip()
MESSAGE_BODY_FILE = BASE_DIR / 'message' / SETTINGS.get('message_body_file').strip()
MESSAGE_IMAGE_DIR = BASE_DIR / 'message' / SETTINGS.get('message_image_folder').strip()
MESSAGE_LINK_RENDERED_TITLE = SETTINGS.get('message_link_rendered_title').strip()
COUNTRY_CODE = SETTINGS.get('country_code').strip()

HEALTH_CHECK_URL = SETTINGS.get('health_check_url').strip()
HEALTH_CHECK_TITLE = SETTINGS.get('health_check_title').strip()
ACTIVE_HOURS = float(SETTINGS.get('active_hours').strip())
PROCESSING_DURATION_EARCH_ROWS_SECS = float(SETTINGS.get('processing_duration_earch_rows_secs').strip())
NUMBER_OF_ROWS_TO_PROCESS = ACTIVE_HOURS*3600/PROCESSING_DURATION_EARCH_ROWS_SECS
