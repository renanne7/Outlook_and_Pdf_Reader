import configparser

CONFIGURATION_FILE = r"conf/application.conf"

config = configparser.RawConfigParser()
config.read(CONFIGURATION_FILE)

# HS Build Details
HS_BUILD_VERSION = config.get('HS_BUILD_DETAILS', 'version')

# PM Logging :
LOG_BASE_PATH = config.get('LOGGER', 'base_path')
LOG_LEVEL = config.get('LOGGER', 'log_level')
FILE_BACKUP_WHEN = config.get('LOGGER', 'file_backup_when')
FILE_BACKUP_INTERVAL = config.getint('LOGGER', 'file_backup_interval')
FILE_BACKUP_COUNT = config.getint('LOGGER', 'file_backup_count')
FILE_NAME = LOG_BASE_PATH + config.get('LOGGER', 'file_name')
FILE_NAME_JSON = LOG_BASE_PATH + config.get('LOGGER', 'file_name_json')
LOG_HANDLERS = config.get('LOGGER', 'handlers')
