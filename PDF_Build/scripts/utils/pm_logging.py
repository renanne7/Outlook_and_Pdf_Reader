import logging.handlers
import sys
import time
from logging import StreamHandler

from scripts.constants import pm_configurations as app_conf
from scripts.utils import logger_util

__LOG_FILE = app_conf.FILE_NAME
__LOG_FILE_JSON = app_conf.FILE_NAME_JSON
__LOG_HANDLERS = app_conf.LOG_HANDLERS

__logger = logging.getLogger("HS_Build_Logs")
__logger.setLevel(app_conf.LOG_LEVEL)

__formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s  - %(filename)s - %(module)s: %(funcName)s: '
                                '%(lineno)d - %(message)s')

if 'file' in __LOG_HANDLERS:
    # Adding the log file handler to the logger
    __file_handler = logging.FileHandler(__LOG_FILE + "." + time.strftime("%Y-%m-%d"))
    __file_handler.setFormatter(__formatter)
    __logger.addHandler(__file_handler)

if 'jsonFile' in __LOG_HANDLERS:
    # Adding the log Json file handler to the logger
    __json_handler = logging.FileHandler(__LOG_FILE_JSON + "." + time.strftime("%Y-%m-%d"))
    # Calling custom formatter
    __json_formatter = logger_util.JSONFormatter(
        ['asctime', 'name', 'levelname', 'filename', 'module', 'funcName', 'lineno'])

    __json_handler.setFormatter(__json_formatter)
    __logger.addHandler(__json_handler)

if 'console' in __LOG_HANDLERS:
    # Adding the log Console handler to the logger
    __console_handler = StreamHandler(sys.stdout)
    __console_handler.setFormatter(__formatter)
    __logger.addHandler(__console_handler)


def get_logger():
    return __logger
