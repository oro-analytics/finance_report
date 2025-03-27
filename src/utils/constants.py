import datetime
import os
from src.utils.service_utils import check_path_exist


# Настройка путей
#
CONFIG_BOT_DIR = os.path.dirname(__file__)

#
REPORT_PATH = os.path.abspath(os.path.join(CONFIG_BOT_DIR, '../../report'))
check_path_exist(REPORT_PATH)

# logging
LOG_FILE = os.path.abspath(
    os.path.join(CONFIG_BOT_DIR, '../..', 'log', 'finance_report-{:%Y-%m-%d-%H-%M-%S}.log'.format(datetime.datetime.now())))
LOGGER_NAME = ''
NULL_LOGGER_NAME = 'null'
check_path_exist(os.path.split(LOG_FILE)[0])

