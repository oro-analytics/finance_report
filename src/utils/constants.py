import datetime
import os
from src.utils.service_utils import check_path_exist


# Настройка путей
#
CONFIG_BOT_DIR = os.path.dirname(__file__)
date_stamp = '{:%Y-%m-%d-%H-%M-%S}'.format(datetime.datetime.now())

#
REPORT_PATH = os.path.abspath(os.path.join(CONFIG_BOT_DIR, '../../report', f'finance_report-{date_stamp}.xlsx'))
check_path_exist(os.path.split(REPORT_PATH)[0])

# logging
LOG_FILE = os.path.abspath(
    os.path.join(CONFIG_BOT_DIR, '../..', 'log', f'finance_report-{date_stamp}.log'))
LOGGER_NAME = ''
NULL_LOGGER_NAME = 'null'
check_path_exist(os.path.split(LOG_FILE)[0])

