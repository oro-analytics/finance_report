import logging
import os
import shutil
import sys
from logging.handlers import TimedRotatingFileHandler


def check_path_exist(dir_):
    #dir_ = os.path.split(filename)[0]
    # create all nested directories
    if not os.path.isdir(dir_):
        os.makedirs(dir_)


def setup_logger(log_file, logger_name='', level=logging.DEBUG):
    """ Config path for messages of any level

    :param log_file:
    :param logger_name:
    :param level: обязательно logging.DEBUG, что бы включать logging входящих библиотек
    :return:
    """
    check_path_exist(os.path.split(log_file)[0])
    print(f"ADD LOGGING (LOGGER_NAME={logger_name}) to: {log_file}")
    handler = TimedRotatingFileHandler(log_file,
                                       when="midnight",
                                       interval=1,
                                       encoding='utf-8',
                                       backupCount=0)
    handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    logger_internal = logging.getLogger(logger_name)
    logger_internal.setLevel(level)
    if logger_internal.hasHandlers():
        logger_internal.handlers.clear()
    logger_internal.addHandler(handler)

    # В stdout
    sh = logging.StreamHandler(stream=sys.stdout)
    sh.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
    #sh.setFormatter(formatter)
    logger_internal.addHandler(sh)
    return logger_internal
