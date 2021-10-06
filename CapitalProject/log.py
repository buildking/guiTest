import logging
import os
from config import ConfigUtil

def setLogging(_logName="log"):
    log_level = ConfigUtil.config.get("log", 'log_level', fallback='INFO')
    log_dir = ConfigUtil.config.get("log", 'log_dir', fallback='log')
    log_file = ConfigUtil.config.get("log", 'log_file', fallback='project.log')

    # log 디렉토리가 없으면 생성한다.
    try:
        os.makedirs(log_dir)
    except OSError:
        if not os.path.isdir(log_dir):
            raise

    logger = logging.getLogger(_logName)

    if log_level == 'INFO':
        logger.setLevel(logging.INFO)
    elif log_level == 'DEBUG':
        logger.setLevel(logging.DEBUG)
    elif log_level == 'ERROR':
        logger.setLevel(logging.ERROR)
    elif log_level == 'WARNING':
        logger.setLevel(logging.WARNING)
    else:
        logger.setLevel(logging.INFO)

    handler = logging.StreamHandler()

    formatter = logging.Formatter('[%(asctime)s %(name)s:%(lineno)d] %(levelname)-8s %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    file_handler = logging.FileHandler('./' + log_dir + '/' + log_file, mode='a')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    logger.debug("logger module is ready")
    return logger

#logger = setLogging("main")
#logger.info("test")