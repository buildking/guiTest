import logging
import os

def setLogging(_logName="log"):

    # log 디렉토리가 없으면 생성한다.
    try:
        os.makedirs('./log')
    except OSError:
        if not os.path.isdir('./log'):
            raise

    logger = logging.getLogger(_logName)
    logger.setLevel(logging.INFO)

    handler = logging.StreamHandler()

    formatter = logging.Formatter('[%(asctime)s %(name)s:%(lineno)d] %(levelname)-8s %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    file_handler = logging.FileHandler('./log/project.log')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)


    print('logging is ready')
    return logger