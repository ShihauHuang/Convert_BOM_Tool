import coloredlogs, logging, datetime
from os import mkdir, path

def Create_Logger() :

    if not path.exists('log'): # 建立資料夾存放轉換後的檔案
        mkdir('log')

    logger = logging.getLogger()
    coloredlogs.install(
        level="INFO", 
        fmt = '[%(asctime)s %(module)s:%(lineno)d] %(message)s'
    )

    log_filename = datetime.datetime.now().strftime("%Y-%m-%d_%H_%M_%S.log")
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter(
        '[%(asctime)s %(module)s:%(lineno)d] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S')
    fh = logging.FileHandler('log/' + log_filename, encoding='utf-8-sig')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(formatter)
    logger.addHandler(fh)