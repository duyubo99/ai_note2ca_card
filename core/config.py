'''
@Project ：code 
@File    ：config.py
@Author  ：Sito
@Date    ：2025/3/5 14:39 
@Description    ：
'''
import os
import sys
import logging

CHUNK_SIZE = 10000


def logger_configuration(task='server'):
    '''
        配置日志 handler
    Returns: logger
    '''
    logger = logging.getLogger(f'{task}-log')
    logger.setLevel(logging.INFO)

    # 禁用父级日志记录器的传播
    logger.propagate = False

    # stream handler
    rf_handler = logging.StreamHandler(sys.stderr)
    rf_handler.setLevel(logging.INFO)
    rf_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(message)s"))

    # file handler
    # if not os.path.isdir('log'):
    #     os.mkdir('log')
    # f_handler = logging.FileHandler(f'log/{task}.log', encoding='utf-8')
    # f_handler.setLevel(logging.INFO)
    # f_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(filename)s[:%(lineno)d] - %(message)s"))

    # 添加处理器前检查是否已经存在
    if not any(isinstance(h, logging.StreamHandler) for h in logger.handlers):
        logger.addHandler(rf_handler)

    # if not any(isinstance(h, logging.FileHandler) for h in logger.handlers):
    #     logger.addHandler(f_handler)

    return logger
