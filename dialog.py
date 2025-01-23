"""
日志
"""
import logging
import inspect

from tkinter import messagebox as mb

class Dialog:
    DEBUG = logging.DEBUG
    INFO = logging.INFO
    WARNING = logging.WARNING
    ERROR = logging.ERROR
    CRITICAL = logging.CRITICAL

    def __init__(self, dialogFileName: str, dialogFormat: str, dateFormat: str, dialogLevel: int, dialogEncoding: str):
        """
        日志记录器初始化

        Params:
            dialogFileName: 日志文件名
            dialogFormat: 日志格式
            dateFormat: 日期格式
            dialogLevel: 日志级别
            dialogEncoding: 日志文件编码
        """
        logging.basicConfig(
            filename=dialogFileName,
            format=dialogFormat,
            datefmt=dateFormat,
            level=dialogLevel,
            encoding=dialogEncoding
        )

    @staticmethod
    def log(message: str, dialogLevel: int = DEBUG):
        """
        写日志

        Params:
            dialogLevel: 日志类型
            message: 日志信息
        """
        if not logging.root.hasHandlers():
            mb.showerror('错误', '日志记录器未初始化，请初始化后使用！')
            return
        callerFunctionName = inspect.currentframe().f_back.f_code.co_name # 获取调用者函数名
        message = f'{callerFunctionName} -> {message}'
        logging.log(dialogLevel, message)
        match  dialogLevel:
            case Dialog.DEBUG:
                pass
            case Dialog.INFO:
                mb.showinfo('信息', message)
            case Dialog.WARNING:
                mb.showwarning('警告', message)
            case Dialog.ERROR:
                mb.showerror('错误', message)
            case Dialog.CRITICAL:
                mb.showerror('严重错误', message)