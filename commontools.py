import os
import sys
import pathlib
import datetime

class CommonTools:
    @staticmethod
    def getAbsPath(path: str) -> str:
        """
        获取绝对路径

        Params:
            path (str): 路径

        Returns:
            str: 绝对路径
        """
        return os.path.abspath(path)
    
    @staticmethod
    def getPackagedStatus():
        """
        获取打包状态

        Returns:
            bool: 是否为打包状态
        """
        return getattr(sys, 'frozen', False)

    @staticmethod
    def getCurrentPath() -> str:
        """
        获取程序所在目录

        Returns:
            str: 程序所在目录
        """
        if CommonTools.getPackagedStatus():
            return os.path.dirname(CommonTools.getAbsPath(sys.executable))
        else:
            return os.path.dirname(CommonTools.getAbsPath(__file__))
        
    @staticmethod
    def getNameFromPath(path: str) -> str:
        """
        从路径中获取文件名

        Params:
            path (str): 文件名
        """
        return os.path.basename(path)
    
    @staticmethod
    def removeFileExtension(path: str) -> str:
        """
        移除文件扩展名
        """
        return pathlib.Path(path).stem

    @staticmethod
    def getFilesFromFolder(folder: str, exts: tuple) -> list[str]:
        """
        遍历文件夹，获取指定扩展名的文件路径

        Params:
            folder (str): 文件夹路径
            exts (tuple): 扩展名列表

        Returns:
            list[str]: 文件路径列表
        """
        filesPath = []
        for root, dirs, files in os.walk(folder):
            for file in files:
                if file.endswith(exts):
                    filePath = os.path.join(root, file)
                    filesPath.append(filePath)
        return filesPath

    @staticmethod
    def checkFolderExist(folder: str) -> bool:
        """
        检查文件夹是否存在

        Params:
            folder (str): 文件夹路径

        Returns:
            bool: 文件夹是否存在
        """
        if os.path.exists(folder) and os.path.isdir(folder):
            return True
        else:
            return False
        
    @staticmethod
    def getTime() -> str:
        """
        获取系统时间

        Returns:
            str: 系统时间
        """
        format = '%Y%m%d_%H%M%S'
        now = datetime.datetime.now()
        return now.strftime(format)
