from commontools import CommonTools as ct
from dialog import Dialog
import buildtime

import os
import sys

class Path:
    rootDir = os.path.dirname(os.path.abspath(__file__))
    """文件根目录"""

    programFileDir = os.path.dirname(os.path.abspath(sys.argv[0]))
    """程序文件所在目录"""

class Basic:
    projectName = '质量通报提取工具'
    version = buildtime.buildTime
    author = 'IYATT-yx'
    email = 'iyatt@iyatt.com'
    logoName = os.path.join(Path.rootDir, 'icon.ico')

class Dialog:
    fileName = os.path.join(Path.programFileDir, 'quality-report-harvester.log')
    format = '[ %(asctime)s %(levelname)-8s ] %(message)s'
    dateFormat = '%Y-%m-%d %H:%M:%S'
    level = Dialog.DEBUG
    encoding = 'utf-8'

class CustomDict:
    """自定义词典路径
    默认在软件同目录下或上级目录下
    """
    file1 = os.path.join(Path.programFileDir, 'dict.txt')
    file2 = os.path.join(os.path.dirname(Path.programFileDir), 'dict.txt')
    file = file1 if os.path.exists(file1) else file2
