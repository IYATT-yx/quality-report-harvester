from commontools import CommonTools as ct
from dialog import Dialog
import buildtime

import os
import sys

class Status:
    packaged = not sys.argv[0].endswith('.py')
    """软件是否处于打包状态"""

class Path:
    executableCommand = os.path.abspath(sys.argv[0]) if Status.packaged else sys.executable + ' ' + os.path.abspath(sys.argv[0])
    """软件的执行命令"""

    rootDir = os.path.dirname(
          sys.executable if Status.packaged else os.path.abspath(sys.argv[0])
    )
    """根目录"""

    programFileDir = os.path.dirname(os.path.abspath(sys.argv[0]))
    """程序文件所在目录"""

    executableFilePath = os.path.abspath(sys.argv[0]) if Status.packaged else sys.executable
    """可执行文件路径"""

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
    file = os.path.join(Path.programFileDir, 'dict.txt')
