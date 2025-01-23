import os

from commontools import CommonTools as ct
from dialog import Dialog

class Constant:
    class Basic:
        projectName = '质量通报提取工具'
        version = '202501230001'
        author = 'IYATT-yx'
        email = 'iyatt@iyatt.com'
        defaultDict = 'dict.txt'
        logoName = 'icon.ico'

    class Dialog:
        fileName = os.path.join(ct.getCurrentPath(), 'quality-report-harvester.log')
        format = '[ %(asctime)s %(levelname)-8s ] %(message)s'
        dateFormat = '%Y-%m-%d %H:%M:%S'
        level = Dialog.DEBUG
        encoding = 'utf-8'

    class CustomDict:
        file = os.path.join(ct.getCurrentPath(), 'dict.txt')
