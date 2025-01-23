import tkinter.messagebox
import sys
import os

from constant import Constant
from dialog import Dialog
from commontools import CommonTools
from mainui import MainUi

def main():
    # 创建主窗口
    root = tkinter.Tk()

    # 设置窗口标题
    root.title(Constant.Basic.projectName)

    # 窗口大小、位置
    defaultWidth = 800
    defaultHeight = 150
    defaultX = int((root.winfo_screenwidth() - defaultWidth) / 2)
    defaultY = int((root.winfo_screenheight() - defaultHeight) / 2)
    root.geometry(f'{defaultWidth}x{defaultHeight}+{defaultX}+{defaultY}')
    root.minsize(defaultWidth, defaultHeight)

    # 图标
    if CommonTools.getPackagedStatus():
        root.iconbitmap(os.path.join(sys._MEIPASS, Constant.Basic.logoName))
    else:
        root.iconbitmap(Constant.Basic.logoName)

    # 初始化日志
    Dialog(
        Constant.Dialog.fileName,
        Constant.Dialog.format,
        Constant.Dialog.dateFormat,
        Constant.Dialog.level,
        Constant.Dialog.encoding
    )

    Dialog.log(f'程序启动，版本号：{Constant.Basic.version}')

    # 创建主界面
    mu = MainUi(root)
    mu.mainWindow()

    # 进入主循环
    root.mainloop()

if __name__ == '__main__':
    
    main()