import tkinter.messagebox

import constant
from dialog import Dialog
from mainui import MainUi

def main():
    # 创建主窗口
    root = tkinter.Tk()

    # 设置窗口标题
    root.title(constant.Basic.projectName)

    # 窗口大小、位置
    defaultWidth = 1000
    defaultHeight = 150
    defaultX = int((root.winfo_screenwidth() - defaultWidth) / 2)
    defaultY = int((root.winfo_screenheight() - defaultHeight) / 2)
    root.geometry(f'{defaultWidth}x{defaultHeight}+{defaultX}+{defaultY}')
    root.minsize(defaultWidth, defaultHeight)

    # 图标
    root.iconbitmap(constant.Basic.logoName)

    # 初始化日志
    Dialog(
        constant.Dialog.fileName,
        constant.Dialog.format,
        constant.Dialog.dateFormat,
        constant.Dialog.level,
        constant.Dialog.encoding
    )

    Dialog.log(f'程序启动，版本号：{constant.Basic.version}')

    # 创建主界面
    mu = MainUi(root)
    mu.mainWindow()

    # 进入主循环
    root.mainloop()

if __name__ == '__main__':
    main()