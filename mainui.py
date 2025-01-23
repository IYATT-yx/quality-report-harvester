import tkinter
from tkinter import filedialog
import os

from commontools import CommonTools
from constant import Constant
from operatortype import OperatorType
from wordproc import WordProc
from dialog import Dialog
from excelproc import ExcelProc
from wordsegmentation import WordSegmentation

class MainUi(tkinter.Frame):
    def __init__(self, master: tkinter.Tk = None):
        super().__init__(master)
        self.master: tkinter.Tk = master
        self.pack(fill='both')

        self.wordFileTypes = [('Word 文档', '*.docx')]
        self.exts = tuple(ext[1:] for _, ext in self.wordFileTypes if ext != '*.*')

    def onOpenFolderButton(self):
        directory = filedialog.askdirectory(initialdir=CommonTools.getCurrentPath(), title='选择目标文件所在文件夹')
        if directory:
            directory = os.path.normpath(directory)
            self.folderEntry.delete(0, tkinter.END)
            self.folderEntry.insert(0, directory)
        else:
            Dialog.log('取消选择文件夹')
            return

    def onOpenFileButton(self):
        files = filedialog.askopenfilenames(initialdir=CommonTools.getCurrentPath(), title='选择目标文件', filetypes=self.wordFileTypes)
        if files:
            files = [os.path.normpath(file) for file in files]
            filesString = ''
            for file in files:
                filesString += (file + ';')
            self.fileEntry.delete(0, tkinter.END)
            self.fileEntry.insert(0, filesString)
        else:
            Dialog.log('取消选择文件')
            return

    def onOpenDictButton(self):
        file = filedialog.askopenfilename(initialdir=CommonTools.getCurrentPath(), initialfile='diCommonTools.txt', title='选择自定义字典文件', filetypes=[('文本文件', '*.txt')])
        if file:
            file = os.path.normpath(file)
            self.dictEntry.delete(0, tkinter.END)
            self.dictEntry.insert(0, file)
        else:
            Dialog.log('取消选择词典文件')
            return

    def onExtractButton(self):
        operatorType = self.operatorType.get()
        excelProc = ExcelProc(operatorType, self.folderEntry.get())
        saveFile = excelProc.openExcel()
        if saveFile is None:
            Dialog.log('取消保存文件')
            return
        else:
            Dialog.log('已选择保存文件路径：' + saveFile)

        match self.operatorType.get():
            case OperatorType.FILE:
                files = self.fileEntry.get()
                files = files.split(';')
                files = [file for file in files if file != '']
            case OperatorType.FOLDER:
                folder = self.folderEntry.get()
                if not CommonTools.checkFolderExist(folder):
                    Dialog.log('文件夹不存在', Dialog.ERROR)
                    return
                files = CommonTools.getFilesFromFolder(folder, self.exts)
            case OperatorType.ALL:
                folder = '.'
                folder = CommonTools.getAbsPath(folder)
                if not CommonTools.checkFolderExist(folder):
                    Dialog.log('文件夹不存在', Dialog.ERROR)
                    return
                files = CommonTools.getFilesFromFolder(folder, self.exts)
            case _:
                Dialog.log('未知操作类型', Dialog.CRITICAL)
                return
            
        WordSegmentation.loadCustomDict(self.dictEntry.get()) # 加载自定义词典

        Dialog.log('开始处理通报文件：')
        for file in files:
            # 忽略临时文件
            if CommonTools.getNameFromPath(file).startswith('~$'):
                continue
            status = self.wordProc.openWord(file)
            if not status:
                continue
            status, content = self.wordProc.do()
            if status:
                Dialog.log(f'{CommonTools.getNameFromPath(file)}')
                excelProc.writeContent(content, file)
            else:
                continue
        Dialog.log('通报文件处理完成', Dialog.INFO)
        excelProc.closeExcel()
        
    def onOperatorTypeRadio(self):
        match self.operatorType.get():
            case OperatorType.FOLDER:
                Dialog.log('已选择文件夹模式')
                self.fileEntry.config(state='readonly')
                self.openFileButton.config(state='disabled')
                self.folderEntry.config(state='normal')
                self.openFolderButton.config(state='normal')
            case OperatorType.FILE:
                Dialog.log('已选择文件模式')
                self.folderEntry.config(state='readonly')
                self.openFolderButton.config(state='disabled')
                self.fileEntry.config(state='normal')
                self.openFileButton.config(state='normal')
            case OperatorType.ALL:
                Dialog.log('已选择全部模式')
                self.folderEntry.config(state='readonly')
                self.fileEntry.config(state='readonly')
                self.openFolderButton.config(state='disabled')
                self.openFileButton.config(state='disabled')
            case _:
                Dialog.log('未知操作模式', Dialog.CRITICAL)
                return
            
    def onHelpMenuAbout(self):
        tkinter.messagebox.showinfo('关于', f'{Constant.Basic.projectName}\n版本：{Constant.Basic.version}\n作者：{Constant.Basic.author}\n邮箱：{Constant.Basic.email}')

    def getResolution(self):
        """
        获取分辨率
        """
        Dialog.log(f'屏幕分辨率：{self.winfo_screenwidth()}×{self.winfo_screenheight()}，窗口大小：{self.master.winfo_width()}×{self.master.winfo_height()}')

    def onClosing(self):
        Dialog.log('程序退出')
        self.master.destroy()

    def mainWindow(self):
        # 文字标签
        tkinter.Label(self, text='文件夹').grid(row=0, column=0, sticky='w')
        tkinter.Label(self, text='文件').grid(row=1, column=0, sticky='w')
        tkinter.Label(self, text='词典').grid(row=2, column=0, sticky='w')
        tkinter.Label(self, text='范围').grid(row=3, column=0, sticky='w')
        
        # 路径输入框
        self.folderEntry = tkinter.Entry(self)
        self.fileEntry = tkinter.Entry(self)
        self.dictEntry = tkinter.Entry(self)
        self.dictEntry.insert(0, Constant.CustomDict.file)

        # 操作按钮
        self.openFolderButton = tkinter.Button(self, text='打开文件夹', command=self.onOpenFolderButton, width=10)
        self.openFileButton = tkinter.Button(self, text='打开文件', command=self.onOpenFileButton, width=10)
        self.extractButton = tkinter.Button(self, text='提取', width=4, command=self.onExtractButton)
        self.openDictButton = tkinter.Button(self, text='打开词典', command=self.onOpenDictButton, width=10)

        # 模式单选
        self.operatorType = tkinter.IntVar()
        self.operatorType.set(OperatorType.FOLDER) # 默认文件夹模式
        tkinter.Radiobutton(self, text='当前目录下所有文件', variable=self.operatorType, value=OperatorType.ALL, command=self.onOperatorTypeRadio) \
        .grid(row=3, column=1, sticky='w')
        tkinter.Radiobutton(self, text='文件夹', variable=self.operatorType, value=OperatorType.FOLDER, command=self.onOperatorTypeRadio) \
        .grid(row=3, column=2, sticky='w')
        tkinter.Radiobutton(self, text='文件', variable=self.operatorType, value=OperatorType.FILE, command=self.onOperatorTypeRadio) \
        .grid(row=3, column=3, sticky='w')

        # 菜单
        menuBar = tkinter.Menu(self)
        self.master.config(menu=menuBar)

        helpMenu = tkinter.Menu(menuBar, tearoff=0)
        menuBar.add_cascade(label='帮助', menu=helpMenu)

        helpMenu.add_command(label='关于', command=self.onHelpMenuAbout)

        # 路径输入框和操作按钮的布局
        self.grid_columnconfigure(1, weight=1)
        self.folderEntry.grid(row=0, column=1, columnspan=2, sticky='ew')        
        self.openFolderButton.grid(row=0, column=3, sticky='w')
        self.fileEntry.grid(row=1, column=1, columnspan=2, sticky='ew')
        self.openFileButton.grid(row=1, column=3, sticky='w')
        self.dictEntry.grid(row=2, column=1, columnspan=2, sticky='ew')
        self.openDictButton.grid(row=2, column=3, sticky='w')
        self.extractButton.grid(row=0, column=4, rowspan=3, sticky='nsew')
        self.grid_rowconfigure(0, weight=1) # 高度不能自动调整

        self.onOperatorTypeRadio() # 初始界面状态
        self.wordProc = WordProc()

        self.master.protocol("WM_DELETE_WINDOW", self.onClosing)

        self.after(1000, self.getResolution)