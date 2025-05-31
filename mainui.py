import tkinter
from tkinter import filedialog, messagebox
import os

from commontools import CommonTools
import constant
from operatortype import OperatorType
from wordproc import WordProc
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
        directory = filedialog.askdirectory(initialdir=constant.Path.programFileDir, title='选择目标文件所在文件夹')
        if directory:
            directory = os.path.normpath(directory)
            self.folderEntry.delete(0, tkinter.END)
            self.folderEntry.insert(0, directory)
        else:
            messagebox.showwarning('警告', '未选择文件夹')
            return

    def onOpenFileButton(self):
        files = filedialog.askopenfilenames(initialdir=constant.Path.programFileDir, title='选择目标文件', filetypes=self.wordFileTypes)
        if files:
            files = [os.path.normpath(file) for file in files]
            filesString = ''
            for file in files:
                filesString += (file + ';')
            self.fileEntry.delete(0, tkinter.END)
            self.fileEntry.insert(0, filesString)
        else:
            messagebox.showwarning('警告', '未选择文件')
            return

    def onOpenDictButton(self):
        file = filedialog.askopenfilename(initialdir=constant.Path.programFileDir, initialfile='diCommonTools.txt', title='选择自定义字典文件', filetypes=[('文本文件', '*.txt')])
        if file:
            file = os.path.normpath(file)
            self.dictEntry.delete(0, tkinter.END)
            self.dictEntry.insert(0, file)
        else:
            messagebox.showwarning('警告', '未选择词典文件')
            return

    def onExtractButton(self):
        operatorType = self.operatorType.get()
        excelProc = ExcelProc(operatorType, self.folderEntry.get())
        saveFile = excelProc.openExcel()
        if saveFile is None:
            messagebox.showwarning('警告', '未选择保存文件路径')
            return

        match self.operatorType.get():
            case OperatorType.FILE:
                files = self.fileEntry.get()
                files = files.split(';')
                files = [file for file in files if file != '']
            case OperatorType.FOLDER:
                folder = self.folderEntry.get()
                if not CommonTools.checkFolderExist(folder):
                    messagebox.showwarning('警告', '文件夹不存在')
                    return
                files = CommonTools.getFilesFromFolder(folder, self.exts)
            case OperatorType.ALL:
                folder = '.'
                folder = CommonTools.getAbsPath(folder)
                if not CommonTools.checkFolderExist(folder):
                    messagebox.showwarning('警告', '当前目录不存在')
                    return
                files = CommonTools.getFilesFromFolder(folder, self.exts)
            case _:
                messagebox.showwarning('警告', '未知操作类型')
                return
            
        WordSegmentation.loadCustomDict(self.dictEntry.get()) # 加载自定义词典

        for file in files:
            # 忽略临时文件
            if CommonTools.getNameFromPath(file).startswith('~$'):
                continue
            self.modifyExtractButtonText(f'正在处理：{CommonTools.getNameFromPath(file)}')
            status = self.wordProc.openWord(file)
            if not status:
                continue
            status, content = self.wordProc.do()
            if status:
                excelProc.writeContent(content, file)
            else:
                continue
        self.modifyExtractButtonText('提取完成')
        self.resetExtractButtonText()
        excelProc.closeExcel()
        
    def onOperatorTypeRadio(self):
        match self.operatorType.get():
            case OperatorType.FOLDER:
                self.fileEntry.config(state='readonly')
                self.openFileButton.config(state='disabled')
                self.folderEntry.config(state='normal')
                self.openFolderButton.config(state='normal')
            case OperatorType.FILE:
                self.folderEntry.config(state='readonly')
                self.openFolderButton.config(state='disabled')
                self.fileEntry.config(state='normal')
                self.openFileButton.config(state='normal')
            case OperatorType.ALL:
                self.folderEntry.config(state='readonly')
                self.fileEntry.config(state='readonly')
                self.openFolderButton.config(state='disabled')
                self.openFileButton.config(state='disabled')
            case _:
                messagebox.showwarning('警告', '未知操作类型')
                return
            
    def onHelpMenuAbout(self):
        messagebox.showinfo('关于', f'{constant.Basic.projectName}\n版本：{constant.Basic.version}\n作者：{constant.Basic.author}\n邮箱：{constant.Basic.email}')

    def onClosing(self):
        self.master.destroy()

    def modifyExtractButtonText(self, text: str):
        """
        修改提取按钮文本

        Args:
            text (str): 新的文本
            wrap (bool): 是否自动换行，默认True
        """
        self.extractButton.config(text=text)
        self.extractButton.update_idletasks()

    def resetExtractButtonText(self, time=3000):
        """
        重置提取按钮文本

        Args:
            time (int): 重置时间，单位为毫秒，默认3000毫秒
        """
        self.after(time, self.modifyExtractButtonText, '开始提取')

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
        self.dictEntry.insert(0, constant.CustomDict.file)

        # 操作按钮
        self.openFolderButton = tkinter.Button(self, text='打开文件夹', command=self.onOpenFolderButton, width=10)
        self.openFileButton = tkinter.Button(self, text='打开文件', command=self.onOpenFileButton, width=10)
        self.extractButton = tkinter.Button(self, text='开始提取', bd=5, width=50, command=self.onExtractButton)
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