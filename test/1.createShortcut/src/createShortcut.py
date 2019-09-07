import tkinter
from tkinter import messagebox as tkMessageBox
from tkinter import filedialog as tkFileDialog
import os
import sys
from pathlib import Path
import os.path
import win32com.client
import datetime

class FolderDialog:
    def __init__(self):
        pass
    def showFD(self):
        root=tkinter.Tk()
        root.withdraw()

        iDir='c:/' 

        dirPath=tkFileDialog.askdirectory(initialdir=iDir)

        return dirPath

class ShortCut:
    def __init__(self,parentFolderPath):
        self.folderPath = parentFolderPath
    def createShortCut(self,saveFilePath):
        shell   = win32com.client.Dispatch('WScript.shell')

        shCut                  = shell.CreateShortcut(os.path.join(self.folderPath,os.path.basename(saveFilePath)+".lnk"))
        shCut.TargetPath       = saveFilePath
        shCut.WindowStyle      = 1
        shCut.IconLocation     = saveFilePath
        shCut.WorkingDirectory = self.folderPath
        shCut.Save()

#ショートカット作成のための親フォルダ取得
fd = FolderDialog()
fileFolderPath = fd.showFD()

if fileFolderPath in (None,''):
    sys.exit()

#ショートカット保存フォルダ作成
dt = datetime.datetime.now()
dtStr = dt.strftime('%Y%m%d_%H%M%S_%f')[:-3]
saveFolderPath = os.path.join(os.getcwd(), 'shortcut' + '_' + dtStr)
os.mkdir(saveFolderPath)

#ショートカット作成
shortCut = ShortCut(saveFolderPath)
p = Path(fileFolderPath)
for filePath in list(p.iterdir()):
    shortCut.createShortCut(str(filePath))




# for path in list(p.glob('*')):
#     message += str(path) + '\n'

# if message != '':
#     tkMessageBox.showinfo('FolderInfo', message)
# else:
#     tkMessageBox.showinfo('FolderInfo', 'NoFile')
