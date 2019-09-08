import os
import sys
import shutil
import tkinter
from tkinter import messagebox as tkMessageBox
from tkinter import filedialog as tkFileDialog
import tkinter.simpledialog as sd
from pathlib import Path

#フォルダダイアログを開くクラス
class FolderDialog:
    def __init__(self):
        pass
    def showFD(self):
        root=tkinter.Tk()
        root.withdraw()

        iDir='c:/' 

        dirPath=tkFileDialog.askdirectory(initialdir=iDir)

        return dirPath

#ショートカット作成フォルダ選択
fd = FolderDialog()
fileFolderPath = fd.showFD()

#フォルダ未選択の場合
if fileFolderPath in (None,''):
    sys.exit()

#ファイル名を選択
filename = sd.askstring('入力', 'input', initialvalue='')
if(filename in (None,'')):
    sys.exit()

#すでに存在している場合はフォルダ作成を行わない
if not os.path.exists(filename):
    os.mkdir(filename)

p = Path(fileFolderPath)
#カウンタ
i=0

#ファイル数分処理を行う
for filePath in p.iterdir():
    if not filePath.is_file():
        continue
    i += 1
    shutil.copy2(str(filePath), os.path.join(filename,filename +'_' + str(i) + filePath.suffix))