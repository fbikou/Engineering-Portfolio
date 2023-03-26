import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

import glob
import win32api
import win32print
import sys
import os
import time

class PDFAutoPrint:
    def __init__(self):
        self.IDDataFiles=[]
        self.inputs=[]
        # rootの作成
        self.root = tk.Tk()
        self.makeGUI()

    def makeGUI(self):
        #tkのGUI画面のタイトル設定
        self.root.title("フォルダ選択画面")
        #GUI画面のサイズ設定
        self.root.geometry('800x200')
        #Frame割り当て用にGUI画面に行列を設定する[今回は7×3]
        self.root.rowconfigure(index=1, weight=1)
        self.root.rowconfigure(index=2, weight=1)
        self.root.columnconfigure(index=1, weight=1)
        self.makeFrame(0 ,'参照PDFファイル一覧',True)
        frame = ttk.Frame(self.root, padding=10,relief="groove")
        FinishButton = ttk.Button(frame, text="実行", command=lambda:self.getInputButton())
        FinishButton.pack()
        frame.grid(row=1,column=1)
        self.root.mainloop() 

    def makeFrame(self,rowNum,LabelName,ReferOrNot):
        #Frameの作成(Frameを含むシステムの構成について[URL]:https://denno-sekai.com/tkinter-frame/)
        frame = ttk.Frame(self.root, padding=10,relief="groove")
        #「参照PDFファイル一覧」の見出し作成(位置:上)
        IDirLabel = ttk.Label(frame, text=LabelName, padding=(5, 2))
        IDirLabel.pack(side=tk.TOP)
        #参照データ(Excel)のファイルパスの表示欄を設定(位置:中央)
        IDDataFile = ttk.Entry(frame, textvariable=tk.StringVar(), state='readonly',width=120)
        IDDataFile.pack()
        self.IDDataFiles.append(IDDataFile)
        #この行はラムダ式で模擬関数を定義し, 
        #GUI表示時点による実行を避け, 参照ボタンを押した時のみ実行する仕様に設定(位置:右(下))
        if ReferOrNot:
            IDDataButton = ttk.Button(frame, text="参照", command=lambda:self.referenceButton(False,rowNum))
            IDDataButton.pack(side=tk.RIGHT)
        #gridはrow(column)figureで行列を設定しなければ使えない
        frame.grid(row=rowNum,column=1)

    def referenceButton(self,writeable,positionNum):
        #フォルダーパスの取得
        FolderPath = filedialog.askdirectory(initialdir=os.getcwd())
        if FolderPath:
            if not writeable:
                self.IDDataFiles[positionNum].configure(state='normal')
            self.IDDataFiles[positionNum].delete(0,tk.END)
            self.IDDataFiles[positionNum].insert(tk.END,FolderPath)
            if not writeable:
                self.IDDataFiles[positionNum].configure(state='readonly')
    
    def getInputButton(self):
        DestroyJudge=False
        for EntryInput in self.IDDataFiles:
            self.inputs.append(EntryInput.get())
        self.root.withdraw()
        if self.inputs[0]=="":
            messagebox.showerror('入力不足エラー', '全ての欄を入力して下さい。')
            self.reinitialization()            
        else:
            self.root.destroy()
            DestroyJudge=True
        if DestroyJudge==False:
            self.makeGUI()

    def reinitialization(self):
        for EntryInput in self.IDDataFiles:
            EntryInput.delete(0,tk.END)
        self.inputs.clear()
        self.IDDataFiles.clear()
        # rootの再表示
        self.root.deiconify()
    
    def AutoPrint(self):
        if len(self.inputs)!=0:
            PDFPaths=glob.glob(self.inputs[0]+'/*.pdf')
            if len(PDFPaths)==0:
                messagebox.showerror('ファイルエラー',\
                '選択フォルダ内にPDFファイルはありません。\n'+\
                '選択フォルダをご確認の上, 再度, お試し下さい。')
            else:
                for PDFPath in PDFPaths:
                    win32api.ShellExecute(0,"print",PDFPath,"/c:""%s" % win32print.GetDefaultPrinter(),".",0)
                    time.sleep(3)
                messagebox.showinfo('実行完了',
                    '　プログラムの実行が完了しました。\n'+\
                    '　印刷状況をご確認下さい。\n'+\
                    '　正常に印刷されない場合はご使用中のPCのプリンター・印刷設定をご確認の上, 再度, お試し下さい。')

PDFAutoPrint().AutoPrint()