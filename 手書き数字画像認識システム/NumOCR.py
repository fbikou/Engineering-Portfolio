import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

import glob

import requests
import base64
import json

class NumOCR:
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
        self.root.rowconfigure(index=3, weight=1)
        self.root.columnconfigure(index=1, weight=1)
        self.makeFrame(0 ,'手書き数字画像ファイル(png形式)保存フォルダ',True)
        self.makeFrame(1 ,'Google Cloud Vision APIのAPIキー',False)
        frame = ttk.Frame(self.root, padding=10,relief="groove")
        FinishButton = ttk.Button(frame, text="実行", command=lambda:self.getInputButton())
        FinishButton.pack()
        frame.grid(row=2,column=1)
        self.root.mainloop() 

    def makeFrame(self,rowNum,LabelName,ReferOrNot):
        #Frameの作成(Frameを含むシステムの構成について[URL]:https://denno-sekai.com/tkinter-frame/)
        frame = ttk.Frame(self.root, padding=10,relief="groove")
        #「参照PDFファイル一覧」の見出し作成(位置:上)
        IDirLabel = ttk.Label(frame, text=LabelName, padding=(5, 2))
        IDirLabel.pack(side=tk.TOP)
        #参照データ(Excel)のファイルパスの表示欄を設定(位置:中央)
        StateType=''
        if LabelName=='手書き数字画像ファイル(png形式)':
            StateType='readonly'
        else:
            StateType='normal'
        IDDataFile = ttk.Entry(frame, textvariable=tk.StringVar(), state=StateType,width=120)
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
        if self.inputs[0]=="" or self.inputs[1]=="":
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
    
        # APIを呼び、認識結果をjson型で返す
    def request_cloud_vison_api(self,image_base64):
        GOOGLE_CLOUD_VISION_API_URL = 'https://vision.googleapis.com/v1/images:annotate?key='
        API_KEY = self.inputs[1]
        api_url = GOOGLE_CLOUD_VISION_API_URL + API_KEY
        req_body = json.dumps({
            'requests': [{
                'image': {
                    # jsonに変換するためにstring型に変換する
                    'content': image_base64.decode('utf-8')
                },
                'features': [{
                    # ここを変更することで分析内容を変更できる
                    'type': 'TEXT_DETECTION',
                    'maxResults': 10,
                }]
            }]
        })
        res = requests.post(api_url, data=req_body)
        return res.json()

    # 画像読み込み
    def img_to_base64(self,filepath):
        with open(filepath, 'rb') as img:
            img_byte = img.read()
        return base64.b64encode(img_byte)

    def OutputNum(self):
        if len(self.inputs)!=0:
            ImgPaths=glob.glob(self.inputs[0]+'/*.png')
            if len(ImgPaths)==0:
                messagebox.showerror('ファイルエラー',\
                '選択フォルダ内にPDFファイルはありません。\n'+\
                '選択フォルダをご確認の上, 再度, お試し下さい。')
            else:
                ResultList=[]
                for ImgPath in ImgPaths:
                    # 文字認識させたい画像を設定
                    img_base64 = self.img_to_base64(ImgPath)
                    result = self.request_cloud_vison_api(img_base64)
                    # 認識した文字を出力
                    text_r = result["responses"][0]["textAnnotations"][1]["description"]
                    if text_r!='':
                        ResultList.append(text_r)
                    else:
                        ResultList.append('認識不可')
                messagebox.showinfo('実行完了',
                    '　プログラムの実行が完了しました。\n'+\
                    '認識結果が以下のようになります。(左から順に)\n'+\
                    ','.join(ResultList))

NumOCR().OutputNum()