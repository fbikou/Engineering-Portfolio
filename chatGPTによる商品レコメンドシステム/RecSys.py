import openai
import tkinter as tk
import tkinter.scrolledtext as scrolledtext
import tkinter.messagebox as messagebox
import json
import re
import webbrowser
import requests
from bs4 import BeautifulSoup
class RecSys:
    def __init__(self):
        openai.api_key = "OpenAIのAPIキー"
        openai.api_requestor.API_REQUEST_TIMEOUT = 5
        self.window=None
        self.output_text=None
        self.input_state=None
        self.age_spinbox=None
        self.male_button=None
        self.female_button=None
        self.job_text=None
        self.category_text=None
        self.preference_text=None
        self.company_text=None
        self.input_text=None
        self.start_pos="1.0"
    def is_profile_complete(self):
        return self.age_var.get() and self.gender_value.get() and self.job_text.get('1.0', tk.END).strip()
    def is_category_complete(self):
        return self.category_text.get('1.0', tk.END).strip()
    def get_response(self,prompt):
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "user", "content": prompt}
            ],
            #API_REQUEST_TIMEOUTとtemperatureが両方, 大きいと応答のバリエーションが増える
            temperature=0.5,  # 応答を速くする為にtemperatureを0.5に変更する
            #max_tokensが大きいと, よりユーザーの要望にフィットした提案と説明を応答として返す
            max_tokens=512,  # 応答を速くする為にmax_tokensを512に変更する
            n=1,
            stop=None,
            frequency_penalty=0,
            presence_penalty=0
        )
        message = json.loads(json.dumps(response.choices[0]))["message"]["content"]
        return message
    def copy_to_clipboard(self):
        self.window.clipboard_clear()
        self.window.clipboard_append(self.output_text.get('1.0', tk.END))
    def toggle_input_state(self):
        if self.input_state.get():
            self.age_spinbox.configure(state="disabled")
            self.male_button.configure(state="disabled")
            self.female_button.configure(state="disabled")
            self.job_text.configure(state="disabled")
            self.category_text.configure(state="disabled")
            self.preference_text.configure(state="disabled")
            self.company_text.configure(state="disabled")
            self.input_text.configure(state="normal")
        else:
            self.age_spinbox.configure(state="normal")
            self.male_button.configure(state="normal")
            self.female_button.configure(state="normal")
            self.job_text.configure(state="normal")
            self.category_text.configure(state="normal")
            self.preference_text.configure(state="normal")
            self.company_text.configure(state="normal")
            self.input_text.configure(state="disabled")
    def get_next_position(self, start, count):
        line, col = map(int, start.split('.'))
        col += count
        while True:
            line_end = f'{line + 1}.0'
            next_char = self.output_text.get(f'{line}.{col}')
            if next_char == '\n' or not next_char:
                break
            col += 1
            if f'{line}.{col}' >= line_end:
                line += 1
                col = 0
                if not self.output_text.get(f'{line}.0', line_end):
                    break
        return f'{line}.{col}'
    def words_in_txt(self,words,txt):
        if len(words)>0:
            for word in words:
                if word in txt:
                    return True
        return False
    def check_url(self,url):
        try:
            response = requests.get(url)
            response.raise_for_status() # レスポンスがエラーなら例外を発生
            flgs=[]
            #レスポンス状態が404だったらエラー判定
            flgs.append(response.status_code == 404)
            #タイトル内に「ページが見つかりません」やエラー要素があったらエラー判定
            soup = BeautifulSoup(response.content, 'html.parser')
            title_tag = soup.find('title')
            flgs.append(self.words_in_txt(["ページが見つかりません","Not Found","error"],title_tag))
            #レスポンステキスト内に「ページが見つかりません」があったらエラー判定
            flgs.append(self.words_in_txt(["ページが見つかりません","Not Found"],response.text))
            #url内にエラー要素があったらエラー判定
            flgs.append(self.words_in_txt(["error","404"],response.url))
            return  any(flgs)
        except (requests.exceptions.HTTPError, requests.exceptions.ConnectionError, requests.exceptions.Timeout,
                requests.exceptions.TooManyRedirects, requests.exceptions.RequestException, AttributeError):
            # 上記の例外が発生した場合、空の文字列を返す
            return True
    def linkfy_text(self):
        start = self.start_pos
        txt=self.output_text.get(start, tk.END).strip()
        pattern = r'(https?://[^\s]+)'
        urls = re.findall(pattern, txt)
        url_error_msg="[URLエラーが発生した為, 削除しました。]"
        for url in urls:
            link_text = url
            tag = 'link-' + url
            self.output_text.tag_config(tag, foreground='blue', underline=True)
            def callback(event, url=url):
                webbrowser.open_new(url)
            find_first = self.output_text.search(url, start, 'end')
            find_end=self.get_next_position(find_first,len(link_text))
            if not self.check_url(url):
                self.output_text.replace(find_first, find_end, link_text, tag)
                self.output_text.tag_bind(tag, '<Button-1>', callback)
            else:
                self.output_text.replace(find_first, find_end, url_error_msg)
                find_end=self.get_next_position(find_first,len(url_error_msg))
            start = self.get_next_position(find_end,1)
    def get_chat_response(self):
        if not self.is_profile_complete():
            messagebox.showerror("エラー", "プロフィールを入力してください。")
            return
        if not self.is_category_complete():
            messagebox.showerror("エラー", "商品カテゴリを入力してください。")
            return
        if self.input_state.get():
            input_prompt = self.input_text.get('1.0', tk.END).strip()
        else:
            self.output_text.delete('1.0', 'end')
            input_prompt = "私は"
            if self.age_var.get():  # 年齢の値を取得する
                input_prompt += f"{self.age_var.get()}歳"
            if self.gender_value.get():
                input_prompt += f"{self.gender_value.get()}の"
            if self.job_text.get('1.0', tk.END).strip():
                input_prompt += f"{self.job_text.get('1.0', tk.END).strip()}です！"
            category_prompt = self.category_text.get('1.0', tk.END).strip()
            if category_prompt:
                input_prompt += f"\n私にオススメの{category_prompt}を提案して下さい！"
            else:
                input_prompt += "\n私にオススメの商品を提案して下さい！"
            preference = self.preference_text.get('1.0', tk.END).strip()
            if preference:
                input_prompt += f"\nこだわりポイントは「{preference}」です！"
            company_prompt = self.company_text.get('1.0', tk.END).strip()
            if company_prompt:
                input_prompt += f"\n尚、{company_prompt}で販売している商品を提案して下さい！"
            input_prompt += "\n提案する商品に関するサイトがあれば、URLも出力して下さい！"+\
                "(URLのレスポンスコードが404の場合やURL先のサイトに「ページが見つかりません」という旨の文章がある場合は出力しないで下さい！)"
            input_prompt += "\nもし、商品提案に必要な情報が不足している場合は、何でも私に聞いて下さい！"
            #input_prompt += "\n又, 端的に提案して下さい！"
        response = self.get_response(input_prompt)
        if response:
            output_prompt = f"user: {input_prompt}\nAI: {response}\n"
        else:
            output_prompt = "API error: No response received"
        #前の文章が終わる位置をstart_posに格納し, 更新
        self.start_pos=self.output_text.index(tk.END)
        self.output_text.insert(tk.END, output_prompt)
        self.linkfy_text()
    def create_gui(self):
        self.window = tk.Tk()
        self.window.title("ChatGPT APIによる商品レコメンドシステム")
        self.window.geometry("800x700")  # ウィンドウサイズを800x700に変更
        profile_frame = tk.LabelFrame(self.window, text="プロフィール(必須)")
        profile_frame.pack()
        age_label = tk.Label(profile_frame, text="年齢")
        age_label.pack(side=tk.LEFT, padx=5)
        self.age_var = tk.IntVar(value=0)
        self.age_spinbox = tk.Spinbox(profile_frame, from_=0, to=120, width=10, textvariable=self.age_var, state="normal")
        self.age_spinbox.pack(side=tk.LEFT, padx=5)
        gender_label = tk.Label(profile_frame, text="性別")
        gender_label.pack(side=tk.LEFT, padx=5)
        self.gender_value = tk.StringVar(value=None)  # デフォルトはNoneにして初期選択なしにする
        self.male_button = tk.Radiobutton(profile_frame, text="男性", variable=self.gender_value, value="男性", state="normal")
        self.female_button = tk.Radiobutton(profile_frame, text="女性", variable=self.gender_value, value="女性", state="normal")
        self.male_button.pack(side=tk.LEFT, padx=5)
        self.female_button.pack(side=tk.LEFT, padx=5)
        job_label = tk.Label(profile_frame, text="職業")
        job_label.pack(side=tk.LEFT, padx=5)
        self.job_text = tk.Text(profile_frame, width=30, height=1, state="normal")
        self.job_text.pack(side=tk.LEFT, padx=5)
        category_label = tk.Label(self.window, text="商品カテゴリ(必須)")
        category_label.pack()
        self.category_text = tk.Text(self.window, width=30, height=1, state="normal")
        self.category_text.pack()
        preference_label = tk.Label(self.window, text="こだわりポイント(任意)")
        preference_label.pack()
        self.preference_text = tk.Text(self.window, width=30, height=3, state="normal")
        self.preference_text.pack()
        company_label = tk.Label(self.window, text="販売会社(任意)")
        company_label.pack()
        self.company_text = tk.Text(self.window, width=30, height=1, state="normal") 
        self.company_text.pack()
        self.input_state = tk.BooleanVar(value=False)
        input_frame = tk.Frame(self.window)
        input_frame.pack()
        input_text_frame = tk.Frame(input_frame)
        input_text_frame.pack(side=tk.LEFT, padx=5)
        input_text_label = tk.Label(input_text_frame, text="チャット入力欄")
        input_text_label.pack()
        self.input_text = tk.Text(input_text_frame, width=50, height=10, state="disabled")
        self.input_text.pack()
        toggle_input_button_frame = tk.Frame(input_frame)
        toggle_input_button_frame.pack(side=tk.LEFT)
        toggle_input_button = tk.Checkbutton(toggle_input_button_frame, text="チャット機能を使う", variable=self.input_state, command=self.toggle_input_state)
        toggle_input_button.pack(side=tk.LEFT, padx=5)
        output_label = tk.Label(self.window, text="出力欄")
        output_label.pack()
        self.output_text = scrolledtext.ScrolledText(self.window, width=80, height=20, wrap=tk.WORD)
        self.output_text.pack()
        copy_button = tk.Button(self.window, text="コピー", command=self.copy_to_clipboard)
        copy_button.pack()
        chat_button = tk.Button(self.window, text="送信", command=self.get_chat_response)
        chat_button.pack()
        self.window.mainloop()
if __name__=="__main__":
    RecSys().create_gui()