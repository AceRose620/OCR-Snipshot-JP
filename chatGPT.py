from PIL import Image, ImageTk
import os
import stat
import pyautogui
import time
import openai
import base64
import tkinter as tk
from tkinter import messagebox, ttk
import re
import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import platform
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph
import sys
import datetime
import shutil
import cv2
import threading
import logging

# =============================
# ファイル・ディレクトリ定義
# =============================
# ベースディレクトリ取得
def get_base_dir() :
    if getattr(sys, 'frozen', False):
        # PyInstallerで固められたアプリケーションとして実行されている場合
        base_dir = os.path.dirname(sys.executable)

        # macOSの場合は[実行ファイル名/Contents/MacOS]がついてしまうため2つ上に上がる
        if sys.platform == "darwin":
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(base_dir)))

        return base_dir
    else:
        # 通常のPythonスクリプトとして実行されている場合
        return os.path.dirname(os.path.abspath(__file__))

# ディレクトリ作成
def create_grant_dir(dirname) :
    absolute_path = os.path.join(get_base_dir(), dirname)
    os.makedirs(absolute_path, exist_ok=True)
    os.chmod(absolute_path, stat.S_IRWXU)
    return absolute_path

# テキストファイル作成
def write_txt(file, write_value) :
    if os.path.isfile(file) or not os.path.exists(file) :
        with open(file, 'w', encoding='utf-8') as output_file:
            output_file.write(write_value)
        return file
    return ""

# PDFファイル作成
def write_pdf(file, write_value) :
    doc = SimpleDocTemplate(file, pagesize=letter)
    styles = getSampleStyleSheet()
    font_path = os.path.join(conf_dir, "NotoSansJP-VariableFont_wght.ttf")
    pdfmetrics.registerFont(TTFont('NotoSansJP', font_path))
    styles['Normal'].fontName = 'NotoSansJP'
    processed_text = write_value.replace("\t", "    ")
    flowables = []
    paragraph = Paragraph(processed_text, styles['Normal'])
    flowables.append(paragraph)
    doc.build(flowables)

# xlsxファイル作成
def write_xlsx(file, write_value) :
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = write_value
    wb.save(file)

# ベースディレクトリ
base_dir = get_base_dir()

# 設定ファイルディレクトリ
conf_dir = create_grant_dir("config")

# 出力ファイル格納ディレクトリ
out_dir = create_grant_dir("output")

# ログファイル格納ディレクトリ
log_dir = create_grant_dir("log")

temp_dir = ""

# 設定ファイル
conf_file = os.path.join(conf_dir, "api.conf")

# ログファイルを設定
log_file = os.path.join(log_dir, "appError.log")
if os.path.isfile(log_file):
    os.remove(log_file)

logging.basicConfig(filename=log_file, 
                    level=logging.ERROR, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# エラー出力をログにリダイレクト
class StdErrToLogger(object):
    def write(self, message):
        if message.strip() != "":
            logging.error(message)

    def flush(self):
        pass

# 標準出力と標準エラー出力のリダイレクト
sys.stderr = StdErrToLogger()

# グローバル変数
response = ""
output_sentences = ""

# =============================
# ダイアログ共通処理
# =============================

# tkオブジェクト作成
def create_dialog_root(title) :
    root = tk.Tk()
    root.title(title)
    root.resizable(False, False)
    
    root.attributes('-topmost', True)

    return root

def geometry_center(root) :

    root.update_idletasks()

    # ダイアログの幅と高さを取得
    width = root.winfo_width()
    height = root.winfo_height()

    # スクリーンの幅と高さを取得
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # ウィンドウの位置を計算
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    
    # ウィンドウの位置を設定
    root.geometry(f"+{x}+{y}")

def show_info(title, message) :
    root.attributes('-topmost', True)
    messagebox.showinfo(title, message)

def show_warning(title, message) :
    root.attributes('-topmost', True)
    messagebox.showwarning(title, message)

# =============================
# 主処理
# =============================

def init() :
    
    # Macの場合はスクショ処理前の確認ダイアログを表示
    if sys.platform == "darwin":
        show_info(
            "通知",
            "事前に画面収録とコンピュータ制御の権限が必要になります。\n\n" +
            "システム設定の「プライバシーとセキュリティ」から「画面収録とシステムオーディオ録音」と「アクセシビリティ」に" +
            "アプリが追加されている状態で実行してください。\n"
        )

    global temp_dir
    # tempディレクトリがある場合は続きから実行するか確認    
    files_and_dirs = os.listdir(base_dir)
    valid_files = sorted([item for item in files_and_dirs if item.startswith('temp_')])

    if len(valid_files) > 0:

        def do_continue() :
            global temp_dir
            root.destroy()
            temp_dir = os.path.join(get_base_dir(), valid_files[len(valid_files)-1])
            with open(conf_file, 'r', encoding='utf-8') as output_file:
                openai.api_key = output_file.read()
            open_second_dialog()

        def do_init() :
            global temp_dir
            root.destroy()

            # 既存のtempフォルダを削除
            for item in valid_files:
                shutil.rmtree(os.path.join(base_dir, item))

            # キャプチャファイル格納ディレクトリ作成
            temp_dir = create_grant_dir(f"temp_{datetime.datetime.now().strftime("%Y%m%d%H%M%S")}")
            open_first_dialog()

        # ダイアログ生成
        root = create_dialog_root("実行方式確認")
        frame = tk.Frame(root)
        frame.grid(column=1, row=3)

        ttk.Label(frame, text="前回文字起こしされていないキャプチャが残っています。\n前回の続きから実行するか選択してください。").grid(column=0, row=0, padx=15, pady=20)
        ttk.Button(frame, text="続きから", command=do_continue, width=15, padding=[0,5]).grid(column=0, row=1, pady=[0, 10])
        ttk.Button(frame, text="はじめから", command=do_init, width=15, padding=[0,5]).grid(column=0, row=2, pady=[0, 20])

        geometry_center(root)

        root.mainloop()

    else:
        # キャプチャファイル格納ディレクトリ作成
        temp_dir = create_grant_dir(f"temp_{datetime.datetime.now().strftime("%Y%m%d%H%M%S")}")
        
        # 最初のダイアログを呼び出して処理開始
        open_first_dialog()

# 1つ目のダイアログ
def open_first_dialog():
    # バリデーション
    def validate(api_key, page):
        if not re.match(r'^[A-Za-z0-9_-]{1,200}$', api_key):
            show_warning("入力エラー", "APIキーが無効です。\n1〜200文字の有効な英数字・記号を入力してください。")
            return False

        page = page.strip()
        if not (page.isdigit() and 1 <= int(page) <= 1000):
            show_warning("入力エラー", "無効なページ数です。1〜1000の整数を入力してください。")
            return False

        return True

    # ボタン押下時処理
    def on_first_submit():
        val_entry_key = entry_key.get()
        val_entry_page = entry_page.get()
        val_page_tran = page_tran.get()

        # バリデーション
        if not validate(val_entry_key, val_entry_page):
            return

        # ダイアログを閉じて入力値を取得
        root.destroy()
        page = int(val_entry_page)
        write_txt(conf_file, val_entry_key)
        openai.api_key = val_entry_key

        # スクショ取得処理へ
        get_screen_shot(page, val_page_tran)

    # gridLayoutのダイアログ生成
    root = create_dialog_root("キャプチャ取得情報設定")
    frame_top = ttk.Frame(root, padding=(30, 20))
    frame_top.grid()
    
    # APIキー
    label_key = ttk.Label(frame_top, text="APIキー", width=15)
    label_key.grid(column=0, row=0, columnspan=2, pady=[0, 10], sticky=tk.W)

    entry_key = ttk.Entry(frame_top, width=30)
    entry_key.grid(column=1, row=0, columnspan=2, pady=[0, 10], sticky=tk.E)
    api_key = ""
    if os.path.exists(conf_file) and os.path.isfile(conf_file) :
        with open(conf_file, 'r', encoding='utf-8') as output_file:
            api_key = output_file.read()
        entry_key.insert(0, api_key)

    # ページ数
    label_page = ttk.Label(frame_top, text="ページ数", width=15)
    label_page.grid(column=0, row=1, columnspan=2, pady=[0, 15], sticky=tk.W)

    entry_page = ttk.Entry(frame_top, width=30)
    entry_page.grid(column=1, row=1, columnspan=2, pady=[0, 15], sticky=tk.E)

    # ページ送りの方向
    radio_direction = ttk.Label(frame_top, text="ページ送りの方向", width=15)
    radio_direction.grid(column=0, row=2, sticky=tk.W)

    page_tran = tk.StringVar(value="left")
    radio_left = ttk.Radiobutton(frame_top, text="左に進む", variable=page_tran, value="left", padding=(5, 0))
    radio_right = ttk.Radiobutton(frame_top, text="右に進む", variable=page_tran, value="right", padding=(5, 0))
    radio_left.grid(column=1, row=2, sticky=tk.W)
    radio_right.grid(column=2, row=2, sticky=tk.W)

    # 次へボタン
    ttk.Button(frame_top, text="次へ", command=on_first_submit, width=10, padding=[0,5]).grid(columnspan=3, column=0, row=3, sticky=tk.S, pady=[15, 0])
    
    geometry_center(root)

    root.mainloop()

# スクショ取得
def get_screen_shot(page, val_page_tran) :
        # スクショ処理前の確認ダイアログを表示
        root = tk.Tk()
        root.attributes('-topmost', True)
        root.withdraw()
        messagebox.showinfo("確認", "ダイアログを閉じたら10秒後に処理を開始します。\nKindleの画面を表示してください。\n")
        # 10秒の間に、スクショしたいkindleの画面に移動
        root.after(10000)
        root.destroy()
        
        # スクショ処理
        for p in range(page):
            # 出力ファイル名(頭文字_連番.png)
            out_filename = "picture_" + str(p + 1).zfill(4) + '.png'
            # 画面全体のスクリーンショットを保存
            s = pyautogui.screenshot().save(os.path.join(temp_dir, out_filename))
            # 次のページ
            pyautogui.keyDown(val_page_tran)
            pyautogui.keyUp(val_page_tran)

        # 2つ目のダイアログの処理へ
        open_second_dialog()

# 2つ目のダイアログ
def open_second_dialog():
    # バリデーション
    def validate_second(file_name, export_format):
        if not file_name:
            show_warning("入力エラー", "ファイル名を入力してください。")
            return False

        invalid_chars = set('/\\:*?"<>|') if platform.system() == 'Windows' else set('/')
        if any(char in invalid_chars for char in file_name):
            show_warning("入力エラー", "ファイル名に無効な文字が含まれています。")
            return False

        if not export_format:
            show_warning("入力エラー", "書き出し形式を選択してください。")
            return False

        valid_options = ["txt", "pdf", "xlsx"]
        if export_format not in valid_options:
            show_warning("入力エラー", "無効な書き出し形式が選択されています。")
            return False

        return True

    # ボタン押下時処理
    def on_second_submit():

        val_entry_file_name = entry_file_name.get()
        val_format_var = format_var.get()

        if not validate_second(val_entry_file_name, val_format_var):
            return

        # ダイアログを閉じてファイル名を生成
        root.destroy()
        file_path = os.path.join(out_dir, f"{val_entry_file_name.replace('\n', '')}.{val_format_var}")

        # ファイル作成処理
        convert_img2txt(file_path)

    # ダイアログ生成
    root = create_dialog_root("保存情報の入力")
    frame_top = ttk.Frame(root, padding=(30, 10))
    frame_top.grid()

    ttk.Label(frame_top, text="画面イメージの取得が完了しました。\n出力ファイル名を指定してください。").grid(row=0, column=0, pady=(10, 0))
    ttk.Label(frame_top, text="ファイル名").grid(row=1, column=0, pady=(20, 0))
    entry_file_name = ttk.Entry(frame_top, width=40)
    entry_file_name.grid(row=2, column=0, pady=(5, 10))

    ttk.Label(frame_top, text="書き出し形式").grid(row=3, column=0, pady=(10, 0))
    option = ["txt", "pdf", "xlsx"]
    format_var = tk.StringVar(value=option[0])
    combobox = ttk.Combobox(frame_top, values=option, textvariable=format_var, state="readonly")
    combobox.grid(row=4, column=0, pady=(5, 10))

    ttk.Button(frame_top, text="確認", command=on_second_submit, width=10, padding=[0,5]).grid(row=5, column=0, pady=20)
    
    geometry_center(root)

    root.mainloop()

# 文字起こし処理
def convert_img2txt(file_path) :
    # 画像ファイルのエンコード
    def encode_image(image_path):
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')

    # APIからのレスポンスを取得してレスポンスを後続処理に使用
    def execute_sub_thread(root, valid_files):
        # 文字起こし
        def call_gpt_api(root, valid_files) :
            def update_progress(progress):
                # プログレスバーとパーセンテージの更新
                progress_var.set(progress)  # プログレスバーを更新
                progress_label.config(text=f"{progress}%")

            str_result = ""
            i = 0
            global output_sentences

            try:
                # 文字起こし処理開始
                for item in valid_files:
                    item_path = os.path.join(temp_dir, item)

                    response = ""
                    prompt = f"横書きで文字起こしをしてください。出力は文字起こしの回答のみにしてください。"
                    response = openai.chat.completions.create(
                        model="gpt-4o",
                        messages=[
                            {
                                "role": "user",
                                "content": [
                                    {"type": "text", "text": prompt},
                                    {
                                        "type": "image_url",
                                        "image_url": {
                                            "url": f"data:image/png;base64,{encode_image(item_path)}",
                                        },
                                    },
                                ],
                            }
                        ],
                        max_tokens=4096,
                    )
                    
                    os.remove(item_path)
                    print(response)

                    # レスポンス文字列化
                    if response:
                        output_sentences += "\n" + str(response.choices[0].message.content)

                    # 進捗バー更新（GUIの更新をメインスレッドで行うために after() を使用）
                    i += 1
                    progress = int(((i) / len(valid_files)) * 100)  # パーセント計算
                    root.after(0, update_progress, progress)
                
                root.withdraw()
                shutil.rmtree(temp_dir)
    
            except openai.RateLimitError as e:
                root.withdraw()
                root.after(1000, show_warning,
                    "APIエラー（RateLimitError）",
                    "APIの利用上限を超過しました。\nOpenAI APIの利用上限を確認してください。\n実行できたページまでの文字起こしを行います。\n"
                )

            except openai.APITimeoutError as e:
                root.withdraw()
                root.after(1000, show_warning,
                    "APIエラー（APITimeoutError）",
                    "APIのタイムアウトが発生しました。\n再度実行するか接続状況を確認してください。\n実行できたページまでの文字起こしを行います。\n"
                )
        
            except openai.BadRequestError as e:
                root.withdraw()
                root.after(1000, show_warning,
                    "APIエラー（BadRequestError）",
                    "文字起こし処理にて想定外のエラーが発生しました。\n再度実行するか接続状況を確認してください。\n実行できたページまでの文字起こしを行います。\n"
                )
            
            except openai.AuthenticationError as e:
                root.withdraw()
                root.after(1000, show_warning,
                    "APIエラー（AuthenticationError）",
                    "APIの認証に失敗しました。\n再度はじめから実行し、APIのキーを確認してください。\n"
                )
                sys.exit()
    
        call_gpt_api(root, valid_files)
    
    # メインスレッドでプログレスバーを動的に更新
    def main_thread_task(root, valid_files):
        # スレッドが完了したかを確認
        def check_sub_thread():
            if sub_thread.is_alive():
                root.after(10, check_sub_thread)  # まだスレッドが実行中の場合、再度確認
                # 動画のフレームを更新する
                ret, frame = cap.read()
                if ret:
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    img = Image.fromarray(frame)
                    imgtk = ImageTk.PhotoImage(image=img)
                    video_canvas.create_image(0, 0, anchor=tk.NW, image=imgtk)
                    video_canvas.imgtk = imgtk  # 参照を保持することで画像が表示される
                else:
                    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)  # 動画を最初に戻す
            else:
                # 後続処理
                time.sleep(1)
                create_file(file_path)

        # 別スレッドで文字起こし開始
        global sub_thread
        sub_thread = threading.Thread(target=execute_sub_thread, args=(root, valid_files,))
        sub_thread.start()
        root.after(1000, check_sub_thread)  # 1000ミリ秒後にスレッドの状態を確認

    # 文字起こし対象のファイルを取得
    files_and_dirs = os.listdir(temp_dir)
    valid_files = sorted([item for item in files_and_dirs if item != '.DS_Store'])

    # 最終メッセージで通知
    root = tk.Tk()
    root.attributes('-topmost', True)
    root.withdraw()
    messagebox.showinfo(
        "出力処理内容",
        f"【APIキー】\n{openai.api_key}\n\n【ページ数】\n{str(len(valid_files))}\n\n【出力ファイルパス】\n{file_path}"
    )
    root.destroy()

    root = create_dialog_root("文字起こし実行状況")
    frame_top = ttk.Frame(root, padding=(0, 0))
    frame_top.grid()

    # OpenCVを使用して動画をキャプチャ
    cap = cv2.VideoCapture(os.path.join(conf_dir, "now_processing.mp4"))

    video_canvas = tk.Canvas(frame_top, width=334, height=150)  # 動画サイズに合わせて指定
    video_canvas.grid(column=0, row=0)

    # プログレスバーの下にパーセンテージ表示
    progress_label = ttk.Label(frame_top, text="0%")
    progress_label.grid(column=0, row=1, pady=[15, 10])

    # プログレスバーの変数を設定
    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(frame_top, maximum=100, variable=progress_var, length=300)
    progress_bar.grid(column=0, row=2, pady=[0, 25])

    geometry_center(root)

    # ウィンドウの表示後にプログレスバーと動画再生を自動開始
    root.after(10, lambda: main_thread_task(root, valid_files))

    root.mainloop()

def create_file(file_path) :

    if not output_sentences:
        logging.error("書き出しする文字がありません。")
        show_warning(
            "文字起こしエラー",
            "書き出しする文字がありません。処理を終了します。\n"
        )
        sys.exit()

    # txtで出力
    if  file_path.endswith(".txt"):
        write_txt(file_path, output_sentences)
    
    # pdfで出力
    elif file_path.endswith(".pdf"):
        write_pdf(file_path, output_sentences)

    # xlsxで出力
    elif file_path.endswith(".xlsx"):
        write_xlsx(file_path, output_sentences)
    
    show_info(
        "通知",
        "出力が完了しました。\noutputフォルダ配下を確認してください。\n"
    )

init()