from PIL import Image, ImageTk, ImageEnhance
import os
import stat
import pyautogui
import time
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
import pytesseract
import numpy as np

# =============================
# ファイル・ディレクトリ定義
# =============================
def get_base_dir():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
        if sys.platform == "darwin":
            base_dir = os.path.dirname(os.path.dirname(os.path.dirname(base_dir)))
        return base_dir
    else:
        return os.path.dirname(os.path.abspath(__file__))

def create_grant_dir(dirname):
    absolute_path = os.path.join(get_base_dir(), dirname)
    os.makedirs(absolute_path, exist_ok=True)
    os.chmod(absolute_path, stat.S_IRWXU)
    return absolute_path

def write_txt(file, write_value):
    if os.path.isfile(file) or not os.path.exists(file):
        with open(file, 'w', encoding='utf-8') as output_file:
            output_file.write(write_value)
        return file
    return ""

def write_pdf(file, write_value):
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

def write_xlsx(file, write_value):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws['A1'] = write_value
    wb.save(file)

base_dir = get_base_dir()
conf_dir = create_grant_dir("config")
out_dir = create_grant_dir("output")
log_dir = create_grant_dir("log")
temp_dir = ""

tesseract_path = os.path.join(get_base_dir(), "Tesseract-OCR", "tesseract.exe")
pytesseract.pytesseract.tesseract_cmd = tesseract_path

log_file = os.path.join(log_dir, "appError.log")
if os.path.isfile(log_file):
    os.remove(log_file)

logging.basicConfig(filename=log_file, 
                    level=logging.ERROR, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

class StdErrToLogger(object):
    def write(self, message):
        if message.strip() != "":
            logging.error(message)

    def flush(self):
        pass

sys.stderr = StdErrToLogger()

response = ""
output_sentences = ""

# =============================
# ダイアログ共通処理
# =============================
def create_dialog_root(title):
    root = tk.Tk()
    root.title(title)
    root.resizable(False, False)
    root.attributes('-topmost', True)
    return root

def geometry_center(root):
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    root.geometry(f"+{x}+{y}")

def show_info(title, message):
    root.attributes('-topmost', True)
    messagebox.showinfo(title, message)

def show_warning(title, message):
    root.attributes('-topmost', True)
    messagebox.showwarning(title, message)

# =============================
# 主処理
# =============================
def init():
    if sys.platform == "darwin":
        show_info(
            "通知",
            "事前に画面収録とコンピュータ制御の権限が必要になります。\n\n" +
            "システム設定の「プライバシーとセキュリティ」から「画面収録とシステムオーディオ録音」と「アクセシビリティ」に" +
            "アプリが追加されている状態で実行してください。\n"
        )

    global temp_dir
    files_and_dirs = os.listdir(base_dir)
    valid_files = sorted([item for item in files_and_dirs if item.startswith('temp_')])

    if len(valid_files) > 0:
        def do_continue():
            global temp_dir
            root.destroy()
            temp_dir = os.path.join(get_base_dir(), valid_files[len(valid_files)-1])
            open_second_dialog()

        def do_init():
            global temp_dir
            root.destroy()
            for item in valid_files:
                shutil.rmtree(os.path.join(base_dir, item))
            temp_dir = create_grant_dir(f"temp_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}")
            open_first_dialog()

        root = create_dialog_root("実行方式確認")
        frame = tk.Frame(root)
        frame.grid(column=1, row=3)

        ttk.Label(frame, text="前回文字起こしされていないキャプチャが残っています。\n前回の続きから実行するか選択してください。").grid(column=0, row=0, padx=15, pady=20)
        ttk.Button(frame, text="続きから", command=do_continue, width=15, padding=[0,5]).grid(column=0, row=1, pady=[0, 10])
        ttk.Button(frame, text="はじめから", command=do_init, width=15, padding=[0,5]).grid(column=0, row=2, pady=[0, 20])

        geometry_center(root)
        root.mainloop()
    else:
        temp_dir = create_grant_dir(f"temp_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}")
        open_first_dialog()

def open_first_dialog():
    def validate(page):
        page = page.strip()
        if not (page.isdigit() and 1 <= int(page) <= 1000):
            show_warning("入力エラー", "無効なページ数です。1〜1000の整数を入力してください。")
            return False
        return True

    def on_first_submit():
        val_entry_page = entry_page.get()
        val_page_tran = page_tran.get()
        if not validate(val_entry_page):
            return
        root.destroy()
        page = int(val_entry_page)
        get_screen_shot(page, val_page_tran)

    root = create_dialog_root("キャプチャ取得情報設定")
    frame_top = ttk.Frame(root, padding=(30, 20))
    frame_top.grid()

    label_page = ttk.Label(frame_top, text="ページ数", width=15)
    label_page.grid(column=0, row=1, columnspan=2, pady=[0, 15], sticky=tk.W)

    entry_page = ttk.Entry(frame_top, width=30)
    entry_page.grid(column=1, row=1, columnspan=2, pady=[0, 15], sticky=tk.E)

    radio_direction = ttk.Label(frame_top, text="ページ送りの方向", width=15)
    radio_direction.grid(column=0, row=2, sticky=tk.W)

    page_tran = tk.StringVar(value="left")
    radio_left = ttk.Radiobutton(frame_top, text="左に進む", variable=page_tran, value="left", padding=(5, 0))
    radio_right = ttk.Radiobutton(frame_top, text="右に進む", variable=page_tran, value="right", padding=(5, 0))
    radio_left.grid(column=1, row=2, sticky=tk.W)
    radio_right.grid(column=2, row=2, sticky=tk.W)

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
        
    def perform_ocr(root, valid_files):
        def update_progress(progress):
            progress_var.set(progress)  # Update progress bar value
            progress_label.config(text=f"{progress}%")  # Update progress percentage label

        str_result = ""
        i = 0
        global output_sentences

        for item in valid_files:
            item_path = os.path.join(temp_dir, item)
            image = cv2.imread(item_path, cv2.IMREAD_GRAYSCALE)

            # Increase contrast
            pil_img = Image.fromarray(image)
            enhancer = ImageEnhance.Contrast(pil_img)
            pil_img = enhancer.enhance(2)
            image = np.array(pil_img)  # Ensure 'image' is re-assigned

            # Apply thresholding
            _, thresh_image = cv2.threshold(image, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

            # Denoise and apply GaussianBlur
            blurred_image = cv2.GaussianBlur(thresh_image, (5, 5), 0)

            try:
                text = pytesseract.image_to_string(blurred_image, lang='jpn+jpn_vert', config='--oem 3 --psm 1')
                output_sentences += "\n" + text
            except Exception as e:
                print(f"Error occurred during OCR: {e}")

            os.remove(item_path)

            # Update progress bar
            i += 1
            progress = int((i / len(valid_files)) * 100)
            root.after(0, update_progress, progress)

        root.withdraw()
        shutil.rmtree(temp_dir)
    
    # メインスレッドでプログレスバーを動的に更新
    def main_thread_task(root, valid_files):
        def check_sub_thread():
            if sub_thread.is_alive():
                root.after(10, check_sub_thread)
                ret, frame = cap.read()
                if ret:
                    frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    img = Image.fromarray(frame)
                    imgtk = ImageTk.PhotoImage(image=img)
                    video_canvas.create_image(0, 0, anchor=tk.NW, image=imgtk)
                    video_canvas.imgtk = imgtk
                else:
                    cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            else:
                time.sleep(1)
                create_file(file_path)

        global sub_thread
        sub_thread = threading.Thread(target=perform_ocr, args=(root, valid_files,))
        sub_thread.start()
        root.after(1000, check_sub_thread)

    # Retrieve files for OCR
    files_and_dirs = os.listdir(temp_dir)
    valid_files = sorted([item for item in files_and_dirs if item != '.DS_Store'])

    root = tk.Tk()
    root.attributes('-topmost', True)
    root.withdraw()
    messagebox.showinfo(
        "Output Processing Info",
        f"【Number of Pages】\n{str(len(valid_files))}\n\n【Output File Path】\n{file_path}"
    )
    root.destroy()

    root = create_dialog_root("文字起こし実行状況")
    frame_top = ttk.Frame(root, padding=(0, 0))
    frame_top.grid()

    cap = cv2.VideoCapture(os.path.join(conf_dir, "now_processing.mp4"))

    video_canvas = tk.Canvas(frame_top, width=334, height=150)
    video_canvas.grid(column=0, row=0)

    progress_label = ttk.Label(frame_top, text="0%")
    progress_label.grid(column=0, row=1, pady=[15, 10])

    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(frame_top, maximum=100, variable=progress_var, length=300)
    progress_bar.grid(column=0, row=2, pady=[0, 25])

    geometry_center(root)

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