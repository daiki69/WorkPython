import csv
import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
from tkcalendar import Calendar, DateEntry
from ttkthemes import ThemedStyle
from PIL import Image, ImageTk
import os

def generate_csv():
    start_time_str = start_time_entry.get()
    end_time_str = end_time_entry.get()
    interval_count = interval_count_entry.get()
    description = description_text.get("1.0", "end-1c")  # テキストフィールドから説明文を取得

    try:
        start_time = datetime.datetime.strptime(start_time_str, "%Y-%m-%d")
        end_time = datetime.datetime.strptime(end_time_str, "%Y-%m-%d")
        
        day_interval = (end_time - start_time) / int(interval_count)

        time_list = [start_time + i * day_interval for i in range(int(interval_count) + 1)]

        # 保存先のディレクトリを選択するダイアログを表示する
        save_directory = filedialog.askdirectory()

        if save_directory:
            file_path = os.path.join(save_directory, 'time_list.csv')

            with open(file_path, mode='w', newline='', encoding='utf-8') as csv_file:
                fieldnames = ['Time', 'Description']  # Descriptionフィールドを追加
                writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                writer.writeheader()
                for time in time_list:
                    time = time.replace(hour=0, minute=0, second=0)
                    writer.writerow({'Time': time.strftime('%Y-%m-%dT%H:%M:%S'), 'Description': description})

            result = messagebox.askquestion("成功", "CSVファイルが正常に生成されました！\nExcelで開きますか？")
            if result == "yes":
                os.startfile(file_path)
    except ValueError:
        messagebox.showerror("エラー", "無効な入力です。正しい値を入力してください。")

# カレンダーの表示を日本語に設定する
def set_japanese_locale():
    return 'ja_JP'

# GUIウィンドウを作成する
window = tk.Tk()
window.title("CSVジェネレータ")
style = ThemedStyle(window)
style.set_theme("arc")

# ロゴ画像を読み込む
logo_image = Image.open("FuzorLogo.png")
logo_photo = ImageTk.PhotoImage(logo_image)

# ロゴ画像を表示するラベルを作成する
logo_label = tk.Label(window, image=logo_photo)
logo_label.pack()

# 開始日時の入力フィールドを作成する
start_time_label = tk.Label(window, text="開始日時:", font=("メイリオ", 12, "bold"))
start_time_label.pack()
start_time_entry = DateEntry(window, date_pattern='yyyy-mm-dd', locale=set_japanese_locale(), font=("メイリオ", 12))
start_time_entry.pack()

# 終了日時の入力フィールドを作成する
end_time_label = tk.Label(window, text="終了日時:", font=("メイリオ", 12, "bold"))
end_time_label.pack()
end_time_entry = DateEntry(window, date_pattern='yyyy-mm-dd', locale=set_japanese_locale(), font=("メイリオ", 12))
end_time_entry.pack()

# 分割数の入力フィールドを作成する
interval_count_label = tk.Label(window, text="分割数:", font=("メイリオ", 12, "bold"))
interval_count_label.pack()
interval_count_entry = tk.Entry(window, font=("メイリオ", 12))
interval_count_entry.pack()

# 説明文の入力フィールドを作成する
description_label = tk.Label(window, text="説明文:", font=("メイリオ", 12, "bold"))
description_label.pack()
description_text = tk.Text(window, font=("メイリオ", 12), height=5, width=30)
description_text.pack()

# CSVファイルを生成するボタンを作成する
generate_button = tk.Button(window, text="CSV生成", command=generate_csv, font=("メイリオ", 12, "bold"), bg="green", fg="white")
generate_button.pack(pady=10)

# GUIイベントループを開始する
window.mainloop()