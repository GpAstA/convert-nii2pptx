import os
from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def check_image_sizes(files):
    result_text.delete(1.0, tk.END)  # 結果表示エリアをクリア
    for file_path in files:
        try:
            with Image.open(file_path) as img:
                width, height = img.size
                result_text.insert(tk.END, f"{os.path.basename(file_path)}: 幅={width}px, 高さ={height}px\n")
        except Exception as e:
            result_text.insert(tk.END, f"ファイルを開けませんでした: {os.path.basename(file_path)}, エラー: {e}\n")

def select_files():
    files = filedialog.askopenfilenames(filetypes=[("画像ファイル", "*.jpg *.jpeg *.png *.bmp *.gif")])
    if files:
        check_image_sizes(files)
    else:
        messagebox.showwarning("警告", "ファイルが選択されていません")

# Tkinter GUIの設定
root = tk.Tk()
root.title("画像サイズチェックツール")

# ファイル選択ボタン
select_button = tk.Button(root, text="画像ファイルを選択", command=select_files)
select_button.pack(pady=10)

# 結果表示用のテキストエリア
result_text = scrolledtext.ScrolledText(root, width=60, height=20, wrap=tk.WORD)
result_text.pack(pady=10)

root.mainloop()
