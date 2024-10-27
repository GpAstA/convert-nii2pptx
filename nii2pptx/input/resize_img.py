import os
from PIL import Image, ImageOps
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

def check_and_resize_images(files):
    result_text.delete(1.0, tk.END)  # 結果表示エリアをクリア
    for file_path in files:
        try:
            with Image.open(file_path) as img:
                # 画像サイズ取得
                original_width, original_height = img.size
                # 背景色の指定
                background_color = (104, 187, 227)
                # 画像を480x480に余白付きでフィットさせる
                resized_img = ImageOps.pad(img, (480, 480), color=background_color)
                
                # 保存先ファイルパスの設定（同階層に保存）
                base_name = os.path.splitext(os.path.basename(file_path))[0]  # 拡張子を除くファイル名
                save_path = os.path.join(os.path.dirname(file_path), f"resized_{base_name}.jpg")
                # JPEG形式で保存
                resized_img = resized_img.convert("RGB")  # RGBモードに変換 (JPEG形式のため)
                resized_img.save(save_path, "JPEG")
                # 結果表示
                result_text.insert(tk.END, f"{base_name}: 幅={original_width}px, 高さ={original_height}px -> 480x480にリサイズし保存しました: {save_path}\n")
        except Exception as e:
            result_text.insert(tk.END, f"ファイルを処理できませんでした: {os.path.basename(file_path)}, エラー: {e}\n")

def select_files():
    files = filedialog.askopenfilenames(filetypes=[("画像ファイル", "*.jpg *.jpeg *.png *.bmp *.gif")])
    if files:
        check_and_resize_images(files)
    else:
        messagebox.showwarning("警告", "ファイルが選択されていません")

# Tkinter GUIの設定
root = tk.Tk()
root.title("画像リサイズツール")

# ファイル選択ボタン
select_button = tk.Button(root, text="画像ファイルを選択", command=select_files)
select_button.pack(pady=10)

# 結果表示用のテキストエリア
result_text = scrolledtext.ScrolledText(root, width=60, height=20, wrap=tk.WORD)
result_text.pack(pady=10)

root.mainloop()
