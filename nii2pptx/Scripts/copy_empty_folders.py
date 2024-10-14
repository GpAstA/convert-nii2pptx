import os
import tkinter as tk
from tkinter import filedialog

def main():
    # Tkinterのルートウィンドウを作成して非表示にする
    root = tk.Tk()
    root.withdraw()

    # ソースフォルダを選択
    source_folder = filedialog.askdirectory(title="ソースフォルダを選択してください")
    if not source_folder:
        print("ソースフォルダが選択されていません。")
        return

    # 出力先フォルダを選択
    output_folder = filedialog.askdirectory(title="出力先フォルダを選択してください")
    if not output_folder:
        print("出力先フォルダが選択されていません。")
        return

    # ソースフォルダ内の直下のアイテムを取得
    items = os.listdir(source_folder)

    for item in items:
        item_path = os.path.join(source_folder, item)
        if os.path.isdir(item_path):
            # 出力先に同名のフォルダを作成
            new_dir_path = os.path.join(output_folder, item)
            if not os.path.exists(new_dir_path):
                os.mkdir(new_dir_path)
                print(f"ディレクトリを作成しました: {new_dir_path}")
        else:
            # ファイルの場合は何もしない
            continue

    print("選択したフォルダの直下にあるフォルダのみが空でコピーされました。")

if __name__ == "__main__":
    main()
