import nibabel as nib
import os
from tkinter import Tk, filedialog

def check_nii_file_size(nii_file):
    # NIfTIファイルを読み込む
    img = nib.load(nii_file)
    
    # データの形状を取得（各次元のサイズ）
    data_shape = img.shape
    print(f"Image shape: {data_shape}")
    
    # データ型を取得
    data_type = img.get_data_dtype()
    print(f"Data type: {data_type}")
    
    # データ全体のサイズを計算
    data_size = img.get_fdata().nbytes
    print(f"Total data size: {data_size / (1024 ** 2):.2f} MB")  # メガバイトで表示

def select_nii_file():
    # tkinterを使ってファイル選択ダイアログを表示
    root = Tk()
    root.withdraw()  # メインウィンドウを隠す
    nii_file = filedialog.askopenfilename(title="Select NIfTI file", filetypes=[("NIfTI files", "*.nii *.nii.gz")])
    
    if not nii_file:
        print("No file selected.")
        return None
    
    return nii_file

if __name__ == "__main__":
    # ファイル選択
    nii_file = select_nii_file()

    if nii_file:
        # 選択したNIfTIファイルを処理
        check_nii_file_size(nii_file)
