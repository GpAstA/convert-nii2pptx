import fitz  # PyMuPDF
from PIL import Image
import numpy as np
import cv2
import matplotlib.pyplot as plt
import nibabel as nib
import os
import glob
import tkinter as tk
from tkinter import filedialog

# 色の範囲を設定 (最小値と最大値)
target_color_ranges = [
    (np.array([40, 180, 240]), np.array([70, 255, 255])),   # 青
    (np.array([200, 0, 0]), np.array([255, 100, 100])),     # 赤
    (np.array([80, 180, 0]), np.array([120, 255, 100])),    # 緑
    (np.array([200, 180, 0]), np.array([255, 230, 100]))    # 黄
]

def get_pdf_images(pdf_path):
    doc = fitz.open(pdf_path)
    images_list = []

    for page_num in range(doc.page_count):
        page = doc[page_num]
        img = page.get_pixmap()

        # Convert the RGB image to a PIL Image
        pil_img = Image.frombytes("RGB", (img.width, img.height), img.samples)

        # Append the PIL Image to the list
        images_list.append(pil_img)

    doc.close()
    return images_list

def images_to_numpy(images_list):
    return [np.array(img) for img in images_list]

# 色の範囲に基づいてマスクを抽出する関数
def mask_extracted(image_array, target_color_ranges):
    combined_images = []
    
    for i in range(len(image_array)):
        image_rgb = image_array[i]  # 画像を取得

        # マスクの初期化
        mask = np.zeros((image_rgb.shape[0], image_rgb.shape[1]), dtype=np.uint8)

        # 各色の範囲を使ってピクセルを抽出
        for lower_color, upper_color in target_color_ranges:
            color_mask = cv2.inRange(image_rgb, lower_color, upper_color)
            mask = cv2.bitwise_or(mask, color_mask)  # 複数の色範囲を組み合わせる

        # マスクを使って元の画像から対象部分を抽出（1 or 0）
        binary_mask = np.where(mask > 0, 1, 0)  # マスク部分を1、それ以外を0
        combined_images.append(binary_mask)

    return combined_images

# 長い方の両端を切り取る関数
def crop_to_square(image, target_size=540):
    """
    長い方の両端を切り取って正方形にクロップする。
    画像が2次元または3次元（カラー）の場合に対応。
    """
    height, width = image.shape[:2]  # 画像の高さと幅を取得
    
    # モノクロ画像の場合は2次元、カラー画像の場合は3次元
    if len(image.shape) == 3:
        # 縦長か横長かを判定して長辺を540に合わせる
        if height > width:
            # 縦長の場合、高さを540に合わせて中央から両端を切り取る
            crop_start = (height - target_size) // 2
            cropped_image = image[crop_start:crop_start + target_size, :, :]  # 高さを540に
        elif width > height:
            # 横長の場合、幅を540に合わせて中央から両端を切り取る
            crop_start = (width - target_size) // 2
            cropped_image = image[:, crop_start:crop_start + target_size, :]  # 幅を540に
        else:
            # 既に正方形の場合はそのまま
            cropped_image = image
    else:  # 画像が2次元の場合（モノクロ画像の場合）
        if height > width:
            # 縦長の場合、高さを540に合わせて中央から両端を切り取る
            crop_start = (height - target_size) // 2
            cropped_image = image[crop_start:crop_start + target_size, :]  # 高さを540に
        elif width > height:
            # 横長の場合、幅を540に合わせて中央から両端を切り取る
            crop_start = (width - target_size) // 2
            cropped_image = image[:, crop_start:crop_start + target_size]  # 幅を540に
        else:
            # 既に正方形の場合はそのまま
            cropped_image = image
    
    return cropped_image

# 画像を90度単位で回転させる関数
def rotate_image(image, num_rotations):
    """
    画像を90度単位で回転させる。
    num_rotations: 90度単位の回転回数。1なら90度、2なら180度、3なら270度。
    """
    if num_rotations == 1:
        return cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
    elif num_rotations == 2:
        return cv2.rotate(image, cv2.ROTATE_180)
    elif num_rotations == 3:
        return cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)
    else:
        return image  # 回転しない

# リサイズ処理
def resize_images(images, target_size=(540, 540), num_rotations=0):
    resized_images = []
    for image in images:
        # 長い方の両端を切り取って540x540にする
        cropped_image = crop_to_square(image, target_size[0])

        # 幅がtarget_sizeに収まっていない場合のみリサイズ
        if cropped_image.shape[1] != target_size[1]:
            resized_image = cv2.resize(cropped_image, target_size, interpolation=cv2.INTER_NEAREST)
        else:
            resized_image = cropped_image

        # 540x540にリサイズ後、90度単位で回転させる
        rotated_image = rotate_image(resized_image, num_rotations)
        resized_images.append(rotated_image)
    
    return resized_images

# NIfTIファイルに変換して保存
def save_as_nii(images, output_filename="output_image.nii"):
    # z軸方向に画像をスタック
    stacked_images = np.stack(images, axis=-1)

    # NIfTIイメージを作成
    nii_img = nib.Nifti1Image(stacked_images, affine=np.eye(4))

    # NIfTIファイルとして保存
    nib.save(nii_img, output_filename)
    print(f"NIfTIファイルを保存しました: {output_filename}")

# 指定されたフォルダ内のpptxファイルを処理
def process_pptx_in_folder(folder_path):
    # フォルダ内のすべてのpptxファイルを検索
    pptx_files = glob.glob(os.path.join(folder_path, "*.pptx"))

    # 各pptxファイルに対して処理を実行
    for pptx_file in pptx_files:
        print(f"Processing {pptx_file}...")
        
        # 元のファイル名を取得し、出力ファイル名を設定
        base_name = os.path.splitext(os.path.basename(pptx_file))[0]
        output_filename = os.path.join(os.path.dirname(pptx_file), f"convert_{base_name}.nii")

        # PDFファイルを画像に変換
        pdf_images = get_pdf_images(pptx_file)

        # 画像をNumPy配列に変換
        numpy_images = images_to_numpy(pdf_images)

        # 色範囲に基づいてマスクを抽出
        mask_images = mask_extracted(numpy_images, target_color_ranges)

        # 540x540にリサイズ（長い方を両端カット）、90度単位で回転（num_rotationsにより調整）
        num_rotations = 1  # 90度右回り
        resized_images = resize_images(mask_images, target_size=(540, 540), num_rotations=num_rotations)

        # NIfTIファイルとして保存
        save_as_nii(resized_images, output_filename)

# tkinterでフォルダを選択させる
def select_folder():
    root = tk.Tk()
    root.withdraw()  # GUIウィンドウを表示しない
    folder_path = filedialog.askdirectory(title="処理するフォルダを選択してください")
    return folder_path

if __name__ == "__main__":
    # tkinterを使ってフォルダを選択
    folder_path = select_folder()
    if folder_path:
        process_pptx_in_folder(folder_path)
    else:
        print("フォルダが選択されていません。")
