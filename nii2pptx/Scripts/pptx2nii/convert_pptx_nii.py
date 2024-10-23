import os
import comtypes.client  # PowerPoint to PDF conversion (only works on Windows)
import fitz  # PyMuPDF
from PIL import Image
import numpy as np
import cv2
import nibabel as nib
import tkinter as tk
from tkinter import filedialog

# 色の範囲を設定 (最小値と最大値)
target_color_ranges = [
    (np.array([40, 180, 240]), np.array([70, 255, 255])),   # 青
    (np.array([200, 0, 0]), np.array([255, 100, 100])),     # 赤
    (np.array([80, 180, 0]), np.array([120, 255, 100])),    # 緑
    (np.array([200, 180, 0]), np.array([255, 230, 100]))    # 黄
]

# PowerPointファイルをPDFに変換する関数（Windows限定）
def convert_pptx_to_pdf(pptx_path, output_pdf_path):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    ppt = powerpoint.Presentations.Open(pptx_path)
    ppt.SaveAs(output_pdf_path, 32)  # 32はPDFとして保存
    ppt.Close()
    powerpoint.Quit()

# PDFファイルを画像リストに変換する関数
def get_pdf_images(pdf_path):
    doc = fitz.open(pdf_path)
    images_list = []
    for page_num in range(doc.page_count):
        page = doc[page_num]
        img = page.get_pixmap()
        pil_img = Image.frombytes("RGB", (img.width, img.height), img.samples)
        images_list.append(pil_img)
    doc.close()
    return images_list

# 画像リストをNumPy配列に変換する関数
def images_to_numpy(images_list):
    return [np.array(img) for img in images_list]

# 色の範囲に基づいてマスクを抽出する関数
def mask_extracted(image_array, target_color_ranges):
    combined_images = []
    for image_rgb in image_array:
        mask = np.zeros((image_rgb.shape[0], image_rgb.shape[1]), dtype=np.uint8)
        for lower_color, upper_color in target_color_ranges:
            color_mask = cv2.inRange(image_rgb, lower_color, upper_color)
            mask = cv2.bitwise_or(mask, color_mask)
        binary_mask = np.where(mask > 0, 1, 0)
        combined_images.append(binary_mask)
    return combined_images

# 画像を540x540にクロップする関数
def crop_to_square(image, target_size=540):
    height, width = image.shape[:2]
    if len(image.shape) == 3:
        if height > width:
            crop_start = (height - target_size) // 2
            cropped_image = image[crop_start:crop_start + target_size, :, :]
        elif width > height:
            crop_start = (width - target_size) // 2
            cropped_image = image[:, crop_start:crop_start + target_size, :]
        else:
            cropped_image = image
    else:
        if height > width:
            crop_start = (height - target_size) // 2
            cropped_image = image[crop_start:crop_start + target_size, :]
        elif width > height:
            crop_start = (width - target_size) // 2
            cropped_image = image[:, crop_start:crop_start + target_size]
        else:
            cropped_image = image
    return cropped_image

# 画像を90度単位で回転させる関数
def rotate_image(image, num_rotations):
    if num_rotations == 1:
        return cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
    elif num_rotations == 2:
        return cv2.rotate(image, cv2.ROTATE_180)
    elif num_rotations == 3:
        return cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)
    else:
        return image

# 画像をリサイズして回転する関数
def resize_images(images, target_size=(540, 540), num_rotations=0):
    resized_images = []
    for image in images:
        cropped_image = crop_to_square(image, target_size[0])
        resized_image = cv2.resize(cropped_image, target_size, interpolation=cv2.INTER_NEAREST)
        rotated_image = rotate_image(resized_image, num_rotations)
        resized_images.append(rotated_image)
    return resized_images

# NIfTIファイルに変換して保存する関数
def save_as_nii(images, output_filename="output_image.nii"):
    stacked_images = np.stack(images, axis=-1)
    nii_img = nib.Nifti1Image(stacked_images, affine=np.eye(4))
    nib.save(nii_img, output_filename)
    print(f"NIfTIファイルを保存しました: {output_filename}")

# tkinterでフォルダを選択させる関数
def select_folder():
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="処理するフォルダを選択してください")
    return folder_path

# 指定フォルダ内のすべてのpptxファイルを処理する関数
def process_pptx_in_folder(folder_path):
    for dirpath, _, filenames in os.walk(folder_path):
        for filename in filenames:
            if filename.endswith(".pptx"):
                pptx_file = os.path.join(dirpath, filename)
                base_name = os.path.splitext(os.path.basename(pptx_file))[0]
                output_pdf_path = os.path.join(dirpath, f"{base_name}.pdf")
                output_nii_path = os.path.join(dirpath, f"convert_{base_name}.nii")

                # PPTXをPDFに変換
                convert_pptx_to_pdf(pptx_file, output_pdf_path)

                # PDFを画像に変換して処理
                pdf_images = get_pdf_images(output_pdf_path)
                numpy_images = images_to_numpy(pdf_images)
                mask_images = mask_extracted(numpy_images, target_color_ranges)
                resized_images = resize_images(mask_images, target_size=(540, 540), num_rotations=1)

                # NIfTI形式で保存
                save_as_nii(resized_images, output_nii_path)

                # 変換後のPDFは不要であれば削除
                os.remove(output_pdf_path)

if __name__ == "__main__":
    folder_path = select_folder()
    if folder_path:
        process_pptx_in_folder(folder_path)
    else:
        print("フォルダが選択されていません。")
