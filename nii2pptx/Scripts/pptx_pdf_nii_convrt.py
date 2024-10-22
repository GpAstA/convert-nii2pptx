import os
import fitz  # PyMuPDF
from PIL import Image
import numpy as np
import cv2
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog
import win32com.client  # Windows環境でのPPTX→PDF変換用


# PPTXをPDFに変換する関数 (Windows専用)
def pptx_to_pdf(pptx_path, output_pdf_path):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1

    # PPTXファイルを開く
    presentation = powerpoint.Presentations.Open(pptx_path)

    # PDFとして保存
    presentation.SaveAs(output_pdf_path, 32)  # 32はPDFフォーマット
    presentation.Close()
    powerpoint.Quit()


# PDFを画像リストに変換する関数
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


# フォルダ内の全てのPPTXファイルを探索してPDFに変換
def process_pptx_files(input_dir, output_dir):
    # 入力フォルダ内のすべてのPPTXファイルを探索
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            # .pptxファイル以外をスキップ
            if not file.lower().endswith(".pptx"):
                print(f"Skipping non-pptx file: {file}")
                continue

            pptx_file = os.path.join(root, file)
            output_pdf_path = os.path.join(output_dir, file.replace(".pptx", ".pdf"))

            # ファイルが存在するか確認
            if not os.path.exists(pptx_file):
                print(f"File not found: {pptx_file}")
                continue

            print(f"Processing: {pptx_file}")
            # PPTXをPDFに変換
            try:
                pptx_to_pdf(pptx_file, output_pdf_path)
                print(f"Converted {pptx_file} to PDF.")
            except Exception as e:
                print(f"Error converting {pptx_file}: {str(e)}")


# tkを使用してフォルダを選択
def select_input_output_folders():
    root = Tk()
    root.withdraw()  # メインウィンドウを隠す

    # 入力フォルダの選択
    input_dir = filedialog.askdirectory(title="Select input folder (PPTX)")
    if not input_dir:
        print("No input folder selected. Exiting.")
        return None, None

    # 出力フォルダの選択
    output_dir = filedialog.askdirectory(title="Select output folder (PDF)")
    if not output_dir:
        print("No output folder selected. Exiting.")
        return None, None

    return input_dir, output_dir


if __name__ == "__main__":
    # フォルダの選択
    input_dir, output_dir = select_input_output_folders()

    if input_dir and output_dir:
        # 入力フォルダ内のPPTXファイルをすべて処理
        process_pptx_files(input_dir, output_dir)
