import os
from comtypes import client

def pptx_to_pdf(pptx_path, pdf_path):
    powerpoint = client.CreateObject('Powerpoint.Application')
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(pptx_path)
    presentation.SaveAs(pdf_path, 32)  # 32はPDF形式
    presentation.Close()
    powerpoint.Quit()

def convert_pptx_in_folder(input_folder):
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            if file.endswith('.pptx'):
                pptx_file = os.path.join(root, file)
                pdf_file = os.path.splitext(pptx_file)[0] + '.pdf'
                pptx_to_pdf(pptx_file, pdf_file)
                print(f"{pptx_file} を {pdf_file} に変換しました。")

input_folder = r"C:\Users\ME-PC2\OneDrive - Hiroshima University (1)\ドキュメント\4年\MRI脳画像\損傷度\241015‗エジンバラ\20241016_色塗り\20241016_色塗り_convrt_nii"
convert_pptx_in_folder(input_folder)
