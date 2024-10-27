# -*- coding: utf-8 -*-
"""
Created on Sun Oct 27 18:25:43 2024

@author: ME-PC2
"""

import nibabel as nib
import numpy as np
import os


# 元のファイルパス
file_path1 = r"C:\Users\ME-PC2\OneDrive - Hiroshima University\ドキュメント\01.241014_MRI\H000276\002HeadERFLAIRPROPELLER_0003\s2023-03-31_14-00-140952-00001-00001-1.nii"
file_path3 = r"C:/Users/ME-PC2/OneDrive - Hiroshima University (1)/ドキュメント/4年/MRI脳画像/損傷度/241015‗エジンバラ/20241016_色塗り/20241016_色塗り_convrt_nii/色塗り/H000276〇/002HeadERFLAIRPROPELLER_0003/convert_auto_drawed_s2023-03-31_14-00-140952-00001-00001-1_済み.nii"

# 1. ファイルの読み込み
img1 = nib.load(file_path1)
img3 = nib.load(file_path3)

# 2. 画像データのコピー（file_path3の画像データを使用）
data3 = img3.get_fdata()

# 3. file_path1のヘッダ情報をコピーして新しいヘッダ作成
header1 = img1.header.copy()

# 4. 空間参照情報を更新（qform_codeとsform_codeを2に設定）
header1['qform_code'] = 2
header1['sform_code'] = 2

# 5. 新しい画像データを元にfile_path3のヘッダ情報を適用して保存
affine = img1.affine  # 座標変換行列もfile_path1に合わせる場合
new_img = nib.Nifti1Image(data3, affine, header=header1)

# 出力ファイルのパスをfile_path3と同階層に設定
directory = os.path.dirname(file_path3)  # file_path3と同じディレクトリ
original_filename = os.path.basename(file_path3)  # file_path3のファイル名
new_filename = f"header_fixed_{original_filename}"
new_file_path = os.path.join(directory, new_filename)

# 新しいファイルとして保存
nib.save(new_img, new_file_path)

print(f"Updated file saved to {new_file_path}")
