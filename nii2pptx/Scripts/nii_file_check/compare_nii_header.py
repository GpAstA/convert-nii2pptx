import nibabel as nib

# 比較したい2つのファイルパスを指定
file_path1 = r"C:\Users\ME-PC2\OneDrive - Hiroshima University\ドキュメント\01.241014_MRI\H000276\002HeadERFLAIRPROPELLER_0003\s2023-03-31_14-00-140952-00001-00001-1.nii"
file_path2 = r"C:\Users\ME-PC2\OneDrive - Hiroshima University\ドキュメント\01.241014_MRI\H000276\002HeadERFLAIRPROPELLER_0003\s2023-03-31_14-00-140952-00001-00001-1_NoName_All_mask.nii"
# file_path3 = r"C:/Users/ME-PC2/OneDrive - Hiroshima University (1)/ドキュメント/4年/MRI脳画像/損傷度/241015‗エジンバラ/20241016_色塗り/20241016_色塗り_convrt_nii/色塗り/H000276〇/002HeadERFLAIRPROPELLER_0003/convert_auto_drawed_s2023-03-31_14-00-140952-00001-00001-1_済み.nii"
file_path3 = r"C:\Users\ME-PC2\OneDrive - Hiroshima University (1)\ドキュメント\4年\MRI脳画像\損傷度\241015‗エジンバラ\20241016_色塗り\20241016_色塗り_convrt_nii\色塗り\H000276〇\002HeadERFLAIRPROPELLER_0003\header_fixed_convert_auto_drawed_s2023-03-31_14-00-140952-00001-00001-1_済み.nii"


##標準返還後
file_path2 = r"C:\Users\ME-PC2\OneDrive - Hiroshima University\デスクトップ\雑用\M001\002HeadERFLAIRPROPELLER_0003\ws2023-03-31_14-00-140952-00001-00001-1_NoName_All_mask.nii"
file_path3 = r"C:\Users\ME-PC2\OneDrive - Hiroshima University\デスクトップ\雑用\M001\002HeadERFLAIRPROPELLER_0003\ws2023-03-31_14-00-140952-00001-00001-1_header_fixed_convert_auto_drawed_.nii"


# ファイルを読み込み、ヘッダ情報を取得
img1 = nib.load(file_path1)
img2 = nib.load(file_path2)
img3 = nib.load(file_path3)


header1 = img1.header
header2 = img2.header
header3 = img3.header


# ヘッダ情報をディクショナリ形式で取得
header_dict1 = dict(header1)
header_dict2 = dict(header2)
header_dict3 = dict(header3)


# 差分を表示する関数
def compare_three_headers(header1, header2, header3):
    for key in header1:
        if key in header2 and key in header3:
            value1 = header1[key]
            value2 = header2[key]
            value3 = header3[key]

            # 3つの値がすべて同じでない場合にのみ出力
            if not (value1 == value2).all() or not (value2 == value3).all() or not (value1 == value3).all():
                print(f"{key}:")
                print(f"  File 1: {value1}")
                print(f"  File 2: {value2}")
                print(f"  File 3: {value3}")
                print("-" * 40)

print("Differences between headers of file 1, file 2, and file 3:")
compare_three_headers(header1, header2, header3)
