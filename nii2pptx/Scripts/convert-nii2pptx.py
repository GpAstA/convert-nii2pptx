import os
import nibabel as nib
import numpy as np
import imageio
from pptx import Presentation
from pptx.util import Inches, Pt

SLIDE_WIDTH, SLIDE_HEIGHT = 9144000, 6858000
IMG_CENTER_X, IMG_CENTER_Y = SLIDE_WIDTH / 2, SLIDE_HEIGHT / 2
SLIDE_ASPECT_RATIO = SLIDE_WIDTH / SLIDE_HEIGHT

def add_slide(prs):
    blank_slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_slide_layout)
    return slide

def add_picture(slide, save_png, i):
    from PIL import Image
    im = Image.open(save_png)
    im_width, im_height = im.size
    aspect_ratio = im_width / im_height
    if aspect_ratio > SLIDE_ASPECT_RATIO:
        img_display_width = SLIDE_WIDTH
        img_display_height = img_display_width / aspect_ratio
    else:
        img_display_height = SLIDE_HEIGHT
        img_display_width = img_display_height * aspect_ratio
    left = IMG_CENTER_X - img_display_width / 2
    top = IMG_CENTER_Y - img_display_height / 2
    slide.shapes.add_picture(save_png, left, top, width=img_display_width if aspect_ratio > SLIDE_ASPECT_RATIO else None, height=None if aspect_ratio > SLIDE_ASPECT_RATIO else img_display_height)
    txbox = slide.shapes.add_textbox(left + 10, top + 10, Inches(1), Inches(1))
    tf = txbox.text_frame
    tf.text = 'z={}'.format(i)
    tf.paragraphs[0].font.size = Pt(28)

def create_pptx_for_nii(nii_file, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    base_filename = os.path.basename(nii_file).replace('.nii', '.pptx')
    save_name = os.path.join(output_dir, base_filename)

    prs = Presentation()
    prs.slide_width = SLIDE_WIDTH
    prs.slide_height = SLIDE_HEIGHT

    img = nib.load(nii_file)
    img_fdata = img.get_fdata()
    z_slices = img.shape[2]

    for i in range(z_slices):
        slice_data = img_fdata[:, :, i]
        
        # スライスが2Dでない場合はスキップ
        if len(slice_data.shape) != 2:
            print(f"Skipping non-2D slice at index {i} in file {nii_file}")
            continue
        
        slice_data = np.rot90(slice_data)
        slice_data = ((slice_data - np.min(slice_data)) / (np.max(slice_data) - np.min(slice_data)) * 255).astype(np.uint8)

        save_png = os.path.join(output_dir, f'z={i}.png')
        imageio.imwrite(save_png, slice_data)

        slide = add_slide(prs)
        add_picture(slide, save_png, i)

        os.remove(save_png)

    prs.save(save_name)
    print(f"ファイルの書き出し完了しました: {save_name}")

def find_nii_and_create_pptx(input_dir, output_base_dir):
    for root, dirs, files in os.walk(input_dir):
        for file in files:
            # ファイル名が 's' で始まり '.nii' で終わるものをフィルタリング
            if file.startswith('s') and file.endswith('.nii'):
                nii_file = os.path.join(root, file)
                relative_path = os.path.relpath(root, input_dir)
                output_dir = os.path.join(output_base_dir, relative_path)
                create_pptx_for_nii(nii_file, output_dir)

if __name__ == "__main__":
    from tkinter import filedialog
    input_dir = filedialog.askdirectory(title="フォルダを選択してください")
    output_dir = filedialog.askdirectory(title="出力フォルダを選択してください")
    find_nii_and_create_pptx(input_dir, output_dir)
