import tkinter as tk
from tkinter import filedialog
import numpy as np
import nibabel as nib
import cv2
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

# 描画状態管理クラス
class DrawingApp:
    def __init__(self, root, images, original_filename):
        self.root = root
        self.images = images  # z軸スライスのリスト (3D配列)
        self.current_z = 0  # 表示中のzスライス
        self.undo_stack = []  # 作業の履歴
        self.tool = "pen"  # pen, fill, eraseのどれか
        self.original_filename = original_filename  # 元のファイル名
        
        # UI構成
        self.canvas_frame = tk.Frame(self.root)
        self.canvas_frame.pack(side=tk.TOP)
        self.control_frame = tk.Frame(self.root)
        self.control_frame.pack(side=tk.BOTTOM)
        
        # Matplotlibのfigureを使ったキャンバス
        self.fig, self.ax = plt.subplots()
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.canvas_frame)
        self.canvas.get_tk_widget().pack()
        self.update_canvas()

        # ボタン配置
        pen_button = tk.Button(self.control_frame, text="ペン", command=self.set_pen_tool)
        fill_button = tk.Button(self.control_frame, text="塗りつぶし", command=self.set_fill_tool)
        erase_button = tk.Button(self.control_frame, text="消しゴム", command=self.set_erase_tool)
        undo_button = tk.Button(self.control_frame, text="戻す", command=self.undo_last_action)
        prev_button = tk.Button(self.control_frame, text="前のスライス", command=self.prev_slice)
        next_button = tk.Button(self.control_frame, text="次のスライス", command=self.next_slice)
        save_button = tk.Button(self.control_frame, text="保存", command=self.save_nii_file)

        pen_button.grid(row=0, column=0)
        fill_button.grid(row=0, column=1)
        erase_button.grid(row=0, column=2)
        undo_button.grid(row=0, column=3)
        prev_button.grid(row=0, column=4)
        next_button.grid(row=0, column=5)
        save_button.grid(row=0, column=6)
        
        # マウスイベントのバインディング
        self.canvas.mpl_connect("button_press_event", self.on_press)
        self.canvas.mpl_connect("motion_notify_event", self.on_drag)
        self.drawing = False

    def update_canvas(self):
        """現在のzスライスを描画する"""
        self.ax.clear()
        self.ax.imshow(self.images[:, :, self.current_z], cmap="gray", aspect='auto')
        self.canvas.draw()

    def set_pen_tool(self):
        self.tool = "pen"

    def set_fill_tool(self):
        self.tool = "fill"

    def set_erase_tool(self):
        self.tool = "erase"

    def on_press(self, event):
        """マウスクリック時のイベント"""
        self.drawing = True
        self.start_x, self.start_y = int(event.xdata), int(event.ydata)
        self.undo_stack.append(np.copy(self.images[:, :, self.current_z]))  # 現在の状態を保存

    def on_drag(self, event):
        """マウスドラッグ時のイベント"""
        if self.drawing and event.xdata and event.ydata:
            x, y = int(event.xdata), int(event.ydata)

            # zスライスの画像を取得し、2D配列に変換
            slice_img = self.images[:, :, self.current_z]

            # スライスが3D配列であれば2Dに変換する
            if len(slice_img.shape) > 2:
                slice_img = slice_img[:, :, 0]  # 必要なら3次元を2次元に

            # OpenCVで扱うためにデータ型をuint8に変換
            slice_img = slice_img.astype(np.uint8)

            # 描画処理
            if self.tool == "pen":
                cv2.line(slice_img, (self.start_x, self.start_y), (x, y), 1, 2)  # 色1、太さ2
            elif self.tool == "erase":
                cv2.line(slice_img, (self.start_x, self.start_y), (x, y), 0, 2)  # 色0、太さ2

            # 変更を元の配列に反映
            self.images[:, :, self.current_z] = slice_img

            # 次の描画位置を更新
            self.start_x, self.start_y = x, y

            # 描画を更新
            self.update_canvas()


    def undo_last_action(self):
        """ひとつ前の作業を戻す"""
        if self.undo_stack:
            self.images[:, :, self.current_z] = self.undo_stack.pop()
            self.update_canvas()

    def next_slice(self):
        """次のスライスに進む"""
        if self.current_z < self.images.shape[2] - 1:
            self.current_z += 1
            self.update_canvas()

    def prev_slice(self):
        """前のスライスに戻る"""
        if self.current_z > 0:
            self.current_z -= 1
            self.update_canvas()

    def save_nii_file(self):
        """修正後の画像をNIfTIファイルとして保存する"""
        save_as_nii(self.images, self.original_filename)
        
def load_nii_file(filepath):
    """NIfTIファイルを読み込んで画像データを取得"""
    nii_data = nib.load(filepath)
    images = nii_data.get_fdata()
    
    # 画像データの形状を確認
    print(f"Loaded NIfTI data shape: {images.shape}")  # 形状確認
    images = np.round(images).astype(np.uint8)  # uint8 に変換
    return images

# NIfTIファイルに変換して保存
def save_as_nii(images, original_filename, output_dir=""):
    """
    NIfTIファイルを保存する。
    :param images: 保存する画像データ（3D配列）
    :param original_filename: 元のNIfTIファイルの名前（ファイル名のみ、パスを含まない）
    :param output_dir: 保存先のディレクトリ
    """
    # z軸方向に画像をスタック
    stacked_images = np.stack(images, axis=-1)

    # NIfTIイメージを作成
    nii_img = nib.Nifti1Image(stacked_images, affine=np.eye(4))

    # 元のファイル名から新しい名前を生成
    base_name = os.path.splitext(os.path.basename(original_filename))[0]  # ファイル名のみ取得
    output_filename = f"drawed_{base_name}.nii"  # 新しいファイル名を作成

    # ディレクトリが指定されていればそのパスに、指定されていなければカレントディレクトリに保存
    if output_dir:
        output_filepath = os.path.join(output_dir, output_filename)
    else:
        output_filepath = output_filename

    # NIfTIファイルとして保存
    nib.save(nii_img, output_filepath)
    print(f"NIfTIファイルを保存しました: {output_filepath}")

def select_file():
    """NIfTIファイルを選択する"""
    filepath = filedialog.askopenfilename(title="NIfTIファイルを選択", filetypes=[("NIfTI files", "*.nii *.nii.gz")])
    return filepath

if __name__ == "__main__":
    root = tk.Tk()
    root.title("NIfTI 画像編集ツール")


    # ファイルを選択
    file_path = select_file()
    if file_path:
        images = load_nii_file(file_path)
        
        
        # 編集作業後のNIfTIファイルの保存
        app = DrawingApp(root, images, file_path)

    root.mainloop()
