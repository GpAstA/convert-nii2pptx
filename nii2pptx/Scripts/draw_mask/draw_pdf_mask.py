import tkinter as tk
from tkinter import filedialog
import numpy as np
import fitz  # PyMuPDF
import cv2
from PIL import Image, ImageTk
import os

class PDFDrawingApp:
    def __init__(self, root, pdf_path):
        self.root = root
        self.pdf_path = pdf_path
        self.doc = fitz.open(pdf_path)
        self.page_index = 0
        self.page_count = self.doc.page_count
        self.images = self.load_pdf_page(self.page_index)  # 最初のページをロード

        self.undo_stack = []  # 作業の履歴
        self.tool = "pen"  # "pen", "erase", "fill" のどれか
        self.pen_color = (102, 204, 0)  # デフォルトペンの色を緑に設定
        self.points = []  # 領域を囲むための座標リスト

        # キャンバス設定
        self.canvas_frame = tk.Frame(self.root)
        self.canvas_frame.pack(side=tk.TOP)
        self.control_frame = tk.Frame(self.root)
        self.control_frame.pack(side=tk.BOTTOM)

        # 描画用のキャンバスを作成
        self.canvas = tk.Canvas(self.canvas_frame, width=self.images.width, height=self.images.height)
        self.canvas.pack()
        self.tk_img = ImageTk.PhotoImage(self.images)
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_img)

        # ボタンの配置
        pen_button = tk.Button(self.control_frame, text="ペン", command=self.set_pen_tool)
        erase_button = tk.Button(self.control_frame, text="消しゴム", command=self.set_erase_tool)
        fill_button = tk.Button(self.control_frame, text="塗りつぶし", command=self.set_fill_tool)
        undo_button = tk.Button(self.control_frame, text="戻す", command=self.undo_last_action)
        prev_button = tk.Button(self.control_frame, text="前のページ", command=self.prev_page)
        next_button = tk.Button(self.control_frame, text="次のページ", command=self.next_page)
        save_button = tk.Button(self.control_frame, text="保存", command=self.save_pdf)

        pen_button.grid(row=0, column=0)
        erase_button.grid(row=0, column=1)
        fill_button.grid(row=0, column=2)
        undo_button.grid(row=0, column=3)
        prev_button.grid(row=0, column=4)
        next_button.grid(row=0, column=5)
        save_button.grid(row=0, column=6)

        # マウスイベントのバインディング
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.start_x, self.start_y = None, None
        self.drawing = False

    def load_pdf_page(self, page_index):
        """PDFのページを画像として読み込む"""
        page = self.doc.load_page(page_index)
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img

    def update_canvas(self):
        """キャンバスの画像を更新"""
        self.tk_img = ImageTk.PhotoImage(self.images)
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.tk_img)

    def set_pen_tool(self):
        self.tool = "pen"

    def set_erase_tool(self):
        self.tool = "erase"
        self.undo_stack.append(self.images.copy())  # 元の状態をバックアップ

    def set_fill_tool(self):
        """塗りつぶしツールを有効化"""
        self.tool = "fill"

    def on_click(self, event):
        """マウスクリック時のイベント"""
        x, y = event.x, event.y

        if self.tool == "fill":
            # 塗りつぶし用の領域を初期化
            self.points = [(x, y)]

        self.start_x, self.start_y = x, y

    def on_drag(self, event):
        """マウスドラッグ時のイベント"""
        x, y = event.x, event.y
        if self.start_x and self.start_y:
            # PIL画像をOpenCV画像に変換
            cv_img = np.array(self.images)

            # ペンの色をBGRタプルに変換
            pen_color_bgr = (int(self.pen_color[2]), int(self.pen_color[1]), int(self.pen_color[0]))  # RGBからBGRへ

            if self.tool == "pen":
                cv2.line(cv_img, (self.start_x, self.start_y), (x, y), pen_color_bgr, 5)  # 緑色で描画
            elif self.tool == "erase":
                cv2.line(cv_img, (self.start_x, self.start_y), (x, y), (255, 255, 255), 10)  # 白で消去
            elif self.tool == "fill":
                # ドラッグ時に領域を作成
                self.points.append((x, y))
                cv2.line(cv_img, (self.start_x, self.start_y), (x, y), pen_color_bgr, 2)  # 領域を表示

            # 変更後の画像をPILに変換
            self.images = Image.fromarray(cv_img)
            self.update_canvas()

        self.start_x, self.start_y = x, y

    def on_release(self, event):
        """マウスボタンを離した時のイベント"""
        if self.tool == "fill" and self.points:
            # 領域が閉じられたら塗りつぶし
            cv_img = np.array(self.images)

            # ペンの色をBGRタプルに変換
            fill_color_bgr = (int(self.pen_color[2]), int(self.pen_color[1]), int(self.pen_color[0]))  # RGBからBGR

            # 塗りつぶす領域を閉じてポリゴンとして塗りつぶす
            points_array = np.array(self.points, dtype=np.int32)
            cv2.fillPoly(cv_img, [points_array], fill_color_bgr)

            # 塗りつぶした後、画像を更新
            self.images = Image.fromarray(cv_img)
            self.update_canvas()

            # 塗りつぶし終了後、ポイントをリセット
            self.points = []

    def undo_last_action(self):
        """ひとつ前の作業を戻す"""
        if self.undo_stack:
            # 最後の操作をスタックから取り出し
            self.images = self.undo_stack.pop()
            self.update_canvas()

    def next_page(self):
        """次のPDFページに進む"""
        if self.page_index < self.page_count - 1:
            self.page_index += 1
            self.images = self.load_pdf_page(self.page_index)
            self.update_canvas()

    def prev_page(self):
        """前のPDFページに戻る"""
        if self.page_index > 0:
            self.page_index -= 1
            self.images = self.load_pdf_page(self.page_index)
            self.update_canvas()

    def save_pdf(self):
        """現在の状態をPDFとして保存"""
        output_pdf_path = f"drawed_{os.path.basename(self.pdf_path)}"
        pdf_writer = fitz.open()

        for i in range(self.page_count):
            page = self.doc.load_page(i)
            pix = page.get_pixmap()

            # 画像に描画した部分を反映させたページ画像を取得
            if i == self.page_index:
                img = np.array(self.images)
                pix = fitz.Pixmap(fitz.csRGB, fitz.Matrix(pix.width, pix.height), img.tobytes())

            # 新しいPDFにページを追加
            pdf_writer.insert_pdf(self.doc, from_page=i, to_page=i)

        pdf_writer.save(output_pdf_path)
        pdf_writer.close()
        print(f"PDFを保存しました: {output_pdf_path}")

def select_pdf_file():
    """PDFファイルを選択する"""
    filepath = filedialog.askopenfilename(title="PDFファイルを選択", filetypes=[("PDF files", "*.pdf")])
    return filepath

if __name__ == "__main__":
    root = tk.Tk()
    root.title("PDF 描画ツール")

    # PDFファイルを選択
    file_path = select_pdf_file()
    if file_path:
        app = PDFDrawingApp(root, file_path)

    root.mainloop()
