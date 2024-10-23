import matplotlib.pyplot as plt
import numpy as np
import cv2

def display_image_and_get_rgb(image):
    fig, ax = plt.subplots()
    ax.imshow(image)

    # クリックした位置のRGB値を取得するイベントハンドラ
    def onclick(event):
        if event.xdata is not None and event.ydata is not None:
            x, y = int(event.xdata), int(event.ydata)
            rgb = image[y, x, :]  # (y, x)でRGB取得
            print(f"Coordinates: ({x}, {y}) - RGB: {rgb}")

    cid = fig.canvas.mpl_connect('button_press_event', onclick)
    plt.show()

# 画像ファイルを読み込む（例としてCV2で画像を読み込む）
image_path = r"C:\Users\ME-PC2\convert-nii2pptx\nii2pptx\input\check_rgb.png"  # ここに画像のパスを指定
image = cv2.imread(image_path)
image_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

# 画像を表示してクリックした場所のRGBを取得
display_image_and_get_rgb(image_rgb)
