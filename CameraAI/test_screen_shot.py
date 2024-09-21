import cv2
import numpy as np
import mss
import time

# スクリーンキャプチャする領域を定義（x, y, 幅, 高さ）
monitor = {"top": 100, "left": 100, "width": 800, "height": 600}

# mssの初期化
sct = mss.mss()

# ウィンドウ名
window_name = "ScreenCapture"

# OpenCVウィンドウを作成
cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)

# ウィンドウを最前面に設定
cv2.setWindowProperty(window_name, cv2.WND_PROP_TOPMOST, 1)

while True:
    # mssを使ってスクリーン領域をキャプチャ
    screenshot = sct.grab(monitor)

    # 画像をnumpy配列に変換
    img = np.array(screenshot)

    # BGRフォーマットに変換（mssがキャプチャする画像はBGRAフォーマットなので）
    img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)

    # OpenCVを使って画像を表示
    cv2.imshow(window_name, img)

    # 各フレームで画像処理を実行（ここにリアルタイム処理のロジックを挿入可能）
    # ...

    # 'q' キーが押されたらループを終了
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# リソースを解放
cv2.destroyAllWindows()



