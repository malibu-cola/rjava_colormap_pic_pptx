from PIL import Image
import os
from posixpath import normcase

status = ["NSE", "freezeout", "last"]

Ye = ["010", "013", "016"]

T0 = ["4e9", "7e9", \
    "1e10", "4e10", "7e10", \
        "1e11", "4e11", "7e11", \
            "1e12"]
rho0 = ["1e10", "4e10", "7e10", \
    "1e11", "4e11", "7e11", \
        "1e12", "4e12", "7e12", \
            "1e13", "4e13"]


for w in status:
    for x in Ye:
        for y in T0:
            for z in rho0:
                # 読み込むファイル名を指定
                filename = w + "_Ye_" + x + "_T0_" + y + "_rho0_" + z + ".png"
                # 読み込むファイルの相対パス
                input_path = "./pic/" + filename
                # 読み込むファイルが存在しなければ、飛ばす。
                exsist = os.path.exists(input_path)
                if not exsist:
                    continue
                # ファイルを読み込む
                cat = Image.open(input_path)
                # 画像を切り抜く（左上のx, 左上のy, 右下のx, 右下のy）
                crop_cat = cat.crop((140, 250, 1300, 940))
                # 切り抜いた画像を保存
                crop_cat.save("./pic_edit/" + filename)
