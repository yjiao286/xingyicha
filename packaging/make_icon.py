"""Convert 星小纪.png into a multi-size .ico for the exe / installer icon.

构建期脚本（CI 与 build_windows.bat 共用），需要 pillow。"""
import os

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, '..', '星小纪.png')
DST = os.path.join(HERE, 'star.ico')

img = Image.open(SRC).convert('RGBA')
img.save(DST, sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
print('icon written:', DST)
