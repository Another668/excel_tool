from PIL import Image
import sys

try:
    img = Image.open('图标.png')
    ico_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    img.save('图标.ico', format='ICO', sizes=ico_sizes)
    print('图标转换成功：图标.png -> 图标.ico')
    sys.exit(0)
except Exception as e:
    print(f'图标转换失败：{e}')
    sys.exit(1)
