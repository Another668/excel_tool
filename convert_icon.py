<<<<<<< HEAD
from PIL import Image
import sys
import os

def get_script_dir():
    """获取脚本所在目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def main():
    script_dir = get_script_dir()
    png_path = os.path.join(script_dir, '图标.png')
    ico_path = os.path.join(script_dir, '图标.ico')
    
    print(f'[convert_icon] 脚本目录: {script_dir}')
    print(f'[convert_icon] PNG源文件: {png_path}')
    print(f'[convert_icon] ICO输出文件: {ico_path}')
    
    if not os.path.exists(png_path):
        print(f'[错误] 源文件不存在: {png_path}')
        sys.exit(1)
    
    try:
        img = Image.open(png_path)
        ico_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
        img.save(ico_path, format='ICO', sizes=ico_sizes)
        print(f'[成功] 图标转换完成：{png_path} -> {ico_path}')
        sys.exit(0)
    except Exception as e:
        print(f'[错误] 图标转换失败：{e}')
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == '__main__':
    main()
=======
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
>>>>>>> cad4c8583212cb6a13af68b9f1b547f5ccd4c4e5
