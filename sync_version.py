"""
版本同步工具
从README.md读取最新版本号，并更新到Python主程序文件中
"""

import re
import sys
import os


def get_script_dir():
    """获取脚本所在目录"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def read_version_from_readme(readme_path):
    """从README.md读取最新版本号"""
    try:
        with open(readme_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 匹配版本号模式：### [v3.3] - 2026-05-15
        version_pattern = r'### \[v(\d+\.\d+)\]'
        matches = re.findall(version_pattern, content)
        
        if matches:
            latest_version = matches[0]  # 第一个匹配项是最新版本
            print(f"[信息] 从README.md读取到版本号: {latest_version}")
            return latest_version
        
        # 备用方案：检查徽章
        badge_pattern = r'badge/version-(\d+\.\d+)-blue'
        badge_match = re.search(badge_pattern, content)
        if badge_match:
            latest_version = badge_match.group(1)
            print(f"[信息] 从README.md徽章读取到版本号: {latest_version}")
            return latest_version
        
        print("[警告] 未找到版本号信息")
        return None
    except Exception as e:
        print(f"[错误] 读取README.md失败: {e}")
        return None


def update_version_in_python(file_path, new_version):
    """更新Python文件中的版本号"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        
        # 更新 __version__ 变量
        content = re.sub(
            r'__version__\s*=\s*["\'][\d.]+["\']',
            f'__version__ = "{new_version}"',
            content
        )
        
        # 更新文档注释中的版本号
        content = re.sub(
            r'版本:\s*[\d.]+',
            f'版本: {new_version}',
            content
        )
        
        # 更新徽章中的版本号（如果有）
        content = re.sub(
            r'badge/version-[\d.]+-blue',
            f'badge/version-{new_version}-blue',
            content
        )
        
        if content != original_content:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"[成功] 已更新版本号到: {new_version}")
            return True
        else:
            print(f"[信息] 版本号已是最新: {new_version}")
            return False
            
    except Exception as e:
        print(f"[错误] 更新Python文件失败: {e}")
        return False


def verify_version_in_python(file_path, expected_version):
    """验证Python文件中的版本号是否正确"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 检查 __version__ 变量
        version_match = re.search(r'__version__\s*=\s*["\']([\d.]+)["\']', content)
        
        if version_match:
            actual_version = version_match.group(1)
            if actual_version == expected_version:
                print(f"[成功] 版本验证通过: {actual_version}")
                return True
            else:
                print(f"[错误] 版本不匹配: 期望 {expected_version}, 实际 {actual_version}")
                return False
        else:
            print("[错误] 未找到版本号变量")
            return False
            
    except Exception as e:
        print(f"[错误] 验证版本号失败: {e}")
        return False


def main():
    script_dir = get_script_dir()
    readme_path = os.path.join(script_dir, 'README.md')
    python_path = os.path.join(script_dir, 'Excel批量解密工具与密码管理.py')
    
    print("=" * 50)
    print("版本同步工具")
    print("=" * 50)
    print()
    
    # 检查文件是否存在
    if not os.path.exists(readme_path):
        print(f"[错误] README.md文件不存在: {readme_path}")
        return 1
    
    if not os.path.exists(python_path):
        print(f"[错误] Python文件不存在: {python_path}")
        return 1
    
    # 读取版本号
    version = read_version_from_readme(readme_path)
    if not version:
        return 1
    
    print()
    
    # 更新版本号
    print("[步骤] 正在更新Python文件中的版本号...")
    updated = update_version_in_python(python_path, version)
    
    print()
    
    # 验证版本号
    print("[步骤] 正在验证版本号...")
    success = verify_version_in_python(python_path, version)
    
    print()
    print("=" * 50)
    if success:
        print("版本同步完成！")
    else:
        print("版本同步失败！")
    print("=" * 50)
    
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
