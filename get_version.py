"""build.bat 辅助脚本：输出 __version__ 的纯 ASCII 内容
设计要点：
1. 仅输出 ASCII 字符（避免 cmd 中文编码问题）
2. 通过 stdout 输出，结果用 exit code 0 表示成功
3. 用法：py -3.13 get_version.py <source_py_file>
"""
import re
import sys
import os

if len(sys.argv) < 2:
    print("ERR_NO_FILE", flush=True)
    sys.exit(1)

target = sys.argv[1]
if not os.path.exists(target):
    print("ERR_NOT_FOUND", flush=True)
    sys.exit(2)

try:
    with open(target, 'r', encoding='utf-8') as f:
        content = f.read()
    m = re.search(r'__version__\s*=\s*["\']([\d.]+)["\']', content)
    if m:
        # 关键：纯 ASCII 输出，无中文/特殊字符
        print(m.group(1), flush=True)
        sys.exit(0)
    else:
        print("ERR_NO_MATCH", flush=True)
        sys.exit(3)
except Exception as e:
    print(f"ERR:{e}", flush=True)
    sys.exit(4)
