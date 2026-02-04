#!/usr/bin/env python3
"""
PyInstaller 打包脚本 - 生成单文件 exe
"""
import PyInstaller.__main__
import os
import sys

# 获取项目根目录
project_dir = os.path.dirname(os.path.abspath(__file__))

# PyInstaller 参数
args = [
    os.path.join(project_dir, "swfit_app.py"),
    "--onefile",
    "--windowed",
    "--name=SWIFT_Data_Collection",
    f"--distpath={os.path.join(project_dir, 'dist')}",
    f"--buildpath={os.path.join(project_dir, 'build')}",
    f"--specpath={os.path.join(project_dir, 'spec')}",
    "--hidden-import=PySide6",
    "--hidden-import=pandas",
    "--hidden-import=openpyxl",
    "--hidden-import=extract_msg",
    "--collect-all=PySide6",
]

# 如果有 icon，添加 icon 参数
icon_path = os.path.join(project_dir, "app.ico")
if os.path.exists(icon_path):
    args.append(f"--icon={icon_path}")

print("开始打包...")
print(f"项目目录: {project_dir}")
print(f"参数: {args}\n")

PyInstaller.__main__.run(args)

print("\n✅ 打包完成！")
print(f"输出文件: {os.path.join(project_dir, 'dist', 'SWIFT_Data_Collection.exe')}")
