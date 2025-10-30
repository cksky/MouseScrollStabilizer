# build.py
import os
import sys
import shutil
from PyInstaller.__main__ import run

def build_exe():
    # 检查图标文件是否存在
    if not os.path.exists('icon.ico'):
        print("警告: 未找到 icon.ico 文件，将使用默认图标")
        icon_option = []
    else:
        icon_option = ['--icon=icon.ico']

    # 清理之前的构建文件
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    if os.path.exists('MouseScrollStabilizer.spec'):
        os.remove('MouseScrollStabilizer.spec')

    # PyInstaller 配置
    opts = [
        '111.py',  # 主程序文件
        '--name=MouseScrollStabilizer',  # 输出名称
        '--onefile',  # 打包成单个exe文件
        '--windowed',  # 不显示控制台窗口
        '--add-data=config;config',  # 包含配置目录（如果存在）
        '--hidden-import=win32timezone',  # 隐藏导入
        '--hidden-import=win32gui',
        '--hidden-import=win32con', 
        '--hidden-import=win32process',
        '--hidden-import=psutil',
        '--uac-admin',  # 请求管理员权限
        '--clean',  # 清理临时文件
    ] + icon_option  # 添加图标选项

    # 运行打包
    run(opts)

    print("打包完成！输出文件在 dist 目录中")
    print("注意: 首次运行会在exe同目录创建config文件夹")

if __name__ == '__main__':
    build_exe()
