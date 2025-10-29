# MouseScrollStabilizer
This project aims to address the issue of frequent jittering back and forth when scrolling up and down with a mouse, caused by a damaged or worn-out scroll wheel.

## How It Works
1. Time Threshold: If a reverse scroll event occurs within the threshold time after the last scroll event, it will be blocked.
2. Count Threshold: When consecutive reverse scroll events (each within the time threshold) reach the count threshold, the initial direction changes and the event is allowed.
3. This effectively prevents accidental scrolling caused by mouse wheel jitter

## Usage Instructions
1. Install the Python runtime environment. Download link: https://www.python.org/downloads/windows/
2. Download the MouseScrollStabilizer.py file.
3. Right-click the file, select "Open with," and set the default program to pythonw.
4. Double-click the file to run it.
5. Alternatively, you can run it via Command Prompt or PowerShell.

# MouseScrollStabilizer
此项目是一个鼠标滚轮防抖工具，用于解决鼠标滚轮上下滚动时经常出现来回跳动的问题。

## 工作原理
1. 时间阈值：在距离上次滚轮事件的时间阈值内，如果出现反方向滚轮事件，则阻塞该事件。
2. 次数阈值：连续的反方向滚轮事件（每个事件距离上次滚轮时间在时间阈值内）达到次数阈值，则改变初始方向并允许此事件。
3. 这样可以有效防止因鼠标滚轮抖动导致的意外滚动。

## 使用说明
1. 安装python运行环境，下载地址: https://www.python.org/downloads/windows/
2. 下载MouseScrollStabilizer.py文件
3. 右键选择打开方式，将打开方式设置为pythonw
4. 双击打开即可。
5. 也可以用cmd或者powershell运行。
