# 人脸识别考勤系统 (Face Recognition Attendance System)

## 📌 简介
一句话说明项目用途（如：一款基于Python的自动化考勤工具，支持实时识别、陌生人告警、数据导出）。

## ✨ 功能特性
- ✅ 人脸录入：拍照提取特征，以姓名保存独立文件
- ✅ 实时考勤：摄像头识别，自动记录到SQLite（每日单次）
- ✅ 陌生人警示：红框标注 + 语音播报 + 弹窗提醒（间隔可调）
- ✅ 记录管理：按日期查询、导出Excel
- ✅ 图形界面：Tkinter实现，操作直观

## 🛠️ 技术栈
- Python 3.8+
- OpenCV / face_recognition / dlib
- Tkinter / Pillow
- SQLite / pandas / openpyxl
- pyttsx3 (语音)

## 🚀 快速开始
### 安装依赖
```bash
pip install -r requirements.txt
# 注意内容
## face_recognition，它会自动依赖 dlib 和 face_recognition_models，但 dlib 在 Windows 上可能需要额外处理
可自行搜索解决方法或 使用如下指令
conda install -c conda-forge dlib face_recognition
