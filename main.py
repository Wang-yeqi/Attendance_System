import cv2
import face_recognition
import numpy as np
import pickle
import os
import glob
import sqlite3
import csv
from datetime import datetime, date
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from PIL import Image, ImageTk
import threading
import time
import pyttsx3
import pandas as pd

# ==================== 配置参数 ====================
THRESHOLD = 0.5  # 识别阈值
DB_FILE = "attendance.db"  # SQLite数据库文件
STRANGER_LOG_FILE = "stranger_log.csv"  # 陌生人日志文件
STRANGER_ALERT_INTERVAL = 10  # 陌生人弹窗/语音最小间隔（秒）


# ==================== 数据库初始化 ====================
def init_db():
    """创建数据库和考勤表（如果不存在）"""
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS attendance
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT NOT NULL,
                  timestamp DATETIME NOT NULL,
                  date TEXT NOT NULL)''')
    conn.commit()
    conn.close()


def record_attendance_db(name):
    """记录考勤到数据库，如果今天已记录则返回False"""
    today = date.today().isoformat()
    now = datetime.now()
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id FROM attendance WHERE name=? AND date=?", (name, today))
    if c.fetchone():
        conn.close()
        return False
    c.execute("INSERT INTO attendance (name, timestamp, date) VALUES (?, ?, ?)",
              (name, now, today))
    conn.commit()
    conn.close()
    return True


# ==================== 人脸数据管理 ====================
def load_known_faces():
    encodings = []  # 存储人脸特征向量的列表
    names = []  # 存储对应姓名的列表
    for file in glob.glob("*.pkl"):  # 获取当前目录下所有扩展名为.pkl的文件
        try:
            with open(file, 'rb') as f:
                data = pickle.load(f)  # 通过Python的pickle模块将二进制文件还原为Python对象
            if 'encoding' in data and 'name' in data:  # 确保字典包含'encoding'（人脸特征向量）和'name'（用户标识）
                encodings.append(data['encoding'])  # 将分散在多个.pkl文件中的单个人脸数据（每个文件通常对应一个注册用户）整合为两个集中化列表
                names.append(data['name'])
        except Exception as e:
            print(f"加载文件 {file} 失败：{e}")
    return encodings, names


def save_single_face(encoding, name):
    # 保存单个人脸特征为 {name}.pkl
    filename = f"{name}.pkl"
    if os.path.exists(filename):
        overwrite = messagebox.askyesno("文件已存在", f"{filename} 已存在，是否覆盖？")
        if not overwrite:
            return False
    try:
        with open(filename, 'wb') as f:
            pickle.dump({'encoding': encoding, 'name': name}, f)
        return True
    except Exception as e:
        messagebox.showerror("保存失败", str(e))
        return False


def log_stranger():
    """记录陌生人出现时间到CSV文件"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_exists = os.path.isfile(STRANGER_LOG_FILE)
    with open(STRANGER_LOG_FILE, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)  # a表示追加写入（不会覆盖已有数据）
        if not file_exists:
            writer.writerow(['时间'])
        writer.writerow([now])


# ==================== 陌生人提示功能 ====================
class StrangerAlert:
    """陌生人提示管理类（控制弹窗和语音频率）"""

    def __init__(self):
        self.last_alert_time = 0  # 记录上次提示时间戳，实现频率控制
        self.engine = pyttsx3.init()  # 初始化语音引擎 （pyttsx3）
        self.engine.setProperty('rate', 150)  # 语速

    def should_alert(self):
        # 检查是否达到提示间隔用当前时间减去上次时间，大于设定间隔（如10分钟）则触发
        now = time.time()
        if now - self.last_alert_time > STRANGER_ALERT_INTERVAL:
            self.last_alert_time = now
            return True
        return False

    def voice_alert(self):
        """语音播报陌生人"""

        def speak():
            self.engine.say("检测到陌生人")
            self.engine.runAndWait()

        # 在新线程中运行语音，避免阻塞视频流
        threading.Thread(target=speak, daemon=True).start()

    def popup_alert(self):
        """弹窗提示陌生人（在主线程中执行）"""

        # 由于Tkinter是线程不安全的，需通过root.after在主线执行
        def show():
            messagebox.showwarning("陌生人警告", "检测到陌生人！")

        # 需要传入root对象，在调用时处理
        return show


# ==================== GUI 应用程序 ====================
class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("人脸识别考勤系统")
        self.root.geometry("800x600")

        # 初始化数据库
        init_db()

        # 加载已知人脸
        self.known_encodings, self.known_names = load_known_faces()
        print(f"已加载 {len(self.known_names)} 个人脸数据")

        # 创建陌生人提示管理器
        self.stranger_alert = StrangerAlert()

        # 创建GUI组件
        self.create_widgets()

        # 视频相关变量
        self.cap = None
        self.video_running = False
        self.current_frame = None

    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        ttk.Button(btn_frame, text="录入新用户", command=self.register_user).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="开始实时考勤", command=self.start_attendance).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="停止考勤", command=self.stop_attendance).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="查看考勤记录", command=self.show_records).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="退出", command=self.root.quit).pack(side=tk.LEFT, padx=5)

        # 视频显示区域
        self.video_label = ttk.Label(main_frame, relief=tk.SUNKEN, anchor=tk.CENTER)
        self.video_label.pack(fill=tk.BOTH, expand=True, pady=10)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, pady=5)

    def register_user(self):
        """录入新用户"""
        cap = cv2.VideoCapture(1)
        if not cap.isOpened():
            messagebox.showerror("错误", "无法打开摄像头")
            return

        cv2.namedWindow("录入 - 按s保存，按q取消", cv2.WINDOW_NORMAL)
        while True:
            ret, frame = cap.read()
            if not ret:
                break
            cv2.imshow("录入 - 按s保存，按q取消", frame)
            key = cv2.waitKey(1) & 0xFF
            if key == ord('s'):
                rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                locations = face_recognition.face_locations(rgb)
                if len(locations) != 1:
                    messagebox.showwarning("警告", f"检测到 {len(locations)} 个人脸，请确保只有一人")
                    continue
                encodings = face_recognition.face_encodings(rgb, locations)
                if not encodings:
                    messagebox.showwarning("警告", "无法提取人脸特征")
                    continue
                name = tk.simpledialog.askstring("输入姓名", "请输入姓名：")
                if not name:
                    continue
                if save_single_face(encodings[0], name):
                    messagebox.showinfo("成功", f"用户 {name} 录入成功")
                    self.known_encodings, self.known_names = load_known_faces()
                break
            elif key == ord('q'):
                break
        cap.release()
        cv2.destroyAllWindows()

    def start_attendance(self):
        """开始实时考勤"""
        if self.video_running:
            return
        self.cap = cv2.VideoCapture(1)
        if not self.cap.isOpened():
            messagebox.showerror("错误", "无法打开摄像头")
            return
        self.video_running = True
        self.update_video()
        self.status_var.set("考勤进行中...")

    def stop_attendance(self):
        """停止实时考勤"""
        self.video_running = False
        if self.cap:
            self.cap.release()
            self.cap = None
        self.video_label.config(image='')
        self.status_var.set("就绪")

    def update_video(self):
        """更新视频帧"""
        if not self.video_running:
            return
        ret, frame = self.cap.read()
        if ret:
            processed_frame = self.process_frame(frame)
            img = cv2.cvtColor(processed_frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(img)
            imgtk = ImageTk.PhotoImage(image=img)
            self.video_label.imgtk = imgtk
            self.video_label.config(image=imgtk)
        self.root.after(30, self.update_video)

    def process_frame(self, frame):
        """人脸识别处理，包含陌生人提示"""
        small = cv2.resize(frame, (0, 0), fx=0.5, fy=0.5)  # 缩小到原图1/2
        rgb_small = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)  # BGR转RGB

        locations = face_recognition.face_locations(rgb_small)
        # 示例输出：[(50,150,100,100)]  # 上、右、下、左坐标
        encodings = face_recognition.face_encodings(rgb_small, locations)
        # 128维人脸特征向量 每个脸都有独特的“数字指纹”（128个数字组成）
        for i, face_encoding in enumerate(encodings):
            if self.known_encodings:
                distances = face_recognition.face_distance(self.known_encodings, face_encoding)
                min_dist = np.min(distances)  # 最小距离
                best_idx = np.argmin(distances)  # 最匹配的索引 ：0=完全相同，1=完全不同（实际阈值通常设0.6）

                if min_dist < THRESHOLD:  ## 熟人处理
                    name = self.known_names[best_idx]
                    color = (0, 255, 0)  # 绿色框
                    if record_attendance_db(name):  # 记录考勤
                        print(f"考勤记录：{name} {datetime.now()}")
                else:  # 陌生人处理
                    name = "Unknown"
                    color = (0, 0, 255)
                    log_stranger()
                    # 陌生人提示（语音+弹窗）
                    if self.stranger_alert.should_alert():  # 检查提示间隔
                        self.stranger_alert.voice_alert()  # 语音提示
                        # 弹窗需在主线程执行
                        self.root.after(0, self.stranger_alert.popup_alert())  # 主线程弹窗
            else:
                name = "Unknown"
                color = (0, 0, 255)
                log_stranger()
                if self.stranger_alert.should_alert():
                    self.stranger_alert.voice_alert()
                    self.root.after(0, self.stranger_alert.popup_alert())

            top, right, bottom, left = locations[i]
            top *= 2
            right *= 2
            bottom *= 2
            left *= 2
            cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
            label = f"{name} ({min_dist:.2f})" if self.known_encodings and min_dist < THRESHOLD else name
            cv2.putText(frame, label, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.7, color, 2)

        return frame

    def show_records(self):
        """显示考勤记录查询窗口，增加导出Excel功能"""
        records_win = tk.Toplevel(self.root)
        records_win.title("考勤记录查询")
        records_win.geometry("650x450")

        # 查询框架
        query_frame = ttk.Frame(records_win, padding="5")
        query_frame.pack(fill=tk.X)

        ttk.Label(query_frame, text="选择日期：").pack(side=tk.LEFT)
        date_entry = ttk.Entry(query_frame, width=12)
        date_entry.pack(side=tk.LEFT, padx=5)
        date_entry.insert(0, date.today().isoformat())

        # 存储当前查询结果的变量
        current_rows = []

        def query():
            nonlocal current_rows
            selected_date = date_entry.get().strip()
            conn = sqlite3.connect(DB_FILE)
            c = conn.cursor()
            c.execute("SELECT name, timestamp FROM attendance WHERE date=? ORDER BY timestamp", (selected_date,))
            rows = c.fetchall()
            conn.close()
            current_rows = rows
            for row in tree.get_children():
                tree.delete(row)
            for name, ts in rows:
                tree.insert("", tk.END, values=(name, ts))

        ttk.Button(query_frame, text="查询", command=query).pack(side=tk.LEFT, padx=5)

        # 导出Excel按钮
        def export_excel():
            if not current_rows:
                messagebox.showwarning("无数据", "当前查询结果为空，无法导出")
                return
            # 选择保存路径
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                     filetypes=[("Excel文件", "*.xlsx")])
            if not file_path:
                return
            # 创建DataFrame并保存
            df = pd.DataFrame(current_rows, columns=["姓名", "时间"])
            try:
                df.to_excel(file_path, index=False, engine='openpyxl')
                messagebox.showinfo("导出成功", f"已保存至：{file_path}")
            except Exception as e:
                messagebox.showerror("导出失败", str(e))

        ttk.Button(query_frame, text="导出Excel", command=export_excel).pack(side=tk.LEFT, padx=5)

        # 表格显示
        columns = ("姓名", "时间")
        tree = ttk.Treeview(records_win, columns=columns, show="headings")
        tree.heading("姓名", text="姓名")
        tree.heading("时间", text="时间")
        tree.column("姓名", width=150)
        tree.column("时间", width=250)

        scrollbar = ttk.Scrollbar(records_win, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 初始加载今天记录
        query()


# ==================== 程序入口 ====================
if __name__ == "__main__":
    root = tk.Tk()
    app = AttendanceApp(root)
    root.mainloop()
