import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import threading
import time
import json
import os
import sys
import pyautogui
import pyperclip
from datetime import datetime
from PIL import Image, ImageDraw, ImageGrab

# Windows 디스플레이 확대/축소(DPI) 설정 시 캡쳐 영역 어긋남 방지
try:
    import ctypes

    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

# Windows 전용 트레이 라이브러리
try:
    import pystray
    from pystray import MenuItem as item
except ImportError:
    messagebox.showerror("라이브러리 누락", "pystray 라이브러리가 필요합니다.\n명령어: pip install pystray")
    sys.exit()

# 부팅 시 자동 실행을 위한 winreg
try:
    import winreg

    WINREG_AVAILABLE = True
except ImportError:
    WINREG_AVAILABLE = False

APP_NAME = "EMR_Sequencer"

# EXE 실행 환경 여부 확인 및 경로 설정
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(application_path, 'sequence_config.json')


# --- 화면 캡쳐를 위한 오버레이 클래스 ---
class SnippingTool:
    def __init__(self, root):
        self.root = root
        self.result_img = None
        self.snip_window = tk.Toplevel(root)

        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        self.snip_window.geometry(f"{screen_width}x{screen_height}+0+0")
        self.snip_window.overrideredirect(True)

        self.snip_window.attributes('-alpha', 0.3)
        self.snip_window.config(bg="black", cursor="cross")
        self.snip_window.attributes("-topmost", True)

        self.canvas = tk.Canvas(self.snip_window, cursor="cross", bg="black", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)

        self.start_x = None
        self.start_y = None
        self.rect = None

        self.canvas.bind("<ButtonPress-1>", self.on_press)
        self.canvas.bind("<B1-Motion>", self.on_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_release)
        self.snip_window.bind("<Escape>", lambda e: self.cancel())

    def on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline='red',
                                                 width=2, fill="black")

    def on_drag(self, event):
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_release(self, event):
        end_x, end_y = event.x, event.y
        self.snip_window.withdraw()
        self.snip_window.update_idletasks()

        x1 = min(self.start_x, end_x)
        y1 = min(self.start_y, end_y)
        x2 = max(self.start_x, end_x)
        y2 = max(self.start_y, end_y)

        if x2 - x1 > 5 and y2 - y1 > 5:
            self.root.after(100, self.capture_screen, x1, y1, x2, y2)
        else:
            self.snip_window.destroy()

    def capture_screen(self, x1, y1, x2, y2):
        self.result_img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        self.snip_window.destroy()

    def cancel(self):
        self.snip_window.destroy()


class EMRSequenceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EMR 자동화 시퀀서")
        self.root.geometry("650x720")
        self.root.resizable(False, False)

        self.is_running = False
        self.tray_icon = None

        self.processes = {}
        self.current_process = ""
        self.default_delay = 2.0
        self.schedules = {}
        self.tray_enabled = tk.BooleanVar(value=False)
        self.autostart_enabled = tk.BooleanVar(value=False)

        self.last_run_date = {}

        self.load_config()
        self.create_widgets()
        self.update_listbox()
        self.update_schedule_ui()
        self.sync_autostart_checkbox()

        threading.Thread(target=self.schedule_checker, daemon=True).start()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def create_widgets(self):
        """UI 구성"""
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10, fill=tk.X, padx=10)

        tk.Label(top_frame, text="프로세스:").grid(row=0, column=0, padx=2, pady=2, sticky="w")

        self.combo_process = ttk.Combobox(top_frame, values=list(self.processes.keys()), state="readonly", width=25)
        self.combo_process.grid(row=0, column=1, padx=2, pady=2)
        if self.current_process:
            self.combo_process.set(self.current_process)
        self.combo_process.bind("<<ComboboxSelected>>", self.on_process_change)

        tk.Button(top_frame, text="새로 만들기", command=self.add_process).grid(row=0, column=2, padx=2, pady=2)
        tk.Button(top_frame, text="이름 변경", command=self.rename_process).grid(row=0, column=3, padx=2, pady=2)
        tk.Button(top_frame, text="삭제", command=self.delete_process).grid(row=0, column=4, padx=2, pady=2)
        tk.Button(top_frame, text="설정 저장", command=self.manual_save, bg="#e6e6ff").grid(row=0, column=5, padx=2, pady=2)

        tk.Label(top_frame, text="예약 시간:").grid(row=1, column=0, padx=2, pady=5, sticky="w")

        # --- 시간 입력 위젯 (Spinbox) ---
        time_frame = tk.Frame(top_frame)
        time_frame.grid(row=1, column=1, padx=2, pady=5, sticky="w")

        self.hour_var = tk.StringVar(value="09")
        self.hour_spin = tk.Spinbox(time_frame, from_=0, to=23, textvariable=self.hour_var, width=3, format="%02.0f")
        self.hour_spin.pack(side=tk.LEFT)

        tk.Label(time_frame, text=":").pack(side=tk.LEFT, padx=2)

        self.minute_var = tk.StringVar(value="00")
        self.minute_spin = tk.Spinbox(time_frame, from_=0, to=59, textvariable=self.minute_var, width=3, format="%02.0f")
        self.minute_spin.pack(side=tk.LEFT)
        # --- 시간 입력 위젯 끝 ---

        tk.Button(top_frame, text="예약 저장", command=self.save_schedule, bg="#e6ffe6").grid(row=1, column=2, padx=2,
                                                                                          pady=5, sticky="we")
        tk.Button(top_frame, text="예약 취소", command=self.cancel_schedule, bg="#ffe6e6").grid(row=1, column=3, padx=2,
                                                                                            pady=5, sticky="we",
                                                                                            columnspan=2)

        self.status_label = tk.Label(self.root, text="대기 중...", font=("맑은 고딕", 11, "bold"), fg="blue")
        self.status_label.pack(pady=2)

        middle_frame = tk.Frame(self.root)
        middle_frame.pack(pady=5, fill=tk.BOTH, expand=True, padx=10)

        left_frame = tk.Frame(middle_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(left_frame, width=35, height=14)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(left_frame, orient="vertical")
        scrollbar.config(command=self.listbox.yview)
        scrollbar.pack(side=tk.LEFT, fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)

        order_frame = tk.Frame(left_frame)
        order_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        tk.Button(order_frame, text="▲ 위로", command=self.move_up).pack(pady=2, fill=tk.X)
        tk.Button(order_frame, text="▼ 아래로", command=self.move_down).pack(pady=2, fill=tk.X)

        right_frame = tk.Frame(middle_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5)

        tk.Button(right_frame, text="+ 이미지 클릭", command=self.add_click, width=17).pack(pady=3)
        tk.Button(right_frame, text="+ 클릭 & 텍스트", command=self.add_type, width=17).pack(pady=3)
        tk.Button(right_frame, text="+ 키 입력(엔터 등)", command=self.add_key, width=17).pack(pady=3)
        tk.Button(right_frame, text="+ 단순 대기(초)", command=self.add_wait, width=17).pack(pady=3)
        tk.Button(right_frame, text="+ 이미지 확인(대기)", command=self.add_wait_image, width=17, bg="#fffde6").pack(pady=3)

        tk.Button(right_frame, text="선택한 작업 이름 변경", fg="purple", command=self.rename_action, width=17).pack(
            pady=(15, 3))

        tk.Button(right_frame, text="선택한 작업(내용) 수정", fg="blue", command=self.edit_action, width=17).pack(pady=3)
        tk.Button(right_frame, text="선택한 작업 삭제", fg="red", command=self.delete_action, width=17).pack(pady=3)

        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(pady=10, fill=tk.X, padx=10)

        options_frame = tk.Frame(bottom_frame)
        options_frame.pack(side=tk.TOP, pady=5, fill=tk.X)

        tk.Button(options_frame, text="환경설정", command=self.open_settings_window).pack(side=tk.LEFT, padx=5)

        delay_frame = tk.Frame(bottom_frame)
        delay_frame.pack(side=tk.TOP, pady=5)
        tk.Label(delay_frame, text="동작 간 기본 대기시간(초):").pack(side=tk.LEFT)

        self.delay_var = tk.DoubleVar(value=self.default_delay)
        delay_spin = tk.Spinbox(delay_frame, from_=0.0, to=10.0, increment=0.5, textvariable=self.delay_var, width=5,
                                command=self.save_config)
        delay_spin.pack(side=tk.LEFT, padx=5)
        delay_spin.bind("<KeyRelease>", lambda e: self.save_config())

        ctrl_frame = tk.Frame(bottom_frame)
        ctrl_frame.pack(side=tk.TOP, pady=5)

        self.start_btn = tk.Button(ctrl_frame, text="▶ 시작", width=12, bg="#e6f2ff", command=self.start_rpa)
        self.start_btn.grid(row=0, column=0, padx=5)

        self.stop_btn = tk.Button(ctrl_frame, text="■ 정지", width=12, bg="#ffe6e6", state=tk.DISABLED,
                                  command=self.stop_rpa)
        self.stop_btn.grid(row=0, column=1, padx=5)

    def open_settings_window(self):
        settings_win = tk.Toplevel(self.root)
        settings_win.title("환경설정")
        settings_win.geometry("300x150")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()

        frame = tk.Frame(settings_win, padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        tray_cb = tk.Checkbutton(frame, text="종료 시 트레이로 최소화", variable=self.tray_enabled, command=self.save_config)
        tray_cb.pack(anchor="w", pady=5)

        autostart_cb = tk.Checkbutton(frame, text="윈도우 부팅 시 자동 실행", variable=self.autostart_enabled,
                                      command=self.toggle_autostart)
        autostart_cb.pack(anchor="w", pady=5)
        if not WINREG_AVAILABLE:
            autostart_cb.config(state=tk.DISABLED)

        close_btn = tk.Button(frame, text="닫기", command=settings_win.destroy)
        close_btn.pack(pady=10)

    # --- 프로세스 및 예약 관리 기능 ---
    def on_process_change(self, event=None):
        self.current_process = self.combo_process.get()
        self.update_listbox()
        self.update_schedule_ui()

    def update_schedule_ui(self):
        scheduled_time = self.schedules.get(self.current_process)
        if scheduled_time:
            try:
                hour, minute = scheduled_time.split(":")
                self.hour_var.set(f"{int(hour):02d}")
                self.minute_var.set(f"{int(minute):02d}")
                self.status_label.config(text=f"'{self.current_process}' 예약됨 ({scheduled_time})", fg="green")
            except (ValueError, TypeError):
                self.hour_var.set("09")
                self.minute_var.set("00")
                self.status_label.config(text="대기 중...", fg="blue")
        else:
            self.hour_var.set("09")
            self.minute_var.set("00")
            self.status_label.config(text="대기 중...", fg="blue")

    def save_schedule(self):
        try:
            hour = int(self.hour_var.get())
            minute = int(self.minute_var.get())
            time_str = f"{hour:02d}:{minute:02d}"

            self.schedules[self.current_process] = time_str
            self.save_config()
            self.update_schedule_ui()
            messagebox.showinfo("예약 완료", f"'{self.current_process}' 프로세스가 매일 {time_str}에 실행됩니다.")
        except ValueError:
            messagebox.showerror("오류", "시간 형식이 올바르지 않습니다. 숫자를 입력해주세요.")

    def cancel_schedule(self):
        if self.current_process in self.schedules:
            del self.schedules[self.current_process]
            self.save_config()
            self.update_schedule_ui()
            messagebox.showinfo("예약 취소", "예약이 취소되었습니다.")

    def add_process(self):
        new_name = simpledialog.askstring("새 프로세스", "새로운 프로세스의 이름을 입력하세요:")
        if new_name:
            if new_name in self.processes:
                messagebox.showwarning("경고", "이미 존재하는 프로세스 이름입니다.")
                return
            self.processes[new_name] = []
            self.current_process = new_name
            self.update_combo_box()
            self.update_listbox()
            self.update_schedule_ui()
            self.save_config()

    def rename_process(self):
        old_name = self.current_process
        new_name = simpledialog.askstring("이름 변경", "새로운 프로세스 이름을 입력하세요:", initialvalue=old_name)
        if new_name and new_name != old_name:
            if new_name in self.processes:
                messagebox.showwarning("경고", "이미 존재하는 프로세스 이름입니다.")
                return
            self.processes[new_name] = self.processes.pop(old_name)
            if old_name in self.schedules:
                self.schedules[new_name] = self.schedules.pop(old_name)
            self.current_process = new_name
            self.update_combo_box()
            self.save_config()
            messagebox.showinfo("완료", "프로세스 이름이 변경되었습니다.")

    def delete_process(self):
        if len(self.processes) <= 1:
            messagebox.showwarning("경고", "최소 1개의 프로세스는 유지해야 합니다.")
            return
        confirm = messagebox.askyesno("삭제 확인", f"'{self.current_process}' 프로세스를 삭제하시겠습니까?")
        if confirm:
            del self.processes[self.current_process]
            if self.current_process in self.schedules:
                del self.schedules[self.current_process]
            self.current_process = list(self.processes.keys())[0]
            self.update_combo_box()
            self.update_listbox()
            self.update_schedule_ui()
            self.save_config()

    def update_combo_box(self):
        self.combo_process['values'] = list(self.processes.keys())
        self.combo_process.set(self.current_process)

    def get_image_path(self, is_edit=False):
        if is_edit:
            msg = "기존 이미지를 변경하시겠습니까?\n\n[예] 새로 직접 화면 캡쳐\n[아니오] 새 파일 선택\n[취소] 기존 이미지 유지"
        else:
            msg = "이미지를 어떻게 지정하시겠습니까?\n\n[예] 화면 직접 캡쳐 (추천)\n[아니오] 기존 이미지 파일 선택\n[취소] 작업 취소"

        choice = messagebox.askyesnocancel("이미지 지정", msg)

        if choice is True:
            self.root.withdraw()
            self.root.update_idletasks()

            snipper = SnippingTool(self.root)
            self.root.wait_window(snipper.snip_window)

            self.root.deiconify()

            if snipper.result_img:
                save_dir = os.path.join(application_path, "captures")
                if not os.path.exists(save_dir):
                    os.makedirs(save_dir)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_path = os.path.join(save_dir, f"cap_{timestamp}.png")
                snipper.result_img.save(file_path)
                return file_path
            return None

        elif choice is False:
            return filedialog.askopenfilename(title="이미지 선택", filetypes=[("Image files", "*.png *.jpg")])

        else:
            return None

    # --- 액션 추가/삭제/수정/이름변경 기능 ---
    def add_click(self):
        file_path = self.get_image_path()
        if file_path:
            action = {"type": "click", "image": file_path, "alias": ""}
            self.processes[self.current_process].append(action)
            self.update_listbox()
            self.save_config()

    def add_type(self):
        file_path = self.get_image_path()
        if file_path:
            text = simpledialog.askstring("텍스트 입력", "입력할 텍스트를 적어주세요:")
            if text is not None:
                action = {"type": "type", "image": file_path, "text": text, "alias": ""}
                self.processes[self.current_process].append(action)
                self.update_listbox()
                self.save_config()

    def add_wait_image(self):
        file_path = self.get_image_path()
        if file_path:
            timeout = simpledialog.askfloat("타임아웃 설정", "이미지가 나타날 때까지 기다릴 최대 시간(초)을 입력하세요\n(예: 10초 이내에 안 나타나면 오류 처리):",
                                            initialvalue=10.0)
            if timeout is not None:
                action = {"type": "wait_image", "image": file_path, "timeout": timeout, "alias": ""}
                self.processes[self.current_process].append(action)
                self.update_listbox()
                self.save_config()

    def add_key(self):
        key = simpledialog.askstring("키보드 입력", "입력할 키를 적어주세요 (예: enter, tab, esc):")
        if key:
            action = {"type": "key", "key": key.lower(), "alias": ""}
            self.processes[self.current_process].append(action)
            self.update_listbox()
            self.save_config()

    def add_wait(self):
        sec = simpledialog.askfloat("대기 시간", "기다릴 시간(초)을 입력하세요 (예: 1.5):")
        if sec is not None:
            action = {"type": "wait", "time": sec, "alias": ""}
            self.processes[self.current_process].append(action)
            self.update_listbox()
            self.save_config()

    def rename_action(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showwarning("선택 오류", "이름을 변경할 작업을 선택해주세요.")
            return

        idx = selected[0]
        act = self.processes[self.current_process][idx]
        current_alias = act.get("alias", "")

        new_alias = simpledialog.askstring(
            "작업 이름 설정",
            "이 작업이 리스트에 표시될 이름을 입력하세요:\n(예: '로그인 버튼 클릭')\n\n※ 비워두면 기본 파일명/값으로 표시됩니다.",
            initialvalue=current_alias
        )

        if new_alias is not None:
            act["alias"] = new_alias.strip()
            self.processes[self.current_process][idx] = act
            self.update_listbox()
            self.save_config()

    def edit_action(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showwarning("선택 오류", "수정할 작업을 선택해주세요.")
            return

        idx = selected[0]
        act = self.processes[self.current_process][idx]

        if act["type"] == "click":
            file_path = self.get_image_path(is_edit=True)
            if file_path: act["image"] = file_path

        elif act["type"] == "type":
            file_path = self.get_image_path(is_edit=True)
            if file_path: act["image"] = file_path
            new_text = simpledialog.askstring("텍스트 수정", "새로운 텍스트를 입력하세요:", initialvalue=act.get("text", ""))
            if new_text is not None: act["text"] = new_text

        elif act["type"] == "wait_image":
            file_path = self.get_image_path(is_edit=True)
            if file_path: act["image"] = file_path
            new_timeout = simpledialog.askfloat("타임아웃 수정", "새로운 최대 대기 시간(초)을 입력하세요:",
                                                initialvalue=act.get("timeout", 10.0))
            if new_timeout is not None: act["timeout"] = new_timeout

        elif act["type"] == "key":
            new_key = simpledialog.askstring("키보드 입력 수정", "새로운 키를 입력하세요:", initialvalue=act.get("key", ""))
            if new_key: act["key"] = new_key.lower()

        elif act["type"] == "wait":
            new_time = simpledialog.askfloat("대기 시간 수정", "새로운 대기 시간(초)을 입력하세요:", initialvalue=act.get("time", 1.0))
            if new_time is not None: act["time"] = new_time

        self.processes[self.current_process][idx] = act
        self.update_listbox()
        self.save_config()

    def delete_action(self):
        selected = self.listbox.curselection()
        if selected:
            idx = selected[0]
            del self.processes[self.current_process][idx]
            self.update_listbox()
            self.save_config()

    # --- 스케줄러, UI 제어 및 기타 유틸 ---
    def move_up(self):
        selected = self.listbox.curselection()
        if not selected: return
        idx = selected[0]
        if idx == 0: return
        actions = self.processes[self.current_process]
        actions[idx], actions[idx - 1] = actions[idx - 1], actions[idx]
        self.update_listbox()
        self.listbox.selection_set(idx - 1)
        self.save_config()

    def move_down(self):
        selected = self.listbox.curselection()
        if not selected: return
        idx = selected[0]
        actions = self.processes[self.current_process]
        if idx == len(actions) - 1: return
        actions[idx], actions[idx + 1] = actions[idx + 1], actions[idx]
        self.update_listbox()
        self.listbox.selection_set(idx + 1)
        self.save_config()

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        current_actions = self.processes.get(self.current_process, [])
        for i, act in enumerate(current_actions):
            alias = act.get("alias", "")

            if act["type"] == "click":
                display = alias if alias else os.path.basename(act["image"])
                self.listbox.insert(tk.END, f"{i + 1}. [클릭] {display}")

            elif act["type"] == "type":
                display = alias if alias else f"{os.path.basename(act['image'])} -> '{act['text']}'"
                self.listbox.insert(tk.END, f"{i + 1}. [입력] {display}")

            elif act["type"] == "key":
                display = alias if alias else act['key']
                self.listbox.insert(tk.END, f"{i + 1}. [키보드] {display}")

            elif act["type"] == "wait":
                display = alias if alias else f"{act['time']}초"
                self.listbox.insert(tk.END, f"{i + 1}. [대기] {display}")

            elif act["type"] == "wait_image":
                display = alias if alias else f"{os.path.basename(act['image'])} (최대 {act['timeout']}초)"
                self.listbox.insert(tk.END, f"{i + 1}. [이미지 확인] {display}")

    def save_config(self):
        data = {
            "settings": {
                "default_delay": self.delay_var.get(),
                "tray_enabled": self.tray_enabled.get(),
                "autostart_enabled": self.autostart_enabled.get(),
            },
            "schedules": self.schedules,
            "processes": self.processes
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def manual_save(self):
        self.save_config()
        messagebox.showinfo("저장 완료", "모든 프로세스와 설정이 성공적으로 저장되었습니다.")

    def load_config(self):
        self.processes = {"기본 프로세스": []}
        self.default_delay = 2.0
        self.schedules = {}
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        self.processes = {"기본 프로세스": data}
                    elif isinstance(data, dict):
                        if "processes" in data:
                            self.processes = data["processes"]
                            settings = data.get("settings", {})
                            self.default_delay = settings.get("default_delay", 2.0)
                            self.tray_enabled.set(settings.get("tray_enabled", False))
                            # autostart_enabled는 sync_autostart_checkbox에서 설정하므로 여기서 로드하지 않음
                            self.schedules = data.get("schedules", {})
                        else:
                            self.processes = data
            except Exception as e:
                print(f"설정 파일 로드 오류: {e}")
        if self.processes:
            self.current_process = list(self.processes.keys())[0]

        for proc, actions in self.processes.items():
            for act in actions:
                if "alias" not in act:
                    act["alias"] = ""

    def schedule_checker(self):
        while True:
            now_time = datetime.now().strftime("%H:%M")
            now_date = datetime.now().strftime("%Y-%m-%d")
            for proc, scheduled_time in dict(self.schedules).items():
                run_key = f"{proc}_{now_date}"
                if now_time == scheduled_time and self.last_run_date.get(proc) != run_key and not self.is_running:
                    self.last_run_date[proc] = run_key
                    self.root.after(0, self.execute_scheduled_task, proc)
            time.sleep(10)

    def execute_scheduled_task(self, proc_name):
        self.current_process = proc_name
        self.update_combo_box()
        self.update_listbox()
        self.start_rpa()

    def on_closing(self):
        if self.tray_enabled.get():
            self.root.withdraw()
            self.show_tray_icon()
        else:
            self.root.quit()

    def create_tray_image(self):
        image = Image.new('RGB', (64, 64), color=(0, 122, 204))
        draw = ImageDraw.Draw(image)
        draw.rectangle([16, 16, 48, 48], fill="white")
        return image

    def show_tray_icon(self):
        menu = pystray.Menu(item('열기', self.restore_window), item('종료', self.quit_app))
        self.tray_icon = pystray.Icon("EMR_Sequencer", self.create_tray_image(), "EMR 예약 시퀀서", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def restore_window(self):
        if self.tray_icon: self.tray_icon.stop()
        self.root.after(0, self.root.deiconify)

    def quit_app(self):
        if self.tray_icon: self.tray_icon.stop()
        self.root.quit()

    def get_autostart_command(self):
        if getattr(sys, 'frozen', False):
            return f'"{sys.executable}"'
        else:
            script_path = os.path.abspath(__file__)
            pythonw = sys.executable.replace("python.exe", "pythonw.exe")
            if not os.path.exists(pythonw): pythonw = sys.executable
            return f'"{pythonw}" "{script_path}"'

    def is_autostart_registered(self):
        if not WINREG_AVAILABLE: return False
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0,
                                winreg.KEY_READ) as key:
                winreg.QueryValueEx(key, APP_NAME)
                return True
        except (FileNotFoundError, Exception):
            return False

    def sync_autostart_checkbox(self):
        if WINREG_AVAILABLE:
            is_registered = self.is_autostart_registered()
            self.autostart_enabled.set(is_registered)
            # config 파일의 상태와 레지스트리 상태를 동기화
            current_config_val = False
            if os.path.exists(CONFIG_FILE):
                try:
                    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        current_config_val = data.get("settings", {}).get("autostart_enabled", False)
                except Exception:
                    pass
            if is_registered != current_config_val:
                self.save_config()


    def toggle_autostart(self):
        if not WINREG_AVAILABLE:
            messagebox.showwarning("지원 안 됨", "이 기능은 Windows에서만 사용할 수 있습니다.")
            self.autostart_enabled.set(False)
            return
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0,
                                winreg.KEY_SET_VALUE) as key:
                if self.autostart_enabled.get():
                    winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, self.get_autostart_command())
                else:
                    try:
                        winreg.DeleteValue(key, APP_NAME)
                    except FileNotFoundError:
                        pass
            self.save_config()
        except Exception as e:
            messagebox.showerror("오류", f"자동 실행 설정 중 오류가 발생했습니다.\n\n{str(e)}")
            self.sync_autostart_checkbox()

    # --- RPA 실행 제어 ---
    def start_rpa(self):
        current_actions = self.processes.get(self.current_process, [])
        if not current_actions:
            messagebox.showwarning("경고", "실행할 작업이 없습니다.")
            return

        global_delay = self.delay_var.get()
        self.is_running = True
        self.status_label.config(text=f"'{self.current_process}' 진행 중...", fg="red")
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.combo_process.config(state=tk.DISABLED)
        self.hour_spin.config(state=tk.DISABLED)
        self.minute_spin.config(state=tk.DISABLED)

        threading.Thread(target=self.rpa_task, args=(current_actions, global_delay), daemon=True).start()

    def stop_rpa(self):
        self.is_running = False
        self.status_label.config(text="정지 중...", fg="orange")

    def execute_click(self, image_path):
        try:
            img = Image.open(image_path)
            location = pyautogui.locateCenterOnScreen(img, confidence=0.8)
            if location:
                pyautogui.click(location)
                time.sleep(0.5)
                return True
            return False
        except Exception as e:
            print(f"이미지 읽기 오류: {e}")
            return False

    def type_text_char_by_char(self, text, char_interval=0.05):
        original_clipboard = None
        try:
            original_clipboard = pyperclip.paste()
        except Exception:
            pass

        for ch in text:
            if not self.is_running: break
            pyperclip.copy(ch)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(char_interval)

        if original_clipboard is not None:
            try:
                pyperclip.copy(original_clipboard)
            except Exception:
                pass

    def rpa_task(self, actions, global_delay):
        try:
            for i, act in enumerate(actions):
                if not self.is_running: break

                self.listbox.selection_clear(0, tk.END)
                self.listbox.selection_set(i)
                self.listbox.see(i)

                if act["type"] == "click":
                    if not self.execute_click(act["image"]):
                        raise Exception(f"이미지를 찾을 수 없습니다: {os.path.basename(act['image'])}")

                elif act["type"] == "type":
                    if not self.execute_click(act["image"]):
                        raise Exception(f"이미지를 찾을 수 없습니다: {os.path.basename(act['image'])}")
                    time.sleep(0.5)
                    self.type_text_char_by_char(act["text"])
                    time.sleep(0.3)

                elif act["type"] == "key":
                    pyautogui.press(act["key"])
                    time.sleep(0.5)

                elif act["type"] == "wait":
                    wait_time = act["time"]
                    elapsed = 0
                    while elapsed < wait_time:
                        if not self.is_running: break
                        time.sleep(0.1)
                        elapsed += 0.1

                elif act["type"] == "wait_image":
                    timeout = act.get("timeout", 10.0)
                    img_path = act["image"]
                    try:
                        img = Image.open(img_path)
                    except Exception:
                        raise Exception(f"이미지를 열 수 없습니다: {os.path.basename(img_path)}")

                    start_time = time.time()
                    found = False

                    while time.time() - start_time < timeout:
                        if not self.is_running: break
                        try:
                            location = pyautogui.locateOnScreen(img, confidence=0.8)
                            if location:
                                found = True
                                break
                        except Exception:
                            pass
                        time.sleep(0.3)

                    if self.is_running and not found:
                        raise Exception(f"시간 초과: {timeout}초 내에 이미지를 찾을 수 없습니다\n({os.path.basename(img_path)})")

                if i < len(actions) - 1 and self.is_running:
                    delay_elapsed = 0
                    while delay_elapsed < global_delay:
                        if not self.is_running: break
                        time.sleep(0.1)
                        delay_elapsed += 0.1

            if self.is_running:
                self.status_label.config(text="작업 완료!", fg="green")
            else:
                self.status_label.config(text="사용자 중단", fg="orange")

        except Exception as e:
            self.status_label.config(text="오류 발생!", fg="red")
            messagebox.showerror("오류", f"작업 중 오류가 발생했습니다.\n\n{str(e)}")

        finally:
            self.reset_ui()

    def reset_ui(self):
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.combo_process.config(state="readonly")
        self.hour_spin.config(state=tk.NORMAL)
        self.minute_spin.config(state=tk.NORMAL)
        self.root.after(3000, self.update_schedule_ui)


if __name__ == "__main__":
    root = tk.Tk()
    app = EMRSequenceApp(root)
    root.mainloop()
