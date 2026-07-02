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
from PIL import Image, ImageDraw, ImageGrab, ImageTk

# --- 유틸리티 함수 ---
def center_window(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (width // 2)
    y = (win.winfo_screenheight() // 2) - (height // 2)
    win.geometry(f'{width}x{height}+{x}+{y}')

# Windows 디스플레이 확대/축소(DPI) 설정 시 캡쳐 영역 어긋남 방지
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
    GET_TICK_COUNT_AVAILABLE = True
except Exception:
    GET_TICK_COUNT_AVAILABLE = False

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

# 바로가기(.lnk) 해석 및 ShellExecute를 위한 pywin32
try:
    import win32com.client
    import win32api
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False


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


# --- 클릭 위치 선택을 위한 클래스 ---
class ClickPointSelector:
    def __init__(self, parent, image_path, current_click_type="single"):
        self.parent = parent
        self.image_path = image_path
        self.click_pos = None
        self.click_type = current_click_type

        self.win = tk.Toplevel(parent)
        self.win.title("클릭 위치 및 유형 선택")
        self.win.transient(parent)
        self.win.grab_set()

        try:
            self.pil_image = Image.open(image_path)
            self.tk_image = ImageTk.PhotoImage(self.pil_image)
        except Exception as e:
            messagebox.showerror("오류", f"이미지를 열 수 없습니다:\n{e}", parent=self.win)
            self.win.destroy()
            return

        max_w = self.win.winfo_screenwidth() * 0.8
        max_h = self.win.winfo_screenheight() * 0.8
        
        img_w, img_h = self.pil_image.size
        
        win_w = min(img_w + 40, max_w)
        win_h = min(img_h + 160, max_h)
        
        self.win.geometry(f"{int(win_w)}x{int(win_h)}")
        center_window(self.win)

        frame = tk.Frame(self.win)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        h_scroll = tk.Scrollbar(frame, orient="horizontal")
        v_scroll = tk.Scrollbar(frame, orient="vertical")

        self.canvas = tk.Canvas(frame, width=img_w, height=img_h, cursor="hand2",
                                xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        h_scroll.config(command=self.canvas.xview)
        v_scroll.config(command=self.canvas.yview)

        h_scroll.pack(side="bottom", fill="x")
        v_scroll.pack(side="right", fill="y")
        
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_image)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

        self.canvas.bind("<Button-1>", self.on_image_click)

        info_label = tk.Label(self.win, text="이미지에서 클릭할 지점을 선택하세요.", fg="blue")
        info_label.pack(pady=5)
        
        self.double_click_var = tk.BooleanVar(value=(current_click_type == "double"))
        double_click_cb = ttk.Checkbutton(self.win, text="더블 클릭", variable=self.double_click_var)
        double_click_cb.pack(pady=5)

        close_button = tk.Button(self.win, text="확인 (기본값: 중앙)", command=self.on_close, font=("맑은 고딕", 10))
        close_button.pack(pady=(5, 10), ipadx=10, ipady=4)

        self.win.protocol("WM_DELETE_WINDOW", self.on_close)
        self.parent.wait_window(self.win)

    def on_image_click(self, event):
        self.click_pos = (self.canvas.canvasx(event.x), self.canvas.canvasy(event.y))
        self.click_type = "double" if self.double_click_var.get() else "single"
        self.win.destroy()

    def on_close(self):
        if not self.click_pos:
            self.click_pos = (self.pil_image.width // 2, self.pil_image.height // 2)
        self.click_type = "double" if self.double_click_var.get() else "single"
        self.win.destroy()


class EMRSequenceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EMR 자동화 시퀀서")
        
        # 스타일 설정
        style = ttk.Style()
        style.configure("TButton", font=("맑은 고딕", 9))
        style.configure("TLabel", font=("맑은 고딕", 9))
        style.configure("TCheckbutton", font=("맑은 고딕", 9))
        style.configure("TRadiobutton", font=("맑은 고딕", 9))
        style.configure("TCombobox", font=("맑은 고딕", 9))
        style.configure("Treeview.Heading", font=("맑은 고딕", 10, "bold"))
        style.configure("Accent.TButton", font=("맑은 고딕", 9, "bold"))
        
        style.map('Treeview',
          foreground=self._fixed_map(style, 'foreground'),
          background=self._fixed_map(style, 'background'))


        self.is_running = False
        self.tray_icon = None

        self.processes = {}
        self.current_process = ""
        self.default_delay = 2.0
        self.schedules = {}
        self.tray_enabled = tk.BooleanVar(value=False)
        self.autostart_enabled = tk.BooleanVar(value=False)
        self.confidence_var = tk.DoubleVar(value=0.8)

        self.last_run_date = {}
        
        self.schedule_type_var = tk.StringVar(value="time")
        self.hour_var = tk.StringVar(value="09")
        self.minute_var = tk.StringVar(value="00")
        self.boot_minute_var = tk.StringVar(value="05")
        self.boot_second_var = tk.StringVar(value="00")

        self.load_config()
        
        if self.window_geometry:
            self.root.geometry(self.window_geometry)
        else:
            self.root.geometry("650x750")
            center_window(self.root)
        
        if self.window_state == 'zoomed':
            self.root.state('zoomed')

        self.root.resizable(True, True)
        self.root.minsize(650, 750)

        # 우클릭 메뉴 생성
        self.tree_menu = tk.Menu(self.root, tearoff=0)
        self.tree_menu.add_command(label="선택 항목 활성화", command=lambda: self.set_actions_enabled(True))
        self.tree_menu.add_command(label="선택 항목 비활성화", command=lambda: self.set_actions_enabled(False))

        self.create_widgets()
        self.update_treeview()
        self.update_schedule_ui()
        self.sync_autostart_checkbox()

        threading.Thread(target=self.schedule_checker, daemon=True).start()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _fixed_map(self, style, option):
        return [elm for elm in style.map('Treeview', query_opt=option) if
                elm[:2] != ('!disabled', '!selected')]

    def create_widgets(self):
        """UI 구성"""
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(fill=tk.X)
        top_frame.columnconfigure(6, weight=1)

        ttk.Label(top_frame, text="프로세스:").grid(row=0, column=0, padx=2, pady=2, sticky="w")

        self.combo_process = ttk.Combobox(top_frame, values=list(self.processes.keys()), state="readonly", width=25)
        self.combo_process.grid(row=0, column=1, padx=2, pady=2, columnspan=2)
        if self.current_process:
            self.combo_process.set(self.current_process)
        self.combo_process.bind("<<ComboboxSelected>>", self.on_process_change)

        ttk.Button(top_frame, text="새로 만들기", command=self.add_process).grid(row=0, column=3, padx=2, pady=2)
        ttk.Button(top_frame, text="이름 변경", command=self.rename_process).grid(row=0, column=4, padx=2, pady=2)
        ttk.Button(top_frame, text="삭제", command=self.delete_process).grid(row=0, column=5, padx=2, pady=2)
        ttk.Button(top_frame, text="환경설정", command=self.open_settings_window).grid(row=0, column=6, padx=(10, 2), pady=2, sticky="e")

        schedule_frame = ttk.LabelFrame(self.root, text="프로세스 실행 예약", padding=(10, 5))
        schedule_frame.pack(fill=tk.X, padx=10, pady=5)
        schedule_frame.columnconfigure(1, weight=1)

        self.time_radio = ttk.Radiobutton(schedule_frame, text="특정 시간 (HH:MM)", variable=self.schedule_type_var, value="time", command=self.toggle_schedule_widgets)
        self.time_radio.grid(row=0, column=0, sticky="w", padx=5)

        self.boot_radio = ttk.Radiobutton(schedule_frame, text="부팅 후 (MM:SS)", variable=self.schedule_type_var, value="boot", command=self.toggle_schedule_widgets)
        self.boot_radio.grid(row=1, column=0, sticky="w", padx=5)
        if not GET_TICK_COUNT_AVAILABLE:
            self.boot_radio.config(state=tk.DISABLED)

        self.time_input_frame = ttk.Frame(schedule_frame)
        self.time_input_frame.grid(row=0, column=1, sticky="w")
        self.hour_spin = tk.Spinbox(self.time_input_frame, from_=0, to=23, textvariable=self.hour_var, width=3, format="%02.0f")
        self.hour_spin.pack(side=tk.LEFT)
        ttk.Label(self.time_input_frame, text=":").pack(side=tk.LEFT, padx=2)
        self.minute_spin = tk.Spinbox(self.time_input_frame, from_=0, to=59, textvariable=self.minute_var, width=3, format="%02.0f")
        self.minute_spin.pack(side=tk.LEFT)

        self.boot_input_frame = ttk.Frame(schedule_frame)
        self.boot_input_frame.grid(row=1, column=1, sticky="w")
        self.boot_minute_spin = tk.Spinbox(self.boot_input_frame, from_=0, to=59, textvariable=self.boot_minute_var, width=3, format="%02.0f")
        self.boot_minute_spin.pack(side=tk.LEFT)
        ttk.Label(self.boot_input_frame, text=":").pack(side=tk.LEFT, padx=2)
        self.boot_second_spin = tk.Spinbox(self.boot_input_frame, from_=0, to=59, textvariable=self.boot_second_var, width=3, format="%02.0f")
        self.boot_second_spin.pack(side=tk.LEFT)

        btn_frame = ttk.Frame(schedule_frame)
        btn_frame.grid(row=0, column=2, rowspan=2, padx=(10,0))
        ttk.Button(btn_frame, text="예약 저장", command=self.save_schedule, style="Accent.TButton").pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="예약 취소", command=self.cancel_schedule).pack(fill=tk.X, pady=2)

        self.toggle_schedule_widgets()

        self.status_label = ttk.Label(self.root, text="대기 중...", font=("맑은 고딕", 11, "bold"), foreground="blue")
        self.status_label.pack(pady=2)

        middle_frame = ttk.Frame(self.root)
        middle_frame.pack(pady=5, fill=tk.BOTH, expand=True, padx=10)
        middle_frame.rowconfigure(1, weight=1)
        middle_frame.columnconfigure(0, weight=1)

        action_btn_frame = ttk.Frame(middle_frame)
        action_btn_frame.grid(row=0, column=0, sticky="w", pady=(0, 5))
        ttk.Button(action_btn_frame, text="선택 활성화", command=lambda: self.set_actions_enabled(True)).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(action_btn_frame, text="선택 비활성화", command=lambda: self.set_actions_enabled(False)).pack(side=tk.LEFT)

        tree_container = ttk.Frame(middle_frame)
        tree_container.grid(row=1, column=0, sticky="nsew")
        tree_container.rowconfigure(0, weight=1)
        tree_container.columnconfigure(0, weight=1)

        columns = ("#", "활성", "구분", "내용")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("#", text="번호", anchor="center")
        self.tree.heading("활성", text="활성", anchor="center")
        self.tree.heading("구분", text="구분", anchor="center")
        self.tree.heading("내용", text="내용")
        
        self.tree.column("#", width=40, anchor="center", stretch=False)
        self.tree.column("활성", width=50, anchor="center", stretch=False)
        self.tree.column("구분", width=120, anchor="center", stretch=False)
        self.tree.column("내용", width=300, stretch=True)
        
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        if sys.platform == "darwin":
            self.tree.bind("<Button-2>", self.show_tree_menu)
        else:
            self.tree.bind("<Button-3>", self.show_tree_menu)

        order_frame = ttk.Frame(middle_frame)
        order_frame.grid(row=1, column=1, sticky="ns", padx=5)
        ttk.Button(order_frame, text="▲ 위로", command=self.move_up).pack(pady=2, fill=tk.X)
        ttk.Button(order_frame, text="▼ 아래로", command=self.move_down).pack(pady=2, fill=tk.X)

        right_frame = ttk.Frame(middle_frame)
        right_frame.grid(row=1, column=2, sticky="ns", padx=5)

        ttk.Button(right_frame, text="+ 이미지 클릭", command=self.add_click).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 클릭 & 텍스트", command=self.add_type).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 키 입력(엔터 등)", command=self.add_key).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 단순 대기(초)", command=self.add_wait).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 이미지 확인(대기)", command=self.add_wait_image).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 파일 실행", command=self.add_exec_file).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 경로 열기", command=self.add_open_path).pack(pady=3, fill=tk.X)

        ttk.Button(right_frame, text="선택한 작업 이름 변경", command=self.rename_action).pack(
            pady=(15, 3), fill=tk.X)

        ttk.Button(right_frame, text="선택한 작업(내용) 수정", command=self.edit_action).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="선택한 작업 삭제", command=self.delete_action).pack(pady=3, fill=tk.X)

        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(pady=10, fill=tk.X, padx=10)

        delay_frame = ttk.Frame(bottom_frame)
        delay_frame.pack(side=tk.TOP, pady=5)
        ttk.Label(delay_frame, text="동작 간 기본 대기시간(초):").pack(side=tk.LEFT)

        self.delay_var = tk.DoubleVar(value=self.default_delay)
        delay_spin = tk.Spinbox(delay_frame, from_=0.0, to=10.0, increment=0.5, textvariable=self.delay_var, width=5,
                                command=self.save_config)
        delay_spin.pack(side=tk.LEFT, padx=5)
        delay_spin.bind("<KeyRelease>", lambda e: self.save_config())

        ctrl_frame = ttk.Frame(bottom_frame)
        ctrl_frame.pack(side=tk.TOP, pady=5)

        self.start_btn = ttk.Button(ctrl_frame, text="▶ 처음부터 시작", command=self.start_rpa)
        self.start_btn.grid(row=0, column=0, padx=5)

        self.start_from_btn = ttk.Button(ctrl_frame, text="▶ 선택부터 시작", command=self.start_from_selected)
        self.start_from_btn.grid(row=0, column=1, padx=5)

        self.stop_btn = ttk.Button(ctrl_frame, text="■ 정지", state=tk.DISABLED,
                                  command=self.stop_rpa)
        self.stop_btn.grid(row=0, column=2, padx=5)

    def show_tree_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            if item not in self.tree.selection():
                self.tree.selection_set(item)
                self.tree.focus(item)
            self.tree_menu.post(event.x_root, event.y_root)

    def set_actions_enabled(self, enabled):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("선택 오류", "상태를 변경할 작업을 하나 이상 선택해주세요.")
            return

        for item_id in selected_items:
            idx = self.tree.index(item_id)
            act = self.processes[self.current_process][idx]
            act["enabled"] = enabled
        
        self.save_config()
        self.update_treeview()

    def start_from_selected(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("선택 오류", "시작할 작업을 리스트에서 선택해주세요.")
            return
        
        idx = self.tree.index(selected_item[0])
        self.start_rpa(start_index=idx)

    def toggle_schedule_widgets(self):
        if self.schedule_type_var.get() == "time":
            for child in self.time_input_frame.winfo_children():
                child.configure(state='normal')
            for child in self.boot_input_frame.winfo_children():
                child.configure(state='disabled')
        else:
            for child in self.time_input_frame.winfo_children():
                child.configure(state='disabled')
            for child in self.boot_input_frame.winfo_children():
                child.configure(state='normal')

    def open_settings_window(self):
        settings_win = tk.Toplevel(self.root)
        settings_win.title("환경설정")
        settings_win.geometry("300x200")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()
        
        center_window(settings_win)

        frame = ttk.Frame(settings_win, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        tray_cb = ttk.Checkbutton(frame, text="종료 시 트레이로 최소화", variable=self.tray_enabled, command=self.save_config)
        tray_cb.pack(anchor="w", pady=5)

        autostart_cb = ttk.Checkbutton(frame, text="윈도우 부팅 시 자동 실행", variable=self.autostart_enabled,
                                      command=self.toggle_autostart)
        autostart_cb.pack(anchor="w", pady=5)
        if not WINREG_AVAILABLE:
            autostart_cb.config(state=tk.DISABLED)

        confidence_label = ttk.Label(frame, text=f"이미지 인식 정확도: {self.confidence_var.get():.2f}")
        confidence_label.pack(anchor="w", pady=(10, 0))
        confidence_scale = tk.Scale(frame, from_=0.1, to=1.0, resolution=0.05, orient=tk.HORIZONTAL, variable=self.confidence_var,
                                    command=lambda val: confidence_label.config(text=f"이미지 인식 정확도: {float(val):.2f}"))
        confidence_scale.pack(fill=tk.X)
        confidence_scale.bind("<ButtonRelease-1>", lambda e: self.save_config())


        close_btn = ttk.Button(frame, text="닫기", command=settings_win.destroy)
        close_btn.pack(pady=10)

    def on_process_change(self, event=None):
        self.current_process = self.combo_process.get()
        self.update_treeview()
        self.update_schedule_ui()

    def update_schedule_ui(self):
        schedule_info = self.schedules.get(self.current_process)
        if schedule_info:
            schedule_type = schedule_info.get("type", "time")
            schedule_value = schedule_info.get("value", "")
            self.schedule_type_var.set(schedule_type)

            if schedule_type == "time":
                try:
                    hour, minute = schedule_value.split(":")
                    self.hour_var.set(f"{int(hour):02d}")
                    self.minute_var.set(f"{int(minute):02d}")
                    self.status_label.config(text=f"'{self.current_process}' 예약됨 (매일 {schedule_value})", foreground="green")
                except (ValueError, TypeError):
                    self.status_label.config(text="대기 중...", foreground="blue")
            elif schedule_type == "boot":
                try:
                    minute, second = schedule_value.split(":")
                    self.boot_minute_var.set(f"{int(minute):02d}")
                    self.boot_second_var.set(f"{int(second):02d}")
                    self.status_label.config(text=f"'{self.current_process}' 예약됨 (부팅 후 {schedule_value})", foreground="purple")
                except (ValueError, TypeError):
                     self.status_label.config(text="대기 중...", foreground="blue")
        else:
            self.schedule_type_var.set("time")
            self.status_label.config(text="대기 중...", foreground="blue")
        
        self.toggle_schedule_widgets()

    def save_schedule(self):
        schedule_type = self.schedule_type_var.get()
        
        if schedule_type == "time":
            try:
                hour = int(self.hour_var.get())
                minute = int(self.minute_var.get())
                value = f"{hour:02d}:{minute:02d}"
                msg = f"프로세스가 매일 {value}에 실행됩니다."
            except ValueError:
                messagebox.showerror("오류", "시간 형식이 올바르지 않습니다. 숫자를 입력해주세요.")
                return
        elif schedule_type == "boot":
            try:
                minute = int(self.boot_minute_var.get())
                second = int(self.boot_second_var.get())
                value = f"{minute:02d}:{second:02d}"
                msg = f"프로세스가 부팅 후 {minute}분 {second}초 뒤에 실행됩니다."
            except ValueError:
                messagebox.showerror("오류", "시간 형식이 올바르지 않습니다. 숫자를 입력해주세요.")
                return
        else:
            return

        self.schedules[self.current_process] = {"type": schedule_type, "value": value}
        self.save_config()
        self.update_schedule_ui()
        messagebox.showinfo("예약 완료", f"'{self.current_process}' {msg}")

    def cancel_schedule(self):
        if self.current_process in self.schedules:
            del self.schedules[self.current_process]
            self.save_config()
            self.update_schedule_ui()
            messagebox.showinfo("예약 취소", "예약이 취소되었습니다.")

    def add_process(self):
        new_name = simpledialog.askstring("새 프로세스", "새로운 프로세스의 이름을 입력하세요:", parent=self.root)
        if new_name:
            if new_name in self.processes:
                messagebox.showwarning("경고", "이미 존재하는 프로세스 이름입니다.")
                return
            self.processes[new_name] = []
            self.current_process = new_name
            self.update_combo_box()
            self.update_treeview()
            self.update_schedule_ui()
            self.save_config()

    def rename_process(self):
        old_name = self.current_process
        new_name = simpledialog.askstring("이름 변경", "새로운 프로세스 이름을 입력하세요:", initialvalue=old_name, parent=self.root)
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
            self.update_treeview()
            self.update_schedule_ui()
            self.save_config()

    def update_combo_box(self):
        self.combo_process['values'] = list(self.processes.keys())
        self.combo_process.set(self.current_process)

    def get_image_and_click_pos(self, is_edit=False, current_action=None):
        file_path = None
        click_pos = None
        click_type = "single"

        if is_edit and current_action:
            click_type = current_action.get("click_type", "single")

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
                process_capture_dir = os.path.join(application_path, "captures", self.current_process)
                os.makedirs(process_capture_dir, exist_ok=True)
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_path = os.path.join(process_capture_dir, f"cap_{timestamp}.png")
                snipper.result_img.save(file_path)

        elif choice is False:
            file_path = filedialog.askopenfilename(title="이미지 선택", filetypes=[("Image files", "*.png *.jpg")])

        else: # Cancel
            if is_edit and current_action:
                return current_action.get("image"), current_action.get("click_pos"), current_action.get("click_type")
            return None, None, None

        if file_path:
            selector = ClickPointSelector(self.root, file_path, current_click_type=click_type)
            click_pos = selector.click_pos
            click_type = selector.click_type
            return file_path, click_pos, click_type

        return None, None, None
    
    def get_image_path_for_wait(self, is_edit=False, current_action=None):
        file_path = None
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
                process_capture_dir = os.path.join(application_path, "captures", self.current_process)
                os.makedirs(process_capture_dir, exist_ok=True)
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                file_path = os.path.join(process_capture_dir, f"cap_{timestamp}.png")
                snipper.result_img.save(file_path)
                return file_path

        elif choice is False:
            return filedialog.askopenfilename(title="이미지 선택", filetypes=[("Image files", "*.png *.jpg")])
        
        else:
            if is_edit and current_action:
                return current_action.get("image")
            return None
        return None

    def add_click(self):
        file_path, click_pos, click_type = self.get_image_and_click_pos()
        if file_path:
            action = {"type": "click", "image": file_path, "alias": "", "click_pos": click_pos, "click_type": click_type, "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def add_type(self):
        file_path, click_pos, click_type = self.get_image_and_click_pos()
        if file_path:
            text = simpledialog.askstring("텍스트 입력", "입력할 텍스트를 적어주세요:", parent=self.root)
            if text is not None:
                action = {"type": "type", "image": file_path, "text": text, "alias": "", "click_pos": click_pos, "click_type": click_type, "enabled": True}
                self.processes[self.current_process].append(action)
                self.update_treeview()
                self.save_config()

    def add_wait_image(self):
        file_path = self.get_image_path_for_wait()
        if file_path:
            timeout = simpledialog.askfloat("타임아웃 설정", "이미지가 나타날 때까지 기다릴 최대 시간(초)을 입력하세요\n(예: 10초 이내에 안 나타나면 오류 처리):",
                                            initialvalue=10.0, parent=self.root)
            if timeout is not None:
                action = {"type": "wait_image", "image": file_path, "timeout": timeout, "alias": "", "enabled": True}
                self.processes[self.current_process].append(action)
                self.update_treeview()
                self.save_config()

    def add_key(self):
        key = simpledialog.askstring("키보드 입력", "입력할 키를 적어주세요 (예: enter, tab, esc):", parent=self.root)
        if key:
            action = {"type": "key", "key": key.lower(), "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def add_wait(self):
        sec = simpledialog.askfloat("대기 시간", "기다릴 시간(초)을 입력하세요 (예: 1.5):", parent=self.root)
        if sec is not None:
            action = {"type": "wait", "time": sec, "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def add_exec_file(self):
        file_path = filedialog.askopenfilename(title="실행할 파일 선택")
        if file_path:
            action = {"type": "exec_file", "path": file_path, "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()
    
    def add_open_path(self):
        dir_path = filedialog.askdirectory(title="열고 싶은 폴더 선택")
        if dir_path:
            action = {"type": "open_path", "path": dir_path, "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def rename_action(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("선택 오류", "이름을 변경할 작업을 선택해주세요.")
            return

        idx = self.tree.index(selected_item[0])
        act = self.processes[self.current_process][idx]
        current_alias = act.get("alias", "")

        new_alias = simpledialog.askstring(
            "작업 이름 설정",
            "이 작업이 리스트에 표시될 이름을 입력하세요:\n(예: '로그인 버튼 클릭')\n\n※ 비워두면 기본 파일명/값으로 표시됩니다.",
            initialvalue=current_alias, parent=self.root
        )

        if new_alias is not None:
            act["alias"] = new_alias.strip()
            self.update_treeview()
            self.save_config()

    def edit_action(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("선택 오류", "수정할 작업을 선택해주세요.")
            return

        idx = self.tree.index(selected_item[0])
        act = self.processes[self.current_process][idx]

        if act["type"] in ["click", "type"]:
            file_path, click_pos, click_type = self.get_image_and_click_pos(is_edit=True, current_action=act)
            if file_path:
                act["image"] = file_path
                act["click_pos"] = click_pos
                act["click_type"] = click_type

            if act["type"] == "type":
                new_text = simpledialog.askstring("텍스트 수정", "새로운 텍스트를 입력하세요:", initialvalue=act.get("text", ""), parent=self.root)
                if new_text is not None:
                    act["text"] = new_text

        elif act["type"] == "wait_image":
            file_path = self.get_image_path_for_wait(is_edit=True, current_action=act)
            if file_path:
                act["image"] = file_path
            new_timeout = simpledialog.askfloat("타임아웃 수정", "새로운 최대 대기 시간(초)을 입력하세요:",
                                                initialvalue=act.get("timeout", 10.0), parent=self.root)
            if new_timeout is not None:
                act["timeout"] = new_timeout

        elif act["type"] == "key":
            new_key = simpledialog.askstring("키보드 입력 수정", "새로운 키를 입력하세요:", initialvalue=act.get("key", ""), parent=self.root)
            if new_key:
                act["key"] = new_key.lower()

        elif act["type"] == "wait":
            new_time = simpledialog.askfloat("대기 시간 수정", "새로운 대기 시간(초)을 입력하세요:", initialvalue=act.get("time", 1.0), parent=self.root)
            if new_time is not None:
                act["time"] = new_time
        
        elif act["type"] == "exec_file":
            new_path = filedialog.askopenfilename(title="실행할 파일 선택", initialfile=act.get("path"))
            if new_path:
                act["path"] = new_path
        
        elif act["type"] == "open_path":
            new_path = filedialog.askdirectory(title="열고 싶은 폴더 선택", initialdir=act.get("path"))
            if new_path:
                act["path"] = new_path

        self.update_treeview()
        self.save_config()

    def delete_action(self):
        selected_items = self.tree.selection()
        if selected_items:
            indices = sorted([self.tree.index(item) for item in selected_items], reverse=True)
            for idx in indices:
                del self.processes[self.current_process][idx]
            self.update_treeview()
            self.save_config()

    def move_up(self):
        selected_items = self.tree.selection()
        if not selected_items: return
        
        indices = sorted([self.tree.index(item) for item in selected_items])
        
        if indices[0] == 0: return

        for idx in indices:
            item = self.tree.get_children()[idx]
            self.tree.move(item, "", idx - 1)
            actions = self.processes[self.current_process]
            actions.insert(idx - 1, actions.pop(idx))
        
        self.save_config()
        self.update_treeview()
        
        new_selection_ids = [self.tree.get_children()[i-1] for i in indices]
        self.tree.selection_set(new_selection_ids)
        self.tree.focus(new_selection_ids[0])


    def move_down(self):
        selected_items = self.tree.selection()
        if not selected_items: return

        indices = sorted([self.tree.index(item) for item in selected_items], reverse=True)

        if indices[0] >= len(self.processes[self.current_process]) - 1: return

        for idx in indices:
            item = self.tree.get_children()[idx]
            self.tree.move(item, "", idx + 1)
            actions = self.processes[self.current_process]
            actions.insert(idx + 1, actions.pop(idx))
            
        self.save_config()
        self.update_treeview()

        new_selection_ids = [self.tree.get_children()[i+1] for i in indices]
        self.tree.selection_set(new_selection_ids)
        self.tree.focus(new_selection_ids[0])


    def update_treeview(self):
        selection = self.tree.selection()
        scroll_pos = self.tree.yview()

        for i in self.tree.get_children():
            self.tree.delete(i)
            
        current_actions = self.processes.get(self.current_process, [])
        for i, act in enumerate(current_actions):
            alias = act.get("alias", "")
            enabled = act.get("enabled", True)
            enabled_display = 'O' if enabled else 'X'
            
            act_type_display = ""
            content_display = ""

            if act["type"] == "click":
                act_type_display = "[더블클릭]" if act.get("click_type") == "double" else "[클릭]"
                content_display = alias if alias else os.path.basename(act["image"])
            elif act["type"] == "type":
                act_type_display = "[더블클릭 & 입력]" if act.get("click_type") == "double" else "[클릭 & 입력]"
                content_display = alias if alias else f"{os.path.basename(act['image'])} -> '{act['text']}'"
            elif act["type"] == "key":
                act_type_display = "[키보드]"
                content_display = alias if alias else act['key']
            elif act["type"] == "wait":
                act_type_display = "[대기]"
                content_display = alias if alias else f"{act['time']}초"
            elif act["type"] == "wait_image":
                act_type_display = "[이미지 확인]"
                content_display = alias if alias else f"{os.path.basename(act['image'])} (최대 {act['timeout']}초)"
            elif act["type"] == "exec_file":
                act_type_display = "[파일 실행]"
                content_display = alias if alias else os.path.basename(act["path"])
            elif act["type"] == "open_path":
                act_type_display = "[경로 열기]"
                content_display = alias if alias else os.path.basename(act["path"])

            tags = []
            if not enabled:
                tags.append('disabled')
            if act.get("click_pos"):
                tags.append('has_click_pos')

            self.tree.tag_configure('disabled', foreground='gray')
            self.tree.tag_configure('has_click_pos', foreground='purple')

            self.tree.insert("", "end", values=(i + 1, enabled_display, act_type_display, content_display), tags=tuple(tags))
        
        if selection:
            self.tree.selection_set(selection)
        self.tree.yview_moveto(scroll_pos[0])


    def save_config(self):
        geometry = ""
        if self.root.state() == 'normal':
            geometry = self.root.geometry()

        data = {
            "settings": {
                "default_delay": self.delay_var.get(),
                "tray_enabled": self.tray_enabled.get(),
                "autostart_enabled": self.autostart_enabled.get(),
                "window_geometry": geometry,
                "window_state": self.root.state(),
                "confidence": self.confidence_var.get()
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
        self.window_geometry = ""
        self.window_state = "normal"
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    self.processes = {"기본 프로세스": data}
                elif isinstance(data, dict):
                    self.processes = data.get("processes", {"기본 프로세스": []})
                    settings = data.get("settings", {})
                    self.default_delay = settings.get("default_delay", 2.0)
                    self.tray_enabled.set(settings.get("tray_enabled", False))
                    self.window_geometry = settings.get("window_geometry", "")
                    self.window_state = settings.get("window_state", "normal")
                    self.confidence_var.set(settings.get("confidence", 0.8))
                    
                    loaded_schedules = data.get("schedules", {})
                    for proc, val in loaded_schedules.items():
                        if isinstance(val, str):
                            self.schedules[proc] = {"type": "time", "value": val}
                        else:
                            self.schedules[proc] = val

            except Exception as e:
                print(f"설정 파일 로드 오류: {e}")

        if not self.processes:
            self.processes = {"기본 프로세스": []}
        self.current_process = list(self.processes.keys())[0]

        for proc, actions in self.processes.items():
            for act in actions:
                if "alias" not in act:
                    act["alias"] = ""
                if "click_pos" not in act:
                    act["click_pos"] = None
                if "click_type" not in act:
                    act["click_type"] = "single"
                if "enabled" not in act:
                    act["enabled"] = True

    def schedule_checker(self):
        while True:
            now_time_str = datetime.now().strftime("%H:%M")
            now_date_str = datetime.now().strftime("%Y-%m-%d")
            
            uptime_seconds = 0
            if GET_TICK_COUNT_AVAILABLE:
                uptime_seconds = ctypes.windll.kernel32.GetTickCount64() / 1000

            for proc, schedule_info in dict(self.schedules).items():
                run_key = f"{proc}_{now_date_str}"
                schedule_type = schedule_info.get("type")
                schedule_value = schedule_info.get("value")

                if self.last_run_date.get(proc) == run_key or self.is_running:
                    continue

                should_run = False
                if schedule_type == "time" and now_time_str == schedule_value:
                    should_run = True
                elif schedule_type == "boot" and uptime_seconds > 0:
                    try:
                        minutes, seconds = map(int, schedule_value.split(':'))
                        target_seconds = minutes * 60 + seconds
                        if abs(uptime_seconds - target_seconds) < 10:
                            should_run = True
                    except (ValueError, TypeError):
                        continue
                
                if should_run:
                    self.last_run_date[proc] = run_key
                    self.root.after(0, self.execute_scheduled_task, proc)

            time.sleep(10)

    def execute_scheduled_task(self, proc_name):
        self.current_process = proc_name
        self.update_combo_box()
        self.update_treeview()
        self.start_rpa()

    def on_closing(self):
        self.save_config()
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

    def start_rpa(self, start_index=0):
        current_actions = self.processes.get(self.current_process, [])
        if not current_actions:
            messagebox.showwarning("경고", "실행할 작업이 없습니다.")
            return

        global_delay = self.delay_var.get()
        self.is_running = True
        self.status_label.config(text=f"'{self.current_process}' 진행 중...", foreground="red")
        self.start_btn.config(state=tk.DISABLED)
        self.start_from_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.combo_process.config(state=tk.DISABLED)
        self.hour_spin.config(state=tk.DISABLED)
        self.minute_spin.config(state=tk.DISABLED)

        threading.Thread(target=self.rpa_task, args=(current_actions, global_delay, start_index), daemon=True).start()

    def stop_rpa(self):
        self.is_running = False
        self.status_label.config(text="정지 중...", foreground="orange")

    def execute_click(self, image_path, click_pos, click_type="single"):
        try:
            img = Image.open(image_path)
            location = pyautogui.locateOnScreen(img, confidence=self.confidence_var.get())
            if location:
                click_x = location.left + location.width / 2
                click_y = location.top + location.height / 2
                if click_pos:
                    click_x = location.left + click_pos[0]
                    click_y = location.top + click_pos[1]
                
                if click_type == "double":
                    pyautogui.doubleClick(click_x, click_y)
                else:
                    pyautogui.click(click_x, click_y)

                time.sleep(0.5)
                return True
            return False
        except Exception as e:
            print(f"이미지 읽기 또는 클릭 오류: {e}")
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

    def rpa_task(self, actions, global_delay, start_index=0):
        try:
            for i in range(start_index, len(actions)):
                act = actions[i]
                
                if not self.is_running: break

                self.root.after(0, lambda item_id=self.tree.get_children()[i]: (
                    self.tree.selection_set(item_id), self.tree.see(item_id)
                ))
                
                if not act.get("enabled", True):
                    continue

                if act["type"] == "click":
                    if not self.execute_click(act["image"], act.get("click_pos"), act.get("click_type", "single")):
                        raise Exception(f"이미지를 찾을 수 없습니다: {os.path.basename(act['image'])}")

                elif act["type"] == "type":
                    if not self.execute_click(act["image"], act.get("click_pos"), act.get("click_type", "single")):
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
                            location = pyautogui.locateOnScreen(img, confidence=self.confidence_var.get())
                            if location:
                                found = True
                                break
                        except Exception:
                            pass
                        time.sleep(0.3)

                    if self.is_running and not found:
                        raise Exception(f"시간 초과: {timeout}초 내에 이미지를 찾을 수 없습니다\n({os.path.basename(img_path)})")
                
                elif act["type"] == "exec_file":
                    try:
                        os.startfile(act["path"])
                    except Exception as e:
                        raise Exception(f"파일을 실행할 수 없습니다: {os.path.basename(act['path'])}\n{e}")
                
                elif act["type"] == "open_path":
                    try:
                        os.startfile(act["path"])
                    except Exception as e:
                        raise Exception(f"경로를 열 수 없습니다: {os.path.basename(act['path'])}\n{e}")


                if i < len(actions) - 1 and self.is_running:
                    delay_elapsed = 0
                    while delay_elapsed < global_delay:
                        if not self.is_running: break
                        time.sleep(0.1)
                        delay_elapsed += 0.1

            if self.is_running:
                self.root.after(0, lambda: self.status_label.config(text="작업 완료!", foreground="green"))
            else:
                self.root.after(0, lambda: self.status_label.config(text="사용자 중단", foreground="orange"))

        except Exception as e:
            self.root.after(0, lambda: self.status_label.config(text="오류 발생!", foreground="red"))
            self.root.after(0, lambda: messagebox.showerror("오류", f"작업 중 오류가 발생했습니다.\n\n{str(e)}"))

        finally:
            self.root.after(0, self.reset_ui)

    def reset_ui(self):
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.start_from_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.combo_process.config(state="readonly")
        self.toggle_schedule_widgets()
        self.root.after(3000, self.update_schedule_ui)


if __name__ == "__main__":
    root = tk.Tk()
    app = EMRSequenceApp(root)
    root.mainloop()
