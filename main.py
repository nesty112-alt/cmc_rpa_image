import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
import threading
import time
import json
import os
import sys
import subprocess
import pyautogui
import pyperclip
import base64
from datetime import datetime, timedelta
from PIL import Image, ImageDraw, ImageGrab, ImageTk
import copy

# --- COM 스레딩 모델 설정 (Tkinter 폴더 대화상자와 pywinauto 충돌로 인한 응답없음 방지) ---
sys.coinit_flags = 2  # COINIT_APARTMENTTHREADED (STA 모드 강제 지정)

# --- UI 객체 인식(pywinauto) 라이브러리 로드 ---
try:
    from pywinauto import Desktop

    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False


# --- 유틸리티 함수 ---
def center_window(win):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()

    x, y = 0, 0
    try:
        # PYWIN32를 사용하여 주 모니터의 중앙에 위치시키기
        import win32api
        import win32con

        # 기본 모니터 핸들 가져오기
        monitor_handle = win32api.MonitorFromPoint((0, 0), win32con.MONITOR_DEFAULTTOPRIMARY)
        monitor_info = win32api.GetMonitorInfo(monitor_handle)
        monitor_area = monitor_info['Monitor']
        work_area = monitor_info['Work']

        # 모니터의 작업 영역 중앙 계산
        mon_x = work_area[0]
        mon_y = work_area[1]
        mon_width = work_area[2] - mon_x
        mon_height = work_area[3] - mon_y

        x = mon_x + (mon_width - width) // 2
        y = mon_y + (mon_height - height) // 2

    except (ImportError, AttributeError):
        # PYWIN32가 없거나 관련 함수 실패 시 기존 방식으로 대체
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

# tkcalendar 라이브러리 (날짜 선택기용)
try:
    from tkcalendar import Calendar

    TKCALENDAR_AVAILABLE = True
except ImportError:
    TKCALENDAR_AVAILABLE = False

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

        self.is_multi_monitor = False
        self.v_left, self.v_top = 0, 0

        # For multi-monitor support on Windows
        if GET_TICK_COUNT_AVAILABLE:  # This implies ctypes is available
            try:
                SM_XVIRTUALSCREEN = 76
                SM_YVIRTUALSCREEN = 77
                SM_CXVIRTUALSCREEN = 78
                SM_CYVIRTUALSCREEN = 79

                self.v_left = ctypes.windll.user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
                self.v_top = ctypes.windll.user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
                v_width = ctypes.windll.user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
                v_height = ctypes.windll.user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)

                self.snip_window.geometry(f"{v_width}x{v_height}+{self.v_left}+{self.v_top}")
                self.is_multi_monitor = True
            except Exception:
                self.is_multi_monitor = False

        if not self.is_multi_monitor:
            # Fallback for non-Windows or if ctypes fails
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
        try:
            if self.is_multi_monitor:
                # Adjust coordinates for the virtual screen
                abs_x1 = self.v_left + x1
                abs_y1 = self.v_top + y1
                abs_x2 = self.v_left + x2
                abs_y2 = self.v_top + y2
                self.result_img = ImageGrab.grab(bbox=(abs_x1, abs_y1, abs_x2, abs_y2), all_screens=True)
            else:
                # Original behavior
                self.result_img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        except Exception as e:
            messagebox.showerror("캡쳐 오류", f"화면을 캡쳐하는 중 오류가 발생했습니다:\n{e}")
            self.result_img = None

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


# --- UI 객체 속성 입력을 위한 클래스 (pywinauto 용) ---
class UIObjectSettingDialog:
    def __init__(self, parent, initial_data=None):
        self.parent = parent
        self.result = None
        if initial_data is None:
            initial_data = {}

        self.win = tk.Toplevel(parent)
        self.win.title("UI 객체 속성 입력")
        self.win.geometry("450x420")
        self.win.transient(parent)
        self.win.grab_set()
        center_window(self.win)

        main_frame = ttk.Frame(self.win, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        info_text = ("Accessibility Insights 등의 도구를 통해 알아낸\n"
                     "객체의 고유 속성을 입력하세요. (최소 1개 이상 입력)")
        ttk.Label(main_frame, text=info_text, justify=tk.CENTER, foreground="blue").pack(pady=(0, 15))

        # Auto ID
        frame_id = ttk.Frame(main_frame)
        frame_id.pack(fill="x", pady=5)
        ttk.Label(frame_id, text="Auto ID (AutomationId):", width=20).pack(side=tk.LEFT)
        self.auto_id_var = tk.StringVar(value=initial_data.get("auto_id", ""))
        ttk.Entry(frame_id, textvariable=self.auto_id_var).pack(side=tk.LEFT, fill="x", expand=True)

        # Title / Name
        frame_title = ttk.Frame(main_frame)
        frame_title.pack(fill="x", pady=5)
        ttk.Label(frame_title, text="Title / Name:", width=20).pack(side=tk.LEFT)
        self.title_var = tk.StringVar(value=initial_data.get("title", ""))
        ttk.Entry(frame_title, textvariable=self.title_var).pack(side=tk.LEFT, fill="x", expand=True)

        # Control Type
        frame_type = ttk.Frame(main_frame)
        frame_type.pack(fill="x", pady=5)
        ttk.Label(frame_type, text="Control Type (선택):", width=20).pack(side=tk.LEFT)
        self.control_type_var = tk.StringVar(value=initial_data.get("control_type", ""))
        ttk.Entry(frame_type, textvariable=self.control_type_var).pack(side=tk.LEFT, fill="x", expand=True)
        ttk.Label(main_frame, text="예: Button, Edit, Pane, Window 등", font=("맑은 고딕", 8), foreground="gray").pack(
            anchor="w", padx=140)

        self.double_click_var = tk.BooleanVar(value=initial_data.get("double_click", False))
        ttk.Checkbutton(main_frame, text="더블 클릭", variable=self.double_click_var).pack(pady=(15, 0), anchor="w")

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=(30, 0))

        ttk.Button(btn_frame, text="확인", command=self.on_ok, style="Accent.TButton", width=15).pack(side=tk.LEFT,
                                                                                                    padx=10)
        ttk.Button(btn_frame, text="취소", command=self.win.destroy, width=15).pack(side=tk.LEFT, padx=10)

    def on_ok(self):
        auto_id = self.auto_id_var.get().strip()
        title = self.title_var.get().strip()
        control_type = self.control_type_var.get().strip()

        if not auto_id and not title and not control_type:
            messagebox.showerror("오류", "최소 하나 이상의 속성을 입력해야 객체를 찾을 수 있습니다.", parent=self.win)
            return

        self.result = {
            "auto_id": auto_id,
            "title": title,
            "control_type": control_type,
            "double_click": self.double_click_var.get()
        }
        self.win.destroy()


# --- 키보드 입력을 순차적으로 기록하기 위한 클래스 ---
class KeyRecorder:
    def __init__(self, parent):
        self.parent = parent
        self.recorded_actions = []
        self.is_recording = False
        self.current_modifiers = set()

        self.win = tk.Toplevel(parent)
        self.win.title("키보드 입력 레코더")
        self.win.geometry("500x450")
        self.win.transient(parent)
        self.win.grab_set()
        center_window(self.win)

        lbl = ttk.Label(self.win, text="[기록 시작] 버튼을 누른 후 키보드를 입력하세요.\n조합키(Ctrl+C 등)도 지원됩니다.",
                        justify=tk.CENTER, font=("맑은 고딕", 10))
        lbl.pack(pady=15)

        self.display_text = tk.Text(self.win, height=10, width=55, state='disabled', font=("Consolas", 11))
        self.display_text.pack(pady=10, padx=20)

        btn_frame = ttk.Frame(self.win)
        btn_frame.pack(pady=10)

        self.start_btn = ttk.Button(btn_frame, text="기록 시작", command=self.toggle_recording, width=15)
        self.start_btn.pack(side=tk.LEFT, padx=10)

        self.clear_btn = ttk.Button(btn_frame, text="초기화", command=self.clear_record, width=10)
        self.clear_btn.pack(side=tk.LEFT, padx=10)

        self.save_btn = ttk.Button(self.win, text="작업 리스트에 추가 (종료)", command=self.save_and_close,
                                   style="Accent.TButton")
        self.save_btn.pack(pady=20, ipadx=20, ipady=5)

        self.win.bind("<KeyPress>", self.on_key_press)
        self.win.bind("<KeyRelease>", self.on_key_release)

    def toggle_recording(self):
        if not self.is_recording:
            self.is_recording = True
            self.current_modifiers.clear()
            self.start_btn.config(text="기록 중지 (정지)")
            self.win.focus_set()
        else:
            self.is_recording = False
            self.current_modifiers.clear()
            self.start_btn.config(text="기록 시작")

    def clear_record(self):
        self.recorded_actions = []
        self.current_modifiers.clear()
        self.update_display()

    def _get_pyautogui_key(self, event):
        key_name = event.keysym.lower()

        # PyAutoGUI 호환 키 매핑
        mapping = {
            "return": "enter",
            "escape": "esc",
            "backspace": "backspace",
            "tab": "tab",
            "space": "space",
            "up": "up",
            "down": "down",
            "left": "left",
            "right": "right",
            "prior": "pgup",
            "next": "pgdn",
            "end": "end",
            "home": "home",
            "delete": "delete",
            "caps_lock": "capslock",
            "shift_l": "shift",
            "shift_r": "shift",
            "control_l": "ctrl",
            "control_r": "ctrl",
            "alt_l": "alt",
            "alt_r": "alt"
        }

        if key_name in mapping:
            return mapping[key_name]
        elif len(event.char) == 1 and event.char.isprintable():
            return event.char.lower()
        else:
            return key_name

    def on_key_press(self, event):
        if not self.is_recording:
            return

        key = self._get_pyautogui_key(event)

        # 모디파이어 키 처리
        if key in ['shift', 'ctrl', 'alt', 'win']:
            self.current_modifiers.add(key)
            return "break"

        # 조합키 생성 (ex: ctrl+c)
        if self.current_modifiers:
            # 모디파이어 키들을 정렬해서 일관성 유지 (ctrl, alt, shift 순)
            mods = sorted(list(self.current_modifiers), key=lambda x: ['ctrl', 'alt', 'shift', 'win'].index(x) if x in ['ctrl', 'alt', 'shift', 'win'] else 99)
            combo = "+".join(mods + [key])
            self.recorded_actions.append(combo)
        else:
            self.recorded_actions.append(key)

        self.update_display()
        return "break"

    def on_key_release(self, event):
        if not self.is_recording:
            return
            
        key = self._get_pyautogui_key(event)
        if key in self.current_modifiers:
            self.current_modifiers.remove(key)
            
        return "break"

    def update_display(self):
        self.display_text.config(state='normal')
        self.display_text.delete("1.0", tk.END)
        self.display_text.insert("1.0", " -> ".join(self.recorded_actions))
        self.display_text.config(state='disabled')
        self.display_text.see(tk.END)

    def save_and_close(self):
        self.is_recording = False
        self.win.destroy()


# --- 날짜 입력 설정을 위한 클래스 ---
class DateSettingDialog:
    def __init__(self, parent, initial_offset=0, initial_format="%Y%m%d"):
        self.parent = parent
        self.result = None

        self.win = tk.Toplevel(parent)
        self.win.title("날짜 입력 설정")
        self.win.geometry("400x520")
        self.win.transient(parent)
        self.win.grab_set()
        center_window(self.win)

        today = datetime.now().date()
        target_date = today + timedelta(days=initial_offset)

        main_frame = ttk.Frame(self.win, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="1. 날짜 선택 (달력)", font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(0, 5))

        if TKCALENDAR_AVAILABLE:
            self.cal = Calendar(main_frame, selectmode='day',
                                year=target_date.year, month=target_date.month, day=target_date.day,
                                date_pattern='yyyy-mm-dd')
            self.cal.pack(fill="x", pady=5)
            # 초기 선택 설정
            self.cal.selection_set(target_date)
            self.cal.bind("<<CalendarSelected>>", self.on_date_selected)
        else:
            err_lbl = ttk.Label(main_frame,
                                text="tkcalendar 라이브러리가 설치되어 있지 않습니다.\n달력 기능을 사용하려면 아래 명령어를 실행하세요:\npip install tkcalendar",
                                foreground="red", justify=tk.CENTER)
            err_lbl.pack(pady=20)
            self.cal = None

        ttk.Label(main_frame, text="2. 기준일(오늘)로부터 차이 (일)", font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(15, 5))

        offset_frame = ttk.Frame(main_frame)
        offset_frame.pack(fill="x")

        self.offset_var = tk.IntVar(value=initial_offset)
        self.offset_spin = tk.Spinbox(offset_frame, from_=-3650, to=3650, textvariable=self.offset_var, width=10)
        self.offset_spin.pack(side=tk.LEFT, padx=5)

        # trace를 사용하여 값이 변경될 때마다 달력과 라벨 업데이트
        self.offset_var.trace_add("write", lambda *args: self.update_from_offset())

        self.calc_label = ttk.Label(offset_frame, text=f"(계산된 날짜: {target_date.strftime('%Y-%m-%d')})",
                                    foreground="blue")
        self.calc_label.pack(side=tk.LEFT, padx=10)

        ttk.Label(main_frame, text="3. 날짜 출력 형식", font=("맑은 고딕", 10, "bold")).pack(anchor="w", pady=(15, 5))

        self.format_var = tk.StringVar(value=initial_format)
        formats = ["%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y년 %m월 %d일"]
        self.format_combo = ttk.Combobox(main_frame, values=formats, textvariable=self.format_var)
        self.format_combo.pack(fill="x", pady=5)

        ttk.Label(main_frame, text="예: %Y%m%d -> 20231027", font=("맑은 고딕", 8), foreground="gray").pack(anchor="w")

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=(30, 0))

        ttk.Button(btn_frame, text="확인", command=self.on_ok, style="Accent.TButton", width=15).pack(side=tk.LEFT,
                                                                                                    padx=10)
        ttk.Button(btn_frame, text="취소", command=self.win.destroy, width=15).pack(side=tk.LEFT, padx=10)

    def on_date_selected(self, event=None):
        if not self.cal: return
        try:
            selected_date = self.cal.selection_get()
            today = datetime.now().date()
            diff = (selected_date - today).days

            # 무한 루프 방지를 위해 값 확인 후 업데이트
            if self.offset_var.get() != diff:
                self.offset_var.set(diff)

            self.calc_label.config(text=f"(계산된 날짜: {selected_date.strftime('%Y-%m-%d')})")
        except:
            pass

    def update_from_offset(self):
        try:
            val = self.offset_var.get()
            target_date = datetime.now().date() + timedelta(days=val)
            self.calc_label.config(text=f"(계산된 날짜: {target_date.strftime('%Y-%m-%d')})")

            if self.cal:
                # 현재 달력 선택과 다를 때만 업데이트
                current_cal_date = self.cal.selection_get()
                if current_cal_date != target_date:
                    self.cal.selection_set(target_date)
        except (tk.TclError, ValueError):
            # 입력 중인 상태일 수 있음
            pass

    def on_ok(self):
        try:
            offset = self.offset_var.get()
            fmt = self.format_var.get()
            self.result = (offset, fmt)
            self.win.destroy()
        except:
            messagebox.showerror("오류", "올바른 숫자를 입력해주세요.")


class FailureActionDialog:
    def __init__(self, parent, action_list, current_action):
        self.parent = parent
        self.result = None
        self.action_list = action_list

        self.win = tk.Toplevel(parent)
        self.win.title("실패 시 동작 설정")
        self.win.geometry("400x300")
        self.win.transient(parent)
        self.win.grab_set()
        center_window(self.win)

        main_frame = ttk.Frame(self.win, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        on_failure = current_action.get("on_failure", {})

        self.enabled_var = tk.BooleanVar(value=on_failure.get("enabled", False))
        self.retries_var = tk.IntVar(value=on_failure.get("retries", 3))
        self.goto_var = tk.StringVar(value=on_failure.get("goto", ""))

        ttk.Checkbutton(main_frame, text="실패 시 재시도/이동 사용", variable=self.enabled_var).pack(anchor="w", pady=(0, 10))

        retries_frame = ttk.Frame(main_frame)
        retries_frame.pack(fill="x", pady=5)
        ttk.Label(retries_frame, text="재시도 횟수:").pack(side=tk.LEFT, padx=(0, 10))
        tk.Spinbox(retries_frame, from_=0, to=100, textvariable=self.retries_var, width=5).pack(side=tk.LEFT)

        goto_frame = ttk.Frame(main_frame)
        goto_frame.pack(fill="x", pady=5)
        ttk.Label(goto_frame, text="이동할 단계:").pack(side=tk.LEFT, padx=(0, 10))

        # 단계 목록 생성
        step_options = [f"{i + 1}: {act.get('alias') or act['type']}" for i, act in enumerate(action_list)]
        self.goto_combo = ttk.Combobox(goto_frame, textvariable=self.goto_var, values=step_options, width=30)
        if on_failure.get("goto"):
            self.goto_combo.set(
                f"{on_failure['goto']}: {action_list[on_failure['goto'] - 1].get('alias') or action_list[on_failure['goto'] - 1]['type']}")
        self.goto_combo.pack(side=tk.LEFT)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=(30, 0))

        ttk.Button(btn_frame, text="확인", command=self.on_ok, style="Accent.TButton").pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="설정 초기화", command=self.on_clear).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="취소", command=self.win.destroy).pack(side=tk.LEFT, padx=10)

    def on_ok(self):
        goto_val = self.goto_var.get()
        goto_step = None
        if goto_val:
            try:
                goto_step = int(goto_val.split(":")[0])
            except (ValueError, IndexError):
                messagebox.showerror("오류", "이동할 단계 형식이 올바르지 않습니다.", parent=self.win)
                return

        self.result = {
            "enabled": self.enabled_var.get(),
            "retries": self.retries_var.get(),
            "goto": goto_step
        }
        self.win.destroy()

    def on_clear(self):
        self.result = {}
        self.win.destroy()


# --- 텍스트 입력 및 암호화 설정을 위한 클래스 ---
class TextEncryptionDialog:
    def __init__(self, parent, initial_text="", initial_encrypted=False):
        self.parent = parent
        self.result = None

        self.win = tk.Toplevel(parent)
        self.win.title("텍스트 입력")
        self.win.geometry("400x200")
        self.win.transient(parent)
        self.win.grab_set()
        center_window(self.win)

        main_frame = ttk.Frame(self.win, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="입력할 텍스트를 적어주세요:").pack(anchor="w")

        self.text_var = tk.StringVar(value=initial_text)
        ttk.Entry(main_frame, textvariable=self.text_var, width=50).pack(fill="x", pady=5)

        self.encrypt_var = tk.BooleanVar(value=initial_encrypted)
        ttk.Checkbutton(main_frame, text="암호화하여 저장 (리스트에 ***로 표시)", variable=self.encrypt_var).pack(anchor="w", pady=10)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=(10, 0))

        ttk.Button(btn_frame, text="확인", command=self.on_ok, style="Accent.TButton").pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="취소", command=self.win.destroy).pack(side=tk.LEFT, padx=10)

    def on_ok(self):
        text = self.text_var.get()
        is_encrypted = self.encrypt_var.get()

        if is_encrypted and not text:
            messagebox.showwarning("경고", "암호화할 내용이 없습니다.", parent=self.win)
            return

        self.result = (text, is_encrypted)
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

        # 왼쪽 정렬 버튼 스타일 (패딩 추가하여 들여쓰기 구현)
        style.configure("Left.TButton", font=("맑은 고딕", 9), anchor="w", padding=(10, 0, 0, 0))

        # Treeview 비활성화 상태 색상 문제 수정을 위한 스타일 맵
        style.map('Treeview',
                  foreground=self._fixed_map(style, 'foreground'),
                  background=self._fixed_map(style, 'background'))

        self.is_running = False
        self.tray_icon = None

        self.processes = {}
        self.current_process = ""
        self.default_delay = 2.0
        self.delay_var = tk.DoubleVar(value=2.0)
        self.schedules = {}
        self.tray_enabled = tk.BooleanVar(value=False)
        self.autostart_enabled = tk.BooleanVar(value=False)
        self.confidence_var = tk.DoubleVar(value=0.8)
        self.pre_type_delay_var = tk.DoubleVar(value=0.0)
        self.char_interval_var = tk.DoubleVar(value=0.0)
        self.post_type_delay_var = tk.DoubleVar(value=0.0)

        self.last_run_date = {}
        self.retry_counts = {}

        self.schedule_type_var = tk.StringVar(value="time")
        self.hour_var = tk.StringVar(value="09")
        self.minute_var = tk.StringVar(value="00")
        self.boot_minute_var = tk.StringVar(value="05")
        self.boot_second_var = tk.StringVar(value="00")

        self.undo_stack = []
        self.redo_stack = []

        self.load_config()

        if self.window_geometry:
            self.root.geometry(self.window_geometry)
        else:
            self.root.geometry("900x750")
            center_window(self.root)

        if self.window_state == 'zoomed':
            self.root.state('zoomed')

        self.root.resizable(True, True)
        self.root.minsize(900, 750)

        # 우클릭 메뉴 생성
        self.tree_menu = tk.Menu(self.root, tearoff=0)
        self.tree_menu.add_command(label="동작 활성화/비활성화", command=self.toggle_action_enabled_from_menu)

        self.create_widgets()
        self.update_treeview()
        self.update_schedule_ui()
        self.sync_autostart_checkbox()
        self.update_undo_redo_buttons()

        # 스케줄러 시작 시 시간을 기록해두기 위한 변수
        self.app_start_time = time.time()
        threading.Thread(target=self.schedule_checker, daemon=True).start()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _fixed_map(self, style, option):
        # Treeview 스타일 관련 버그 수정을 위한 헬퍼 함수
        return [elm for elm in style.map('Treeview', query_opt=option) if
                elm[:2] != ('!disabled', '!selected')]

    def _resolve_path(self, path):
        """저장된 상대 경로를 절대 경로로 변환. 이미 절대 경로인 경우 그대로 반환."""
        if not path or os.path.isabs(path):
            return path
        return os.path.join(application_path, path)

    def _get_relative_path(self, path):
        """절대 경로를 앱 기준의 상대 경로로 변환. 앱 폴더 밖의 경로는 절대 경로로 유지."""
        if not path:
            return ""
        abs_path = os.path.abspath(path)
        app_abs_path = os.path.abspath(application_path)

        # Windows에서 드라이브가 다른 경우 relpath가 오류를 일으키므로 절대 경로 반환
        if os.name == 'nt':
            if os.path.splitdrive(abs_path)[0].lower() != os.path.splitdrive(app_abs_path)[0].lower():
                return abs_path

        # 경로가 애플리케이션 폴더 내에 있는 경우 상대 경로로 변환
        if abs_path.lower().startswith(app_abs_path.lower()):
            try:
                return os.path.relpath(abs_path, app_abs_path)
            except ValueError:
                return abs_path  # 예외 발생 시 절대 경로로 대체

        return abs_path

    # --- 다중 모니터 대응 전체 캡처 및 오프셋 계산 헬퍼 함수 추가 ---
    def get_full_screenshot_and_offset(self):
        v_left, v_top = 0, 0
        if GET_TICK_COUNT_AVAILABLE:
            try:
                import ctypes
                SM_XVIRTUALSCREEN = 76
                SM_YVIRTUALSCREEN = 77
                v_left = ctypes.windll.user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
                v_top = ctypes.windll.user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
            except Exception:
                pass

        screenshot = ImageGrab.grab(all_screens=True)
        return screenshot, v_left, v_top

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

        ttk.Button(top_frame, text="환경설정", command=self.open_settings_window).grid(row=0, column=6, padx=(10, 2),
                                                                                   pady=2, sticky="e")

        schedule_frame = ttk.LabelFrame(self.root, text="프로세스 실행 예약", padding=(10, 5))
        schedule_frame.pack(fill=tk.X, padx=10, pady=5)
        schedule_frame.columnconfigure(1, weight=1)

        self.time_radio = ttk.Radiobutton(schedule_frame, text="특정 시간 (HH:MM)", variable=self.schedule_type_var,
                                          value="time", command=self.toggle_schedule_widgets)
        self.time_radio.grid(row=0, column=0, sticky="w", padx=5)

        self.boot_radio = ttk.Radiobutton(schedule_frame, text="부팅 후 (MM:SS)", variable=self.schedule_type_var,
                                          value="boot", command=self.toggle_schedule_widgets)
        self.boot_radio.grid(row=1, column=0, sticky="w", padx=5)
        if not GET_TICK_COUNT_AVAILABLE:
            self.boot_radio.config(state=tk.DISABLED)

        self.time_input_frame = ttk.Frame(schedule_frame)
        self.time_input_frame.grid(row=0, column=1, sticky="w")
        self.hour_spin = tk.Spinbox(self.time_input_frame, from_=0, to=23, textvariable=self.hour_var, width=3,
                                    format="%02.0f")
        self.hour_spin.pack(side=tk.LEFT)
        ttk.Label(self.time_input_frame, text=":").pack(side=tk.LEFT, padx=2)
        self.minute_spin = tk.Spinbox(self.time_input_frame, from_=0, to=59, textvariable=self.minute_var, width=3,
                                      format="%02.0f")
        self.minute_spin.pack(side=tk.LEFT)

        self.boot_input_frame = ttk.Frame(schedule_frame)
        self.boot_input_frame.grid(row=1, column=1, sticky="w")
        self.boot_minute_spin = tk.Spinbox(self.boot_input_frame, from_=0, to=59, textvariable=self.boot_minute_var,
                                           width=3, format="%02.0f")
        self.boot_minute_spin.pack(side=tk.LEFT)
        ttk.Label(self.boot_input_frame, text=":").pack(side=tk.LEFT, padx=2)
        self.boot_second_spin = tk.Spinbox(self.boot_input_frame, from_=0, to=59, textvariable=self.boot_second_var,
                                           width=3, format="%02.0f")
        self.boot_second_spin.pack(side=tk.LEFT)

        btn_frame = ttk.Frame(schedule_frame)
        btn_frame.grid(row=0, column=2, rowspan=2, padx=(10, 0))
        ttk.Button(btn_frame, text="예약 저장", command=self.save_schedule, style="Accent.TButton").pack(fill=tk.X, pady=2)
        ttk.Button(btn_frame, text="예약 취소", command=self.cancel_schedule).pack(fill=tk.X, pady=2)

        self.toggle_schedule_widgets()

        self.status_label = ttk.Label(self.root, text="예약 없음", font=("맑은 고딕", 11, "bold"), foreground="blue")
        self.status_label.pack(pady=2)

        middle_frame = ttk.Frame(self.root)
        middle_frame.pack(pady=5, fill=tk.BOTH, expand=True, padx=10)
        middle_frame.rowconfigure(0, weight=1)
        middle_frame.columnconfigure(0, weight=1)

        left_frame = ttk.Frame(middle_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        columns = ("#", "활성", "구분", "내용", "실패 시 이동", "재시도")
        self.tree = ttk.Treeview(left_frame, columns=columns, show="headings", selectmode="extended")
        self.tree.heading("#", text="번호", anchor="center")
        self.tree.heading("활성", text="활성", anchor="center")
        self.tree.heading("구분", text="구분", anchor="center")
        self.tree.heading("내용", text="내용")
        self.tree.heading("실패 시 이동", text="실패 시 이동", anchor="center")
        self.tree.heading("재시도", text="재시도", anchor="center")

        self.tree.column("#", width=40, anchor="center", stretch=False)
        self.tree.column("활성", width=50, anchor="center", stretch=False)
        self.tree.column("구분", width=120, anchor="center", stretch=False)
        self.tree.column("내용", width=300, stretch=True)
        self.tree.column("실패 시 이동", width=100, anchor="center", stretch=False)
        self.tree.column("재시도", width=60, anchor="center", stretch=False)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.LEFT, fill="y")

        # 이벤트 바인딩
        self.tree.bind("<Button-1>", self.on_tree_click)
        if sys.platform == "darwin":
            self.tree.bind("<Button-2>", self.show_tree_menu)
        else:
            self.tree.bind("<Button-3>", self.show_tree_menu)

        order_frame = ttk.Frame(middle_frame)
        order_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Button(order_frame, text="▲", command=self.move_up, width=3).pack(pady=2, fill=tk.X)
        ttk.Button(order_frame, text="▼", command=self.move_down, width=3).pack(pady=2, fill=tk.X)

        right_frame = ttk.Frame(middle_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5)

        # 버튼 가로 사이즈 조정 (기존 25에서 22로)
        btn_width = 22

        ttk.Button(right_frame, text="+ 이미지 클릭", command=self.add_click, style="Left.TButton", width=btn_width).pack(
            pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 클릭 & 텍스트", command=self.add_type, style="Left.TButton", width=btn_width).pack(
            pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 텍스트(비밀번호)", command=self.add_password, style="Left.TButton",
                   width=btn_width).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 키 입력(연속 기록)", command=self.add_key, style="Left.TButton", width=btn_width).pack(
            pady=3, fill=tk.X)

        # --- UI 객체 제어 버튼 추가 ---
        ttk.Button(right_frame, text="+ UI 객체 클릭", command=self.add_ui_click, style="Left.TButton",
                   width=btn_width).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ UI 객체 텍스트", command=self.add_ui_type, style="Left.TButton",
                   width=btn_width).pack(pady=3, fill=tk.X)

        ttk.Button(right_frame, text="+ 날짜 입력", command=self.add_date_input, style="Left.TButton",
                   width=btn_width).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 단순 대기(초)", command=self.add_wait, style="Left.TButton", width=btn_width).pack(
            pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 이미지 확인(대기)", command=self.add_wait_image, style="Left.TButton",
                   width=btn_width).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 파일 실행", command=self.add_exec_file, style="Left.TButton", width=btn_width).pack(
            pady=3, fill=tk.X)
        ttk.Button(right_frame, text="+ 폴더 열기", command=self.add_open_path, style="Left.TButton", width=btn_width).pack(
            pady=3, fill=tk.X)

        ttk.Button(right_frame, text="실패 시 동작 설정", command=self.set_failure_action, style="Left.TButton",
                   width=btn_width).pack(pady=(15, 3), fill=tk.X)
        ttk.Button(right_frame, text="작업 이름 변경", command=self.rename_action, style="Left.TButton",
                   width=btn_width).pack(pady=3, fill=tk.X)
        ttk.Button(right_frame, text="작업(내용) 수정", command=self.edit_action, style="Left.TButton", width=btn_width).pack(
            pady=3, fill=tk.X)
        ttk.Button(right_frame, text="작업 삭제", command=self.delete_action, style="Left.TButton", width=btn_width).pack(
            pady=3, fill=tk.X)

        self.undo_btn = ttk.Button(right_frame, text="↶ Undo", command=self.undo_action, style="Left.TButton",
                                   width=btn_width)
        self.undo_btn.pack(pady=(10, 3), fill=tk.X)
        self.redo_btn = ttk.Button(right_frame, text="↷ Redo", command=self.redo_action, style="Left.TButton",
                                   width=btn_width)
        self.redo_btn.pack(pady=3, fill=tk.X)

        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(pady=10, fill=tk.X, padx=10)

        ctrl_frame = ttk.Frame(bottom_frame)
        ctrl_frame.pack(side=tk.TOP, pady=5)

        self.start_btn = ttk.Button(ctrl_frame, text="▶ 처음부터 시작", command=self.start_rpa)
        self.start_btn.grid(row=0, column=0, padx=5)

        self.start_from_btn = ttk.Button(ctrl_frame, text="▶ 선택부터 시작", command=self.start_from_selected)
        self.start_from_btn.grid(row=0, column=1, padx=5)

        self.stop_btn = ttk.Button(ctrl_frame, text="■ 정지", state=tk.DISABLED,
                                   command=self.stop_rpa)
        self.stop_btn.grid(row=0, column=2, padx=5)

    def save_state_for_undo(self):
        self.undo_stack.append(copy.deepcopy(self.processes))
        self.redo_stack.clear()
        self.update_undo_redo_buttons()

    def undo_action(self):
        if self.undo_stack:
            self.redo_stack.append(copy.deepcopy(self.processes))
            self.processes = self.undo_stack.pop()
            self.update_treeview()
            self.update_undo_redo_buttons()
            self.save_config()

    def redo_action(self):
        if self.redo_stack:
            self.undo_stack.append(copy.deepcopy(self.processes))
            self.processes = self.redo_stack.pop()
            self.update_treeview()
            self.update_undo_redo_buttons()
            self.save_config()

    def update_undo_redo_buttons(self):
        self.undo_btn.config(state=tk.NORMAL if self.undo_stack else tk.DISABLED)
        self.redo_btn.config(state=tk.NORMAL if self.redo_stack else tk.DISABLED)

    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column_id = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)

        if not item_id:
            return

        # '활성' 컬럼(#2)이 클릭되었는지 확인
        if column_id == '#2':
            self.toggle_action_enabled(item_id)

    def show_tree_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.tree.focus(item)
            self.tree_menu.post(event.x_root, event.y_root)

    def toggle_action_enabled_from_menu(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        self.toggle_action_enabled(selected_item[0])

    def toggle_action_enabled(self, item_id):
        self.save_state_for_undo()
        idx = self.tree.index(item_id)
        act = self.processes[self.current_process][idx]
        act["enabled"] = not act.get("enabled", True)
        self.save_config()

        # 전체를 다시 그리지 않고 해당 항목만 업데이트
        enabled_display = '☑' if act["enabled"] else '☐'
        self.tree.set(item_id, column="활성", value=enabled_display)

        current_tags = list(self.tree.item(item_id, 'tags'))
        if 'disabled' in current_tags:
            current_tags.remove('disabled')

        if not act["enabled"]:
            current_tags.append('disabled')

        self.tree.item(item_id, tags=tuple(current_tags))

        self.tree.selection_set(item_id)
        self.tree.focus(item_id)

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
        settings_win.geometry("400x500")
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

        delay_frame = ttk.Frame(frame)
        delay_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(delay_frame, text="동작 간 기본 대기시간(초):").pack(side=tk.LEFT)
        delay_spin = tk.Spinbox(delay_frame, from_=0.0, to=10.0, increment=0.5, textvariable=self.delay_var, width=5,
                                command=self.save_config)
        delay_spin.pack(side=tk.RIGHT)
        delay_spin.bind("<KeyRelease>", lambda e: self.save_config())

        confidence_frame = ttk.Frame(frame)
        confidence_frame.pack(fill=tk.X, pady=(10, 0))

        confidence_label = ttk.Label(confidence_frame, text=f"이미지 인식 정확도: {self.confidence_var.get():.2f}")
        confidence_label.pack(anchor="w")

        confidence_scale = tk.Scale(confidence_frame, from_=0.1, to=1.0, resolution=0.05, orient=tk.HORIZONTAL,
                                    variable=self.confidence_var,
                                    command=lambda val: confidence_label.config(text=f"이미지 인식 정확도: {float(val):.2f}"))
        confidence_scale.pack(fill=tk.X)
        confidence_scale.bind("<ButtonRelease-1>", lambda e: self.save_config())

        help_text = "Tip: 이미지를 잘 찾지 못하면 이 값을 낮춰보세요. (예: 0.7)"
        help_label = ttk.Label(confidence_frame, text=help_text, foreground="gray", font=("맑은 고딕", 8))
        help_label.pack(anchor="e")

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=15)

        typing_frame = ttk.LabelFrame(frame, text="타이핑 속도 설정 (초)")
        typing_frame.pack(fill=tk.X, pady=5)

        pre_delay_frame = ttk.Frame(typing_frame)
        pre_delay_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(pre_delay_frame, text="클릭 후 입력 전 대기:").pack(side=tk.LEFT)
        pre_delay_spin = tk.Spinbox(pre_delay_frame, from_=0.0, to=5.0, increment=0.1,
                                    textvariable=self.pre_type_delay_var, width=5, command=self.save_config)
        pre_delay_spin.pack(side=tk.RIGHT)
        pre_delay_spin.bind("<KeyRelease>", lambda e: self.save_config())

        char_interval_frame = ttk.Frame(typing_frame)
        char_interval_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(char_interval_frame, text="글자 입력 간격:").pack(side=tk.LEFT)
        char_interval_spin = tk.Spinbox(char_interval_frame, from_=0.0, to=1.0, increment=0.01,
                                        textvariable=self.char_interval_var, width=5, command=self.save_config)
        char_interval_spin.pack(side=tk.RIGHT)
        char_interval_spin.bind("<KeyRelease>", lambda e: self.save_config())

        post_delay_frame = ttk.Frame(typing_frame)
        post_delay_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(post_delay_frame, text="입력 후 다음 동작 전 대기:").pack(side=tk.LEFT)
        post_delay_spin = tk.Spinbox(post_delay_frame, from_=0.0, to=5.0, increment=0.1,
                                     textvariable=self.post_type_delay_var, width=5, command=self.save_config)
        post_delay_spin.pack(side=tk.RIGHT)
        post_delay_spin.bind("<KeyRelease>", lambda e: self.save_config())

        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=15)

        clean_btn = ttk.Button(frame, text="사용하지 않는 이미지 정리", command=self.clean_unused_images)
        clean_btn.pack(pady=5)

        close_btn = ttk.Button(frame, text="닫기", command=settings_win.destroy)
        close_btn.pack(pady=10)

    def clean_unused_images(self):
        captures_path = os.path.join(application_path, "captures")
        if not os.path.isdir(captures_path):
            messagebox.showinfo("알림", "captures 폴더가 존재하지 않습니다.", parent=self.root)
            return

        used_images = set()
        for proc_name, actions in self.processes.items():
            for act in actions:
                if "image" in act and act["image"]:
                    # 정규화된 경로로 비교하기 위해 절대 경로로 변환
                    used_images.add(os.path.normpath(self._resolve_path(act["image"])))

        all_images_in_folder = set()
        for root, _, files in os.walk(captures_path):
            for name in files:
                if name.lower().endswith(('.png', '.jpg', '.jpeg')):
                    all_images_in_folder.add(os.path.normpath(os.path.join(root, name)))

        unused_images = all_images_in_folder - used_images

        if not unused_images:
            messagebox.showinfo("알림", "사용하지 않는 이미지가 없습니다.", parent=self.root)
            return

        msg = f"{len(unused_images)}개의 사용하지 않는 이미지를 찾았습니다. 삭제하시겠습니까?\n\n"
        msg += "\n".join([os.path.basename(p) for p in list(unused_images)[:10]])  # 최대 10개만 미리보기
        if len(unused_images) > 10:
            msg += "\n..."

        if messagebox.askyesno("이미지 정리", msg, parent=self.root):
            deleted_count = 0
            error_files = []
            for img_path in unused_images:
                try:
                    os.remove(img_path)
                    deleted_count += 1
                except OSError as e:
                    error_files.append(os.path.basename(img_path))

            final_msg = f"{deleted_count}개의 이미지를 삭제했습니다."
            if error_files:
                final_msg += f"\n\n다음 파일들은 삭제하지 못했습니다:\n" + "\n".join(error_files)

            messagebox.showinfo("완료", final_msg, parent=self.root)

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
                    self.status_label.config(text=f"'{self.current_process}' 예약됨 (매일 {schedule_value})",
                                             foreground="green")
                except (ValueError, TypeError):
                    self.status_label.config(text="예약 없음", foreground="blue")
            elif schedule_type == "boot":
                try:
                    minute, second = schedule_value.split(":")
                    self.boot_minute_var.set(f"{int(minute):02d}")
                    self.boot_second_var.set(f"{int(second):02d}")
                    self.status_label.config(text=f"'{self.current_process}' 예약됨 (부팅 후 {schedule_value})",
                                             foreground="purple")
                except (ValueError, TypeError):
                    self.status_label.config(text="예약 없음", foreground="blue")
        else:
            self.schedule_type_var.set("time")
            self.status_label.config(text="예약 없음", foreground="blue")

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
            initial_dir = os.path.normpath(os.path.join(application_path, "captures"))
            self.root.update_idletasks()  # UI 안정화
            file_path = filedialog.askopenfilename(title="이미지 선택",
                                                   initialdir=initial_dir,
                                                   filetypes=[("Image files", "*.png *.jpg")])
            self.root.focus_force()  # 포커스 강제 회수

        else:  # Cancel
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
            initial_dir = os.path.normpath(os.path.join(application_path, "captures"))
            self.root.update_idletasks()  # UI 안정화
            selected = filedialog.askopenfilename(title="이미지 선택",
                                                  initialdir=initial_dir,
                                                  filetypes=[("Image files", "*.png *.jpg")])
            self.root.focus_force()  # 포커스 강제 회수
            return selected

        else:
            if is_edit and current_action:
                return current_action.get("image")
            return None
        return None

    def add_click(self):
        self.save_state_for_undo()
        file_path, click_pos, click_type = self.get_image_and_click_pos()
        if file_path:
            action = {"type": "click", "image": self._get_relative_path(file_path), "alias": "", "click_pos": click_pos,
                      "click_type": click_type, "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def add_type(self):
        self.save_state_for_undo()
        file_path, click_pos, click_type = self.get_image_and_click_pos()
        if file_path:
            text = simpledialog.askstring("텍스트 입력", "입력할 텍스트를 적어주세요:", parent=self.root)
            if text is not None:
                action = {"type": "type", "image": self._get_relative_path(file_path), "text": text, "alias": "",
                          "click_pos": click_pos, "click_type": click_type, "enabled": True}
                self.processes[self.current_process].append(action)
                self.update_treeview()
                self.save_config()

    def add_password(self):
        self.save_state_for_undo()
        dialog = TextEncryptionDialog(self.root)
        self.root.wait_window(dialog.win)

        if dialog.result:
            text, is_encrypted = dialog.result
            action = {"type": "password", "text": text, "encrypted": is_encrypted, "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    # --- UI 객체 동작(pywinauto) 추가 ---
    def add_ui_click(self):
        if not PYWINAUTO_AVAILABLE:
            messagebox.showerror("오류", "pywinauto 라이브러리가 필요합니다.\npip install pywinauto", parent=self.root)
            return
        self.save_state_for_undo()
        dialog = UIObjectSettingDialog(self.root)
        self.root.wait_window(dialog.win)

        if dialog.result:
            action = {"type": "ui_click", "identifiers": dialog.result, "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def add_ui_type(self):
        if not PYWINAUTO_AVAILABLE:
            messagebox.showerror("오류", "pywinauto 라이브러리가 필요합니다.\npip install pywinauto", parent=self.root)
            return
        self.save_state_for_undo()
        dialog = UIObjectSettingDialog(self.root)
        self.root.wait_window(dialog.win)

        if dialog.result:
            text = simpledialog.askstring("텍스트 입력", "UI 객체에 입력할 텍스트를 적어주세요:", parent=self.root)
            if text is not None:
                action = {"type": "ui_type", "identifiers": dialog.result, "text": text, "alias": "", "enabled": True}
                self.processes[self.current_process].append(action)
                self.update_treeview()
                self.save_config()

    def add_wait_image(self):
        self.save_state_for_undo()
        file_path = self.get_image_path_for_wait()
        if file_path:
            timeout = simpledialog.askfloat("타임아웃 설정", "이미지가 나타날 때까지 기다릴 최대 시간(초)을 입력하세요\n(예: 10초 이내에 안 나타나면 오류 처리):",
                                            initialvalue=10.0, parent=self.root)
            if timeout is not None:
                action = {"type": "wait_image", "image": self._get_relative_path(file_path), "timeout": timeout,
                          "alias": "", "enabled": True}
                self.processes[self.current_process].append(action)
                self.update_treeview()
                self.save_config()

    def add_key(self):
        self.save_state_for_undo()
        recorder = KeyRecorder(self.root)
        self.root.wait_window(recorder.win)

        if recorder.recorded_actions:
            action = {"type": "key", "keys": recorder.recorded_actions, "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def add_date_input(self):
        self.save_state_for_undo()
        dialog = DateSettingDialog(self.root)
        self.root.wait_window(dialog.win)
        if dialog.result:
            offset, fmt = dialog.result
            action = {"type": "date_input", "offset": offset, "format": fmt, "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def add_wait(self):
        self.save_state_for_undo()
        sec = simpledialog.askfloat("대기 시간", "기다릴 시간(초)을 입력하세요 (예: 1.5):", parent=self.root)
        if sec is not None:
            action = {"type": "wait", "time": sec, "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def add_exec_file(self):
        self.save_state_for_undo()
        initial_dir = os.path.normpath(application_path)
        self.root.update_idletasks()  # UI 안정화
        # parent 인자 제거 및 포커스 관리 로직 추가
        file_path = filedialog.askopenfilename(title="실행할 파일 선택", initialdir=initial_dir)
        self.root.focus_force()  # 포커스 강제 회수

        if file_path:
            action = {"type": "exec_file", "path": self._get_relative_path(file_path), "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def add_open_path(self):
        self.save_state_for_undo()
        initial_dir = os.path.normpath(application_path)
        self.root.update_idletasks()  # 다이얼로그 호출 전 UI 상태 업데이트

        # parent=self.root를 제거하여 COM 스레딩 충돌이나 숨김 현상 방지
        dir_path = filedialog.askdirectory(title="열고 싶은 폴더 선택", initialdir=initial_dir)

        self.root.focus_force()  # 선택 후 포커스 복귀

        if dir_path:
            action = {"type": "open_path", "path": self._get_relative_path(dir_path), "alias": "", "enabled": True}
            self.processes[self.current_process].append(action)
            self.update_treeview()
            self.save_config()

    def rename_action(self):
        self.save_state_for_undo()
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
        self.save_state_for_undo()
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("선택 오류", "수정할 작업을 선택해주세요.")
            return

        idx = self.tree.index(selected_item[0])
        act = self.processes[self.current_process][idx]

        if act["type"] in ["click", "type"]:
            file_path, click_pos, click_type = self.get_image_and_click_pos(is_edit=True, current_action=act)
            if file_path:
                act["image"] = self._get_relative_path(file_path)
                act["click_pos"] = click_pos
                act["click_type"] = click_type

            if act["type"] == "type":
                new_text = simpledialog.askstring("텍스트 수정", "새로운 텍스트를 입력하세요:", initialvalue=act.get("text", ""),
                                                  parent=self.root)
                if new_text is not None:
                    act["text"] = new_text

        elif act["type"] == "password":
            is_currently_encrypted = act.get("encrypted", True)
            dialog = TextEncryptionDialog(self.root, initial_text=act.get("text", ""),
                                          initial_encrypted=is_currently_encrypted)
            self.root.wait_window(dialog.win)

            if dialog.result:
                text, is_encrypted = dialog.result
                act["text"] = text
                act["encrypted"] = is_encrypted

        elif act["type"] in ["ui_click", "ui_type"]:
            if not PYWINAUTO_AVAILABLE:
                messagebox.showerror("오류", "pywinauto 라이브러리가 필요합니다.", parent=self.root)
                return
            dialog = UIObjectSettingDialog(self.root, initial_data=act.get("identifiers", {}))
            self.root.wait_window(dialog.win)
            if dialog.result:
                act["identifiers"] = dialog.result

            if act["type"] == "ui_type" and dialog.result:
                new_text = simpledialog.askstring("텍스트 수정", "새로운 텍스트를 입력하세요:", initialvalue=act.get("text", ""),
                                                  parent=self.root)
                if new_text is not None:
                    act["text"] = new_text

        elif act["type"] == "wait_image":
            file_path = self.get_image_path_for_wait(is_edit=True, current_action=act)
            if file_path:
                act["image"] = self._get_relative_path(file_path)
            new_timeout = simpledialog.askfloat("타임아웃 수정", "새로운 최대 대기 시간(초)을 입력하세요:",
                                                initialvalue=act.get("timeout", 10.0), parent=self.root)
            if new_timeout is not None:
                act["timeout"] = new_timeout

        elif act["type"] == "key":
            current_keys = ", ".join(act.get("keys", [act.get("key", "")]))
            new_keys_str = simpledialog.askstring("키보드 입력 수정", "새로운 키들을 콤마(,)로 구분하여 입력하세요:",
                                                  initialvalue=current_keys, parent=self.root)
            if new_keys_str:
                act["keys"] = [k.strip().lower() for k in new_keys_str.split(",")]
                if "key" in act: del act["key"]

        elif act["type"] == "date_input":
            dialog = DateSettingDialog(self.root, initial_offset=act.get("offset", 0),
                                       initial_format=act.get("format", "%Y%m%d"))
            self.root.wait_window(dialog.win)
            if dialog.result:
                offset, fmt = dialog.result
                act["offset"] = offset
                act["format"] = fmt

        elif act["type"] == "wait":
            new_time = simpledialog.askfloat("대기 시간 수정", "새로운 대기 시간(초)을 입력하세요:", initialvalue=act.get("time", 1.0),
                                             parent=self.root)
            if new_time is not None:
                act["time"] = new_time

        elif act["type"] == "exec_file":
            initial_file = self._resolve_path(act.get("path"))
            initial_dir = os.path.dirname(initial_file) if initial_file else application_path
            initial_dir = os.path.normpath(initial_dir)
            self.root.update_idletasks()
            new_path = filedialog.askopenfilename(title="실행할 파일 선택", initialdir=initial_dir, initialfile=initial_file)
            self.root.focus_force()
            if new_path:
                act["path"] = self._get_relative_path(new_path)

        elif act["type"] == "open_path":
            initial_dir = self._resolve_path(act.get("path")) if act.get("path") else application_path
            initial_dir = os.path.normpath(initial_dir)
            self.root.update_idletasks()
            new_path = filedialog.askdirectory(title="열고 싶은 폴더 선택", initialdir=initial_dir)
            self.root.focus_force()
            if new_path:
                act["path"] = self._get_relative_path(new_path)

        self.update_treeview()
        self.save_config()

    def set_failure_action(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("선택 오류", "설정할 작업을 선택해주세요.")
            return

        idx = self.tree.index(selected_item[0])
        current_actions = self.processes[self.current_process]
        act = current_actions[idx]

        dialog = FailureActionDialog(self.root, current_actions, act)
        self.root.wait_window(dialog.win)

        if dialog.result is not None:
            self.save_state_for_undo()
            if not dialog.result:
                if "on_failure" in act:
                    del act["on_failure"]
            else:
                act["on_failure"] = dialog.result
            self.update_treeview()
            self.save_config()

    def delete_action(self):
        self.save_state_for_undo()
        selected_items = self.tree.selection()
        if selected_items:
            # 여러 항목 삭제 시 인덱스가 변경되므로 뒤에서부터 삭제
            indices = sorted([self.tree.index(item) for item in selected_items], reverse=True)
            for idx in indices:
                del self.processes[self.current_process][idx]
            self.update_treeview()
            self.save_config()

    def move_up(self):
        self.save_state_for_undo()
        selected_items = self.tree.selection()
        if not selected_items:
            return

        selected_indices = sorted([self.tree.index(item) for item in selected_items])

        if selected_indices[0] == 0:
            return

        actions = self.processes[self.current_process]

        for idx in selected_indices:
            item_to_move = actions.pop(idx)
            actions.insert(idx - 1, item_to_move)

            # UI 상에서 아이템 이동
            item_id = self.tree.get_children()[idx]
            self.tree.move(item_id, '', idx - 1)

        self.save_config()
        self.renumber_treeview()

        # 선택 상태 복원
        new_selection_ids = [self.tree.get_children()[i - 1] for i in selected_indices]
        self.tree.selection_set(new_selection_ids)
        if new_selection_ids:
            self.tree.focus(new_selection_ids[0])
            self.tree.see(new_selection_ids[0])

    def move_down(self):
        self.save_state_for_undo()
        selected_items = self.tree.selection()
        if not selected_items:
            return

        selected_indices = sorted([self.tree.index(item) for item in selected_items], reverse=True)

        if selected_indices[0] >= len(self.processes[self.current_process]) - 1:
            return

        actions = self.processes[self.current_process]

        for idx in selected_indices:
            item_to_move = actions.pop(idx)
            actions.insert(idx + 1, item_to_move)

            # UI 상에서 아이템 이동
            item_id = self.tree.get_children()[idx]
            self.tree.move(item_id, '', idx + 1)

        self.save_config()
        self.renumber_treeview()

        # 선택 상태 복원
        new_selection_ids = [self.tree.get_children()[i + 1] for i in selected_indices]
        self.tree.selection_set(new_selection_ids)
        if new_selection_ids:
            self.tree.focus(new_selection_ids[-1])
            self.tree.see(new_selection_ids[-1])

    def renumber_treeview(self):
        for i, item_id in enumerate(self.tree.get_children()):
            self.tree.set(item_id, column="#", value=i + 1)

    def update_treeview(self):
        # 현재 선택 및 스크롤 위치 저장
        selection = self.tree.selection()
        scroll_pos = self.tree.yview()

        for i in self.tree.get_children():
            self.tree.delete(i)

        current_actions = self.processes.get(self.current_process, [])
        for i, act in enumerate(current_actions):
            alias = act.get("alias", "")
            enabled = act.get("enabled", True)
            enabled_display = '☑' if enabled else '☐'

            act_type_display = ""
            content_display = ""

            path_in_act = act.get("image") or act.get("path")

            if act["type"] == "click":
                act_type_display = "[더블클릭]" if act.get("click_type") == "double" else "[클릭]"
                content_display = alias if alias else os.path.basename(self._resolve_path(act["image"]))
            elif act["type"] == "type":
                act_type_display = "[더블클릭 & 입력]" if act.get("click_type") == "double" else "[클릭 & 입력]"
                image_name = os.path.basename(self._resolve_path(act['image']))
                content_display = alias if alias else f"{image_name} -> '{act['text']}'"
            elif act["type"] == "password":
                is_encrypted = act.get("encrypted", True)
                act_type_display = "[텍스트 입력]"
                if is_encrypted:
                    content_display = alias if alias else "'***' 텍스트 입력"
                else:
                    content_display = alias if alias else f"'{act.get('text', '')}' 텍스트 입력"
            elif act["type"] == "ui_click":
                ids = act.get("identifiers", {})
                act_type_display = "[UI 더블클릭]" if ids.get("double_click") else "[UI 클릭]"
                display_id = ids.get("title") or ids.get("auto_id") or ids.get("control_type")
                content_display = alias if alias else f"객체({display_id}) 클릭"
            elif act["type"] == "ui_type":
                ids = act.get("identifiers", {})
                act_type_display = "[UI 더블클릭&입력]" if ids.get("double_click") else "[UI 텍스트]"
                display_id = ids.get("title") or ids.get("auto_id") or ids.get("control_type")
                content_display = alias if alias else f"객체({display_id}) -> '{act.get('text', '')}'"
            elif act["type"] == "key":
                act_type_display = "[키보드]"
                keys_display = " -> ".join(act.get("keys", [act.get("key", "")]))
                content_display = alias if alias else keys_display
            elif act["type"] == "date_input":
                act_type_display = "[날짜 입력]"
                content_display = alias if alias else f"오늘로부터 {act['offset']}일 ({act['format']})"
            elif act["type"] == "wait":
                act_type_display = "[대기]"
                content_display = alias if alias else f"{act['time']}초"
            elif act["type"] == "wait_image":
                act_type_display = "[이미지 확인]"
                image_name = os.path.basename(self._resolve_path(act['image']))
                content_display = alias if alias else f"{image_name} (최대 {act['timeout']}초)"
            elif act["type"] == "exec_file":
                act_type_display = "[파일 실행]"
                content_display = alias if alias else os.path.basename(self._resolve_path(act["path"]))
            elif act["type"] == "open_path":
                act_type_display = "[폴더 열기]"
                content_display = alias if alias else os.path.basename(self._resolve_path(act["path"]))

            on_failure = act.get("on_failure")
            goto_val = ""
            retries_val = ""
            if on_failure and on_failure.get("enabled"):
                goto_val = on_failure.get('goto', '')
                retries_val = on_failure.get('retries', '')

            tags = []
            if not enabled:
                tags.append('disabled')
            if act.get("click_pos") or act["type"] in ["ui_click", "ui_type"]:
                tags.append('has_click_pos')

            self.tree.tag_configure('disabled', foreground='gray')
            self.tree.tag_configure('has_click_pos', foreground='purple')

            self.tree.insert("", "end",
                             values=(i + 1, enabled_display, act_type_display, content_display, goto_val, retries_val),
                             tags=tuple(tags))

        # 선택 및 스크롤 위치 복원
        try:
            if selection:
                self.tree.selection_set(selection)
                if selection:
                    self.tree.focus(selection[0])
        except tk.TclError:
            # 항목이 존재하지 않으면 선택을 복원하지 않음
            pass
        self.tree.yview_moveto(scroll_pos[0])

    def save_config(self):
        geometry = ""
        if self.root.state() == 'normal':
            geometry = self.root.geometry()

        processes_to_save = json.loads(json.dumps(self.processes))
        for proc, actions in processes_to_save.items():
            for act in actions:
                if act.get("type") == "password":
                    if act.get("encrypted", True) and "text" in act:
                        act["text"] = base64.b64encode(act["text"].encode('utf-8')).decode('utf-8')

        data = {
            "settings": {
                "default_delay": self.delay_var.get(),
                "tray_enabled": self.tray_enabled.get(),
                "autostart_enabled": self.autostart_enabled.get(),
                "window_geometry": geometry,
                "window_state": self.root.state(),
                "confidence": self.confidence_var.get(),
                "pre_type_delay": self.pre_type_delay_var.get(),
                "char_interval": self.char_interval_var.get(),
                "post_type_delay": self.post_type_delay_var.get()
            },
            "schedules": self.schedules,
            "processes": processes_to_save
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
                    self.delay_var.set(self.default_delay)
                    self.tray_enabled.set(settings.get("tray_enabled", False))
                    self.window_geometry = settings.get("window_geometry", "")
                    self.window_state = settings.get("window_state", "normal")
                    self.confidence_var.set(settings.get("confidence", 0.8))
                    self.pre_type_delay_var.set(settings.get("pre_type_delay", 0.0))
                    self.char_interval_var.set(settings.get("char_interval", 0.0))
                    self.post_type_delay_var.set(settings.get("post_type_delay", 0.0))

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

                if act.get("type") in ["ui_click", "ui_type"]:
                    if "identifiers" in act and "double_click" not in act["identifiers"]:
                        act["identifiers"]["double_click"] = False

                if act.get("type") == "password":
                    is_encrypted = act.get("encrypted", True)
                    if is_encrypted and "text" in act:
                        try:
                            act["text"] = base64.b64decode(act["text"].encode('utf-8')).decode('utf-8')
                        except:
                            pass

    def schedule_checker(self):
        # 시작 시 대기 (EMR 시퀀스 프로그램 UI가 뜨고 로드될 시간을 줌)
        time.sleep(5)
        
        # 프로그램이 시작된 후 기준 시간을 잡기 위한 시작 uptime 기록
        start_uptime = 0
        if GET_TICK_COUNT_AVAILABLE:
            start_uptime = ctypes.windll.kernel32.GetTickCount64() / 1000
            
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
                        
                        # "부팅 후(uptime_seconds)" 시간이 프로그램이 켜진 시점(start_uptime)보다 이후인지,
                        # 혹은 프로그램이 부팅 후 지정 시간 근처에서 시작되었는지를 판단
                        # 목표 시간에 도달했거나 이미 지난 경우이면서,
                        # 프로그램이 켜진 시점이 목표 시간 이전이었을 때만 실행하도록 변경
                        if uptime_seconds >= target_seconds and start_uptime <= target_seconds + 30:
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
        self.retry_counts = {}
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

    # --- 클릭 동작(다중 모니터 좌표 보정 반영) ---
    def execute_click(self, image_path, click_pos, click_type="single"):
        try:
            resolved_path = self._resolve_path(image_path)
            img = Image.open(resolved_path)

            # 전체 화면 캡처 및 오프셋 정보 획득
            screenshot, v_left, v_top = self.get_full_screenshot_and_offset()

            # pyautogui.locate 대신 전체 스크린샷 위에서 locate
            location = pyautogui.locate(img, screenshot, confidence=self.confidence_var.get(), grayscale=True)

            if location:
                # 찾은 이미지 내의 상대 클릭 좌표
                if click_pos:
                    rel_x, rel_y = click_pos
                else:
                    rel_x = location.width / 2
                    rel_y = location.height / 2

                # 전체 가상 모니터 기준 오프셋(v_left, v_top) 더하기
                click_x = v_left + location.left + rel_x
                click_y = v_top + location.top + rel_y

                if click_type == "double":
                    pyautogui.doubleClick(click_x, click_y)
                else:
                    pyautogui.click(click_x, click_y)

                time.sleep(0.5)
                return True
            return False
        except Exception as e:
            print(f"이미지 읽기 또는 클릭 오류 ({resolved_path}): {e}")
            return False

    def type_text_char_by_char(self, text):
        original_clipboard = None
        try:
            original_clipboard = pyperclip.paste()
        except Exception:
            pass

        for ch in text:
            if not self.is_running: break
            pyperclip.copy(ch)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(self.char_interval_var.get())

        if original_clipboard is not None:
            try:
                pyperclip.copy(original_clipboard)
            except Exception:
                pass

    def rpa_task(self, actions, global_delay, start_index=0):
        i = start_index
        while i < len(actions):
            if not self.is_running:
                break

            act = actions[i]

            self.root.after(0, lambda item_id=self.tree.get_children()[i]: (
                self.tree.selection_set(item_id), self.tree.see(item_id)
            ))

            if not act.get("enabled", True):
                i += 1
                continue

            try:
                if act["type"] == "click":
                    if not self.execute_click(act["image"], act.get("click_pos"), act.get("click_type", "single")):
                        raise Exception(f"이미지를 찾을 수 없습니다: {os.path.basename(act['image'])}")

                elif act["type"] == "type":
                    if not self.execute_click(act["image"], act.get("click_pos"), act.get("click_type", "single")):
                        raise Exception(f"이미지를 찾을 수 없습니다: {os.path.basename(act['image'])}")
                    time.sleep(self.pre_type_delay_var.get())
                    self.type_text_char_by_char(act["text"])
                    time.sleep(self.post_type_delay_var.get())

                elif act["type"] == "password":
                    time.sleep(self.pre_type_delay_var.get())
                    self.type_text_char_by_char(act["text"])
                    time.sleep(self.post_type_delay_var.get())

                elif act["type"] in ["ui_click", "ui_type"]:
                    if not PYWINAUTO_AVAILABLE:
                        raise Exception("pywinauto 라이브러리가 설치되어 있지 않습니다.")

                    ids = act.get("identifiers", {})
                    kwargs = {}
                    if ids.get("auto_id"): kwargs["auto_id"] = ids["auto_id"]
                    if ids.get("title"): kwargs["title"] = ids["title"]
                    if ids.get("control_type"): kwargs["control_type"] = ids["control_type"]

                    if not kwargs:
                        raise Exception("찾을 UI 객체 조건이 없습니다.")

                    try:
                        # Desktop에서 전체 윈도우를 대상으로 객체 검색
                        desktop = Desktop(backend="uia")
                        element = desktop.window(**kwargs)
                        # 요소가 준비될 때까지 잠시 대기 (최대 10초)
                        element.wait('ready', timeout=10)

                        # 창을 최상단으로 올리기 (포커스)
                        try:
                            element.set_focus()
                        except Exception:
                            pass

                        # 요소 클릭
                        if ids.get("double_click"):
                            element.double_click_input()
                        else:
                            element.click_input()

                        # ui_type인 경우 텍스트 입력 추가
                        if act["type"] == "ui_type":
                            time.sleep(self.pre_type_delay_var.get())
                            self.type_text_char_by_char(act["text"])
                            time.sleep(self.post_type_delay_var.get())

                    except Exception as e:
                        raise Exception(f"UI 객체를 찾거나 제어할 수 없습니다.\n상세: {e}")

                elif act["type"] == "key":
                    keys = act.get("keys", [act.get("key", "")])
                    for k in keys:
                        if not self.is_running: break
                        if k:
                            if "+" in k:
                                # 복합키 처리 (ex: ctrl+c)
                                k_parts = k.split("+")
                                pyautogui.hotkey(*k_parts)
                            else:
                                pyautogui.press(k)
                            time.sleep(0.01)
                    time.sleep(0.5)

                elif act["type"] == "date_input":
                    target_date = datetime.now() + timedelta(days=act.get("offset", 0))
                    date_str = target_date.strftime(act.get("format", "%Y-%m-%d"))
                    self.type_text_char_by_char(date_str)
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
                    img_path = self._resolve_path(act["image"])
                    try:
                        img = Image.open(img_path)
                    except Exception:
                        raise Exception(f"이미지를 열 수 없습니다: {os.path.basename(img_path)}")

                    start_time = time.time()
                    found = False

                    while time.time() - start_time < timeout:
                        if not self.is_running: break
                        try:
                            # 대기 시 이미지 찾기(다중 모니터 대응)
                            screenshot, _, _ = self.get_full_screenshot_and_offset()
                            location = pyautogui.locate(img, screenshot, confidence=self.confidence_var.get(),
                                                        grayscale=True)

                            if location:
                                found = True
                                break
                        except Exception:
                            pass
                        time.sleep(0.3)

                    if self.is_running and not found:
                        raise Exception(f"시간 초과: {timeout}초 내에 이미지를 찾을 수 없습니다\n({os.path.basename(img_path)})")

                elif act["type"] == "exec_file" or act["type"] == "open_path":
                    path = self._resolve_path(act["path"])
                    if not os.path.exists(path):
                        raise FileNotFoundError(f"경로를 찾을 수 없습니다: {path}")
                    try:
                        if sys.platform == "win32":
                            os.startfile(path)
                        elif sys.platform == "darwin":
                            subprocess.run(["open", path])
                        else:
                            subprocess.run(["xdg-open", path])
                    except Exception as e:
                        raise Exception(f"경로를 열 수 없습니다: {os.path.basename(path)}\n{e}")

                # 성공 시 재시도 횟수 초기화
                if i in self.retry_counts:
                    self.retry_counts[i] = 0

                i += 1  # 다음 단계로

            except Exception as e:
                failure_config = act.get("on_failure")
                if failure_config and failure_config.get("enabled"):
                    goto_step = failure_config.get("goto")
                    if goto_step is not None and 1 <= goto_step <= len(actions):
                        retry_key = (i, goto_step - 1)
                        current_retries = self.retry_counts.get(retry_key, 0)
                        max_retries = failure_config.get("retries", 0)

                        if current_retries < max_retries:
                            self.retry_counts[retry_key] = current_retries + 1
                            i = goto_step - 1
                            time.sleep(1)
                            continue
                        else:
                            self.root.after(0, lambda e=e: self.status_label.config(text="오류 발생!", foreground="red"))
                            self.root.after(0, lambda e=e: messagebox.showerror("오류",
                                                                                f"최대 재시도 횟수({max_retries}회)를 초과했습니다.\n\n{str(e)}"))
                            break
                    else:  # goto_step이 설정되지 않은 경우
                        self.root.after(0, lambda e=e: self.status_label.config(text="오류 발생!", foreground="red"))
                        self.root.after(0, lambda e=e: messagebox.showerror("오류", f"작업 중 오류가 발생했습니다.\n\n{str(e)}"))
                        break
                else:
                    self.root.after(0, lambda e=e: self.status_label.config(text="오류 발생!", foreground="red"))
                    self.root.after(0, lambda e=e: messagebox.showerror("오류", f"작업 중 오류가 발생했습니다.\n\n{str(e)}"))
                    break

            if i < len(actions) and self.is_running:
                delay_elapsed = 0
                while delay_elapsed < global_delay:
                    if not self.is_running: break
                    time.sleep(0.1)
                    delay_elapsed += 0.1

        if i == len(actions) and self.is_running:
            self.root.after(0, lambda: self.status_label.config(text="작업 완료!", foreground="green"))
        elif not self.is_running:
            self.root.after(0, lambda: self.status_label.config(text="사용자 중단", foreground="orange"))

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
