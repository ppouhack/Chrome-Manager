import tkinter as tk
from tkinter import ttk, messagebox
import os
import subprocess
import win32gui
import win32process
import win32con
import win32api
import win32com.client
import json
from typing import List, Dict, Optional
import math
import ctypes
from ctypes import wintypes
import threading
import time
import sys
import keyboard
import mouse
import webbrowser
import sv_ttk

def is_admin():
    # 관리자 권한 확인
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    # 관리자 권한으로 프로그램 재실행
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

class ChromeManager:
    def __init__(self):
        
        if not is_admin():
            if messagebox.askyesno("권한 부족", "동기화 기능을 실행하려면 관리자 권한이 필요합니다.\n프로그램을 관리자 권한으로 다시 시작하시겠습니까?"):
                run_as_admin()
                sys.exit()
                
        self.root = tk.Tk()
        self.root.title("NoBiggie 커뮤니티 Chrome 다중 창 관리자 V1.0")
        
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "app.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"아이콘 설정 실패: {str(e)}")
        
        last_position = self.load_window_position()
        if last_position:
            self.root.geometry(last_position)
        
        sv_ttk.set_theme("light")  # 사용 light 테마
        
        self.window_list = None  # 먼저 None으로 초기화
        self.windows = []
        self.master_window = None
        self.shortcut_path = self.load_settings().get('shortcut_path', '')
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.select_all_var = tk.StringVar(value="전체 선택")
        
        self.is_syncing = False
        self.sync_button = None
        self.mouse_hook_id = None
        self.keyboard_hook = None
        self.hook_thread = None
        self.user32 = ctypes.WinDLL('user32', use_last_error=True)
        self.sync_windows = []
        
        self.chrome_drivers = {}
        self.debug_ports = {}
        self.base_debug_port = 9222
        
        self.DWMWA_BORDER_COLOR = 34
        self.DWM_MAGIC_COLOR = 0x00FF0000
        
        self.popup_mappings = {}
        
        self.popup_monitor_thread = None
        
        self.mouse_threshold = 3
        self.last_mouse_position = (0, 0)
        self.last_move_time = 0
        self.move_interval = 0.016
        
        self.shortcut_hook = None
        self.current_shortcut = None
        
        # 설정에서 단축키 로드
        settings = self.load_settings()
        if 'sync_shortcut' in settings:
            self.set_shortcut(settings['sync_shortcut'])
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 인터페이스 요소 생성
        self.create_widgets()  
        self.create_styles()   
        
       
        self.root.update()
        current_width = self.root.winfo_width()
        current_height = self.root.winfo_height()
        self.root.geometry(f"{current_width}x{current_height}")
        self.root.resizable(False, False)

    def create_styles(self):
        style = ttk.Style()
        
        default_font = ('Microsoft YaHei UI', 9)
        
        style.configure('Small.TEntry',
            padding=(4, 0),
            font=default_font
        )
                
        style.configure('TButton', font=default_font)
        style.configure('TLabel', font=default_font)
        style.configure('TEntry', font=default_font)
        style.configure('Treeview', font=default_font)
        style.configure('Treeview.Heading', font=default_font)
        style.configure('TLabelframe.Label', font=default_font)
        style.configure('TNotebook.Tab', font=default_font)
        
        if self.window_list:
            self.window_list.tag_configure("master", 
                background="#0d6efd",
                foreground='white'
            )
        
        # 링크 스타일
        style.configure('Link.TLabel',
            foreground='#0d6efd',
            cursor='hand2',
            font=('Microsoft YaHei UI', 9, 'underline')
        )

    def create_widgets(self):
        # 인터페이스 요소 생성
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.X, padx=10, pady=5)
        
        upper_frame = ttk.Frame(main_frame)
        upper_frame.pack(fill=tk.X)
        
        arrange_frame = ttk.LabelFrame(upper_frame, text="사용자 지정 배열")
        arrange_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(3, 0))
        
        manage_frame = ttk.LabelFrame(upper_frame, text="창 관리")
        manage_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        button_frame = ttk.Frame(manage_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="창 가져오기", command=self.import_windows, style='Accent.TButton').pack(side=tk.LEFT, padx=2)
        select_all_label = ttk.Label(button_frame, textvariable=self.select_all_var, style='Link.TLabel')
        select_all_label.pack(side=tk.LEFT, padx=5)
        select_all_label.bind('<Button-1>', self.toggle_select_all)
        ttk.Button(button_frame, text="자동 배열", command=self.auto_arrange_windows).pack(side=tk.LEFT, padx=2)
        ttk.Button(button_frame, text="선택 닫기", command=self.close_selected_windows).pack(side=tk.LEFT, padx=2)
        
        self.sync_button = ttk.Button(
            button_frame, 
            text="▶ 시작 동기화",
            command=self.toggle_sync,
            style='Accent.TButton'
        )
        self.sync_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame,
            text="단축키",
            command=self.show_shortcut_dialog,
            style='Accent.TButton'
        ).pack(side=tk.LEFT, padx=5)
        
        list_frame = ttk.Frame(manage_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=2)
        
        # 창 목록 생성
        self.window_list = ttk.Treeview(list_frame, 
            columns=("select", "number", "title", "master", "hwnd"),
            show="headings", 
            height=4,  
            style='Accent.Treeview'
        )
        self.window_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.window_list.heading("select", text="선택")
        self.window_list.heading("number", text="번호")
        self.window_list.heading("title", text="제목")
        self.window_list.heading("master", text="마스터")
        self.window_list.heading("hwnd", text="")
        
        self.window_list.column("select", width=40, anchor="center")
        self.window_list.column("number", width=40, anchor="center")
        self.window_list.column("title", width=300)
        self.window_list.column("master", width=40, anchor="center")
        self.window_list.column("hwnd", width=0, stretch=False)  # hwnd 열 숨기기
        
        self.window_list.tag_configure("master", background="lightblue")
        
        self.window_list.bind('<Button-1>', self.on_click)
        
        # 스크롤 바 추가
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.window_list.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.window_list.configure(yscrollcommand=scrollbar.set)
        
        params_frame = ttk.Frame(arrange_frame)
        params_frame.pack(fill=tk.X, padx=5, pady=2)
        
        left_frame = ttk.Frame(params_frame)
        left_frame.pack(side=tk.LEFT, padx=(0, 5))
        right_frame = ttk.Frame(params_frame)
        right_frame.pack(side=tk.LEFT)
        
        ttk.Label(left_frame, text="시작 X 좌표").pack(anchor=tk.W)
        self.start_x = ttk.Entry(left_frame, width=8, style='Small.TEntry')
        self.start_x.pack(fill=tk.X, pady=(0, 2))
        self.start_x.insert(0, "0")
        
        ttk.Label(left_frame, text="창 너비").pack(anchor=tk.W)
        self.window_width = ttk.Entry(left_frame, width=8, style='Small.TEntry')
        self.window_width.pack(fill=tk.X, pady=(0, 2))
        self.window_width.insert(0, "500")
        
        ttk.Label(left_frame, text="가로 간격").pack(anchor=tk.W)
        self.h_spacing = ttk.Entry(left_frame, width=8, style='Small.TEntry')
        self.h_spacing.pack(fill=tk.X, pady=(0, 2))
        self.h_spacing.insert(0, "0")
        
        ttk.Label(right_frame, text="시작 Y 좌표").pack(anchor=tk.W)
        self.start_y = ttk.Entry(right_frame, width=8, style='Small.TEntry')
        self.start_y.pack(fill=tk.X, pady=(0, 2))
        self.start_y.insert(0, "0")
        
        ttk.Label(right_frame, text="창 높이").pack(anchor=tk.W)
        self.window_height = ttk.Entry(right_frame, width=8, style='Small.TEntry')
        self.window_height.pack(fill=tk.X, pady=(0, 2))
        self.window_height.insert(0, "400")
        
        ttk.Label(right_frame, text="세로 간격").pack(anchor=tk.W)
        self.v_spacing = ttk.Entry(right_frame, width=8, style='Small.TEntry')
        self.v_spacing.pack(fill=tk.X, pady=(0, 2))
        self.v_spacing.insert(0, "0")
        
        for widget in left_frame.winfo_children() + right_frame.winfo_children():
            if isinstance(widget, ttk.Entry):
                widget.pack_configure(pady=(0, 2))
        
        bottom_frame = ttk.Frame(arrange_frame)
        bottom_frame.pack(fill=tk.X, padx=5, pady=2)
        
        row_frame = ttk.Frame(bottom_frame)
        row_frame.pack(side=tk.LEFT)
        ttk.Label(row_frame, text="행당 창 개수").pack(anchor=tk.W)
        self.windows_per_row = ttk.Entry(row_frame, width=8, style='Small.TEntry')
        self.windows_per_row.pack(pady=(2, 0))
        self.windows_per_row.insert(0, "5")
        
        ttk.Button(bottom_frame, text="사용자 지정 배열", 
            command=self.custom_arrange_windows,
            style='Accent.TButton'
        ).pack(side=tk.RIGHT, pady=(15, 0))
        
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, padx=10, pady=(5, 0))
        
        self.tab_control = ttk.Notebook(bottom_frame)
        self.tab_control.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        open_window_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(open_window_tab, text="창 열기")
        
        input_frame = ttk.Frame(open_window_tab)
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(input_frame, text="단축키 디렉토리:").pack(side=tk.LEFT)
        self.path_entry = ttk.Entry(input_frame)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.path_entry.insert(0, self.shortcut_path)
        
        numbers_frame = ttk.Frame(input_frame)
        numbers_frame.pack(pady=5, padx=10, fill=tk.X)
        ttk.Label(numbers_frame, text="창 번호:").pack(side=tk.LEFT)
        self.numbers_entry = ttk.Entry(numbers_frame)
        self.numbers_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        settings = self.load_settings()
        if 'last_window_numbers' in settings:
            self.numbers_entry.insert(0, settings['last_window_numbers'])
            
        self.numbers_entry.bind('<Return>', lambda e: self.open_windows())
        
        ttk.Button(
            numbers_frame,
            text="창 열기",
            command=self.open_windows
        ).pack(side=tk.LEFT)
        
        url_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(url_tab, text="일괄 URL 열기")
        
        url_frame = ttk.Frame(url_tab)
        url_frame.pack(fill=tk.X, padx=5, pady=5)
        ttk.Label(url_frame, text="URL:").pack(side=tk.LEFT)
        self.url_entry = ttk.Entry(url_frame)
        self.url_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.url_entry.insert(0, "www.google.com")
        
        self.url_entry.bind('<Return>', lambda e: self.batch_open_urls())
        
        ttk.Button(url_frame, text="일괄 열기", command=self.batch_open_urls).pack(side=tk.LEFT, padx=5)
        
        icon_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(icon_tab, text="아이콘 변경")
        
        icon_frame = ttk.Frame(icon_tab)
        icon_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(icon_frame, text="아이콘 디렉토리:").pack(side=tk.LEFT)
        self.icon_path_entry = ttk.Entry(icon_frame)
        self.icon_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(icon_frame, text="창 번호:").pack(side=tk.LEFT, padx=(10, 0))
        self.icon_window_numbers = ttk.Entry(icon_frame, width=15)
        self.icon_window_numbers.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Label(icon_frame, text="예시: 1-5,7,9-12").pack(side=tk.LEFT)
        ttk.Button(icon_frame, text="아이콘 변경", command=self.set_taskbar_icons).pack(side=tk.LEFT, padx=5)
        
        footer_frame = ttk.Frame(self.root)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)

        author_frame = ttk.Frame(footer_frame)
        author_frame.pack(side=tk.RIGHT)

        ttk.Label(author_frame, text="Compiled by Devilflasher").pack(side=tk.LEFT)

        ttk.Label(author_frame, text="  ").pack(side=tk.LEFT)

        twitter_label = ttk.Label(
            author_frame, 
            text="Twitter",
            cursor="hand2",
            font=("Arial", 9)
        )
        twitter_label.pack(side=tk.LEFT)
        twitter_label.bind("<Button-1>", lambda e: webbrowser.open("https://x.com/DevilflasherX"))

        ttk.Label(author_frame, text="  ").pack(side=tk.LEFT)

        telegram_label = ttk.Label(
            author_frame, 
            text="Telegram",
            cursor="hand2",
            font=("Arial", 9)
        )
        telegram_label.pack(side=tk.LEFT)
        telegram_label.bind("<Button-1>", lambda e: webbrowser.open("https://t.me/devilflasher0"))

    def toggle_select_all(self, event=None):
        # 전체 선택 상태 전환
        try:
            items = self.window_list.get_children()
            if not items:
                return
                
            
            current_text = self.select_all_var.get()
            
            
            if current_text == "전체 선택":
                
                for item in items:
                    self.window_list.set(item, "select", "√")
            else:  
                
                for item in items:
                    self.window_list.set(item, "select", "")
            
            # 버튼 상태 업데이트
            self.update_select_all_status()
            
        except Exception as e:
            print(f"전체 선택 상태 실패: {str(e)}")

    def update_select_all_status(self):
        # 전체 선택 상태 업데이트
        try:
            # 모든 프로젝트 가져오기
            items = self.window_list.get_children()
            if not items:
                self.select_all_var.set("전체 선택")
                return
            
            # 모든 프로젝트가 선택되었는지 확인
            selected_count = sum(1 for item in items if self.window_list.set(item, "select") == "√")
            
            # 선택된 프로젝트 수에 따라 버튼 텍스트 설정
            if selected_count == len(items):
                self.select_all_var.set("취소 전체 선택")
            else:
                self.select_all_var.set("전체 선택")
            
        except Exception as e:
            print(f"전체 선택 상태 업데이트 실패: {str(e)}")

    def on_click(self, event):
        # 클릭 이벤트 처리
        try:
            region = self.window_list.identify_region(event.x, event.y)
            if region == "cell":
                column = self.window_list.identify_column(event.x)
                item = self.window_list.identify_row(event.y)
                
                if column == "#1":  # 선택 열
                    current = self.window_list.set(item, "select")
                    self.window_list.set(item, "select", "" if current == "√" else "√")
                    # 전체 선택 버튼 상태 업데이트
                    self.update_select_all_status()
                elif column == "#4":  # 마스터 열
                    self.set_master_window(item)
        except Exception as e:
            print(f"클릭 이벤트 처리 실패: {str(e)}")

    def set_master_window(self, item):
        # 마스터 창 설정
        if not item:
            return
        
        try:
            # 다른 창의 마스터 상태 및 제목 제거
            for i in self.window_list.get_children():
                values = self.window_list.item(i)['values']
                if values and len(values) >= 5:
                    hwnd = int(values[4])
                    title = values[2]
                    if title.startswith("[마스터]"):
                        new_title = title.replace("[마스터]", "").strip()
                        win32gui.SetWindowText(hwnd, new_title)
                    # 기본 테두리 색상 복원
                    try:
                        ctypes.windll.dwmapi.DwmSetWindowAttribute(
                            hwnd,
                            self.DWMWA_BORDER_COLOR,
                            ctypes.byref(ctypes.c_int(0)),
                            ctypes.sizeof(ctypes.c_int)
                        )
                    except:
                        pass
                self.window_list.set(i, "master", "")
                self.window_list.item(i, tags=())
            
            # 새로운 마스터 창 설정
            values = self.window_list.item(item)['values']
            self.master_window = int(values[4])
            
            # 마스터 표시 및 파란색 배경
            self.window_list.set(item, "master", "√")
            self.window_list.item(item, tags=("master",))
            
            # 창 제목 및 테두리 색상 수정
            title = values[2]
            if not title.startswith("[마스터]"):
                new_title = f"[마스터] {title}"
                win32gui.SetWindowText(self.master_window, new_title)
                try:
                    # 빨간색 테두리 설정
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        self.master_window,
                        self.DWMWA_BORDER_COLOR,
                        ctypes.byref(ctypes.c_int(0x000000FF)),
                        ctypes.sizeof(ctypes.c_int)
                    )
                except:
                    pass
            
        except Exception as e:
            print(f"마스터 창 설정 실패: {str(e)}")

    def toggle_sync(self, event=None):
        # 동기화 상태 전환
        if not self.window_list.get_children():
            messagebox.showinfo("알림", "먼저 창을 가져오세요!")
            return
        
        # 선택된 창 가져오기
        selected = []
        for item in self.window_list.get_children():
            if self.window_list.set(item, "select") == "√":
                selected.append(item)
        
        if not selected:
            messagebox.showinfo("알림", "동기화할 창을 선택하세요!")
            return
        
        # 마스터 창 확인
        master_items = [item for item in self.window_list.get_children() 
                       if self.window_list.set(item, "master") == "√"]
        
        if not master_items:
            # 마스터 창이 없으면 첫 번째 선택된 창을 마스터로 설정
            self.set_master_window(selected[0])
        
        # 동기화 상태 전환
        if not self.is_syncing:
            try:
                self.start_sync(selected)
                self.sync_button.configure(text="■ 중지 동기화", style='Accent.TButton')
                self.is_syncing = True
                print("동기화 시작")
            except Exception as e:
                print(f"동기화 시작 실패: {str(e)}")
                # 상태 보장
                self.is_syncing = False
                self.sync_button.configure(text="▶ 시작 동기화", style='Accent.TButton')
                # 다시 오류 메시지 표시
                messagebox.showerror("오류", str(e))
        else:
            try:
                self.stop_sync()
                self.sync_button.configure(text="▶ 시작 동기화", style='Accent.TButton')
                self.is_syncing = False
                print("동기화 중지")
            except Exception as e:
                print(f"동기화 중지 실패: {str(e)}")

    def start_sync(self, selected_items):
        try:
            # 마스터 창 존재 보장
            if not self.master_window:
                raise Exception("마스터 창 설정 안됨")
            
            # 선택된 창 목록 저장 및 번호 정렬
            self.sync_windows = []
            window_info = []
            
            # 모든 선택된 창 수집
            for item in selected_items:
                values = self.window_list.item(item)['values']
                if values and len(values) >= 5:
                    number = int(values[1])
                    hwnd = int(values[4])
                    if hwnd != self.master_window:  # 마스터 창 제외
                        window_info.append((number, hwnd))
            
            # 번호 정렬
            window_info.sort(key=lambda x: x[0])
            
            # 모든 동기화 창 핸들 저장
            self.sync_windows = [hwnd for _, hwnd in window_info]
            
            # 키보드 및 마우스 후킹 시작
            if not self.hook_thread:
                self.is_syncing = True
                self.hook_thread = threading.Thread(target=self.message_loop)
                self.hook_thread.daemon = True
                self.hook_thread.start()
                
                keyboard.hook(self.on_keyboard_event)
                mouse.hook(self.on_mouse_event)
                
                # 버튼 상태 업데이트
                self.sync_button.configure(text="■ 중지 동기화", style='Accent.TButton')
                
                # 플러그인 창 모니터 스레드 시작
                self.popup_monitor_thread = threading.Thread(target=self.monitor_popups)
                self.popup_monitor_thread.daemon = True
                self.popup_monitor_thread.start()
                
                print(f"동기화 시작, 마스터 창: {self.master_window}, 동기화 창: {self.sync_windows}")
                
        except Exception as e:
            self.stop_sync()  # 자원 정리 보장
            print(f"동기화 시작 실패: {str(e)}")
            raise e

    def message_loop(self):
        # 메시지 루프
        while self.is_syncing:
            time.sleep(0.001)

    def on_mouse_event(self, event):
        try:
            if self.is_syncing:
                current_window = win32gui.GetForegroundWindow()
                
                # 마스터 창 또는 플러그인 창인지 확인
                is_master = current_window == self.master_window
                master_popups = self.get_chrome_popups(self.master_window)
                is_popup = current_window in master_popups
                
                if is_master or is_popup:
                    # 이동 이벤트에 대한 최적화
                    if isinstance(event, mouse.MoveEvent):
                        # 이동 거리 및 시간 간격 확인
                        current_time = time.time()
                        if current_time - self.last_move_time < self.move_interval:
                            return
                            
                        dx = abs(event.x - self.last_mouse_position[0])
                        dy = abs(event.y - self.last_mouse_position[1])
                        if dx < self.mouse_threshold and dy < self.mouse_threshold:
                            return
                            
                        self.last_mouse_position = (event.x, event.y)
                        self.last_move_time = current_time

                    # 마우스 위치 가져오기
                    x, y = mouse.get_position()
                    
                    # 현재 창의 상대 좌표 가져오기
                    current_rect = win32gui.GetWindowRect(current_window)
                    rel_x = (x - current_rect[0]) / (current_rect[2] - current_rect[0])
                    rel_y = (y - current_rect[1]) / (current_rect[3] - current_rect[1])
                    
                    # 다른 창으로 동기화
                    for hwnd in self.sync_windows:
                        try:
                            # 목표 창 확인
                            if is_master:
                                target_hwnd = hwnd
                            else:
                                # 해당 확장 프로그램 창 찾기
                                target_popups = self.get_chrome_popups(hwnd)
                                # 상대 위치 일치
                                best_match = None
                                min_diff = float('inf')
                                for popup in target_popups:
                                    popup_rect = win32gui.GetWindowRect(popup)
                                    master_rect = win32gui.GetWindowRect(current_window)
                                    # 상대 위치 차이 계산
                                    master_rel_x = master_rect[0] - win32gui.GetWindowRect(self.master_window)[0]
                                    master_rel_y = master_rect[1] - win32gui.GetWindowRect(self.master_window)[1]
                                    popup_rel_x = popup_rect[0] - win32gui.GetWindowRect(hwnd)[0]
                                    popup_rel_y = popup_rect[1] - win32gui.GetWindowRect(hwnd)[1]
                                    
                                    diff = abs(master_rel_x - popup_rel_x) + abs(master_rel_y - popup_rel_y)
                                    if diff < min_diff:
                                        min_diff = diff
                                        best_match = popup
                                target_hwnd = best_match if best_match else hwnd
                            
                            if not target_hwnd:
                                continue
                            
                            # 목표 창 크기 가져오기
                            target_rect = win32gui.GetWindowRect(target_hwnd)
                            
                            # 목표 좌표 계산
                            client_x = int((target_rect[2] - target_rect[0]) * rel_x)
                            client_y = int((target_rect[3] - target_rect[1]) * rel_y)
                            lparam = win32api.MAKELONG(client_x, client_y)
                            
                            # 롤링 이벤트 처리
                            if isinstance(event, mouse.WheelEvent):
                                try:
                                    wheel_delta = int(event.delta)
                                    if keyboard.is_pressed('ctrl'):
                                        
                                        if wheel_delta > 0:                                            
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, 0xBB, 0)  # VK_OEM_PLUS
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, 0xBB, 0)
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
                                        else:
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, 0xBD, 0)  # VK_OEM_MINUS
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, 0xBD, 0)
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
                                    else:
                                        vk_code = win32con.VK_UP if wheel_delta > 0 else win32con.VK_DOWN
                                        vk_code = win32con.VK_UP if wheel_delta > 0 else win32con.VK_DOWN
                                        repeat_count = min(abs(wheel_delta) * 3, 6)
                                        for _ in range(repeat_count):
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                                            win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, vk_code, 0)
                                
                                except Exception as e:
                                    print(f"롤링 이벤트 처리 실패: {str(e)}")
                                    continue
                            
                            # 마우스 클릭 처리
                            elif isinstance(event, mouse.ButtonEvent):
                                if event.event_type == mouse.DOWN:
                                    if event.button == mouse.LEFT:
                                        win32gui.PostMessage(target_hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam)
                                    elif event.button == mouse.RIGHT:
                                        win32gui.PostMessage(target_hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, lparam)
                                elif event.event_type == mouse.UP:
                                    if event.button == mouse.LEFT:
                                        win32gui.PostMessage(target_hwnd, win32con.WM_LBUTTONUP, 0, lparam)
                                    elif event.button == mouse.RIGHT:
                                        win32gui.PostMessage(target_hwnd, win32con.WM_RBUTTONUP, 0, lparam)
                            
                            # 마우스 이동 처리
                            elif isinstance(event, mouse.MoveEvent):
                                win32gui.PostMessage(target_hwnd, win32con.WM_MOUSEMOVE, 0, lparam)
                                
                        except Exception as e:
                            print(f"창 {target_hwnd} 동기화 실패: {str(e)}")
                            continue
                            
        except Exception as e:
            print(f"마우스 이벤트 처리 실패: {str(e)}")

    def on_keyboard_event(self, event):
        # 개선된 키보드 이벤트 처리
        try:
            if self.is_syncing:
                current_window = win32gui.GetForegroundWindow()
                
                # 마스터 창 또는 플러그인 창인지 확인
                is_master = current_window == self.master_window
                master_popups = self.get_chrome_popups(self.master_window)
                is_popup = current_window in master_popups
                
                if is_master or is_popup:
                    # 실제 입력 목표 창 가져오기
                    input_hwnd = win32gui.GetFocus()
                    
                    # 다른 창으로 동기화
                    for hwnd in self.sync_windows:
                        try:
                            # 목표 창 확인
                            if is_master:
                                target_hwnd = hwnd
                            else:
                                # 해당 확장 프로그램 창 찾기
                                target_popups = self.get_chrome_popups(hwnd)
                                # 상대 위치 일치
                                best_match = None
                                min_diff = float('inf')
                                for popup in target_popups:
                                    popup_rect = win32gui.GetWindowRect(popup)
                                    master_rect = win32gui.GetWindowRect(current_window)
                                    # 상대 위치 차이 계산
                                    master_rel_x = master_rect[0] - win32gui.GetWindowRect(self.master_window)[0]
                                    master_rel_y = master_rect[1] - win32gui.GetWindowRect(self.master_window)[1]
                                    popup_rel_x = popup_rect[0] - win32gui.GetWindowRect(hwnd)[0]
                                    popup_rel_y = popup_rect[1] - win32gui.GetWindowRect(hwnd)[1]
                                    
                                    diff = abs(master_rel_x - popup_rel_x) + abs(master_rel_y - popup_rel_y)
                                    if diff < min_diff:
                                        min_diff = diff
                                        best_match = popup
                                target_hwnd = best_match if best_match else hwnd

                            if not target_hwnd:
                                continue

                            # Ctrl 조합키 처리
                            if keyboard.is_pressed('ctrl'):
                                # Ctrl 눌림 전송
                                win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
                                
                                # 일반적인 조합키 처리
                                if event.name in ['a', 'c', 'v', 'x']:
                                    vk_code = ord(event.name.upper())
                                    if event.event_type == keyboard.KEY_DOWN:
                                        win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                                        win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, vk_code, 0)
                                    win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
                                    continue
                                    
                            # 일반 키 처리
                            if event.name in ['enter', 'backspace', 'tab', 'esc', 'space', 
                                            'up', 'down', 'left', 'right',  # 좌우 키 추가
                                            'home', 'end', 'page up', 'page down', 'delete']:  
                                vk_map = {
                                    'enter': win32con.VK_RETURN,
                                    'backspace': win32con.VK_BACK,
                                    'tab': win32con.VK_TAB,
                                    'esc': win32con.VK_ESCAPE,
                                    'space': win32con.VK_SPACE,
                                    'up': win32con.VK_UP,
                                    'down': win32con.VK_DOWN,
                                    'left': win32con.VK_LEFT,      
                                    'right': win32con.VK_RIGHT,    
                                    'home': win32con.VK_HOME,
                                    'end': win32con.VK_END,
                                    'page up': win32con.VK_PRIOR,
                                    'page down': win32con.VK_NEXT,
                                    'delete': win32con.VK_DELETE  
                                }
                                vk_code = vk_map[event.name]
                            else:
                                # 일반 문자 처리
                                if len(event.name) == 1:
                                    vk_code = win32api.VkKeyScan(event.name[0]) & 0xFF
                                    if event.event_type == keyboard.KEY_DOWN:
                                        # 문자 메시지 전송
                                        win32gui.PostMessage(target_hwnd, win32con.WM_CHAR, ord(event.name[0]), 0)
                                    continue
                                else:
                                    continue

                            # 키 메시지 전송
                            if event.event_type == keyboard.KEY_DOWN:
                                win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, vk_code, 0)
                            else:
                                win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, vk_code, 0)
                                
                            # 조합키 해제
                            if keyboard.is_pressed('ctrl'):
                                win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
                                
                        except Exception as e:
                            print(f"창 {target_hwnd} 동기화 실패: {str(e)}")
                            
        except Exception as e:
            print(f"키보드 이벤트 처리 실패: {str(e)}")

    def stop_sync(self):
        # 동기화 중지
        try:
            self.is_syncing = False
            
            # 키보드 후킹 제거
            if self.keyboard_hook:
                keyboard.unhook(self.keyboard_hook)
                self.keyboard_hook = None
            
            # 마우스 후킹 제거
            if self.mouse_hook_id:
                mouse.unhook(self.mouse_hook_id)
                self.mouse_hook_id = None
            
            # 모니터 스레드 종료 대기
            if self.hook_thread and self.hook_thread.is_alive():
                self.hook_thread.join(timeout=1.0)
            
            # 자원 정리 (마스터 창 설정 유지)
            self.sync_windows.clear()
            self.popup_mappings.clear()
            
            # 디버그 포트 매핑 정리
            self.debug_ports.clear()
            
            # 마우스 상태 재설정
            self.last_mouse_position = (0, 0)
            self.last_move_time = 0
            
            # 버튼 상태 업데이트
            if self.sync_button:
                self.sync_button.configure(text="▶ 시작 동기화", style='Accent.TButton')
            
        except Exception as e:
            print(f"동기화 중지 실패: {str(e)}")

    def on_closing(self):
        # 창 닫기 이벤트
        try:
            self.stop_sync()
            # 단축키 정리
            if self.shortcut_hook:
                keyboard.clear_all_hotkeys()
                keyboard.unhook_all()
                self.shortcut_hook = None
            self.save_settings()
        except Exception as e:
            print(f"프로그램 닫을 때 오류: {str(e)}")
        finally:
            self.root.destroy()

    def auto_arrange_windows(self):
        # 자동 배열 창
        try:
            # 먼저 동기화 중지
            was_syncing = self.is_syncing
            if was_syncing:
                self.stop_sync()
            
            # 선택된 창 및 번호 정렬
            selected = []
            for item in self.window_list.get_children():
                if self.window_list.set(item, "select") == "√":
                    values = self.window_list.item(item)['values']
                    if values and len(values) >= 5:
                        number = int(values[1])  
                        hwnd = int(values[4])
                        selected.append((number, hwnd, item))
            
            if not selected:
                messagebox.showinfo("알림", "먼저 배열할 창을 선택하세요!")
                return
            
            # 번호 순서대로 정렬
            selected.sort(key=lambda x: x[0])  
            
            # 화면 크기 가져오기
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
            
            # 최적 레이아웃 계산
            count = len(selected)
            cols = int(math.sqrt(count))
            if cols * cols < count:
                cols += 1
            rows = (count + cols - 1) // cols
            
            # 창 크기 계산
            width = screen_width // cols
            height = screen_height // rows
            
            # 위치 매핑(좌측에서 우측으로, 상단에서 하단으로) 생성
            positions = []
            # 먼저 완전한 위치 목록 생성
            for i in range(count):
                row = i // cols
                col = i % cols
                x = col * width
                y = row * height
                positions.append((x, y))
            
            # 창 위치 적용
            for i, (_, hwnd, _) in enumerate(selected):
                x, y = positions[i]
                # 창 표시 및 지정된 위치로 이동
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.MoveWindow(hwnd, x, y, width, height, True)
            
            # 이전에 동기화 중이면 다시 동기화 시작
            if was_syncing:
                self.start_sync([item for _, _, item in selected])
            
        except Exception as e:
            messagebox.showerror("오류", f"자동 배열 실패: {str(e)}")

    def custom_arrange_windows(self):
        # 사용자 지정 배열 창
        try:
            # 먼저 동기화 중지
            was_syncing = self.is_syncing
            if was_syncing:
                self.stop_sync()
            
            selected = []
            for item in self.window_list.get_children():
                if self.window_list.set(item, "select") == "√":
                    selected.append(item)
                    
            if not selected:
                messagebox.showinfo("알림", "배열할 창을 선택하세요!")
                return
            
            try:
                # 매개변수 가져오기
                start_x = int(self.start_x.get())
                start_y = int(self.start_y.get())
                width = int(self.window_width.get())
                height = int(self.window_height.get())
                h_spacing = int(self.h_spacing.get())
                v_spacing = int(self.v_spacing.get())
                windows_per_row = int(self.windows_per_row.get())
                
                # 창 배열
                for i, item in enumerate(selected):
                    values = self.window_list.item(item)['values']
                    if values and len(values) >= 5:  # 충분한 값 확인
                        hwnd = int(values[4])
                        row = i // windows_per_row
                        col = i % windows_per_row
                        
                        x = start_x + col * (width + h_spacing)
                        y = start_y + row * (height + v_spacing)
                        
                        # 창 표시 및 지정된 위치로 이동
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        win32gui.MoveWindow(hwnd, x, y, width, height, True)
                
                # 매개변수 저장
                self.save_settings()
                    
            except ValueError:
                messagebox.showerror("오류", "유효한 숫자 매개변수를 입력하세요!")
            except Exception as e:
                messagebox.showerror("오류", f"창 배열 실패: {str(e)}")
            
            # 이전에 동기화 중이면 다시 동기화 시작
            if was_syncing:
                self.start_sync(selected)
            
        except Exception as e:
            messagebox.showerror("오류", f"창 배열 실패: {str(e)}")

    def load_settings(self) -> dict:
        # 설정 로드
        try:
            with open('settings.json', 'r') as f:
                return json.load(f)
        except:
            return {}

    def save_settings(self):
        # 설정 저장
        try:
            settings = {
                'shortcut_path': self.path_entry.get(),
                'window_position': self.root.geometry(),
                'last_window_numbers': self.numbers_entry.get(),  # 창 번호 저장
                'arrange_params': {
                    'start_x': self.start_x.get(),
                    'start_y': self.start_y.get(),
                    'window_width': self.window_width.get(),
                    'window_height': self.window_height.get(),
                    'h_spacing': self.h_spacing.get(),
                    'v_spacing': self.v_spacing.get(),
                    'windows_per_row': self.windows_per_row.get()
                },
                'sync_shortcut': self.current_shortcut
            }
            with open('settings.json', 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"설정 저장 실패: {str(e)}")

    def load_arrange_params(self):
        # 배열 매개변수 로드
        settings = self.load_settings()
        if 'arrange_params' in settings:
            params = settings['arrange_params']
            self.start_x.delete(0, tk.END)
            self.start_x.insert(0, params.get('start_x', '0'))
            self.start_y.delete(0, tk.END)
            self.start_y.insert(0, params.get('start_y', '0'))
            self.window_width.delete(0, tk.END)
            self.window_width.insert(0, params.get('window_width', '500'))
            self.window_height.delete(0, tk.END)
            self.window_height.insert(0, params.get('window_height', '400'))
            self.h_spacing.delete(0, tk.END)
            self.h_spacing.insert(0, params.get('h_spacing', '0'))
            self.v_spacing.delete(0, tk.END)
            self.v_spacing.insert(0, params.get('v_spacing', '0'))
            self.windows_per_row.delete(0, tk.END)
            self.windows_per_row.insert(0, params.get('windows_per_row', '5'))

    def parse_window_numbers(self, numbers_str: str) -> List[int]:
        # 창 번호 문자열 분석
        if not numbers_str.strip():
            return list(range(1, 49))  # 빈 경우 모든 번호 반환
            
        result = []
        # 쉼표로 구분된 부분 분석
        parts = numbers_str.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                # 범위 처리, 예: "1-5"
                start, end = map(int, part.split('-'))
                result.extend(range(start, end + 1))
            else:
                # 단일 숫자 처리
                result.append(int(part))
        return sorted(list(set(result)))  # 중복 제거 및 정렬

    def open_windows(self):
        # Chrome 창 열기
        path = self.path_entry.get()
        numbers = self.numbers_entry.get()
        
        if not path or not numbers:
            messagebox.showwarning("경고", "단축키 경로 및 창 번호를 입력하세요!")
            return
            
        try:
            window_numbers = self.parse_window_numbers(numbers)
            for num in window_numbers:
                shortcut = os.path.join(path, f"{num}.lnk")
                if os.path.exists(shortcut):
                    subprocess.Popen(["start", "", shortcut], shell=True)
                    time.sleep(0.5)  # 지연 보장 순서로 열기
                else:
                    messagebox.showwarning("경고", f"단축키 없음: {shortcut}")
            
            # 경로 저장
            self.save_settings()
            
        except Exception as e:
            messagebox.showerror("오류", f"창 열기 실패: {str(e)}")

    def get_shortcut_number(self, shortcut_path):
        # 단축키에서 창 번호 가져오기
        handle = None
        try:
            # 목표 프로세스 가져오기
            shortcut = self.shell.CreateShortCut(shortcut_path)
            cmd_line = shortcut.Arguments
            
            # 프로세스 열기 명령행 가져오기
            handle = win32api.OpenProcess(
                win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ,
                False,
                self.pid
            )
            
            # 명령행 처리
            if '--user-data-dir=' in cmd_line:
                data_dir = cmd_line.split('--user-data-dir=')[1].strip('"')
                number = os.path.basename(data_dir)
                return number
                
            return None
            
        except Exception as e:
            print(f"단축키 번호 가져오기 실패: {str(e)}")
            return None
            
        finally:
            # 핸들 닫기 보장
            if handle:
                try:
                    win32api.CloseHandle(handle)
                except Exception as e:
                    print(f"프로세스 핸들 닫기 실패: {str(e)}")

    def import_windows(self):
        # 현재 열린 Chrome 창 가져오기
        try:
            # 목록 비우기
            for item in self.window_list.get_children():
                self.window_list.delete(item)
            
            def callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    try:
                        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                        path = win32process.GetModuleFileNameEx(handle, 0)
                        win32api.CloseHandle(handle)
                        
                        if "chrome.exe" in path.lower():
                            title = win32gui.GetWindowText(hwnd)
                            if title and not title.startswith("Chrome 전달"):
                                # 각 창에 디버그 포트 할당
                                port = self.base_debug_port + len(windows)
                                self.debug_ports[hwnd] = port
                                windows.append((title, hwnd))
                            
                    except Exception as e:
                        print(f"프로세스 정보 가져오기 실패: {str(e)}")
                    
            windows = []
            win32gui.EnumWindows(callback, windows)
            
            # windows 목록 반전, 이렇게 하면 역순으로 목록에 추가됨
            windows.reverse()
            
            # 목록에 추가
            for i, (title, hwnd) in enumerate(windows, 1):
                self.window_list.insert("", "end", values=("", f"{i}", title, "", hwnd))
            
        except Exception as e:
            messagebox.showerror("오류", f"창 가져오기 실패: {str(e)}")

    def enum_window_callback(self, hwnd, windows):
        # 창 열거 콜백 함수
        try:
            # 창이 표시되는지 확인
            if not win32gui.IsWindowVisible(hwnd):
                return
            
            # 창 제목 가져오기
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return
            
            # Chrome 창인지 확인
            if " - Google Chrome" in title:
                # 창 번호 추출
                number = None
                if title.startswith("[마스터]"):
                    title = title[4:].strip()  # [마스터] 표시 제거
                
                # 프로세스 명령행 인수에서 창 번호 가져오기
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, False, pid)
                    if handle:
                        cmd_line = win32process.GetModuleFileNameEx(handle, 0)
                        win32api.CloseHandle(handle)
                        
                        # 경로에서 번호 추출
                        if "\\Data\\" in cmd_line:
                            number = int(cmd_line.split("\\Data\\")[-1].split("\\")[0])
                except:
                    pass
                
                if number is not None:
                    windows.append({
                        'hwnd': hwnd,
                        'title': title,
                        'number': number
                    })
                
        except Exception as e:
            print(f"창 열거 실패: {str(e)}")

    def close_selected_windows(self):
        # 선택된 창 닫기
        selected = []
        for item in self.window_list.get_children():
            if self.window_list.set(item, "select") == "√":
                selected.append(item)
                
        if not selected:
            messagebox.showinfo("알림", "먼저 닫을 창을 선택하세요!")
            return
            
        try:
            for item in selected:
                # values에서 hwnd 가져오기
                hwnd = int(self.window_list.item(item)['values'][4])
                try:
                    # 창이 아직 있는지 확인
                    if win32gui.IsWindow(hwnd):
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                except:
                    pass  # 이미 닫힌 창 오류 무시
            
            # 창 닫힌 후 목록 새로고침
            time.sleep(0.5)
            self.import_windows()
            
        except Exception as e:
            print(f"창 닫기 실패: {str(e)}")  # 오류만 인쇄, 오류 대화상자 표시 안함

    def set_taskbar_icons(self):
        # 독립 작업 표시줄 아이콘 설정
        if not self.path_entry.get():
            messagebox.showinfo("알림", "먼저 단축키 경로를 설정하세요!")
            return
            
        if not os.path.exists(self.path_entry.get()):
            messagebox.showerror("오류", "단축키 경로 없음!")
            return
            
        # 작업 선택
        choice = messagebox.askyesnocancel("작업 선택", "수행할 작업 선택:\n예 - 사용자 지정 아이콘 설정\n아니오 - 원래 설정 복원\n취소 - 아무 작업도 수행하지 않음")
        if choice is None:  # 사용자 취소
            return
            
        try:
            data_dir = self.path_entry.get()
            icon_dir = self.icon_path_entry.get()
            shell = win32com.client.Dispatch("WScript.Shell")
            modified_count = 0
            
            # 수정할 창 번호 목록 가져오기
            window_numbers = self.parse_window_numbers(self.icon_window_numbers.get())
            
            if choice:  # 사용자 지정 아이콘 설정
                # 아이콘 디렉토리 존재 보장
                if not os.path.exists(icon_dir):
                    os.makedirs(icon_dir)
                
                # 지정된 단축키 수정
                for i in window_numbers:
                    shortcut_path = os.path.join(data_dir, f"{i}.lnk")
                    if not os.path.exists(shortcut_path):
                        continue
                        
                    # 단축키 수정
                    shortcut = shell.CreateShortCut(shortcut_path)
                    
                    # 사용자 지정 아이콘 설정
                    icon_path = os.path.join(icon_dir, f"{i}.ico")
                    if os.path.exists(icon_path):
                        shortcut.IconLocation = icon_path
                        # 수정 저장
                        shortcut.save()
                        modified_count += 1
                
                messagebox.showinfo("성공", f"성공적으로 {modified_count} 개 단축키의 아이콘을 수정했습니다!")
            else:  # 원래 설정 복원
                chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
                if not os.path.exists(chrome_path):
                    chrome_path = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
                
                # 지정된 단축키 복원
                for i in window_numbers:
                    shortcut_path = os.path.join(data_dir, f"{i}.lnk")
                    if not os.path.exists(shortcut_path):
                        continue
                        
                    # 단축키 수정
                    shortcut = shell.CreateShortCut(shortcut_path)
                    
                    # 기본 아이콘 복원
                    shortcut.IconLocation = f"{chrome_path},0"
                    
                    # 원래 시작 인수 복원
                    original_args = f'--user-data-dir="C:\\python_project\\Chrome-Manager\\Data\\{i}"'
                    shortcut.TargetPath = chrome_path
                    shortcut.Arguments = original_args
                    
                    # 수정 저장
                    shortcut.save()
                    modified_count += 1
                
                messagebox.showinfo("성공", f"성공적으로 {modified_count} 개 단축키의 원래 설정을 복원했습니다!")
            
        except Exception as e:
            messagebox.showerror("오류", f"작업 실패: {str(e)}")

    def batch_open_urls(self):
        # 일괄 URL 열기
        try:
            # 입력된 URL 가져오기
            url = self.url_entry.get() 
            if not url:
                messagebox.showwarning("경고", "열려는 URL을 입력하세요!")
                return
            
            # URL 형식 확인
            if not url.startswith(('http://', 'https://')):
                url = 'https://' + url
            
            # 선택된 창 가져오기
            selected_windows = []
            for item in self.window_list.get_children():
                if self.window_list.set(item, "select") == "√":
                    hwnd = int(self.window_list.item(item)['values'][-1])
                    selected_windows.append(hwnd)
            
            if not selected_windows:
                messagebox.showwarning("경고", "먼저 작업할 창을 선택하세요!")
                return
            
            # 선택된 각 창에서 URL 열기
            for hwnd in selected_windows:
                try:
                    # 창 활성화
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.1) 
                    
                    # 새 탭 페이지 열기
                    keyboard.press_and_release('ctrl+t')
                    time.sleep(0.1) 
                    
                    # URL 입력 및 엔터
                    keyboard.write(url)
                    time.sleep(0.1) 
                    keyboard.press_and_release('enter')
                    time.sleep(0.2) 
                    
                except Exception as e:
                    print(f"창 {hwnd}에서 URL 열기 실패: {str(e)}")
            
            messagebox.showinfo("성공", "일괄 URL 열기 완료!")
            
        except Exception as e:
            messagebox.showerror("오류", f"일괄 URL 열기 실패: {str(e)}")

    def run(self):
        """프로그램 실행"""
        self.root.mainloop()

    def load_window_position(self):
        # settings.json에서 창 위치 로드
        try:
            settings = self.load_settings()
            return settings.get('window_position')
        except:
            return None

    def save_window_position(self):
        # settings.json에 창 위치 저장
        try:
            position = self.root.geometry()
            settings = self.load_settings()
            settings['window_position'] = position
            
            with open('settings.json', 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"창 위치 저장 실패: {str(e)}")

    def get_chrome_popups(self, chrome_hwnd):
        # 개선된 플러그인 창 탐지
        popups = []
        def enum_windows_callback(hwnd, _):
            try:
                if not win32gui.IsWindowVisible(hwnd):
                    return
                    
                class_name = win32gui.GetClassName(hwnd)
                title = win32gui.GetWindowText(hwnd)
                _, chrome_pid = win32process.GetWindowThreadProcessId(chrome_hwnd)
                _, popup_pid = win32process.GetWindowThreadProcessId(hwnd)
                
                # Chrome 관련 창인지 확인
                if popup_pid == chrome_pid:
                    # 창 유형 확인
                    if "Chrome_WidgetWin_1" in class_name:
                        # 확장 프로그램 관련 창인지 확인, 감도 조건 확인
                        style = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
                        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                        
                        # 확장 창의 특성
                        is_popup = (
                            "확장 프로그램" in title or 
                            "플러그인" in title or
                            win32gui.GetParent(hwnd) == chrome_hwnd or
                            (style & win32con.WS_POPUP) != 0 or
                            (style & win32con.WS_CHILD) != 0 or
                            (ex_style & win32con.WS_EX_TOOLWINDOW) != 0  # 도구 창 확인 추가
                        )
                        
                        if is_popup:
                            popups.append(hwnd)
                    
            except Exception as e:
                print(f"창 열거 실패: {str(e)}")
                
        win32gui.EnumWindows(enum_windows_callback, None)
        return popups

    def monitor_popups(self):
        # 플러그인 창 변경 모니터링
        while self.is_syncing:
            try:
                self.sync_chrome_popups()
            except:
                pass
            time.sleep(0.1)  

    def show_shortcut_dialog(self):
        # 단축키 설정 대화상자 표시
        # 대화상자 생성
        dialog = tk.Toplevel(self.root)
        dialog.title("동기화 기능 단축키 설정")
        dialog.geometry("300x150")
        dialog.resizable(False, False)
        
        # 대화상자 모달 설정
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 현재 단축키 표시
        current_label = ttk.Label(dialog, text=f"현재 단축키: {self.current_shortcut}")
        current_label.pack(pady=10)
        
        # 단축키 입력 상자
        shortcut_var = tk.StringVar(value="하단 버튼을 클릭하여 단축키 녹음 시작...")
        shortcut_label = ttk.Label(dialog, textvariable=shortcut_var)
        shortcut_label.pack(pady=5)
        
        # 키 상태 기록
        keys_pressed = set()
        recording = False
        
        def start_recording():
            # 단축키 녹음 시작
            nonlocal recording
            recording = True
            keys_pressed.clear()
            shortcut_var.set("키 조합을 눌러 녹음...")
            record_btn.configure(state='disabled')
            
            def on_key_event(e):
                if not recording:
                    return
                if e.event_type == keyboard.KEY_DOWN:
                    keys_pressed.add(e.name)
                    shortcut_var.set('+'.join(sorted(keys_pressed)))
                elif e.event_type == keyboard.KEY_UP:
                    if e.name in keys_pressed:
                        keys_pressed.remove(e.name)
                    if not keys_pressed:  
                        stop_recording()
            
            keyboard.hook(on_key_event)
        
        def stop_recording():
            # 단축키 녹음 중지
            nonlocal recording
            recording = False
            keyboard.unhook_all()
            # 다시 현재 단축키 설정
            if self.current_shortcut:
                self.set_shortcut(self.current_shortcut)
            record_btn.configure(state='normal')
        
        # 녹음 버튼
        record_btn = ttk.Button(
            dialog,
            text="녹음 시작",
            command=start_recording
        )
        record_btn.pack(pady=10)
        
        def save_shortcut():
            # 단축키 설정 저장
            new_shortcut = shortcut_var.get()
            if new_shortcut and new_shortcut != "하단 버튼을 클릭하여 단축키 녹음 시작..." and new_shortcut != "키 조합을 눌러 녹음...":
                try:
                    # 새 단축키 설정
                    self.set_shortcut(new_shortcut)
                    
                    # 설정 파일에 저장
                    settings = self.load_settings()
                    settings['sync_shortcut'] = new_shortcut
                    with open('settings.json', 'w', encoding='utf-8') as f:
                        json.dump(settings, f, ensure_ascii=False, indent=4)
                    
                    messagebox.showinfo("성공", f"단축키가 {new_shortcut}로 설정되었습니다!")
                    dialog.destroy()
                except Exception as e:
                    messagebox.showerror("오류", f"단축키 설정 실패: {str(e)}")
            else:
                messagebox.showwarning("경고", "먼저 단축키를 녹음하세요!")
        
        # 저장 버튼
        ttk.Button(
            dialog,
            text="저장",
            command=save_shortcut
        ).pack(pady=5)
        
        # 대화상자 닫을 때 녹음 중지
        dialog.protocol("WM_DELETE_WINDOW", lambda: [stop_recording(), dialog.destroy()])
        
        # 대화상자 중앙 표시
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f'{width}x{height}+{x}+{y}')

    def set_shortcut(self, shortcut):
        # 단축키 설정
        try:
            # 먼저 모든 단축키 및 후킹 제거
            keyboard.unhook_all()
            keyboard.clear_all_hotkeys()
            
            if self.shortcut_hook:
                self.shortcut_hook = None
            
            if shortcut:
                # 오류 재시도 메커니즘 추가
                max_retries = 3
                for attempt in range(max_retries):
                    try:
                        # 새 단축키 등록
                        self.shortcut_hook = keyboard.add_hotkey(
                            shortcut,
                            self.toggle_sync,
                            suppress=True,
                            trigger_on_release=True  # 키 해제 시 트리거, 카드 방지
                        )
                        self.current_shortcut = shortcut
                        print(f"단축키 {shortcut} 설정 성공")
                        break
                    except Exception as e:
                        print(f"시도 {attempt + 1} 단축키 설정 실패: {str(e)}")
                        if attempt == max_retries - 1:
                            raise
                        time.sleep(0.5)  # 재시도 전 대기
                    
        except Exception as e:
            print(f"단축키 설정 실패: {str(e)}")
            self.shortcut_hook = None
            # 이전 단축키 복원 시도
            if self.current_shortcut and self.current_shortcut != shortcut:
                try:
                    self.shortcut_hook = keyboard.add_hotkey(
                        self.current_shortcut,
                        self.toggle_sync,
                        suppress=True,
                        trigger_on_release=True
                    )
                except:
                    self.current_shortcut = None

if __name__ == "__main__":
    app = ChromeManager()
    app.run() 