import os
import sys
from datetime import datetime
from tkinter import Tk, Button, Label, Text, Scrollbar, filedialog, Frame, END, StringVar, Entry, Toplevel
from tkcalendar import Calendar as TkCalendarWidget

from calendar_processor import CalendarFormatter
from config import REGION_KEYWORD_MAP, SEMINAR_KEYWORDS

# --- .exe 파일 경로 설정 함수 ---
def get_base_path():
    """
    실행 파일(.exe)로 만들었을 때와 파이썬 스크립트로 실행했을 때
    모두 올바른 경로를 찾도록 도와주는 함수입니다.
    """
    if getattr(sys, 'frozen', False):
        # .exe 파일로 실행될 경우, .exe 파일이 있는 폴더를 기준으로 경로를 설정합니다.
        return os.path.dirname(sys.executable)
    else:
        # 파이썬 스크립트로 실행될 경우, 이 파일이 있는 폴더를 기준으로 경로를 설정합니다.
        return os.path.dirname(os.path.abspath(__file__))

BASE_PATH = get_base_path()

class TextRedirector:
    """STDOUT, STDERR 출력을 Tkinter Text 위젯으로 리디렉션하는 클래스"""

    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, string):
        self.widget.configure(state='normal')
        self.widget.insert(END, string, (self.tag,))
        self.widget.see(END)
        self.widget.configure(state='disabled')

    def flush(self):
        pass



class App:
    """GUI 애플리케이션 클래스"""

    def __init__(self, master):
        self.master = master
        master.title("투쟁일정 캘린더 텍스트 생성기 v1.1")
        master.geometry("800x750")

        self.selected_folder_path = StringVar()
        self.selected_folder_path.set("폴더를 선택해주세요.")

        # --- 상단 컨트롤 프레임 ---
        top_frame = Frame(master)
        top_frame.pack(pady=10, padx=10, fill='x')

        self.select_button = Button(top_frame, text="ICS 폴더 선택", command=self.select_folder, height=2)
        self.select_button.pack(side='left', padx=(0, 5))

        path_label_frame = Frame(top_frame, relief='sunken', borderwidth=1)
        path_label_frame.pack(side='left', fill='x', expand=True, ipady=4)
        self.path_label = Label(path_label_frame, textvariable=self.selected_folder_path, anchor='w', bg='white')
        self.path_label.pack(fill='x', padx=5)

        # --- 날짜 입력 프레임 ---
        date_frame = Frame(master)
        date_frame.pack(pady=5, padx=10, fill='x')

        start_date_label = Label(date_frame, text="시작 날짜 (YYYY-MM-DD):")
        start_date_label.pack(side='left', padx=(0, 5))
        self.start_date_entry = Entry(date_frame, width=15)
        self.start_date_entry.pack(side='left', padx=5)
        start_cal_button = Button(date_frame, text="🗓️", command=lambda: self._open_calendar(self.start_date_entry))
        start_cal_button.pack(side='left')

        end_date_label = Label(date_frame, text="종료 날짜 (YYYY-MM-DD):")
        end_date_label.pack(side='left', padx=(10, 5))
        self.end_date_entry = Entry(date_frame, width=15)
        self.end_date_entry.pack(side='left', padx=5)
        end_cal_button = Button(date_frame, text="🗓️", command=lambda: self._open_calendar(self.end_date_entry))
        end_cal_button.pack(side='left')

        date_info_label = Label(date_frame, text="*날짜를 비우면 오늘부터 7일간의 일정을 생성합니다.")
        date_info_label.pack(side='right', padx=10)

        self.run_button = Button(master, text="일정 생성 시작", command=self.run_processing, state='disabled', height=2,
                                 font=('Helvetica', 10, 'bold'))
        self.run_button.pack(pady=(0, 10), padx=10, fill='x')

        # --- 결과 텍스트 창 ---
        result_frame = Frame(master, pady=5)
        result_frame.pack(fill='both', expand=True, padx=10)

        result_label = Label(result_frame, text="📋 생성된 텍스트 (복사하여 사용)")
        result_label.pack(anchor='w')

        result_text_frame = Frame(result_frame)
        result_text_frame.pack(fill='both', expand=True)

        self.result_text = Text(result_text_frame, wrap='word', height=15)
        self.result_text.pack(side='left', fill='both', expand=True)

        result_scrollbar = Scrollbar(result_text_frame, command=self.result_text.yview)
        result_scrollbar.pack(side='right', fill='y')
        self.result_text.config(yscrollcommand=result_scrollbar.set)

        # --- 로그 텍스트 창 ---
        log_frame = Frame(master, pady=5)
        log_frame.pack(fill='both', expand=True, padx=10)

        log_label = Label(log_frame, text="⚙️ 처리 과정 로그")
        log_label.pack(anchor='w')

        log_text_frame = Frame(log_frame)
        log_text_frame.pack(fill='both', expand=True)

        self.log_text = Text(log_text_frame, wrap='word', state='disabled', height=10)
        self.log_text.pack(side='left', fill='both', expand=True)

        log_scrollbar = Scrollbar(log_text_frame, command=self.log_text.yview)
        log_scrollbar.pack(side='right', fill='y')
        self.log_text.config(yscrollcommand=log_scrollbar.set)

        # stdout, stderr 출력을 로그 창으로 리디렉션
        sys.stdout = TextRedirector(self.log_text, "stdout")
        sys.stderr = TextRedirector(self.log_text, "stderr")

        print("안녕하세요! 'ICS 폴더 선택' 버튼을 눌러 .ics 파일이 담긴 폴더를 지정해주세요.")

    def _open_calendar(self, entry_widget):
        """달력 위젯을 새 창에 띄우고 선택된 날짜를 Entry에 입력하는 함수"""

        def set_date():
            # Calendar 위젯에서 yyyy-mm-dd 형식의 문자열로 날짜 호출
            selected_date = cal.get_date()
            entry_widget.delete(0, END)
            entry_widget.insert(0, selected_date)
            top.destroy()

        top = Toplevel(self.master)
        top.title("날짜 선택")
        top.grab_set()  # 다른 창과 상호작용하지 못하도록 설정

        # 현재 Entry 위젯에 있는 날짜를 파싱하여 달력의 초기 날짜로 설정
        try:
            initial_date = datetime.strptime(entry_widget.get(), '%Y-%m-%d')
            cal = TkCalendarWidget(top, selectmode='day', year=initial_date.year, month=initial_date.month,
                                   day=initial_date.day,
                                   date_pattern='yyyy-mm-dd', locale='ko_KR')
        except ValueError:
            # Entry가 비어있거나 형식이 잘못된 경우 오늘 날짜로 설정
            cal = TkCalendarWidget(top, selectmode='day', date_pattern='yyyy-mm-dd', locale='ko_KR')

        cal.pack(pady=10, padx=10)

        select_button = Button(top, text="선택", command=set_date)
        select_button.pack(pady=10)

    def select_folder(self):
        default_path = os.path.join(BASE_PATH, 'data')
        if not os.path.exists(default_path):
            default_path = BASE_PATH

        folder_path = filedialog.askdirectory(initialdir=default_path, title="ICS 파일이 있는 폴더를 선택하세요")
        if folder_path:
            self.selected_folder_path.set(folder_path)
            self.run_button.config(state='normal')

            # 이전 내용 초기화
            self.result_text.delete(1.0, END)
            self.log_text.configure(state='normal')
            self.log_text.delete(1.0, END)
            self.log_text.configure(state='disabled')

            print(f"선택된 폴더: {folder_path}\n날짜를 입력하거나 비워둔 채로 '일정 생성 시작' 버튼을 눌러주세요.")

    def run_processing(self):
        folder_path = self.selected_folder_path.get()
        if not os.path.isdir(folder_path):
            print("오류: 유효한 폴더가 선택되지 않았습니다.", "stderr")
            return

        # 이전 내용 초기화
        self.result_text.delete(1.0, END)
        self.log_text.configure(state='normal')
        self.log_text.delete(1.0, END)
        self.log_text.configure(state='disabled')

        self.run_button.config(state='disabled')
        self.select_button.config(state='disabled')

        start_date = self.start_date_entry.get()
        end_date = self.end_date_entry.get()

        try:
            formatter = CalendarFormatter(folder_path, REGION_KEYWORD_MAP, SEMINAR_KEYWORDS)
            final_text = formatter.run(start_date, end_date)

            if final_text:
                self.result_text.insert(END, final_text)

        except Exception as e:
            print(f"GUI 처리 중 예외 발생: {e}")
        finally:
            self.run_button.config(state='normal')
            self.select_button.config(state='normal')

