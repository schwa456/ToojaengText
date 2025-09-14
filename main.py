import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta

from dateutil.rrule import rrulestr
from icalendar import Calendar, vRecur
import pytz

from tkinter import Tk, Button, Label, Text, Scrollbar, filedialog, Frame, END, StringVar, Entry, Toplevel
from tkcalendar import Calendar as TkCalendarWidget


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


# --- 설정 (Configuration) ---
# 1. .ics 파일들이 모여 있는 폴더 경로 (GUI를 통해 선택되므로 직접 수정할 필요 없음)
BASE_PATH = get_base_path()

# 2. 지역 그룹과 매칭할 키워드 목록 (출력 순서대로 작성)
REGION_KEYWORD_MAP = {
    '강원': ['강원', '강릉', '동해', '삼척', '속초', '원주', '춘천', '태백', '고성', '양구', '양양', '영월', '인제', '정선', '철원', '평창', '홍천', '화천', '횡성'],
    '경상/대구/울산/부산': ['경상', '대구', '울산', '부산', '경산', '경주', '구미', '김천', '문경', '상주', '안동', '영주', '영천', '포항', '고령', '봉화',
                    '성주', '영덕', '영양', '예천', '울릉', '울진', '의성', '청도', '청송', '칠곡', '창원', '거제', '김해', '밀양', '사천', '양산',
                    '진주', '통영', '거창', '고성', '남해', '산청', '의령', '창녕', '하동', '함안', '함양', '합천'],
    '전라/광주': ['전라', '광주', '군산', '김제', '남원', '익산', '전주', '정읍', '고창', '무주', '부안', '순창', '완주', '임실', '장수', '진안', '목포',
              '여수', '순천', '나주', '광양', '담양', '곡성', '구례', '고흥', '보성', '화순', '장흥', '강진', '해남', '영암', '무안', '함평', '영광',
              '장성', '완도', '진도', '신안'],
    '제주': ['제주', '서귀포'],
    '충청/대전/세종': ['충청', '대전', '세종', '제천', '청주', '충주', '괴산', '단양', '보은', '영동', '옥천', '음성', '증평', '진천', '계룡', '공주', '논산',
                 '당진', '보령', '서산', '아산', '천안', '금산', '부여', '서천', '예산', '청양', '태안', '홍성'],
    '경기/인천': ['경기', '인천', '수원', '고양', '용인', '화성', '성남', '의정부', '안양', '부천', '광명', '평택', '동두천', '안산', '과천', '구리', '남양주',
              '오산', '시흥', '군포', '의왕', '하남', '파주', '이천', '안성', '김포', '광주', '양주', '포천', '여주', '연천', '가평', '양평'],
    '서울': ['서울'],
    '온라인': ['온라인'],
}

# 3. 세미나/강연으로 분류할 키워드
SEMINAR_KEYWORDS = ['세미나', '아카데미', '공부모임', '영화제', '글쓰기', '북토크', '강연', '전시', '플리마켓', '기획전시', '상영회', '토론회', '강좌']


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


class CalendarFormatter:
    def __init__(self, directory_path, region_keyword_map, seminar_keywords):
        self.directory_path = directory_path
        self.region_keyword_map = region_keyword_map
        self.seminar_keywords = seminar_keywords
        self.display_timezone = pytz.timezone('Asia/Seoul')
        self.korean_weekday = ['월', '화', '수', '목', '금', '토', '일']
        self.weekday_map = {'MO': '월', 'TU': '화', 'WE': '수', 'TH': '목', 'FR': '금', 'SA': '토', 'SU': '일'}
        self.region_order = list(region_keyword_map.keys())

    def _get_week_of_month(self, dt):
        start_of_week = dt - timedelta(days=dt.weekday())
        first_day_of_month = dt.replace(day=1)
        start_of_first_week = first_day_of_month - timedelta(days=first_day_of_month.weekday())
        week_diff = (start_of_week - start_of_first_week).days // 7
        return week_diff + 1

    def _format_rrule_for_display(self, rrule_data):
        """RRULE 데이터를 한글로 변환"""
        if not isinstance(rrule_data, vRecur):
            return ""

        freq = rrule_data.get('FREQ', [None])[0]
        byday = rrule_data.get('BYDAY')

        if freq == 'DAILY':
            return "(매일) "
        elif freq == 'WEEKLY':
            if byday:
                if sorted(byday) == ['FR', 'MO', 'TH', 'TU', 'WE']:
                    return "(평일) "

                day_order = {day: i for i, day in enumerate(self.weekday_map.keys())}
                sorted_byday = sorted(byday, key=lambda d: day_order.get(d, 99))

                days_korean = ','.join([self.weekday_map.get(day, '') for day in sorted_byday])
                return f"(매주 {days_korean}) "
        return ""

    def _parse_event_data(self, component, region_group, is_all_day):
        summary = str(component.get('summary', '제목 없음'))
        location = str(component.get('location', ''))
        description = str(component.get('description', ''))

        url = component.get('url')
        if url:
            url = str(url)
        else:
            url_pattern = r'https?://[^\s<>"]+|www\.[^\s<>"]+'
            html_url_pattern = r'<a\s+(?:[^>]*?\s+)?href="([^"]*)"'
            html_match = re.search(html_url_pattern, description)
            if html_match:
                url = html_match.group(1)
            else:
                plain_match = re.search(url_pattern, description)
                url = plain_match.group(0) if plain_match else None

        event_type = '세미나' if any(kw in summary for kw in self.seminar_keywords) else '투쟁'

        return {
            'title': summary,
            'location': location,
            'description': description,
            'url': url,
            'region_group': region_group,
            'event_type': event_type,
            'is_all_day': is_all_day,
        }

    def _generate_output_string(self, structured_events, start_date_utc, end_date_utc):
        start_date_local = start_date_utc.astimezone(self.display_timezone)
        end_date_local = (end_date_utc - timedelta(days=1)).astimezone(self.display_timezone)

        start_str = start_date_local.strftime('%Y년 %m월 %d일')
        if start_date_local.year == end_date_local.year:
            end_str = end_date_local.strftime('%m월 %d일')
        else:
            end_str = end_date_local.strftime('%Y년 %m월 %d일')

        date_range_str = f"🗓️ {start_str} ~ {end_str} 일정"

        output_parts = [
            date_range_str,
            "⬇️ 타래로 각 지역의 투쟁 캘린더가 올라갑니다 ⬇️\n"
        ]
        sorted_regions = sorted(structured_events.keys(),
                                key=lambda r: self.region_order.index(r) if r in self.region_order else 99)
        region_counter = 1

        for region_name in sorted_regions:
            display_region_name = region_name.replace('/', '&')  # 트위터 핸들 문제 방지
            for event_type in ['투쟁', '세미나']:
                region_events = structured_events[region_name][event_type]
                recurring_events = region_events.get('recurring', [])
                single_events = region_events.get('single', [])

                if not recurring_events and not single_events:
                    continue

                if single_events:
                    first_event_time = sorted(single_events, key=lambda x: x['start_time'])[0]['start_time']
                else:
                    # 반복일정만 있을 경우, 조회 시작 날짜를 기준으로 주차를 계산
                    first_event_time = start_date_local

                week_num = self._get_week_of_month(first_event_time)
                month = first_event_time.month
                type_icon = "🔥" if event_type == '투쟁' else "📓"
                type_text = "투쟁" if event_type == '투쟁' else "세미나"

                output_parts.append(f"({region_counter}) ({display_region_name} {type_text})")
                output_parts.append(f"{month}월 {week_num}주차 ({display_region_name} {type_text})\n")
                output_parts.append(f"{type_icon}{type_text}일정 안내{type_icon}\n")
                region_counter += 1

                if recurring_events:
                    date_icon = "✊" if event_type == '투쟁' else "📓"
                    output_parts.append(f"📢 {week_num}주차 반복일정 {date_icon}\n")

                    def get_sort_key_for_recurring(event):
                        start_time = event['start_time_orig']
                        if isinstance(start_time, datetime):
                            return start_time.time()
                        return datetime.min.time()

                    sorted_recurring = sorted(recurring_events, key=get_sort_key_for_recurring)

                    for event in sorted_recurring:
                        rrule_text = self._format_rrule_for_display(event['rrule_obj'])
                        time_str = ""
                        if not event['is_all_day']:
                            time_str = f"{event['start_time_orig'].strftime('%H:%M')} "

                        output_parts.append(f"↓ {event['title']}")
                        output_parts.append(f"{rrule_text}{time_str}{event['location']}")

                        if event['url']:
                            output_parts.append(f"{event['url']}\n")
                        else:
                            output_parts.append("\n")

                if single_events:
                    events_by_date = defaultdict(list)
                    for e in single_events:
                        events_by_date[e['start_time'].date()].append(e)

                    for date in sorted(events_by_date.keys()):
                        day_str = self.korean_weekday[date.weekday()]
                        date_icon = "✊" if event_type == '투쟁' else "📓"
                        output_parts.append(f"📢 {date.month}월 {date.day}일 ({day_str}) {date_icon}\n")

                        sorted_day_events = sorted(events_by_date[date], key=lambda x: x['start_time'])
                        for event in sorted_day_events:
                            output_parts.append(f"↓ {event['title']}")
                            if event['is_all_day']:
                                output_parts.append(f"{event['location']}")
                            else:
                                output_parts.append(f"{event['start_time'].strftime('%H:%M')} {event['location']}")

                            if event['url']:
                                output_parts.append(f"{event['url']}\n")
                            else:
                                output_parts.append("\n")

        output_parts.append(f"({region_counter})\n\n🎤 투쟁 기자들에게 제보하기 🎤\n")
        output_parts.append(
            "제보 폼: https://docs.google.com/forms/d/e/1FAIpQLSfr0XK6NPsuXAWHGxJaGv8DALJAAA3rQ8rDv3F3ZWo6hmZUfw/viewform")
        return "\n".join(output_parts)

    def run(self, start_date_str=None, end_date_str=None):
        if start_date_str and end_date_str:
            # 사용자가 날짜를 입력한 경우
            try:
                start_dt_local = self.display_timezone.localize(datetime.strptime(start_date_str, '%Y-%m-%d'))
                # 종료일은 그 날의 마지막까지 포함하기 위해 하루를 더하고, 시간은 0시 0분으로 설정
                end_dt_local = self.display_timezone.localize(datetime.strptime(end_date_str, '%Y-%m-%d')) + timedelta(
                    days=1)

                start_date_utc = start_dt_local.astimezone(pytz.utc)
                end_date_utc = end_dt_local.astimezone(pytz.utc)
                print(f"지정된 기간으로 일정을 검색합니다: {start_date_str} ~ {end_date_str}")
            except ValueError:
                print("❌ [오류] 날짜 형식이 잘못되었습니다. 'YYYY-MM-DD' 형식으로 입력해주세요.")
                return None
        else:
            # 기본값 (오늘부터 7일)
            start_date_utc = datetime.now(pytz.utc)
            end_date_utc = start_date_utc + timedelta(days=7)
            print("기본값(오늘부터 7일)으로 일정을 검색합니다.")

        structured_events = defaultdict(lambda: defaultdict(lambda: {'recurring': [], 'single': []}))

        try:
            all_files = os.listdir(self.directory_path)
        except FileNotFoundError:
            print(f"❌ [오류] 폴더를 찾을 수 없습니다: '{self.directory_path}'")
            print("   '폴더 선택' 버튼으로 올바른 폴더를 선택해주세요.")
            return None

        ics_files = [f for f in all_files if f.endswith('.ics')]
        if not ics_files:
            print(f"⚠️  [경고] '{self.directory_path}' 폴더에 .ics 파일이 없습니다.")
            return None

        for filename in ics_files:
            full_path = os.path.join(self.directory_path, filename)
            try:
                with open(full_path, 'rb') as f:
                    cal_content = f.read()

                cal_text = cal_content.decode('utf-8', errors='ignore')
                pattern = re.compile(r'UNTIL=([0-9]{8}(?:T[0-9]{6})?)(?=[;]|$)', re.IGNORECASE)

                def replacer(match):
                    original_until_value = match.group(1)
                    if 'T' in original_until_value:
                        new_until_value = original_until_value + 'Z'
                    else:
                        new_until_value = original_until_value + 'T235959Z'
                    return f"UNTIL={new_until_value}"

                cal_text = pattern.sub(replacer, cal_text)
                cal = Calendar.from_ical(cal_text.encode('utf-8'))

                cal_name = str(cal.get('X-WR-CALNAME', ''))
                if not cal_name:
                    print(f"⚠️  [경고] '{filename}' 파일에 캘린더 이름이 없어 건너뜁니다.")
                    continue

                matched_region = None
                for region_group, keywords in self.region_keyword_map.items():
                    if any(keyword in cal_name for keyword in keywords):
                        matched_region = region_group
                        break

                if not matched_region:
                    print(f"⚠️  [경고] '{cal_name}'(파일: {filename})은(는) 설정된 지역과 매칭되지 않아 건너뜁니다.")
                    continue

                for component in cal.walk('VEVENT'):
                    dtstart_prop = component.get('dtstart')
                    if not dtstart_prop:
                        continue

                    dtstart = dtstart_prop.dt
                    is_all_day = not isinstance(dtstart, datetime)

                    rrule = component.get('RRULE')
                    if rrule:
                        if is_all_day:
                            start_time_for_rule = pytz.utc.localize(datetime.combine(dtstart, datetime.min.time()))
                        elif dtstart.tzinfo is None:
                            start_time_for_rule = pytz.utc.localize(dtstart)
                        else:
                            start_time_for_rule = dtstart.astimezone(pytz.utc)

                        try:
                            rule = rrulestr(rrule.to_ical().decode(), dtstart=start_time_for_rule)
                            if rule.between(start_date_utc, end_date_utc, inc=True):
                                event_details = self._parse_event_data(component, matched_region, is_all_day)
                                event_details['rrule_obj'] = rrule
                                event_details['start_time_orig'] = dtstart
                                structured_events[matched_region][event_details['event_type']]['recurring'].append(
                                    event_details)
                        except Exception as rule_error:
                            print(f"⚠️  [경고] '{filename}' 파일의 반복 규칙 처리 중 오류 발생: {rule_error}")

                    else:
                        utc_start_time = None
                        if is_all_day:
                            utc_start_time = pytz.utc.localize(datetime.combine(dtstart, datetime.min.time()))
                        elif dtstart.tzinfo is None:
                            utc_start_time = pytz.utc.localize(dtstart)
                        else:
                            utc_start_time = dtstart.astimezone(pytz.utc)

                        if start_date_utc <= utc_start_time < end_date_utc:
                            event_details = self._parse_event_data(component, matched_region, is_all_day)
                            event_details['start_time'] = utc_start_time.astimezone(self.display_timezone)
                            structured_events[matched_region][event_details['event_type']]['single'].append(
                                event_details)

            except Exception as e:
                print(f"❌ [오류] '{filename}' 파일 처리 중 문제가 발생했습니다: {e}")

        if not structured_events:
            message = "선택하신 기간에 해당하는 일정이 없습니다."
            print(message)
            return message

        final_text = self._generate_output_string(structured_events, start_date_utc, end_date_utc)
        print("\n✅ 작업이 완료되었습니다. 위쪽 창에서 생성된 텍스트를 복사하여 사용하세요.")
        return final_text


class App:
    """GUI 애플리케이션 클래스"""

    def __init__(self, master):
        self.master = master
        master.title("투쟁일정 캘린더 텍스트 생성기 v1.0")
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


if __name__ == '__main__':
    root = Tk()
    app = App(root)
    root.mainloop()
