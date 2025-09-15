import os
from icalendar import Calendar
import pytz

# --- ⚙️ 설정 ---
# 1. 내용을 확인하고 싶은 .ics 파일의 '전체 경로'를 여기에 입력하세요.
#    (이전에 사용한 check_ics_files.py로 강원도 캘린더의 파일명을 먼저 확인하세요)
#
ICS_FILE_TO_INSPECT = '../data/malbeolsimin@gmail.com.ical/a02fcc932dd53f4525fc19abdf0171db111f3956367ca2c1dfa0dd7d5d9c0b74@group.calendar.google.com.ics'


# -----------------

def inspect_one_file(file_path):
    """
    지정된 .ics 파일 하나를 열어, 캘린더 이름과 '모든' 일정의
    상세 정보(시작일, 제목, 장소)를 출력합니다.
    """
    print(f"🔍 파일 상세 분석 시작: '{os.path.basename(file_path)}'\n")

    try:
        with open(file_path, 'rb') as f:
            cal = Calendar.from_ical(f.read())
    except FileNotFoundError:
        print(f"❌ [오류] 파일을 찾을 수 없습니다: '{file_path}'")
        print("ICS_FILE_TO_INSPECT 변수에 올바른 파일 경로를 입력했는지 확인해주세요.")
        return
    except Exception as e:
        print(f"❌ [오류] 파일을 읽는 중 문제가 발생했습니다: {e}")
        return

    # 1. 캘린더 이름 확인
    cal_name = cal.get('X-WR-CALNAME')
    if cal_name:
        print(f" * 캘린더 이름: {cal_name}\n")
    else:
        print(" * 캘린더 이름: (설정되지 않음)\n")

    # 2. 파일 내의 '모든' 일정 상세 정보 확인
    print("--- 포함된 전체 일정 목록 ---")
    event_count = 0
    for component in cal.walk('VEVENT'):
        event_count += 1
        summary = component.get('summary', '제목 없음')
        dtstart = component.get('dtstart').dt
        location = component.get('location', '장소 정보 없음')
        is_recurring = component.get('RRULE') is not None

        print(f"\n✅ [{event_count}] {summary}")
        print(f"  - 시작: {dtstart} (타입: {type(dtstart).__name__})")
        print(f"  - 장소: {location}")
        print(f"  - 반복: {is_recurring}")

    if event_count == 0:
        print("\n  (파일에 포함된 일정이 없습니다)")

    print("\n" + "=" * 40)
    print("분석 완료.")


if __name__ == '__main__':
    inspect_one_file(ICS_FILE_TO_INSPECT)
