import os
from icalendar import Calendar

# --- ⚙️ 설정 ---
# 여기에 다운로드한 .ics 파일들이 모여 있는 **폴더 경로**를 입력하세요.
# 예시: 'C:/Users/YourUser/Downloads/MyCalendars'
#      'data/malbeolsimin@gmail.com.ical'
ICS_DIRECTORY_PATH = 'data/malbeolsimin@gmail.com.ical'


# -----------------

def inspect_ics_files(directory_path):
    """
    지정된 디렉토리의 모든 .ics 파일 내용을 검사하여
    캘린더 이름과 샘플 일정을 출력합니다.
    """
    print(f"🔍 '{directory_path}' 폴더에서 .ics 파일을 검사합니다...\n")

    try:
        # 디렉토리의 모든 파일 목록을 가져옵니다.
        all_files = os.listdir(directory_path)
    except FileNotFoundError:
        print(f"❌ [오류] 폴더를 찾을 수 없습니다: '{directory_path}'")
        print("ICS_DIRECTORY_PATH 변수에 올바른 폴더 경로를 입력했는지 확인해주세요.")
        return

    # .ics 파일만 필터링합니다.
    ics_files = [f for f in all_files if f.endswith('.ics')]

    if not ics_files:
        print("해당 폴더에 .ics 파일이 없습니다.")
        return

    # 각 .ics 파일을 순회하며 내용을 확인합니다.
    for filename in ics_files:
        print(f"--- 파일: {filename} ---")
        full_path = os.path.join(directory_path, filename)

        try:
            with open(full_path, 'rb') as f:
                cal = Calendar.from_ical(f.read())

            # 1. 캘린더 이름 확인 (가장 확실한 단서)
            cal_name = cal.get('X-WR-CALNAME')
            if cal_name:
                print(f"* 캘린더 이름: {cal_name}")
            else:
                print("* 캘린더 이름: (설정되지 않음)")

            # 2. 샘플 일정 제목 확인 (내용으로 유추)
            print("* 샘플 일정 (최대 5개):")
            event_count = 0
            for component in cal.walk('VEVENT'):
                if event_count >= 5:
                    break
                summary = component.get('summary')
                if summary:
                    print(f"  - {summary}")
                    event_count += 1

            if event_count == 0:
                print("  (일정이 없습니다)")

        except Exception as e:
            print(f"  [오류] 파일을 읽는 중 문제가 발생했습니다: {e}")

        print("-" * (len(filename) + 10) + "\n")


if __name__ == '__main__':
    inspect_ics_files(ICS_DIRECTORY_PATH)
