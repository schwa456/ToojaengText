import os
import re
from collections import defaultdict
from datetime import datetime, timedelta

from dateutil.rrule import rrulestr
from icalendar import Calendar, vRecur
import pytz


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

                    # 반복 일정은 시작 시간으로 정렬
                    sorted_recurring = sorted(recurring_events, key=lambda e: e['start_time'])

                    for event in sorted_recurring:
                        # [수정] 생성된 반복 정보를 그대로 사용
                        recurrence_text = event.get('recurrence_info', '(반복) ')

                        time_str = ""
                        if not event['is_all_day']:
                            time_str = f"{event['start_time'].strftime('%H:%M')} "

                        output_parts.append(f"↓ {event['title']}")
                        output_parts.append(f"{recurrence_text}{time_str}{event['location']}")

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
            try:
                start_dt_local = self.display_timezone.localize(datetime.strptime(start_date_str, '%Y-%m-%d'))
                end_dt_local = self.display_timezone.localize(
                    datetime.strptime(end_date_str, '%Y-%m-%d')) + timedelta(days=1)
                start_date_utc = start_dt_local.astimezone(pytz.utc)
                end_date_utc = end_dt_local.astimezone(pytz.utc)
                print(f"지정된 기간으로 일정을 검색합니다: {start_date_str} ~ {end_date_str}")
            except ValueError:
                print("❌ [오류] 날짜 형식이 잘못되었습니다. 'YYYY-MM-DD' 형식으로 입력해주세요.")
                return None
        else:
            start_date_utc = datetime.now(pytz.utc)
            end_date_utc = start_date_utc + timedelta(days=7)
            print("기본값(오늘부터 7일)으로 일정을 검색합니다.")

        # 1. 모든 발생 건을 저장할 임시 리스트
        all_occurrences = []

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
                    event_details = self._parse_event_data(component, matched_region, is_all_day)

                    rrule = component.get('RRULE')
                    if rrule:
                        start_time_for_rule = dtstart
                        if is_all_day:
                            start_time_for_rule = pytz.utc.localize(datetime.combine(dtstart, datetime.min.time()))
                        elif dtstart.tzinfo is None:
                            start_time_for_rule = pytz.utc.localize(dtstart)
                        else:
                            start_time_for_rule = dtstart.astimezone(pytz.utc)

                        try:
                            rule = rrulestr(rrule.to_ical().decode(), dtstart=start_time_for_rule)
                            # 조회 기간 내 모든 발생 건을 생성
                            for occurrence_dt in rule.between(start_date_utc,
                                                              end_date_utc - timedelta(microseconds=1), inc=True):
                                new_event = event_details.copy()
                                new_event['start_time'] = occurrence_dt.astimezone(self.display_timezone)
                                # [수정] RRULE 객체를 이벤트 정보에 추가
                                new_event['rrule_obj'] = rrule
                                all_occurrences.append(new_event)
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
                            event_details['start_time'] = utc_start_time.astimezone(self.display_timezone)
                            all_occurrences.append(event_details)

            except Exception as e:
                print(f"❌ [오류] '{filename}' 파일 처리 중 문제가 발생했습니다: {e}")

        # 2. 동일한 이벤트를 그룹화
        grouped_events = defaultdict(list)
        for event in all_occurrences:
            # 제목, 장소, 지역, 타입으로 고유한 이벤트를 식별
            event_key = (event['title'], event['location'], event['region_group'], event['event_type'])
            grouped_events[event_key].append(event)

        # 3. 그룹화된 이벤트를 '반복'과 '단일'로 최종 분류
        structured_events = defaultdict(lambda: defaultdict(lambda: {'recurring': [], 'single': []}))
        for key, occurrences in grouped_events.items():
            region_group = key[2]
            event_type = key[3]

            # 발생 횟수가 2번 이상이면 '반복'으로 처리
            if len(occurrences) >= 2:
                representative_event = occurrences[0]

                #  RRULE 객체가 있는지 확인하여 반복 정보 생성
                rrule_obj = representative_event.get('rrule_obj')
                if rrule_obj:
                    # RRULE이 있으면 요일 정보로 포맷팅
                    recurrence_text = self._format_rrule_for_display(rrule_obj)
                    # 포맷팅 결과가 비어있을 경우, 횟수 정보로 대체
                    if not recurrence_text.strip():
                        recurrence_text = f"(기간 내 {len(occurrences)}회 반복) "
                else:
                    # RRULE이 없으면 횟수 정보로 표시
                    recurrence_text = f"(기간 내 {len(occurrences)}회 반복) "

                representative_event['recurrence_info'] = recurrence_text
                structured_events[region_group][event_type]['recurring'].append(representative_event)
            # 발생 횟수가 1번이면 '단일'로 처리
            else:
                structured_events[region_group][event_type]['single'].append(occurrences[0])

        if not structured_events:
            message = "선택하신 기간에 해당하는 일정이 없습니다."
            print(message)
            return message

        final_text = self._generate_output_string(structured_events, start_date_utc, end_date_utc)
        print("\n✅ 작업이 완료되었습니다. 위쪽 창에서 생성된 텍스트를 복사하여 사용하세요.")
        return final_text

