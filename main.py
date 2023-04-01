from dataclasses import dataclass, field, asdict
import requests
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from io import BytesIO
from itertools import islice
import datetime
import re
import sys


@dataclass
class Exam:
    educator_id: str
    educator: str = field(init=False)
    time: tuple[int, int]
    date: datetime.date
    title: str
    where: str
    group_full: str
    group: str = field(init=False)

    def __post_init__(self):
        self.educator = ''
        self.group = ''
        self.double = False

    def __str__(self):
        return '{ed:<3} {date.day:0>2}.{date.month:0>2} {time[0]:0>2}:{time[1]:0>2} - {group} {title} @ {where}'.format(
            ed=self.educator,
            date=self.date,
            time=self.time,
            group=self.group,
            title=self.title,
            where=self.where
        )

    def is_double(self, other):
        exam1 = asdict(self)
        exam2 = asdict(other)
        return [entry for entry in exam1 if exam1[entry] != exam2[entry]] == ['educator_id']


def parse_date(date_str, single_date_regex=re.compile('^(\d+)\.(\d+)$')):
    m = single_date_regex.match(date_str)
    if m:
        today = datetime.date.today()
        day, month = map(int, [m.group(1), m.group(2)])
        year = today.year
        if 9 <= month <= 12 and 1 <= today.month <= 8:
            year -= 1
        elif 9 <= today.month <= 12 and 1 <= month <= 8:
            year += 1
        return datetime.date(year, month, day)
    return None


def parse_time(time_str, time_regex=re.compile('^(\d+):(\d+)–(\d+):(\d+)$')):
    m = time_regex.match(time_str)
    if m:
        return int(m.group(1)), int(m.group(2))
    return 0, 0


def parse_tt_excel(ws: Worksheet, educator_id):
    exams_list = []
    for row in islice(ws.rows, 4, None, None):
        exam = Exam(
            educator_id=educator_id,
            time=parse_time(row[1].value),
            date=parse_date(row[2].value),
            title=row[3].value,
            where=row[4].value,
            group_full=row[5].value,
        )
        if any(exam_type in exam.title for exam_type in ['зачёт', 'экзамен']):
            exams_list.append(exam)
    return exams_list


def compile_exams_table(educator_aliases=None, excluded_depts=(), group_aliases=None):
    tt_base_link = 'https://timetable.spbu.ru/'
    tt_educator_link_suffix = 'EducatorEvents/'
    educators_number = len(educator_aliases)
    exams_list = []
    for i, educator_id in enumerate(educator_aliases):
        url = '{}{}{}/Excel'.format(tt_base_link, tt_educator_link_suffix, educator_id)
        resp = requests.get(url, cookies={'_culture': 'ru'})
        sys.stdout.write('\rLoading: [{}░{}]'.format('█' * i, ' ' * (educators_number - i - 1)))
        sys.stdout.flush()
        wb = openpyxl.load_workbook(BytesIO(resp.content))
        educator_tt = wb.active
        sys.stdout.write('\rLoading: [{}▒{}]'.format('█' * i, ' ' * (educators_number - i - 1)))
        sys.stdout.flush()
        educator_exams_list = parse_tt_excel(educator_tt, educator_id)
        exams_list.extend(educator_exams_list)
        sys.stdout.write('\rLoading: [{:<{}}]'.format('█' * (i + 1), educators_number))
        sys.stdout.flush()
    print()
    today = datetime.date.today()
    buckets = {
        'Обычные': [],
        'Комиссии': [],
        # 'Ничейные': [],
        'Прошедшие': [],
    }
    for exam in exams_list:
        if all(dept not in exam.group_full for dept in excluded_depts):
            bucket_name = 'Обычные'
            if exam.date < today:
                bucket_name = 'Прошедшие'
            elif 'комиссия' in exam.title:
                bucket_name = 'Комиссии'
            buckets[bucket_name].append(exam)
    for bucket_name, bucket in buckets.items():
        bucket.sort(key=lambda exam: (exam.date, exam.title), reverse=True if bucket_name == 'Прошедшие' else False)
        exam1 = bucket.pop(0)
        bucket_dedoubled = [exam1]
        for exam2 in bucket:
            if exam1.is_double(exam2):
                exam1.double = True
            else:
                bucket_dedoubled.append(exam2)
                exam1 = exam2
        buckets[bucket_name] = bucket_dedoubled
    outputs = []
    for bucket_name, bucket in buckets.items():
        if bucket:
            for exam in bucket:
                exam.educator = render_educator(exam.educator_id, educator_aliases)
                if exam.double:
                    exam.educator += ' et al.'
                exam.group = render_group(exam.group_full, group_aliases)
    for bucket_name, bucket in buckets.items():
        outputs.append('{}:\n\n{}'.format(bucket_name, '\n'.join(map(str, bucket))))
    return '\n\n'.join(outputs)



def render_group(full_group_name, group_aliases, group_regex=re.compile('^(\d+)\.([БС0-9]+-мм).*(\d) курс\)')):
    group_name = full_group_name[:full_group_name.index(' ')]
    if group_aliases:
        m = group_regex.match(full_group_name)
        if m:
            group_code = m.group(2)
            if group_code in group_aliases:
                group_number = '{}{}'.format(m.group(3), group_aliases[group_code])
                return '{} ({})'.format(group_name, group_number)
    return group_name


def render_educator(educator_id, educator_aliases):
    return educator_aliases[educator_id]


if __name__ == '__main__':
    group_aliases = {}
    with open('groups.txt', 'r', encoding='UTF-8') as file:
        for line in file:
            group_name, group_number = line.split()
            group_aliases[group_name] = group_number
    educators_aliases = {}
    with open('educators.txt', 'r', encoding='UTF-8') as file:
        for line in file:
            educator, ed_id = line.split()
            educators_aliases[ed_id] = educator
    exams_table = compile_exams_table(educators_aliases, group_aliases=group_aliases, excluded_depts=['мкн'])
    filename = 'exams {}.txt'.format(datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S'))
    with open(filename, 'w', encoding='UTF-8') as file:
        file.write(exams_table)
