#!/usr/bin/env python
# coding: utf-8

# In[1]:


import hwp5
from os import remove
from os import listdir
from os.path import isfile, join
from hwp5.transforms import BaseTransform
from hwp5.utils import cached_property
from hwp5.xmlmodel import Hwp5File
from hwp5.utils import make_open_dest_file
from contextlib import closing
from hwp5.dataio import ParseError
from hwp5.errors import InvalidHwp5FileError
import datetime
import json
import re
from os import path
#import pandas as pd
import csv


# In[2]:


class TextTransform(BaseTransform):

    @property
    def transform_hwp5_to_text(self):
        transform_xhwp5 = self.transform_xhwp5_to_text
        return self.make_transform_hwp5(transform_xhwp5)

    @cached_property
    def transform_xhwp5_to_text(self):
        return self.xslt_compile(RESOURCE_PATH_XSL_TEXT)


class Report:
    file_type = None
    date = None
    time_from = None
    time_to = None
    assets_time = None
    location_list = None
    assets_list = None
    assets_count = 0
    # 입수자산 개수 카운트
    identifed_objects = {}
    data=None
    def to_csv(self):

        temp = []
        pd_columns=["수집자산","날짜","시간","지역","data"]
        #list.append(self.date.strftime("%Y-%m-%d"))
        # list.append(self.time_from.strftime('%H:%M'))
        # list.append(self.time_to.strftime('%H:%M'))
        #for army_type in army_types:
        #    print(",".join(self.identifed_objects[army_type]))
        #    list.append(",".join(self.identifed_objects[army_type]))
        #print(self.assets_list)
        #print(self.location_list)
        #print(self.date)
        #print(self.file_type)
        #print(self.assets_time)
        #print(self.data)
        output_file = open("output.tsv", mode="a", encoding="utf-8", newline="")
        output_writer = csv.writer(output_file, delimiter='\t')

        temp = self.data
        for a in temp:
            if a[2:] == []:
                # 데이터가 없을 경우 넘기기
                continue
            list=[]
            list.append(a[0])#수집자산
            list.append(self.date)
            for b in range(len(self.assets_list)):
                if a[0] == self.assets_list[b]:
                    list.append("".join(self.assets_time[b]))
                    list.append("".join(self.location_list[b]))
            list.append(a[1:])
            
            output_writer.writerow(list)
        return

    def __init__(self):
        self.file_type = None
        self.date = None
        self.time_from = None
        self.time_to = None
        self.assets_time = None
        self.location_list = None
        self.assets_list = None
        self.assets_count = 0
        self.identifed_objects = {}
        self.data=[]

    def __str__(self):
        #r = ''
        #r += 'date: ' + self.date.strftime("%Y-%m-%d") + "\n"
        #r += 'time_from: ' + self.time_from.strftime('%H:%M') + "\n"
        #r += 'time_to: ' + self.time_to.strftime('%H:%M') + "\n"
        #r += 'location_list: ' + ",".join(self.location_list) + "\n"

        #identifed_objects_str = json.dumps(self.identifed_objects, ensure_ascii=False)

        #r += 'identifed_objects: ' + identifed_objects_str + "\n"
        print(self.assets_list)
        print(self.location_list)
        print(self.date)
        print(self.file_type)
        print(self.assets_time)
        print(self.data)
        return "yes?"

    def filename_proc(self, filename):
        # 파일이름에서 날자 데이터 따오기
        # ex) c2011131730
        # 앞 문자는 양식 종류추정 그 뒤로는 yymmddhhmm
        date_str = filename.split(".")[0]  # .뒷부분 날리기
        self.file_type = re.findall(r"\D", date_str)  # 양식 종류
        # 중요할경우 output에 추가예정
        date_str = re.findall(r"\d{10}", date_str)  # yymmddhhmm
        self.date = "20" + date_str[0][0:6]
        #datetime.datetime.strptime("20" + date_str[0][0:6], "%Y%m%d")

    def task_status_proc(self, lines, line_index):
        # 임무현황 발견했을때
        while line_index < len(lines):
            line = lines[line_index].strip()
            line_index += 1
            if "입수자산" in line or "입수 자산" in line:
                # 입수자산이 여러개일 경우 시간및 촬영 지역이 다를수 있음.
                self.assets_list = []
                tntl=0#같은 수집자산 (수시)붙은거 나올시 예외처리
                while 1:
                    line = lines[line_index].strip()
                    if "촬영면적" in line:
                        break
                    line_index += 1
                    if "수시" in line:
                        tntl=1
                    if "(" == line[0]:
                        continue
                    self.assets_count += 1
                    self.assets_list.append(line.split("(")[0])
                if tntl == 1:
                    kk = len(self.assets_list)
                    for ii in range(0,kk):
                        for jj in range(ii+1,kk):
                            if self.assets_list[ii]==self.assets_list[jj]:
                                self.assets_list[ii]+="(수시)"

            if "촬영지역" in line or "촬영 지역" in line:
                self.location_list = []
                for i in range(self.assets_count):
                    # 입수 자산개수만큼 촬영지역이 존재
                    line = lines[line_index].strip()
                    line_index += 1
                    self.location_list.append(line)
                    while line[-1:]==',':
                        line = lines[line_index].strip()
                        line_index += 1
                        self.location_list[-1]+=line

            if "촬영시간" in line or "촬영 시간" in line:
                self.assets_time = []
                #for i in range(self.assets_count):
                while True:
                    line = lines[line_index].strip()
                    line_index += 1
                    if ":" in line:
                        self.assets_time.append(line)
                    if "입수시간" in line:
                            break
                # self.time_from = datetime.datetime.strptime(line.split(" ~ ")[0], "%H:%M")
                # self.time_to = datetime.datetime.strptime(line.split(" ~ ")[1], "%H:%M")

            if self.location_list is not None and self.assets_list is not None and self.assets_time is not None:
                # 임무현황에서 수집할수 있는 모든 데이터를 수집하였을때 종료
                break
        return line_index

    def idenfied_objects_proc(self, lines, line_index):
        my_assets=None
        while line_index < len(lines):
            line = lines[line_index].strip()
            for count in range(len(self.assets_list)):
                check = 0
                if line_index+1 >= len(lines):
                    print("line_index >= len(lines)")
                    break
                if "(" in self.assets_list[count]:#수집자산별 분류

                    if re.compile("\D{0,2}\.{0,2}\s{0,2}"+self.assets_list[count].split("(")[0]).match(line):

                        my_assets = self.assets_list[count].split("(")[0]
                        check = 1
                else:

                    if re.compile("\D{0,2}\.{0,2}\s{0,2}"+self.assets_list[count]).match(line):
                        
                        my_assets = self.assets_list[count]
                        check = 1

                if check == 1:
                    line_index += 1
                    line = lines[line_index].strip()

                    if begin_with_number.match(line) is not None:
                        army_type = line.split(" ")[1]
                        line_index = self.idenfied_objects_internal_proc(army_type, lines, line_index, my_assets, count)
                        break

            line_index += 1
            if re.compile("\d{0,2}\.{0,2}\s{0,2}영상별\s{0,2}좌표\s{0,2}현황").match(line) is not None:
                # 5. 영상별 좌표 현황이 나올경우 종료
                break
        return line_index

    def idenfied_objects_internal_proc(self, army_type, lines, line_index,my_assets,my_assets_count):
        self.identifed_objects[army_type] = []
        left_words = [':']
        right_words = ['식별']
        line = lines[line_index].strip()
        print(my_assets)
        print(line.replace(" ", ""))
        while line_index < len(lines):
            line = lines[line_index].strip()
            line_index += 1
            #print(line)


            #print(my_assets+line)
            for i in range(4):
                if army_types[i] in line.replace(" ", ""):
                    j=1

                    list=[]
                    list.append(my_assets)
                    list.append(army_types[i])
                    while j and line_index < len(lines)-1:
                        line = lines[line_index].strip()
                        
                        line_index += 1
                        if re.compile("\D\.\s{0,2}").match(line) is not None:
                            self.data.append(list)
                            return line_index - 2
                        if re.compile("\d{1,2}\.{1,2}\s{0,2}\D").match(line) is not None:
                            # 숫자. 문자 매치 ex)5. 영상별~, 6. 등등
                            self.data.append(list)
                            return line_index - 2
                        #if len(self.assets_list) != 1 and len(self.assets_list) - 1 != my_assets_count:
                            # print(line)
                        #    temp = self.assets_list[my_assets_count + 1].split("(")[0]
                        #    begin_with_number = re.compile(".*" + temp + ".*")
                        #    if begin_with_number.match(line) is not None:
                                #self.data.append(my_assets)
                                #self.data.append(army_types[i])
                        #        self.data.append(list)
                        #        print(line)
                        #        return line_index - 2

                        for k in range(4):
                            if army_types[k]+":" in line.replace(" ", ""):
                                print(line.replace(" ", ""))
                                j=0
                                #line_index += 1
                                
                                break
                        if line == '':
                            continue
                        if j == 0:
                            continue
                        list.append(line)

                    #self.data.append(my_assets)
                    #self.data.append(army_types[i])
                    self.data.append(list)


###
#            if line.startswith("sdafsd:asdfsfd"):
#                #print(army_types)
#                for left_word in left_words:
#                    if left_word in line:
#                        line = line.split(left_word)[1]
#                        #print(line)
#                        before_line = line
#                        for right_word in right_words:
#                            if right_word in line:
#                                if "미식별" not in line:
#                                    line = line.split(right_word)[0].strip()
#                        if before_line == line:
#                            line = right_number_regex.split(line)[0].strip()
#
#                        self.identifed_objects[army_type].append(line)
###
        #print(self.identifed_objects)
        return line_index - 2


# In[9]:


text = '63식 장갑차 3대 식별'
right_number_regex = re.compile("\d대")
ww = right_number_regex.split(text)
# tsv파일 만든후 칼럼 저장
# hs : D:/hs/ws/add/add/input
# add : ".\\input"
home_dir = ".\\input"  # 한글파일 경로
if path.exists("output.tsv"):  # tsv파일이 존재할경우 삭제
    remove("output.tsv")
output_file = open("output.tsv", mode="a", encoding="utf-8", newline="")
output_writer = csv.writer(output_file, delimiter='\t')

army_types = ['지상군', '해군', '공군', '전략']
colnames = ["수집자산", "날짜", "시간", "지역", "data"]

#colnames = ['날짜', '촬영시간(시작)', '촬영시간(종료)']
#colnames.extend(army_types)  # 칼럼
# army_types = ['지상군', '해', '공', '전']
output_writer.writerow(colnames)  # 만든 tsv파일에 칼럼 추가 \t
output_file.close()  # 파일접근권한 해제
RESOURCE_PATH_XSL_TEXT = 'xsl/plaintext_with_table.xsl'
# 표도 변환하기 위해 수정한 xsl
begin_with_number = re.compile("^\d\)")
# left_words = ['추정','에','지역에']
left_words = [':']
right_words = ['식별']

# In[11]:


text_transform = TextTransform()  # 객체생성
transform = text_transform.transform_hwp5_to_text
onlyfiles = [f for f in listdir(home_dir) if isfile(join(home_dir, f))]  # 디렉토리의 모든 파일명
for file in onlyfiles:

    if not file.endswith(".hwp"):
        # 파일의 확장자가 hwp가 아닐경우 수행하지 않음
        continue

    hwp5path = join(home_dir, file)  # 경로
    outputfile_name = hwp5path + ".txt"
    try:
        remove(outputfile_name)
    except FileNotFoundError as e:
        # hwp.txt 파일이 존재하지 않을경우 건너뛰기
        print(e)

    open_dest = make_open_dest_file(outputfile_name)
    try:
        with closing(Hwp5File(hwp5path)) as hwp5file:
            with open_dest() as dest:
                transform(hwp5file, dest)
    except ParseError as e:
        print(e)
    except InvalidHwp5FileError as e:
        print(e)
    except Exception as ex:
        print(ex)
        continue
        # hwp.txt파일 만드는 과정
    #####################################################################
    try:
        outputfile = open(outputfile_name, mode="r", encoding='utf-8')
    except FileNotFoundError as e:
        # hwp.txt 파일이 존재하지 않을경우 건너뛰기
        print(e)
        continue
    lines = outputfile.read().splitlines()
    outputfile.close()
    # content = [x.strip() for x in content]
    report = Report()
    print(file)
    report.filename_proc(file)
    # 파일이름 분석

    line_index = 0
    while line_index < len(lines):
        line = lines[line_index].strip()  # 양쪽 끝에 있는 공백과 \n기호 삭제
        line_index += 1
        if line == '':  # 공백일경우 건너 뛰기
            continue
        elif re.compile("\d{0,2}\.{0,2}\s{0,2}임무\s{0,2}현황").match(line) is not None:
            line_index = report.task_status_proc(lines, line_index)
            continue
        elif re.compile("\d{0,2}\.{0,2}\s{0,2}수집\s{0,2}자산").match(line) is not None:
            line_index = report.idenfied_objects_proc(lines, line_index)
            continue

    # 결과물 저장
    #output_file = open("output.tsv", mode="a", encoding="utf-8", newline="")
    #output_writer = csv.writer(output_file, delimiter='\t')
    #output_writer.writerow(report.to_csv())
    report.to_csv()
    output_file.close()
    #print(report)
    del report

# In[ ]:
