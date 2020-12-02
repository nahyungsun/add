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
import pandas as pd
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
    location_list = None
    
    identifed_objects = {}

    def to_csv(self):
        list = []

        list.append(self.date.strftime("%Y-%m-%d"))
        list.append(self.time_from.strftime('%H:%M'))
        list.append(self.time_to.strftime('%H:%M'))

        for army_type in army_types:
            list.append(",".join(self.identifed_objects[army_type]))


        return list
    def __str__(self):
        r = ''
        r += 'date: ' + self.date.strftime("%Y-%m-%d") +"\n"
        r += 'time_from: ' + self.time_from.strftime('%H:%M') +"\n"
        r += 'time_to: ' + self.time_to.strftime('%H:%M') +"\n"
        r += 'location_list: ' + ",".join(self.location_list) +"\n"

        identifed_objects_str = json.dumps(self.identifed_objects, ensure_ascii=False)

        r += 'identifed_objects: ' + identifed_objects_str +"\n"

        return r
    def filename_proc(self,filename):
        #파일이름에서 날자 데이터 따오기
        #ex) c2011131730
        #앞 문자는 양식 종류추정 그 뒤로는 yymmddhhmm
        date_str = filename.split(".")[0]#.뒷부분 날리기
        self.file_type = re.findall(r"\D",date_str)#양식 종류
        #중요할경우 output에 추가예정
        date_str = re.findall(r"\d{10}",date_str)#yymmddhhmm
        self.date = datetime.datetime.strptime("20"+date_str[0][0:6], "%Y%m%d")
        

    def task_status_proc(self, lines, line_index):
        while line_index < len(lines):
            line = lines[line_index].strip()
            line_index += 1

            if "촬영지역" in line:
                line = lines[line_index].strip()
                line_index += 1
                self.location_list = line.split(",")
            if "촬영시간" in line:
                line = lines[line_index].strip()
                line_index += 1
                self.time_from = datetime.datetime.strptime(line.split("~")[0], "%H:%M")
                self.time_to = datetime.datetime.strptime(line.split("~")[1], "%H:%M")

            if self.location_list is not None and self.time_from is not None and self.time_to is not None:
                break

        return line_index


    def idenfied_objects_proc(self, lines, line_index):

        while line_index < len(lines):
            line = lines[line_index].strip()
            line_index += 1

            if begin_with_number.match(line) is not None:
                army_type = line.split(" ")[1]
                line_index = self.idenfied_objects_internal_proc(army_type,  lines, line_index)

        return line_index
    def idenfied_objects_internal_proc(self, army_type, lines, line_index):

        self.identifed_objects[army_type] = []
        #print("/"+army_type)
        while line_index < len(lines):
            line = lines[line_index].strip()
            line_index += 1
            #print("\t",line)
            if begin_with_number.match(line) is not None:
                #print("\t",line)
                return line_index - 1
            if line.startswith(":"):

                for left_word in left_words:
                    if left_word in line:
                        line = line.split(left_word)[1]
                        before_line = line
                        for right_word in right_words:
                            if right_word in line:
                                line = line.split(right_word)[0]
                        if before_line == line:
                            line = right_number_regex.split(line)[0].strip()

                        self.identifed_objects[army_type].append(line)
        return line_index


# In[9]:


#text = '63식 장갑차 3대 식별'
right_number_regex = re.compile("\d대")
#ww = right_number_regex.split(text)
###tsv파일 만든후 칼럼 저장
#hs : D:/hs/ws/add/add/input
home_dir = "D:/hs/ws/add/add/input"#한글파일 경로
if path.exists("output.tsv"):#tsv파일이 존재할경우 삭제
    remove("output.tsv")
output_file = open("output.tsv", mode="a", encoding="utf-8", newline="")
output_writer = csv.writer(output_file, delimiter='\t')

army_types = ['지상군','해군','공군']
colnames = ['날짜','촬영시간(시작)','촬영시간(종료)']
colnames.extend(army_types)#칼럼
output_writer.writerow(colnames)#만든 tsv파일에 칼럼 추가 \t
output_file.close()#파일접근권한 해제
RESOURCE_PATH_XSL_TEXT = 'xsl/plaintext_with_table.xsl'
#표도 변환하기 위해 수정한 xsl
begin_with_number = re.compile("^\d\)")
left_words = ['인근에','주변']
right_words = ['추정']


# In[11]:


text_transform = TextTransform()#객체생성
transform = text_transform.transform_hwp5_to_text
onlyfiles = [f for f in listdir(home_dir) if isfile(join(home_dir, f))]#디렉토리의 모든 파일명
for file in onlyfiles:

    if not file.endswith(".hwp"):
        #파일의 확장자가 hwp가 아닐경우 수행하지 않음
        continue

    hwp5path = join(home_dir, file)#경로
    outputfile_name = hwp5path+".txt"
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
        #hwp.txt파일 만드는 과정
#####################################################################
    outputfile = open(outputfile_name, mode="r", encoding='utf-8')
    lines = outputfile.read().splitlines()
    outputfile.close()
    #content = [x.strip() for x in content]

    report = Report()
    report.filename_proc(file)
    #파일이름 분석
    
    line_index = 0
    while line_index < len(lines):
        line = lines[line_index].strip()#양쪽 끝에 있는 공백과 \n기호 삭제
        line_index += 1
        if line == '':#공백일경우 건너 뛰기
            continue

        #print(line)
        if "임무현황" in line:
            line_index = report.task_status_proc(lines, line_index)
            continue

        if "4. 수집 자산" in line:
            line_index = report.idenfied_objects_proc(lines, line_index)
            continue
    
    #결과물 저장
    output_file = open("output.tsv", mode="a", encoding="utf-8", newline="")
    output_writer = csv.writer(output_file, delimiter='\t')
    output_writer.writerow(report.to_csv())
    output_file.close()
    print(report)
    del report


# In[ ]:




