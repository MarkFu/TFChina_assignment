# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import os
import datetime as dt
import pandas as pd

df_teacher = pd.read_excel('data_cleaned.xlsx', sheetname = 'Teachers', header = 0, na_values = "", encoding="gbk")
df_school = pd.read_excel('data_cleaned.xlsx', sheetname = 'Schools', header = 0, na_values = "", encoding="gbk")

## DATA PREPARATION
df_school = df_school[df_school['人数1'.decode('utf-8')] != 0]

## GENERATE LIST FOR EACH SCHOOL
for index, row in df_school.iterrows():
	id_school = row['序号'.decode('utf-8')]
	str_df_school = "df_" + str(id_school)
	df_temp = pd.DataFrame(columns = ['Subject','Teacher'])
	str_sub1 = row['科目1'.decode('utf-8')]
	count_ppl1 = row['人数1'.decode('utf-8')]
	list_pos1 = [str_sub1] * count_ppl1
	globals()[str_df_school]['Subject'] = list_pos1


## EXPORT OUTPUT
writer = pd.ExcelWriter('Output.xlsx')
for s in df_school['序号'.decode('utf-8')]:
	df_exp = globals()[str_df_school]
	df_exp.to_excel(writer, str(s), index = False)
