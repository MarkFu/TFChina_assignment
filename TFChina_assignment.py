# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import os
import datetime as dt
import pandas as pd

df_teacher = pd.read_excel('data_cleaned.xlsx', sheetname = 'Teachers', header = 0, na_values = "", encoding="gbk")
df_school = pd.read_excel('data_cleaned.xlsx', sheetname = 'Schools', header = 0, na_values = "", encoding="gbk")

writer = pd.ExcelWriter('Output.xlsx')

## DATA PREPARATION
df_school = df_school[df_school['人数1'.decode('utf-8')] != 0]

## GENERATE LIST FOR EACH SCHOOL
for index, row in df_school.iterrows():
	id_school = row['序号'.decode('utf-8')]
	str_df_school = "df_" + str(id_school)
	df_temp = pd.DataFrame(columns = ['Subject','Teacher'])
	str_sub = row['科目1'.decode('utf-8')]
	count_ppl = row['人数1'.decode('utf-8')]
	list_pos = [str_sub] * count_ppl
	df_temp['Subject'] = list_pos
	## EXPORT OUTPUT
	df_temp.to_excel(writer, str(id_school), index = False)


