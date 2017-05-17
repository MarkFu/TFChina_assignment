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
df_school = df_school[df_school['科目1'.decode('utf-8')] != 0]

## GENERATE LIST FOR EACH SCHOOL
for index, row in df_school.iterrows():
	id_school = row['序号'.decode('utf-8')]
	str_df_school = "df_" + str(id_school)
	df_s = pd.DataFrame(columns = ['科目'.decode('utf-8'),'老师'.decode('utf-8'),'优先级'.decode('utf-8')])
	for p in range(1,5):
		df_temp = pd.DataFrame(columns = ['科目'.decode('utf-8'),'老师'.decode('utf-8'),'优先级'.decode('utf-8')])
		sub = '科目'.decode('utf-8') + str(p)
		ppl = '人数'.decode('utf-8') + str(p)
		str_sub = row[sub]
		count_ppl = row[ppl]
		if count_ppl > 0:
			list_pos = [str_sub] * count_ppl
			list_pri = [p] * count_ppl
			df_temp['科目'.decode('utf-8')] = list_pos
			df_temp['优先级'.decode('utf-8')] = list_pri
		df_s = pd.concat([df_s, df_temp])
	## EXPORT OUTPUT
	df_s.to_excel(writer, str(id_school), index = False)