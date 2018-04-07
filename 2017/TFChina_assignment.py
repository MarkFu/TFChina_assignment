# -*- coding: utf-8 -*-
## Built in Python 2.7
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os
import datetime as dt
import pandas as pd

## LOAD DATA
input_data = 'data_cleaned_v2.xlsx'
df_teacher = pd.read_excel(input_data, sheetname = 'Teachers', header = 0, na_values = "", encoding="gbk")
df_school = pd.read_excel(input_data, sheetname = 'Schools', header = 0, na_values = "", encoding="gbk")
writer = pd.ExcelWriter('Output.xlsx')

## DATA PREPARATION
df_school = df_school.sort_values(by = '地级市'.decode('utf-8')) # group by city (sort by count if needed)
list_city = df_school['地级市'.decode('utf-8')].unique()
cols_exp = ['学校名称'.decode('utf-8'),'科目'.decode('utf-8'),'老师'.decode('utf-8'),'科目优先级'.decode('utf-8')]
df_teacher = df_teacher.sample(frac = 1).reset_index(drop=True) # shuffle teachers; to be removed
df_cand = df_teacher.copy()

# define function to match teacher to each 
def teacher_to_school(df_req, df_school, df_cand):
	req_ppl = len(df_req)
	if req_ppl > 0:
		name_school = list(df_req['学校名称'.decode('utf-8')])[0].decode('utf-8')
		df_target_school = df_school[df_school['学校名称'.decode('utf-8')] == name_school]
		## rule 1: medical attention
		school_med_condition = list(df_target_school['就医条件'.decode('utf-8')])[0]
		if school_med_condition == '一般'.decode('utf-8'):
			df_cand = df_cand[pd.isnull(df_cand['12、我的过往病史为：'.decode('utf-8')])]
		## add more school level matching criteria here
	df_selected = df_cand.head(n = req_ppl)
	list_selected_names = list(df_selected['全名'.decode('utf-8')])
	if len(list_selected_names) < req_ppl:
		# if not enough candidates, leave blank for now and wait for adjustment
		list_selected_names = list_selected_names + ['']*(req_ppl - len(list_selected_names))
	df_req['老师'.decode('utf-8')] = list_selected_names
	df_cand = df_cand[~df_cand['全名'.decode('utf-8')].isin(list_selected_names)]
	return df_req, df_cand

## PRIORITY SCHOOLS
df_sp = df_school[df_school['有无包班'.decode('utf-8')] == 1]
df_exp = pd.DataFrame(columns = cols_exp)
for index, row in df_sp.iterrows():
	for p in range(1,5):
		df_temp = pd.DataFrame(columns = cols_exp)
		sub = '科目'.decode('utf-8') + str(p)
		ppl = '人数'.decode('utf-8') + str(p)
		str_sub = row[sub]
		count_ppl = row[ppl]
		if count_ppl > 0:
			list_pos = [str_sub] * count_ppl
			list_pri = [p] * count_ppl
			df_temp['科目'.decode('utf-8')] = list_pos
			df_temp['科目优先级'.decode('utf-8')] = list_pri
		df_temp['学校名称'.decode('utf-8')] = [row['学校名称'.decode('utf-8')]] * len(df_temp)
		df_temp, df_cand = teacher_to_school(df_temp, df_school, df_cand)
		df_exp = pd.concat([df_exp, df_temp])

## GENERATE LIST FOR EACH SCHOOL
df_non_sp = df_school[df_school['有无包班'.decode('utf-8')] != 1]
list_city_non_sp = df_non_sp['地级市'.decode('utf-8')].unique()
for c in list_city_non_sp:
	df_city_school = df_non_sp[df_non_sp['地级市'.decode('utf-8')] == c]
	for index, row in df_city_school.iterrows():
		id_school = row['序号'.decode('utf-8')]
		str_df_school = "df_" + str(id_school)
		df_s = pd.DataFrame(columns = cols_exp)
		for p in range(1,5):
			df_temp = pd.DataFrame(columns = cols_exp)
			sub = '科目'.decode('utf-8') + str(p)
			ppl = '人数'.decode('utf-8') + str(p)
			str_sub = row[sub]
			count_ppl = row[ppl]
			if count_ppl > 0:
				list_pos = [str_sub] * count_ppl
				list_pri = [p] * count_ppl
				df_temp['科目'.decode('utf-8')] = list_pos
				df_temp['科目优先级'.decode('utf-8')] = list_pri
				df_temp['学校名称'.decode('utf-8')] = [row['学校名称'.decode('utf-8')]] * len(df_temp)
			df_s = pd.concat([df_s, df_temp])
		if len(df_s) > 0:
			df_s, df_cand = teacher_to_school(df_s, df_school, df_cand)
		df_exp = pd.concat([df_exp, df_s])

df_exp.to_excel(writer,'Output', index = False)