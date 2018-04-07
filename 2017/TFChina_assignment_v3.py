# -*- coding: utf-8 -*-
## Built in Python 2.7
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os
import datetime as dt
import pandas as pd
import numpy as np

## PENDING CHANGES
# change teacher name to ID, if possible

## LOAD DATA
input_data = 'data_cleaned_v4.xlsx'
df_teacher = pd.read_excel(input_data, sheetname = 'Teachers', header = 0, na_values = "", encoding="gbk")
df_school = pd.read_excel(input_data, sheetname = 'Schools', header = 0, na_values = "", encoding="gbk")
df_svh = pd.read_excel('mapping table.xlsx', sheetname = 'science_v_humanity', header = 0, na_values = "", encoding="gbk")
df_english = pd.read_excel('mapping table.xlsx', sheetname = 'Englisth_competency', header = 0, na_values = "", encoding="gbk")
writer = pd.ExcelWriter('Output.xlsx')

## DATA PREPARATION
df_school = df_school.sort_values(by = '地级市'.decode('utf-8')) # group by city (sort by count if needed)
list_city = df_school['地级市'.decode('utf-8')].unique()
cols_exp = ['学校名称'.decode('utf-8'),'科目'.decode('utf-8'),'老师'.decode('utf-8'),'科目优先级'.decode('utf-8')]
df_teacher = df_teacher.sample(frac = 1).reset_index(drop=True) # shuffle teachers; to be removed
df_cand = df_teacher.copy()

# province
df_school['省'.decode('utf-8')] = df_school['省'.decode('utf-8')].apply(lambda x: x[:2]) # standardize province format

# science v. humanity
df_cand['sub_cat'] = df_cand['我高中时期所学科类'.decode('utf-8')]
df_cand['sub_cat'] = df_cand['sub_cat'].replace('综合'.decode('utf-8'),np.nan)
df_cand['sub_cat'] = df_cand['sub_cat'].fillna(df_cand['本科专业所属类别：'.decode('utf-8')])

# English competency
eng_comp = '英语能力分级'.decode('utf-8')
qual_in_data = '英语水平考试及成绩：'.decode('utf-8')
df_cand[eng_comp] = [4] * len(df_cand)
for index, row in df_cand.iterrows():
	if (('专业八级'.decode('utf-8') in row[qual_in_data]) or ('专业四级'.decode('utf-8') in row[qual_in_data]) or ('GRE' in row[qual_in_data]) or ('SAT' in row[qual_in_data]) or ('专八'.decode('utf-8') in row[qual_in_data]) or ('专四'.decode('utf-8') in row[qual_in_data])):
		df_cand.loc[index, eng_comp] = 1
	elif (('雅思'.decode('utf-8') in row[qual_in_data]) or ('托福'.decode('utf-8') in row[qual_in_data])):
		df_cand.loc[index, eng_comp] = 2
	elif (('大学英语六级'.decode('utf-8') in row[qual_in_data]) or ('大学英语四级'.decode('utf-8') in row[qual_in_data])):
		df_cand.loc[index, eng_comp] = 3
	

# specialty - 3 types
col_PE = '体育'.decode('utf-8')
col_music = '音乐'.decode('utf-8')
col_arts = '美术'.decode('utf-8')
cols_spec = [col_PE, col_music, col_arts]
for col in cols_spec:
	df_cand[col] = [0] * len(df_cand)
	for index, row in df_cand.iterrows():
		if col in str(row['特长类型'.decode('utf-8')]):
			df_cand.loc[index, col] = 1

# calculate all ratios


'''
# define function to match teacher to each list of requests
def teacher_to_school(df_req, df_school, df_cand):
	req_ppl = len(df_req)
	if req_ppl > 0:
		df_selected = pd.DataFrame()
		name_school = list(df_req['学校名称'.decode('utf-8')])[0].decode('utf-8')
		df_target_school = df_school[df_school['学校名称'.decode('utf-8')] == name_school]
		## rule 1: medical attention
		school_med_condition = list(df_target_school['就医条件'.decode('utf-8')])[0]
		if school_med_condition == '一般'.decode('utf-8'):
			df_cand = df_cand[pd.isnull(df_cand['12、我的过往病史为：'.decode('utf-8')])]
		## rule 4: requirement on province
		school_province = list(df_target_school['省'.decode('utf-8')])[0]
		df_cand = df_cand[(df_cand['我认为我在下列哪个地区能发挥更大的影响力，并说明理由：'.decode('utf-8')] == school_province)| (df_cand['我认为我在下列哪个地区能发挥更大的影响力，并说明理由：'.decode('utf-8')] == '都可以'.decode('utf-8'))]
		## rule 5: science vs. humanity
		#for r in df_req['文理科']：
			#if r == '文科'.decode('utf-8'):

		## add more school level matching criteria here
	df_selected = df_cand.head(n = req_ppl)
	list_selected_names = list(df_selected['全名'.decode('utf-8')]) 
	if len(list_selected_names) < req_ppl:
		# if not enough candidates, leave blank for now and wait for adjustment
		list_selected_names = list_selected_names + ['']*(req_ppl - len(list_selected_names))
	df_req['老师'.decode('utf-8')] = list_selected_names
	df_cand = df_cand[~df_cand['全名'.decode('utf-8')].isin(list_selected_names)]
	return df_req, df_cand

## GENERATE LIST FOR EACH SCHOOL
df_exp = pd.DataFrame(columns = cols_exp)
list_city = df_school['地级市'.decode('utf-8')].unique()
for c in list_city:
	df_city_school = df_school[df_school['地级市'.decode('utf-8')] == c]
	df_city_school = df_city_school.sort_values(by = ['是否有持续开展的音体美项目（1-是；2-否）'.decode('utf-8'), '美丽中国班数'.decode('utf-8')], ascending = [False, False]) # prioritize schools
	for index, row in df_city_school.iterrows():
		id_school = row['序号'.decode('utf-8')]
		str_df_school = "df_" + str(id_school)
		df_s = pd.DataFrame(columns = cols_exp)
		for p in range(1,9):
			# up to 4 subject requests
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
			df_s['学校优先级'.decode('utf-8')] = [row['学校优先级'.decode('utf-8')]] * len(df_s)
			df_s['地级市'.decode('utf-8')] = [c] * len(df_s)
			df_s = pd.merge(df_s, df_svh, how = 'left', left_on = '科目'.decode('utf-8'), right_on = '科目'.decode('utf-8'))
			df_s = df_s.sort_values(by = '科目优先级'.decode('utf-8'), ascending = True)
			df_s, df_cand = teacher_to_school(df_s, df_school, df_cand)
		df_exp = pd.concat([df_exp, df_s])

df_exp.to_excel(writer,'Output', index = False)
'''