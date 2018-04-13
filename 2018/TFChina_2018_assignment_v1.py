# -*- coding: utf-8 -*-
## Built in Python 2.7
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os
import datetime as dt
import pandas as pd
import numpy as np
pd.options.mode.chained_assignment = None

######################################################################################################
## VARIABLES
######################################################################################################
dir_input_file = '2017-2019_20170623104814_cleaned.xlsx'
dir_mapping_file = 'mapping table.xlsx'
dir_output_file = '分配结果'.decode('utf-8') + '.xlsx'

######################################################################################################
## LOAD DATA
######################################################################################################
df_teacher = pd.read_excel(dir_input_file, sheetname = 'Sheet1', header = 0, na_values = "", encoding="gbk")
df_prov_ratio = pd.read_excel(dir_mapping_file, sheetname = 'Province_ratio', header = 0, na_values = "", encoding="gbk")

######################################################################################################
## DATA PREPARATION
######################################################################################################

## create global column names (string)
###################################################
str_colname_ID = '序号'.decode('utf-8')
str_colname_name = '1.姓名'.decode('utf-8')
str_colname_gender = '性别'.decode('utf-8')
str_colname_prov_preference = '13.您认为您在下列哪个地区能发挥更大的影响力？'.decode('utf-8')
str_colname_spouse_in_TFC = '11.您的丈夫/妻子是否也是美丽中国的项目老师？'.decode('utf-8')
str_colname_spouse_name = '您丈夫/妻子的姓名是：'.decode('utf-8')
str_colname_bfgf_in_TFC = '您的男/女朋友是否也是美丽中国的项目老师？'.decode('utf-8')
str_colname_bfgf_name = '您男/女朋友的姓名是：'.decode('utf-8')
str_colname_medical_history = '12.您是否有较严重的过往病史？'.decode('utf-8')
str_colname_undergrad_school = '8.毕业学校（本科阶段）：'.decode('utf-8')
str_colname_english_test = '15.您已经完成下列哪类英语水平考试？'.decode('utf-8')
str_colname_english_score = '16.以上英语考试的成绩为：（例：大学英语四级：570；托福：90）'.decode('utf-8')
str_colname_degree_category = '6.您高中时期所学科目属于：'.decode('utf-8')

str_colname_prov_name = '省'.decode('utf-8')
str_colname_prov_ppl_ratio = '人数比例'.decode('utf-8')

## create placeholder dataframes
###################################################
# copy of teacher dataframe
df_cand = df_teacher.copy()
# copy of ID-name mapping
df_id_name_mapping = df_cand[[str_colname_ID,str_colname_name]]

## calculate assignment criteria
###################################################
## total number of teachers
num_count_total_teachers = len(df_cand)

## ratio： has_condition (i.e. has medical history)
num_count_has_condition = len(df_cand[df_cand[str_colname_medical_history] == '是'.decode('utf-8')])
num_frac_has_condition = num_count_has_condition/float(num_count_total_teachers)

## ratio: is_985211

## ratio: is_foreign

## ratio: is_male

## ratio: has_comb_degree
num_count_has_comb_degree = len(df_cand[df_cand[str_colname_degree_category] == '综合'.decode('utf-8')]) # need to manually check this column in input data
num_frac_has_comb_degree = num_count_has_comb_degree/float(num_count_total_teachers)

## ratio: has_science_degree
num_count_has_science_degree = len(df_cand[df_cand[str_colname_degree_category] == '理科'.decode('utf-8')]) # need to manually check this column in input data
num_frac_has_science_degree = num_count_has_science_degree/float(num_count_total_teachers)

## ratio: can_teach_English
df_cand['can_teach_English'] = [False] * len(df_cand) # create placeholder column
df_cand[str_colname_english_score] = df_cand[str_colname_english_score].apply(str)
for idx, row in df_cand.iterrows():
	str_score = row[str_colname_english_score].decode('utf-8').replace('：', ': ') # clean up strings
	str_score = str_score.replace('级', '级 ') # clen up strings
	list_num_score = [int(s) for s in str_score.split() if s.isdigit()] # extract list of integers in string
	if not list_num_score:
		num_max_score = 0 # if list is empty, then score is 0 (default)
	else:
		num_max_score = max(list_num_score) # if list is not empty, take max score
	if (
		('专业八级'.decode('utf-8') in row[str_colname_english_test]) 
		or ('专业四级'.decode('utf-8') in row[str_colname_english_test]) 
		or ('GRE' in row[str_colname_english_test]) 
		or ('SAT' in row[str_colname_english_test]) 
		or ('专八'.decode('utf-8') in row[str_colname_english_test]) 
		or ('专四'.decode('utf-8') in row[str_colname_english_test])
		or ('雅思'.decode('utf-8') in row[str_colname_english_test]) 
		or ('托福'.decode('utf-8') in row[str_colname_english_test])
		or (num_max_score >= 500)
	):
		df_cand.loc[idx, 'can_teach_English'] = True
num_count_can_teach_English = len(df_cand[df_cand['can_teach_English']])
num_frac_can_teach_English = num_count_can_teach_English/float(num_count_total_teachers)

## calculate each province's total headcount
df_prov_ratio['headcount_province_total'] = df_prov_ratio[str_colname_prov_ppl_ratio] * num_count_total_teachers
df_prov_ratio['headcount_province_total'] = df_prov_ratio['headcount_province_total'].apply(round)

## sort province by total headcount
df_prov_ratio = df_prov_ratio.sort_values(by = 'headcount_province_total', ascending = True)

## calculate each province's headcount/quota by criteria
list_criteria = [
	'has_condition',
	'has_comb_degree',
	'has_science_degree',
	'can_teach_English'
]

for c in list_criteria:
	str_colname_headcount = 'headcount_' + c
	str_criteria_frac = 'num_frac_' + c
	df_prov_ratio[str_colname_headcount] = df_prov_ratio['headcount_province_total'] * eval(str_criteria_frac)
	df_prov_ratio[str_colname_headcount] = df_prov_ratio[str_colname_headcount].apply(round)

## keep track of spouse/bfgf
###################################################
df_spouse = df_cand[
	(df_cand[str_colname_spouse_in_TFC] == '是'.decode('utf-8')) | (df_cand[str_colname_bfgf_in_TFC] == '是'.decode('utf-8'))
]
df_spouse['spouse_name'] = df_spouse[str_colname_spouse_name] # create spouse name column (by copying spouse name)
df_spouse['spouse_name'] = df_spouse['spouse_name'].fillna(df_spouse[str_colname_bfgf_name]) # fill null with bf/gf name
df_spouse_name_mapping = df_id_name_mapping.rename(columns = {
	str_colname_ID: 'spouse_ID',
	str_colname_name: 'spouse_name'
}) # rename columns in mapping table for merging purpose
df_spouse = pd.merge(df_spouse, df_spouse_name_mapping, how = 'left', on = 'spouse_name') # obtain spouse ID
df_spouse = df_spouse[pd.notnull(df_spouse['spouse_ID'])] # drop null spouse ID

######################################################################################################
## CORE ALGORITHM
######################################################################################################

writer = pd.ExcelWriter(dir_output_file) # initiate output file

def counter_updater(df_top_cand, dict_counter):
	# has_condition
	if df_top_cand.loc[0, str_colname_medical_history] == '是'.decode('utf-8'):
		dict_counter['has_condition'] += 1
	# is_985211
	# is_foreign
	# is_male
	# has_comb_degree
	if df_top_cand.loc[0, str_colname_degree_category] == '综合'.decode('utf-8'):
		dict_counter['has_comb_degree'] += 1

	# has_science_degree
	if df_top_cand.loc[0, str_colname_degree_category] == '理科'.decode('utf-8'):
		dict_counter['has_science_degree'] += 1

	# can_teach_English
	if df_top_cand.loc[0, 'can_teach_English']:
		dict_counter['can_teach_English'] += 1
	
	return dict_counter

def candidate_updater(dict_counter, dict_quota, df_cand_prov):
	# has_condition
	if dict_counter['has_condition'] >= dict_quota['has_condition']:
		df_cand_prov = df_cand_prov[df_cand_prov[str_colname_medical_history] != '是'.decode('utf-8')]
	# is_985211
	# is_foreign
	# is_male
	
	# has_comb_degree
	if dict_counter['has_comb_degree'] >= dict_quota['has_comb_degree']:
		df_cand_prov = df_cand_prov[df_cand_prov[str_colname_degree_category] != '综合'.decode('utf-8')]
	
	# has_science_degree
	if dict_counter['has_science_degree'] >= dict_quota['has_science_degree']:
		df_cand_prov = df_cand_prov[df_cand_prov[str_colname_degree_category] != '理科'.decode('utf-8')]
	
	# can_teach_English
	if dict_counter['can_teach_English'] >= dict_quota['can_teach_English']:
		df_cand_prov = df_cand_prov[~df_cand_prov['can_teach_English']]
	
	return df_cand_prov

for p in list(df_prov_ratio[str_colname_prov_name]): # loop through each province
	print(p + '...')
	# setup
	df_cand_prov = df_cand.copy() # create local copy of candidate list for this province only
	list_prov_assigned_ID = list() # create placeholder list for assigned candidates to this province
	df_prov_ratio_sub = df_prov_ratio.loc[df_prov_ratio[str_colname_prov_name] == p] # subset of quota for this province only
	df_prov_ratio_sub = df_prov_ratio_sub.reset_index()
	num_prov_quota = df_prov_ratio.loc[df_prov_ratio[str_colname_prov_name] == p, 'headcount_province_total'] # total headcount for the province
	
	# create dictionaries to keep track of all criteria
	dict_quota = {c: int(df_prov_ratio_sub.loc[0, 'headcount_' + c]) for c in list_criteria} # create dictionary for quota
	dict_counter = {c: 0 for c in list_criteria} # create dictionary of counter for each criteria (initiate with 0)
	print(dict_quota)

	# sort by province preference
	df_cand_prov['sort_prov'] = df_cand_prov[str_colname_prov_preference].map({
		p: 1,
		'都可以'.decode('utf-8'): 2
	})
	df_cand_prov['sort_prov'] = df_cand_prov['sort_prov'].fillna(3)
	df_cand_prov = df_cand_prov.sort_values(by = 'sort_prov', ascending = True) # sort candidates by preference
	
	# assign top candidate to the province
	while (not df_cand_prov.empty) and (len(list_prov_assigned_ID) <= int(num_prov_quota)):
		# extract information of top candidate
		df_top_cand = df_cand_prov.iloc[[0]]
		df_top_cand = df_top_cand.reset_index()
		num_cand_ID = df_top_cand.loc[0, str_colname_ID] # extract ID
		# assign to province
		list_prov_assigned_ID.append(num_cand_ID)
		# update province criteria counter
		dict_counter = counter_updater(df_top_cand, dict_counter)
		# update candidate pool if quota is met
		df_cand_prov = candidate_updater(dict_counter, dict_quota, df_cand_prov)
		# remove assigned person from candidate list
		df_cand_prov = df_cand_prov[df_cand_prov[str_colname_ID] != num_cand_ID]
		df_cand = df_cand[df_cand[str_colname_ID] != num_cand_ID]
	
	# generate assignment table for the province
	print(dict_counter)
	df_prov = df_teacher[df_teacher[str_colname_ID].isin(list_prov_assigned_ID)]
	df_prov.to_excel(writer, p, index = False)

writer.save()
writer.close()


'''
# for testing purpose only
writer_temp = pd.ExcelWriter('temp.xlsx')
df_cand.to_excel(writer_temp, 'full', index = False)
writer_temp.save()
writer_temp.close() 
exit()
'''