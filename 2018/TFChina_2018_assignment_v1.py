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
dir_input_file = 'input_cleaned.xlsx'
dir_mapping_file = 'mapping table.xlsx'
dir_output_file = '分配结果'.decode('utf-8') + '.xlsx'

######################################################################################################
## LOAD DATA
######################################################################################################
df_teacher = pd.read_excel(dir_input_file, sheetname = 'List', header = 0, na_values = "", encoding="gbk")
df_fixed = pd.read_excel(dir_input_file, sheetname = 'Fixed', header = 0, na_values = "", encoding="gbk")
df_prov_ratio = pd.read_excel(dir_mapping_file, sheetname = 'Province_ratio', header = 0, na_values = "", encoding="gbk")

######################################################################################################
## DATA PREPARATION
######################################################################################################

## create global column names (string)
###################################################
## df_teacher
str_colname_ID = 'ID'
str_colname_name = '1.姓名'.decode('utf-8')
str_colname_gender = '2.性别'.decode('utf-8')
str_colname_is_985211 = '毕业院校：985/211'.decode('utf-8')
str_colname_is_foreign = '毕业院校: 海外'.decode('utf-8')
str_colname_prov_preference = '28.您认为您在下列哪个地区能发挥更大的影响力？'.decode('utf-8')
str_colname_spouse_in_TFC = '26.您的丈夫/妻子是否也是美丽中国的项目老师？'.decode('utf-8')
str_colname_spouse_name = '您丈夫/妻子的姓名是：'.decode('utf-8')
str_colname_bfgf_in_TFC = '您的男/女朋友是否也是美丽中国的项目老师？'.decode('utf-8')
str_colname_bfgf_name = '您男/女朋友的姓名是：'.decode('utf-8')
str_colname_medical_history = '27.您是否有较严重的过往病史？'.decode('utf-8')
str_colname_english_test = '30.您已经完成下列哪类英语水平考试？'.decode('utf-8')
str_colname_CET4_score = '大学英语四级'.decode('utf-8')
str_colname_CET6_score = '大学英语六级'.decode('utf-8')
str_colname_degree_category = '15.您高中时期所学科目属于：'.decode('utf-8')

## df_fixed
str_colname_fixed_pref = '项目老师个人意愿'.decode('utf-8')

## df_prov_ratio
str_colname_prov_name = '省'.decode('utf-8')
str_colname_prov_ppl_ratio = '人数比例'.decode('utf-8')

## create placeholder dataframes
###################################################
## copy of teacher dataframe
df_cand = df_teacher.copy()
## copy of ID-name mapping
df_id_name_mapping = df_cand[[str_colname_ID,str_colname_name]]

## calculate assignment criteria
###################################################
## total number of teachers
num_count_total_teachers = len(df_cand)

## ratio： has_condition (i.e. has medical history)
num_count_has_condition = len(df_cand[df_cand[str_colname_medical_history] == '是'.decode('utf-8')])
num_frac_has_condition = num_count_has_condition/float(num_count_total_teachers)

## ratio: is_foreign
num_count_is_foreign = df_cand[str_colname_is_foreign].sum()
num_frac_is_foreign = num_count_is_foreign/float(num_count_total_teachers)

## ratio: is_985211
num_count_is_985211 = df_cand[str_colname_is_985211].sum()
num_frac_is_985211 = num_count_is_985211/float(num_count_total_teachers)

## ratio: is_male
num_count_is_male = len(df_cand[df_cand[str_colname_gender] == '男士'.decode('utf-8')])
num_frac_is_male = num_count_is_male/float(num_count_total_teachers)

## ratio: has_comb_degree
num_count_has_comb_degree = len(df_cand[df_cand[str_colname_degree_category] == '综合'.decode('utf-8')]) # need to manually check this column in input data
num_frac_has_comb_degree = num_count_has_comb_degree/float(num_count_total_teachers)

## ratio: has_science_degree
num_count_has_science_degree = len(df_cand[df_cand[str_colname_degree_category] == '理科'.decode('utf-8')]) # need to manually check this column in input data
num_frac_has_science_degree = num_count_has_science_degree/float(num_count_total_teachers)

## ratio: can_teach_English
df_cand['can_teach_English'] = [False] * len(df_cand) # create placeholder column
df_cand[str_colname_CET4_score] = df_cand[str_colname_CET4_score].fillna(0)
df_cand[str_colname_CET6_score] = df_cand[str_colname_CET6_score].fillna(0)
for idx, row in df_cand.iterrows():
	if (
		('专业八级'.decode('utf-8') in row[str_colname_english_test]) 
		or ('专业四级'.decode('utf-8') in row[str_colname_english_test]) 
		or ('GRE' in row[str_colname_english_test]) 
		or ('SAT' in row[str_colname_english_test]) 
		or ('专八'.decode('utf-8') in row[str_colname_english_test]) 
		or ('专四'.decode('utf-8') in row[str_colname_english_test])
		or ('雅思'.decode('utf-8') in row[str_colname_english_test]) 
		or ('托福'.decode('utf-8') in row[str_colname_english_test])
		or (row[str_colname_CET4_score] > 550)
		or (row[str_colname_CET6_score] > 500)
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
	'is_foreign',
	'is_985211',
	'is_male',
	'has_comb_degree',
	'has_science_degree',
	'can_teach_English'
]

## calculate each criteria's corresponding headcount (quota)
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
	
	# is_foreign
	if df_top_cand.loc[0, str_colname_is_foreign] == 1:
		dict_counter['is_foreign'] += 1
	
	# is_985211
	if df_top_cand.loc[0, str_colname_is_985211] == 1:
		dict_counter['is_985211'] += 1
	
	# is_male
	if df_top_cand.loc[0, str_colname_gender] == '男士'.decode('utf-8'):
		dict_counter['is_male'] += 1
	
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
	
	# is_foreign
	if dict_counter['is_foreign'] >= dict_quota['is_foreign']:
		df_cand_prov = df_cand_prov[df_cand_prov[str_colname_is_foreign] != 1]
	
	# is_985211
	if dict_counter['is_985211'] >= dict_quota['is_985211']:
		df_cand_prov = df_cand_prov[df_cand_prov[str_colname_is_985211] != 1]
	
	# is_male
	if dict_counter['is_male'] >= dict_quota['is_male']:
		df_cand_prov = df_cand_prov[df_cand_prov[str_colname_gender] != '男士'.decode('utf-8')]
	
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

## create placeholder dictionaries
dict_quota = {}
dict_counter = {}
list_prov_assigned_ID = {}
dict_num_fixed = {}

## assign fixed candidates
for p in list(df_prov_ratio[str_colname_prov_name]): # loop through each province
	## setup
	df_cand_prov = df_cand.copy() # create local copy of candidate list for this province only
	list_prov_assigned_ID[p] = list() # create placeholder list for assigned candidates to this province
	df_prov_ratio_sub = df_prov_ratio.loc[df_prov_ratio[str_colname_prov_name] == p] # subset of quota for this province only
	df_prov_ratio_sub = df_prov_ratio_sub.reset_index()
	num_prov_quota = df_prov_ratio.loc[df_prov_ratio[str_colname_prov_name] == p, 'headcount_province_total'] # total headcount for the province
	
	## create dictionaries to keep track of all criteria
	dict_quota[p] = {c: int(df_prov_ratio_sub.loc[0, 'headcount_' + c]) for c in list_criteria} # create dictionary for quota
	dict_counter[p] = {c: 0 for c in list_criteria} # create dictionary of counter for each criteria (initiate with 0)

	## find fixed candidates
	list_fixed_ID = list(df_fixed[df_fixed[str_colname_fixed_pref] == p][str_colname_ID])
	df_fixed_prov = df_cand_prov[df_cand_prov[str_colname_ID].isin(list_fixed_ID)]

	## assign fixed candidate
	for r in range(len(df_fixed_prov)):
		df_top_cand = df_fixed_prov.iloc[[r]]
		df_top_cand = df_top_cand.reset_index()
		dict_counter[p] = counter_updater(df_top_cand, dict_counter[p])
		num_cand_ID = df_top_cand.loc[0, str_colname_ID] # extract ID
		## assign to province
		list_prov_assigned_ID[p].append(num_cand_ID)

	## remove fixed candidates from eligible pool
	df_cand = df_cand[~df_cand[str_colname_ID].isin(list_fixed_ID)]

	## record number of fixed candidates
	dict_num_fixed[p] = len(df_fixed_prov)

## assign all remaining candidates
for p in list(df_prov_ratio[str_colname_prov_name]): # loop through each province
	## set up
	df_cand_prov = df_cand.copy()
	df_prov_ratio_sub = df_prov_ratio.loc[df_prov_ratio[str_colname_prov_name] == p] # subset of quota for this province only
	df_prov_ratio_sub = df_prov_ratio_sub.reset_index()
	num_prov_quota = df_prov_ratio.loc[df_prov_ratio[str_colname_prov_name] == p, 'headcount_province_total'] # total headcount for the province
	print(p + ': ' + str(num_prov_quota) + " teachers needed")

	## update candidate pool if quota is met
	df_cand_prov = candidate_updater(dict_counter[p], dict_quota[p], df_cand_prov)

	## sort by province preference
	df_cand_prov['sort_prov'] = df_cand_prov[str_colname_prov_preference].map({
		p: 1,
		'都可以'.decode('utf-8'): 2
	})
	df_cand_prov['sort_prov'] = df_cand_prov['sort_prov'].fillna(3)
	df_cand_prov = df_cand_prov.sort_values(by = 'sort_prov', ascending = True) # sort candidates by preference
	
	## assign top candidate to the province
	while (not df_cand_prov.empty) and (len(list_prov_assigned_ID[p]) < int(num_prov_quota)):
		## extract information of top candidate
		df_top_cand = df_cand_prov.iloc[[0]]
		df_top_cand = df_top_cand.reset_index()
		num_cand_ID = df_top_cand.loc[0, str_colname_ID] # extract ID
		## assign to province
		list_prov_assigned_ID[p].append(num_cand_ID)
		## update province criteria counter
		dict_counter[p] = counter_updater(df_top_cand, dict_counter[p])
		## update candidate pool if quota is met
		df_cand_prov = candidate_updater(dict_counter[p], dict_quota[p], df_cand_prov)
		## remove assigned person from candidate list
		df_cand_prov = df_cand_prov[df_cand_prov[str_colname_ID] != num_cand_ID]
		df_cand = df_cand[df_cand[str_colname_ID] != num_cand_ID]
		## assign spouse if applicable
		if num_cand_ID in list(df_spouse[str_colname_ID]):
			num_spouse_ID = df_spouse.loc[df_spouse[str_colname_ID] == num_cand_ID, 'spouse_ID']
			num_spouse_ID = num_spouse_ID.tolist()[0]
			if num_spouse_ID in list(df_cand_prov[str_colname_ID]):
				df_spouse_cand = df_cand_prov[df_cand_prov[str_colname_ID] == num_spouse_ID]
				df_spouse_cand = df_spouse_cand.reset_index()
				list_prov_assigned_ID[p].append(num_spouse_ID) ## assign to province
				dict_counter[p] = counter_updater(df_spouse_cand, dict_counter[p]) ## update province criteria counter
				df_cand_prov = candidate_updater(dict_counter[p], dict_quota[p], df_cand_prov) ## update candidate pool if quota is met
				df_cand_prov = df_cand_prov[df_cand_prov[str_colname_ID] != num_spouse_ID] ## remove assigned person from candidate list
				df_cand = df_cand[df_cand[str_colname_ID] != num_spouse_ID]
	
	## generate assignment table for the province
	print(dict_quota[p])
	print(dict_counter[p])
	df_prov = df_teacher[df_teacher[str_colname_ID].isin(list_prov_assigned_ID[p])]
	df_prov.to_excel(writer, p, index = False)
	print(str(len(df_prov)) + " teachers assigned")

writer.save()
writer.close()

######################################################################################################
## Quality check
######################################################################################################
list_all_assigned_ID = []
for k in list_prov_assigned_ID:
	list_all_assigned_ID = list_all_assigned_ID + list_prov_assigned_ID[k]

print(len(list_all_assigned_ID))
print(len(df_teacher))

df_remaining = df_teacher[~df_teacher[str_colname_ID].isin(list_all_assigned_ID)]
df_remaining = df_remaining[~df_remaining[str_colname_ID].isin(df_fixed[str_colname_ID])]

writer_temp = pd.ExcelWriter('remaining.xlsx')
df_remaining.to_excel(writer_temp, 'full', index = False)
writer_temp.save()
writer_temp.close() 

######################################################################################################
## Export report
######################################################################################################
df_quota = pd.DataFrame()
for p in list(df_prov_ratio[str_colname_prov_name]): # loop through each province
	df_quota_p = pd.DataFrame(dict_quota[p], columns= dict_quota[p].keys(), index=[0])
	df_quota_p['Prov'] = [p]
	df_quota = pd.concat([df_quota, df_quota_p])

df_counter = pd.DataFrame()
for p in list(df_prov_ratio[str_colname_prov_name]): # loop through each province
	df_counter_p = pd.DataFrame(dict_counter[p], columns= dict_counter[p].keys(), index=[0])
	df_counter_p['Prov'] = [p]
	df_counter = pd.concat([df_counter, df_counter_p])

writer_report = pd.ExcelWriter('report.xlsx')
df_prov_ratio.to_excel(writer_report, 'headcount', index = False)
df_quota.to_excel(writer_report, 'quota', index = False)
df_counter.to_excel(writer_report, 'counter', index = False)
writer_report.save()
writer_report.close() 

'''
# for testing purpose only
writer_temp = pd.ExcelWriter('temp.xlsx')
df_cand.to_excel(writer_temp, 'full', index = False)
writer_temp.save()
writer_temp.close() 
exit()
'''