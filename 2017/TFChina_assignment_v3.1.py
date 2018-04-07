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

## GLOBAL CONSTANT
quota_buffer = 0.15 # % of buffer for each ratio balance category
headcount_lb = 10 # lowerbound to enforce ratio balance (# headcount in a city)

## LOAD DATA
input_data = 'data_cleaned_v5.xlsx'
df_teacher = pd.read_excel(input_data, sheetname = 'Teachers', header = 0, na_values = "", encoding="gbk")
df_school = pd.read_excel(input_data, sheetname = 'Schools', header = 0, na_values = "", encoding="gbk")
df_svh = pd.read_excel('mapping table.xlsx', sheetname = 'science_v_humanity', header = 0, na_values = "", encoding="gbk")
df_english = pd.read_excel('mapping table.xlsx', sheetname = 'Englisth_competency', header = 0, na_values = "", encoding="gbk")
df_priority = pd.read_excel('mapping table.xlsx', sheetname = 'TF_China', header = 0, na_values = "", encoding="gbk")
writer = pd.ExcelWriter('Output_v0624.xlsx')

## DATA PREPARATION
df_school = df_school.sort_values(by = '地级市'.decode('utf-8')) # group by city (sort by count if needed)
list_city = df_school['地级市'.decode('utf-8')].unique()
cols_exp = ['项目学校'.decode('utf-8'),'科目'.decode('utf-8'),'老师'.decode('utf-8'),'科目优先级'.decode('utf-8')]
# df_teacher = df_teacher.sample(frac = 1).reset_index(drop=True) # shuffle teachers; to be removed
df_cand = df_teacher.copy()

# province
df_school['省份'.decode('utf-8')] = df_school['省份'.decode('utf-8')].apply(lambda x: x[:2]) # standardize province format

# science v. humanity
df_cand['sub_cat'] = df_cand['高中科目类型'.decode('utf-8')]
df_cand['sub_cat'] = df_cand['sub_cat'].replace('综合'.decode('utf-8'),np.nan)
df_cand['sub_cat'] = df_cand['sub_cat'].fillna(df_cand['本科专业所属类别'.decode('utf-8')])

# English competency
eng_comp = '英语能力分级'.decode('utf-8')
qual_in_data = '英语水平考试'.decode('utf-8')
df_cand[eng_comp] = [4] * len(df_cand)
for index, row in df_cand.iterrows():
	if (('专业八级'.decode('utf-8') in row[qual_in_data]) or ('专业四级'.decode('utf-8') in row[qual_in_data]) or ('GRE' in row[qual_in_data]) or ('SAT' in row[qual_in_data]) or ('专八'.decode('utf-8') in row[qual_in_data]) or ('专四'.decode('utf-8') in row[qual_in_data])):
		df_cand.loc[index, eng_comp] = 1
	elif (('雅思'.decode('utf-8') in row[qual_in_data]) or ('托福'.decode('utf-8') in row[qual_in_data])):
		df_cand.loc[index, eng_comp] = 2
	elif (('大学英语六级'.decode('utf-8') in row[qual_in_data]) or ('大学英语四级'.decode('utf-8') in row[qual_in_data])):
		df_cand.loc[index, eng_comp] = 3

# calculate all ratios
# gender
list_gender = list(df_cand['性别'.decode('utf-8')])
dict_gender = {x.decode('utf-8'):list_gender.count(x) for x in list_gender}
female_ratio = float(dict_gender['女士'.decode('utf-8')])/(sum(dict_gender.values()))
male_ratio = 1 - female_ratio
# college
ct_oversea = sum(list(df_cand['毕业院校: 海外'.decode('utf-8')]))
ct_985 = sum(list(df_cand['毕业院校: 985'.decode('utf-8')]))
ct_211 = sum(list(df_cand['毕业院校: 211'.decode('utf-8')]))
ct_total = len(df_cand)
ratio_985 = float(ct_985) / ct_total
ratio_oversea = float(ct_oversea) / ct_total
ratio_211 = 1 - ratio_985 - ratio_oversea
# normal
df_cand['normal'] = df_cand['本科毕业学校'.decode('utf-8')].str.contains('师范'.decode('utf-8'))
df_cand['normal'] = df_cand['normal'].replace(True, 1)
df_cand['normal'] = df_cand['normal'].replace(False, 0)
list_normal = list(df_cand['normal'])
dict_normal = {x:list_normal.count(x) for x in list_normal}
normal_ratio = float(dict_normal[1])/(sum(dict_normal.values()))
non_normal_ratio = 1 - normal_ratio

# specialty 
list_specialty = list(df_cand['17.除文化课知识外您有艺术、体育等方面的特长？'.decode('utf-8')])
dict_specialty = {x.decode('utf-8'):list_specialty.count(x) for x in list_specialty}
with_spec_ratio = float(dict_specialty['是'.decode('utf-8')])/(sum(dict_specialty.values()))
no_spec_ratio = 1 - with_spec_ratio

## GENERATE LIST FOR EACH SCHOOL

def ratio_balance_trigger(c, df_city_cand, df_sub_selected, dict_quota, dict_counter):
	teacher_gender = list(df_sub_selected['性别'.decode('utf-8')])[0]
	teacher_985 = list(df_sub_selected['毕业院校: 985'.decode('utf-8')])[0]
	teacher_211 = list(df_sub_selected['毕业院校: 211'.decode('utf-8')])[0]
	teacher_normal = list(df_sub_selected['normal'])[0]
	teacher_spec = list(df_sub_selected['17.除文化课知识外您有艺术、体育等方面的特长？'.decode('utf-8')])[0]
	if teacher_gender == '女士'.decode('utf-8'):
		dict_counter['city_female_count'] += 1
		if dict_counter['city_female_count'] > dict_quota['quota_female']:
			df_city_cand = df_city_cand[df_city_cand['性别'.decode('utf-8')] != '女士'.decode('utf-8')]
			print('Gender constraint triggered (need more male) - City: ' + c.decode('utf-8'))
	if teacher_985 == 1:
		dict_counter['city_985_count'] += 1
		if dict_counter['city_985_count'] > dict_quota['quota_985']:
			df_city_cand = df_city_cand[df_city_cand['毕业院校: 985'.decode('utf-8')] != 1]
			print('985 constraint triggered (need more oversea) - City: ' + c.decode('utf-8'))
	if teacher_211 == 1:
		dict_counter['city_211_count'] += 1
		if dict_counter['city_211_count'] > dict_quota['quota_211']:
			df_city_cand = df_city_cand[df_city_cand['毕业院校: 211'.decode('utf-8')] != 1]
			print('211 constraint triggered (need more oversea) - City: ' + c.decode('utf-8'))
	if teacher_normal == 0:
		dict_counter['city_normal_count'] += 1
		if dict_counter['city_normal_count'] > dict_quota['quota_normal']:
			df_city_cand = df_city_cand[df_city_cand['normal'] != 0]
			print('Normal constraint triggered (need more normal) - City: ' + c.decode('utf-8'))
	if teacher_spec == '是'.decode('utf-8'):
		dict_counter['city_spec_count'] += 1
		if dict_counter['city_spec_count'] > dict_quota['quota_spec']:
			df_city_cand = df_city_cand[df_city_cand['17.除文化课知识外您有艺术、体育等方面的特长？'.decode('utf-8')] != '是'.decode('utf-8')]
			print('Specialty constraint triggered (need more non-specialty) - City: ' + c.decode('utf-8'))
	return df_city_cand, dict_counter

df_exp = pd.DataFrame(columns = cols_exp)
list_city = df_school['地级市'.decode('utf-8')].unique()
df_school = pd.merge(df_school, df_priority, how = 'left', left_on = '具体派遣方式说明'.decode('utf-8'), right_on = '具体派遣方式说明'.decode('utf-8'))
for c in list_city:
	df_city_school = df_school[df_school['地级市'.decode('utf-8')] == c]
	df_city_school = df_city_school.sort_values(by = ['美丽中国班优先级'.decode('utf-8')], ascending = [False]) # prioritize schools
	city_total_headcount = sum([(list(df_city_school['人数'.decode('utf-8') + str(p)])).count(1) for p in range(1, 9)])
	df_city_cand = df_cand.copy()
	if city_total_headcount >= headcount_lb:
		# enforce limit on majority
		quota_female = int(female_ratio * city_total_headcount * (1 + quota_buffer))
		quota_211 = int(ratio_211 * city_total_headcount * (1 + quota_buffer))
		quota_985 = int(ratio_985 * city_total_headcount * (1 + quota_buffer))
		quota_normal = int(non_normal_ratio * city_total_headcount * (1 + quota_buffer))
		quota_spec = int(with_spec_ratio * city_total_headcount * (1 + quota_buffer))
	else:
		quota_female = city_total_headcount
		quota_211 = city_total_headcount
		quota_985 = city_total_headcount
		quota_normal = city_total_headcount
		quota_spec = city_total_headcount
	dict_quota = {'quota_female': quota_female, 'quota_985': quota_985, 'quota_211': quota_211, 'quota_normal': quota_normal, 'quota_spec': quota_spec}
	dict_counter = {'city_female_count': 0, 'city_985_count': 0, 'city_211_count': 0, 'city_normal_count': 0, 'city_spec_count': 0}
	for index, row in df_city_school.iterrows():
		id_school = row['序号'.decode('utf-8')]
		str_df_school = "df_" + str(id_school)
		df_s = pd.DataFrame(columns = cols_exp)
		for p in range(1,9):
			# up to 8 subject requests
			df_temp = pd.DataFrame(columns = cols_exp)
			sub = '科目'.decode('utf-8') + str(p)
			ppl = '人数'.decode('utf-8') + str(p)
			str_sub = row[sub]
			count_ppl = row[ppl]
			if count_ppl > 0:
				count_ppl = int(count_ppl)
				list_pos = [str_sub] * count_ppl
				list_pri = [p] * count_ppl
				df_temp['科目'.decode('utf-8')] = list_pos
				df_temp['科目优先级'.decode('utf-8')] = list_pri
				df_temp['项目学校'.decode('utf-8')] = [row['项目学校'.decode('utf-8')]] * len(df_temp)
			df_s = pd.concat([df_s, df_temp])
		if len(df_s) > 0:
			df_s['地级市'.decode('utf-8')] = [c] * len(df_s)
			df_s['主科目'.decode('utf-8')] = df_s['科目'.decode('utf-8')].apply(lambda x: x[:2])
			df_s = pd.merge(df_s, df_svh, how = 'left', left_on = '主科目'.decode('utf-8'), right_on = '主科目'.decode('utf-8'))
			df_s = df_s.sort_values(by = '科目优先级'.decode('utf-8'), ascending = True)
			df_s['老师'.decode('utf-8')] = [''] * len(df_s)
			req_ppl = len(df_s)
			if req_ppl > 0:
				df_selected = pd.DataFrame()
				name_school = list(df_s['项目学校'.decode('utf-8')])[0].decode('utf-8')
				df_target_school = df_school[df_school['项目学校'.decode('utf-8')] == name_school]
				## rule 3: medical attention
				school_med_condition = list(df_target_school['是否能够接受有病史的项目老师'.decode('utf-8')])[0]
				if school_med_condition == 0:
					df_city_cand = df_city_cand[pd.isnull(df_city_cand['过往病史和就医需求'.decode('utf-8')])]
				## rule 5: requirement on province
				school_province = list(df_target_school['省份'.decode('utf-8')])[0]
				df_city_cand = df_city_cand[(df_city_cand['选择地区'.decode('utf-8')] == school_province)| (df_city_cand['选择地区'.decode('utf-8')] == '都可以'.decode('utf-8'))]
				## rule 6-7: science vs. humanity - core matching
				for index, row in df_s.iterrows():
					if row['文理科'.decode('utf-8')] == '文科'.decode('utf-8'):
						df_sub_pool = df_city_cand[df_city_cand['sub_cat'] == '文科'.decode('utf-8')]
						if len(df_sub_pool) > 0:
							df_sub_selected = df_sub_pool.iloc[[0]]
							selected_teacher_name = list(df_sub_selected['姓名'.decode('utf-8')])[0]
							df_s.loc[index, '老师'.decode('utf-8')] = selected_teacher_name
							df_cand = df_cand[df_cand['姓名'.decode('utf-8')] != selected_teacher_name]
							df_city_cand = df_city_cand[df_city_cand['姓名'.decode('utf-8')] != selected_teacher_name]
							df_city_cand, dict_counter = ratio_balance_trigger(c, df_city_cand, df_sub_selected, dict_quota, dict_counter)
					if row['文理科'.decode('utf-8')] == '理科'.decode('utf-8'):
						df_sub_pool = df_city_cand[df_city_cand['sub_cat'] == '理科'.decode('utf-8')]
						if len(df_sub_pool) > 0:
							df_sub_selected = df_sub_pool.iloc[[0]]
							selected_teacher_name = list(df_sub_selected['姓名'.decode('utf-8')])[0]
							df_s.loc[index, '老师'.decode('utf-8')] = selected_teacher_name
							df_cand = df_cand[df_cand['姓名'.decode('utf-8')] != selected_teacher_name]
							df_city_cand = df_city_cand[df_city_cand['姓名'.decode('utf-8')] != selected_teacher_name]
							df_city_cand, dict_counter = ratio_balance_trigger(c, df_city_cand, df_sub_selected, dict_quota, dict_counter)
					if row['文理科'.decode('utf-8')] == '特长'.decode('utf-8'):
						df_sub_pool = df_city_cand[df_city_cand['17.除文化课知识外您有艺术、体育等方面的特长？'.decode('utf-8')] == '是'.decode('utf-8')] # potentially add another source of specialty
						if len(df_sub_pool) > 0:
							df_sub_selected = df_sub_pool.iloc[[0]]
							selected_teacher_name = list(df_sub_selected['姓名'.decode('utf-8')])[0]
							df_s.loc[index, '老师'.decode('utf-8')] = selected_teacher_name
							df_cand = df_cand[df_cand['姓名'.decode('utf-8')] != selected_teacher_name]
							df_city_cand = df_city_cand[df_city_cand['姓名'.decode('utf-8')] != selected_teacher_name]
							df_city_cand, dict_counter = ratio_balance_trigger(c, df_city_cand, df_sub_selected, dict_quota, dict_counter)
					if row['文理科'.decode('utf-8')] == '英语'.decode('utf-8'):
						df_sub_pool = df_city_cand[df_city_cand[eng_comp] < 4]
						if len(df_sub_pool) > 0:
							df_sub_pool = df_sub_pool.sort_values(by = eng_comp, ascending = True) # sort candidates by English competency
							df_sub_selected = df_sub_pool.iloc[[0]]
							selected_teacher_name = list(df_sub_selected['姓名'.decode('utf-8')])[0]
							df_s.loc[index, '老师'.decode('utf-8')] = selected_teacher_name
							df_cand = df_cand[df_cand['姓名'.decode('utf-8')] != selected_teacher_name]
							df_city_cand = df_city_cand[df_city_cand['姓名'.decode('utf-8')] != selected_teacher_name]
							df_city_cand, dict_counter = ratio_balance_trigger(c, df_city_cand, df_sub_selected, dict_quota, dict_counter)
				## add more school level matching criteria here
		df_exp = pd.concat([df_exp, df_s])
	print(c.decode('utf-8'))
	print(dict_quota)
	print(dict_counter)

cols_school = ['项目学校'.decode('utf-8'), '省份'.decode('utf-8'),'是否能够接受有病史的项目老师'.decode('utf-8'),'是否至少匹配一名男生'.decode('utf-8')]
df_exp = pd.merge(df_exp, df_school[cols_school], how = 'left', left_on = '项目学校'.decode('utf-8'), right_on = '项目学校'.decode('utf-8'))

df_exp.to_excel(writer,'Output', index = False)


df_remain = df_teacher[~df_teacher['姓名'.decode('utf-8')].isin(df_exp['老师'.decode('utf-8')].unique())]
del df_remain['序号'.decode('utf-8')]
df_remain = df_remain.drop_duplicates()

if len(df_remain) > 0:
	print('Warning: Not all teachers are assigned - {0} candidates available'.format(len(df_remain)))
	df_remain.to_excel(writer, 'Unassigned', index = False)

writer.save()