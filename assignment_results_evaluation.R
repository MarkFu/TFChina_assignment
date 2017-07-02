# Setup
# ----------------------------------------

Required_Packages <- c("xlsx", "xtable", "zoo", "lubridate", "ggplot2", "gridExtra", "gdata", "brew", "reshape")
Remaining_Packages <- Required_Packages[!(Required_Packages %in% installed.packages()[,"Package"])]

if (length(Remaining_Packages)) install.packages(Remaining_Packages, repos='http://cran.us.r-project.org')
for(package_name in Required_Packages) suppressMessages(library(package_name,character.only=TRUE,quietly=TRUE))

source("code/aux_functions.R")

# Data Preparation
# ----------------------------------------

# Input data

teacher_data = read.xlsx2("data/test_run/data_cleaned_v4.xlsx", sheetIndex = 1)
teacher_data = apply(teacher_data, 2, trim) # trim off white spaces
colnames(teacher_data) = paste0("td_", colnames(teacher_data))

education_school_history = subset(teacher_data, select = colnames(teacher_data)[grep("学校", colnames(teacher_data))])
education_school_history = apply(education_school_history, 2, FUN = function(x) as.character(x))
teacher_data = data.frame(teacher_data)
teacher_data$has_attended_normal_school = 0 
teacher_data$has_attended_normal_school[grep("师范", 
                                             sapply(c(1:nrow(teacher_data)), 
                                                    FUN = function(i) paste0(education_school_history[i, ], collapse = "")
                                                           ))] = 1

teacher_data$prov_preference = 
  as.character(teacher_data$td_我认为我在下列哪个地区能发挥更大的影响力.并说明理由.)

teacher_data$td_我高中时期所学科类 = as.character(teacher_data$td_我高中时期所学科类)
teacher_data$td_本科专业所属类别. = as.character(teacher_data$td_本科专业所属类别.)
teacher_data$discipline = teacher_data$td_我高中时期所学科类
teacher_data$discipline[which(!teacher_data$td_我高中时期所学科类 %in% c("文科","理科"))] = 
  teacher_data$td_本科专业所属类别.[which(!teacher_data$td_我高中时期所学科类 %in% c("文科","理科"))]


teacher_data$english_skills 
# eng_comp = '英语能力分级'.decode('utf-8')
# qual_in_data = '英语水平考试及成绩：'.decode('utf-8')
# df_cand[eng_comp] = [4] * len(df_cand)
# for index, row in df_cand.iterrows():
#   if (('专业八级'.decode('utf-8') in row[qual_in_data]) or ('专业四级'.decode('utf-8') in row[qual_in_data]) or ('GRE' in row[qual_in_data]) or ('SAT' in row[qual_in_data]) or ('专八'.decode('utf-8') in row[qual_in_data]) or ('专四'.decode('utf-8') in row[qual_in_data])):
#   df_cand.loc[index, eng_comp] = 1
# elif (('雅思'.decode('utf-8') in row[qual_in_data]) or ('托福'.decode('utf-8') in row[qual_in_data])):
#   df_cand.loc[index, eng_comp] = 2
# elif (('大学英语六级'.decode('utf-8') in row[qual_in_data]) or ('大学英语四级'.decode('utf-8') in row[qual_in_data])):
#   df_cand.loc[index, eng_comp] = 3


school_data = read.xlsx2("data/test_run/data_cleaned_v4.xlsx", sheetIndex = 2)
school_data = apply(school_data, 2, trim)
colnames(school_data) = paste0("sd_", colnames(school_data))

past_assignments = read.xlsx2("data/test_run/data_cleaned_v4.xlsx", sheetIndex = 3)
past_assignments = apply(past_assignments, 2, trim)
colnames(past_assignments) = paste0("past_", colnames(past_assignments))

# Matching results

matching_results = read.xlsx2("data/test_run/匹配结果_0618.xlsx", sheetIndex = 1)
matching_results = apply(matching_results, 2, trim)

school_names = split_uneven_string(matching_results[, 3], sep.by = "/", max.length = 2)
colnames(school_names) = c("school_name_english", "school_name_chinese")
matching_results = cbind(matching_results, school_names)

matching_results_with_info = merge(matching_results, teacher_data, by.x = "老师", by.y = "td_全名", all.x = T)
matching_results_with_info = merge(matching_results_with_info, school_data, by.x = "学校名称", by.y = "sd_学校名称", all.x = T)

# Request Hierarchy (# 4)
# Tier 1: 美丽中国校美丽中国班，所有需求
# Tier 2: 美丽中国校非美丽中国班，所有需求
# Tier 3: 非美丽中国校，所有需求（先满足一个学校的所有需求，再满足下一个学校的所有需求）

matching_results_with_info$sd_美丽中国班数 = as.numeric(as.character(matching_results_with_info$sd_美丽中国班数))
matching_results_with_info$科目优先级 = as.numeric(as.character(matching_results_with_info$科目优先级))

matching_results_with_info$is_mlzg_school = ifelse(matching_results_with_info$sd_美丽中国班数 > 0, 1, 0)
matching_results_with_info$is_mlzg_class = ifelse(matching_results_with_info$科目优先级 %in% c(1:4), 1, 0)

matching_results_with_info$request_hierarchy = ifelse(matching_results_with_info$is_mlzg_class == 1, "tier_1", 
                                                      ifelse(matching_results_with_info$is_mlzg_school == 1, "tier_2", "tier_3"))

# ----------------------------------------
# Diagnostics
# ----------------------------------------

# Overall fullfillment rate

matching_results_with_info$matched = ifelse(matching_results_with_info$老师 == "", 0, 1)

table(matching_results_with_info$request_hierarchy) # Number of teachers requested, by priority level

aggregate(matched ~ request_hierarchy, FUN = mean, data = matching_results_with_info)

matching_results_with_info$地级市 = as.character(matching_results_with_info$地级市)

fullfillment.district = NULL

for (d in unique(matching_results_with_info$地级市)){
  
  results_d = subset(matching_results_with_info, 地级市 == d)
  p.fullfilled = aggregate(matched ~ request_hierarchy, FUN = mean, data = results_d)

  fullfillment.district = rbind(fullfillment.district, data.frame(district = rep(d, nrow(p.fullfilled)), priority = p.fullfilled$request_hierarchy, p.fullfilled = p.fullfilled$matched))
}

# Matching Rules
# ----------------------------------------

## 1: Balanced distributions of the following factors at district level
## 1.1: Gender 
## 1.2.1 - 1.2.3: School Tiers (abroad/985/211)
## 1.3: Students with specialties 
## 1.4: Normal School Students

prop.table(table(subset(teacher_data, select = "td_性别"))) # 7:3

prop.table(table(matching_results_with_info$td_毕业院校..海外)) # 8%
prop.table(table(matching_results_with_info$td_毕业院校..985)) # 41%
prop.table(table(matching_results_with_info$td_毕业院校..211)) # 66%

prop.table(table(subset(teacher_data, select = "td_除文化课知识外我有艺术.体育.竞技等方面的爱好.特长."))) 

prop.table(table(subset(teacher_data, select = "has_attended_normal_school"))) 

gender.ratio.district = NULL
school.ratio.district = NULL
specialty.ratio.district = NULL
normal.ratio.district = NULL

for (d in unique(matching_results_with_info$地级市)){
  
  results_d = subset(matching_results_with_info, 地级市 == d & matched == 1)

  gender.district = prop.table(table(results_d$td_性别))
  gender.ratio.district = rbind(gender.ratio.district, 
                                data.frame(district = d, p.girl = gender.district[which(names(gender.district) == "女士")]))

  school.district = prop.table(table(results_d$td_毕业院校..海外))
  school.ratio.district = rbind(school.ratio.district, 
                                data.frame(district = d, type = "abroad", p = 1 - school.district[which(names(school.district) == "0")]))
  school.district = prop.table(table(results_d$td_毕业院校..985))
  school.ratio.district = rbind(school.ratio.district, 
                                data.frame(district = d, type = "985", p = 1 - school.district[which(names(school.district) == "0")]))
  school.district = prop.table(table(results_d$td_毕业院校..211))
  school.ratio.district = rbind(school.ratio.district, 
                                data.frame(district = d, type = "211", p = 1 - school.district[which(names(school.district) == "0")]))
  
  specialty.district = prop.table(table(results_d$td_除文化课知识外我有艺术.体育.竞技等方面的爱好.特长.))
  specialty.ratio.district = rbind(specialty.ratio.district, 
                                data.frame(district = d, p.specialty = specialty.district[which(names(specialty.district) == "是")]))
  
  normal.district = prop.table(table(results_d$has_attended_normal_school))
  normal.ratio.district = rbind(normal.ratio.district, 
                                data.frame(district = d, p.normal = 1 - normal.district[which(names(normal.district) == "0")]))
  
}

school.ratio.district = dcast(school.ratio.district, district ~ type, value.var = "p")
ratios.district = merge(gender.ratio.district, school.ratio.district, by = "district")
ratios.district = merge(ratios.district, specialty.ratio.district, by = "district")
ratios.district = merge(ratios.district, normal.ratio.district, by = "district")

## 2: Medical Requirement

matching_results_with_info$needs_medical_care = ifelse(matching_results_with_info$td_X..我有较严重的过往病史.. == "是", 1, 0)
table(matching_results_with_info$needs_medical_care)

with(subset(matching_results_with_info, needs_medical_care == 1), sd_就医条件) ## All "良好"

## 5: At least 1 boy every 2 years per school

## 6: Teacher's Preference for Destination Province

prop.table(table(teacher_data$prov_preference))
matching_results_with_info$sd_省 = as.character(matching_results_with_info$sd_省)
matching_results_with_info$prov_preference_met = unlist(lapply(c(1:nrow(matching_results_with_info)), 
                                                               function(x) max(grep(matching_results_with_info$prov_preference[x], matching_results_with_info$sd_省[x]), 0, na.rm = T)))

matching_results_with_info$prov_preference_met[matching_results_with_info$prov_preference == "都可以"] = 1
summary(with(subset(matching_results_with_info, matched == 1), prov_preference_met)) ## 99%

## 7: Science vs Liberal Arts
disc.table = table(matching_results_with_info$文理科, matching_results_with_info$discipline)
rownames(disc.table) = paste("项目", rownames(disc.table), sep = "")
disc.table[which(rownames(disc.table) %in% c("项目文科", "项目理科")), ]

## 8: English skills


## 9: PE/Music/Art 
ECA_requests = subset(matching_results_with_info, sd_是否有持续开展的音体美项目.1.是.2.否. == 1)



## 9: Soft Skills


## Miscellaneous 
## ---------------------------
# prov_preference = cbind(split_uneven_string(as.character(matching_results_16_with_info$td_我了解下列哪个地区的风俗文化.), 
#                                       max.length = 4, sep.by = "┋" )
#                         , split_uneven_string(as.character(matching_results_16_with_info$td_我会说.听得懂下列哪个地区的方言.), 
#                                               max.length = 4, sep.by = "┋" ))

# prov_preference = apply(prov_preference, 1, function(x) paste0(sort(setdiff(unique(x), c("都不了解", NA))), collapse = ","))
# matching_results_16_with_info = cbind(matching_results_16_with_info, prov_preference)
