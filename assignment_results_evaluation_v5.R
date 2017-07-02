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

teacher_data = read.xlsx2("data/0701/data_cleaned_v9.xlsx", sheetIndex = 1)
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
summary(teacher_data$has_attended_normal_school)

teacher_data$prov_preference = as.character(teacher_data$td_选择地区)
table(teacher_data$prov_preference)

teacher_data$td_高中科目类型 = as.character(teacher_data$td_高中科目类型)
teacher_data$td_本科专业所属类别 = as.character(teacher_data$td_本科专业所属类别)
teacher_data$discipline = teacher_data$td_高中科目类型
teacher_data$discipline[which(!teacher_data$td_高中科目类型 %in% c("文科","理科"))] = 
  teacher_data$td_本科专业所属类别[which(!teacher_data$td_高中科目类型 %in% c("文科","理科"))]

table(teacher_data$discipline)

teacher_data$english_skills = split_uneven_string(as.character(teacher_data$td_英语水平考试), max.length = 5, 
                                                               sep.by = "，")

teacher_data$english_skills_level = 0 

teacher_data$english_skills_level[grep("四级|4级", teacher_data$td_英语水平考试)] = 4
teacher_data$english_skills_level[grep("六级|6级", teacher_data$td_英语水平考试)] = 3
teacher_data$english_skills_level[grep("雅思|托福|GRE|gre|TOEFL|toefl", teacher_data$td_英语水平考试)] = 2
teacher_data$english_skills_level[grep("专业", teacher_data$td_英语水平考试)] = 1

table(teacher_data$english_skills_level)

school_data = read.xlsx2("data/0701/data_cleaned_v9.xlsx", sheetIndex = 2)
school_data = data.frame(apply(school_data, 2, trim))
colnames(school_data) = paste0("sd_", colnames(school_data))

# Matching results

matching_results = read.xlsx2("data/0701/Output_v0630_v1.xlsx", sheetIndex = 1)
matching_results = data.frame(apply(matching_results, 2, trim))

# school_names = split_uneven_string(as.character(school_data[, 2]), sep.by = "/", max.length = 2)
# colnames(school_names) = c("school_name_english", "school_name_chinese")
# school_data = cbind(school_data, school_names)

matching_results$老师 = as.character(matching_results$老师)
teacher_data$td_姓名 = as.character(teacher_data$td_姓名)
matching_results$项目学校 = as.character(matching_results$项目学校)
school_data$sd_项目学校 = as.character(school_data$sd_项目学校)

matching_results_with_info = merge(matching_results, teacher_data, by.x = "老师", by.y = "td_姓名", all.x = T)
matching_results_with_info = merge(matching_results_with_info, school_data, by.x = "项目学校", by.y = "sd_项目学校", all.x = T)

# Request Hierarchy (# 4)
# Tier 1: 美丽中国校美丽中国班，所有需求
# Tier 2: 美丽中国校非美丽中国班，所有需求
# Tier 3: 非美丽中国校，所有需求（先满足一个学校的所有需求，再满足下一个学校的所有需求）

matching_results_with_info$sd_具体派遣方式说明 = trim(as.character(matching_results_with_info$sd_具体派遣方式说明))
matching_results_with_info$科目优先级 = as.numeric(as.character(matching_results_with_info$科目优先级))

matching_results_with_info$is_mlzg_school = ifelse(matching_results_with_info$sd_具体派遣方式说明 == "仅派遣非美丽中国班老师", 0, 1)
matching_results_with_info$is_mlzg_class = ifelse(matching_results_with_info$科目优先级 %in% c(1:4), 1, 0)

matching_results_with_info$request_hierarchy = ifelse(matching_results_with_info$is_mlzg_class == 1, "tier_1", 
                                                      ifelse(matching_results_with_info$is_mlzg_school == 1, "tier_2", "tier_3"))
table(matching_results_with_info$request_hierarchy)

# ----------------------------------------
# Diagnostics
# ----------------------------------------

# Overall fullfillment rate

matching_results_with_info$matched = ifelse(matching_results_with_info$老师 == "", 0, 1)

table(matching_results_with_info$request_hierarchy) # Number of teachers requested, by priority level

aggregate(matched ~ request_hierarchy, FUN = mean, data = matching_results_with_info)

matching_results_with_info$地级市 = as.character(matching_results_with_info$地级市)

fullfillment.district = NULL

for (d in unique(matching_results_with_info$省份)){
  
  results_d = subset(matching_results_with_info, 省份 == d)
  p.fullfilled = aggregate(matched ~ request_hierarchy, FUN = mean, data = results_d)

  fullfillment.district = rbind(fullfillment.district, data.frame(district = rep(d, nrow(p.fullfilled)), priority = p.fullfilled$request_hierarchy, p.fullfilled = p.fullfilled$matched))
}

for (d in unique(matching_results_with_info$地级市)){
  
  results_d = subset(matching_results_with_info, 地级市 == d)
  p.fullfilled = aggregate(matched ~ request_hierarchy, FUN = mean, data = results_d)
  
  fullfillment.district = rbind(fullfillment.district, data.frame(district = rep(d, nrow(p.fullfilled)), priority = p.fullfilled$request_hierarchy, p.fullfilled = p.fullfilled$matched))
}

fullfillment.district = dcast(fullfillment.district, district ~ priority, value.var = "p.fullfilled")

matching_results_with_info$count = 1

fullfillment.count.s = aggregate(cbind(count, matched) ~ 省份 + request_hierarchy, FUN = sum, data = matching_results_with_info)
fullfillment.count.d = aggregate(cbind(count, matched) ~ 地级市 + request_hierarchy, FUN = sum, data = matching_results_with_info)

colnames(fullfillment.count.s) = c("district", "priority", "count", "matched")
colnames(fullfillment.count.d) = c("district", "priority", "count", "matched")
fullfillment.count = rbind(fullfillment.count.s, fullfillment.count.d)

fullfillment.count = dcast(fullfillment.count, district ~ priority, value.var = "count")
fullfillment.district = dcast(fullfillment.district, district ~ priority, value.var = "p.fullfilled")

fullfillment = merge(fullfillment.count, fullfillment.district, by = "district")
colnames(fullfillment) = c("district", "tier_1_total_requests",  "tier_2_total_requests",  "tier_3_total_requests",
                           "tier_1_pct_matched", "tier_2_pct_matched",  "tier_3_pct_matched")
#fullfillment[, c(5:7)] = round(100*fullfillment[, c(5:7)], 1)

# Matching Rules
# ----------------------------------------

## 1: Balanced distributions of the following factors at district level
## 1.1: Gender 
## 1.2.1 - 1.2.3: School Tiers (abroad/985/211)
## 1.3: Students with specialties 
## 1.4: Normal School Students

global.gender = prop.table(table(subset(teacher_data, select = "td_性别"))) # 74:26
global.p.girl = global.gender[which(names(global.gender) == "女士")]

global.school.type = data.frame(
  rbind(prop.table(table(matching_results_with_info$td_毕业院校..海外)), # 11.9%
        prop.table(table(matching_results_with_info$td_毕业院校..985)), # 30.6%
        prop.table(table(matching_results_with_info$td_毕业院校..211)))) # 54.4%

colnames(global.school.type) = c("No", "Yes")
global.school.type$type = c("abroad", "985", "211")
global.p.abroad = global.school.type$Yes[1]
global.p.985 = global.school.type$Yes[2]
global.p.211 = global.school.type$Yes[3]

global.specialty = prop.table(table(!subset(teacher_data, select = "td_您有下列哪些方面的特长.") == "")) # 48.5%
global.p.speciaty = global.specialty[which(names(global.specialty) == "TRUE")]

global.nomarl = prop.table(table(subset(teacher_data, select = "has_attended_normal_school"))) # 15.8%
global.p.nomarl = global.nomarl[which(names(global.nomarl) == 1)]

gender.ratio.district = NULL
school.ratio.district = NULL
specialty.ratio.district = NULL
normal.ratio.district = NULL

for (d in unique(matching_results_with_info$省份)){
  
  results_d = subset(matching_results_with_info, 省份 == d & matched == 1)

  gender.district = prop.table(table(results_d$td_性别))
  gender.ratio.district = rbind(gender.ratio.district, 
                                data.frame(level = "省", district = d, p.girl = gender.district[which(names(gender.district) == "女士")]))

  school.district = prop.table(table(results_d$td_毕业院校..海外))
  school.ratio.district = rbind(school.ratio.district, 
                                data.frame(level = "省", district = d, type = "abroad", p = 1 - school.district[which(names(school.district) == "0")]))
  school.district = prop.table(table(results_d$td_毕业院校..985))
  school.ratio.district = rbind(school.ratio.district, 
                                data.frame(level = "省", district = d, type = "985", p = 1 - school.district[which(names(school.district) == "0")]))
  school.district = prop.table(table(results_d$td_毕业院校..211))
  school.ratio.district = rbind(school.ratio.district, 
                                data.frame(level = "省", district = d, type = "211", p = 1 - school.district[which(names(school.district) == "0")]))
  
  specialty.district = prop.table(table(!results_d$td_您有下列哪些方面的特长. == ""))
  specialty.ratio.district = rbind(specialty.ratio.district, 
                                data.frame(level = "省", district = d, p.specialty = specialty.district[which(names(specialty.district) == "TRUE")]))
  
  normal.district = prop.table(table(results_d$has_attended_normal_school))
  normal.ratio.district = rbind(normal.ratio.district, 
                                data.frame(level = "省", district = d, p.normal = 1 - normal.district[which(names(normal.district) == "0")]))
  
}

for (d in unique(matching_results_with_info$地级市)){
  
  results_d = subset(matching_results_with_info, 地级市 == d & matched == 1)
  
  gender.district = prop.table(table(results_d$td_性别))
  gender.ratio.district = rbind(gender.ratio.district, 
                                data.frame(level = "地级市", district = d, p.girl = gender.district[which(names(gender.district) == "女士")]))
  
  school.district = prop.table(table(results_d$td_毕业院校..海外))
  school.ratio.district = rbind(school.ratio.district, 
                                data.frame(level = "地级市", district = d, type = "abroad", p = 1 - school.district[which(names(school.district) == "0")]))
  school.district = prop.table(table(results_d$td_毕业院校..985))
  school.ratio.district = rbind(school.ratio.district, 
                                data.frame(level = "地级市", district = d, type = "985", p = 1 - school.district[which(names(school.district) == "0")]))
  school.district = prop.table(table(results_d$td_毕业院校..211))
  school.ratio.district = rbind(school.ratio.district, 
                                data.frame(level = "地级市", district = d, type = "211", p = 1 - school.district[which(names(school.district) == "0")]))
  
  specialty.district = prop.table(table(!results_d$td_您有下列哪些方面的特长. == ""))
  specialty.ratio.district = rbind(specialty.ratio.district, 
                                   data.frame(level = "地级市", district = d, p.specialty = specialty.district[which(names(specialty.district) == "TRUE")]))
  
  normal.district = prop.table(table(results_d$has_attended_normal_school))
  normal.ratio.district = rbind(normal.ratio.district, 
                                data.frame(level = "地级市", district = d, p.normal = 1 - normal.district[which(names(normal.district) == "0")]))
  
}

school.ratio.district = dcast(school.ratio.district, district ~ type, value.var = "p")
colnames(school.ratio.district)[2:4] = paste0("p.", colnames(school.ratio.district)[2:4])
ratios.district = merge(gender.ratio.district, school.ratio.district, by = "district")
ratios.district = merge(ratios.district, specialty.ratio.district, by = c("district", "level"))
ratios.district = merge(ratios.district, normal.ratio.district, by = c("district", "level"))

ratios.district = rbind(ratios.district, data.frame(district = "(所有报名者)", level = "", 
                                                    p.girl = global.p.girl, p.abroad = global.p.abroad, p.985 = global.p.985, p.211 = global.p.211, 
                                                    p.specialty = global.p.speciaty, p.normal = global.p.nomarl))

#ratios.district[, c(3:8)] = round(100*ratios.district[, c(3:8)], 1)

ratios.district = merge(ratios.district, fullfillment, by = "district", all.x = T)
ratios.district = ratios.district[order(ratios.district$level), ]

write.csv(ratios.district, "ratios_district_and_state.csv", row.names = F)

## 2: Medical Requirement

matching_results_with_info$needs_medical_care = matching_results_with_info$td_过往病史.1.是.0.否. 
table(matching_results_with_info$needs_medical_care)

medical_requirement_result = table(with(subset(matching_results_with_info, needs_medical_care == 1), sd_是否能够接受有病史的项目老师)) ## All "良好"

medical_requirement_result_string = paste0("有病史的老师总数: ", medical_requirement_result[which(names(medical_requirement_result) == 1)][[1]], '\n',
                                    "被分配到符合医疗条件的老师占总人数的比例: ", round(prop.table(medical_requirement_result)[which(names(prop.table(medical_requirement_result)) == 1)][[1]]*100, 1), "%")
## 5: At least 1 boy every 2 years per school

schools_with_at_least_one_boy = unique(with(subset(matching_results_with_info, td_性别 == "男士"), sd_序号))
schools_need_at_least_one_boy = unique(with(subset(school_data, sd_是否至少匹配一名男生 == 1), sd_序号))
schools_with_at_least_one_teacher = unique(with(subset(matching_results_with_info, matched == 1), sd_序号))

gender_requirement_result = table(schools_with_at_least_one_boy %in% schools_need_at_least_one_boy)

gender_requirement_result_string = paste0("需要至少一名男生的学校数: ", length(schools_need_at_least_one_boy), '\n',
                                          "分配到至少一名男生的学校数: ", length(schools_with_at_least_one_boy), '\n',
                                    "需要至少一名男生的学校中, 分配到至少一名男老师的学校数: ", gender_requirement_result[which(names(gender_requirement_result) == TRUE)])


## 6: Teacher's Preference for Destination Province

prop.table(table(teacher_data$prov_preference))
matching_results_with_info$sd_省 = as.character(matching_results_with_info$sd_省)
matching_results_with_info$prov_preference_met = unlist(lapply(c(1:nrow(matching_results_with_info)), 
                                                               function(x) max(grep(matching_results_with_info$prov_preference[x], matching_results_with_info$sd_省[x]), 0, na.rm = T)))

matching_results_with_info$prov_preference_met[matching_results_with_info$prov_preference == "都可以"] = 1
prov_preference_string = paste0("对省份的偏好得到满足的老师的占比: ", 
                                          round(mean(with(subset(matching_results_with_info, matched == 1), prov_preference_met))*100, 1), "%") ## 99%

## 7: Science vs Liberal Arts

disc.table = table(matching_results_with_info$文理科, matching_results_with_info$discipline)
rownames(disc.table) = paste("项目", rownames(disc.table), sep = "")
disc.table = disc.table[which(rownames(disc.table) %in% c("项目文科", "项目理科")), ]

to.check = matching_results_with_info[which(matching_results_with_info$文理科 != matching_results_with_info$discipline 
                                      & matching_results_with_info$文理科 != "英语"),]

## 8: English skills
matching_results_with_info$is_English = 0
matching_results_with_info$is_English[grep("英语", as.character(matching_results_with_info$科目))] = 1
English_sub = subset(matching_results_with_info, is_English == 1)
# English_qual_result = table(English_sub$english_skills_level)
English_qual_result = table(English_sub$english_skills_level > 0)
English_qual_result_string = paste0("英语能力符合要求的老师占所有英语授课老师的比例: ", 
  round(prop.table(English_qual_result)[which(names(English_qual_result) == TRUE)]*100, 1), "%")

## 9: PE/Music/Art
matching_results_with_info$sd_该校有几个持续到17.18学年开展的音体美项目 = as.numeric(as.character(matching_results_with_info$sd_该校有几个持续到17.18学年开展的音体美项目))
ECA_requests = subset(matching_results_with_info, sd_该校有几个持续到17.18学年开展的音体美项目 > 0)

ECA_requests_m = ECA_requests[grep("音乐", ECA_requests$科目), ]
ECA_requests_m_table_perfect = table(ECA_requests_m$td_特长类型.音乐.1.是.0.否.)
ECA_requests_m_table = table(ECA_requests_m$td_X17.除文化课知识外您有艺术.体育等方面的特长.)

ECA_requests_a = ECA_requests[grep("美术", ECA_requests$科目), ]
ECA_requests_a_table_perfect = table(ECA_requests_a$td_特长类型.美术.1.是.0.否.)
ECA_requests_a_table = table(ECA_requests_a$td_X17.除文化课知识外您有艺术.体育等方面的特长.)

ECA_requests_c = ECA_requests[grep("体育", ECA_requests$科目), ]
ECA_requests_c_table_perfect = table(ECA_requests_c$td_特长类型.体育.1.是.0.否.)
ECA_requests_c_table = table(ECA_requests_c$td_X17.除文化课知识外您有艺术.体育等方面的特长.)

perfect_table = rbind(ECA_requests_m_table_perfect, ECA_requests_a_table_perfect, ECA_requests_c_table_perfect)
general_table = rbind(ECA_requests_m_table, ECA_requests_a_table, ECA_requests_c_table)

# ECA_result_string = paste0("有几个持续到17.18学年开展的音体美项目的学校，音体美项目老师的特长匹配比例", 
#                                     round(prop.table(English_qual_result)[which(names(English_qual_result) == TRUE)]*100, 1), "%")


## Create EXL file

output = xlsx::createWorkbook(type="xlsx")

sheet1 = xlsx::createSheet(output, sheetName = "ratios")
xlsx::addDataFrame(ratios.district, sheet1, startRow = 1, startColumn = 1, 
                   row.names = F )

sheet2 = xlsx::createSheet(output, sheetName = "other_rules")
strings = paste(medical_requirement_result_string, gender_requirement_result_string, prov_preference_string, English_qual_result_string, sep = "\n")
xlsx::addDataFrame(strings, sheet2, startRow = 1, startColumn = 1, 
                   row.names = F )

sheet3 = xlsx::createSheet(output, sheetName = "science_vs_liberal_arts")
xlsx::addDataFrame(disc.table, sheet3, startRow = 1, startColumn = 1, 
                   row.names = F )

sheet4 = xlsx::createSheet(output, sheetName = "ECA_requests")
xlsx::addDataFrame(rbind(perfect_table[ , c(2:3)], general_table), sheet4, startRow = 1, startColumn = 1, 
                   row.names = F )

sheet5 = xlsx::createSheet(output, sheetName = "ref_table")
xlsx::addDataFrame(matching_results_with_info[ ,c(1:94)], sheet5, startRow = 1, startColumn = 1, 
                   row.names = F )

xlsx::saveWorkbook(output, paste("matching_evaluation_", Sys.Date(), ".xlsx", sep = ""))



## Miscellaneous 
## ---------------------------
# prov_preference = cbind(split_uneven_string(as.character(matching_results_16_with_info$td_我了解下列哪个地区的风俗文化.), 
#                                       max.length = 4, sep.by = "┋" )
#                         , split_uneven_string(as.character(matching_results_16_with_info$td_我会说.听得懂下列哪个地区的方言.), 
#                                               max.length = 4, sep.by = "┋" ))

# prov_preference = apply(prov_preference, 1, function(x) paste0(sort(setdiff(unique(x), c("都不了解", NA))), collapse = ","))
# matching_results_16_with_info = cbind(matching_results_16_with_info, prov_preference)
