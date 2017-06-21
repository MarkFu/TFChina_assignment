# Setup
# ----------------------------------------

Required_Packages <- c("xlsx", "xtable", "zoo", "lubridate", "ggplot2", "gridExtra", "gdata", "brew")
Remaining_Packages <- Required_Packages[!(Required_Packages %in% installed.packages()[,"Package"])]

if (length(Remaining_Packages)) install.packages(Remaining_Packages, repos='http://cran.us.r-project.org')
for(package_name in Required_Packages) suppressMessages(library(package_name,character.only=TRUE,quietly=TRUE))

# Data Preparation
# ----------------------------------------

# Input data

teacher_data = read.xlsx2("data/test_run/data_cleaned_v4.xlsx", sheetIndex = 1)
teacher_data = apply(teacher_data, 2, trim) # trim off white spaces
colnames(teacher_data) = paste0("td_", colnames(teacher_data))
 
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

# ----------------------------------------
# Diagnostics
# ----------------------------------------

# Overall fullfillment rate

table(matching_results_with_info$科目优先级) # Number of teachers requested, by priority level

prop.table(table(matching_results_with_info$科目优先级, matching_results_with_info$老师 == ""), 1)

matching_results_with_info$地级市 = as.character(matching_results_with_info$地级市)

fullfillment.district = NULL

for (d in unique(matching_results_with_info$地级市)){
  
  results_d = subset(matching_results_with_info, 地级市 == d)
  p.fullfilled = data.frame(prop.table(table(results_d$科目优先级, results_d$老师 == ""), 1))
  p.fullfilled = subset(p.fullfilled, Var2 == "FALSE" & !is.na(Freq))
  fullfillment.district = rbind(fullfillment.district, data.frame(district = rep(d, nrow(p.fullfilled)), priority = p.fullfilled$Var1, p.fullfilled = p.fullfilled$Freq))
}

# Matching Rules
# ----------------------------------------

## 1: Medical Requirement

matching_results_with_info$needs_medical_care = ifelse(matching_results_with_info$td_X..我有较严重的过往病史.. == "是", 1, 0)
table(matching_results_with_info$needs_medical_care)

with(subset(matching_results_with_info, needs_medical_care == 1), sd_就医条件) ## All "良好")

## 2 - ?: Balanced distributions of the following factors at district level
## 2: Gender 
## 3: School Tiers
## 4: Normal School Students

prop.table(table(subset(teacher_data, select = "td_性别"))) # 7:3

prop.table(table(matching_results_with_info$td_毕业院校..海外)) # 8%
prop.table(table(matching_results_with_info$td_毕业院校..985)) # 41%
prop.table(table(matching_results_with_info$td_毕业院校..211)) # 66%

gender.ratio.district = NULL
school.ratio.district = NULL

for (d in unique(matching_results_with_info$地级市)){
  
  results_d = subset(matching_results_with_info, 地级市 == d & !(老师 == ""))
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
  
}


## Teacher's Preference for Destination Province

prop.table(table(matching_results_with_info$td_我了解下列哪个地区的风俗文化.))

prov_preference = cbind(split_uneven_string(as.character(matching_results_16_with_info$td_我了解下列哪个地区的风俗文化.), 
                                      max.length = 4, sep.by = "┋" )
                        , split_uneven_string(as.character(matching_results_16_with_info$td_我会说.听得懂下列哪个地区的方言.), 
                                              max.length = 4, sep.by = "┋" ))

prov_preference = apply(prov_preference, 1, function(x) paste0(sort(setdiff(unique(x), c("都不了解", NA))), collapse = ","))
matching_results_16_with_info = cbind(matching_results_16_with_info, prov_preference)

matching_results_16_with_info$sd_省 = as.character(matching_results_16_with_info$sd_省)
matching_results_16_with_info$sd_省 = gsub("省", "", matching_results_16_with_info$sd_省)
matching_results_16_with_info$sd_省 = gsub("壮族自治区", "", matching_results_16_with_info$sd_省)

matching_results_16_with_info$prov_preference_met = unlist(lapply(c(1:nrow(matching_results_16_with_info)), 
                                                                  function(x) max(grep(matching_results_16_with_info$sd_省[x], matching_results_16_with_info$prov_preference[x]), 0, na.rm = T)))

matching_results_16_with_info$prov_preference_met[matching_results_16_with_info$prov_preference == ""] = 1
summary( matching_results_16_with_info$prov_preference_met) ## 70%

## 5: Science vs Liberal Arts


## 6: English skills


## 7: PE/Music/Art 


## 9: Soft Skills

