# Copyright 2022 Province of British Columbia
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
# http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and limitations under the License.
# LIBRARIES--------------
# library("profvis")
library("tidyverse")
library("readxl")
library("XLConnect")
library("here")
library("lubridate")

# "Constants"----------
current_year <- as.numeric(year(today())) - 1 # delete the -1 once we have current data.
previous_year <- current_year - 1
current_plus_5 <- current_year + 5
current_plus_10 <- current_year + 10

# FUNCTIONS----------------
source(here("R", "functions.R"))

# LOAD DATA---------

career_trek <- read_excel(here("Data Sources", list.files(here("Data Sources"), pattern = "newtemplate"))) %>%
  clean_tbbl()

industry_mapping <- read_csv(here("Data Sources", "industry_mapping.csv")) %>%
  clean_tbbl()

empl_bc_regions_industry <- read_excel(here("Data Sources", list.files(here("Data Sources"), pattern = "Empl_BC_Regions_Industry")), skip = 3) %>%
  clean_tbbl()

job_openings_bc_regions_industry <- read_excel(here("Data Sources", list.files(here("Data Sources"), pattern = "JobOpenings_BC_Regions_Industry")), skip = 2) %>%
  clean_tbbl()

lmo_work_bc_unemployment <- read_excel(here("Data Sources", list.files(here("Data Sources"), pattern = "Unemployment_rate")), skip = 3) %>%
  clean_tbbl()

# make one long df--------

long_emp <- pl_wrap(empl_bc_regions_industry) %>%
  filter(year != previous_year) # empl_bc_regions_industry contains previous year data
long_jo <- pl_wrap(job_openings_bc_regions_industry)
long_un <- pl_wrap(lmo_work_bc_unemployment)

long <- bind_rows(long_emp, long_jo) %>%
  mutate(value = round(value, -1)) %>% # employment and job openings data rounded to nearest 10
  bind_rows(long_un) %>% # unemployment rates NOT rounded
  full_join(industry_mapping, by = "industry")

# 3.3.1 CAREER PROFILE PROVINCIAL---------------
#' This is the "by NOC" breakdown of labour market.
#' 1. Suppression rule:
#'   If the current year's employment is below 20,
#'   then all variables should be changed to "N/A"
#'
provincial_career_profiles <- long %>%
  filter(
    geographic_area == "british_columbia",
    industry == "all_industries"
  ) %>%
  select(-geographic_area, industry) %>%
  group_by(noc, description) %>%
  nest() %>%
  mutate(
    cagr = map(data, get_cagrs, "employment", all = TRUE),
    current_employment = map_dbl(data, current_value, "employment"),
    jo = map(data, current_5_10, "job_openings"),
    ten_year_job_openings = map_dbl(data, ten_sum, "job_openings"),
    ten_year_sum_expan = map_dbl(data, ten_sum, "expansion_demand"),
    ten_year_sum_replace = map_dbl(data, ten_sum, "replacement_demand"),
    expansion_percent = round(ten_year_sum_expan / (ten_year_sum_expan + ten_year_sum_replace), 3) * 100,
    replacement_percent = 100 - expansion_percent
  ) %>%
  select(-data) %>%
  unnest(jo, names_sep = "_") %>%
  unnest(cagr, names_sep = "_") %>%
  relocate(replacement_percent, .after = ten_year_job_openings) %>%
  relocate(ten_year_sum_replace, .after = replacement_percent) %>%
  relocate(expansion_percent, .after = ten_year_sum_replace) %>%
  mutate(
    expansion_percent = if_else(is.na(expansion_percent), 0, expansion_percent),
    replacement_percent = if_else(is.na(replacement_percent), 0, replacement_percent),
    expansion_percent = if_else(expansion_percent < 0, 0, expansion_percent),
    ten_year_sum_replace = if_else(ten_year_sum_expan < 0 & replacement_percent > 0, ten_year_sum_replace + ten_year_sum_expan, ten_year_sum_replace),
    ten_year_sum_expan = if_else(ten_year_sum_expan < 0, 0, ten_year_sum_expan),
    expansion_percent = if_else(expansion_percent > 100, 0, expansion_percent),
    replacement_percent = if_else(replacement_percent < 0, 100, replacement_percent)
  ) %>%
  mutate(across(where(is.numeric), ~ if_else(current_employment < 20, NA_real_, .x))) %>%
  select(-current_employment)%>%
  camel_to_title()

# 3.3.1 CAREER PROFILE REGIONAL----------
#' This is the "by NOC and geographic_area" breakdown of the labour market.
#' 1. Suppression rule:
#'   If the current year's employment is below 20, then all variables should be changed to "N/A"

career_profile_regional <- long %>%
  filter(industry == "all_industries") %>%
  select(-industry) %>%
  group_by(noc, description, geographic_area) %>%
  nest() %>%
  mutate(
    current_employment = map_dbl(data, current_value, "employment"),
    ten_year_employment_growth = map_dbl(data, get_cagrs, "employment", all = FALSE),
    ten_year_job_openings = map_dbl(data, ten_sum, "job_openings")
  ) %>%
  select(-data) %>%
  arrange(geographic_area) %>%
  mutate(across(where(is.numeric), ~ if_else(current_employment < 20, NA_real_, .x))) %>%
  pivot_wider(
    names_from = geographic_area,
    values_from = -all_of(c("noc", "description", "geographic_area")),
    names_glue = "{geographic_area}_{.value}",
    names_vary = "slowest"
  ) %>%
  select(!contains("british_columbia"), everything())%>%
  camel_to_title()

# 3.3.2 INDUSTRY PROFILE-------------
# This is the "by aggregate_industry" breakdown of the labour market.

industry_profile <- long %>%
  filter(
    geographic_area == "british_columbia",
    description == "total"
  ) %>%
  select(-geographic_area, -description) %>%
  group_by(aggregate_industry) %>%
  nest() %>%
  mutate(data = map(data, aggregate_by_year, "employment"))

tot_emp <- industry_profile %>%
  filter(aggregate_industry == "all_industries") %>%
  unnest(data)

industry_profile <- industry_profile %>%
  mutate(employment_share = map(data, get_shares, "employment")) %>%
  unnest(employment_share, names_sep = "_") %>%
  mutate(employment = map(data, current_5_10, "employment")) %>%
  unnest(employment, names_sep = "_") %>%
  mutate(cagr = map(data, get_cagrs, "employment", all = TRUE)) %>%
  unnest(cagr, names_sep = "_") %>%
  select(-data) %>%
  filter(aggregate_industry != "all_industries") %>%
  rename(industry = aggregate_industry) %>%
  arrange(industry)%>%
  camel_to_title() 

##########################################################
# 3.3.3 REGIONAL PROFILE - REGIONAL PROFILES LMO-------------
# COMPOSITION OF JOB OPENINGS SECTION - REGIONAL PROFILES LMO
# This is a "by geographic_area" breakdown of job openings.

composition_job_openings <- long %>%
  filter(
    industry == "all_industries",
    description == "total"
  ) %>%
  select(-industry, -description) %>%
  group_by(geographic_area) %>%
  nest() %>%
  mutate(
    ten_year_sum_expansion_demand = map_dbl(data, ten_sum, "expansion_demand"),
    ten_year_sum_replacement_demand = map_dbl(data, ten_sum, "replacement_demand"),
    replacement_percent = round(ten_year_sum_replacement_demand / (ten_year_sum_replacement_demand + ten_year_sum_expansion_demand), 3) * 100,
    expansion_percent = 100 - replacement_percent
  ) %>%
  select(-data)%>%
  camel_to_title()

# EMPLOYMENT OUTLOOK SECTION - REGIONAL PROFILES LMO
# This is a "by geographic_area" breakdown of employment and job openings.

employment_outlook <- long %>%
  filter(
    industry == "all_industries",
    description == "total"
  ) %>%
  select(-industry, -description) %>%
  group_by(geographic_area) %>%
  nest() %>%
  mutate(ten_year_cagr = map_dbl(data, get_cagrs, "employment", all = FALSE)) %>%
  mutate(
    employment = map(data, current_5_10, "employment"),
    ten_sum_jo = map_dbl(data, ten_sum, "job_openings")
  ) %>%
  unnest(employment, names_sep = "_")%>%
  select(-data)%>%
  camel_to_title()

# 3.3.3 REGIONAL PROFILE - TOP 10 INDUSTRIES---------------
# This is the "by geographic_area and aggregate_industry" breakdown of the labour market.

regional_profile_maindf2 <- long %>%
  filter(
    industry != "all_industries",
    description == "total"
  ) %>%
  select(-description) %>%
  group_by(geographic_area, aggregate_industry) %>%
  nest() %>%
  mutate(
    jo = map(data, aggregate_by_year, "job_openings"),
    ten_year_job_openings = map_dbl(jo, ten_sum, "job_openings"),
    emp = map(data, aggregate_by_year, "employment"),
    ten_year_cagr = map_dbl(emp, get_cagrs, "employment", all = FALSE)
  ) %>%
  select(-data, -jo, -emp) %>%
  group_by(geographic_area, .add = FALSE) %>%
  slice_max(order_by = ten_year_job_openings, n = 10)%>%
  camel_to_title()

# 3.3.3 REGIONAL PROFILE - TOP 10 OCCUPATIONS-----------
# This is the "by noc and geographic_area" breakdown of the labour market.

occupations_top_ten <- long %>%
  filter(industry == "all_industries") %>%
  select(-industry) %>%
  group_by(noc, description, geographic_area) %>%
  nest() %>%
  mutate(
    ten_year_job_openings = map_dbl(data, ten_sum, "job_openings"),
    ten_year_cagr = map_dbl(data, get_cagrs, "employment", all = FALSE)
  ) %>%
  select(-data) %>%
  filter(noc != "#t") %>%
  group_by(geographic_area, .add = FALSE) %>%
  slice_max(order_by = ten_year_job_openings, n = 10)%>%
  camel_to_title()

# 3.4 CAREER COMPASS - BROWSE CAREERS---------
# This is the "by noc" breakdown of the labour market.
lmo_career_compass <- long %>%
  filter(
    geographic_area == "british_columbia",
    industry == "all_industries"
  ) %>%
  select(-geographic_area, -industry) %>%
  group_by(noc, description) %>%
  nest() %>%
  mutate(
    ten_year_job_openings = map_dbl(data, ten_sum, "job_openings"),
    jo = map(data, current_5_10, "job_openings"),
    unemployment = map(data, current_5_10, "unemployment_rate"),
    cagr = map(data, get_cagrs, "employment", all = TRUE)
  ) %>%
  unnest(cagr, names_sep = "_") %>%
  unnest(unemployment, names_sep = "_") %>%
  unnest(jo, names_sep = "_") %>%
  select(-data) %>%
  ungroup()

# 3.5 CAREER TREK-----------

career_trek_merged <- lmo_career_compass %>%
  select(-contains("unemployment"), -contains("jo_")) %>%
  semi_join(career_trek, by = "noc")%>%
  camel_to_title()

# foolery to get data into excel templates...
#
# career_profile_regional_template <- career_profile_regional%>%
#   add_column(` `="", .after = "cariboo_ten_year_job_openings")%>%
#   add_column(`  `="", .after = "kootenay_ten_year_job_openings")%>%
#   add_column(`   `="", .after = "mainland_south_west_ten_year_job_openings")%>%
#   add_column(`    `="", .after = "north_coast_&_nechako_ten_year_job_openings")%>%
#   add_column(`     `="", .after = "north_east_ten_year_job_openings")%>%
#   add_column(`      `="", .after = "thompson_okanagan_ten_year_job_openings")%>%
#   select(-contains("british_columbia"))
# # EXPORTING DATA DIRECTLY TO EXCEL----------
# # 3.3.1
# wb <- loadWorkbook(here("templates", "3.3.1_WorkBC_Career_Profile_Data.xlsx"))
# # Provincial Outlook Data
# write_workbook(Provincial_Career_Profiles[, c(3:12)], "Provincial Outlook", 4, 3)
# # Regional Outlook Data - Cariboo
# write_workbook(Career_Profile_Regional[, c(3:5)], "Regional Outlook", 4, 3)
# # Regional Outlook Data - Kootney
# write_workbook(Career_Profile_Regional[, c(6:8)], "Regional Outlook", 4, 7)
# # Regional Outlook Data - Mainland Southwest
# write_workbook(Career_Profile_Regional[, c(9:11)], "Regional Outlook", 4, 11)
# # Regional Outlook Data - North Coast and Nechako
# write_workbook(Career_Profile_Regional[, c(12:14)], "Regional Outlook", 4, 15)
# # Regional Outlook Data - Northeast
# write_workbook(Career_Profile_Regional[, c(15:17)], "Regional Outlook", 4, 19)
# # Regional Outlook Data - Thompson Okanagan
# write_workbook(Career_Profile_Regional[, c(18:20)], "Regional Outlook", 4, 23)
# # Regional Outlook Data - Vancouver Island Coast
# write_workbook(Career_Profile_Regional[, c(21:23)], "Regional Outlook", 4, 27)
# saveWorkbook(
#   wb,
#   here(
#     "Send to WorkBC",
#     paste0(
#       "3.3.1_WorkBC_Career_Profile_Data",
#       current_year,
#       "-",
#       current_plus_10,
#       ".xlsx"
#     )
#   )
# )
# # export 3.3.2-------------
# wb <- loadWorkbook(here("templates", "3.3.2_WorkBC_Industry_Profile.xlsx"))
# # BC Industry Data
# write_workbook(Industry_Profile[, c(2:9)], "BC", 4, 2)
# saveWorkbook(
#   wb,
#   here(
#     "Send to WorkBC",
#     paste0(
#       "3.3.2_WorkBC_Industry_Profile",
#       current_year,
#       "-",
#       current_plus_10,
#       ".xlsx"
#     )
#   )
# )
#
# # export 3.3.3-----------
# wb <- loadWorkbook(here("templates", "3.3.3_WorkBC_Regional_Profile_Data.xlsx"))
# write_workbook(composition_job_openings[, c(2:5)], "Regional Profiles - LMO", 5, 2)
# # Regional Data
# write_workbook(employment_outlook, "Regional Profiles - LMO", 5, 6)
# # Top Occupations - BC
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "British Columbia", 2:5], "Top Occupation", 5, 1)
# # Top Occupation - Cariboo
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Cariboo", 2:5], "Top Occupation", 17, 1)
# # Top Occupation - Kootenay
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Kootenay", 2:5], "Top Occupation", 29, 1)
# # Top Occupation - Mainland SouthWest
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Mainland South West", 2:5], "Top Occupation", 41, 1)
# # Top Occupation - North Coast & Nechako
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North Coast & Nechako", 2:5], "Top Occupation", 53, 1)
# # Top Occupation - North East
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North East", 2:5], "Top Occupation", 65, 1)
# # Top Occupation - Thompson Okanagan
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Thompson Okanagan", 2:5], "Top Occupation", 77, 1)
# # Top Occupation - Vancouver Island Coast
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Vancouver Island Coast", 2:5], "Top Occupation", 89, 1)
# # Top Industries - BC
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "British Columbia", 2:4], "Top Industries", 5, 1)
# # Top Industries - Cariboo
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Cariboo", 2:4], "Top Industries", 17, 1)
# # Top Industries - Kootenay
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Kootenay", 2:4], "Top Industries", 29, 1)
# # Top Industries - Mainland SouthWest
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Mainland South West", 2:4], "Top Industries", 42, 1)
# # Top Industries - North Coast & Nechako
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North Coast & Nechako", 2:4], "Top Industries", 54, 1)
# # Top Industries - North East
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North East", 2:4], "Top Industries", 67, 1)
# # Top Industries - Thompson Okanagan
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Thompson Okanagan", 2:4], "Top Industries", 79, 1)
# # Top Industries - Vancouver Island Coast
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Vancouver Island Coast", 2:4], "Top Industries", 92, 1)
# saveWorkbook(
#   wb,
#   here(
#     "Send to WorkBC",
#     paste0(
#       "3.3.3_WorkBC_Regional_Profile_Data",
#       current_year,
#       "-",
#       current_plus_10,
#       ".xlsx"
#     )
#   )
# )
# # export 3.4---------------
# wb <- loadWorkbook(here("templates", "3.4_WorkBC_Career_Compass.xlsx"))
# # Browse Careers
# write_workbook(LMO_Career_Compass[, c(3:10)], "Browse Careers", 4, 3)
# # Top Occupation - Cariboo
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Cariboo", 2:4], "Regions", 5, 1)
# # Top Occupation - Kootenay
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Kootenay", 2:4], "Regions", 17, 1)
# # Top Occupation - Mainland SouthWest
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Mainland South West", 2:4], "Regions", 29, 1)
# # Top Occupation - North Coast & Nechako
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North Coast & Nechako", 2:4], "Regions", 41, 1)
# # Top Occupation - North East
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North East", 2:4], "Regions", 53, 1)
# # Top Occupation - Thompson Okanagan
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Thompson Okanagan", 2:4], "Regions", 65, 1)
# # Top Occupation - Vancouver Island Coast
# write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Vancouver Island Coast", 2:4], "Regions", 77, 1)
# # Top Industries - Cariboo
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Cariboo", 2:3], "Regions", 5, 5)
# # Top Industries - Kootenay
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Kootenay", 2:3], "Regions", 17, 5)
# # Top Industries - Mainland SouthWest
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Mainland South West", 2:3], "Regions", 29, 5)
# # Top Industries - North Coast & Nechako
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North Coast & Nechako", 2:3], "Regions", 41, 5)
# # Top Industries - North East
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North East", 2:3], "Regions", 53, 5)
# # Top Industries - Thompson Okanagan
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Thompson Okanagan", 2:3], "Regions", 65, 5)
# # Top Industries - Vancouver Island Coast
# write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Vancouver Island Coast", 2:3], "Regions", 77, 5)
# saveWorkbook(
#   wb,
#   here(
#     "Send to WorkBC",
#     paste0(
#       "3.4_WorkBC_Career_Compass",
#       current_year,
#       "-",
#       current_plus_10,
#       ".xlsx"
#     )
#   )
# )
# # export 3.5-------------------
# wb <- loadWorkbook(here("templates", "3.5_WorkBC_Career_Trek.xlsx"))
# # Browse Careers
# write_workbook(Career_Trek[, c(3:4)], "LMO", 2, 5)
# saveWorkbook(
#   wb,
#   here(
#     "Send to WorkBC",
#     paste0(
#       "3.5_WorkBC_Career_Trek",
#       current_year,
#       "-",
#       current_plus_10,
#       ".xlsx"
#     )
#   )
# )
#
# # export 3.7--------------
# wb <- loadWorkbook(here("templates", "3.7_WorkBC_Buildprint_Builder.xlsx"))
# # Regional Data
# write_workbook(composition_job_openings[, c(2:5)], "Regional Data", 5, 2)
# # Jobs in Demand
# write_workbook(Occupations_top_ten[order(Occupations_top_ten$`Geographic Area`), c(3, 2)], "Jobs in Demand", 4, 2)
# saveWorkbook(
#   wb,
#   here(
#     "Send to WorkBC",
#     paste0(
#       "3.7_WorkBC_Buildprint_Builder",
#       current_year,
#       "-",
#       current_plus_10,
#       ".xlsx"
#     )
#   )
# )
