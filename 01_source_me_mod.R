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
tictoc::tic()
# LIBRARIES--------------
library("tidyverse")
library("readxl")
# library("XLConnect")
library("here")
library("lubridate")
# "Constants"----------
current_year <- as.numeric(year(today())) - 1 # delete the -1 once we have current data.
previous_year <- current_year - 1
first_five_years <- current_year + 5
second_five_years <- current_year + 10
region_list <- list(
  "Cariboo",
  "Kootenay",
  "Mainland South West",
  "North Coast & Nechako",
  "North East",
  "Thompson Okanagan",
  "Vancouver Island Coast",
  "British Columbia"
)

# FUNCTIONS----------------
source(here("R", "functions.R"))

# LOAD DATA---------

industry_mapping <- read_csv(here(
  "Data Sources",
  "industry_mapping.csv"
)) %>%
  clean_tbbl()

career_trek <- read_excel(here(
  "Data Sources",
  list.files(here("Data Sources"),
    pattern = "newtemplate"
  )
)) %>%
  clean_tbbl()

empl_bc_regions_industry <- read_excel(here(
  "Data Sources",
  list.files(here("Data Sources"),
    pattern = "Empl_BC_Regions_Industry"
  )
),
skip = 3
) %>%
  clean_tbbl()
# industry_profiles <- read_excel(here(
#   "Data Sources",
#   list.files(here("Data Sources"),
#     pattern = "Industry_Profiles"
#   )
# )) %>%
#   clean_tbbl()
job_openings_bc_regions_industry <- read_excel(here(
  "Data Sources",
  list.files(here("Data Sources"),
    pattern = "JobOpenings_BC_Regions_Industry"
  )
),
skip = 2
) %>%
  clean_tbbl()
lmo_work_bc_unemployment <- read_excel(here(
  "Data Sources",
  list.files(here("Data Sources"),
    pattern = "Unemployment_rate"
  )
),
skip = 3
) %>%
  clean_tbbl()

# 3.3.1 CAREER PROFILE PROVINCIAL---------------
#' 1. Suppression rule:
#'   If the current year's employment is below 20,
#'   then all variables should be changed to "N/A"

bc <- empl_bc_regions_industry %>%
  filter(
    geographic_area == "british_columbia",
    industry == "all_industries"
  ) %>%
  select(-geographic_area, -industry, -variable) %>%
  pivot_longer(
    cols = -c(noc, description),
    names_to = "year",
    values_to = "value"
  ) %>%
  group_by(noc, description) %>%
  nest() %>%
  mutate(cagrs = map(data, get_cagrs, all = TRUE)) %>%
  unnest(cagrs) %>%
  mutate(current_employment = map_dbl(data, current_jobs))

# Job Openings---------------

job_openings <- job_openings_bc_regions_industry %>%
  filter(
    geographic_area == "british_columbia",
    industry == "all_industries"
  ) %>%
  select(-geographic_area, -industry) %>%
  pivot_longer(
    cols = -c(noc, description, variable),
    names_to = "year",
    values_to = "value"
  ) %>%
  mutate(value = round(value, -1)) %>%
  group_by(noc, description) %>%
  nest() %>%
  mutate(
    jo_select_years = map(data, jo_select),
    ten_year_job_openings = map_dbl(data, ten_sum, "job_openings"),
    ten_year_sum_expan = map_dbl(data, ten_sum, "expansion_demand"),
    ten_year_sum_replace = map_dbl(data, ten_sum, "replacement_demand"),
    `expansion_%` = round(ten_year_sum_expan / (ten_year_sum_expan + ten_year_sum_replace), 3) * 100,
    `replacement_%` = 100 - `expansion_%`
  ) %>%
  select(-data) %>%
  unnest(jo_select_years)

provincial_career_profiles <- full_join(bc, job_openings, by = c("noc", "description")) %>%
  relocate(`replacement_%`, .after = ten_year_job_openings) %>%
  relocate(ten_year_sum_replace, .after = `replacement_%`) %>%
  relocate(`expansion_%`, .after = ten_year_sum_replace) %>%
  mutate(
    `expansion_%` = if_else(is.na(`expansion_%`), 0, `expansion_%`),
    `replacement_%` = if_else(is.na(`replacement_%`), 0, `replacement_%`),
    `expansion_%` = if_else(`expansion_%` < 0, 0, `expansion_%`),
    ten_year_sum_replace = if_else(ten_year_sum_expan < 0 & `replacement_%` > 0, ten_year_sum_replace + ten_year_sum_expan, ten_year_sum_replace),
    ten_year_sum_expan = if_else(ten_year_sum_expan < 0, 0, ten_year_sum_expan),
    `expansion_%` = if_else(`expansion_%` > 100, 0, `expansion_%`),
    `replacement_%` = if_else(`replacement_%` < 0, 100, `replacement_%`)
  ) %>%
  mutate(across(where(is.numeric), ~ if_else(current_employment < 20, NA_real_, .x))) %>%
  select(-current_employment, -data)

# 3.3.1 CAREER PROFILE REGIONAL----------
#' 1. Suppression rule:
#'   If the current year's employment is below 20, then all variables should be changed to "N/A"

regions <- empl_bc_regions_industry %>%
  filter(industry == "all_industries") %>%
  select(-industry, -variable) %>%
  pivot_longer(
    cols = -c(noc, description, geographic_area),
    names_to = "year",
    values_to = "value"
  ) %>%
  mutate(value = round(value, -1)) %>%
  group_by(noc, description, geographic_area) %>%
  nest() %>%
  mutate(
    employment = map_dbl(data, current_jobs),
    "avg_employ_growth_{current_year}_{second_five_years}" := map_dbl(data, get_cagrs, all = FALSE)
  ) %>%
  select(-data)

regional_openings <- job_openings_bc_regions_industry %>%
  filter(industry == "all_industries") %>%
  select(-industry) %>%
  pivot_longer(
    cols = -c(noc, description, variable, geographic_area),
    names_to = "year",
    values_to = "value"
  ) %>%
  mutate(value = round(value, -1)) %>%
  group_by(noc, description, geographic_area) %>%
  nest() %>%
  mutate(ten_year_job_openings = map_dbl(data, ten_sum, "job_openings")) %>%
  select(-data)

id_cols <- c("noc", "description", "geographic_area")
career_profile_regional <- full_join(regions,
  regional_openings,
  by = c("noc", "description", "geographic_area")
) %>%
  arrange(geographic_area) %>%
  mutate(across(where(is.numeric), ~ if_else(employment < 20, NA_real_, .x))) %>%
  rename("employment_{current_year}" := employment) %>%
  pivot_wider(
    names_from = geographic_area,
    values_from = -all_of(id_cols),
    names_glue = "{geographic_area}_{.value}",
    names_vary = "slowest"
  ) %>%
  select(!contains("british_columbia"), everything())

# 3.3.2 INDUSTRY PROFILE-------------

industry_profile <- empl_bc_regions_industry%>%
  full_join(industry_mapping, by = "industry")%>%
  filter(geographic_area=="british_columbia",
         description=="total")%>%
  select(-geographic_area, -description, -noc, -industry, -variable)%>%
  pivot_longer(
    cols = -aggregate_industry,
    names_to = "year",
    values_to = "value"
  ) %>%
  mutate(value = round(value, -1)) %>%
  group_by(aggregate_industry) %>%
  nest()%>%
  mutate(data=map(data, aggregate_by_year))

tot_emp <- industry_profile%>%
  filter(aggregate_industry == "all_industries")%>%
  unnest(data)

industry_profile <- industry_profile%>%
  mutate(shares = map(data, get_shares))%>%
  unnest(shares)%>%
  mutate(jobs = map(data, get_values))%>%
  unnest(jobs)%>%
  mutate(cagrs = map(data, get_cagrs, all=TRUE))%>%
  unnest(cagrs)%>%
  select(-data)%>%
  filter(aggregate_industry != "all_industries")%>%
  rename(industry = aggregate_industry)%>%
  arrange(industry)
 
# 3.3.3 REGIONAL PROFILE - REGIONAL PROFILES LMO-------------
# COMPOSITION OF JOB OPENINGS SECTION - REGIONAL PROFILES LMO

composition_job_openings <- job_openings_bc_regions_industry%>%
  filter(variable %in% c("expansion_demand","replacement_demand"),
         industry == "all_industries",
         description == "total")%>%
  select(-noc, -description, -industry)%>%
  pivot_longer(
    cols = -c(variable, geographic_area),
    names_to = "year",
    values_to = "value"
  )%>%
  mutate(value = round(value, -1))%>%
  group_by(geographic_area)%>%
  nest()%>%
  mutate(ten_year_sum_expansion_demand = map_dbl(data, ten_sum, "expansion_demand"),
         ten_year_sum_replacement_demand = map_dbl(data, ten_sum, "replacement_demand"),
         `replacement_%` = round(ten_year_sum_replacement_demand/(ten_year_sum_replacement_demand+ten_year_sum_expansion_demand), 3)*100,
         `expansion_%` = 100 - `replacement_%`
         )%>%
  select(-data)
  #might need to relocate columns in order geo, 10yreplace, replace%, 10yexpand, expand% if final output

# EMPLOYMENT OUTLOOK SECTION - REGIONAL PROFILES LMO

employment_outlook <-empl_bc_regions_industry%>%
  filter(variable=="employment",
         industry=="all_industries",
         description=="total")%>%
  select(-noc, -variable,-industry,-description)%>%
  pivot_longer(cols=-geographic_area, 
               names_to = "year", 
               values_to = "value")%>%
  mutate(value = round(value, -1)) %>%
  group_by(geographic_area)%>%
  nest()%>%
  mutate("{current_year}-{second_five_years}":=map_dbl(data, get_cagrs, all=FALSE))%>%
  mutate(jobs=map(data, get_values, short_names = TRUE))%>%
  unnest(jobs)%>%
  select(-data)

jobopenings <- job_openings_bc_regions_industry%>%
  filter(variable=="job_openings",
         industry=="all_industries",
         description=="total")%>%
  select(-noc, -variable,-industry,-description)%>%
  pivot_longer(cols=-geographic_area, 
               names_to = "year", 
               values_to = "value")%>%
  mutate(value = round(value, -1)) %>%
  group_by(geographic_area)%>%
  nest()%>%
  mutate(ten_year_sum=map_dbl(data, ten_sum))%>%
  unnest(ten_year_sum)%>%
  select(-data)

employment_outlook <- employment_outlook%>%
  full_join(jobopenings, by = "geographic_area")
employment_outlook <- employment_outlook[, c(1,3,4,5,6,2)]

# 3.3.3 REGIONAL PROFILE - TOP 10 INDUSTRIES---------------

regional_profile_jo <- job_openings_bc_regions_industry%>%
  filter(industry != "all_industries",
        variable=="job_openings",
        description=="total")%>%
  left_join(industry_mapping, by = "industry")%>%
  select(-variable, -description, -noc, -industry)%>%
  pivot_longer(cols=-c(geographic_area, aggregate_industry),
               names_to = "year",
               values_to = "value")%>%
  mutate(value = round(value, -1)) %>%
  group_by(geographic_area, aggregate_industry)%>%
  nest()%>%
  mutate(data = map(data, aggregate_by_year),
         ten_year_job_openings = map_dbl(data, ten_sum)
         )%>%
  select(-data)

regional_profile_cagr <- empl_bc_regions_industry%>%
  filter(industry != "all_industries",
         variable=="employment",
         description=="total")%>%
  left_join(industry_mapping, by = "industry")%>%
  select(-variable, -description, -noc, -industry)%>%
  pivot_longer(cols=-c(geographic_area, aggregate_industry),
               names_to = "year",
               values_to = "value")%>%
  mutate(value = round(value, -1)) %>%
  group_by(geographic_area, aggregate_industry)%>%
  nest()%>%
  mutate(data = map(data, aggregate_by_year),
         "ave_employ_growth_{current_year}_{second_five_years}" := map_dbl(data, get_cagrs, all=FALSE)       
  )%>%
  select(-data)

regional_profile_maindf2 <- full_join(regional_profile_jo,
                                      regional_profile_cagr,
                                      by = c("geographic_area", 
                                             "aggregate_industry"))%>%
  group_by(geographic_area, .add=FALSE)%>%
  slice_max(order_by = ten_year_job_openings, n = 10)


# 3.3.3 REGIONAL PROFILE - TOP 10 OCCUPATIONS-----------
jo_noc <- job_openings_bc_regions_industry%>%
  filter(industry == "all_industries",
        variable == "job_openings")%>%
  select(-industry, -variable)%>%
  pivot_longer(cols=-c(noc, description, geographic_area),
               names_to = "year",
               values_to = "value")%>%
  mutate(value = round(value, -1)) %>%
  group_by(noc, description, geographic_area)%>%
  nest()%>%
  mutate(ten_year_job_openings = map_dbl(data, ten_sum))%>%
  select(-data)

emp_noc <- empl_bc_regions_industry%>%
  filter(industry == "all_industries",
        variable=="employment")%>%
  select(-industry, -variable)%>%
  pivot_longer(cols=-c(noc, description, geographic_area),
               names_to = "year",
               values_to = "value")%>%
  mutate(value = round(value, -1)) %>%
  group_by(noc, description, geographic_area)%>%
  nest()%>%
  mutate("ave_employ_growth_{current_year}_{second_five_years}":=map_dbl(data, get_cagrs, all=FALSE))%>%
  select(-data)

occupations_top_ten <- full_join(jo_noc, 
                                 emp_noc,
                                 by = c("noc", "description", "geographic_area")
                                 )%>%
  filter(noc!="#t")%>%
  group_by(geographic_area, .add=FALSE)%>%
  slice_max(order_by = ten_year_job_openings, n = 10)

# 3.4 CAREER COMPASS - BROWSE CAREERS---------

noc_unemp <- lmo_work_bc_unemployment%>%
  select(-industry, -variable)%>%
  pivot_longer(cols=-c(noc, description, geographic_area),
               names_to = "year",
               values_to = "value")%>%
  group_by(noc, description, geographic_area)%>%
  nest()%>%
  mutate(unemployment = map(data, get_values, short_names=TRUE))%>%
  unnest(unemployment, names_sep = "_")%>%
  select(-data)
  
noc_cagr <- emp_noc%>%
  filter(geographic_area=="british_columbia")

#basically duplicated code...
  
noc_jo <- job_openings_bc_regions_industry%>%
  filter(geographic_area=="british_columbia",
        industry == "all_industries",
        variable == "job_openings")%>%
  select(-industry, -variable)%>%
  pivot_longer(cols=-c(noc, description, geographic_area),
               names_to = "year",
               values_to = "value")%>%
  mutate(value = round(value, -1)) %>%
  group_by(noc, description, geographic_area)%>%
  nest()%>%
  mutate(ten_year_job_openings = map_dbl(data, ten_sum),
        jo=map(data, get_values, short_names=TRUE))%>%
  unnest(jo, names_sep = "_")%>%
  select(-data)

lmo_career_compass <- noc_unemp%>%
  full_join(noc_cagr, by = c("noc", "description", "geographic_area"))%>%
  full_join(noc_jo, by = c("noc", "description", "geographic_area"))%>%
  ungroup()%>%
  select(-geographic_area)

# 3.5 CAREER TREK-----------

career_trek_merged <- lmo_career_compass%>%
  select(-contains("unemployment"),-contains("jo_"))%>%
  semi_join(career_trek, by = "noc")

#foolery to get data into excel templates...

career_profile_regional_template <- career_profile_regional%>%
  add_column(` `="", .after = "cariboo_ten_year_job_openings")%>%
  add_column(`  `="", .after = "kootenay_ten_year_job_openings")%>%
  add_column(`   `="", .after = "mainland_south_west_ten_year_job_openings")%>%
  add_column(`    `="", .after = "north_coast_&_nechako_ten_year_job_openings")%>%
  add_column(`     `="", .after = "north_east_ten_year_job_openings")%>%
  add_column(`      `="", .after = "thompson_okanagan_ten_year_job_openings")%>%
  select(-contains("british_columbia"))
# EXPORTING DATA DIRECTLY TO EXCEL----------
# 3.3.1
wb <- loadWorkbook(here("templates", "3.3.1_WorkBC_Career_Profile_Data.xlsx"))
# Provincial Outlook Data
write_workbook(Provincial_Career_Profiles[, c(3:12)], "Provincial Outlook", 4, 3)
# Regional Outlook Data - Cariboo
write_workbook(Career_Profile_Regional[, c(3:5)], "Regional Outlook", 4, 3)
# Regional Outlook Data - Kootney
write_workbook(Career_Profile_Regional[, c(6:8)], "Regional Outlook", 4, 7)
# Regional Outlook Data - Mainland Southwest
write_workbook(Career_Profile_Regional[, c(9:11)], "Regional Outlook", 4, 11)
# Regional Outlook Data - North Coast and Nechako
write_workbook(Career_Profile_Regional[, c(12:14)], "Regional Outlook", 4, 15)
# Regional Outlook Data - Northeast
write_workbook(Career_Profile_Regional[, c(15:17)], "Regional Outlook", 4, 19)
# Regional Outlook Data - Thompson Okanagan
write_workbook(Career_Profile_Regional[, c(18:20)], "Regional Outlook", 4, 23)
# Regional Outlook Data - Vancouver Island Coast
write_workbook(Career_Profile_Regional[, c(21:23)], "Regional Outlook", 4, 27)
saveWorkbook(
  wb,
  here(
    "Send to WorkBC",
    paste0(
      "3.3.1_WorkBC_Career_Profile_Data",
      current_year,
      "-",
      second_five_years,
      ".xlsx"
    )
  )
)
# export 3.3.2-------------
wb <- loadWorkbook(here("templates", "3.3.2_WorkBC_Industry_Profile.xlsx"))
# BC Industry Data
write_workbook(Industry_Profile[, c(2:9)], "BC", 4, 2)
saveWorkbook(
  wb,
  here(
    "Send to WorkBC",
    paste0(
      "3.3.2_WorkBC_Industry_Profile",
      current_year,
      "-",
      second_five_years,
      ".xlsx"
    )
  )
)

# export 3.3.3-----------
wb <- loadWorkbook(here("templates", "3.3.3_WorkBC_Regional_Profile_Data.xlsx"))
write_workbook(composition_job_openings[, c(2:5)], "Regional Profiles - LMO", 5, 2)
# Regional Data
write_workbook(employment_outlook, "Regional Profiles - LMO", 5, 6)
# Top Occupations - BC
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "British Columbia", 2:5], "Top Occupation", 5, 1)
# Top Occupation - Cariboo
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Cariboo", 2:5], "Top Occupation", 17, 1)
# Top Occupation - Kootenay
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Kootenay", 2:5], "Top Occupation", 29, 1)
# Top Occupation - Mainland SouthWest
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Mainland South West", 2:5], "Top Occupation", 41, 1)
# Top Occupation - North Coast & Nechako
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North Coast & Nechako", 2:5], "Top Occupation", 53, 1)
# Top Occupation - North East
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North East", 2:5], "Top Occupation", 65, 1)
# Top Occupation - Thompson Okanagan
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Thompson Okanagan", 2:5], "Top Occupation", 77, 1)
# Top Occupation - Vancouver Island Coast
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Vancouver Island Coast", 2:5], "Top Occupation", 89, 1)
# Top Industries - BC
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "British Columbia", 2:4], "Top Industries", 5, 1)
# Top Industries - Cariboo
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Cariboo", 2:4], "Top Industries", 17, 1)
# Top Industries - Kootenay
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Kootenay", 2:4], "Top Industries", 29, 1)
# Top Industries - Mainland SouthWest
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Mainland South West", 2:4], "Top Industries", 42, 1)
# Top Industries - North Coast & Nechako
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North Coast & Nechako", 2:4], "Top Industries", 54, 1)
# Top Industries - North East
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North East", 2:4], "Top Industries", 67, 1)
# Top Industries - Thompson Okanagan
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Thompson Okanagan", 2:4], "Top Industries", 79, 1)
# Top Industries - Vancouver Island Coast
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Vancouver Island Coast", 2:4], "Top Industries", 92, 1)
saveWorkbook(
  wb,
  here(
    "Send to WorkBC",
    paste0(
      "3.3.3_WorkBC_Regional_Profile_Data",
      current_year,
      "-",
      second_five_years,
      ".xlsx"
    )
  )
)
# export 3.4---------------
wb <- loadWorkbook(here("templates", "3.4_WorkBC_Career_Compass.xlsx"))
# Browse Careers
write_workbook(LMO_Career_Compass[, c(3:10)], "Browse Careers", 4, 3)
# Top Occupation - Cariboo
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Cariboo", 2:4], "Regions", 5, 1)
# Top Occupation - Kootenay
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Kootenay", 2:4], "Regions", 17, 1)
# Top Occupation - Mainland SouthWest
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Mainland South West", 2:4], "Regions", 29, 1)
# Top Occupation - North Coast & Nechako
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North Coast & Nechako", 2:4], "Regions", 41, 1)
# Top Occupation - North East
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North East", 2:4], "Regions", 53, 1)
# Top Occupation - Thompson Okanagan
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Thompson Okanagan", 2:4], "Regions", 65, 1)
# Top Occupation - Vancouver Island Coast
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Vancouver Island Coast", 2:4], "Regions", 77, 1)
# Top Industries - Cariboo
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Cariboo", 2:3], "Regions", 5, 5)
# Top Industries - Kootenay
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Kootenay", 2:3], "Regions", 17, 5)
# Top Industries - Mainland SouthWest
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Mainland South West", 2:3], "Regions", 29, 5)
# Top Industries - North Coast & Nechako
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North Coast & Nechako", 2:3], "Regions", 41, 5)
# Top Industries - North East
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North East", 2:3], "Regions", 53, 5)
# Top Industries - Thompson Okanagan
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Thompson Okanagan", 2:3], "Regions", 65, 5)
# Top Industries - Vancouver Island Coast
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Vancouver Island Coast", 2:3], "Regions", 77, 5)
saveWorkbook(
  wb,
  here(
    "Send to WorkBC",
    paste0(
      "3.4_WorkBC_Career_Compass",
      current_year,
      "-",
      second_five_years,
      ".xlsx"
    )
  )
)
# export 3.5-------------------
wb <- loadWorkbook(here("templates", "3.5_WorkBC_Career_Trek.xlsx"))
# Browse Careers
write_workbook(Career_Trek[, c(3:4)], "LMO", 2, 5)
saveWorkbook(
  wb,
  here(
    "Send to WorkBC",
    paste0(
      "3.5_WorkBC_Career_Trek",
      current_year,
      "-",
      second_five_years,
      ".xlsx"
    )
  )
)

# export 3.7--------------
wb <- loadWorkbook(here("templates", "3.7_WorkBC_Buildprint_Builder.xlsx"))
# Regional Data
write_workbook(composition_job_openings[, c(2:5)], "Regional Data", 5, 2)
# Jobs in Demand
write_workbook(Occupations_top_ten[order(Occupations_top_ten$`Geographic Area`), c(3, 2)], "Jobs in Demand", 4, 2)
saveWorkbook(
  wb,
  here(
    "Send to WorkBC",
    paste0(
      "3.7_WorkBC_Buildprint_Builder",
      current_year,
      "-",
      second_five_years,
      ".xlsx"
    )
  )
)
tictoc::toc()