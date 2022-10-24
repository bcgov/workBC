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

#' BEFORE SOURCING FILE:
#'
#' 1. Ensure that the following (updated) files are in folder raw_data:
#'
#' occupational_characteristics
#' hoo_list
#' pop_municipal
#' career_trek_jobs
#' lmo64_characteristics
#' empl_bc_regions_industry
#' job_openings_bc_regions_industry
#' unemployment_rate
#' career_search_tool_job_openings
#' wages
#'
#' 2.  Ensure that the following files are in folder templates: (request updated versions from WorkBC)
#'
#' 3.3.1_WorkBC_Career_Profile_Data.xlsx
#' 3.3.2_WorkBC_Industry_Profile.xlsx
#' 3.3.3_WorkBC_Regional_Profile_Data.xlsx
#' 3.4_WorkBC_Career_Compass.xlsx
#' 3.5_WorkBC_Career_Trek.xlsx
#' 3.7_WorkBC_Buildprint_Builder.xlsx
#' BC Population Distribution.xlsx
#' Career Search Tool Job Openings.xlsx
#' HOO BC and Region for new tool.xlsx
#' HOO List.xlsx
tictoc::tic()
# LIBRARIES--------------
library("tidyverse")
library("readxl")
library("XLConnect")
library("here")
library("lubridate")

# "Constants"----------
current_year <- as.numeric(year(today()))
previous_year <- current_year - 1
current_plus_5 <- current_year + 5
current_plus_10 <- current_year + 10

# FUNCTIONS----------------
source(here("R", "functions.R"))

# Load the data sets---------------

occ_char <- read_excel(here("raw_data", list.files(here("raw_data"), pattern = "occupational_characteristics")), skip = 3) %>%
  rename(typical_education_background = starts_with("Typical")) %>%
  mutate(typical_education_background = str_replace(typical_education_background, "Excluding Apprenticeship", "")) %>%
  clean_tbbl() %>%
  select(
    noc = noc_2016,
    typical_education_background = starts_with("typical"),
    occupational_interests = interests,
    skills_and_compentencies = `skills:_top_3`,
    first = skill1,
    second = skill2,
    third = skill3,
    typical_education_background = starts_with("typical")
  )

whos_hoo <- read_excel(here("raw_data", list.files(here("raw_data"), pattern = "hoo_list"))) %>%
  clean_tbbl() %>%
  filter(`high_opportunity_occupation?` == "yes") %>%
  select(noc, geographic_area)

pop <- read_excel(here("raw_data", list.files(here("raw_data"), pattern = "pop_municipal")),
  skip = 2,
  sheet = "Dev-RD", n_max = 45
) %>%
  select(Name, `2021`)

career_trek <- read_excel(here("raw_data", list.files(here("raw_data"), pattern = "career_trek_jobs"))) %>%
  clean_tbbl()

mapping <- read_csv(here("raw_data", list.files(here("raw_data"), pattern = "lmo64_characteristics"))) %>%
  select(industry = lmo_detailed_industry, aggregate_industry) %>%
  distinct()

empl_bc_regions_industry <- read_excel(here(
  "raw_data",
  list.files(here("raw_data"),
    pattern = "empl_bc_regions_industry"
  )
),
skip = 3
) %>%
  clean_tbbl()

job_openings_bc_regions_industry <- read_excel(here(
  "raw_data",
  list.files(here("raw_data"),
    pattern = "job_openings_bc_regions_industry"
  )
),
skip = 3
) %>%
  clean_tbbl()

lmo_work_bc_unemployment <- read_excel(here(
  "raw_data",
  list.files(here("raw_data"),
    pattern = "unemployment_rate"
  )
),
skip = 3
) %>%
  clean_tbbl()

cstjo <- read_csv(here("raw_data", list.files(here("raw_data"),
  pattern = "career_search_tool_job_openings"
)))

wages_bc <- read_excel(here("raw_data", list.files(here("raw_data"), pattern = "wages"))) %>%
  clean_tbbl() %>%
  rename(
    noc = noc_2016,
    low_wage = contains("low"),
    median_wage = contains("median"),
    high_wage = contains("high")
  ) %>%
  mutate(
    low_wage = round(low_wage, 2),
    median_wage = round(median_wage, 2),
    high_wage = round(high_wage, 2),
    median_annual_salary = ifelse(median_wage < 1000,
      round(median_wage * 2085.6, 2),
      median_wage
    )
  )

# process the data--------------
# make one long df--------
long_emp <- pl_wrap(empl_bc_regions_industry) %>%
  filter(year != previous_year) # empl_bc_regions_industry contains previous year data
long_jo <- pl_wrap(job_openings_bc_regions_industry)
long_un <- pl_wrap(lmo_work_bc_unemployment)
industry_mapping <- mapping %>%
  rbind(c("all_industries", "all"))

long <- bind_rows(long_emp, long_jo) %>%
  mutate(value = round(value, -1)) %>% # employment and job openings data rounded to nearest 10
  bind_rows(long_un) %>% # unemployment rates NOT rounded
  full_join(industry_mapping, by = "industry")

# Career Search Tool Job Openings
current_jo <- long_jo %>%
  filter(
    variable == "job_openings",
    noc != "#t",
    year > current_year
  ) %>%
  select(-year, -variable) %>%
  group_by(noc, description, industry, geographic_area) %>%
  summarize(job_openings = round(sum(value, na.rm = TRUE), -1))

median_wages_bc <- wages_bc %>%
  select(noc, median_wage)

cstjo_with_jo_and_wages <- full_join(cstjo, current_jo, by = c(
  "noc" = "noc",
  "noc_name" = "description",
  "industry_(sub-industry)" = "industry",
  "region" = "geographic_area"
)) %>%
  full_join(median_wages_bc, by = "noc") %>%
  select(
    noc,
    noc_name,
    `short_noc_name_(not_used_on_the_tool)`,
    `industry_(sub-industry)`,
    region,
    education,
    link,
    jobbank2,
    `industry_(aggregate)`,
    `part-time/full-time`,
    job_openings,
    median_wage
  ) %>%
  relocate("job_openings_{current_year+1}-{current_plus_10}" := job_openings, .after = region) %>%
  relocate(`salary_(calculated_median_salary)` = median_wage, .after = education) %>%
  camel_to_title() %>%
  na.omit()

colnames(cstjo_with_jo_and_wages) <- make_sentence(colnames(cstjo_with_jo_and_wages)) %>%
  str_replace_all("noc", "NOC") %>%
  str_replace_all("Noc", "NOC") %>%
  str_replace_all("Job openings", "Job Openings")

# 3.3.1 CAREER PROFILE PROVINCIAL---------------
#' This is the "by NOC" breakdown of labour market.
#' 1. Suppression rule:
#'   If the current year's employment is below 20,
#'   then all variables should be changed to "N/A"
#'
provincial_career_profiles <- long %>%
  filter(
    geographic_area == "british_columbia",
    industry == "all_industries",
    noc != "#t"
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
  select(-current_employment, -cagr_ten_year) %>%
  camel_to_title()

# 3.3.1 CAREER PROFILE REGIONAL----------
#' This is the "by NOC and geographic_area" breakdown of the labour market.
#' 1. Suppression rule:
#'   If the current year's employment is below 20, then all variables should be changed to "N/A"
career_profile_regional <- long %>%
  filter(
    industry == "all_industries",
    noc != "#t"
  ) %>%
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
  select(!contains("british_columbia"), everything()) %>%
  camel_to_title()

career_profile_regional_excel <- career_profile_regional %>%
  add_column(` ` = "", .after = "cariboo_ten_year_job_openings") %>%
  add_column(`  ` = "", .after = "kootenay_ten_year_job_openings") %>%
  add_column(`   ` = "", .after = "mainland_south_west_ten_year_job_openings") %>%
  add_column(`    ` = "", .after = "north_coast_&_nechako_ten_year_job_openings") %>%
  add_column(`     ` = "", .after = "north_east_ten_year_job_openings") %>%
  add_column(`      ` = "", .after = "thompson_okanagan_ten_year_job_openings") %>%
  select(-contains("british_columbia"))

# 3.3.2 INDUSTRY PROFILE-------------
# This is the "by aggregate_industry" breakdown of the labour market.

jo_by_industry <- long %>%
  filter(
    year>current_year,
    geographic_area == "british_columbia",
    description == "total"
  ) %>%
  select(-geographic_area, -description) %>%
  group_by(aggregate_industry) %>%
  nest() %>%
  mutate(data = map(data, aggregate_by_year, "job_openings"))%>%
  unnest(data)%>%
  summarize(job_openings=sum(value))%>%
  rename(industry = aggregate_industry) %>%
  filter(industry!="all")%>%
  arrange(industry) %>%
  camel_to_title()

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
  filter(aggregate_industry == "all") %>% # check to make sure df not empty
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
  filter(industry!="all")%>%
  arrange(industry) %>%
  camel_to_title()

industry_profile <- inner_join(jo_by_industry, industry_profile)

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
  select(-data) %>%
  camel_to_title() %>%
  select(
    geographic_area,
    ten_year_sum_replacement_demand,
    replacement_percent,
    ten_year_sum_expansion_demand,
    expansion_percent
  ) %>%
  arrange(geographic_area)

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
  unnest(employment, names_sep = "_") %>%
  select(-data) %>%
  camel_to_title()

# join composition_job_openings with employment outlook
jo_and_eo <- full_join(composition_job_openings,
  employment_outlook,
  by = "geographic_area"
) %>%
  select(
    geographic_area,
    ten_year_sum_replacement_demand,
    replacement_percent,
    ten_year_sum_expansion_demand,
    expansion_percent,
    employment_current,
    employment_five,
    employment_ten,
    ten_sum_jo,
    ten_year_cagr
  )

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
  slice_max(order_by = ten_year_job_openings, n = 10, with_ties = FALSE) %>%
  camel_to_title()

# need region headings and spacing
industries_top_ten_excel <- regional_profile_maindf2 %>%
  group_by(geographic_area) %>%
  nest() %>%
  mutate(data = map2(data, geographic_area, add_foot_head, num_col = 3)) %>%
  ungroup() %>%
  select(-geographic_area) %>%
  unnest(data)

# need region headings and spacing
industries_no_bc <- regional_profile_maindf2 %>%
  select(-ten_year_cagr) %>%
  group_by(geographic_area) %>%
  nest() %>%
  mutate(data = map2(data, geographic_area, add_foot_head2, num_col = 2)) %>%
  ungroup() %>%
  filter(geographic_area != "British Columbia") %>%
  select(-geographic_area) %>%
  unnest(data)

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
  slice_max(order_by = ten_year_job_openings, n = 10, with_ties = FALSE) %>%
  camel_to_title()

# need region headings and spacing
occupations_top_ten_excel <- occupations_top_ten %>%
  group_by(geographic_area) %>%
  nest() %>%
  mutate(data = map2(data, geographic_area, add_foot_head, num_col = 4)) %>%
  ungroup() %>%
  select(-geographic_area) %>%
  unnest(data)

# need region headings and spacing
occupations_no_bc <- occupations_top_ten %>%
  select(-ten_year_cagr) %>%
  group_by(geographic_area) %>%
  nest() %>%
  mutate(data = map2(data, geographic_area, add_foot_head2, num_col = 3)) %>%
  ungroup() %>%
  filter(geographic_area != "British Columbia") %>%
  select(-geographic_area) %>%
  unnest(data)


# 3.4 CAREER COMPASS - BROWSE CAREERS---------
# This is the "by noc" breakdown of the labour market.
lmo_career_compass <- long %>%
  filter(
    geographic_area == "british_columbia",
    industry == "all_industries",
    noc != "#t"
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
  ungroup() %>%
  select(
    noc,
    description,
    unemployment_current,
    unemployment_five,
    unemployment_ten,
    cagr_ten_year,
    ten_year_job_openings,
    jo_current,
    jo_five,
    jo_ten
  )

# 3.5 CAREER TREK-----------
career_trek_merged <- lmo_career_compass %>%
  select(-contains("unemployment"), -contains("jo_")) %>%
  right_join(career_trek, by = c("noc" = "noc", "description" = "noc_title")) %>%
  camel_to_title() %>%
  select(
    sr_no,
    noc,
    description,
    starts_with("occupation"),
    cagr_ten_year,
    ten_year_job_openings
  )

# BC Population Distribution--------------
pop_summary <- pop %>%
  filter(Name %in% c(
    "Vancouver Island/Coast",
    "Mainland/Southwest",
    "Thompson/Okanagan",
    "Kootenay",
    "Cariboo",
    "North Coast",
    "Nechako",
    "Northeast",
    "British Columbia"
  )) %>%
  group_by(Name) %>%
  summarize(`2021` = max(`2021`)) %>% # kootenay is both a region and a subregion... this gets the pop for the region.
  arrange(desc(Name))

# top 10 careers by aggregate industry----------------

long_emp_current_bc <- long_emp %>%
  filter(
    year == current_year,
    geographic_area == "british_columbia",
    noc != "#t",
    industry != "all_industries"
  ) %>%
  select(-variable, -year, -geographic_area)

emp_current_with_agg <- long_emp_current_bc %>%
  full_join(mapping, by = "industry")%>%
  group_by(aggregate_industry, noc, description)%>% 
  summarise(value=sum(value))%>% #aggregates data at industry level to aggregate industry level
  group_by(aggregate_industry, .add=FALSE)%>% #to get top 10 by aggregate industry
  slice_max(value, n = 10, with_ties = FALSE) %>%
  mutate(value = round(value, 0)) %>%
  arrange(desc(aggregate_industry)) %>%
  select(
    `Aggregate Industry` = aggregate_industry,
    NOC = noc,
    Occupation = description,
    "Employment-{current_year}" := value
  ) %>%
  camel_to_title()

# HOO BC and Region for new tool----------------

ten_year_jo <- long_jo %>%
  filter(
    variable == "job_openings",
    noc != "#t",
    industry == "all_industries",
    year > current_year
  ) %>%
  group_by(noc, geographic_area) %>%
  summarize(job_openings = round(sum(value), 0))

hoo_new_tool <- left_join(whos_hoo, wages_bc, by = "noc") %>%
  left_join(ten_year_jo) %>%
  left_join(occ_char) %>%
  select(
    noc_title = description,
    job_openings,
    low_wage,
    median_wage,
    high_wage,
    median_annual_salary,
    occupational_interests,
    skills_and_compentencies,
    first,
    second,
    third,
    `#noc` = noc,
    typical_education_background,
    geographic_area
  ) %>%
  mutate(noc = str_replace(`#noc`, "#", ""), .before = `#noc`) %>%
  camel_to_title()

# HOO list----------

whos_hoo_bc <- whos_hoo %>%
  filter(geographic_area == "british_columbia")

ten_year_jo_bc <- long_jo %>%
  filter(
    variable == "job_openings",
    noc != "#t",
    industry == "all_industries",
    year > current_year,
    geographic_area == "british_columbia"
  ) %>%
  group_by(noc) %>%
  summarize(job_openings = round(sum(value), 0))

top_hoo_by_educ <- left_join(whos_hoo_bc, wages_bc, by = "noc") %>%
  left_join(ten_year_jo_bc) %>%
  left_join(occ_char) %>%
  select(
    noc_title = description,
    job_openings,
    low_wage,
    median_wage,
    high_wage,
    occupational_interests,
    skills_and_compentencies,
    noc,
    typical_education_background
  ) %>%
  mutate(typical_education_background = ordered(typical_education_background,
    levels = c(
      "degree",
      "diploma/certificate",
      "apprenticeship_certificate",
      "high_school"
    )
  )) %>%
  group_by(typical_education_background) %>%
  arrange(desc(job_openings), .by_group = TRUE)

# add some text to top of HOO list sheet
counts <- top_hoo_by_educ %>%
  janitor::tabyl(typical_education_background)%>%
  summarize(smushed = paste0(n, " ", typical_education_background, "; ", collapse = ""))%>%
  mutate(
    smushed = str_replace_all(smushed, "_", " "),
    smushed = str_to_title(smushed),
    smushed = str_sub(smushed, end = -3)
  )
text <- paste0("LMO ", current_year, " Edition HOO Top ", nrow(top_hoo_by_educ), " (", counts, ")")

# adds education level headers to sheet
top_hoo_by_educ_headers <- top_hoo_by_educ %>%
  camel_to_title() %>% # converts factors to characters (cant add headers to factors)
  group_by(typical_education_background) %>%
  nest() %>%
  mutate(data = map2(data, typical_education_background, add_header)) %>%
  ungroup() %>%
  select(-typical_education_background) %>%
  unnest(data)

# Write to file-----------------
# export 3.3.1
wb <- loadWorkbook(here("templates", "3.3.1_WorkBC_Career_Profile_Data.xlsx"))
write_workbook(provincial_career_profiles, "Provincial Outlook", 4, 1)
write_workbook(career_profile_regional_excel, "Regional Outlook", 5, 1)
saveWorkbook(wb, here("processed_data", paste0(
  "3.3.1_WorkBC_Career_Profile_Data_",
  current_year,
  "-",
  current_plus_10,
  ".xlsx"
)))
# export 3.3.2
wb <- loadWorkbook(here("templates", "3.3.2_WorkBC_Industry_Profile.xlsx"))
write_workbook(industry_profile, "BC", 4, 1)
saveWorkbook(wb, here("processed_data", paste0(
  "3.3.2_WorkBC_Industry_Profile_",
  current_year,
  "-",
  current_plus_10,
  ".xlsx"
)))
# export 3.3.3
wb <- loadWorkbook(here("templates", "3.3.3_WorkBC_Regional_Profile_Data.xlsx"))
write_workbook(jo_and_eo, "Regional Profiles - LMO", 5, 1)
write_workbook(occupations_top_ten_excel, "Top Occupation", 4, 1)
write_workbook(industries_top_ten_excel, "Top Industries", 4, 1)
saveWorkbook(wb, here("processed_data", paste0(
  "3.3.3_WorkBC_Regional_Profile_Data_",
  current_year,
  "-",
  current_plus_10,
  ".xlsx"
)))
# export 3.4

occupations_industries_no_bc <- bind_cols(occupations_no_bc, industries_no_bc) %>%
  mutate(" " = "", .before = "aggregate_industry")
wb <- loadWorkbook(here("templates", "3.4_WorkBC_Career_Compass.xlsx"))
write_workbook(camel_to_title(lmo_career_compass), "Browse Careers", 4, 1)
write_workbook(occupations_industries_no_bc, "Regions", 4, 1)
saveWorkbook(wb, here("processed_data", paste0(
  "3.4_WorkBC_Career_Compass_",
  current_year,
  "-",
  current_plus_10,
  ".xlsx"
)))
# export 3.5
wb <- loadWorkbook(here("templates", "3.5_WorkBC_Career_Trek.xlsx"))
write_workbook(career_trek_merged, "LMO", 2, 1)
saveWorkbook(wb, here("processed_data", paste0(
  "3.5_WorkBC_Career_Trek_",
  current_year,
  "-",
  current_plus_10,
  ".xlsx"
)))
# export 3.7
wb <- loadWorkbook(here("templates", "3.7_WorkBC_Buildprint_Builder.xlsx"))
write_workbook(composition_job_openings, "Regional Data", 5, 1)
write_workbook(select(occupations_top_ten, description, noc), "Jobs in Demand", 4, 2)
saveWorkbook(wb, here("processed_data", paste0(
  "3.7_WorkBC_Buildprint_Builder_",
  current_year,
  "-",
  current_plus_10,
  ".xlsx"
)))

# export population
wb <- loadWorkbook(here("templates", "BC Population Distribution.xlsx"))
write_workbook(pop_summary, "Region Population Estimates", 3, 1)
write_workbook(pop, "Regional District Population", 5, 1)
saveWorkbook(wb, here("processed_data", paste0(current_year - 1, " BC Population Distribution.xlsx")))

# export top 10 careers by aggregate industry
openxlsx::write.xlsx(emp_current_with_agg, here(
  "processed_data",
  paste(current_year, "top_10_careers_by_aggregate_industry.xlsx")
))

# export HOO BC and Region for new tool
wb <- loadWorkbook(here("templates", "HOO BC and Region for new tool.xlsx"))
write_workbook(hoo_new_tool, "Sheet1", 2, 1)
saveWorkbook(wb, here(
  "processed_data",
  paste(current_year, "HOO BC and Region for new tool.xlsx")
))

# export HOO list
wb <- loadWorkbook(here("templates", "HOO List.xlsx"))
write_workbook(text, "Sheet1", 1, 1)
write_workbook(top_hoo_by_educ_headers, "Sheet1", 5, 1)
saveWorkbook(wb, here(
  "processed_data",
  paste(current_year, "HOO List.xlsx")
))

# export Career Search Tool Job Openings
workbook <- openxlsx::createWorkbook()
openxlsx::addWorksheet(workbook, "Sheet1")
bold.style <- openxlsx::createStyle(textDecoration = "Bold")
openxlsx::writeData(workbook, "Sheet1", cstjo_with_jo_and_wages,
  startRow = 1, startCol = 1,
  headerStyle = bold.style
)
openxlsx::saveWorkbook(workbook, here(
  "processed_data",
  paste(current_year, "Career Search Tool Job Openings.xlsx")
),
overwrite = TRUE
)
tictoc::toc()