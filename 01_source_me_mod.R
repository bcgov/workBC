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
#library("XLConnect")
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
career_trek <- read_excel(here("Data Sources",
                               list.files(here("Data Sources"),
                                          pattern = "newtemplate")))%>%
  clean_tbbl()
empl_bc_regions_industry <- read_excel(here("Data Sources",
                                            list.files(here("Data Sources"),
                                                       pattern = "Empl_BC_Regions_Industry")),
                                       skip=3) %>%
  clean_tbbl()
industry_profiles <- read_excel(here("Data Sources",
                                     list.files(here("Data Sources"),
                                                pattern = "Industry_Profiles")))%>%
  clean_tbbl() 
job_openings_bc_regions_industry <- read_excel(here("Data Sources", 
                                                   list.files(here("Data Sources"), 
                                                              pattern = "JobOpenings_BC_Regions_Industry")),
                                              skip=2)%>%
  clean_tbbl() 
lmo_work_bc_unemployment <- read_excel(here("Data Sources", 
                                           list.files(here("Data Sources"), 
                                                      pattern = "Unemployment_rate")),
                                      skip=3)%>%
  clean_tbbl()

# 3.3.1 CAREER PROFILE PROVINCIAL---------------
#' 1. Suppression rule:
#'   If the current year's employment is below 20, 
#'   then all variables should be changed to "N/A"

bc <- empl_bc_regions_industry%>%
  filter(geographic_area == "british_columbia",
         industry == "all_industries")%>%
  select(-geographic_area, -industry, -variable)%>%
  pivot_longer(cols= -c(noc, description),
               names_to = "year", 
               values_to = "value")%>%
  group_by(noc, description)%>%
  nest()%>%
  mutate(cagrs = map(data, get_cagrs, all=TRUE))%>%
  unnest(cagrs)%>%
  mutate(current_employment=map_dbl(data, current_jobs))%>%
  select(-data)

# Job Openings---------------

job_openings <- job_openings_bc_regions_industry%>%
  filter(geographic_area == "british_columbia",
         industry == "all_industries")%>%
  select(-geographic_area, -industry)%>%
  pivot_longer(cols= -c(noc, description, variable),
               names_to = "year", 
               values_to = "value")%>%
  mutate(value=round(value, -1))%>%
  group_by(noc, description)%>%
  nest()%>%
  mutate(jo_select_years = map(data, jo_select),
         ten_year_job_openings = map_dbl(data, ten_sum, "job_openings"),
         ten_year_sum_expan = map_dbl(data, ten_sum, "expansion_demand"),
         ten_year_sum_replace = map_dbl(data, ten_sum, "replacement_demand"),
         `expansion_%` = round(ten_year_sum_expan/(ten_year_sum_expan+ten_year_sum_replace),3)*100,
         `replacement_%` = 100-`expansion_%`
         )%>%
  select(-data)%>%
  unnest(jo_select_years)

provincial_career_profiles <- full_join(bc, job_openings, by = c("noc", "description"))%>%
  relocate(`replacement_%`, .after = ten_year_job_openings)%>%
  relocate(ten_year_sum_replace, .after = `replacement_%`)%>%
  relocate(`expansion_%`, .after = ten_year_sum_replace)%>%
  mutate(`expansion_%` = if_else(is.na(`expansion_%`), 0, `expansion_%`),
         `replacement_%` = if_else(is.na(`replacement_%`), 0, `replacement_%`),
         `expansion_%` = if_else(`expansion_%` < 0, 0,`expansion_%`),
         ten_year_sum_replace = if_else(ten_year_sum_expan < 0 & `replacement_%` > 0, ten_year_sum_replace + ten_year_sum_expan, ten_year_sum_replace),
         ten_year_sum_expan = if_else(ten_year_sum_expan < 0, 0, ten_year_sum_expan),
         `expansion_%` = if_else(`expansion_%` > 100, 0, `expansion_%`),
         `replacement_%` = if_else(`replacement_%` < 0, 100, `replacement_%`))%>%
  mutate(across(where(is.numeric), ~ if_else(current_employment<20, NA_real_, .x)))%>%
  select(-current_employment)

# 3.3.1 CAREER PROFILE REGIONAL----------
#' 1. Suppression rule:
#'   If the current year's employment is below 20, then all variables should be changed to "N/A"

regions <- empl_bc_regions_industry%>%
  filter(industry == "all_industries")%>%
  select(-industry, -variable)%>%
  pivot_longer(cols= -c(noc, description, geographic_area),
               names_to = "year", 
               values_to = "value")%>%
  mutate(value=round(value, -1))%>%
  group_by(noc, description, geographic_area)%>%
  nest()%>%
  mutate(employment = map_dbl(data, current_jobs),
    "avg_employ_growth_{current_year}_{second_five_years}" := map_dbl(data, get_cagrs, all=FALSE))%>%
  select(-data)

regional_openings <- job_openings_bc_regions_industry%>%
  filter(industry == "all_industries")%>%
  select(-industry)%>%
  pivot_longer(cols= -c(noc, description, variable, geographic_area),
               names_to = "year", 
               values_to = "value")%>%
  mutate(value=round(value, -1))%>%
  group_by(noc, description, geographic_area)%>%
  nest()%>%
  mutate(ten_year_job_openings = map_dbl(data, ten_sum, "job_openings"))%>%
  select(-data)

id_cols=c("noc", "description", "geographic_area")
career_profile_regional <- full_join(regions,
  regional_openings,
  by = c("noc", "description", "geographic_area")) %>%
  arrange(geographic_area)%>%
  mutate(across(where(is.numeric), ~ if_else(employment < 20, NA_real_, .x)))%>%
  rename("employment_{current_year}" := employment)%>%
pivot_wider(names_from = geographic_area, 
              values_from = -all_of(id_cols),
              names_glue = "{geographic_area}_{.value}",
              names_vary = "slowest")%>%
  select(!contains("british_columbia"), everything())






# 3.3.2 INDUSTRY PROFILE-------------
geo <- "British Columbia"
# Subset out the total employment for each industry
temp <- subset(Empl_BC_Regions_Industry, `Geographic Area` == geo & Description == "Total")
# Using the standardized industry names, merge the employment data
Industry_Profile <- merge(Industry_Profiles, temp, by = "Industry", all.x = TRUE)
# Subsetting out the employment values for the 10 year forecast
Industry_Profile <- Industry_Profile[, c(1, (ncol(Industry_Profile) - 10):(ncol(Industry_Profile)))]
# Assigning Industry as the rowname
rownames(Industry_Profile) <- Industry_Profile[, 1]
Industry_Profile <- Industry_Profile[, -1]
# Find the aggregate industry values using custom function
Industry_Profile <- aggregate_industries(Industry_Profile)
# Calculating the industry share of total employment
# Current year share
Industry_Profile[, paste0(current_year, " share of employment")] <- round(prop.table(Industry_Profile[, current_year]) * 100, 1)
# Fifth year share
Industry_Profile[, paste0(first_five_years, " share of employment")] <- round(prop.table(Industry_Profile[, first_five_years]) * 100, 1)
# Tenth year share
Industry_Profile[, paste0(second_five_years, " share of employment")] <- round(prop.table(Industry_Profile[, second_five_years]) * 100, 1)

# Forecasted Employment
# Current year
Industry_Profile[, paste0(current_year, " forecasted employment")] <- round(Industry_Profile[, current_year], -2)
# Fifth year
Industry_Profile[, paste0(first_five_years, " forecasted employment")] <- round(Industry_Profile[, first_five_years], -2)
# Tenth year
Industry_Profile[, paste0(second_five_years, " forecasted employment")] <- round(Industry_Profile[, second_five_years], -2)
# Calculate CAGRs
Industry_Profile <- CAGR(Industry_Profile)
# Creating a column for the industry names
Industry_Profile[, 1] <- row.names(Industry_Profile)
colnames(Industry_Profile)[1] <- "Industry"
# Extracting the final data
Industry_Profile <- Industry_Profile[, c(1, 12:ncol(Industry_Profile))]
# Naming column
colnames(Industry_Profile)[1] <- "Industry"
# Removing rownames
rownames(Industry_Profile) <- NULL
# Ordering the data to the LMO industry order
Industry_Profile <- with(Industry_Profile, Industry_Profile[order(Industry, decreasing = FALSE), ])
# 3.3.3 REGIONAL PROFILE - REGIONAL PROFILES LMO-------------
# COMPOSITION OF JOB OPENINGS SECTION - REGIONAL PROFILES LMO
# Copy the data that I've already produced
composition_job_openings <- subset(
  JobOpenings_BC_Regions_Industry,
  Variable %in% c("Expansion Demand", "Replacement Demand") &
    Industry == "All industries" &
    Description == "Total" &
    `Geographic Area` %in% c(
      "Cariboo",
      "Kootenay",
      "Mainland South West",
      "North Coast & Nechako",
      "North East",
      "Thompson Okanagan",
      "Vancouver Island Coast",
      "British Columbia"
    )
)
# Find 10 year sum of job openings
composition_job_openings$ten_year_sum <- apply(composition_job_openings[, 7:16], 1, sum)
# Finding the replacement and expansion percentages
composition_job_openings <- setDT(composition_job_openings)[, `Replacement %` :=
  round((ten_year_sum[Variable == "Replacement Demand"] /
    (ten_year_sum[Variable == "Replacement Demand"] +
      ten_year_sum[Variable == "Expansion Demand"])) * 100, 1),
by = .(`Geographic Area`)
]
composition_job_openings <- setDT(composition_job_openings)[, `Expansion %` :=
  round((ten_year_sum[Variable == "Expansion Demand"] /
    (ten_year_sum[Variable == "Replacement Demand"] +
      ten_year_sum[Variable == "Expansion Demand"])) * 100, 1),
by = .(`Geographic Area`)
]
composition_job_openings <- composition_job_openings[, c(
  "Geographic Area",
  "Variable",
  "ten_year_sum",
  "Replacement %",
  "Expansion %"
)]
# Widen dataset
composition_job_openings <- dcast(composition_job_openings, `Geographic Area` ~ Variable,
  value.var = c(
    "ten_year_sum",
    "Replacement %",
    "Expansion %"
  )
)
# Selecting out the final data
composition_job_openings <- composition_job_openings[, c(
  "Geographic Area",
  "ten_year_sum_Replacement Demand",
  "Replacement %_Expansion Demand",
  "ten_year_sum_Expansion Demand",
  "Expansion %_Expansion Demand"
)]
# Rounding the values
composition_job_openings[, c(2, 4)] <- lapply(composition_job_openings[, c(2, 4)], function(x) {
  round(x, -2)
})
# EMPLOYMENT OUTLOOK SECTION - REGIONAL PROFILES LMO
employment_outlook <- subset(
  Empl_BC_Regions_Industry,
  Variable %in% c("Employment") &
    Industry == "All industries" &
    Description == "Total" &
    `Geographic Area` %in% c(
      "Cariboo",
      "Kootenay",
      "Mainland South West",
      "North Coast & Nechako",
      "North East",
      "Thompson Okanagan",
      "Vancouver Island Coast",
      "British Columbia"
    )
)
# Find 10 year sums by merging Job Openings 10 year sum to employment
jobopenings <- subset(
  JobOpenings_BC_Regions_Industry,
  Variable %in% c("Job Openings") &
    Industry == "All industries" &
    Description == "Total" &
    `Geographic Area` %in% c(
      "Cariboo",
      "Kootenay",
      "Mainland South West",
      "North Coast & Nechako",
      "North East",
      "Thompson Okanagan",
      "Vancouver Island Coast",
      "British Columbia"
    )
)

jobopenings <- as.data.frame(cbind(jobopenings[, "Geographic Area"], round(apply(jobopenings[, 7:16], 1, sum), -2)))
colnames(jobopenings) <- c("Geographic Area", "ten_year_sum")
# Merging employment and job openings
employment_outlook <- merge(employment_outlook, jobopenings, by = "Geographic Area")
# Calculate CAGRs
employment_outlook <- CAGR(employment_outlook)
# Selecting out the final data needed
employment_outlook <- employment_outlook[, c(
  current_year,
  first_five_years,
  second_five_years,
  "ten_year_sum",
  paste0(current_year, "-", second_five_years)
)]
# Convert to numeric
employment_outlook$ten_year_sum <- as.numeric(as.character(employment_outlook$ten_year_sum))
# Rounding the values
employment_outlook[, c(1:4)] <- lapply(employment_outlook[, c(1:4)], function(x) {
  round(as.numeric(x), -2)
})
# 3.3.3 REGIONAL PROFILE - TOP 10 INDUSTRIES---------------
# Looping over each region to find their aggregated industries-----------
count <- 0
for (geo in region_list) {
  # Find the job openings data for each region
  attach(JobOpenings_BC_Regions_Industry)
  temp_regional_profile <- as.data.frame(JobOpenings_BC_Regions_Industry[Variable %in% c("Job Openings") &
    Description == "Total" &
    `Geographic Area` %in% geo, ])

  # Subset the data even further
  temp_regional_profile <- droplevels(temp_regional_profile)
  temp_regional_profile <- temp_regional_profile[, c(3, 7:16)]
  # Assigning Industry as the rowname
  rownames(temp_regional_profile) <- temp_regional_profile[, 1]
  temp_regional_profile <- temp_regional_profile[, -1]
  # Calculate the aggregate industries
  temp_regional_profile <- aggregate_industries(temp_regional_profile)
  # Calculating ten year job openings
  temp_regional_profile$ten_year_job_openings <- round(apply(temp_regional_profile[, 1:10], 1, sum), -2)
  # GETTING THE DATA INTO THE RIGHT FORMAT AND JOINING THE CAGR DATA FROM THE INDUSTRY PROFILES
  temp_regional_profile[, 1] <- row.names(temp_regional_profile)
  colnames(temp_regional_profile)[1] <- "Industry"
  # Extracting the final data
  temp_regional_profile <- temp_regional_profile[, c(1, 11:ncol(temp_regional_profile))]
  colnames(temp_regional_profile)[1] <- "Industry"
  rownames(temp_regional_profile) <- NULL
  # Ordering the temp_regional data by Industry Name
  temp_regional_profile <- with(temp_regional_profile, temp_regional_profile[order(Industry, decreasing = FALSE), ])
  temp_regional_profile[, paste0("avg_employ_growth_", current_year, "_", second_five_years)] <-
    Industry_Profile[, paste0(current_year, "-", second_five_years)]
  temp_regional_profile$Region <- geo

  # save the data
  if (count == 0) {
    Regional_Profile_maindf <- temp_regional_profile
  } else {
    Regional_Profile_maindf <- rbind(Regional_Profile_maindf, temp_regional_profile)
  }

  count <- count + 1
}
# end of looping over regions-------------
# Finding the top ten industries for the regional profiles
Regional_Profile_maindf2 <- setDT(Regional_Profile_maindf)[order(ten_year_job_openings, decreasing = TRUE), head(.SD, 10), by = Region]
# Sort by region
Regional_Profile_maindf2 <- Regional_Profile_maindf2[order(Regional_Profile_maindf2$Region), ]

# 3.3.3 REGIONAL PROFILE - TOP 10 OCCUPATIONS-----------
#' 1. The top 10 occupations relies upon the top 10 industries to be run first

for (data in region_list) {
  # Subsetting BC and All Industries
  attach(JobOpenings_BC_Regions_Industry)
  assign(
    data,
    subset(
      JobOpenings_BC_Regions_Industry,
      `Geographic Area` == data &
        Industry == "All industries" &
        Variable == "Job Openings"
    )
  )
}
top_ten_occupations <- data.frame()
top_ten_industries <- data.frame()
Occupations_top_10 <- data.frame()
industry_top_10 <- data.frame()

# Second loop to perform the calculations
for (data in region_list) {
  temp <- get(data)
  # Subsetting out the Job Openings data
  temp2 <- subset(
    JobOpenings_BC_Regions_Industry,
    `Geographic Area` == data &
      Industry == "All industries" &
      Variable == "Job Openings"
  )
  # Finding the sum of the job openings over each year
  temp2$ten_year_job_openings <- apply(temp2[, 7:16], 1, sum)
  # Subset All Industries from the employment data
  temp4 <- subset(
    Empl_BC_Regions_Industry,
    `Geographic Area` == data &
      Industry == "All industries"
  )
  # Calculate the CAGR
  temp_regional_profile[, paste0("avg_employ_growth_", current_year, "_", second_five_years)] <-
    Industry_Profile[, paste0(current_year, "-", second_five_years)]
  # Calculating the 10 year CAGR
  temp4[, paste("avg_employ_growth", current_year, second_five_years, sep = "_")] <-
    round((temp4[, second_five_years] / temp4[, current_year])^(1 / 10) -
      1, digits = 3) * 100
  # Convert any NANs to zeros
  # temp4[, paste0("avg_employ_growth_2021_2031")] =
  # ifelse(is.na(temp4[, paste0("avg_employ_growth_2021_2031")]),
  #       "NA",
  #      temp4[, paste0("avg_employ_growth_2021_2031")])

  # Merge CAGR to job openings (temp2)
  temp2 <- merge(temp2, temp4[, colnames(temp4) %in% c(
    "NOC",
    paste0("avg_employ_growth_", current_year, "_", second_five_years)
  )])
  # Subsetting the data to keep just the NOC and ten year job openings
  temp2 <- temp2[, c(1, length(temp2) - 1, length(temp2))]
  # merge the temp2 data to temp
  temp_occ <- merge(temp, temp2, by = "NOC", all.x = TRUE)
  # Finding the top 10 occupations
  temp_occ <- head(temp_occ[order(temp_occ$ten_year_job_openings, decreasing = TRUE), ], n = 11)
  # Combine all of the top 10 occupations from each region together
  Occupations_top_10 <- rbind(Occupations_top_10, temp_occ, fill = TRUE)
}
### checking
# Remove duplicates
Occupations_top_10 <- Occupations_top_10[!duplicated(Occupations_top_10), ]
Occupations_top_10 <- na.omit(Occupations_top_10)

attach(Occupations_top_10)
Occupations_top_ten <- Occupations_top_10[
  Description != "Total",
  colnames(Occupations_top_10) %in% c(
    "Geographic Area",
    "NOC",
    "Description",
    "ten_year_job_openings",
    paste0("avg_employ_growth_", current_year, "_", second_five_years)
  )
]
Occupations_top_ten$ten_year_job_openings <- round(Occupations_top_ten$ten_year_job_openings, -1)
# Reordering the columns
Occupations_top_ten <- Occupations_top_ten[, c(
  "Geographic Area",
  "NOC",
  "Description",
  "ten_year_job_openings",
  paste0("avg_employ_growth_", current_year, "_", second_five_years)
)]

# 3.4 CAREER COMPASS - BROWSE CAREERS---------
# Calculate the unemployment rate for the current year
LMO_WorkBC_Unemployment[, paste0("Unemployment_", current_year)] <-
  round(LMO_WorkBC_Unemployment[, paste0(current_year)], 1)
# Fifth year unemployment rate
LMO_WorkBC_Unemployment[, paste0("Unemployment_", first_five_years)] <-
  round(LMO_WorkBC_Unemployment[, paste0(first_five_years)], 1)
# Tenth year unemployment rate
LMO_WorkBC_Unemployment[, paste0("Unemployment_", second_five_years)] <-
  round(LMO_WorkBC_Unemployment[, paste0(second_five_years)], 1)
# Adding the CAGR and job openings data to the unemployment rate data
LMO_WorkBC_Unemployment <- merge(LMO_WorkBC_Unemployment, BC[, colnames(BC) %in% c(
  "NOC",
  paste0(
    current_year, "-",
    second_five_years
  ),
  "ten_year_job_openings",
  paste0("JO_", current_year),
  paste0("JO_", first_five_years),
  paste0("JO_", second_five_years)
)],
by = "NOC", all.x = TRUE
)

# Adding the 10 year CAGR
LMO_WorkBC_Unemployment[, paste0(current_year, "-", second_five_years)] <-
  round((BC[, second_five_years] / BC[, current_year])^(1 / 10) - 1, 3) * 100
LMO_WorkBC_Unemployment <- LMO_WorkBC_Unemployment[complete.cases(LMO_WorkBC_Unemployment), ]
# Subsetting the final dataset to be exported
LMO_Career_Compass <- LMO_WorkBC_Unemployment[, c(
  "NOC",
  "Description",
  paste0("Unemployment_", current_year),
  paste0("Unemployment_", first_five_years),
  paste0("Unemployment_", second_five_years),
  paste0(current_year, "-", second_five_years),
  "ten_year_job_openings",
  paste0("JO_", current_year),
  paste0("JO_", first_five_years),
  paste0("JO_", second_five_years)
)]
LMO_Career_Compass <- na.omit(LMO_Career_Compass)

# 3.5 CAREER TREK-----------
# Adding zeroes and hashtages to the beginnning of the NOC column
Career_Trek$NOC <- ifelse(nchar(Career_Trek$NOC) == 3, paste0("#0", Career_Trek$NOC), Career_Trek$NOC)
Career_Trek$NOC <- ifelse(nchar(Career_Trek$NOC) == 2, paste0("#00", Career_Trek$NOC), Career_Trek$NOC)
Career_Trek$NOC <- ifelse(nchar(Career_Trek$NOC) == 4, paste0("#", Career_Trek$NOC), Career_Trek$NOC)
# Merge the career trek data with the bc data
Career_Trek <- merge(Career_Trek, BC, by = c("NOC"), all.x = TRUE)
# Add 10 year CAGR
Career_Trek[, paste0(current_year, "-", second_five_years)] <-
  round((Career_Trek[, second_five_years] / Career_Trek[, current_year])^(1 / 10) - 1, 3) * 100
# Subsetting data using column names
Career_Trek <- Career_Trek[, c("NOC", "Description", paste0(current_year, "-", second_five_years), "ten_year_job_openings")]
# Rounding the ten year job openings
Career_Trek$ten_year_job_openings <- round(Career_Trek$ten_year_job_openings, -2)

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