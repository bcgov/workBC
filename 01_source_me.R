# LIBRARIES--------------
library("data.table")
library("readxl")
library("XLConnect")
library("here")
library("lubridate")
library("dplyr")
# "Constants"----------
current_year = as.numeric(year(today())) -1 #delete the -1 once we have current data.
previous_year = as.character(current_year - 1)
first_five_years = as.character(current_year + 5) 
second_five_years = as.character(current_year + 10)
current_year = as.character(current_year)
region_list = list("Cariboo", 
                   "Kootenay", 
                   "Mainland South West", 
                   "North Coast & Nechako", 
                   "North East", 
                   "Thompson Okanagan", 
                   "Vancouver Island Coast", 
                   "British Columbia")

# FUNCTIONS----------------
source(here("R","functions.R"))
# LOAD DATA---------
Career_Trek <- read_xlsx(here("Data Sources", list.files(here("Data Sources"), pattern="newtemplate")))%>%
  mutate(NOC=as.character(NOC))
Empl_BC_Regions_Industry <- read_xlsx(here("Data Sources", list.files(here("Data Sources"), pattern="Empl_BC_Regions_Industry")))
Industry_Profiles <- read_xlsx(here("Data Sources", list.files(here("Data Sources"), pattern="Industry_Profiles")))
JobOpenings_BC_Regions_Industry <- read_xlsx(here("Data Sources", list.files(here("Data Sources"), pattern="JobOpenings_BC_Regions_Industry")))
LMO_WorkBC_Unemployment <- read_xlsx(here("Data Sources", list.files(here("Data Sources"), pattern="Unemployment_rate")))

#Cleaning-----------
names(Empl_BC_Regions_Industry) <-  as.matrix(Empl_BC_Regions_Industry[2,])
Empl_BC_Regions_Industry <-  Empl_BC_Regions_Industry[-c(1:2),]
names(JobOpenings_BC_Regions_Industry) <-  as.matrix(JobOpenings_BC_Regions_Industry[2,])
JobOpenings_BC_Regions_Industry <-  JobOpenings_BC_Regions_Industry[-c(1:2),]
names(LMO_WorkBC_Unemployment) <-  as.matrix(LMO_WorkBC_Unemployment[2,])
LMO_WorkBC_Unemployment <-  LMO_WorkBC_Unemployment[-c(1:2),]

# 3.3.1 CAREER PROFILE PROVINCIAL---------------
#' 1. Suppression rule:
#'   If the current year's employment is below 20, then all variables should be changed to "N/A"
#Subsetting BC and All Industries
BC <-  subset(Empl_BC_Regions_Industry, (`Geographic Area` %in% 'British Columbia') & Industry == 'All industries')
#Calculate CAGRs
BC <-  CAGR(BC)
#Job Openings---------------
attach(JobOpenings_BC_Regions_Industry)
#Subsetting out the All Industries data for BC
`Job_Openings` = subset(JobOpenings_BC_Regions_Industry, 
                        `Geographic Area` == 'British Columbia' & 
                          (Variable == 'Job Openings') & 
                          Industry == 'All industries')
#Round values to nearest 10
`Job_Openings`[, 7:16] = lapply(`Job_Openings`[, 7:16], function(x){round(x, -1)})
#Calculate the 10 year job openings
`Job_Openings`$ten_year_job_openings = apply(`Job_Openings`[, 7:16], 1,sum)
#Replacement of Retiring Workers % & number-------------
attach(JobOpenings_BC_Regions_Industry)
#Subset out BC's replacement and expansion data for all industries
BC_Expan_Replace = JobOpenings_BC_Regions_Industry[
  `Geographic Area` == 'British Columbia' & 
    !Variable %in% c("Deaths", "Retirements", "Job Openings") & 
    Industry == "All industries",]
BC_Expan_Replace$ten_year_sum = round(apply(BC_Expan_Replace[7:16], 1, sum), -1)
#Calculating the Replacement Percentage
BC_Expan_Replace = setDT(BC_Expan_Replace)[, `Replacement %` :=
            round((ten_year_sum[Variable == "Replacement Demand"]/
            (ten_year_sum[Variable == "Replacement Demand"] + 
            ten_year_sum[Variable == "Expansion Demand"]))*100, 1), 
            by = .(NOC)]
BC_Expan_Replace = setDT(BC_Expan_Replace)[, `Expansion %` :=
            round((ten_year_sum[Variable == "Expansion Demand"]/
            (ten_year_sum[Variable == "Replacement Demand"] + 
            ten_year_sum[Variable == "Expansion Demand"]))*100, 1), 
            by = .(NOC)]
#Merge the replacement and expansion data with the main BC dataset----------
BC = merge(BC, BC_Expan_Replace[BC_Expan_Replace$Variable == "Replacement Demand", c(1,17:19)], by = "NOC", all.x = TRUE)
colnames(BC)[21] = "ten_year_sum_Replace"
BC = merge(BC, BC_Expan_Replace[BC_Expan_Replace$Variable == "Expansion Demand", c(1,17)], by = "NOC", all.x = TRUE)
colnames(BC)[ncol(BC)] = "ten_year_sum_Expan"
#Changing the job openings column names
colnames(`Job_Openings`)[c(6, 11, 16)] = c( paste0("JO_", current_year), 
                                                 paste0("JO_", first_five_years), 
                                                 paste0("JO_", second_five_years))
#Changing to column names
BC = merge(BC, `Job_Openings`[ , c(1, 6, 11, 16, 17)], 
           by = "NOC", 
           all.x = TRUE)
#Reordering the columns
Provincial_Career_Profiles = BC[, c("NOC", 
                                   "Description", 
                                   paste0(current_year, "-", first_five_years),
                                   paste0(first_five_years, "-", second_five_years),
                                   paste0("JO_", current_year),
                                   paste0("JO_", first_five_years),
                                   paste0("JO_", second_five_years),
                                   "ten_year_job_openings", 
                                   "Replacement %", 
                                   "ten_year_sum_Replace", 
                                   "Expansion %", 
                                   "ten_year_sum_Expan",
                                   current_year
                                   )]
#Dealing NAN in replacement and expansion percentage columns
#If the cell contains an NA then replace with zero
Provincial_Career_Profiles$`Expansion %` = ifelse(is.na(Provincial_Career_Profiles$`Expansion %`), 0, Provincial_Career_Profiles$`Expansion %`)
Provincial_Career_Profiles$`Replacement %` = ifelse(is.na(Provincial_Career_Profiles$`Replacement %`), 0, Provincial_Career_Profiles$`Replacement %`)
#Dealing with the replacement and expansion percentages/values not adding up to 100
# If the replacement % is above 100 then add the replacement % to the expansion % to make it 100
Provincial_Career_Profiles$`Replacement %` = ifelse(Provincial_Career_Profiles$`Replacement %` > 100, Provincial_Career_Profiles$`Replacement %` + Provincial_Career_Profiles$`Expansion %`, Provincial_Career_Profiles$`Replacement %`)
#If expansion is less than zero then convert it to be zero
Provincial_Career_Profiles$`Expansion %` = ifelse(Provincial_Career_Profiles$`Expansion %` < 0, 0, Provincial_Career_Profiles$`Expansion %`)
#Correcting the replacement sum when the expansion sum is negative
Provincial_Career_Profiles$ten_year_sum_Replace = ifelse(Provincial_Career_Profiles$ten_year_sum_Expan < 0 &
                                                           Provincial_Career_Profiles$`Replacement %` > 0 , Provincial_Career_Profiles$ten_year_sum_Replace + Provincial_Career_Profiles$ten_year_sum_Expan, Provincial_Career_Profiles$ten_year_sum_Replace)
#Now that the replacement sum has been corrected, the negative expansion sums can be corrected to zero
Provincial_Career_Profiles$ten_year_sum_Expan = ifelse(Provincial_Career_Profiles$ten_year_sum_Expan < 0, 0, Provincial_Career_Profiles$ten_year_sum_Expan)
#If the expansion % is greater then 100 then change it to be 100
Provincial_Career_Profiles$`Expansion %` = ifelse(Provincial_Career_Profiles$`Expansion %` > 100, 0, Provincial_Career_Profiles$`Expansion %`)
#If the replacement % is less then 0 then change it be zero
Provincial_Career_Profiles$`Replacement %` = ifelse(Provincial_Career_Profiles$`Replacement %` < 0, 100, Provincial_Career_Profiles$`Replacement %`)
#Replace rows with current year employment less than 20 with NA
Provincial_Career_Profiles[Provincial_Career_Profiles[, current_year] < 20, 3:12] <- "N/A"
#Subsetting out the final columns needed
Provincial_Career_Profiles = Provincial_Career_Profiles[,c("NOC", 
                                                           "Description", 
                                                           paste0(current_year, "-", first_five_years),
                                                           paste0(first_five_years, "-", second_five_years),
                                                           paste0("JO_", current_year),
                                                           paste0("JO_", first_five_years),
                                                           paste0("JO_", second_five_years),
                                                           "ten_year_job_openings", 
                                                           "Replacement %", 
                                                           "ten_year_sum_Replace", 
                                                           "Expansion %", 
                                                           "ten_year_sum_Expan")]


# 3.3.1 CAREER PROFILE REGIONAL---------- 
#' 1. Suppression rule:
#'   If the current year's employment is below 20, then all variables should be changed to "N/A"
#Initiating a list to contain all of the regional datasets
regional_data_list = list()
count = 1
for (data in region_list){
  #Subsetting BC and All Industries
  attach(Empl_BC_Regions_Industry)
  assign(data, Empl_BC_Regions_Industry[`Geographic Area` == data & Industry == 'All industries',])
}
#Second loop to perform the calculations-----------
for (data in region_list){
  #Calculating the average annual employment growth rate between different time periods
  #Compound Average Annual Growth Rate
  temp = get(data)
  paste0("avg_employ_growth_", current_year, "_", second_five_years)
  temp[, paste0("avg_employ_growth_", current_year, "_", second_five_years)] =
    round(((temp[, second_five_years] /
              temp[, current_year]) ^ (1 / 10)) - 1, digits = 3) * 100
  colnames(temp)[length(temp)] = paste0(data,
                                        "_",
                                        "avg_employ_growth_",
                                        current_year,
                                        "_",
                                        second_five_years)
  #10 Year job openings
  attach(JobOpenings_BC_Regions_Industry)
  temp2 = JobOpenings_BC_Regions_Industry[`Geographic Area` == data &
                                                   Variable == 'Job Openings' &
                                                   Industry == 'All industries',]
  
  #Finding the 10 year sum of the job openings - sum of 2022 to 2031
  temp2$ten_year_job_openings = round(apply(temp2[, 7:16], 1,sum), -1)
  #Adding a column name
  colnames(temp2)[length(temp2)] = paste0(data, "_ten_year_job_openings")
  #Subsetting the data to keep just the NOC and ten year job openings
  temp2 = temp2[,c("NOC", paste0(data, "_ten_year_job_openings"))]
  #merge the temp2 data to temp
  temp = merge(temp, temp2, by = "NOC", all.x = TRUE)
  #Employment in 2021
  attach(Empl_BC_Regions_Industry)
  employ = Empl_BC_Regions_Industry[`Geographic Area` == data & Industry == 'All industries',]
  #Subsetting out the data for future merging
  employ = employ[, colnames(employ) %in% c("NOC", current_year)]
  #Suppressing values less than 200
  employ[employ[, current_year] < 20, current_year] <- NA
  #Changing the column name (keep as index because the year will change)
  colnames(employ)[2] = paste0(data, "_Employment_", current_year)
  #Round the employment data for the current year to the nearest 10th
  employ[, paste0(data, "_Employment_", current_year)] = round(employ[, paste0(data, "_Employment_", current_year)], -1)
  #Merging the job openings and employment data
  temp = merge(temp, employ, by = "NOC", all.x = TRUE)
  #Subsetting the final temp data and then placing it all into a list to be merged together
  #Subsetting out the data for calculations
  temp = temp[,c("NOC", 
                 "Description", 
                 paste0(data, "_Employment_", current_year),
                 paste0(data, "_", "avg_employ_growth_", current_year, "_", second_five_years),
                 paste0(data, "_ten_year_job_openings")
                 )]
  #If an NA exists in the ten year employment then make sure all rows contain NA as well
  temp[apply(is.na(temp[, c(paste0(data, "_Employment_", current_year),
                          paste0(data, "_", "avg_employ_growth_", current_year, "_", second_five_years),
                          paste0(data, "_ten_year_job_openings")
                          )
                        ]), 1, any), 
                          c(paste0(data, "_Employment_", current_year),
                          paste0(data, "_", "avg_employ_growth_", current_year, "_", second_five_years),
                          paste0(data, "_ten_year_job_openings"))] <- "N/A"
  
  if (count == 1){
    maindf = temp
  }else{
    maindf = merge(maindf, temp, by = c("NOC", "Description"))
  }
  count = count + 1
}
#end of Second loop to perform the calculations-----------
Career_Profile_Regional = maindf
#Removing ten_year_employ columns
Career_Profile_Regional = Career_Profile_Regional[ , !grepl("ten_year_employ", colnames(Career_Profile_Regional))]

# 3.3.2 INDUSTRY PROFILE-------------
geo <-  "British Columbia"
#Subset out the total employment for each industry
temp = subset(Empl_BC_Regions_Industry, `Geographic Area` == geo & Description == "Total")
#Using the standardized industry names, merge the employment data
Industry_Profile = merge(Industry_Profiles, temp, by = "Industry", all.x = TRUE)
#Subsetting out the employment values for the 10 year forecast
Industry_Profile = Industry_Profile[, c(1, (ncol(Industry_Profile) - 10):(ncol(Industry_Profile)))]
#Assigning Industry as the rowname
rownames(Industry_Profile) = Industry_Profile[,1]
Industry_Profile = Industry_Profile[,-1]
#Find the aggregate industry values using custom function
Industry_Profile = aggregate_industries(Industry_Profile)
#Calculating the industry share of total employment
#Current year share
Industry_Profile[, paste0(current_year, " share of employment")] = round(prop.table(Industry_Profile[, current_year])*100, 1)
#Fifth year share
Industry_Profile[, paste0(first_five_years, " share of employment")] = round(prop.table(Industry_Profile[, first_five_years])*100, 1)
#Tenth year share
Industry_Profile[, paste0(second_five_years, " share of employment")] = round(prop.table(Industry_Profile[, second_five_years])*100, 1)

#Forecasted Employment
#Current year
Industry_Profile[, paste0(current_year, " forecasted employment")] = round(Industry_Profile[, current_year], -2)
#Fifth year
Industry_Profile[, paste0(first_five_years, " forecasted employment")] = round(Industry_Profile[, first_five_years], -2)
#Tenth year
Industry_Profile[, paste0(second_five_years, " forecasted employment")] = round(Industry_Profile[, second_five_years], -2)
#Calculate CAGRs
Industry_Profile =  CAGR(Industry_Profile)
#Creating a column for the industry names
Industry_Profile[,1] = row.names(Industry_Profile)
colnames(Industry_Profile)[1] = "Industry"
#Extracting the final data
Industry_Profile = Industry_Profile[, c(1, 12:ncol(Industry_Profile))]
#Naming column
colnames(Industry_Profile)[1] = "Industry"
#Removing rownames
rownames(Industry_Profile) = NULL
#Ordering the data to the LMO industry order
Industry_Profile = with(Industry_Profile, Industry_Profile[order(Industry, decreasing = FALSE),])
# 3.3.3 REGIONAL PROFILE - REGIONAL PROFILES LMO-------------
#COMPOSITION OF JOB OPENINGS SECTION - REGIONAL PROFILES LMO
#Copy the data that I've already produced
composition_job_openings = subset(JobOpenings_BC_Regions_Industry, 
                           Variable %in% c("Expansion Demand", "Replacement Demand") & 
                           Industry == "All industries" & 
                           Description == "Total" & 
                           `Geographic Area` %in% c("Cariboo", 
                                                    "Kootenay", 
                                                    "Mainland South West", 
                                                    "North Coast & Nechako", 
                                                    "North East", 
                                                    "Thompson Okanagan", 
                                                    "Vancouver Island Coast", 
                                                    "British Columbia"))
#Find 10 year sum of job openings
composition_job_openings$ten_year_sum = apply(composition_job_openings[ , 7:16], 1, sum)
#Finding the replacement and expansion percentages
composition_job_openings =  setDT(composition_job_openings)[, `Replacement %` := 
                                                      round((ten_year_sum[Variable == "Replacement Demand"]/
                                                      (ten_year_sum[Variable == "Replacement Demand"] + 
                                                      ten_year_sum[Variable == "Expansion Demand"]))*100, 1), 
                                                      by = .(`Geographic Area`)]
composition_job_openings =  setDT(composition_job_openings)[, `Expansion %` :=
                                                      round((ten_year_sum[Variable == "Expansion Demand"]/
                                                      (ten_year_sum[Variable == "Replacement Demand"] + 
                                                      ten_year_sum[Variable == "Expansion Demand"]))*100, 1), 
                                                      by = .(`Geographic Area`)]
composition_job_openings = composition_job_openings[ , c("Geographic Area", 
                                                         "Variable", 
                                                         "ten_year_sum", 
                                                         "Replacement %", 
                                                         "Expansion %")]
#Widen dataset
composition_job_openings = dcast(composition_job_openings, `Geographic Area` ~ Variable, 
                                                           value.var = c("ten_year_sum", 
                                                                         "Replacement %", 
                                                                         "Expansion %"))
#Selecting out the final data
composition_job_openings = composition_job_openings[ , c("Geographic Area",
                                                         "ten_year_sum_Replacement Demand", 
                                                         "Replacement %_Expansion Demand", 
                                                         "ten_year_sum_Expansion Demand", 
                                                         "Expansion %_Expansion Demand")]
#Rounding the values
composition_job_openings[ , c(2, 4)] = lapply(composition_job_openings[ , c(2, 4)], function(x) {round(x, -2)})
#EMPLOYMENT OUTLOOK SECTION - REGIONAL PROFILES LMO
employment_outlook = subset(Empl_BC_Regions_Industry, 
                           Variable %in% c("Employment") & 
                           Industry == "All industries" & 
                           Description == "Total" & 
                           `Geographic Area` %in% c("Cariboo", 
                                                    "Kootenay", 
                                                    "Mainland South West", 
                                                    "North Coast & Nechako", 
                                                    "North East", 
                                                    "Thompson Okanagan", 
                                                    "Vancouver Island Coast", 
                                                    "British Columbia"))
#Find 10 year sums by merging Job Openings 10 year sum to employment
jobopenings = subset(JobOpenings_BC_Regions_Industry, 
                           Variable %in% c("Job Openings") & 
                           Industry == "All industries" & 
                           Description == "Total" & 
                           `Geographic Area` %in% c("Cariboo", 
                                                    "Kootenay", 
                                                    "Mainland South West", 
                                                    "North Coast & Nechako", 
                                                    "North East", 
                                                    "Thompson Okanagan", 
                                                    "Vancouver Island Coast", 
                                                    "British Columbia"))

jobopenings = as.data.frame(cbind(jobopenings[, "Geographic Area"], round(apply(jobopenings[ , 7:16], 1, sum), -2)))
colnames(jobopenings) = c("Geographic Area", "ten_year_sum")
#Merging employment and job openings
employment_outlook = merge(employment_outlook, jobopenings, by = "Geographic Area")
#Calculate CAGRs
employment_outlook = CAGR(employment_outlook)
#Selecting out the final data needed
employment_outlook = employment_outlook[ , c(current_year, 
                                             first_five_years, 
                                             second_five_years, 
                                             "ten_year_sum", 
                                             paste0(current_year, "-", second_five_years))]
#Convert to numeric
employment_outlook$ten_year_sum = as.numeric(as.character(employment_outlook$ten_year_sum))
#Rounding the values
employment_outlook[ , c(1:4)] = lapply(employment_outlook[ , c(1:4)], function(x) {round(as.numeric(x), -2)})
# 3.3.3 REGIONAL PROFILE - TOP 10 INDUSTRIES---------------
#Looping over each region to find their aggregated industries-----------
count = 0
for(geo in region_list){
  #Find the job openings data for each region
  attach(JobOpenings_BC_Regions_Industry)
  temp_regional_profile = as.data.frame(JobOpenings_BC_Regions_Industry[Variable %in% c("Job Openings") & 
                                                                   Description == "Total" & 
                                                                   `Geographic Area` %in% geo, ])
  
  #Subset the data even further
  temp_regional_profile = droplevels(temp_regional_profile)
  temp_regional_profile = temp_regional_profile[, c(3,7:16)]
  #Assigning Industry as the rowname
  rownames(temp_regional_profile) = temp_regional_profile[,1]
  temp_regional_profile = temp_regional_profile[,-1]
  #Calculate the aggregate industries
  temp_regional_profile = aggregate_industries(temp_regional_profile)
  #Calculating ten year job openings
  temp_regional_profile$ten_year_job_openings = round(apply(temp_regional_profile[, 1:10], 1, sum), -2)
  #GETTING THE DATA INTO THE RIGHT FORMAT AND JOINING THE CAGR DATA FROM THE INDUSTRY PROFILES
  temp_regional_profile[,1] = row.names(temp_regional_profile)
  colnames(temp_regional_profile)[1] = "Industry"
  #Extracting the final data
  temp_regional_profile = temp_regional_profile[, c(1, 11:ncol(temp_regional_profile))]
  colnames(temp_regional_profile)[1] = "Industry"
  rownames(temp_regional_profile) = NULL
  #Ordering the temp_regional data by Industry Name
  temp_regional_profile = with(temp_regional_profile, temp_regional_profile[order(Industry, decreasing = FALSE),])
  temp_regional_profile[, paste0("avg_employ_growth_", current_year, "_", second_five_years)] = 
    Industry_Profile[, paste0(current_year, "-", second_five_years)]
  temp_regional_profile$Region = geo
  
  #save the data
  if(count == 0){
    Regional_Profile_maindf = temp_regional_profile
  }else{
    Regional_Profile_maindf = rbind(Regional_Profile_maindf, temp_regional_profile)
  } 
  
  count = count + 1

}
#end of looping over regions-------------  
#Finding the top ten industries for the regional profiles
Regional_Profile_maindf2 = setDT(Regional_Profile_maindf)[order(ten_year_job_openings, decreasing = TRUE), head(.SD, 10), by = Region]
#Sort by region
Regional_Profile_maindf2 = Regional_Profile_maindf2[order(Regional_Profile_maindf2$Region), ]

# 3.3.3 REGIONAL PROFILE - TOP 10 OCCUPATIONS-----------
#' 1. The top 10 occupations relies upon the top 10 industries to be run first

for (data in region_list) {
  #Subsetting BC and All Industries
  attach(JobOpenings_BC_Regions_Industry)
  assign(
    data,
    subset(
      JobOpenings_BC_Regions_Industry,
      `Geographic Area` == data &
      Industry == 'All industries' &
      Variable == "Job Openings"
    )
  )
  
}
top_ten_occupations = data.frame()
top_ten_industries = data.frame()
Occupations_top_10 = data.frame()
industry_top_10 = data.frame()

#Second loop to perform the calculations
for (data in region_list) {
  temp = get(data)
  #Subsetting out the Job Openings data
  temp2 = subset(
    JobOpenings_BC_Regions_Industry,
    `Geographic Area` == data &
      Industry == 'All industries' &
      Variable == "Job Openings"
  )
  #Finding the sum of the job openings over each year
  temp2$ten_year_job_openings = apply(temp2[, 7:16], 1, sum)
    #Subset All Industries from the employment data
  temp4 = subset(Empl_BC_Regions_Industry,
                 `Geographic Area` == data &
                   Industry == "All industries")
  #Calculate the CAGR
  temp_regional_profile[, paste0("avg_employ_growth_", current_year, "_", second_five_years)] =
    Industry_Profile[, paste0(current_year, "-", second_five_years)]
  #Calculating the 10 year CAGR
  temp4[, paste("avg_employ_growth", current_year, second_five_years, sep="_")] =
    round((temp4[, second_five_years] / temp4[, current_year]) ^ (1 / 10) -
            1, digits = 3) * 100
  #Convert any NANs to zeros
  #temp4[, paste0("avg_employ_growth_2021_2031")] =
   # ifelse(is.na(temp4[, paste0("avg_employ_growth_2021_2031")]),
    #       "NA",
     #      temp4[, paste0("avg_employ_growth_2021_2031")])
  
  #Merge CAGR to job openings (temp2)
  temp2 = merge(temp2, temp4[, colnames(temp4) %in% c("NOC",
                                                      paste0("avg_employ_growth_", current_year, "_", second_five_years))])
  #Subsetting the data to keep just the NOC and ten year job openings
  temp2 = temp2[, c(1, length(temp2) - 1, length(temp2))]
  #merge the temp2 data to temp
  temp_occ = merge(temp, temp2, by = "NOC", all.x = TRUE)
  #Finding the top 10 occupations
  temp_occ = head(temp_occ[order(temp_occ$ten_year_job_openings, decreasing = TRUE), ], n = 11)
    #Combine all of the top 10 occupations from each region together
  Occupations_top_10 = rbind(Occupations_top_10, temp_occ, fill = TRUE)
}
###checking 
#Remove duplicates
Occupations_top_10 = Occupations_top_10[!duplicated(Occupations_top_10), ]
Occupations_top_10 = na.omit(Occupations_top_10)

attach(Occupations_top_10)
Occupations_top_ten = Occupations_top_10[Description != "Total",
                                         colnames(Occupations_top_10) %in% c(
                                           "Geographic Area",
                                           "NOC",
                                           "Description",
                                           "ten_year_job_openings",
                                           paste0("avg_employ_growth_", current_year, "_", second_five_years)
                                         )]
Occupations_top_ten$ten_year_job_openings = round(Occupations_top_ten$ten_year_job_openings,-1)
#Reordering the columns
Occupations_top_ten = Occupations_top_ten[, c(
  "Geographic Area",
  "NOC",
  "Description",
  "ten_year_job_openings",
  paste0("avg_employ_growth_", current_year, "_", second_five_years)
)]

# 3.4 CAREER COMPASS - BROWSE CAREERS---------
#Calculate the unemployment rate for the current year
LMO_WorkBC_Unemployment[, paste0("Unemployment_", current_year)] = 
  round(LMO_WorkBC_Unemployment[, paste0(current_year)], 1)
#Fifth year unemployment rate
LMO_WorkBC_Unemployment[, paste0("Unemployment_", first_five_years)] = 
  round(LMO_WorkBC_Unemployment[, paste0(first_five_years)], 1)
#Tenth year unemployment rate
LMO_WorkBC_Unemployment[, paste0("Unemployment_", second_five_years)] = 
  round(LMO_WorkBC_Unemployment[, paste0(second_five_years)], 1)
#Adding the CAGR and job openings data to the unemployment rate data
LMO_WorkBC_Unemployment = merge(LMO_WorkBC_Unemployment, BC[, colnames(BC) %in% c("NOC", 
                                                                                          paste0(current_year, "-",
                                                                                                 second_five_years),
                                                                                          "ten_year_job_openings", 
                                                                                          paste0("JO_", current_year), 
                                                                                          paste0("JO_", first_five_years), 
                                                                                          paste0("JO_", second_five_years))], 
                                                                    by = "NOC", all.x = TRUE)

#Adding the 10 year CAGR
LMO_WorkBC_Unemployment[, paste0(current_year, "-", second_five_years)] = 
  round((BC[, second_five_years]/BC[, current_year])^(1/10)-1, 3)*100
LMO_WorkBC_Unemployment = LMO_WorkBC_Unemployment[complete.cases(LMO_WorkBC_Unemployment), ]
#Subsetting the final dataset to be exported
LMO_Career_Compass = LMO_WorkBC_Unemployment[, c("NOC", 
                                                          "Description",
                                                          paste0("Unemployment_", current_year),
                                                          paste0("Unemployment_", first_five_years),
                                                          paste0("Unemployment_", second_five_years),
                                                          paste0(current_year, "-", second_five_years), 
                                                          "ten_year_job_openings",
                                                          paste0("JO_", current_year), 
                                                          paste0("JO_", first_five_years), 
                                                          paste0("JO_", second_five_years))]
LMO_Career_Compass = na.omit(LMO_Career_Compass)

# 3.5 CAREER TREK-----------
#Adding zeroes and hashtages to the beginnning of the NOC column
Career_Trek$NOC = ifelse(nchar(Career_Trek$NOC) == 3, paste0("#0", Career_Trek$NOC), Career_Trek$NOC)
Career_Trek$NOC = ifelse(nchar(Career_Trek$NOC) == 2, paste0("#00", Career_Trek$NOC), Career_Trek$NOC)
Career_Trek$NOC = ifelse(nchar(Career_Trek$NOC) == 4, paste0("#", Career_Trek$NOC), Career_Trek$NOC)
#Merge the career trek data with the bc data
Career_Trek = merge(Career_Trek, BC, by = c("NOC"), all.x = TRUE)
#Add 10 year CAGR
Career_Trek[, paste0(current_year, "-", second_five_years)] = 
  round((Career_Trek[, second_five_years]/Career_Trek[, current_year])^(1/10)-1, 3)*100
#Subsetting data using column names
Career_Trek = Career_Trek[, c("NOC", "Description", paste0(current_year, "-", second_five_years), "ten_year_job_openings")]
#Rounding the ten year job openings
Career_Trek$ten_year_job_openings = round(Career_Trek$ten_year_job_openings, -2)

# EXPORTING DATA DIRECTLY TO EXCEL----------
#3.3.1
wb <- loadWorkbook(here("templates", "3.3.1_WorkBC_Career_Profile_Data.xlsx"))
#Provincial Outlook Data
write_workbook(Provincial_Career_Profiles[, c(3:12)], "Provincial Outlook", 4, 3)
#Regional Outlook Data - Cariboo
write_workbook(Career_Profile_Regional[, c(3:5)], "Regional Outlook", 4, 3)
#Regional Outlook Data - Kootney
write_workbook(Career_Profile_Regional[, c(6:8)], "Regional Outlook", 4, 7)
#Regional Outlook Data - Mainland Southwest
write_workbook(Career_Profile_Regional[, c(9:11)], "Regional Outlook", 4, 11)
#Regional Outlook Data - North Coast and Nechako
write_workbook(Career_Profile_Regional[, c(12:14)], "Regional Outlook", 4, 15)
#Regional Outlook Data - Northeast
write_workbook(Career_Profile_Regional[, c(15:17)], "Regional Outlook", 4, 19)
#Regional Outlook Data - Thompson Okanagan
write_workbook(Career_Profile_Regional[, c(18:20)], "Regional Outlook", 4, 23)
#Regional Outlook Data - Vancouver Island Coast
write_workbook(Career_Profile_Regional[, c(21:23)], "Regional Outlook", 4, 27)
saveWorkbook(wb, 
  here("Send to WorkBC", 
  paste0("3.3.1_WorkBC_Career_Profile_Data", 
         current_year,
         "-", 
         second_five_years, 
         ".xlsx")))
#export 3.3.2-------------
wb <- loadWorkbook(here("templates", "3.3.2_WorkBC_Industry_Profile.xlsx"))
#BC Industry Data
write_workbook(Industry_Profile[, c(2:9)], "BC", 4, 2)
saveWorkbook(wb, 
             here("Send to WorkBC", 
                  paste0("3.3.2_WorkBC_Industry_Profile", 
                         current_year,
                         "-", 
                         second_five_years, 
                         ".xlsx")))

#export 3.3.3-----------
wb <- loadWorkbook(here("templates", "3.3.3_WorkBC_Regional_Profile_Data.xlsx"))
write_workbook(composition_job_openings[, c(2:5)], "Regional Profiles - LMO", 5, 2)
#Regional Data
write_workbook(employment_outlook, "Regional Profiles - LMO", 5, 6)
#Top Occupations - BC
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "British Columbia", 2:5], "Top Occupation", 5, 1)
#Top Occupation - Cariboo
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Cariboo", 2:5], "Top Occupation", 17, 1)
#Top Occupation - Kootenay
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Kootenay", 2:5], "Top Occupation", 29, 1)
#Top Occupation - Mainland SouthWest
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Mainland South West", 2:5], "Top Occupation", 41, 1)
#Top Occupation - North Coast & Nechako
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North Coast & Nechako", 2:5], "Top Occupation", 53, 1)
#Top Occupation - North East
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North East", 2:5], "Top Occupation", 65, 1)
#Top Occupation - Thompson Okanagan
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Thompson Okanagan", 2:5], "Top Occupation", 77, 1)
#Top Occupation - Vancouver Island Coast
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Vancouver Island Coast", 2:5], "Top Occupation", 89, 1)
#Top Industries - BC
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "British Columbia", 2:4], "Top Industries", 5, 1)
#Top Industries - Cariboo
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Cariboo", 2:4], "Top Industries", 17, 1)
#Top Industries - Kootenay
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Kootenay", 2:4], "Top Industries", 29, 1)
#Top Industries - Mainland SouthWest
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Mainland South West", 2:4], "Top Industries", 42, 1)
#Top Industries - North Coast & Nechako
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North Coast & Nechako", 2:4], "Top Industries", 54, 1)
#Top Industries - North East
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North East", 2:4], "Top Industries", 67, 1)
#Top Industries - Thompson Okanagan
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Thompson Okanagan", 2:4], "Top Industries", 79, 1)
#Top Industries - Vancouver Island Coast
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Vancouver Island Coast", 2:4], "Top Industries", 92, 1)
saveWorkbook(wb, 
             here("Send to WorkBC", 
                  paste0("3.3.3_WorkBC_Regional_Profile_Data", 
                         current_year,
                         "-", 
                         second_five_years, 
                         ".xlsx")))
#export 3.4---------------
wb <- loadWorkbook(here("templates", "3.4_WorkBC_Career_Compass.xlsx"))
#Browse Careers
write_workbook(LMO_Career_Compass[, c(3:10)], "Browse Careers", 4, 3)
#Top Occupation - Cariboo
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Cariboo", 2:4], "Regions", 5, 1)
#Top Occupation - Kootenay
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Kootenay", 2:4], "Regions", 17, 1)
#Top Occupation - Mainland SouthWest
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Mainland South West", 2:4], "Regions", 29, 1)
#Top Occupation - North Coast & Nechako
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North Coast & Nechako", 2:4], "Regions", 41, 1)
#Top Occupation - North East
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "North East", 2:4], "Regions", 53, 1)
#Top Occupation - Thompson Okanagan
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Thompson Okanagan", 2:4], "Regions", 65, 1)
#Top Occupation - Vancouver Island Coast
write_workbook(Occupations_top_ten[Occupations_top_ten$`Geographic Area` == "Vancouver Island Coast", 2:4], "Regions", 77, 1)
#Top Industries - Cariboo
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Cariboo", 2:3], "Regions", 5, 5)
#Top Industries - Kootenay
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Kootenay", 2:3], "Regions", 17, 5)
#Top Industries - Mainland SouthWest
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Mainland South West", 2:3], "Regions", 29, 5)
#Top Industries - North Coast & Nechako
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North Coast & Nechako", 2:3], "Regions", 41, 5)
#Top Industries - North East
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "North East", 2:3], "Regions", 53, 5)
#Top Industries - Thompson Okanagan
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Thompson Okanagan", 2:3], "Regions", 65, 5)
#Top Industries - Vancouver Island Coast
write_workbook(Regional_Profile_maindf2[Regional_Profile_maindf2$Region == "Vancouver Island Coast", 2:3], "Regions", 77, 5)
saveWorkbook(wb, 
             here("Send to WorkBC", 
                  paste0("3.4_WorkBC_Career_Compass", 
                         current_year,
                         "-", 
                         second_five_years, 
                         ".xlsx")))
#export 3.5-------------------
wb <- loadWorkbook(here("templates", "3.5_WorkBC_Career_Trek.xlsx"))
#Browse Careers
write_workbook(Career_Trek[, c(3:4)], "LMO", 2, 5)
saveWorkbook(wb, 
             here("Send to WorkBC", 
                  paste0("3.5_WorkBC_Career_Trek", 
                         current_year,
                         "-", 
                         second_five_years, 
                         ".xlsx")))

#export 3.7--------------
wb <- loadWorkbook(here("templates", "3.7_WorkBC_Buildprint_Builder.xlsx"))
#Regional Data
write_workbook(composition_job_openings[, c(2:5)], "Regional Data", 5, 2)
#Jobs in Demand
write_workbook(Occupations_top_ten[order(Occupations_top_ten$`Geographic Area`), c(3, 2)], "Jobs in Demand", 4, 2)
saveWorkbook(wb, 
             here("Send to WorkBC", 
                  paste0("3.7_WorkBC_Buildprint_Builder", 
                         current_year,
                         "-", 
                         second_five_years, 
                         ".xlsx")))


