CAGR = function(data){
  
  data[, paste0(current_year, "-", first_five_years)] = 
    round((data[, first_five_years]/data[, current_year])^(1/5)-1, 3)*100
  
  data[, paste0(first_five_years, "-", second_five_years)] = 
    round((data[, second_five_years]/data[, first_five_years])^(1/5)-1, 3)*100
  
  data[, paste0(current_year, "-", second_five_years)] = 
    round((data[, second_five_years]/data[, current_year])^(1/10)-1, 3)*100
  
  return(data)
}

#Function to quickly export data
write_workbook = function(data, sheetname, startrow, startcol){
  writeWorksheet(
    wb,
    data,
    sheetname,
    startRow = startrow,
    startCol = startcol,
    header = FALSE
  )
}

#61 industries. Make sure that all names are the same names in the in 4CastViewer data sets, industry profiles excel sheet and this code (look for extra spaces). 
#Calculate the values for the LMO aggregate industries
aggregate_industries = function(data){
  
  data['Agriculture and fishing',] = colSums(data[row.names(data) %in% c("Farms and support activities", 
                                                                         "Fishing, hunting and trapping"), 
                                                  1:ncol(data)])
  
  data['Forestry and logging with support activities',] = 
    colSums(data[row.names(data) %in% c("Forestry, logging and support activities"), 
                 1:ncol(data)])
  
  data['Mining and oil and gas extraction',] = 
    colSums(data[row.names(data) %in% c("Oil and gas extraction",
                                        "Mining",
                                        "Support activities for mining and oil and gas extraction" 
    ), 
    1:ncol(data)])
  
  data['Utilities',] = 
    colSums(data[row.names(data) %in% c("Utilities" 
    ), 
    1:ncol(data)])
  
  data['Construction',] = 
    colSums(data[row.names(data) %in% c("Construction" 
    ), 
    1:ncol(data)])
  
  data['Manufacturing',] = colSums(data[row.names(data) %in% c("Food, beverage and tobacco manufacturing",
                                                               "Wood product manufacturing",
                                                               "Paper manufacturing",
                                                               "Primary metal manufacturing",
                                                               "Fabricated metal product manufacturing",
                                                               "Machinery manufacturing",
                                                               "Ship and boat building",
                                                               "Transportation equipment manufacturing (excluding shipbuilding)",
                                                               "Other manufacturing"
  ), 
  1:ncol(data)])
  
  data['Retail trade',] = colSums(data[row.names(data) %in% c("Motor vehicle and parts dealers",
                                                              "Health and personal care stores",
                                                              "Online shopping",
                                                              "Other retail trade (excluding cars, online shopping and personal care)"
  ),
  1:ncol(data)])
  
  data['Transportation and warehousing',] = colSums(data[row.names(data) %in% c("Air transportation and support activities",
                                                                                "Rail transportation and support activities",
                                                                                "Water transportation",
                                                                                "Ports and freight transportation arrangement",
                                                                                "Truck transportation and support activities",
                                                                                "Transit, sightseeing and pipeline transportation",
                                                                                "Postal service, couriers and messengers",
                                                                                "Warehousing and storage"
  ), 
  1:ncol(data)])
  
  data['Finance, insurance and real estate',] = colSums(data[row.names(data) %in% c("Finance",
                                                                                    "Insurance carriers and related activities",
                                                                                    "Real estate and rental and leasing"
  ),
  1:ncol(data)])
  
  data['Professional, scientific and technical services',] = 
    colSums(data[row.names(data) %in% c("Architectural, engineering and related services", 
                                        "Computer systems design and related services",
                                        "Management, scientific and technical consulting services",
                                        "Legal, accounting, design, research and advertising services"
    ), 
    1:ncol(data)])
  
  data['Business, building and other support services',] = colSums(data[row.names(data) %in% c("Travel arrangement services",
                                                                                               "Business and building support services (excluding travel)"
  ),
  1:ncol(data)])
  
  
  data['Educational services',] = colSums(data[row.names(data) %in% c("Elementary and secondary schools", 
                                                                      "Community colleges", 
                                                                      "Universities", 
                                                                      "Private and trades education"), 1:ncol(data)])
  
  data['Health care and social assistance',] = colSums(data[row.names(data) %in% c("Ambulatory health care services",
                                                                                   "Hospitals",
                                                                                   "Nursing and residential care facilities",
                                                                                   "Social assistance"
  ), 
  1:ncol(data)])
  
  
  data['Information, culture and recreation',] = 
    colSums(data[row.names(data) %in% c("Publishing industries", 
                                        "Motion picture and sound recording industries", 
                                        "Telecommunications", 
                                        "Broadcasting, data processing and information", 
                                        "Performing arts, spectator sports and related industries", 
                                        "Amusement, gambling and recreation industries", 
                                        "Museums, galleries and parks"
    ), 
    1:ncol(data)])
  
  data['Accommodation and food services',] = colSums(data[row.names(data) %in% c("Accommodation services", 
                                                                                 "Food services and drinking places"), 
                                                          1:ncol(data)])
  
  data['Repair, personal and non-profit services',] = colSums(data[row.names(data) %in% c("Automotive repair and maintenance",
                                                                                          "Personal, non-automotive repair and non-profit services"),
                                                                   1:ncol(data)])
  
  data['Public Administration',] = colSums(data[row.names(data) %in% c("Federal government public administration",
                                                                       "Provincial and territorial public administration",
                                                                       "Local and Indigenous public administration"
  ),
  1:ncol(data)])
  
  
  data = na.omit(data)
  
  #Subsetting out the aggregate industries
  data = data[row.names(data) %in% c('Agriculture and fishing',
                                     'Forestry and logging with support activities',
                                     'Mining and oil and gas extraction',
                                     'Utilities',
                                     'Construction',
                                     'Manufacturing',
                                     "Wholesale trade",
                                     'Retail trade',
                                     'Transportation and warehousing',
                                     'Finance, insurance and real estate',
                                     'Professional, scientific and technical services',
                                     "Business, building and other support services",
                                     'Educational services',
                                     'Health care and social assistance',
                                     'Information, culture and recreation',
                                     'Accommodation and food services',
                                     'Repair, personal and non-profit services',
                                     'Public Administration'),]
  
  return(data)
}