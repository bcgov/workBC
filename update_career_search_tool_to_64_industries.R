cstjo_old <- read_excel(here("raw_data","archive","Career Search Tool Job Openings(61).xlsx"))%>%
  mutate(Region=str_replace(Region, "All", "british_columbia"),
         Region=str_replace(Region, "Mainland / Southwest", "mainland_south_west"),
         Region=str_replace(Region, "Vancouver Island / Coast", "vancouver_island_coast"),
         Region=str_replace(Region, "North Coast & Nechako", "north_coast_&_nechako"),
         Region=str_replace(Region, "Cariboo", "cariboo"),
         Region=str_replace(Region, "Kootenay", "kootenay"),
         Region=str_replace(Region, "Northeast", "north_east"),
         Region=str_replace(Region, "Thompson-Okanagan", "thompson_okanagan"),
         `Industry (sub-industry)` = str_replace(`Industry (sub-industry)`, "All", "all_industries")
  )%>%
  clean_tbbl()

cstjo_noc <- cstjo_old%>%
  select(noc, 
         noc_name, 
         `short_noc_name_(not_used_on_the_tool)`,
         education, 
         link, 
         jobbank2, 
         `part-time/full-time`)%>%
  distinct()

cstjo_region <- cstjo_old%>%
  select(region)%>%
  distinct()

cstjo <- crossing(cstjo_noc, cstjo_region, industry_mapping)%>%
  mutate("job_openings{current_year+1}-{current_plus_10}":=NA,
         `salary_(calculated_median_salary)`=NA)%>%
  select(noc,
         noc_name,
         `short_noc_name_(not_used_on_the_tool)`,
         `industry_(sub-industry)`=industry,
         region,
         starts_with("job_openings"),
         education,
         `salary_(calculated_median_salary)`,
         link,
         jobbank2,
         `industry_(aggregate)`=aggregate_industry,
         `part-time/full-time`)

write_csv(cstjo, here("raw_data","career_search_tool_job_openings_2022.csv"))

