[![Lifecycle:Experimental](https://img.shields.io/badge/Lifecycle-Experimental-339999)](<Redirect-URL>)

### Usage:

#### Before sourcing file:

 1. Ensure that the following (updated) files are in folder raw_data:
 
-     occupational_characteristics
-     hoo_list
-     pop_municipal
-     career_trek_jobs
-     lmo64_characteristics
-     empl_bc_regions_industry
-     job_openings_bc_regions_industry
-     unemployment_rate
-     career_search_tool_job_openings
-     wages
 
 2.  Ensure that the following files are in folder templates: (request updated versions from WorkBC)
 
-     3.3.1_WorkBC_Career_Profile_Data.xlsx
-     3.3.2_WorkBC_Industry_Profile.xlsx
-     3.3.3_WorkBC_Regional_Profile_Data.xlsx
-     3.4_WorkBC_Career_Compass.xlsx
-     3.5_WorkBC_Career_Trek.xlsx
-     3.7_WorkBC_Buildprint_Builder.xlsx
-     BC Population Distribution.xlsx
-     Career Search Tool Job Openings.xlsx
-     HOO BC and Region for new tool.xlsx
-     HOO List.xlsx
 
3. Source file "01_source_me.R"

4. Output files can be found in folder "processed_data"

### Getting Help or Reporting an Issue

To report bugs/issues/feature requests, please file an [issue](https://github.com/bcgov/boilerplate/issues/).

### How to Contribute

If you would like to contribute, please see our [CONTRIBUTING](CONTRIBUTING.md) guidelines.

Please note that this project is released with a [Contributor Code of Conduct](CODE_OF_CONDUCT.md). By participating in this project you agree to abide by its terms.

### License

```
Copyright 2022 Province of British Columbia

Licensed under the Apache License, Version 2.0 (the &quot;License&quot;);
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an &quot;AS IS&quot; BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and limitations under the License.
```
---
*This project was created using the [bcgovr](https://github.com/bcgov/bcgovr) package.* 
