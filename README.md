

### Usage

 Procedure:
 1. Place files whose names contain the following strings in folder "Data Sources".  
 Note that the actual file names will likely contain dates so the script pattern matches on the following strings. 
 
 newtemplate
 Empl_BC_Regions_Industry
 Industry_Profiles
 JobOpenings_BC_Regions_Industry
 BC_Unemployment_rate
 
 2. Ensure that folder "templates" contains the following files
 
 "3.3.1_WorkBC_Career_Profile_Data.xlsx" 
 "3.3.2_WorkBC_Industry_Profile.xlsx"               
 "3.3.3_WorkBC_Regional_Profile_Data.xlsx" 
 "3.4_WorkBC_Career_Compass.xlsx"                   
 "3.5_WorkBC_Career_Trek.xlsx"                      
 "3.7_WorkBC_Buildprint_Builder.xlsx"   

 3. Source file "01_source_me.R"

4. Product can be found in folder "Send to WorkBC"

 Previous Richard's Notes:
 The data for this project was downloaded from the 4CastViewer. The datasets for performing the analysis are in wide format.
 We recently received a new 4CastViewer tool which allows you to download the data in long format.
 Be aware that you'll have to change the dataset names in the "DATA CLEANING" section because it picks up the data source names in the "Data Sources" folder.
 Check with WorkBC that all the templates are updated (Sr No and career job titles)

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
