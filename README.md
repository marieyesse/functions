# Scripts & Functions

## Helper functions
CEMIS.R: calculates cumulative values (e.g. emissions, Cumulative EMISsions, athough can be anythign ;-) ).
Use via: df <- cemis(df, startyear, endyear) assuming the availability of cols "Year"  and "value" 

## Full scripts
### IEA World Energy Outlook: from XLSX to data.table
WEO_Elec_CO2.R: Script that automatically reads all WEO20XX_AnnexA.xlsx balance sheets (with future projections) for CO2 emissions. 
Requires manual conversion from .xlsb (original file provided by IEA) to .xlsx (for readxl package). Should work if all files are in same folder.

WEO_Balances.R: Script that automatically reads all WEO20XX_AnnexA.xlsx balance sheets (with future projections) for energy consumption (Mtoe). 
Requires manual conversion from .xlsb (original file provided by IEA) to .xlsx (for readxl package). Should work if all files are in same folder.

### Talk of the day! Sample Twitter!
SampleTwitter.R: A script ready to use to sample talking points for specific Topics for the last 7 days. You can fill in a string of keywords (TopicSearch) and a specific filter (GreenTweet)
