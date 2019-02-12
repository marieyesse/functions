# functions

CEMIS.R: calculates cumulative values (e.g. emissions, Cumulative EMISsions).
Use via: df <- CEMIS(df, startyear, endyear) assuming the availability of cols "Year"  and "value" 


WEO_Elec_CO2.R: Script that automatically reads all WEO20XX_AnnexA.xlsx balance sheets (with future projections) for CO2 emissions. 
Requires manual conversion from .xlsb (original file provided by IEA) to .xlsx (for xlreader package). Should work if all files are in same folder.
