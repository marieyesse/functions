#######################################################################################################################
# THIS FILE: Reading WEO data and casting it into MyM-files
# DATE: May 2019
# Includes: TIMER REGIONS (Native.Region.Code [alphabetic] and DIM_R [numerical])
# Includes: Official ISO2, ISO3, UNI, FAOSTAT,GAUL classification
#######################################################################################################################

# Clear memory
rm(list=ls()) # clear memory

# Working directory
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))


#######################################################################################################################
### Install and/or load required packages / dependencies
#######################################################################################################################

# Easy scripting packages
if(!require(dplyr)){install.packages("dplyr");   library(dplyr)}
if(!require(forcats)){install.packages("forcats");  library(forcats)}
if(!require(data.table)){install.packages("data.table");  library(data.table)}
if(!require(tidyr)){install.packages("tidyr");  library(tidyr)}
if(!require(fuzzyjoin)){install.packages("fuzzyjoin");  library(fuzzyjoin)}
if(!require(stringr)){install.packages("stringr");  library(stringr)}
if(!require(rlang)){install.packages("rlang");  library(rlang)}

# text mining packages
if(!require(tidytext)){install.packages("tidytext");  library(tidytext)}
if(!require(tm)){install.packages("tm");  library(tm)}
if(!require(readxl)){install.packages("readxl");  library(readxl)}
if(!require(tabulizer)){install.packages("tabulizer");  library(tabulizer)}

# data visualisation packages
if(!require(ggplot2)){install.packages("ggplot2");  library(ggplot2)}
if(!require(ggforce)){install.packages("ggforce");  library(ggforce)}
if(!require(plotly)){install.packages("plotly");  library(plotly)}
if(!require(paletteer)){install.packages("paletteer");  library(paletteer)}
if(!require(ggrepel)){install.packages("ggrepel");  library(ggrepel)}
if(!require(igraph)){install.packages("igraph");  library(igraph)}
if(!require(ggraph)){install.packages("ggraph");  library(ggraph)}
if(!require(ggforce)){install.packages("ggforce");  library(ggforce)}
if(!require(shadowtext)){install.packages("shadowtext");  library(shadowtext)}

# MyM cnversion packages
if(!require(mym2r)){install_github("marieyesse/mym2r");  library(mym2r)}


#######################################################################################################################
### Functions (see repository "Timer_tools")
#######################################################################################################################



## Helper function "intrapolate" 

intrapolate <- function(df, startyear, endyear){
  df.intrapolate=seq(unique(df$Year)[1],endyear) 
  
  ## Activate data.table package ###
  dfreal=data.table(df)
  
  #Intrapolatie with approximate
  dfreal = dfreal[,list(approx(x=Year,y=value,xout=df.intrapolate)$y,
                        approx(x=Year,y=value,xout=df.intrapolate)$x),
                  by=eval(dput(names(dfreal[, -c("Year", "value")])))]
  
  setnames(dfreal,"V1","value")
  setnames(dfreal,"V2","Year")
  
  dfreal=data.table(dfreal)
  
}

#######################################################################################################################
### Constants & Conversion tables
#######################################################################################################################


TIMERNEC=data.table(TIMER_CARRIER_NEC8=c("Coal", "Oil", "Gas", "Hydrogen", "CombRen", "Heat", "Traditional Biomass", "Electricity", "Hydro", "Bioenergy", "Biofuels", "Other renewables", "Nuclear"),
                    DIM_2=c(               1,       2,        3,      4,           5,       6,            7,                   8,      5,         5,          5,              5,              5))

IEATIMERNRC=data.table(TIMER_NRC_DIM=c(1:28),
                       TIMER_NRC_REG=c("CAN", "USA","MEX","RCAM", "BRA", "RSAM", "NAF", "WAF", "EAF", "SAF", "WEU", "CEU", "TUR", "UKR", "STAN", "RUS", "ME", "INDIA", "KOR", "CHN",    "SEAS", "INDO", "JAP", "OCE", "RSAS", "RSAF", "dummy", "WORLD"),
                       IEA_reg=c(      "NA",  "US","NA",  "NA",  "NA",   "NA",   "NA",  "NA",  "NA",  "SAFR", "NA", "NA", "NA",   "NA",  "NA",  "NA", "NA",  "INDIA", "NA",  "CHINA",  "NA", "NA",    "JPN", "NA",  "NA",  "NA",   "NA",    "World"  ))


WEO_Energy_indicators=
  '! IEA World Energy Outlook Totals data from the "_Balance" sheet (all in Mtoe),
! The Scenario is indicated in the file name.
! WEO[5,8,28] stands for numbers of WEO publications (5) and Total categories (8) (defined below) and 26 IMAGE regions + dummy[27] + World[28]
! WEOs included:
! 1=2014, 2=2025, 3=2016, 4=2017, 5=2018
! Total categories:
! 1= Buildings,
! 2= Industry,
! 3= Other,
! 4= Other energy sector,
! 5= Power sector,
! 6= TFC,
! 7= TPED,
! 8= Transport'


############################################################################
# Your file(s)
# Mind: it has to be xls or xlsx

WEO_versions <- c(list.files(path = "WEO_publications"))


############################################################################
## Finaly Energy (Mtoe)
############################################################################


## Nested super loop

ALLWEO <- lapply(WEO_versions, function(j) {

  WEODatasheetsENE <- excel_sheets(j)
  WEODatasheetsENE <- WEODatasheetsENE[grepl(c("_Balance"), WEODatasheetsENE)]

  WEORegionsENE  <- gsub("_Balance", "", WEODatasheetsENE, fixed=TRUE) 

  
          # xlsx files
          #Collecting all New Policies (Baseline)
          
          ENERGYDATA.newpol <- lapply( WEODatasheetsENE, function(i) {
                    WEO2018_ENE_0 <- read_excel(j,  sheet=i, range = "A4:H56" )
                    WEO2018_ENE_0$Scenario <- "New Policies Scenarios"
                    
                    reg  <- gsub("_Balance", "", i, fixed=TRUE)  
                    WEO2018_ENE_0$Region <- reg   
                    
                    WEO2018_ENE_0=data.table(WEO2018_ENE_0)
                    setnames(WEO2018_ENE_0, "...1", "Carrier")
                    WEO2018_ENE_0=WEO2018_ENE_0[!Carrier %in% c("Of which: bunkers", "Of which: Bunkers",  "Source: World Energy Outlook 2018") ]
              
                    WEO2018_ENE_0 <- na.omit(WEO2018_ENE_0)
                    WEO2018_ENE_0$Indicator <- if(nrow(WEO2018_ENE_0) == 48) {c(	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Other energy sector",	"Other energy sector",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Transport",	"Transport",	"Transport",	"Transport",	"Transport",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Other")} else {
                      c(	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Other energy sector",	"Other energy sector",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Transport",	"Transport",	"Transport",	"Transport",	"Transport",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Other",	"Other")}
                    WEO2018_ENE_0 <- data.table(WEO2018_ENE_0)
                    
                  })
          
          ENERGYDATA.newpoldat <- rbindlist(ENERGYDATA.newpol)
          
          #Collecting all Current policies 
          
          
          ENERGYDATA.curpol <- lapply( WEODatasheetsENE, function(i) {
                    WEO2018_ENE_0 <- read_excel(j,  sheet=i, range = "M4:Q56" )
                    WEO2018_ENE_0$Scenario <- "Current Policies"
                    
                    reg  <- gsub("_Balance", "", i, fixed=TRUE)  
                    WEO2018_ENE_0$Region <- reg
                    
                    WEO2018_ENE_0=data.table(WEO2018_ENE_0)
                    setnames(WEO2018_ENE_0, "...1", "Carrier")
                    WEO2018_ENE_0=WEO2018_ENE_0[!Carrier %in% c("Of which: bunkers", "Of which: Bunkers",  "Source: World Energy Outlook 2018") ]
                    
                    WEO2018_ENE_0 <- na.omit(WEO2018_ENE_0)
                    WEO2018_ENE_0$Indicator <- if(nrow(WEO2018_ENE_0) == 48) {c(	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Other energy sector",	"Other energy sector",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Transport",	"Transport",	"Transport",	"Transport",	"Transport",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Other")} else {
                      c(	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Other energy sector",	"Other energy sector",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Transport",	"Transport",	"Transport",	"Transport",	"Transport",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Other",	"Other")}
                    WEO2018_ENE_0 <- data.table(WEO2018_ENE_0)
                  })
          
          ENERGYDATA.curpol <- rbindlist(ENERGYDATA.curpol)
          ENERGYDATA.curpoldat <- cbind(ENERGYDATA.curpol,ENERGYDATA.newpoldat[,2:3])
          
          
          #Collecting all Sustainable development scenario
          
          ENERGYDATA.2DS <- lapply( WEODatasheetsENE, function(i) {
                    WEO2018_ENE_0 <- read_excel(j,  sheet=i, range = "Q4:U56" )
                    WEO2018_ENE_1 <- read_excel(j,  sheet=i, range = "M4:M56" )
                    WEO2018_ENE_2 <- read_excel(j,  sheet=i, range = "B4:C56" )
                    
                    WEO2018_ENE_0 <- cbind(WEO2018_ENE_0, WEO2018_ENE_1,WEO2018_ENE_2)
                    WEO2018_ENE_0$Scenario <- "2DS"
                    
                    reg  <- gsub("_Balance", "", i, fixed=TRUE)  
                    WEO2018_ENE_0$Region <- reg 
                    
                    WEO2018_ENE_0=data.table(WEO2018_ENE_0)
                    setnames(WEO2018_ENE_0, "...1", "Carrier")
                    WEO2018_ENE_0=WEO2018_ENE_0[!Carrier %in% c("Of which: bunkers", "Of which: Bunkers",  "Source: World Energy Outlook 2018")]
                    
                    WEO2018_ENE_0 <- na.omit(WEO2018_ENE_0)
                    # Newer WEO versions have included "Traditional biomass" and "Petrochem. feedstock"
                    WEO2018_ENE_0$Indicator <- if(nrow(WEO2018_ENE_0) == 48) {c(	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Other energy sector",	"Other energy sector",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Transport",	"Transport",	"Transport",	"Transport",	"Transport",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Other")} else {
                      c(	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Other energy sector",	"Other energy sector",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Transport",	"Transport",	"Transport",	"Transport",	"Transport",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Other",	"Other")}
                    WEO2018_ENE_0 <- data.table(WEO2018_ENE_0)
                    
                  })
          
          ENERGYDATA.2DS <- rbindlist(ENERGYDATA.2DS)
          # Weird combining effect: TPES Total is not included, have to push rows 1 down
          ENERGYDATA.2DSdat <- ENERGYDATA.2DS
          
          ### Compiling to 1 dataset
          
          WEO.LT.2DS=melt(ENERGYDATA.2DSdat,  id.vars=c("Scenario","Region","Carrier", "Indicator"), variable.name="Year")
          WEO.LT.curpol=melt(ENERGYDATA.curpoldat,  id.vars=c("Scenario","Region","Carrier", "Indicator"), variable.name="Year")
          WEO.LT.newpol=melt(ENERGYDATA.newpoldat,  id.vars=c("Scenario","Region","Carrier", "Indicator"), variable.name="Year")
          
          ENERGYDATA.WEO=rbind(WEO.LT.newpol,WEO.LT.curpol,WEO.LT.2DS)
          #ENERGYDATA.WEO$Year = as.numeric(as.character(substr(ENERGYDATA.WEO$Year, 1, stop=4)))
          
         #ENERGYDATA.WEO = rbind(ENERGYDATA.newpoldat, ENERGYDATA.curpoldat, ENERGYDATA.2DSdat)
          ENERGYDATA.WEO <- na.omit(ENERGYDATA.WEO)
          
          # Here we filter on just the totals. TUSS actually has all...
          # Replacing Carrier names with Totals
          ENERGYDATA.WEO= ENERGYDATA.WEO %>%
            filter(Carrier %in% c(unique(ENERGYDATA.WEO$Indicator))) %>%
            mutate(Carrier="Total")
          
          ENERGYDATA.WEO$WEOYR <- c(as.numeric(substr(j, 4, 7)))
          ENERGYDATA.WEO=data.table(ENERGYDATA.WEO)
          
    })
  

ALLWEOdat   <- rbindlist(ALLWEO)  

## now we start to data clean
ALLWEOdat$filter <- paste(ALLWEOdat$Scenario, ALLWEOdat$Year, ALLWEOdat$WEOYR)
ALLWEOdat$Year <- as.character(ALLWEOdat$Year)
ALLWEOdat=ALLWEOdat[!Year %in% c("...5")]

ALLWEOdat=ALLWEOdat[!filter %in% c("2DS 2040...1 2018", "Current Policies 2025...5 2017", "2DS 2040...4 2017", "Current Policies 2020...5 2016", "2DS 2040...4 2016", "Current Policies 2020...5 2015", "2DS 2040...4 2015", "Current Policies 2020...5 2014", "2DS 2040...4 2014")]
ALLWEOdat$Year <- as.numeric(substr(ALLWEOdat$Year, 1,4))
ALLWEOdat$value <- as.numeric(ALLWEOdat$value)

##########################################################################
## Write table
write.table(ALLWEOdat, file="Compilation_WEO_Balances.csv", dec=".", sep=";", row.names=FALSE)

##########################################################################



############################################################################
####### Transforming to TIMER .dat files
############################################################################

# Similar data
TIMER_ALLWEOdat_SAME=merge(ALLWEOdat %>% rename("IEA_reg"="Region"), IEATIMERNRC,  by=c("IEA_reg"), all=T)

# Template
TIMER_ALLWEOdat_temp=TIMER_ALLWEOdat_SAME %>%
  filter(IEA_reg=="US") %>%
  mutate(value=0, IEA_reg="NA") %>%
  select(-TIMER_NRC_REG, -TIMER_NRC_DIM)

#Missing data
TIMER_ALLWEOdat_NA=merge(IEATIMERNRC, TIMER_ALLWEOdat_temp,  by=c("IEA_reg"), allow.cartesian=TRUE)

#Total file
TIMER_ALLWEOdat=rbind(TIMER_ALLWEOdat_SAME, TIMER_ALLWEOdat_NA)

# transform all columns to numeric DIMs
TIMER_ALLWEOdat_out = TIMER_ALLWEOdat %>%
  na.omit() %>%
  select(Year, Indicator,TIMER_NRC_DIM,  WEOYR, Scenario, value) %>%
  arrange( TIMER_NRC_DIM, Indicator, WEOYR, Scenario, Year) %>%
  intrapolate(., 1990, 2040) %>%
  mutate(value=ifelse(is.na(value)==T, 0, value))  %>%
  select(Year,  Scenario, WEOYR, Indicator, TIMER_NRC_DIM, value) %>%
  arrange(Year, Scenario, WEOYR, Indicator,  TIMER_NRC_DIM ) %>%
  mutate(Scenario=gsub(" ", "", Scenario, fixed=T), WEOYR=as.factor(WEOYR), Indicator=as.factor(Indicator), Year=as.factor(Year), TIMER_NRC_DIM=as.factor(TIMER_NRC_DIM))
  

# Write to M per WEO scenario & per WEO

LoopEnergy <- lapply(unique(TIMER_ALLWEOdat_out$Scenario), function(x) {
 
    TIMER_ALLWEOdat_out3 <- TIMER_ALLWEOdat_out %>%
    mutate(TIMER_NRC_DIM=as.numeric(as.character(TIMER_NRC_DIM))) %>%
    filter(Scenario==x) %>%
    select(-Scenario) 

  
write.r2mym(data = TIMER_ALLWEOdat_out3, 
            outputfile = paste('WEO_Energy_totals_', x , '.dat', sep=""), 
            value.var = 'value', 
            MyM.vartype = 'REAL',
            MyM.varname = 'WEO', 
            comment.line = WEO_Energy_indicators,
            time.dependent = TRUE) 
  
})
