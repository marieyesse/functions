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



WEO_Emission_indicators=
  '! IEA World Energy Outlook Emission Totals data from the "_CO2_EI" sheet (all in Mt),
! The Scenario is indicated in the file name.
! WEO[5,6,28] stands for numbers of WEO publications (5) and Total categories (8) (defined below) and 26 IMAGE regions + dummy[27] + World[28]
! WEOs included:
! 1=2014, 2=2025, 3=2016, 4=2017, 5=2018
! Total categories:
! 1= Buildings,
! 2= Industry,
! 3= Power Generation,
! 4= TFC,
! 5= Total CO2,
! 6= Transport'


############################################################################
# Your file(s)
# Mind: it has to be xls or xlsx

WEO_versions <- c(list.files(path = "WEO_publications"))


############################################################################
## CO2 emissions (Mt)
############################################################################


ALLWEOCO2 <- lapply(WEO_versions, function(j) {

WEODatasheets <- excel_sheets(j)
WEODatasheets <- WEODatasheets[grepl(c("_El"), WEODatasheets)]
     
          # xlsx files
          #Collecting all New Policies (Baseline)
          
          CO2DATA.newpol <- lapply( WEODatasheets, function(i) {

                    WEO2018_CO2_0 <- read_excel(j,  sheet=i, range = if( as.numeric(substr(j, 4,7)) > 2017) {"A38:H56"} else {
                                                                          if(as.numeric(substr(j, 4,7)) < 2017 ) {"A36:H52"} else {"A38:H53"} })
                    WEO2018_CO2_0$Scenario <- "New Policies Scenarios"
            
                    reg  <- gsub("_El_CO2_Ind", "", i, fixed=TRUE)  
                    reg  <- gsub("_El_Ind_CO2", "", reg, fixed=TRUE)  
                    reg  <- gsub("_Elec_CO2_Ind", "", reg, fixed=TRUE)  
                    reg  <- gsub("_Elec_Ind_CO2", "", reg, fixed=TRUE)
                    reg  <- gsub("_Elec_CO2", "", reg, fixed=TRUE)
                    
                    WEO2018_CO2_0$Region <- reg
                    WEO2018_CO2_0=data.table(WEO2018_CO2_0)
                    setnames(WEO2018_CO2_0, "...1", "Carrier",  skip_absent=TRUE)
                    WEO2018_CO2_0 <- na.omit(WEO2018_CO2_0)
                    WEO2018_CO2_0=WEO2018_CO2_0[!Carrier %in% c("Of which: bunkers", "Of which: Bunkers")]
                    WEO2018_CO2_0$Indicator <- if( as.numeric(substr(j, 4,7)) < 2018) {c( "Total CO2",	"Total CO2",	"Total CO2",	"Total CO2",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC")}  else 
                                                          {c( "Total CO2",	"Total CO2",	"Total CO2",	"Total CO2",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",		"Industry",	"Transport", "Buildings")}
                    WEO2018_CO2_0 <- data.table(WEO2018_CO2_0)
                    
                  })
          
          CO2DATA.newpoldat <- rbindlist(CO2DATA.newpol)
          
          #Collecting all Current policies 
          
          
                  CO2DATA.curpol <- lapply( WEODatasheets, function(i) {
                    WEO2018_CO2_0 <- read_excel(j,  sheet=i, range = if (as.numeric(substr(j, 4,7)) >= 2018 ) {"M38:Q56"} else {
                                                                          if (as.numeric(substr(j, 4,7)) < 2017 ) {"M36:P52"} else {"M38:Q53"} } )
                    WEO2018_CO2_0$Scenario <- "Current Policies"
                    
                    reg  <- gsub("_El_CO2_Ind", "", i, fixed=TRUE)  
                    reg  <- gsub("_El_Ind_CO2", "", reg, fixed=TRUE)  
                    reg  <- gsub("_Elec_CO2_Ind", "", reg, fixed=TRUE)  
                    reg  <- gsub("_Elec_Ind_CO2", "", reg, fixed=TRUE)
                    reg  <- gsub("_Elec_CO2", "", reg, fixed=TRUE)
                    
                    WEO2018_CO2_0$Region <- reg 
                    WEO2018_CO2_0=data.table(WEO2018_CO2_0)
                    setnames(WEO2018_CO2_0, "...1", "Carrier",  skip_absent=TRUE)
                    WEO2018_CO2_0=WEO2018_CO2_0[!Carrier %in% c("Of which: bunkers", "Of which: Bunkers")]
                    WEO2018_CO2_0 <- na.omit(WEO2018_CO2_0)
                    WEO2018_CO2_0$Indicator <- if( as.numeric(substr(j, 4,7)) < 2018) {c( "Total CO2",	"Total CO2",	"Total CO2",	"Total CO2",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC")}  else 
                                                                                     {c( "Total CO2",	"Total CO2",	"Total CO2",	"Total CO2",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",		"Industry",	"Transport", "Buildings")}
                    WEO2018_CO2_0 <- data.table(WEO2018_CO2_0)
                  })
          
          CO2DATA.curpoldat <- rbindlist(CO2DATA.curpol)
          CO2DATA.curpoldat <- cbind(CO2DATA.curpoldat,CO2DATA.newpoldat[,2:3])
          
          #Collecting all Sustainable development scenario
          
          CO2DATA.2DS <- lapply( WEODatasheets, function(i) {
                    
                    WEO2018_CO2_0 <- read_excel(j,  sheet=i, range = if(as.numeric(substr(j, 4,7)) > 2017 ) {"Q38:U56"} else {
                       if (as.numeric(substr(j, 4,7)) < 2017 ) {"Q36:S52"} else {"Q38:U53"} } )
                    
                    WEO2018_CO2_1 <- read_excel(j,  sheet=i, range = if(as.numeric(substr(j, 4,7)) > 2017 ) {"M38:M56"} else {
                      if (as.numeric(substr(j, 4,7)) < 2017 ) {"M36:M52"} else {"M38:M53"} } )
                    
                    WEO2018_CO2_2 <- read_excel(j,  sheet=i, range = if(as.numeric(substr(j, 4,7)) > 2017 ) {"B38:C56"} else {
                      if (as.numeric(substr(j, 4,7)) < 2017 ) {"B36:C52"} else {"B38:C53"} } )
                    
                    
                    WEO2018_CO2_0$Scenario <- "2DS"
                    WEO2018_CO2_0 <- cbind(WEO2018_CO2_0, WEO2018_CO2_1,WEO2018_CO2_2)
                    
                    reg  <- gsub("_El_CO2_Ind", "", i, fixed=TRUE)  
                    reg  <- gsub("_El_Ind_CO2", "", reg, fixed=TRUE)  
                    reg  <- gsub("_Elec_CO2_Ind", "", reg, fixed=TRUE)  
                    reg  <- gsub("_Elec_Ind_CO2", "", reg, fixed=TRUE)
                    reg  <- gsub("_Elec_CO2", "", reg, fixed=TRUE)
                    
                    WEO2018_CO2_0$Region <- reg 
                    
                    WEO2018_CO2_0=data.table(WEO2018_CO2_0)
                    setnames(WEO2018_CO2_0, "...1", "Carrier",  skip_absent=TRUE)
                    WEO2018_CO2_0=WEO2018_CO2_0[!Carrier %in% c("Of which: bunkers", "Of which: Bunkers")]
                    WEO2018_CO2_0 <- na.omit(WEO2018_CO2_0)
                    WEO2018_CO2_0$Indicator <- if( as.numeric(substr(j, 4,7)) < 2018) {c( "Total CO2",	"Total CO2",	"Total CO2",	"Total CO2",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC")}  else 
                                                                                       {c( "Total CO2",	"Total CO2",	"Total CO2",	"Total CO2",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",		"Industry",	"Transport", "Buildings")}
                    WEO2018_CO2_0 <- data.table(WEO2018_CO2_0)
            
          })
          
          CO2DATA.2DSdat <- rbindlist(CO2DATA.2DS)
          
          
          ### Compiling to 1 dataset (to Long Table)
          
          WEOCO2.LT.2DS=melt(CO2DATA.2DSdat,  id.vars=c("Scenario","Region","Carrier", "Indicator"), variable.name="Year")
          WEOCO2.LT.curpol=melt(CO2DATA.curpoldat,  id.vars=c("Scenario","Region","Carrier", "Indicator"), variable.name="Year")
          WEOCO2.LT.newpol=melt(CO2DATA.newpoldat,  id.vars=c("Scenario","Region","Carrier", "Indicator"), variable.name="Year")
          
          CO2DATA.WEO=rbind(WEOCO2.LT.newpol,WEOCO2.LT.curpol,WEOCO2.LT.2DS)
          CO2DATA.WEO <- na.omit(CO2DATA.WEO)
          
          # Remove transport (as part of TFC) from list
          CO2DATA.WEO = CO2DATA.WEO %>%
            mutate(Carrier = ifelse(Carrier=="Transport" & Indicator=="TFC", "Transport subset", Carrier))
          
          # Here we filter on just the totals. TUSS actually has all...
          # Replacing Carrier names with Totals
          CO2DATA.WEO= CO2DATA.WEO %>%
            filter(Carrier %in% c(unique(CO2DATA.WEO$Indicator))) %>%
            mutate(Carrier="Total")
          
          
          CO2DATA.WEO$WEOYR <- c(as.numeric(substr(j, 4, 7)))
          CO2DATA.WEO=data.table(CO2DATA.WEO)
          
           })

ALLWEOCO2dat  <- rbindlist(ALLWEOCO2)

## now we start to data clean
ALLWEOCO2dat$filter <- paste(ALLWEOCO2dat$Scenario, ALLWEOCO2dat$Year, ALLWEOCO2dat$WEOYR)
ALLWEOCO2dat$Year <- as.character(ALLWEOCO2dat$Year)
ALLWEOCO2dat=ALLWEOCO2dat[!Year %in% c("...5")]

ALLWEOCO2dat=ALLWEOCO2dat[!filter %in% c("2DS 2040...1 2018", "Current Policies 2025...5 2017", "2DS 2040...4 2017", "Current Policies 2020...5 2016", "2DS 2040...4 2016", "Current Policies 2020...5 2015", "2DS 2040...4 2015", "Current Policies 2020...5 2014", "2DS 2040...4 2014")]
ALLWEOCO2dat$Year <- as.numeric(substr(ALLWEOCO2dat$Year, 1,4))
ALLWEOCO2dat$value <- as.numeric(ALLWEOCO2dat$value)


##########################################################################
## Write table
write.table(ALLWEOCO2dat, file="Compilation_WEO_Elec_CO2.csv", dec=".", sep=";", row.names=FALSE)

##########################################################################


############################################################################
####### Transforming to TIMER .dat files
############################################################################

# Similar data
TIMER_ALLWEOCO2dat_SAME=merge(ALLWEOCO2dat %>% rename("IEA_reg"="Region"), IEATIMERNRC,  by=c("IEA_reg"), all=T)

# unclear how to complete Reg 21 (ASEAN, ASIAPAC, OECDPAC)

# Template
Loop_years <- as.numeric(substr(WEO_versions, 4,7))

TIMER_ALLWEOCO2dat_temp_run <- lapply(Loop_years, function(x){
  TIMER_ALLWEOCO2dat_template=TIMER_ALLWEOCO2dat_SAME %>%
    filter(IEA_reg=="World" & WEOYR=="2018") %>%
    mutate(value=0, IEA_reg="NA") %>%
    select(-TIMER_NRC_REG, -TIMER_NRC_DIM) %>%
    mutate(WEOYR=x)
})
TIMER_ALLWEOCO2dat_temp <- rbindlist(TIMER_ALLWEOCO2dat_temp_run)


TIMER_ALLWEOCO2dat_temp_demand_run <- lapply(Loop_years, function(x){
  TIMER_ALLWEOCO2dat_template=TIMER_ALLWEOCO2dat_SAME %>%
    filter(WEOYR=="2018" & !Indicator %in%  c("TFC", "Total CO2", "Power generation")) %>%
    mutate(value=0) %>%
    select(-TIMER_NRC_REG, -TIMER_NRC_DIM) %>%
    mutate(WEOYR=x)
})

TIMER_ALLWEOCO2dat_demand_temp <-rbindlist(TIMER_ALLWEOCO2dat_temp_demand_run)


#Missing data
TIMER_ALLWEOCO2dat_NA=merge(IEATIMERNRC, rbind(TIMER_ALLWEOCO2dat_temp,TIMER_ALLWEOCO2dat_demand_temp),  by=c("IEA_reg"), allow.cartesian=TRUE)

#Total file
TIMER_ALLWEOCO2dat=rbind(TIMER_ALLWEOCO2dat_SAME, TIMER_ALLWEOCO2dat_NA)



# transform all columns to numeric DIMs
TIMER_ALLWEOCO2dat_out = TIMER_ALLWEOCO2dat %>%
  na.omit() %>%
  select(Year, Indicator,TIMER_NRC_DIM,  WEOYR, Scenario, value) %>%
  arrange( TIMER_NRC_DIM, Indicator, WEOYR, Scenario, Year) %>%
  intrapolate(., 1990, 2040) %>%
  mutate(value=ifelse(is.na(value)==T, 0, value))  %>%
  select(Year,  Scenario, WEOYR, Indicator, TIMER_NRC_DIM, value) %>%
  arrange(Year, Scenario, WEOYR, Indicator,  TIMER_NRC_DIM ) %>%
  mutate(Scenario=gsub(" ", "", Scenario, fixed=T), Scenario=as.factor(Scenario), WEOYR=as.factor(WEOYR), Indicator=as.factor(Indicator), Year=as.factor(Year), TIMER_NRC_DIM=as.factor(TIMER_NRC_DIM))


# Write to M per WEO scenario 

LoopEmissions <- lapply(unique(TIMER_ALLWEOCO2dat_out$Scenario), function(x) {
  
  TIMER_ALLWEOCO2dat_out3 <- TIMER_ALLWEOCO2dat_out %>%
    mutate(TIMER_NRC_DIM=as.numeric(as.character(TIMER_NRC_DIM))) %>%
    filter(Scenario==x) %>%
    select(-Scenario) 
  
  
  write.r2mym(data = TIMER_ALLWEOCO2dat_out3, 
              outputfile = paste('WEO_emission_totals_', x , '.dat', sep=""), 
              value.var = 'value', 
              MyM.vartype = 'REAL',
              MyM.varname = 'WEO', 
              comment.line = WEO_Emission_indicators,
              time.dependent = TRUE) 
  
})

