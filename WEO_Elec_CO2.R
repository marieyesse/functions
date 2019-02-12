# Clear memory
rm(list=ls()) # clear memory

####### Set working directory (all files should be in same level as script)
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

############################################################################
# Requirements
# Plotting and transforming
library(dplyr)
library(data.table)

# Plotting packages
library(ggplot2)

############################################################################

#install.packages("readxl")
library("readxl")

############################################################################
# Your file(s)
# Mind: it has to be xls or xlsx

WEO_versions <- c( "WEO2018_AnnexA.xlsx", "WEO2017_AnnexA.xlsx", "WEO2016_AnnexA.xlsx", "WEO2015_AnnexA.xlsx", "WEO2014_AnnexA.xlsx")


############################################################################
## CO2 emissions (Mt)
############################################################################


ALLWEOCO2 <- lapply(WEO_versions, function(j) {
  
  WEODatasheets <- excel_sheets(j)
  WEODatasheets <- WEODatasheets[grepl(c("_El"), WEODatasheets)]
  
  # xlsx files
  #Collecting all New Policies (Baseline)
  
  CO2DATA.newpol <- lapply( WEODatasheets, function(i) {
    WEO2018_CO2_0 <- read_excel(j,  sheet=i, range = if (as.numeric(substr(j, 4,7)) < 2017 ) {"A36:H56"} else {"A38:H56"} )
    WEO2018_CO2_0$Scenario <- "New Policies Scenarios"
    
    reg  <- gsub("_El_CO2_Ind", "", i, fixed=TRUE)  
    reg  <- gsub("_Elec_CO2_Ind", "", reg, fixed=TRUE)  
    reg  <- gsub("_Elec_Ind_CO2", "", reg, fixed=TRUE)
    reg  <- gsub("_Elec_CO2", "", reg, fixed=TRUE)
    
    WEO2018_CO2_0$Region <- reg
    WEO2018_CO2_0=data.table(WEO2018_CO2_0)
    setnames(WEO2018_CO2_0, "..1", "Carrier",  skip_absent=TRUE)
    WEO2018_CO2_0=WEO2018_CO2_0[!Carrier=="Of which: bunkers"]
    WEO2018_CO2_0$Indicator <- c( "Total CO2",	"Total CO2",	"Total CO2",	"Total CO2",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",		"Industry",	"Transport", "Buildings")
    WEO2018_CO2_0 <- data.table(WEO2018_CO2_0)
    
  })
  
  CO2DATA.newpoldat <- rbindlist(CO2DATA.newpol)
  
  #Collecting all Current policies 
  
  
  CO2DATA.curpol <- lapply( WEODatasheets, function(i) {
    WEO2018_CO2_0 <- read_excel(j,  sheet=i, range =  if (as.numeric(substr(j, 4,7)) < 2017 ) {"M36:P56"} else {"M38:Q56"}  )
    WEO2018_CO2_0$Scenario <- "Current Policies"
    
    reg  <- gsub("_El_CO2_Ind", "", i, fixed=TRUE)  
    reg  <- gsub("_Elec_CO2_Ind", "", reg, fixed=TRUE)  
    reg  <- gsub("_Elec_Ind_CO2", "", reg, fixed=TRUE)
    reg  <- gsub("_Elec_CO2", "", reg, fixed=TRUE)
    
    WEO2018_CO2_0$Region <- reg 
    WEO2018_CO2_0=data.table(WEO2018_CO2_0)
    setnames(WEO2018_CO2_0, "..1", "Carrier",  skip_absent=TRUE)
    WEO2018_CO2_0=WEO2018_CO2_0[!Carrier=="Of which: bunkers"]
    WEO2018_CO2_0$Indicator <- c( "Total CO2",	"Total CO2",	"Total CO2",	"Total CO2",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",		"Industry",	"Transport", "Buildings")
    WEO2018_CO2_0 <- data.table(WEO2018_CO2_0)
  })
  
  CO2DATA.curpoldat <- rbindlist(CO2DATA.curpol)
  CO2DATA.curpoldat <- cbind(CO2DATA.curpoldat,CO2DATA.newpoldat[,2:3])
  
  #Collecting all Sustainable development scenario
  
  CO2DATA.2DS <- lapply( WEODatasheets, function(i) {
    
    WEO2018_CO2_0 <- read_excel(j,  sheet=i, range = if (as.numeric(substr(j, 4,7)) < 2017 ) {"Q36:S56"} else {"Q38:U56"}  )
    WEO2018_CO2_1 <- read_excel(j,  sheet=i, range = if (as.numeric(substr(j, 4,7)) < 2017 ) {"M36:M56"} else {"M38:M56"}  )
    WEO2018_CO2_2 <- read_excel(j,  sheet=i, range = if (as.numeric(substr(j, 4,7)) < 2017 ) {"B36:C56"} else {"B38:C56"}  )
    
    
    WEO2018_CO2_0$Scenario <- "2DS"
    WEO2018_CO2_0 <- cbind(WEO2018_CO2_0, WEO2018_CO2_1,WEO2018_CO2_2)
    
    reg  <- gsub("_El_CO2_Ind", "", i, fixed=TRUE)  
    reg  <- gsub("_Elec_CO2_Ind", "", reg, fixed=TRUE)  
    reg  <- gsub("_Elec_Ind_CO2", "", reg, fixed=TRUE)
    reg  <- gsub("_Elec_CO2", "", reg, fixed=TRUE)
    
    WEO2018_CO2_0$Region <- reg 
    
    WEO2018_CO2_0=data.table(WEO2018_CO2_0)
    setnames(WEO2018_CO2_0, "..1", "Carrier",  skip_absent=TRUE)
    WEO2018_CO2_0=WEO2018_CO2_0[!Carrier=="Of which: bunkers"]
    WEO2018_CO2_0$Indicator <- c( "Total CO2",	"Total CO2",	"Total CO2",	"Total CO2",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",		"Industry",	"Transport", "Buildings")
    WEO2018_CO2_0 <- data.table(WEO2018_CO2_0)
    
  })
  
  CO2DATA.2DSdat <- rbindlist(CO2DATA.2DS)
  
  
  ### Compiling to 1 dataset (to Long Table)
  
  WEOCO2.LT.2DS=melt(CO2DATA.2DSdat,  id.vars=c("Scenario","Region","Carrier", "Indicator"), variable.name="Year")
  WEOCO2.LT.curpol=melt(CO2DATA.curpoldat,  id.vars=c("Scenario","Region","Carrier", "Indicator"), variable.name="Year")
  WEOCO2.LT.newpol=melt(CO2DATA.newpoldat,  id.vars=c("Scenario","Region","Carrier", "Indicator"), variable.name="Year")
  
  CO2DATA.WEO=rbind(WEOCO2.LT.newpol,WEOCO2.LT.curpol,WEOCO2.LT.2DS)
  CO2DATA.WEO <- na.omit(CO2DATA.WEO)
  
  # Replacing top names with Totals
  CO2DATA.WEO= CO2DATA.WEO[Carrier %in% c(unique(CO2DATA.WEO$Indicator)), Carrier:=c(rep("Total", length(unique(CO2DATA.WEO$Indicator))*length(unique(WEODatasheets))))]
  
  CO2DATA.WEO$WEOYR <- c(as.numeric(substr(j, 4, 7)))
  CO2DATA.WEO=data.table(CO2DATA.WEO)
  
})

ALLWEOCO2dat  <- rbindlist(ALLWEOCO2)

## now we start to data clean
ALLWEOCO2dat$filter <- paste(ALLWEOCO2dat$Scenario, ALLWEOCO2dat$Year, ALLWEOCO2dat$WEOYR)
ALLWEOCO2dat$Year <- as.character(ALLWEOCO2dat$Year)
ALLWEOCO2dat=ALLWEOCO2dat[!Year %in% c("..5")]

ALLWEOCO2dat=ALLWEOCO2dat[!filter %in% c("2DS 2040..1 2018", "Current Policies 2025..5 2017", "2DS 2040..4 2017", "Current Policies 2020..5 2016", "2DS 2040..4 2016", "Current Policies 2020..5 2015", "2DS 2040..4 2015", "Current Policies 2020..5 2014", "2DS 2040..4 2014")]
ALLWEOCO2dat$Year <- as.numeric(substr(ALLWEOCO2dat$Year, 1,4))
ALLWEOCO2dat$value <- as.numeric(ALLWEOCO2dat$value)


##########################################################################
## Write table
write.table(ALLWEOCO2dat, file="Compilation_WEO_Elec_CO2.csv", dec=".", sep=";", row.names=FALSE)

##########################################################################
