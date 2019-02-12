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
library(devtools)
#devtools::install_github("cmartin/ggConvexHull")
library(ggConvexHull)


############################################################################

#install.packages("readxl")
library("readxl")

############################################################################
# Your file(s)
# Mind: it has to be xls or xlsx

WEO_versions <- c( "WEO2018_AnnexA.xlsx", "WEO2017_AnnexA.xlsx", "WEO2016_AnnexA.xlsx", "WEO2015_AnnexA.xlsx", "WEO2014_AnnexA.xlsx")


############################################################################
## Finaly Energy (Mtoe)
############################################################################

WEO_versions <- c("WEO2018_AnnexA.xlsx", "WEO2017_AnnexA.xlsx", "WEO2016_AnnexA.xlsx", "WEO2015_AnnexA.xlsx", "WEO2014_AnnexA.xlsx")

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
              setnames(WEO2018_ENE_0, "..1", "Carrier")
              WEO2018_ENE_0=WEO2018_ENE_0[!Carrier=="Of which: bunkers"]
              
              
              WEO2018_ENE_0$Indicator <- c(	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Other energy sector",	"Other energy sector",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Transport",	"Transport",	"Transport",	"Transport",	"Transport",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Other",	"Other")
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
              setnames(WEO2018_ENE_0, "..1", "Carrier")
              WEO2018_ENE_0=WEO2018_ENE_0[!Carrier=="Of which: bunkers"]
              
              
              WEO2018_ENE_0$Indicator <- c(	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Other energy sector",	"Other energy sector",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Transport",	"Transport",	"Transport",	"Transport",	"Transport",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Other",	"Other")
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
              setnames(WEO2018_ENE_0, "..1", "Carrier")
              WEO2018_ENE_0=WEO2018_ENE_0[!Carrier %in% c("Of which: bunkers", "Source: World Energy Outlook 2018")]
              
              
              WEO2018_ENE_0$Indicator <- c("NA","TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"TPED",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Power generation",	"Other energy sector",	"Other energy sector",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"TFC",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Industry",	"Transport",	"Transport",	"Transport",	"Transport",	"Transport",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Buildings",	"Other",	"Other")
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
      
      # Replacing top names with Totals
      ENERGYDATA.WEO= ENERGYDATA.WEO[Carrier %in% c(unique(ENERGYDATA.WEO$Indicator)), Carrier:=c(rep("Total", length(unique(ENERGYDATA.WEO$Indicator))*length(unique(WEODatasheetsENE))))]
      
      ENERGYDATA.WEO$WEOYR <- c(as.numeric(substr(j, 4, 7)))
      ENERGYDATA.WEO=data.table(ENERGYDATA.WEO)
      
  })


ALLWEOdat   <- rbindlist(ALLWEO)  

## now we start to data clean
ALLWEOdat$filter <- paste(ALLWEOdat$Scenario, ALLWEOdat$Year, ALLWEOdat$WEOYR)
ALLWEOdat$Year <- as.character(ALLWEOdat$Year)
ALLWEOdat=ALLWEOdat[!Year %in% c("..5")]

ALLWEOdat=ALLWEOdat[!filter %in% c("2DS 2040..1 2018", "Current Policies 2025..5 2017", "2DS 2040..4 2017", "Current Policies 2020..5 2016", "2DS 2040..4 2016", "Current Policies 2020..5 2015", "2DS 2040..4 2015", "Current Policies 2020..5 2014", "2DS 2040..4 2014")]
ALLWEOdat$Year <- as.numeric(substr(ALLWEOdat$Year, 1,4))
ALLWEOdat$value <- as.numeric(ALLWEOdat$value)

##########################################################################
## Write table
write.table(ALLWEOdat, file="Compilation_WEO_Balances.csv", dec=".", sep=";", row.names=FALSE)
