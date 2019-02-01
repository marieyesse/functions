CEMIS <- function(df, startyear, endyear){
 
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
 
  df = dfreal[Year %in% c(startyear:endyear)] %>% 
    group_by_at(names(dfreal[, -c("Year", "value")])) %>% 
    mutate(CEMIS=cumsum(value)) 
  
  df = data.table(df)
  
}

