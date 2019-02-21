########################################################################################################################
########################################################################################################################

# Clear memory
rm(list=ls()) # clear memory

# Working directory
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

## Required packages
### If not yet availabe, use install.packages("")

# Soc Media API readers
library(rtweet)

# plotting and pipes!
library(ggplot2) 
library(pathwork)
library(dplyr)
library(tidyr) 
library(data.table)
library(paletteer) 
library(ggrepel)
library(igraph)
library(ggraph)

# text mining library
library(tidytext)
library(tm)   

# coupled words analysis
library(widyr)

theme_set(
  theme_bw()
            )

########################################################################################
## API set-up
########################################################################################

token <- create_token(
  app = "RtweetStudies",
  consumer_key = "XsUTxqZNWovUgkNsFctSasLb7", 
  consumer_secret = "XKg0X52FKVEPMBh4lvMIJZaD0qVuj56vaaPLhsWsJ03xLiZnpD",
  access_token = '3092191180-6hcwjXBiyE7SVoHmCUVHOrr5kdi0pgYxIKq0YM6',
  access_secret = 'N6Fk8IZFyE0F4RTWRFoIjFieMNDUS8M162imYoYipgS50')


########################################################################################
## Your analysis input
########################################################################################

Date <-Sys.Date() 

TopicSearch <- c("CEFIC", "CEPI_Paper", "CEMBUREAU", "EUROFER_eu")
GreenTweet <- c("EII", "EULTS", "ClimateNeutralEU", "EP_Environment", "innovation","CO2", "carbon", "emissions", "sustainability", "future", "energyefficiency", "energy","recycling", "circulareconomy", "demand")

########################################################################################

##########################################################################
### Pull recent tweets (7 days)
##########################################################################

Pull_tweets  <- lapply(TopicSearch, function(i) {
                  Tweets <- search_tweets(q = i,
                                  n = 500)
                  Tweets$Keyword <- i
                  Tweets = data.table(Tweets)
            })
        
Pull_tweets_list <- rbindlist(Pull_tweets)


## Save to csv
save_as_csv(Pull_tweets_list, paste("Pull_tweets_list_", Date, " .csv", sep=""), prepend_ids = TRUE, na = "",
            fileEncoding = "UTF-8")

##########################################################################
### Show timeline
##########################################################################

## plot time series (if ggplot2 is installed)
TIMELINE = Pull_tweets_list  %>%
  
  #Filter   
   group_by(Keyword) %>%
   filter(is_retweet==F ) %>%
   
  #Plot
   ts_plot(size=1) +
    theme_bw()+
    theme(text = element_text(size=15),
          strip.text.x = element_text(size=13, face="bold"),
          strip.background = element_rect(colour="white", fill="#FFFFFF"),
          panel.border = element_rect(colour = "black", size=2),
          legend.title = element_text(face="bold"),
          legend.text = element_text(size=16))+
    labs(
          title = "Number of tweets per keyword",
          subtitle = "Tweets collected, parsed, and plotted using `rtweet`") +
    scale_color_paletteer_d(jcolors, rainbow) +
    ggsave(paste("Timeline_Industry_tweets_", Date,".png", sep="")) 
 

##########################################################################
######### Analyse words in last 7 days
##########################################################################

COUNTALL = Pull_tweets_list  %>%

  # Filter out all tweets with mention of Green Tweets
  filter(grepl(paste(GreenTweet, collapse="|"), text)) %>% 
  group_by(Keyword) %>%
  unnest_tokens(word, text) %>%
  count(word, sort = TRUE) %>%

  # Clean up text
  anti_join(stop_words) %>%
  filter(!word %in% c("t.co", "report", "https", tolower(TopicSearch))) %>% 
  top_n(10) %>%
  
  ## reordering
  group_by(word) %>%
  mutate(word_sum=sum(n)) %>%
  ungroup()%>%
  mutate(word = reorder(word, word_sum)) %>%
   
   ggplot(aes(x = word, y = n)) +
  geom_col(aes( fill=Keyword), colour="black") +
  theme_bw() +
  scale_color_paletteer_d(jcolors, rainbow) +
  #facet_wrap(~Keyword, scales = "free") +
  xlab(NULL) +
  #scale_y_log10()+
  coord_flip() +
  scale_fill_paletteer_d(quickpalette, dreaming) +
  guides(fill=guide_legend("Keyword")) +
  theme(text = element_text(size=15),
        strip.text.x = element_text(size=12, face="bold"),
        strip.background = element_rect(colour="white", fill="#FFFFFF"),
        panel.border = element_rect(colour = "black", size=2),
        legend.title = element_text(face="bold"),
        legend.text = element_text(size=16))+
  labs(y = "Number of times used",
       x = "Unique words",
       title = paste("Count of unique words in tweets [", Date, " snapshot]", sep=""))  +
  ggsave(paste("Tweet_industry_BusAssoc_recent_greentweets_",Date,".png", sep=""), width = 15, height = 8, units = c("in")) 


##########################################################################
#### Accounts tweeting on business associations' green tweets:
##########################################################################

COUNTGREEN = Pull_tweets_list  %>%
  filter(grepl(paste(GreenTweet, collapse="|"), text)) %>%
  group_by(Keyword) %>%
  #unnest_tokens(word, text) %>%
  count(screen_name) %>%
  top_n(10) %>%
 ungroup() %>%
  
  # reordering
  group_by(screen_name) %>%
  mutate(screen_name_sum=sum(n)) %>%
  ungroup()%>%
  mutate(screen_name = reorder(screen_name, screen_name_sum)) %>%
  
  #Filter more
  filter(!n <= 1) %>%
  
  #plotting
  ggplot(aes(x = screen_name, y = n)) +
  geom_col(aes( fill=Keyword), colour="black") +
  theme_bw() +
  scale_color_paletteer_d(jcolors, rainbow) +
  #facet_wrap(~Keyword, scales = "free") +
  xlab(NULL) +
  #scale_y_log10()+
  coord_flip() +
  scale_fill_paletteer_d(quickpalette, dreaming) +
  guides(fill=guide_legend("Keyword")) +
  theme(text = element_text(size=15),
        strip.text.x = element_text(size=12, face="bold"),
        strip.background = element_rect(colour="white", fill="#FFFFFF"),
        panel.border = element_rect(colour = "black", size=2),
        legend.title = element_text(face="bold"),
        legend.text = element_text(size=16))+
  labs(y = "Number of times used",
       x = "Unique words",
       title = paste("Count of Tweets per Tweeting person [", Date, " snapshot]", sep=""),
       subtitle=paste("Filtered on:", paste(GreenTweet, collapse=",")))  +
  ggsave(paste("Tweet_industry_BusAssoc_top_tweeters_greentweets",Date,".png", sep=""), width = 10, height = 8, units = c("in")) 


# Compile overview plot (using patvhwork devtools::install_github("thomasp85/patchwork"))

OVERVIEW = TIMELINE + COUNTALL +  plot_layout(ncol =1, heights = c(1, 2))
                 
print(OVERVIEW)
dev.off()
ggsave(paste("OVERVIEW_" ,Date,".png", sep=""), width=8, height=11, unit="in")
