library(tidyverse)
library(tidytext)
library(readxl)
library(openxlsx)
library(stringr)


d <- file.choose() ##Import the tracker data
tracker <- read_excel(d,sheet = 1, col_names = TRUE)
head(tracker)
##Refine to make data more manageable
refine_tracker <- select(tracker, Name_of_Customer, Region, Form, Category, Date_Received, Escalation, Tags)
##Quarter that you are looking for
y19q1refine_tracker <- filter(refine_tracker, Date_Received >= "2019-02-01" & Date_Received <= "2019-04-30")
y19q2refine_tracker <- filter(refine_tracker, Date_Received >= "2019-05-01" & Date_Received <= "2019-07-31")
y19q3refine_tracker <- filter(refine_tracker, Date_Received >= "2019-08-01" & Date_Received <= "2019-10-31")
y19q4refine_tracker <- filter(refine_tracker, Date_Received >= "2019-11-01" & Date_Received <= "2020-01-31")




##Parse RFPs and CE and Other
RFPs <- filter(y19q1refine_tracker, Form == "Questionnaire" | Form == "Individual Question" | Form == "Document Review" | Form == "EICC/GeSI CM Template")
CE <- filter(y19q1refine_tracker, Form == "Customer Engagement")
Other <- filter(y19q1refine_tracker, Form == "Other")

#RFPs Analysis

Category_RFPs <- group_by(RFPs, Category) %>% summarize(count = n()) %>% arrange(desc(count))
Category_RFPs ##gives Category count of RFPs
RFP_analysis <- Category_RFPs %>% mutate(percentage = count/sum(count)*100) ##adds percentage column to Category_RFPs
RFP_analysis 

Category_RFPs %>% ggplot() +
  geom_col(aes(x = Category, y = count)) ##creates bar graph for Category RFPs


#Text Mining 
taglist <- file.choose()
listtags <- read_excel(taglist, col_names = TRUE)
listtags$Tags<-tolower(listtags$Tags)
RFPs$Tags<-tolower(RFPs$Tags)

tags <- c(listtags)
tags

GenCSR <- filter(RFPs, Category == "General CSR") %>%
  select(Tags)
GenCSR <- c(GenCSR)
GenCSR <- as.data.frame(GenCSR)
GCSRtagcount <- sapply(GenCSR, function(x) {
  sapply(tags[["Tags"]], function(y) {
    sum(grepl(y,x))
  })
}) 
GCSRtagcount

t <- trimws(unlist(str_split(GenCSR$Tags,",",simplify = FALSE)))
t <- as.data.frame(t)
t$t <- as.character(t$t)
t <- unique(t)
t <- arrange(t,t)
names(t) <- "tag"
str(t)
<- distinct(t,tag)

AccountGov <- filter(RFPs, Category == "Accountability / Governance") %>%
  select(Tags)
AG <- c(AccountGov)
AG <- as.data.frame(AG)
AGtagcount <- sapply(AG, function(x) {
  sapply(tags[["Tags"]], function(y) {
    sum(grepl(y,x))
  })
})
AGtagcount

Enviro <- filter(RFPs, Category == "Environment")%>%
  select(Tags)
EV <- c(Enviro)
EV <- as.data.frame(EV)
EVtagcount <- sapply(EV, function(x) {
  sapply(tags[["Tags"]], function (y) {
    sum(grepl(y, x))
  })
})
EVtagcount

RespMfr <- filter(RFPs, Category == "Responsible Sourcing & Mfr")%>%
  select(Tags)
RMfr <- c(RespMfr)
RMfr <- as.data.frame(RMfr)
RMfrtagcount <- sapply(RMfr, function(x) {
  sapply(tags[["Tags"]], function(y) {
    sum(grepl(y,x))
  })
})
RMfrtagcount

Repor <- filter(RFPs, Category == "Reporting")%>%
  select(Tags) %>%
  c() %>% 
  as.data.frame() 
Reportagcount <- sapply(Repor, function(x) {
  sapply(tags[["Tags"]], function(y) {
    sum(grepl(y,x))
  })
})
Reportagcount

DandI <- filter(RFPs, Category == "Diversity / Inclusion") %>%
  select(Tags) %>%
  c() %>%
  as.data.frame()
DandItagcount <- sapply(Repor, function(x) {
  sapply(tags[["Tags"]], function(y) {
    sum(grepl(y,x))
  })
})
DandItagcount


##creates a new excel sheet with the tags in their each new sheet
list_of_tags <- list("GenCSRtags" = GCSRtagcount,"Envtags" = EVtagcount, "AGtags" = AGtagcount, "RMtags" = RMfrtagcount, "Rtags" = Reportagcount, "RFPanalysis" = RFP_analysis)
write.xlsx(list_of_tags, file = "q1TAGdata.xlsx")






Region_RFPs <- group_by(RFPs, Region) %>% summarize(count=n()) %>% arrange(desc(count))
Region_RFPs ## gives Region count of RFPs
Region_RFPs %>% mutate(percentage = count/sum(count)*100) ##adds %

Escalation_RFPs <- group_by(RFPs, Escalation) %>% summarize(count=n()) %>% arrange(desc(count))
Escalation_RFPs ##gives Escalation count of RFPs










