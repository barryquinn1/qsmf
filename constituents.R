MyDeleteItems<-ls()
rm(list=MyDeleteItems)
library(tidyverse)
library(readxl)
file='~/Dropbox/SMF/OversightCommittee/Davy/QUB Weekly Balances.xlsm'
excel_sheets(file)
consolidated<-read_excel(file,"Consolidated")
Holdings<-consolidated %>%
  select(Date,ISIN,SecAbbrv,Name,Quantity) %>%
  arrange(Name,Date) %>%
  group_by(Name) %>%
  summarise(ISIN=last(ISIN),
            Invested=first(Date),
            Divested=last(Date),
            Quantity=last(Quantity)) %>%
  drop_na(ISIN)

write.csv(Holdings,'All.csv')

# Latest<-max(Holdings$Divested)
# write.csv(Holdings %>% filter(Divested==Latest)
#           ,file='~/Dropbox/SMF/OversightCommittee/Barry/latest.csv')
