library(dplyr)
library(tidyr)
library(flextable)
library(ggplot2)
library(forcats)
library(officer)
library(readxl)

#In a Special Publication that is produced every year by ADF&G (e.g., SP22-11), ADF&G publishes our best estimates for commercial salmon harvests for the coming year and also documents what commercial salmon harvests were for the previous year. 
#This script will be used to assemble table 1 of that report and will replace forecast analyses performed in the "Table 1_2025-values.xlsx" file. 


# Set the forecast year 
Forecast.year<-2025

#===============================================================================
#                       5-year Average Harvests                                #
#===============================================================================
 # RECENT 5-YEAR AVERAGE HARVESTS BY MANAGMENT AREA
## Harvest data will need to be updated each year through OceanAK quires here: 
# “Shared Folders” -> “Commercial Fisheries” -> “Statewide” -> “Salmon” -> “User Reports” -> “Salmon_Season_Summary_Report”

# A few notes regarding this data set:
# Chinook harvests for UCI, LCI, Bristol Bay, and Chignik
# Upper Cook Inlet harvest does not inlude the EEZ (stat code 244-64)
# Starting in 2023, the Northwestern district was grouped in the N Penn but has been historically grouped into the South Penn.
# Staff had infrequently included test fish harvests in season summary totals

# Data import:
CommercialHarvestData <- read_excel("Data/CommercialHarvestData.xlsx")

# Setting the pink salmon life cycle to forecast for
Life.Cycle<-if_else(Forecast.year%% 2 == 0, "even", "odd")

# Lets pivot the data frame for easier manipulation
Harvest.Longer<- CommercialHarvestData%>%
  pivot_longer(cols=c("Chinook","sockeye","coho","pink","chum"), names_to = "Species",values_to = "Harvest")

# Generate recent 5-year average commercial harvest by area and species. This will also serve as the template for Table 1. However, not all areas and species will use the 5-year average harvest.
AverageAreaHarvest<-Harvest.Longer%>%
  mutate(even_odd = if_else(Year%% 2 == 0, "even", "odd"), 
         Condition=ifelse(Species%in%"pink" & even_odd%in%Life.Cycle | Species%in% c("Chinook","sockeye", "coho","chum"),"Yes","No"))%>% # Filtering to ensure we are pulling the recent 5 brood years for pink salmon 
  
  filter(Condition%in%"Yes")%>%
  select(ManagementArea,Year,Species,Harvest)%>% # Removing unwanted columns
  arrange(ManagementArea,Species,Year)%>% #ensuring the dataframe is arranged correctly for slicing
  group_by(ManagementArea,Species)%>%slice_tail(n=5)%>% #pulling the recent 5 entries (recent 5 years for all species other than pinks)
  group_by(ManagementArea,Species)%>%summarize(Harvest=mean(Harvest))%>% #Estimating the recent 5 average harvest
  mutate(Harvest=ifelse(ManagementArea%in%"Southeast" & Species%in% c("Chinook","pink","chum")|
                        ManagementArea%in%"Kodiak" & Species%in% c("sockeye","coho","pink","chum")|
                        ManagementArea%in%"Chignik" & Species%in% "sockeye"|
                        ManagementArea%in%"South Alaska Peninsula" & Species%in% c("pink","sockeye","chum")|
                        ManagementArea%in%"North Alaska Peninsula" & Species%in% "sockeye"|
                        ManagementArea%in%"Upper Cook Inlet" & Species%in% c("Chinook","sockeye"),NA,Harvest))%>% # Upper Cook Inlet Chinook are not a targeted species and the sockeye forecast is the ADFG forecast doc
  pivot_wider(names_from = Species,values_from = Harvest)%>% # flip the dataframe for easier interpretation
  mutate(Region=ifelse(ManagementArea%in%"Southeast", "Southeast Region", #Assigning regions
         ifelse(ManagementArea%in%c("Upper Cook Inlet", "Lower Cook Inlet", "Bristol Bay","Prince william Sound"),"Central Region", "Westward Region")))%>%
  select(Region,ManagementArea,Chinook,sockeye,coho,pink,chum)%>%
  arrange(Region,ManagementArea)

  
#===============================================================================
# PRINCE WILLIAM SOUND PROJECTION #
#===============================================================================
    # CCPH harvest stands for "Commercial Common Property Harvest"

# Import harvest data for the Copper and Bering River districts. Data to update this file can be found in the same quires provided for the general commercial harvest above.
CopperBering <- read_excel("Data/PWSWildHarvest.xlsx")

# Estimating the recent 10 year average harvest for each area by species
CBAverage<-CopperBering%>%
  filter(Year>=Forecast.year-10)%>%
  group_by(Area)%>%summarize(Chinook=mean(Chinook),
                                     sockeye=mean(sockeye),
                                     coho=mean(coho))

# will fill in a black data frame with appropriate values. These numbers come from an array of sources
PWSProjectionData <- read_excel("Data/PWSFrame.xlsx")

# First fill out natural production data
PWSProjectionData[1,3]<- 5000 # Copper River Chinook Common Property Harvest. This is provided by the PWS research biologist
PWSProjectionData[2,3]<- 1000  # PWS Chinook Common Property Harvest. This is provided by the PWS research biologist
PWSProjectionData[3,3]<- 368000    # PWS Wild forecast (Coghill + 10-yr avg PWS wild CCPH (excluding coghill). cells C12+C13
PWSProjectionData[4,3]<- 1872000 # CR 2025 wild CCPH Fcst (comm harvest fcst (1,920,000) - gulkana CCPH fcst (48,000). C11
PWSProjectionData[5,3]<- CBAverage[1,3] #Pulling the recent 10 year average harvest of sockeye from the Bering River District
PWSProjectionData[6,3]<- CBAverage[1,4] + CBAverage[2,4] # Adding the recent 10 year average harvest of coho from the Bering + Copper river districts 
PWSProjectionData[7,3]<- 443000 # pink cell c23 of ccph forecast
PWSProjectionData[8,3]<- 16788000 # chum cell c21 of ccph forecast

# Then hatchery forecasts
PWSProjectionData[9,4]<- 989093 # sockeye Main Bay hatchery forecast that excludes production for sport and personal uses. cell K128 of hatchery forecast excel file provided by the PNP hatchery coordinator
PWSProjectionData[10,4]<- 63391  # # coho hatchery forecast that excludes production for sport and personal uses. cell N128 of hatchery forecast excel file provided by the PNP hatchery coordinator
PWSProjectionData[11,4]<- 2198794 # # chum hatchery forecast that excludes production for sport and personal uses. cell T128 of hatchery forecast excel file provided by the PNP hatchery coordinator
PWSProjectionData[12,4]<- 46990627 # pink hatchery forecast that excludes production for sport and personal uses. cell Q128 of hatchery doc provided by the PNP hatchery coordinator

# Finally broodstock needs
PWSProjectionData[13,5]<- 17462 # sockeye 5 year average broostock (PWS not Copper) needs from cell L153 of hatchery forecast excel file provided by the PNP hatchery coordinator
PWSProjectionData[14,5]<- 1446 # coho 5 year average broostock needs from cell L152 of hatchery forecast excel file provided by the PNP hatchery coordinator
PWSProjectionData[15,5]<- 241206 # chum 5 year average broostock needs from cell L151 of hatchery forecast excel file provided by the PNP hatchery coordinator
PWSProjectionData[16,5]<- 1847942 # pink 5 year average broostock needs from cell L167 of hatchery forecast excel file provided by the PNP hatchery coordinator

GulkanaHarvest<- 48000 # For sockeye, we will need the Projected Gulkana hatchery harvest provided by the PWS Research biologist
  



# We will estimate hatchery harvest by subtracting hatchery needs from hatchery forecasts
PWSwildhatchery<-PWSProjectionData%>%
  group_by(Species)%>%summarize(CommonProperty=sum(CommonProperty,na.rm=T),
                                                  HatcheryProduction=sum(HatcheryProduction,na.rm=T),
                                                  Broodstock=sum(Broodstock,na.rm=T))%>%
  mutate(HatcheryHarvest=HatcheryProduction-Broodstock)%>% #Subtracting broodstock needs
  mutate(HatcheryHarvest=ifelse(Species%in%"sockeye", HatcheryHarvest + GulkanaHarvest,HatcheryHarvest)) # Correcting to the sockeye hatchery harvest using the Gulkana projected harvest
  
# Lets reformat this data frame for later merging into flextables
PWSProjection<-PWSwildhatchery%>%
  select(Species,CommonProperty,HatcheryHarvest)%>%
  rename("NaturalProduction"=CommonProperty,"HatcheryProduction"=HatcheryHarvest)

P1<-PWSProjection[,1:2]%>%mutate(Origin="Natural production")%>%rename("Harvest"=NaturalProduction)
P2<-PWSProjection[,c(1,3)]%>%mutate(Origin="Hatchery production")%>%rename("Harvest"=HatcheryProduction)

PWSProjectionFormatted<-rbind(P1,P2)%>%pivot_wider(names_from = Species,values_from = Harvest)%>%
  mutate(ManagementArea="Prince William Sound", Region="Central Region")




#===============================================================================
#                   Table 1 formatting                                        #
#===============================================================================

# Build table 1 template. We will first gather all the data we need to fill in this table. Not all areas are included here becuase they will be added in later using everage harevst

Table1<-data.frame(Region=as.character(c("Southeast Region", "Central Region", "Central Region", "Westward Region", "Arctic-Yukon-Kuskokwim Region")),
          ManagementArea=as.character(c("Southeast", "Lower Cook Inlet", "Lower Cook Inlet", "Kodiak", "Arctic-Yukon-Kuskokwim")),
          Origin=as.character(c("Hatchery production", "Natural production", "Hatchery production", "Hatchery production","")),
          Chinook=as.numeric(c("", "", "", "", "")),
          sockeye=as.numeric(c("", "", "", "", "")),
          coho=as.numeric(c("", "", "", "", "")),
          pink=as.numeric(c("", "", "", "", "")),
          chum=as.numeric(c("", "", "", "", "")))


### Southeast Region
# Southeast natural production
Table1<-Table1%>%rbind(AverageAreaHarvest%>%filter(ManagementArea%in%"Southeast")%>%mutate(Origin="Natural production")) 

Table1[6,4]   #SE does not formally forecast chinook harvest but rather a PSE is provided. If available, input here
Table1[6,5]<- Table1[6,5]-103681 # Sockeye forecast - 5yr average brood return from PNP forecast document K137.
Table1[6,6]<- Table1[6,6]-595132 # Recent 5yr average harvest - 5yr average brood return from PNP forecast document N137.
Table1[6,7]<- 29000000 - 599043 #Pink harvest is taken from ADF&G forecast document and hatchery forecast subtracted to estimate natural comp
Table1[6,8]<- (14101446/0.91)-14101446 #using the 5-yr average brood chum forecast from PNP forecast document (cell T127) to divide potential hatchery harvest by 91% to expand to total harvest (wild plus hatchery). 
#Then, subtract hatchery amount from the total to estimate wild harvest.

# Southeast Hatchery Production
Table1[1,4]  #SE does not formally forecast chinook harvest but rather a PSE is provided. If available, input here
Table1[1,5]<- 103681 #sockeye harvest taken from PNP forecast document cell k127
Table1[1,6]<- 595132 #coho harvest taken from PNP forecast document cell n127
Table1[1,7]<- 599043 #pink harvest taken from PNP forecast document cell q127
Table1[1,8]<- 14101446 #chum harvest taken from PNP forecast document cell T127



### Central Region
# Prince William Sound- uses the PWS projection data
Table1<-Table1%>%rbind(PWSProjectionFormatted)

Table1[8,4]<- NA #PWS Hatchery component is very small and historically used an NA value for this.




# Lower Cook Inlet Natural Production
Table1[2,4]<- 160 #Chinook harvest taken from ADFG forecast document 
Table1[2,5]<- 156800 #sockeye harvest taken from ADFG forecast document
Table1[2,6]<- 250 #coho harvest taken from ADFG forecast document
Table1[2,7]<- 1050000  #pink harvest taken from ADFG forecast document
Table1[2,8]<- 15800 #chum harvest taken from ADFG forecast document
# Lower Cook Inlet Hatchery Production
Table1[3,4]
Table1[3,5]<- 218002 #sockeye 5-yr average brood from PNP forecast document cell K130
Table1[3,6]<- 0 #coho 5-yr average brood from PNP forecast document cell N130
Table1[3,7]<- 730674 #Pink 5-yr average brood from PNP forecast document cell Q130
Table1[3,8]<- 0 #chum 5-yr average brood from PNP forecast document cell Q130

# Upper Cook Inlet- coho, pink, and chum uses the 5-year harvest
Table1<-Table1%>%rbind(AverageAreaHarvest%>%filter(ManagementArea%in%"Upper Cook Inlet")%>%mutate(Origin=NA)) #coho, chum and pink use the 5-yr average harvest

Table1[9,4] # Upper Cook Inlet Chinook will not have a forecast because it is not a directed species
Table1[9,5]<- 4930000 #available commercial harvest of sockeye taken from the ADFG forecast document

# Bristol Bay
Table1<-Table1%>%rbind(AverageAreaHarvest%>%filter(ManagementArea%in%"Bristol Bay")%>%mutate(Origin=NA)) #Chinook, coho, chum and pink use the 5-yr average harvest
Table1[10,5]<- 36400000 ##available commercial harvest of sockeye taken from the ADFG forecast document




### Westward Region
# Note: These values should be double checked against the table westward provided prior to publication, All natural production estimates are generated by subtracting 
# the estimated hatchery component from the total harvest. This is why the 5-yr averages will not match what is in Ocean AK.
# Kodiak Natural Production
Table1<-Table1%>%rbind(AverageAreaHarvest%>%filter(ManagementArea%in%"Kodiak")%>%mutate(Origin="Natural production")) #Chinook use the 5-yr average harvest

Table1[11,5]<- 1311000 # sockeye harvest projection taken from westward region projection document 
Table1[11,6]<- 165000 # coho harvest projection provided by westward staff in the "westward document". This is not a simple 5-yr average as historical documents suggest.
Table1[11,7]<- 21005000 # pink harvest projection taken from ADF&G forecast document or provided from Kodak staff
Table1[11,8]<- 372000 # chum harvest projection provided by westward staff in the "westward document". This is not a simple 5-yr average as historical documents suggest.

# Kodiak Hatchery Production
Table1[4,4]  # Chinook harvest projection taken from westward region projection document 
Table1[4,5]<- 264000 #sockeye harvest projection taken from westward region projection document 
Table1[4,6]<- 37000 #coho harvest projection taken from westward region projection document
Table1[4,7]<- 10786000 #pink harvest projection taken from westward region projection document
Table1[4,8]<- 268000 #chum harvest projection taken from westward region projection document

# Chignik
Table1<-Table1%>%rbind(AverageAreaHarvest%>%filter(ManagementArea%in%"Chignik")%>%mutate(Origin=NA)) #Chinook, coho, and chum use the 5-yr average harvest

Table1[12,5]<- 757000 #sockeye harvest projection taken from westward region projection document 

# South Penn
Table1<-Table1%>%rbind(AverageAreaHarvest%>%filter(ManagementArea%in%"South Alaska Peninsula")%>%mutate(Origin=NA)) # Average harvest will not match what staff provide because of PU and Test harvests

Table1[13,5]<- 2621000 # sockeye harvest taken from westward region projection document. This is not a simple 5-yr average.
Table1[13,7]<- 10639000 # pink harvest taken from westward region projection document
Table1[13,8]<- 1142000 # chum harvest taken from westward region projection document. This is not a simple 5-yr average.

# North Penn
Table1<-Table1%>%rbind(AverageAreaHarvest%>%filter(ManagementArea%in%"North Alaska Peninsula")%>%mutate(Origin=NA))

Table1[14,5]<- 2118000 # Sockeye harvest taken from westward region projection document
Table1[14,7]<- 134000 # pink harvest taken from westward region projection document

#AYK
Table1[5,4]<- 0 # Chinook 
Table1[5,5]<- 0 #sockeye 
Table1[5,6]<- ((10+25)/2)*1000 #coho 
Table1[5,7]<- ((5+25)/2)*1000 #pink 
Table1[5,8]<- ((1000/2)+((5+15)/2) + ((50+150)/2))*1000 # Yukon summer chum + Norton Sound Summer Chum + Kotzebue Sound Fall chum. Taken from AYK section of report.



### Re-order data frame

# Define the custom order
region_order <- c("Southeast Region", "Central Region", "Westward Region", "Arctic-Yukon-Kuskokwim Region")
Management_order<-c( "Southeast" ,"Prince William Sound" ,"Lower Cook Inlet" ,"Upper Cook Inlet", "Bristol Bay","Kodiak","Chignik","South Alaska Peninsula","North Alaska Peninsula","Arctic-Yukon-Kuskokwim")
# Reorder the 'Region' column based on the custom order
Table1$Region <- factor(Table1$Region, levels = region_order, ordered = TRUE)
Table1$ManagementArea <- factor(Table1$ManagementArea, levels = Management_order,ordered = T)
Table1$Origin <- factor(Table1$Origin, levels = c("Natural production", "Hatchery production","NA"))

# Order the data frame by 'region' and 'origin'
Table1Ordered <- Table1[order(Table1$ManagementArea, Table1$Origin, Table1$Origin), ]


#generating Region and management Area totals
RegionTotals<-Table1%>%group_by(Region)%>%summarize(Origin=NA,
                                                    ManagementArea="Region total",
                                                    Chinook=sum(Chinook,na.rm=T),
                                                    sockeye=sum(sockeye,na.rm = T),
                                                    coho=sum(coho),
                                                    pink=sum(pink),
                                                    chum=sum(chum))%>%
  filter(!Region%in%"Arctic-Yukon-Kuskokwim Region")

RegionTotals[1,4]<-103150 #need to include SE Treaty allocation which is Total allocation - sport allocation

 
StatewideTotal<-Table1Ordered%>%summarize(Region="Statewide total",
                                   Origin=NA,ManagementArea=NA,
                                   Chinook=sum(Chinook,na.rm=T) + 103150, #We have to manually add in SE Chinook here
                                   sockeye=sum(sockeye,na.rm=T),
                                   coho=sum(coho,na.rm=T),
                                   pink=sum(pink,na.rm=T),
                                   chum=sum(chum,na.rm=T))

Table1Ordered<-Table1Ordered%>%rbind(RegionTotals)

Table1Ordered<-Table1Ordered%>%arrange(Region)%>%rbind(StatewideTotal)

# Further refining values
Table1Ordered[17,2]<-"Region total"

# Reformatting region and area totals
Table1formatted<-Table1Ordered%>% 
  rowwise()%>%
  mutate(Total = round(sum(c_across(Chinook:chum), na.rm = TRUE)/1000,digits = 0)) %>%
  ungroup()%>%
  mutate(Chinook=round(Chinook/1000,digits= 0),
         sockeye=round(sockeye/1000,digits= 0),
         coho=round(coho/1000,digits= 0),
         pink=round(pink/1000,digits= 0),
         chum=round(chum/1000,digits= 0))

Table1formatted[8,5]<-NA # We need to hide the UCI comm harvest forecast because this is allocation in nature


# Setting up table properties
sect_properties <- prop_section(
  page_size = page_size(
    orient = "landscape",
    width = 8.3, height = 11),
  type = "oddPage",
  page_margins = page_mar())

border_style <- fp_border(color = "black", width = 1)

#Developing flextable
Table1Final<-Table1formatted%>%
  flextable()%>%
  theme_apa()%>%
  set_header_labels(ManagementArea="Management area")%>%
  

  merge_v(j=c("Region","ManagementArea"), part="body",combine = T)%>% # Merging cells to remove duplicates
  merge_v(j=1)%>%
  valign(j=1:2, valign = "top")%>%
  
  colformat_double(j=4:9,digits = 0)%>% # Formatting cells to remove decimal places
  
  align(j=1:3, align = "left", part = "all")%>% # Re-aligning cell contents
  align(j=4:9, align = "right", part = "all")%>%
  
  add_header_row(values = c("", "Species",""), colwidths = c(3,5,1))%>% # Adding the spanning header
  align(i=1,j=1:9,align = "center",part = "header")%>%
  border(i=1, j = 1:3, border = fp_border(color = NA, width = 0))%>%
  
  hline(i=2, j = 1:9, border = border_style, part = "body")%>% # Adding borders
  hline(i=3, j = 1:9, border = border_style, part = "body")%>%
  
  hline(i=9, j = 1:9, border = border_style, part = "body")%>%
  hline(i=10, j = 1:9, border = border_style, part = "body")%>%
  
  hline(i=15, j = 1:9, border = border_style, part = "body")%>%
  hline(i=16, j = 1:9, border = border_style, part = "body")%>%
  
  hline(i=17, j = 1:9, border = border_style, part = "body")%>%
  hline(i=18, j = 1:9, border = border_style, part = "body")%>%
  
  
footnote( i = c(1,1,  8,8, 9,9,9, 11,11,11, 13,13,13, 14,14,14,14, 15,15,15, 15), 
            j = c(5,6,  6,8, 4,6,8, 4,6,8, 4,6,8, 4,5,6,8, 4,5,6,8),
            value = as_paragraph("Average harvest of the previous five years (2020–2024)."),
            ref_symbols = c("a"),
            part = "body") %>%
  
footnote( i = 2, 
            j = 3,
            value = as_paragraph("Hatchery salmon projections made by Southern Southeast Regional Aquaculture Association, Northern Southeast Regional Aquaculture Association, Douglas Island Pink and Chum, Armstrong-Keta Inc., 
                                 and Metlakatla Indian Community less broodstock (5-year average), and excess. Wild chum salmon catch estimated as 9% of total catch."),
            ref_symbols = c("b"),
            part = "body") %>%  
  
footnote( i = 3, 
            j = 4,
            value = as_paragraph("The allowable catch of Chinook salmon in Southeast Alaska is determined by the Pacific Salmon Commission."),
            ref_symbols = c("c"),
            part = "body")%>%  
  
  
footnote( i = 4, 
            j = 5,
            value = as_paragraph("Includes formal natural harvest estimates for Prince William Sound and Copper/Bering River districts."),
            ref_symbols = c("d"),
            part = "body")%>%

  
footnote( i = 4, 
            j = 6,
            value = as_paragraph("Average harvest of the previous ten years (2015–2024)."),
            ref_symbols = c("e"),
            part = "body") %>%
  
footnote( i = 5, 
            j = 3,
            value = as_paragraph("Hatchery salmon projections made by Prince William Sound Aquaculture Corporation and Valdez Fisheries Development Association. Gulkana Hatchery projection made by ADF&G, less broodstock (5-year average)."),
            ref_symbols = c("f"),
            part = "body")%>%
  
footnote( i = 7, 
            j = 5,
            value = as_paragraph("Hatchery salmon projections made by Cook Inlet Aquaculture Corporation minus broodstock (5-year average)."),
            ref_symbols = c("g"),
            part = "body")%>%
  
  
footnote( i = c(8, 8, 10, 18, 10, 18), 
            j = c(4, 9, 4, 4, 9, 9),
            value = as_paragraph("An Upper Cook Inlet Chinook salmon harvest forecast is not available for 2025."),
            ref_symbols = c("h"),
            part = "body")%>%

  
  
footnote( i = c(8, 8, 10, 18, 10, 18),
            j = c(5, 9, 5, 5, 9, 9),
            value = as_paragraph("An Upper Cook Inlet sockeye salmon commercial harvest forecast is not available for 2025. Central Region and Statewide totals include 4.93 million fish available for harvest to all user groups."),
            ref_symbols = c("i"),
            part = "body")%>%
    
footnote( i = c(8, 9, 13, 15), 
          j = c(7, 7, 7, 7),
            value = as_paragraph("Average harvest of the previous five odd years (2015–2023)."),
            ref_symbols = c("j"),
            part = "body") %>%
  
  
footnote( i = 11, 
          j = 5,
            value = as_paragraph("Total Kodiak harvest of natural run sockeye salmon includes projected harvests from formally forecasted systems, projected Chignik harvest at Cape Igvak, and projected harvest from additional minor systems."),
            ref_symbols = c("k"),
            part = "body")%>%

footnote( i = 12, 
          j = 3,
          value = as_paragraph("Hatchery projections made by Kodiak Regional Aquaculture Association. Sockeye salmon hatchery projections include enhanced Spiridon Lake sockeye salmon run harvest forecast and other Kodiak Regional Aquaculture Association projections."),
          ref_symbols = c("l"),                     
          part = "body")%>%

  
footnote( i = 13, 
          j = 5,
          value = as_paragraph("Chignik sockeye salmon forecast is projected harvest of sockeye within Chignik Management Area."),
          ref_symbols = c("m"),
          part = "body")%>%

  
line_spacing(space = 1, part = "all")%>%
font(fontname = "Times New Roman",part="all")%>%
fontsize(size=10,part="body")%>%
fontsize(size=9,part = "footer")%>%
padding(padding = 0, part = "all")%>%
fix_border_issues()%>%
autofit()

# saving the table to the forecast folder.
Table1Final%>%save_as_docx(path="C:/Users/kpgatt/OneDrive - State of Alaska/Documents/GitHub/Statewide-Forecast/Figures and Tables/Table 1 Final.docx", pr_section = sect_properties)



#===============================================================================
#                         Forecast Figures                                #
#===============================================================================
ActualvsPredicted<-read_excel("Data/ActualvsProjected.xlsx")

ActualvsPredicted.New<-ActualvsPredicted %>%select(1:3)%>%mutate(Harvest=Actual,Type="Actual Harvest")%>%select(-Actual)%>%
  rbind(ActualvsPredicted %>%select(1:2,4)%>%mutate(Harvest=Projected,Type="Projected Harvest")%>%select(-Projected))%>%
  pivot_wider(names_from = Species,values_from = Harvest)%>%
  rbind(Test<-StatewideTotal%>%select(Chinook:chum)%>%
          mutate(Year=2025, 
                 Type="Projected Harvest", 
                 Chinook=NA,
                 sockeye=sockeye/1000000,
                 coho=coho/1000000,
                 pink=pink/1000000,
                 chum=chum/1000000)%>%select(Year, Type, Chinook, sockeye, coho, pink, chum)) # Lets add into the current years projected totals


# Chinook Harvest 
Figure2<-ActualvsPredicted.New%>%
  ggplot(aes(Year,Chinook*1000,color=Type,lty=Type))+
  geom_line(size=1.1)+
  geom_point()+
  ggtitle("Chinook salmon")+
  theme_bw()+
  theme(plot.title = element_text(hjust = 0.5))+
  scale_color_manual(values=c('black', 'blue'))+
  theme(legend.position = c(0.5, 0.1),
        legend.background = element_rect(fill = "white"),
        legend.title = element_blank(),
        legend.key.size = unit(2,"lines"))+
  theme(text=element_text(family = "serif",size=16))+
  scale_x_continuous(breaks=seq(1970,2025,5))+
  theme(panel.grid.minor = element_blank())+
  xlab("")+
  ylab("Projected Harvest (thousands of fish)")


ggsave(Figure2,filename="Figures and Tables/Figure2.jpeg", dpi=320)

# Sockeye Harvest
Figure3<-ActualvsPredicted.New%>%
  ggplot(aes(Year,sockeye,color=Type,lty=Type))+
  geom_line(size=1.1)+
  geom_point()+
  ggtitle("sockeye salmon")+
  theme_bw()+
  theme(plot.title = element_text(hjust = 0.5))+
  scale_color_manual(values=c('black', 'blue'))+
  theme(legend.position = c(0.5, 0.1),
        legend.background = element_rect(fill = "white"),
        legend.title = element_blank(),
        legend.key.size = unit(2,"lines"))+
  theme(text=element_text(family = "serif",size=16))+
  scale_x_continuous(breaks=seq(1970,2025,5))+
  theme(panel.grid.minor = element_blank())+
  xlab("")+
  ylab("Projected Harvest (millions of fish)")

ggsave(Figure3,filename="Figures and Tables/Figure3.jpeg", dpi=320)

# coho Harvest
Figure4<-ActualvsPredicted.New%>%
  ggplot(aes(Year,coho,color=Type,lty=Type))+
  geom_line(size=1.1)+
  geom_point()+
  ggtitle("coho salmon")+
  theme_bw()+
  theme(plot.title = element_text(hjust = 0.5))+
  scale_color_manual(values=c('black', 'blue'))+
  theme(legend.position = c(0.5, 0.1),
        legend.background = element_rect(fill = "white"),
        legend.title = element_blank(),
        legend.key.size = unit(2,"lines"))+
  theme(text=element_text(family = "serif",size=16))+
  scale_x_continuous(breaks=seq(1970,2025,5))+
  theme(panel.grid.minor = element_blank())+
  xlab("")+
  ylab("Projected Harvest (millions of fish)")

ggsave(Figure4,filename="Figures and Tables/Figure4.jpeg", dpi = 320)

# pink Harvest
Figure5<-ActualvsPredicted.New%>%
  ggplot(aes(Year,pink,color=Type,lty=Type))+
  geom_line(size=1.1)+
  geom_point()+
  ggtitle("pink salmon")+
  theme_bw()+
  theme(plot.title = element_text(hjust = 0.5))+
  scale_color_manual(values=c('black', 'blue'))+
  theme(legend.position = c(0.5, 0.1),
        legend.background = element_rect(fill = "white"),
        legend.title = element_blank(),
        legend.key.size = unit(2,"lines"))+
  theme(text=element_text(family = "serif",size=16))+
  scale_x_continuous(breaks=seq(1970,2025,5))+
  theme(panel.grid.minor = element_blank())+
  xlab("")+
  ylab("Projected Harvest (millions of fish)")

 
ggsave(Figure5,filename="Figures and Tables/Figure5.jpeg", dpi=320)

# Chum harvest
Figure6<-ActualvsPredicted.New%>%
  ggplot(aes(Year,chum,color=Type,lty=Type))+
  geom_line(size=1.1)+
  geom_point()+
  ggtitle("chum salmon")+
  theme_bw()+
  theme(plot.title = element_text(hjust = 0.5))+
  scale_color_manual(values=c('black', 'blue'))+
  theme(legend.position = c(0.5, 0.1),
        legend.background = element_rect(fill = "white"),
        legend.title = element_blank(),
        legend.key.size = unit(2,"lines"))+
  theme(text=element_text(family = "serif",size=16))+
  scale_x_continuous(breaks=seq(1970,2025,5))+
  theme(panel.grid.minor = element_blank())+
  xlab("")+
  ylab("Projected Harvest (millions of fish)")

ggsave(Figure6,filename="Figures and Tables/Figure6.jpeg", dpi=320)



#==============================================================================#
                        #Forecast Statistics#
#==============================================================================#

























