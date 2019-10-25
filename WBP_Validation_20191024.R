#### October 2019 ### 

#This R script was written to help automate data validation to the five needle pine monitoring database. 
#### This R script Will connect to a database, and then produce an Xxcel spreadsheet with a different tab for each error type

rm(list=ls()) # start with a clean slate

setwd("C:/Users/snydera/Desktop/") 

packages <- function(x){
  x <- as.character(match.call()[[2]])
  if (!require(x,character.only = TRUE)) {
    install.packages(pkgs = x,repos = "http://cran.r-project.org")
    require(x,character.only = TRUE)
  }
}

packages(RODBC)
packages(lubridate)
packages(tidyr)
packages(openxlsx)
packages(tidyverse)
packages(stringr)
packages(distr)
packages(dplyr)



library("RODBC")
library("lubridate")
library("tidyr")
library("openxlsx")
library(tidyverse)
library("stringr")
library("distr")
library("dplyr")


options(stringsAsFactors = FALSE) 


##Connect to the correct Whitebark Pine database
connection <- odbcConnectAccess("FNP_MASTER_2019_20191025.mdb")
#odbcDriverConnect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:/Users/snydera/Desktop/FNP_MASTER_2019_20191023.accdb")

# Create a new dataframe object that will store the tab descriptions, only if the QA queiries below return records and a tab gets exported. Adds a header row. This dataframe will get exported into the XLSX at the end of the code.
TableDefs <-data.frame("Tab Name", "Definition")

#Import Location Data
Locations <- sqlFetch(connection,"tbl_Locations")

#Import Species Lookup Table
SpeciesList <- sqlFetch(connection,"tlu_Species_Parks")

#Import Events Data
Events<- sqlQuery(connection, "SELECT tbl_Locations.PlotID_Number, tbl_Events.Start_Date, tbl_Events.Event_ID, tbl_Events.Location_ID
FROM tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID")

##Start XLS file to hold data errors
#write.xlsx(Events, file= "WBP_Validation.xlsx", sheetName = "EventsList_NoErrors", row.names = FALSE, showNA = FALSE)

OutputFilename<- paste("WBP_Validation_",format(now(), "%Y%m%d_%H%M%S"),".xlsx",sep="")
## Create a new workbook, add worksheet, write data to sheet, save and add to TableDefs

wb <- createWorkbook("WBP_Validation")
addWorksheet(wb, "EventsList_NoErrors")
writeData(wb, "EventsList_NoErrors", Events, rowNames = FALSE)
saveWorkbook(wb, OutputFilename, overwrite = TRUE)
TableDefs[nrow(TableDefs) + 1,] = list("EventsList_NoErrors","Events that have no errors")


#Import Tree Data
DataTrees <- sqlQuery(connection, "SELECT tbl_Sites.Unit_Code, tbl_Locations.PlotID_Number, tbl_Events.Start_Date, tbl_Locations.Panel_ID, tbl_Locations.Panel_YrEst, tbl_EvData_TreeData.TreeData_ID, tbl_EvData_TreeData.Event_ID, tbl_EvData_TreeData.TreeData_SubPlot_StripID, tbl_EvData_TreeData.TreeData_SubPlot_StripNotes, tbl_EvData_TreeData.TreeID_Number, tbl_EvData_TreeData.Species_Code, tbl_EvData_TreeData.Clump_Number, tbl_EvData_TreeData.Stem_Letter, tbl_EvData_TreeData.Tag_Moved, tbl_EvData_TreeData.TreeHeight_m, tbl_EvData_TreeData.TreeDBH_cm, tbl_EvData_TreeData.Krumholtz_YN, tbl_EvData_TreeData.Tree_Status, tbl_EvData_TreeData.StatusDead_Cause, tbl_EvData_TreeData.Crown_Health, tbl_EvData_TreeData.CrownKill_Upper_perc, tbl_EvData_TreeData.CrownKill_Mid_perc, tbl_EvData_TreeData.CrownKill_Lower_perc, tbl_EvData_TreeData.BranchCanks_A_Upper_YN, tbl_EvData_TreeData.BranchCanks_I_Upper_YN, tbl_EvData_TreeData.BranchCanks_ITypes_Upper, tbl_EvData_TreeData.BranchCanks_A_Mid_YN, tbl_EvData_TreeData.BranchCanks_I_Mid_YN, tbl_EvData_TreeData.BranchCanks_ITypes_Mid, tbl_EvData_TreeData.BranchCanks_A_Lower_YN, tbl_EvData_TreeData.BranchCanks_I_Lower_YN, tbl_EvData_TreeData.BranchCanks_ITypes_Lower, tbl_EvData_TreeData.BoleCankers_A_Upper_YN, tbl_EvData_TreeData.BoleCankers_I_Upper_YN, tbl_EvData_TreeData.BoleCanks_ITypes_Upper, tbl_EvData_TreeData.BoleCankers_A_Mid_YN, tbl_EvData_TreeData.BoleCankers_I_Mid_YN, tbl_EvData_TreeData.BoleCanks_ITypes_Mid, tbl_EvData_TreeData.BoleCankers_A_Lower_YN, tbl_EvData_TreeData.BoleCankers_I_Lower_YN, tbl_EvData_TreeData.BoleCanks_ITypes_Lower, tbl_EvData_TreeData.PineBeetle_JGalleries_YN, tbl_EvData_TreeData.PineBeetle_PitchTube_YN, tbl_EvData_TreeData.PineBeetle_Frass_YN, tbl_EvData_TreeData.Mistletoe_YN, tbl_EvData_TreeData.FemaleCones_YN, tbl_EvData_TreeData.Cone_Count, tbl_EvData_TreeData.TreeData_Notes, tbl_EvData_TreeData.Mort_Year, tbl_EvData_TreeData.Tree_FlagID, tbl_EvData_TreeData.IsProofed, tbl_EvData_TreeData.Tree_DataCertID
FROM ((tbl_Sites INNER JOIN tbl_Locations ON tbl_Sites.[Site_ID] = tbl_Locations.[Site_ID_F]) INNER JOIN tbl_Events ON tbl_Locations.[Location_ID] = tbl_Events.[Location_ID]) INNER JOIN tbl_EvData_TreeData ON tbl_Events.[Event_ID] = tbl_EvData_TreeData.[Event_ID];")

#Import Seedling Data
DataSeedlings <- sqlQuery(connection, "SELECT tbl_Sites.Unit_Code, tbl_Locations.PlotID_Number, tbl_Events.Start_Date, tbl_EvData_SeedlingCounts.SeedlingCount_ID, tbl_EvData_SeedlingCounts.Event_ID, tbl_EvData_SeedlingCounts.Seedling_SubPlot_ID, tbl_EvData_SeedlingCounts.Species_Code, tbl_EvData_SeedlingCounts.Height_Class, tbl_EvData_SeedlingCounts.SeedlingTag, tbl_EvData_SeedlingCounts.Status, tbl_EvData_SeedlingCounts.Seedling_SubPlot_Notes, tbl_EvData_SeedlingCounts.Seedling_FlagID, tbl_EvData_SeedlingCounts.IsProofed, tbl_EvData_SeedlingCounts.Seedling_DataCertID, tbl_EvData_SeedlingCounts.Death_Cause
FROM ((tbl_Sites INNER JOIN tbl_Locations ON tbl_Sites.Site_ID = tbl_Locations.Site_ID_F) INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) INNER JOIN tbl_EvData_SeedlingCounts ON tbl_Events.Event_ID = tbl_EvData_SeedlingCounts.Event_ID;")

#Import Photo Data
DataPhotos <- sqlQuery(connection, "SELECT tbl_Sites.Unit_Code, tbl_Locations.PlotID_Number, tbl_Events.Start_Date, tbl_EvData_PlotPhotos.PlotPhoto_ID, tbl_EvData_PlotPhotos.Event_ID, tbl_EvData_PlotPhotos.PlotPhoto_Number, tbl_EvData_PlotPhotos.PlotPhoto_File_Name, tbl_EvData_PlotPhotos.PlotPhoto_File_Path, tbl_EvData_PlotPhotos.PlotPhoto_Loc_Ref, tbl_EvData_PlotPhotos.PlotPhoto_Bear_deg, tbl_EvData_PlotPhotos.Camera_ImageID, tbl_EvData_PlotPhotos.PlotPhoto_Notes, tbl_EvData_PlotPhotos.PlotPhoto_Date, tbl_EvData_PlotPhotos.Photo_FlagID, tbl_EvData_PlotPhotos.IsProofed, tbl_EvData_PlotPhotos.Photo_DataCertID
FROM (tbl_Sites INNER JOIN (tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID) ON tbl_Sites.Site_ID = tbl_Locations.Site_ID_F) INNER JOIN tbl_EvData_PlotPhotos ON tbl_Events.Event_ID = tbl_EvData_PlotPhotos.Event_ID;")

Spatial <- sqlFetch(connection, "tbl_Locations")

###Looking for Events with missing data

#Looking for events with no tree records- if records exist, write to xlsx
EventsNoTrees <- merge(Events, DataTrees, by = 'Event_ID', all = TRUE)
EventsNoTrees <- subset(EventsNoTrees, (is.na(EventsNoTrees$PlotID_Number.y) & !is.na(EventsNoTrees$PlotID_Number.x)))
EventsNoTrees <-EventsNoTrees%>%
  select(PlotID_Number.x, Event_ID, Start_Date.x, Unit_Code)


#TestEventNoTrees<- if(nrow(EventsNoTrees)>0) {
#write.xlsx(EventsNoTrees, file= "WBP_Validation.xlsx", sheetName = "EventsNoTrees", append = TRUE, row.names = FALSE, showNA = FALSE)
#} 

TestEventNoTrees<- if(nrow(EventsNoTrees)>0) {
  addWorksheet(wb, "EventsNoTrees")
  writeData(wb, "EventsNoTrees", EventsNoTrees, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("EventsNoTrees","Events with no related tree records")
} 





#Looking for events with no photo records- if records exist, write to xlsx
EventsNoPhotos <- merge(Events, DataPhotos, by = 'Event_ID', all = TRUE)
EventsNoPhotos <- subset(EventsNoPhotos, (is.na(EventsNoPhotos$PlotID_Number.y) & !is.na(EventsNoPhotos$PlotID_Number.x)))
EventsNoPhotos <-EventsNoPhotos%>%
  select(PlotID_Number.x, Event_ID, Start_Date.x, Unit_Code)

TestEventsNoPhotos<- if(nrow(EventsNoPhotos)>0) {
  addWorksheet(wb, "EventsNoPhotos")
  writeData(wb, "EventsNoPhotos", EventsNoPhotos, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("EventsNoPhotos","Events with no related photo records")
} 

#Looking for events with no seedling records- if records exist, write to xlsx
EventsNoSeedlings <- merge(Events, DataSeedlings, by = 'Event_ID', all = TRUE)
EventsNoSeedlings <- subset(EventsNoSeedlings, (is.na(EventsNoSeedlings$PlotID_Number.y) & !is.na(EventsNoSeedlings$PlotID_Number.x)))
EventsNoSeedlings <-EventsNoSeedlings%>%
  select(PlotID_Number.x, Event_ID, Start_Date.x, Unit_Code)

TestEventsNoSeedlings<- if(nrow(EventsNoSeedlings)>0) {
  addWorksheet(wb, "EventsNoSeedlings")
  writeData(wb, "EventsNoSeedlings", EventsNoSeedlings, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("EventsNoSeedlings","Events with no related seedling records")
} 

###Photo Validation

#Looking for photo records with bearing greater than 360- if records exist, write to xlsx
PhotosBearing <- subset(DataPhotos, ((DataPhotos$PlotPhoto_Bear_deg > 360)))
PhotosBearing <-PhotosBearing%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotosBearing<- if(nrow(PhotosBearing)>0) {
  addWorksheet(wb, "PhotosBearing")
  writeData(wb, "PhotosBearing", PhotosBearing, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotosBearing","Photos records with bearing > 360")
} 


#Looking for events where there are less than 4 photos- if records exist, write to xlsx
PhotoClean <- DataPhotos[,c('Unit_Code', 'Start_Date', 'PlotID_Number')]
PhotoCount2 <- add_count(PhotoClean, (PhotoClean$PlotID_Number))
PhotoCount3 <- distinct(PhotoCount2, (PhotoCount2$PlotID_Number), (PhotoCount2$n))
names(PhotoCount3) <- c("PlotID_Number", "n")
PhotoCount4 <- subset(PhotoCount3, ((PhotoCount3$n < 4)))

TestPhotoCount4<- if(nrow(PhotoCount4)>0) {
  addWorksheet(wb, "PhotoCount4")
  writeData(wb, "PhotoCount4", PhotoCount4, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotoCount4","Events where there are less than 4 photos")
} 



#Returns photo records that don't have a domain value for PlotPhoto_Loc_Ref- if records exist, write to xlsx
PhotosLocRef <- subset(DataPhotos, ((DataPhotos$PlotPhoto_Loc_Ref != "SW_Corner" & DataPhotos$PlotPhoto_Loc_Ref != "NW_Corner" & DataPhotos$PlotPhoto_Loc_Ref != "NE_Corner" & DataPhotos$PlotPhoto_Loc_Ref != "SE_Corner" & DataPhotos$PlotPhoto_Loc_Ref != "See_Notes")))

PhotosLocRef <-PhotosLocRef%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotosLocRef<- if(nrow(PhotosLocRef)>0) {
  addWorksheet(wb, "PhotosLocRef")
  writeData(wb, "PhotosLocRef", PhotosLocRef, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotosLocRef","Photo records that don't have a domain value for PlotPhoto_Loc_Ref")
} 

#TestPhotoLocRef<- if(nrow(PhotosLocRef)>0) {
#  write.xlsx(PhotosLocRef, file= "WBP_Validation.xlsx", sheetName = "PhotosNoLocRef", append = TRUE, row.names = FALSE, showNA = FALSE)
#} 

#Returns photo records that are missing photo number - if records exist, write to xlsx
PhotosDataMissingNumber <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_Number)))
PhotosDataMissingNumber <-PhotosDataMissingNumber%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotosDataMissingNumber<- if(nrow(PhotosDataMissingNumber)>0) {
  addWorksheet(wb, "PhotosDataMissingNumber")
  writeData(wb, "PhotosDataMissingNumber", PhotosDataMissingNumber, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotosDataMissingNumber","Photo records that are missing photo number")
} 

#TestPhotoMissingNum<- if(nrow(PhotosDataMissingNumber)>0) {
#  write.xlsx(PhotosDataMissingNumber, file= "WBP_Validation.xlsx", sheetName = "PhotosMissingNum", append = TRUE, row.names = FALSE, showNA = FALSE)
#}

#Returns photo records that are missing file name- if records exist, write to xlsx 
PhotosDataMissingFileName <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_File_Name)))
PhotosDataMissingFileName <-PhotosDataMissingFileName%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotosDataMissingFileName<- if(nrow(PhotosDataMissingFileName)>0) {
  addWorksheet(wb, "PhotosDataMissingFileName")
  writeData(wb, "PhotosDataMissingFileName", PhotosDataMissingFileName, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotosDataMissingFileName","Photo records that are missing file name")
} 


#Returns photo records that are missing file path- if records exist, write to xlsx
PhotosDataMissingFilePath <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_File_Path)))
PhotosDataMissingFilePath <-PhotosDataMissingFilePath%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotosDataMissingFilePath<- if(nrow(PhotosDataMissingFilePath)>0) {
  addWorksheet(wb, "PhotosDataMissingFilePath")
  writeData(wb, "PhotosDataMissingFilePath", PhotosDataMissingFilePath, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotosDataMissingFilePath","Photo records that are missing file path")
} 


#Returns photo records that are missing location reference- if records exist, write to xlsx
PhotosDataMissingLocRef <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_Loc_Ref)))
PhotosDataMissingLocRef <-PhotosDataMissingLocRef%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotosDataMissingLocRef<- if(nrow(PhotosDataMissingLocRef)>0) {
  addWorksheet(wb, "PhotosDataMissingLocRef")
  writeData(wb, "PhotosDataMissingLocRef", PhotosDataMissingLocRef, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotosDataMissingLocRef","Photo records that are missing location reference")
} 



#Returns photo records that are missing bearing- if records exist, write to xlsx 
PhotosDataMissingBearing <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_Bear_deg)))
PhotosDataMissingBearing <-PhotosDataMissingBearing%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)


TestPhotosDataMissingBearing<- if(nrow(PhotosDataMissingBearing)>0) {
  addWorksheet(wb, "PhotosDataMissingBearing")
  writeData(wb, "PhotosDataMissingBearing", PhotosDataMissingBearing, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotosDataMissingBearing","Photo records that are missing bearing")
} 

#Returns photo records that are missing image ID- if records exist, write to xlsx 
PhotosDataMissingImageID <- subset(DataPhotos, (is.na(DataPhotos$Camera_ImageID)))
PhotosDataMissingImageID <-PhotosDataMissingImageID%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotosDataMissingImageID<- if(nrow(PhotosDataMissingImageID)>0) {
  addWorksheet(wb, "PhotosDataMissingImageID")
  writeData(wb, "PhotosDataMissingImageID", PhotosDataMissingImageID, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotosDataMissingImageID","Photo records that are missing image ID")
} 

#Returns photo records that are missing photo date- if records exist, write to xlsx  
PhotosDataMissingDate <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_Date)))
PhotosDataMissingDate <-PhotosDataMissingDate%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Date, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotosDataMissingDate<- if(nrow(PhotosDataMissingDate)>0) {
  addWorksheet(wb, "PhotosDataMissingDate")
  writeData(wb, "PhotosDataMissingDate", PhotosDataMissingDate, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("PhotosDataMissingDate","Photo records that are missing photo date")
} 


###Seedling Validation

#Returns seedling records that are recorded as Dead but don't have a Death Cause- if records exist, write to xlsx 
SeedlingDataDeathCause <- subset(DataSeedlings, ((DataSeedlings$Status == "D" & is.na(DataSeedlings$Death_Cause))))
SeedlingDataDeathCause <-SeedlingDataDeathCause%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedlingDataDeathCause<- if(nrow(SeedlingDataDeathCause)>0) {
  addWorksheet(wb, "SeedlingDataDeathCause")
  writeData(wb, "SeedlingDataDeathCause", SeedlingDataDeathCause, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataDeathCause","Seedling records that are recorded as Dead but don't have a Death Cause")
} 



#Returns seedlings records that don't have a height domain value- if records exist, write to xlsx 
SeedlingDataHeight <- subset(DataSeedlings, ((DataSeedlings$Height_Class != "20 - <50 cm" & DataSeedlings$Height_Class != "50 - <100 cm" & DataSeedlings$Height_Class != "100 - <137 cm")))
SeedlingDataHeight <-SeedlingDataHeight%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedlingDataHeight<- if(nrow(SeedlingDataHeight)>0) {
  addWorksheet(wb, "SeedlingDataHeight")
  writeData(wb, "SeedlingDataHeight", SeedlingDataHeight, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataHeight","Seedling records that don't have a height domain value")
} 


#Returns seedlings records that are missing data in subplot ID- if records exist, write to xlsx 
SeedlingDataMissingSubplot <- subset(DataSeedlings, ((is.na(DataSeedlings$Seedling_SubPlot_ID))))
SeedlingDataMissingSubplot <-SeedlingDataMissingSubplot%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedlingDataMissingSubplot<- if(nrow(SeedlingDataMissingSubplot)>0) {
  addWorksheet(wb, "SeedlingDataMissingSubplot")
  writeData(wb, "SeedlingDataMissingSubplot", SeedlingDataMissingSubplot, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataMissingSubplot","Seedling records that are missing data in subplot ID")
} 

#Returns seedlings records that are missing data in species- if records exist, write to xlsx 
SeedlingDataMissingSpCode <- subset(DataSeedlings, ((is.na(DataSeedlings$Species_Code))))
SeedlingDataMissingSpCode <-SeedlingDataMissingSpCode%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedlingDataMissingSpCode<- if(nrow(SeedlingDataMissingSpCode)>0) {
  addWorksheet(wb, "SeedlingDataMissingSpCode")
  writeData(wb, "SeedlingDataMissingSpCode", SeedlingDataMissingSpCode, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataMissingSpCode","Seedling records that are missing data in species")
} 



#Returns seedlings records where status = "L" and height is missing- if records exist, write to xlsx 
SeedlingDataMissingHt <- subset(DataSeedlings, (((DataSeedlings$Status == "L" | DataSeedlings$Status == "RD") & is.na(DataSeedlings$Height_Class))))
SeedlingDataMissingHt <-SeedlingDataMissingHt%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedlingDataMissingHt<- if(nrow(SeedlingDataMissingHt)>0) {
  addWorksheet(wb, "SeedlingDataMissingHt")
  writeData(wb, "SeedlingDataMissingHt", SeedlingDataMissingHt, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataMissingHt","Seedling records where status = 'L' and height is missing")
} 



#Returns seedlings records that are missing a tag number- if records exist, write to xlsx 
SeedlingDataMissingTag <- subset(DataSeedlings, (((is.na(DataSeedlings$Species_Code) | DataSeedlings$Species_Code != "_NONE" & DataSeedlings$Species_Code != "_NotSampled") & is.na(DataSeedlings$SeedlingTag))))
SeedlingDataMissingTag <-SeedlingDataMissingTag%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedlingDataMissingTag<- if(nrow(SeedlingDataMissingTag)>0) {
  addWorksheet(wb, "SeedlingDataMissingTag")
  writeData(wb, "SeedlingDataMissingTag", SeedlingDataMissingTag, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataMissingTag","Seedling records that are missing a tag number")
} 



#Returns seedlings records that are missing status- if records exist, write to xlsx 
SeedlingDataMissingStatus<- subset(DataSeedlings, (((is.na(DataSeedlings$Species_Code) | DataSeedlings$Species_Code != "_NONE" & DataSeedlings$Species_Code != "_NotSampled") & is.na(DataSeedlings$Status))))
SeedlingDataMissingStatus <-SeedlingDataMissingStatus%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedlingDataMissingStatus<- if(nrow(SeedlingDataMissingStatus)>0) {
  addWorksheet(wb, "SeedlingDataMissingStatus")
  writeData(wb, "SeedlingDataMissingStatus", SeedlingDataMissingStatus, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataMissingStatus","Seedling records that are missing status")
} 


#Returns seedlings records with a species that doesn't match a species in the lookup table- if records exist, write to xlsx 
names(SpeciesList)[1]<- "Species_Code" #renames field so that I can merge data
names(SpeciesList)[2]<- "Unit_Code" #renames field so that I can merge data
SeedlingNoSp <- subset(DataSeedlings, ((!is.na(DataSeedlings$Species_Code)))) #removes records where no seedling species was recorded
SeedlingNoSpJoin <- anti_join(SeedlingNoSp, SpeciesList, by = 'Species_Code', 'Unit_Code') #joins lookup table to data table
SeedlingNoSpJoin <-SeedlingNoSpJoin%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedlingNoSpJoin<- if(nrow(SeedlingNoSpJoin)>0) {
  addWorksheet(wb, "SeedlingNoSpJoin")
  writeData(wb, "SeedlingNoSpJoin", SeedlingNoSpJoin, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingNoSpJoin","Seedling records with a species that doesn't match a species in the lookup table")
} 


#Returns seedlings records where status doesn't equal a domain value (L, RD, D)- if records exist, write to xlsx 
SeedlingDataStatus <- subset(DataSeedlings, ((DataSeedlings$Status != "L" & DataSeedlings$Status != "D" & DataSeedlings$Status != "RD")))
SeedlingDataStatus <-SeedlingDataStatus%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)


TestSeedlingDataStatus<- if(nrow(SeedlingDataStatus)>0) {
  addWorksheet(wb, "SeedlingDataStatus")
  writeData(wb, "SeedlingDataStatus", SeedlingDataStatus, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataStatus","Seedling records where status doesn't equal a domain value (L, RD, D)")
} 


#Returns seedlings records that have a subplot number that is not between 1 and 9- if records exist, write to xlsx 
SeedlingDataSubplot <- subset(DataSeedlings, ((DataSeedlings$Seedling_SubPlot_ID < 1 | DataSeedlings$Seedling_SubPlot_ID > 9)))
SeedlingDataSubplot <-SeedlingDataSubplot%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedlingDataSubplot<- if(nrow(SeedlingDataSubplot)>0) {
  addWorksheet(wb, "SeedlingDataSubplot")
  writeData(wb, "SeedlingDataSubplot", SeedlingDataSubplot, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataSubplot","Seedling records that have a subplot number that is not between 1 and 9")
} 



#Returns event records where there are not 9 seedling subplots- if records exist, write to xlsx 
DataSeedlingsClean <- DataSeedlings[,c('Unit_Code', 'Start_Date', 'PlotID_Number','Seedling_SubPlot_ID')]
SeedlingDataSubplotCount2 <- distinct(DataSeedlingsClean, (DataSeedlingsClean$PlotID_Number), (DataSeedlingsClean$Seedling_SubPlot_ID))
names(SeedlingDataSubplotCount2) <- c("PlotID_Number", "Seedling_Subplot_ID")
SeedlingDataSubplotCount3 <- add_count(SeedlingDataSubplotCount2, (SeedlingDataSubplotCount2$PlotID_Number))
SeedlingDataSubplotCount9 <- subset(SeedlingDataSubplotCount3, ((SeedlingDataSubplotCount3$n != 9)))


TestSeedlingDataSubplotCount9<- if(nrow(SeedlingDataSubplotCount9)>0) {
  addWorksheet(wb, "SeedlingDataSubplotCount9")
  writeData(wb, "SeedlingDataSubplotCount9", SeedlingDataSubplotCount9, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedlingDataSubplotCount9","Event records where there are not 9 seedling subplots")
} 


#Returns seedling records with duplicate tags


SeedDupTag <-DataSeedlings%>%
  select(Unit_Code, PlotID_Number, Start_Date, SeedlingTag)%>%
  group_by(PlotID_Number, SeedlingTag)%>%
  summarize(CountTot = dplyr::n())

SeedDupTag2 <- subset(SeedDupTag, (!is.na(SeedDupTag$SeedlingTag) & (SeedDupTag$CountTot > 1)))


TestSeedDupTag2<- if(nrow(SeedDupTag2)>0) {
  addWorksheet(wb, "SeedDupTag2")
  writeData(wb, "SeedDupTag2", SeedDupTag2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("SeedDupTag2","Seedling records with duplicate tags")
} 


###Tree Validation

#Returns tree records that are recorded as Dead but don't have a Death Cause- if records exist, write to xlsx 
TreeDataDeathCause <- subset(DataTrees, ((DataTrees$Tree_Status == "D" & is.na(DataTrees$StatusDead_Cause))))
TreeDataDeathCause <-TreeDataDeathCause%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID,  Clump_Number, Stem_Letter, Species_Code, Tree_Status, StatusDead_Cause, TreeData_Notes)

TestTreeDataDeathCause<- if(nrow(TreeDataDeathCause)>0) {
  addWorksheet(wb, "TreeDataDeathCause")
  writeData(wb, "TreeDataDeathCause", TreeDataDeathCause, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataDeathCause","Tree records that are recorded as Dead but don't have a Death Cause")
}  

#Returns tree records with duplicate tags
TreeDupTag <-DataTrees%>%
  select(Unit_Code, PlotID_Number, Start_Date, TreeID_Number)%>%
  group_by(PlotID_Number, TreeID_Number) %>%
  summarise(CountTreeTot = dplyr::n())
TreeDupTag2 <- subset(TreeDupTag, ((TreeDupTag$CountTreeTot > 1)))

TestTreeDupTag2<- if(nrow(TreeDupTag2)>0) {
  addWorksheet(wb, "TreeDupTag2")
  writeData(wb, "TreeDupTag2", TreeDupTag2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDupTag2","Tree records with duplicate tags")
}  


#Returns tree records that have a height that should be checked- if records exist, write to xlsx 
TreeDataHeight <- subset(DataTrees, ((DataTrees$TreeHeight_m > 50 & DataTrees$TreeHeight_m != 999)))
TreeDataHeight <-TreeDataHeight%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, StatusDead_Cause, TreeData_Notes)

TestTreeDataHeight<- if(nrow(TreeDataHeight)>0) {
  addWorksheet(wb, "TreeDataHeight")
  writeData(wb, "TreeDataHeight", TreeDataHeight, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataHeight","Tree records that have a height that should be checked")
}  


#Returns tree records where Status = RD and Species = PIAL and Mortality Year is not populated- if records exist, write to xlsx 
TreeMortYear <- subset(DataTrees, ((DataTrees$Tree_Status == "RD" & DataTrees$Species_Code == "PIAL" & is.na(DataTrees$Mort_Year))))
TreeMortYear <-TreeMortYear%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, Mort_Year, TreeHeight_m, Tree_Status, StatusDead_Cause, TreeData_Notes)

TestTreeMortYear<- if(nrow(TreeMortYear)>0) {
  addWorksheet(wb, "TreeMortYear")
  writeData(wb, "TreeMortYear", TreeMortYear, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeMortYear","Tree records where Status = RD and Species = PIAL and Mortality Year is not populated")
}  


#Returns tree records where cones = true and cone count is not populated- if records exist, write to xlsx 
TreeDataCones1 <- subset(DataTrees, ((DataTrees$FemaleCones_YN == 1 & is.na(DataTrees$Cone_Count))))
TreeDataCones1 <-TreeDataCones1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, Event_ID, Clump_Number, Stem_Letter, Species_Code, Mort_Year, TreeHeight_m, Tree_Status, StatusDead_Cause, FemaleCones_YN, Cone_Count, TreeData_Notes)

TestTreeDataCones1<- if(nrow(TreeDataCones1)>0) {
  addWorksheet(wb, "TreeDataCones1")
  writeData(wb, "TreeDataCones1", TreeDataCones1, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataCones1","Tree records where cones = true and cone count is not populated")
} 


#Returns tree records where cones = false and cone count is populated- if records exist, write to xlsx 
TreeDataCones2 <- subset(DataTrees, ((DataTrees$FemaleCones_YN == 0 & !is.na(DataTrees$Cone_Count))))
TreeDataCones2 <-TreeDataCones2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, Mort_Year, TreeHeight_m, Tree_Status, StatusDead_Cause, FemaleCones_YN, Cone_Count, TreeData_Notes)

TestTreeDataCones2<- if(nrow(TreeDataCones2)>0) {
  addWorksheet(wb, "TreeDataCones2")
  writeData(wb, "TreeDataCones2", TreeDataCones2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataCones2","Tree records where cones = false and cone count is populated")
}  


#Returns tree records where crown health doesn't equal domain values (1-5)- if records exist, write to xlsx 
TreeDataCrownHealth <- subset(DataTrees, ((DataTrees$Crown_Health < 1 | DataTrees$Crown_Health > 5)))
TreeDataCrownHealth <-TreeDataCrownHealth%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, Mort_Year, TreeHeight_m, Tree_Status, StatusDead_Cause, Crown_Health, TreeData_Notes)


TestTreeDataCrownHealth<- if(nrow(TreeDataCrownHealth)>0) {
  addWorksheet(wb, "TreeDataCrownHealth")
  writeData(wb, "TreeDataCrownHealth", TreeDataCrownHealth, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataCrownHealth","Tree records where crown health doesn't equal domain values (1-5)")
}  



#Returns tree records where crown kill lower is greater than 100%- if records exist, write to xlsx 
TreeDataCrownKillLow <- subset(DataTrees, ((DataTrees$CrownKill_Lower_perc >100 )))
TreeDataCrownKillLow <-TreeDataCrownKillLow%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestTreeDataCrownKillLow<- if(nrow(TreeDataCrownKillLow)>0) {
  addWorksheet(wb, "TreeDataCrownKillLow")
  writeData(wb, "TreeDataCrownKillLow", TreeDataCrownKillLow, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataCrownKillLow","Tree records where crown kill lower is greater than 100%")
} 


#Returns tree records where crown kill middle is greater than 100%- if records exist, write to xlsx 
TreeDataCrownKillMid <- subset(DataTrees, ((DataTrees$CrownKill_Mid_perc >100 )))
TreeDataCrownKillMid <-TreeDataCrownKillMid%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestTreeDataCrownKillMid<- if(nrow(TreeDataCrownKillMid)>0) {
  addWorksheet(wb, "TreeDataCrownKillMid")
  writeData(wb, "TreeDataCrownKillMid", TreeDataCrownKillMid, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataCrownKillMid","Tree records where crown kill middle is greater than 100%")
}


#Returns tree records where crown kill upper is greater than 100%- if records exist, write to xlsx 
TreeDataCrownKillUp <- subset(DataTrees, ((DataTrees$CrownKill_Upper_perc >100 )))
TreeDataCrownKillUp <-TreeDataCrownKillUp%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestTreeDataCrownKillUp<- if(nrow(TreeDataCrownKillUp)>0) {
  addWorksheet(wb, "TreeDataCrownKillUp")
  writeData(wb, "TreeDataCrownKillUp", TreeDataCrownKillUp, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataCrownKillUp","Tree records where crown kill upper is greater than 100%")
} 


#Returns tree records where DBH is greater than 200 cm- if records exist, write to xlsx 
TreeDataDBH <- subset(DataTrees, ((DataTrees$TreeDBH_cm >200 & DataTrees$TreeDBH_cm != 999)))
TreeDataDBH <-TreeDataDBH%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataDBH<- if(nrow(TreeDataDBH)>0) {
  addWorksheet(wb, "TreeDataDBH")
  writeData(wb, "TreeDataDBH", TreeDataDBH, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataDBH","Tree records where DBH is greater than 200 cm")
}  

#Returns tree records that are missing tree ID- if records exist, write to xlsx 
TreeDataMissingTagNumber <- subset(DataTrees, ((is.na(DataTrees$TreeID_Number))))
TreeDataMissingTagNumber <-TreeDataMissingTagNumber%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataMissingTagNumber<- if(nrow(TreeDataMissingTagNumber)>0) {
  addWorksheet(wb, "TreeDataMissingTagNumber")
  writeData(wb, "TreeDataMissingTagNumber", TreeDataMissingTagNumber, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataMissingTagNumber","Tree records that are missing tree ID")
} 


#Returns tree records that are missing height- if records exist, write to xlsx 
TreeDataMissingHt <- subset(DataTrees, ((is.na(DataTrees$TreeHeight_m))))
TreeDataMissingHt <-TreeDataMissingHt%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataMissingHt<- if(nrow(TreeDataMissingHt)>0) {
  addWorksheet(wb, "TreeDataMissingHt")
  writeData(wb, "TreeDataMissingHt", TreeDataMissingHt, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataMissingHt","Tree records that are missing height")
}  


#Returns tree records that are missing DBH- if records exist, write to xlsx 
TreeDataMissingDBH <- subset(DataTrees, ((is.na(DataTrees$TreeDBH_cm))))
TreeDataMissingDBH <-TreeDataMissingDBH%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataMissingDBH<- if(nrow(TreeDataMissingDBH)>0) {
  addWorksheet(wb, "TreeDataMissingDBH")
  writeData(wb, "TreeDataMissingDBH", TreeDataMissingDBH, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataMissingDBH","Tree records that are missing DBH")
}  


#Returns tree records that are missing species code- if records exist, write to xlsx 
TreeDataMissingSp <- subset(DataTrees, ((is.na(DataTrees$Species_Code))))
TreeDataMissingSp <-TreeDataMissingSp%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataMissingSp<- if(nrow(TreeDataMissingSp)>0) {
  addWorksheet(wb, "TreeDataMissingSp")
  writeData(wb, "TreeDataMissingSp", TreeDataMissingSp, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataMissingSp","Tree records that are missing species code")
}   


#Returns tree records that are missing subplot #- if records exist, write to xlsx 
TreeDataMissingSubplot <- subset(DataTrees, ((is.na(DataTrees$TreeData_SubPlot_StripID))))
TreeDataMissingSubplot <-TreeDataMissingSubplot%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataMissingSubplot<- if(nrow(TreeDataMissingSubplot)>0) {
  addWorksheet(wb, "TreeDataMissingSubplot")
  writeData(wb, "TreeDataMissingSubplot", TreeDataMissingSubplot, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataMissingSubplot","Tree records that are missing subplot #")
}  


#Returns tree records that are missing status- if records exist, write to xlsx 
TreeDataMissingStatus <- subset(DataTrees, ((is.na(DataTrees$Tree_Status))))
TreeDataMissingStatus <-TreeDataMissingStatus%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeDataMissingStatus<- if(nrow(TreeDataMissingStatus)>0) {
  addWorksheet(wb, "TreeDataMissingStatus")
  writeData(wb, "TreeDataMissingStatus", TreeDataMissingStatus, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataMissingStatus","Tree records that are missing status")
}


#Returns live PIAL records where crown kill percent is missing from the upper section but not missing from mid or lower sections- if records exist, write to xlsx 
TreeDataMissingCrownKillUp <- subset(DataTrees, (((DataTrees$Tree_Status == "L" & DataTrees$Species_Code == "PIAL") & is.na(DataTrees$CrownKill_Upper_perc) & ( !is.na(DataTrees$CrownKill_Mid_perc) | !is.na(DataTrees$CrownKill_Lower_perc)))))
TreeDataMissingCrownKillUp <-TreeDataMissingCrownKillUp%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestTreeDataMissingCrownKillUp<- if(nrow(TreeDataMissingCrownKillUp)>0) {
  addWorksheet(wb, "TreeDataMissingCrownKillUp")
  writeData(wb, "TreeDataMissingCrownKillUp", TreeDataMissingCrownKillUp, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataMissingCrownKillUp","Live PIAL records where crown kill percent is missing from the upper section but not missing from mid or lower sections")
}


#Returns live PIAL records where crown kill percent is missing from the middle section but not missing from upper or lower sections- if records exist, write to xlsx 
TreeDataMissingCrownKillMid <- subset(DataTrees, (((DataTrees$Tree_Status == "L" & DataTrees$Species_Code == "PIAL") & is.na(DataTrees$CrownKill_Mid_perc)& (!is.na(DataTrees$CrownKill_Upper_perc) | !is.na(DataTrees$CrownKill_Lower_perc)))))
TreeDataMissingCrownKillMid <-TreeDataMissingCrownKillMid%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestTreeDataMissingCrownKillMid<- if(nrow(TreeDataMissingCrownKillMid)>0) {
  addWorksheet(wb, "TreeDataMissingCrownKillMid")
  writeData(wb, "TreeDataMissingCrownKillMid", TreeDataMissingCrownKillMid, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataMissingCrownKillMid","Live PIAL records where crown kill percent is missing from the middle section but not missing from upper or lower sections")
} 


#Returns live PIAL records where crown kill percent is missing from the lower section but not missing from upper or mid sections- if records exist, write to xlsx 
TreeDataMissingCrownKillLow <- subset(DataTrees, (((DataTrees$Tree_Status == "L" & DataTrees$Species_Code == "PIAL") & is.na(DataTrees$CrownKill_Lower_perc)& (!is.na(DataTrees$CrownKill_Upper_perc) | !is.na(DataTrees$CrownKill_Mid_perc)))))
TreeDataMissingCrownKillLow <-TreeDataMissingCrownKillLow%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestTreeDataMissingCrownKillLow<- if(nrow(TreeDataMissingCrownKillLow)>0) {
  addWorksheet(wb, "TreeDataMissingCrownKillLow")
  writeData(wb, "TreeDataMissingCrownKillLow", TreeDataMissingCrownKillLow, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataMissingCrownKillLow","Live PIAL records where crown kill percent is missing from the lower section but not missing from upper or mid sections")
}


#Returns live PIAL records that are missing crown health- if records exist, write to xlsx 
TreeDataCrownHealthPIAL <- subset(DataTrees, (((DataTrees$Tree_Status == "L" & DataTrees$Species_Code == "PIAL") & is.na(DataTrees$Crown_Health))))
TreeDataCrownHealthPIAL <-TreeDataCrownHealthPIAL%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataCrownHealthPIAL<- if(nrow(TreeDataCrownHealthPIAL)>0) {
  addWorksheet(wb, "TreeDataCrownHealthPIAL")
  writeData(wb, "TreeDataCrownHealthPIAL", TreeDataCrownHealthPIAL, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataCrownHealthPIAL","Live PIAL records that are missing crown health")
}


#Returns tree records with a species that doesn't match a species in the lookup table- if records exist, write to xlsx 
names(SpeciesList)[1]<- "Species_Code" #renames field so that I can merge data
names(SpeciesList)[2]<- "Unit_Code" #renames field so that I can merge data
TreesNoSp <- subset(DataTrees, ((!is.na(DataTrees$Species_Code)))) #removes records where no seedling species was recorded
TreesNoSpJoin <- anti_join(TreesNoSp, SpeciesList, by = 'Species_Code', 'Unit_Code') #joins lookup table to data table
TreesNoSpJoin <-TreesNoSpJoin%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreesNoSpJoin<- if(nrow(TreesNoSpJoin)>0) {
  addWorksheet(wb, "TreesNoSpJoin")
  writeData(wb, "TreesNoSpJoin", TreesNoSpJoin, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreesNoSpJoin","Tree records with a species that doesn't match a species in the lookup table")
}


#Returns tree records where status does not equal a domain value (L, RD, D)- if records exist, write to xlsx 
TreeDataStatus <- subset(DataTrees, ((DataTrees$Tree_Status != "L" & DataTrees$Tree_Status != "D" & DataTrees$Tree_Status != "RD")))
TreeDataStatus <-TreeDataStatus%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataStatus<- if(nrow(TreeDataStatus)>0) {
  addWorksheet(wb, "TreeDataStatus")
  writeData(wb, "TreeDataStatus", TreeDataStatus, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataStatus","Tree records where status does not equal a domain value (L, RD, D)")
}


#Returns tree records that have a stripID that is not 1-5- if records exist, write to xlsx 
TreeDataSubplot <- subset(DataTrees, ((DataTrees$TreeData_SubPlot_StripID < 1 | DataTrees$TreeData_SubPlot_StripID > 5)))
TreeDataSubplot <-TreeDataSubplot%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataSubplot<- if(nrow(TreeDataSubplot)>0) {
  addWorksheet(wb, "TreeDataSubplot")
  writeData(wb, "TreeDataSubplot", TreeDataSubplot, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataSubplot","Tree records that have a stripID that is not 1-5")
}


#Returns tree records where stem letter is not a letter- if records exist, write to xlsx 
TreeDataStem <- subset(DataTrees, !is.na(DataTrees$Stem_Letter))
TreeDataStem2 <- subset(TreeDataStem, !grepl("^[[:alpha:]]+$", TreeDataStem$Stem_Letter))
TreeDataStem2 <-TreeDataStem2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataStem2<- if(nrow(TreeDataStem2)>0) {
  addWorksheet(wb, "TreeDataStem2")
  writeData(wb, "TreeDataStem2", TreeDataStem2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataStem2","Tree records where stem letter is not a letter")
}



#Returns tree records where lower bole canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBoleCankLow1 <- subset(DataTrees, (DataTrees$BoleCankers_I_Lower_YN == 1 & is.na(DataTrees$BoleCanks_ITypes_Lower)))
TreeDataBoleCankLow1 <-TreeDataBoleCankLow1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Lower_YN, BoleCanks_ITypes_Lower, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataBoleCankLow1<- if(nrow(TreeDataBoleCankLow1)>0) {
  addWorksheet(wb, "TreeDataBoleCankLow1")
  writeData(wb, "TreeDataBoleCankLow1", TreeDataBoleCankLow1, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBoleCankLow1","Tree records where lower bole canks infestation checkbox = true but infestation type is null")
}


#Returns tree records where lower bole canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBoleCankLow2 <- subset(DataTrees, (DataTrees$BoleCankers_I_Lower_YN == 0 & !is.na(DataTrees$BoleCanks_ITypes_Lower)))
TreeDataBoleCankLow2 <-TreeDataBoleCankLow2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Lower_YN, BoleCanks_ITypes_Lower, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeDataBoleCankLow2<- if(nrow(TreeDataBoleCankLow2)>0) {
  addWorksheet(wb, "TreeDataBoleCankLow2")
  writeData(wb, "TreeDataBoleCankLow2", TreeDataBoleCankLow2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBoleCankLow2","Tree records where lower bole canks infestation checkbox = false but infestation type is populated")
} 


#Returns tree records where middle bole canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBoleCankMid1 <- subset(DataTrees, (DataTrees$BoleCankers_I_Mid_YN == 1 & is.na(DataTrees$BoleCanks_ITypes_Mid)))
TreeDataBoleCankMid1 <-TreeDataBoleCankMid1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Mid_YN, BoleCanks_ITypes_Mid, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataBoleCankMid1<- if(nrow(TreeDataBoleCankMid1)>0) {
  addWorksheet(wb, "TreeDataBoleCankMid1")
  writeData(wb, "TreeDataBoleCankMid1", TreeDataBoleCankMid1, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBoleCankMid1","Tree records where middle bole canks infestation checkbox = true but infestation type is null")
}  


#Returns tree records where middle bole canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBoleCankMid2 <- subset(DataTrees, (DataTrees$BoleCankers_I_Mid_YN == 0 & !is.na(DataTrees$BoleCanks_ITypes_Mid)))
TreeDataBoleCankMid2 <-TreeDataBoleCankMid2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Mid_YN, BoleCanks_ITypes_Mid, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataBoleCankMid2<- if(nrow(TreeDataBoleCankMid2)>0) {
  addWorksheet(wb, "TreeDataBoleCankMid2")
  writeData(wb, "TreeDataBoleCankMid2", TreeDataBoleCankMid2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBoleCankMid2","Tree records where middle bole canks infestation checkbox = false but infestation type is populated")
} 


#Returns tree records where upper bole canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBoleCankUpp1 <- subset(DataTrees, (DataTrees$BoleCankers_I_Upper_YN == 1 & is.na(DataTrees$BoleCanks_ITypes_Upper)))
TreeDataBoleCankUpp1 <-TreeDataBoleCankUpp1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Upper_YN, BoleCanks_ITypes_Upper, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataBoleCankUpp1<- if(nrow(TreeDataBoleCankUpp1)>0) {
  addWorksheet(wb, "TreeDataBoleCankUpp1")
  writeData(wb, "TreeDataBoleCankUpp1", TreeDataBoleCankUpp1, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBoleCankUpp1","Tree records where upper bole canks infestation checkbox = true but infestation type is null")
}  


#Returns tree records where upper bole canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBoleCankUpp2 <- subset(DataTrees, (DataTrees$BoleCankers_I_Upper_YN == 0 & !is.na(DataTrees$BoleCanks_ITypes_Upper)))
TreeDataBoleCankUpp2 <-TreeDataBoleCankUpp2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Upper_YN, BoleCanks_ITypes_Upper, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataBoleCankUpp2<- if(nrow(TreeDataBoleCankUpp2)>0) {
  addWorksheet(wb, "TreeDataBoleCankUpp2")
  writeData(wb, "TreeDataBoleCankUpp2", TreeDataBoleCankUpp2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBoleCankUpp2","Tree records where upper bole canks infestation checkbox = false but infestation type is populated")
} 


#Returns tree records where lower branch canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBranchCankLow1 <- subset(DataTrees, (DataTrees$BranchCanks_I_Lower_YN == 1 & is.na(DataTrees$BranchCanks_ITypes_Lower)))
TreeDataBranchCankLow1 <-TreeDataBranchCankLow1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Lower_YN, BranchCanks_ITypes_Lower, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataBranchCankLow1<- if(nrow(TreeDataBranchCankLow1)>0) {
  addWorksheet(wb, "TreeDataBranchCankLow1")
  writeData(wb, "TreeDataBranchCankLow1", TreeDataBranchCankLow1, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBranchCankLow1","Tree records where lower branch canks infestation checkbox = true but infestation type is null")
} 


#Returns tree records where lower branch canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBranchCankLow2 <- subset(DataTrees, (DataTrees$BranchCanks_I_Lower_YN == 0 & !is.na(DataTrees$BranchCanks_ITypes_Lower)))
TreeDataBranchCankLow2 <-TreeDataBranchCankLow2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Lower_YN, BranchCanks_ITypes_Lower, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataBranchCankLow2<- if(nrow(TreeDataBranchCankLow2)>0) {
  addWorksheet(wb, "TreeDataBranchCankLow2")
  writeData(wb, "TreeDataBranchCankLow2", TreeDataBranchCankLow2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBranchCankLow2","Tree records where lower branch canks infestation checkbox = false but infestation type is populated")
} 


#Returns tree records where middle branch canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBranchCankMid1 <- subset(DataTrees, (DataTrees$BranchCanks_I_Mid_YN == 1 & is.na(DataTrees$BranchCanks_ITypes_Mid)))
TreeDataBranchCankMid1 <-TreeDataBranchCankMid1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Mid_YN, BranchCanks_ITypes_Mid, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataBranchCankMid1<- if(nrow(TreeDataBranchCankMid1)>0) {
  addWorksheet(wb, "TreeDataBranchCankMid1")
  writeData(wb, "TreeDataBranchCankMid1", TreeDataBranchCankMid1, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBranchCankMid1","Tree records where middle branch canks infestation checkbox = true but infestation type is null")
} 


#Returns tree records where middle branch canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBranchCankMid2 <- subset(DataTrees, (DataTrees$BranchCanks_I_Mid_YN == 0 & !is.na(DataTrees$BranchCanks_ITypes_Mid)))
TreeDataBranchCankMid2 <-TreeDataBranchCankMid2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Mid_YN, BranchCanks_ITypes_Mid, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDataBranchCankMid2<- if(nrow(TreeDataBranchCankMid2)>0) {
  addWorksheet(wb, "TreeDataBranchCankMid2")
  writeData(wb, "TreeDataBranchCankMid2", TreeDataBranchCankMid2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBranchCankMid2","Tree records where middle branch canks infestation checkbox = false but infestation type is populated")
} 



#Returns tree records where upper branch canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBranchUppCank1 <- subset(DataTrees, (DataTrees$BranchCanks_I_Upper_YN == 1 & is.na(DataTrees$BranchCanks_ITypes_Upper)))
TreeDataBranchUppCank1 <-TreeDataBranchUppCank1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Upper_YN, BranchCanks_ITypes_Upper, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeDataBranchUppCank1<- if(nrow(TreeDataBranchUppCank1)>0) {
  addWorksheet(wb, "TreeDataBranchUppCank1")
  writeData(wb, "TreeDataBranchUppCank1", TreeDataBranchUppCank1, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBranchUppCank1","Tree records where upper branch canks infestation checkbox = true but infestation type is null")
} 


#Returns tree records where upper branch canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBranchCankUpp2 <- subset(DataTrees, (DataTrees$BranchCanks_I_Upper_YN == 0 & !is.na(DataTrees$BranchCanks_ITypes_Upper)))
TreeDataBranchCankUpp2 <-TreeDataBranchCankUpp2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Upper_YN, BranchCanks_ITypes_Upper, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeDataBranchCankUpp2<- if(nrow(TreeDataBranchCankUpp2)>0) {
  addWorksheet(wb, "TreeDataBranchCankUpp2")
  writeData(wb, "TreeDataBranchCankUpp2", TreeDataBranchCankUpp2, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("TreeDataBranchCankUpp2","Tree records where upper branch canks infestation checkbox = false but infestation type is populated")
}


#Returns tree records where Branch Canks upper field values don't match domain values
InfestBranchCanksUp1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BranchCanks_ITypes_Upper)
InfestBranchCanksUp2 <- subset(InfestBranchCanksUp1, (!is.na(DataTrees$BranchCanks_ITypes_Upper)))
InfestBranchCanksUp3 <- InfestBranchCanksUp2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBranchCanksUp2$BranchCanks_ITypes_Upper),]

TestInfestBranchCanksUp3<- if(nrow(InfestBranchCanksUp3)>0) {
  addWorksheet(wb, "InfestBranchCanksUp3")
  writeData(wb, "InfestBranchCanksUp3", InfestBranchCanksUp3, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("InfestBranchCanksUp3","Tree records where Branch Canks upper field values don't match domain values")
}

#Returns tree records where Branch Canks middle field values don't match domain values
InfestBranchCanksMid1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BranchCanks_ITypes_Mid)
InfestBranchCanksMid2 <- subset(InfestBranchCanksMid1, (!is.na(DataTrees$BranchCanks_ITypes_Mid)))
InfestBranchCanksMid3 <- InfestBranchCanksMid2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBranchCanksMid2$BranchCanks_ITypes_Mid),]

TestInfestBranchCanksMid3<- if(nrow(InfestBranchCanksMid3)>0) {
  addWorksheet(wb, "InfestBranchCanksMid3")
  writeData(wb, "InfestBranchCanksMid3", InfestBranchCanksMid3, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("InfestBranchCanksMid3","Tree records where Branch Canks middle field values don't match domain values")
}

#Returns tree records where Branch Canks lower field values don't match domain values
InfestBranchCanksLow1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BranchCanks_ITypes_Lower)
InfestBranchCanksLow2 <- subset(InfestBranchCanksLow1, (!is.na(DataTrees$BranchCanks_ITypes_Lower)))
InfestBranchCanksLow3 <- InfestBranchCanksLow2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBranchCanksLow2$BranchCanks_ITypes_Lower),]

TestInfestBranchCanksLow3<- if(nrow(InfestBranchCanksLow3)>0) {
  addWorksheet(wb, "InfestBranchCanksLow3")
  writeData(wb, "InfestBranchCanksLow3", InfestBranchCanksLow3, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("InfestBranchCanksLow3","Tree records where Branch Canks lower field values don't match domain values")
}


#Returns tree records where Bole Canks upper field values don't match domain values
InfestBoleCanksUp1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BoleCanks_ITypes_Upper)
InfestBoleCanksUp2 <- subset(InfestBoleCanksUp1, (!is.na(DataTrees$BoleCanks_ITypes_Upper)))
InfestBoleCanksUp3 <- InfestBoleCanksUp2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBoleCanksUp2$BoleCanks_ITypes_Upper),]

TestInfestBoleCanksUp3<- if(nrow(InfestBoleCanksUp3)>0) {
  addWorksheet(wb, "InfestBoleCanksUp3")
  writeData(wb, "InfestBoleCanksUp3", InfestBoleCanksUp3, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("InfestBoleCanksUp3","Tree records where Bole Canks upper field values don't match domain values")
}


#Returns tree records where Bole Canks middle field values don't match domain values
InfestBoleCanksMid1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BoleCanks_ITypes_Mid)
InfestBoleCanksMid2 <- subset(InfestBoleCanksMid1, (!is.na(DataTrees$BoleCanks_ITypes_Mid)))
InfestBoleCanksMid3 <- InfestBoleCanksMid2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBoleCanksMid2$BoleCanks_ITypes_Mid),]

TestInfestBoleCanksMid3<- if(nrow(InfestBoleCanksMid3)>0) {
  addWorksheet(wb, "InfestBoleCanksMid3")
  writeData(wb, "InfestBoleCanksMid3", InfestBoleCanksMid3, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("InfestBoleCanksMid3","Tree records where Bole Canks middle field values don't match domain values")
}

#Returns tree records where Bole Canks lower field values don't match domain values
InfestBoleCanksLow1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BoleCanks_ITypes_Lower)
InfestBoleCanksLow2 <- subset(InfestBoleCanksLow1, (!is.na(DataTrees$BoleCanks_ITypes_Lower)))
InfestBoleCanksLow3 <- InfestBoleCanksLow2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBoleCanksLow2$BoleCanks_ITypes_Lower),]

TestInfestBoleCanksLow3<- if(nrow(InfestBoleCanksLow3)>0) {
  addWorksheet(wb, "InfestBoleCanksLow3")
  writeData(wb, "InfestBoleCanksLow3", InfestBoleCanksLow3, rowNames = FALSE)
  saveWorkbook(wb, OutputFilename, overwrite = TRUE)
  TableDefs[nrow(TableDefs) + 1,] = list("InfestBoleCanksLow3","Tree records where Bole Canks lower field values don't match domain values")
}



#### write TAbleDefs back to the Excel sheet
addWorksheet(wb, "TableDefs")
writeData(wb, "TableDefs", TableDefs, rowNames = FALSE)
saveWorkbook(wb, OutputFilename, overwrite = TRUE)

#############CLOSE THE CONNECTION TO THE DATABASE#####################################################################################
odbcCloseAll()

