

rm(list=ls()) # start with a clean slate

setwd("C:/Users/snydera/Desktop/R_Validation") 


library("RODBC")
library("lubridate")
library("dplyr")
library("tidyr")
library(xlsx)
library(tidyverse)
library("stringr")
library("distr")

options(stringsAsFactors = FALSE) 

##Connect to the correct Whitebark Pine database
connection <- odbcConnectAccess2007("Master_WBP_20180906_0909_FinalFieldDB_20180912_1409.mdb")

#Import Location Data
Locations <- sqlFetch(connection,"tbl_Locations")

#Import Species Lookup Table
SpeciesList <- sqlFetch(connection,"tlu_Species_Parks")

#Import Events Data
Events<- sqlQuery(connection, "SELECT tbl_Locations.PlotID_Number, tbl_Events.Start_Date, tbl_Events.Event_ID, tbl_Events.Location_ID
FROM tbl_Locations INNER JOIN tbl_Events ON tbl_Locations.Location_ID = tbl_Events.Location_ID")

#Start XLS file to hold data errors
write.xlsx(Events, file= "WBP_Validation.xlsx", sheetName = "EventsList_NoErrors", row.names = FALSE, showNA = FALSE)

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


TestEventNoTrees<- if(nrow(EventsNoTrees)>0) {
write.xlsx(TestEventNoTrees, file= "WBP_Validation.xlsx", sheetName = "EventsNoTrees", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Looking for events with no photo records- if records exist, write to xlsx
EventsNoPhotos <- merge(Events, DataPhotos, by = 'Event_ID', all = TRUE)
EventsNoPhotos <- subset(EventsNoPhotos, (is.na(EventsNoPhotos$PlotID_Number.y) & !is.na(EventsNoPhotos$PlotID_Number.x)))
EventsNoPhotos <-EventsNoPhotos%>%
  select(PlotID_Number.x, Event_ID, Start_Date.x, Unit_Code)

TestEventNoPhotos<- if(nrow(EventsNoPhotos)>0) {
  write.xlsx(EventsNoPhotos, file= "WBP_Validation.xlsx", sheetName = "EventsNoPhotos", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Looking for events with no seedling records- if records exist, write to xlsx
EventsNoSeedlings <- merge(Events, DataSeedlings, by = 'Event_ID', all = TRUE)
EventsNoSeedlings <- subset(EventsNoSeedlings, (is.na(EventsNoSeedlings$PlotID_Number.y) & !is.na(EventsNoSeedlings$PlotID_Number.x)))
EventsNoSeedlings <-EventsNoSeedlings%>%
  select(PlotID_Number.x, Event_ID, Start_Date.x, Unit_Code)

TestEventsNoSeedlings<- if(nrow(EventsNoSeedlings)>0) {
  write.xlsx(EventsNoSeedlings, file= "WBP_Validation.xlsx", sheetName = "EventsNoSeedlings", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

###Photo Validation

#Looking for photo records with bearing greater than 360- if records exist, write to xlsx
PhotosBearing <- subset(DataPhotos, ((DataPhotos$PlotPhoto_Bear_deg > 360)))
PhotosBearing <-PhotosBearing%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotosBearing<- if(nrow(PhotosBearing)>0) {
  write.xlsx(PhotosBearing, file= "WBP_Validation.xlsx", sheetName = "PhotosBearingOver360", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Looking for events where there are less than 4 photos- if records exist, write to xlsx
PhotoClean <- DataPhotos[,c('Unit_Code', 'Start_Date', 'PlotID_Number')]
PhotoCount2 <- add_count(PhotoClean, (PhotoClean$PlotID_Number))
PhotoCount3 <- distinct(PhotoCount2, (PhotoCount2$PlotID_Number), (PhotoCount2$n))
names(PhotoCount3) <- c("PlotID_Number", "n")
PhotoCount4 <- subset(PhotoCount3, ((PhotoCount3$n < 4)))

TestPhotoCount<- if(nrow(PhotoCount4)>0) {
  write.xlsx(PhotoCount4, file= "WBP_Validation.xlsx", sheetName = "PhotosCountLT4", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Returns photo records that don't have a domain value for PlotPhoto_Loc_Ref- if records exist, write to xlsx
PhotosLocRef <- subset(DataPhotos, ((DataPhotos$PlotPhoto_Loc_Ref != "SW_Corner" & DataPhotos$PlotPhoto_Loc_Ref != "NW_Corner" & DataPhotos$PlotPhoto_Loc_Ref != "NE_Corner" & DataPhotos$PlotPhoto_Loc_Ref != "SE_Corner" & DataPhotos$PlotPhoto_Loc_Ref != "See_Notes")))

PhotosLocRef <-PhotosLocRef%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotoLocRef<- if(nrow(PhotosLocRef)>0) {
  write.xlsx(PhotosLocRef, file= "WBP_Validation.xlsx", sheetName = "PhotosNoLocRef", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Returns photo records that are missing photo number - if records exist, write to xlsx
PhotosDataMissingNumber <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_Number)))
PhotosDataMissingNumber <-PhotosDataMissingNumber%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotoMissingNum<- if(nrow(PhotosDataMissingNumber)>0) {
  write.xlsx(PhotosDataMissingNumber, file= "WBP_Validation.xlsx", sheetName = "PhotosMissingNum", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns photo records that are missing file name- if records exist, write to xlsx 
PhotosDataMissingFileName <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_File_Name)))
PhotosDataMissingFileName <-PhotosDataMissingFileName%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotoMissingFileName<- if(nrow(PhotosDataMissingFileName)>0) {
  write.xlsx(PhotosDataMissingFileName, file= "WBP_Validation.xlsx", sheetName = "PhotosMissingFileName", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns photo records that are missing file path- if records exist, write to xlsx
PhotosDataMissingFilePath <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_File_Path)))
PhotosDataMissingFilePath <-PhotosDataMissingFilePath%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotoMissingPath<- if(nrow(PhotosDataMissingFilePath)>0) {
  write.xlsx(PhotosDataMissingFilePath, file= "WBP_Validation.xlsx", sheetName = "PhotosMissingFilePath", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns photo records that are missing location reference- if records exist, write to xlsx
PhotosDataMissingLocRef <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_Loc_Ref)))
PhotosDataMissingLocRef <-PhotosDataMissingLocRef%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotoMissingLocRef<- if(nrow(PhotosDataMissingLocRef)>0) {
  write.xlsx(PhotosDataMissingLocRef, file= "WBP_Validation.xlsx", sheetName = "PhotosMissingLoCRef", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns photo records that are missing bearing- if records exist, write to xlsx 
PhotosDataMissingBearing <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_Bear_deg)))
PhotosDataMissingBearing <-PhotosDataMissingBearing%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)


TestPhotoMissingBearing<- if(nrow(PhotosDataMissingBearing)>0) {
  write.xlsx(PhotosDataMissingBearing, file= "WBP_Validation.xlsx", sheetName = "PhotosMissingBearing", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns photo records that are missing image ID- if records exist, write to xlsx 
PhotosDataMissingImageID <- subset(DataPhotos, (is.na(DataPhotos$Camera_ImageID)))
PhotosDataMissingImageID <-PhotosDataMissingImageID%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotoMissingImageID<- if(nrow(PhotosDataMissingImageID)>0) {
  write.xlsx(PhotosDataMissingImageID, file= "WBP_Validation.xlsx", sheetName = "PhotosMissingImageID", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns photo records that are missing photo date- if records exist, write to xlsx  
PhotosDataMissingDate <- subset(DataPhotos, (is.na(DataPhotos$PlotPhoto_Date)))
PhotosDataMissingDate <-PhotosDataMissingDate%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, PlotPhoto_ID, PlotPhoto_File_Name, PlotPhoto_Number, PlotPhoto_File_Path, PlotPhoto_Loc_Ref, Camera_ImageID, PlotPhoto_Bear_deg)

TestPhotoMissingDate<- if(nrow(PhotosDataMissingDate)>0) {
  write.xlsx(PhotosDataMissingDate, file= "WBP_Validation.xlsx", sheetName = "PhotosMissingDate", append = TRUE, row.names = FALSE, showNA = FALSE)
}

###Seedling Validation

#Returns seedling records that are recorded as Dead but don't have a Death Cause- if records exist, write to xlsx 
SeedlingDataDeathCause <- subset(DataSeedlings, ((DataSeedlings$Status == "D" & is.na(DataSeedlings$Death_Cause))))
SeedlingDataDeathCause <-SeedlingDataDeathCause%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedDeathCause<- if(nrow(SeedlingDataDeathCause)>0) {
  write.xlsx(SeedlingDataDeathCause, file= "WBP_Validation.xlsx", sheetName = "SeedlingDeathCause", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns seedlings records that don't have a height domain value- if records exist, write to xlsx 
SeedlingDataHeight <- subset(DataSeedlings, ((DataSeedlings$Height_Class != "20 - <50 cm" & DataSeedlings$Height_Class != "50 - <100 cm" & DataSeedlings$Height_Class != "100 - <137 cm")))
SeedlingDataHeight <-SeedlingDataHeight%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedDataHt<- if(nrow(SeedlingDataHeight)>0) {
  write.xlsx(SeedlingDataHeight, file= "WBP_Validation.xlsx", sheetName = "SeedlingDataHt", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns seedlings records that are missing data in subplot ID- if records exist, write to xlsx 
SeedlingDataMissingSubplot <- subset(DataSeedlings, ((is.na(DataSeedlings$Seedling_SubPlot_ID))))
SeedlingDataMissingSubplot <-SeedlingDataMissingSubplot%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedDataMissingSub<- if(nrow(SeedlingDataMissingSubplot)>0) {
  write.xlsx(SeedlingDataMissingSubplot, file= "WBP_Validation.xlsx", sheetName = "SeedlingMissingSubplot", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns seedlings records that are missing data in species- if records exist, write to xlsx 
SeedlingDataMissingSpCode <- subset(DataSeedlings, ((is.na(DataSeedlings$Species_Code))))
SeedlingDataMissingSpCode <-SeedlingDataMissingSpCode%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedMissingSp<- if(nrow(SeedlingDataMissingSpCode)>0) {
  write.xlsx(SeedlingDataMissingSpCode, file= "WBP_Validation.xlsx", sheetName = "SeedlingMissingSp", append = TRUE, row.names = FALSE, showNA = FALSE)
}



#Returns seedlings records where status = "L" and height is missing- if records exist, write to xlsx 
SeedlingDataMissingHt <- subset(DataSeedlings, (((DataSeedlings$Status == "L" | DataSeedlings$Status == "RD") & is.na(DataSeedlings$Height_Class))))
SeedlingDataMissingHt <-SeedlingDataMissingHt%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedMissingHt<- if(nrow(SeedlingDataMissingHt)>0) {
  write.xlsx(SeedlingDataMissingHt, file= "WBP_Validation.xlsx", sheetName = "SeedlingMissingHt", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns seedlings records that are missing a tag number- if records exist, write to xlsx 
SeedlingDataMissingTag <- subset(DataSeedlings, (((is.na(DataSeedlings$Species_Code) | DataSeedlings$Species_Code != "_NONE" & DataSeedlings$Species_Code != "_NotSampled") & is.na(DataSeedlings$SeedlingTag))))
SeedlingDataMissingTag <-SeedlingDataMissingTag%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedMissingTag<- if(nrow(SeedlingDataMissingTag)>0) {
  write.xlsx(SeedlingDataMissingTag, file= "WBP_Validation.xlsx", sheetName = "SeedlingMissingTag", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns seedlings records that are missing status- if records exist, write to xlsx 
SeedlingDataMissingStatus<- subset(DataSeedlings, (((is.na(DataSeedlings$Species_Code) | DataSeedlings$Species_Code != "_NONE" & DataSeedlings$Species_Code != "_NotSampled") & is.na(DataSeedlings$Status))))
SeedlingDataMissingStatus <-SeedlingDataMissingStatus%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedMissingStatus<- if(nrow(SeedlingDataMissingStatus)>0) {
  write.xlsx(SeedlingDataMissingStatus, file= "WBP_Validation.xlsx", sheetName = "SeedlingMissingStatus", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns seedlings records with a species that doesn't match a species in the lookup table- if records exist, write to xlsx 
names(SpeciesList)[1]<- "Species_Code" #renames field so that I can merge data
names(SpeciesList)[2]<- "Unit_Code" #renames field so that I can merge data
SeedlingNoSp <- subset(DataSeedlings, ((!is.na(DataSeedlings$Species_Code)))) #removes records where no seedling species was recorded
SeedlingNoSpJoin <- anti_join(SeedlingNoSp, SpeciesList, by = 'Species_Code', 'Unit_Code') #joins lookup table to data table
SeedlingNoSpJoin <-SeedlingNoSpJoin%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedMissingSPinLU<- if(nrow(SeedlingNoSpJoin)>0) {
  write.xlsx(SeedlingNoSpJoin, file= "WBP_Validation.xlsx", sheetName = "SeedlingNoSpInLU_Table", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns seedlings records where status doesn't equal a domain value (L, RD, D)- if records exist, write to xlsx 
SeedlingDataStatus <- subset(DataSeedlings, ((DataSeedlings$Status != "L" & DataSeedlings$Status != "D" & DataSeedlings$Status != "RD")))
SeedlingDataStatus <-SeedlingDataStatus%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)


TestSeedStatusDomain<- if(nrow(SeedlingDataStatus)>0) {
  write.xlsx(SeedlingDataStatus, file= "WBP_Validation.xlsx", sheetName = "SeedlingStatusDomain", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns seedlings records that have a subplot number that is not between 1 and 9- if records exist, write to xlsx 
SeedlingDataSubplot <- subset(DataSeedlings, ((DataSeedlings$Seedling_SubPlot_ID < 1 | DataSeedlings$Seedling_SubPlot_ID > 9)))
SeedlingDataSubplot <-SeedlingDataSubplot%>%
  select(PlotID_Number, Event_ID, Start_Date, Unit_Code, Seedling_SubPlot_ID, Species_Code, Height_Class, SeedlingTag, Status, Death_Cause, Seedling_SubPlot_Notes)

TestSeedSubplot<- if(nrow(SeedlingDataSubplot)>0) {
  write.xlsx(SeedlingDataSubplot, file= "WBP_Validation.xlsx", sheetName = "SeedlingSubplot_Not_1-9", append = TRUE, row.names = FALSE, showNA = FALSE)
}



#Returns event records where there are not 9 seedling subplots- if records exist, write to xlsx 
DataSeedlingsClean <- DataSeedlings[,c('Unit_Code', 'Start_Date', 'PlotID_Number','Seedling_SubPlot_ID')]
SeedlingDataSubplotCount2 <- distinct(DataSeedlingsClean, (DataSeedlingsClean$PlotID_Number), (DataSeedlingsClean$Seedling_SubPlot_ID))
names(SeedlingDataSubplotCount2) <- c("PlotID_Number", "Seedling_Subplot_ID")
SeedlingDataSubplotCount3 <- add_count(SeedlingDataSubplotCount2, (SeedlingDataSubplotCount2$PlotID_Number))
SeedlingDataSubplotCount4 <- subset(SeedlingDataSubplotCount3, ((SeedlingDataSubplotCount3$n != 9)))


TestSeedSubplotNotNine<- if(nrow(SeedlingDataSubplotCount4)>0) {
  write.xlsx(as.data.frame(SeedlingDataSubplotCount4), file= "WBP_Validation.xlsx", sheetName = "SeedlingsNot9Subplots", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns seedling records with duplicate tags
SeedDupTag <-DataSeedlings%>%
  select(Unit_Code, PlotID_Number, Start_Date, SeedlingTag)%>%
  group_by(PlotID_Number, SeedlingTag) %>%
  summarise(CountTot = n())
SeedDupTag2 <- subset(SeedDupTag, (!is.na(SeedDupTag$SeedlingTag) & (SeedDupTag$CountTot > 1)))

TestSeedDupTag<- if(nrow(SeedDupTag2)>0) {
  write.xlsx(as.data.frame(SeedDupTag2), file= "WBP_Validation.xlsx", sheetName = "SeedlingsDupTag", append = TRUE, row.names = FALSE, showNA = FALSE)
}


###Tree Validation

#Returns tree records that are recorded as Dead but don't have a Death Cause- if records exist, write to xlsx 
TreeDataDeathCause <- subset(DataTrees, ((DataTrees$Tree_Status == "D" & is.na(DataTrees$StatusDead_Cause))))
TreeDataDeathCause <-TreeDataDeathCause%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID,  Clump_Number, Stem_Letter, Species_Code, Tree_Status, StatusDead_Cause, TreeData_Notes)

TestTreeDataDeathCause<- if(nrow(TreeDataDeathCause)>0) {
  write.xlsx(TreeDataDeathCause, file= "WBP_Validation.xlsx", sheetName = "TreeDeadNoCause", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Returns tree records with duplicate tags
TreeDupTag <-DataTrees%>%
  select(Unit_Code, PlotID_Number, Start_Date, TreeID_Number)%>%
  group_by(PlotID_Number, TreeID_Number) %>%
  summarise(CountTreeTot = n())
TreeDupTag2 <- subset(TreeDupTag, ((TreeDupTag$CountTreeTot > 1)))

TestTreeDupTag<- if(nrow(TreeDupTag2)>0) {
  write.xlsx(as.data.frame(TreeDupTag2), file= "WBP_Validation.xlsx", sheetName = "TreeDupTag", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Returns tree records that have a height that should be checked- if records exist, write to xlsx 
TreeDataHeight <- subset(DataTrees, ((DataTrees$TreeHeight_m > 50 & DataTrees$TreeHeight_m != 999)))
TreeDataHeight <-TreeDataHeight%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, StatusDead_Cause, TreeData_Notes)

TestTreeHt<- if(nrow(TreeDataHeight)>0) {
  write.xlsx(TreeDataHeight, file= "WBP_Validation.xlsx", sheetName = "TreeHeight", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Returns tree records where Status = RD and Species = PIAL and Mortality Year is not populated- if records exist, write to xlsx 
TreeMortYear <- subset(DataTrees, ((DataTrees$Tree_Status == "RD" & DataTrees$Species_Code == "PIAL" & is.na(DataTrees$Mort_Year))))
TreeMortYear <-TreeMortYear%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, Mort_Year, TreeHeight_m, Tree_Status, StatusDead_Cause, TreeData_Notes)

TestTreeMortYear<- if(nrow(TreeMortYear)>0) {
  write.xlsx(TreeMortYear, file= "WBP_Validation.xlsx", sheetName = "TreeNoMortYear", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where cones = true and cone count is not populated- if records exist, write to xlsx 
TreeDataCones1 <- subset(DataTrees, ((DataTrees$FemaleCones_YN == 1 & is.na(DataTrees$Cone_Count))))
TreeDataCones1 <-TreeDataCones1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, Event_ID, Clump_Number, Stem_Letter, Species_Code, Mort_Year, TreeHeight_m, Tree_Status, StatusDead_Cause, FemaleCones_YN, Cone_Count, TreeData_Notes)

TestTreeCones1<- if(nrow(TreeDataCones1)>0) {
  write.xlsx(TreeDataCones1, file= "WBP_Validation.xlsx", sheetName = "TreeConesTrueNoConeCt", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Returns tree records where cones = false and cone count is populated- if records exist, write to xlsx 
TreeDataCones2 <- subset(DataTrees, ((DataTrees$FemaleCones_YN == 0 & !is.na(DataTrees$Cone_Count))))
TreeDataCones2 <-TreeDataCones2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, Mort_Year, TreeHeight_m, Tree_Status, StatusDead_Cause, FemaleCones_YN, Cone_Count, TreeData_Notes)

TestTreeCones2<- if(nrow(TreeDataCones2)>0) {
  write.xlsx(TreeDataCones2, file= "WBP_Validation.xlsx", sheetName = "TreeConesFalseConeCt", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where crown health doesn't equal domain values (1-5)- if records exist, write to xlsx 
TreeDataCrownHealth <- subset(DataTrees, ((DataTrees$Crown_Health < 1 | DataTrees$Crown_Health > 5)))
TreeDataCrownHealth <-TreeDataCrownHealth%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, Mort_Year, TreeHeight_m, Tree_Status, StatusDead_Cause, Crown_Health, TreeData_Notes)


TestTreeCrownHealth<- if(nrow(TreeDataCrownHealth)>0) {
  write.xlsx(TreeDataCrownHealth, file= "WBP_Validation.xlsx", sheetName = "TreeCrownHealthNot1_5", append = TRUE, row.names = FALSE, showNA = FALSE)
} 



#Returns tree records where crown kill lower is greater than 100%- if records exist, write to xlsx 
TreeDataCrownKillLow <- subset(DataTrees, ((DataTrees$CrownKill_Lower_perc >100 )))
TreeDataCrownKillLow <-TreeDataCrownKillLow%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestTreeCrownKillLow<- if(nrow(TreeDataCrownKillLow)>0) {
  write.xlsx(TreeDataCrownKillLow, file= "WBP_Validation.xlsx", sheetName = "TreeCrownKillLowGT100", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where crown kill middle is greater than 100%- if records exist, write to xlsx 
TreeDataCrownKillMid <- subset(DataTrees, ((DataTrees$CrownKill_Mid_perc >100 )))
TreeDataCrownKillMid <-TreeDataCrownKillMid%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestTreeCrownKillMid<- if(nrow(TreeDataCrownKillMid)>0) {
  write.xlsx(TreeDataCrownKillMid, file= "WBP_Validation.xlsx", sheetName = "TreeCrownKillMidGT100", append = TRUE, row.names = FALSE, showNA = FALSE)
} 



#Returns tree records where crown kill upper is greater than 100%- if records exist, write to xlsx 
TreeDataCrownKillUp <- subset(DataTrees, ((DataTrees$CrownKill_Upper_perc >100 )))
TreeDataCrownKillUp <-TreeDataCrownKillUp%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestTreeCrownKillUp<- if(nrow(TreeDataCrownKillUp)>0) {
  write.xlsx(TreeDataCrownKillUp, file= "WBP_Validation.xlsx", sheetName = "TreeCrownKillUpGT100", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where DBH is greater than 200 cm- if records exist, write to xlsx 
TreeDataDBH <- subset(DataTrees, ((DataTrees$TreeDBH_cm >200 & DataTrees$TreeDBH_cm != 999)))
TreeDataDBH <-TreeDataDBH%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeDBH_GT200<- if(nrow(TreeDataDBH)>0) {
  write.xlsx(TreeDataDBH, file= "WBP_Validation.xlsx", sheetName = "TreeDBH_GT200", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Returns tree records that are missing tree ID- if records exist, write to xlsx 
TreeDataMissingTagNumber <- subset(DataTrees, ((is.na(DataTrees$TreeID_Number))))
TreeDataMissingTagNumber <-TreeDataMissingTagNumber%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeNoTag<- if(nrow(TreeDataMissingTagNumber)>0) {
  write.xlsx(TreeDataMissingTagNumber, file= "WBP_Validation.xlsx", sheetName = "TreeNoTag", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records that are missing height- if records exist, write to xlsx 
TreeDataMissingHt <- subset(DataTrees, ((is.na(DataTrees$TreeHeight_m))))
TreeDataMissingHt <-TreeDataMissingHt%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeNoHt<- if(nrow(TreeDataMissingHt)>0) {
  write.xlsx(TreeDataMissingHt, file= "WBP_Validation.xlsx", sheetName = "TreeNoHt", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records that are missing DBH- if records exist, write to xlsx 
TreeDataMissingDBH <- subset(DataTrees, ((is.na(DataTrees$TreeDBH_cm))))
TreeDataMissingDBH <-TreeDataMissingDBH%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeNoDBH<- if(nrow(TreeDataMissingDBH)>0) {
  write.xlsx(TreeDataMissingDBH, file= "WBP_Validation.xlsx", sheetName = "TreeNoDBH", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records that are missing species code- if records exist, write to xlsx 
TreeDataMissingSp <- subset(DataTrees, ((is.na(DataTrees$Species_Code))))
TreeDataMissingSp <-TreeDataMissingSp%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeNoSp<- if(nrow(TreeDataMissingSp)>0) {
  write.xlsx(TreeDataMissingSp, file= "WBP_Validation.xlsx", sheetName = "TreeNoSp", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records that are missing subplot #- if records exist, write to xlsx 
TreeDataMissingSubplot <- subset(DataTrees, ((is.na(DataTrees$TreeData_SubPlot_StripID))))
TreeDataMissingSubplot <-TreeDataMissingSubplot%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeNoSubplot<- if(nrow(TreeDataMissingSubplot)>0) {
  write.xlsx(TreeDataMissingSubplot, file= "WBP_Validation.xlsx", sheetName = "TreeNoSubplot", append = TRUE, row.names = FALSE, showNA = FALSE)
} 



#Returns tree records that are missing status- if records exist, write to xlsx 
TreeDataMissingStatus <- subset(DataTrees, ((is.na(DataTrees$Tree_Status))))
TreeDataMissingStatus <-TreeDataMissingStatus%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeDBH_cm, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeNoStatus<- if(nrow(TreeDataMissingStatus)>0) {
  write.xlsx(TreeDataMissingStatus, file= "WBP_Validation.xlsx", sheetName = "TreeNoStatus", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Returns live PIAL records where crown kill percent is missing from the upper section but not missing from mid or lower sections- if records exist, write to xlsx 
TreeDataMissingCrownKillUp <- subset(DataTrees, (((DataTrees$Tree_Status == "L" & DataTrees$Species_Code == "PIAL") & is.na(DataTrees$CrownKill_Upper_perc) & ( !is.na(DataTrees$CrownKill_Mid_perc) | !is.na(DataTrees$CrownKill_Lower_perc)))))
TreeDataMissingCrownKillUp <-TreeDataMissingCrownKillUp%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestPIALNoCrownKillUp<- if(nrow(TreeDataMissingCrownKillUp)>0) {
  write.xlsx(TreeDataMissingCrownKillUp, file= "WBP_Validation.xlsx", sheetName = "TreePIALNoCrownKillUp", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns live PIAL records where crown kill percent is missing from the middle section but not missing from upper or lower sections- if records exist, write to xlsx 
TreeDataMissingCrownKillMid <- subset(DataTrees, (((DataTrees$Tree_Status == "L" & DataTrees$Species_Code == "PIAL") & is.na(DataTrees$CrownKill_Mid_perc)& (!is.na(DataTrees$CrownKill_Upper_perc) | !is.na(DataTrees$CrownKill_Lower_perc)))))
TreeDataMissingCrownKillMid <-TreeDataMissingCrownKillMid%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)

TestPIALNoCrownKillMid<- if(nrow(TreeDataMissingCrownKillMid)>0) {
  write.xlsx(TreeDataMissingCrownKillMid, file= "WBP_Validation.xlsx", sheetName = "TreePIALNoCrownKillMid", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns live PIAL records where crown kill percent is missing from the lower section but not missing from upper or mid sections- if records exist, write to xlsx 
TreeDataMissingCrownKillLow <- subset(DataTrees, (((DataTrees$Tree_Status == "L" & DataTrees$Species_Code == "PIAL") & is.na(DataTrees$CrownKill_Lower_perc)& (!is.na(DataTrees$CrownKill_Upper_perc) | !is.na(DataTrees$CrownKill_Mid_perc)))))
TreeDataMissingCrownKillLow <-TreeDataMissingCrownKillLow%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, CrownKill_Lower_perc, CrownKill_Mid_perc, CrownKill_Upper_perc, Crown_Health, TreeData_Notes)


TestPIALNoCrownKillLow<- if(nrow(TreeDataMissingCrownKillLow)>0) {
  write.xlsx(TreeDataMissingCrownKillLow, file= "WBP_Validation.xlsx", sheetName = "TreePIALNoCrownKillLow", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns live PIAL records that are missing crown health- if records exist, write to xlsx 
TreeDataCrownHealthPIAL <- subset(DataTrees, (((DataTrees$Tree_Status == "L" & DataTrees$Species_Code == "PIAL") & is.na(DataTrees$Crown_Health))))
TreeDataCrownHealthPIAL <-TreeDataCrownHealthPIAL%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestPIALNoCrownHealth<- if(nrow(TreeDataCrownHealthPIAL)>0) {
  write.xlsx(TreeDataCrownHealthPIAL, file= "WBP_Validation.xlsx", sheetName = "TreePIALNoCrownHealth", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records with a species that doesn't match a species in the lookup table- if records exist, write to xlsx 
names(SpeciesList)[1]<- "Species_Code" #renames field so that I can merge data
names(SpeciesList)[2]<- "Unit_Code" #renames field so that I can merge data
TreesNoSp <- subset(DataTrees, ((!is.na(DataTrees$Species_Code)))) #removes records where no seedling species was recorded
TreesNoSpJoin <- anti_join(TreesNoSp, SpeciesList, by = 'Species_Code', 'Unit_Code') #joins lookup table to data table
TreesNoSpJoin <-TreesNoSpJoin%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeSpNotInLU<- if(nrow(TreesNoSpJoin)>0) {
  write.xlsx(TreesNoSpJoin, file= "WBP_Validation.xlsx", sheetName = "TreeSpNotInLU", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where status does not equal a domain value (L, RD, D)- if records exist, write to xlsx 
TreeDataStatus <- subset(DataTrees, ((DataTrees$Tree_Status != "L" & DataTrees$Tree_Status != "D" & DataTrees$Tree_Status != "RD")))
TreeDataStatus <-TreeDataStatus%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeStatusDomain<- if(nrow(TreeDataStatus)>0) {
  write.xlsx(TreeDataStatus, file= "WBP_Validation.xlsx", sheetName = "TreeStatusNoDomain", append = TRUE, row.names = FALSE, showNA = FALSE)
} 



#Returns tree records that have a stripID that is not 1-5- if records exist, write to xlsx 
TreeDataSubplot <- subset(DataTrees, ((DataTrees$TreeData_SubPlot_StripID < 1 | DataTrees$TreeData_SubPlot_StripID > 5)))
TreeDataSubplot <-TreeDataSubplot%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeSubplot<- if(nrow(TreeDataSubplot)>0) {
  write.xlsx(TreeDataSubplot, file= "WBP_Validation.xlsx", sheetName = "TreeSubplotNot1_5", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where stem letter is not a letter- if records exist, write to xlsx 
TreeDataStem <- subset(DataTrees, !is.na(DataTrees$Stem_Letter))
TreeDataStem2 <- subset(TreeDataStem, !grepl("^[[:alpha:]]+$", TreeDataStem$Stem_Letter))
TreeDataStem2 <-TreeDataStem2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeStemLetter<- if(nrow(TreeDataStem2)>0) {
  write.xlsx(TreeDataStem2, file= "WBP_Validation.xlsx", sheetName = "TreeStemLetterNotALetter", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where lower bole canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBoleCankLow1 <- subset(DataTrees, (DataTrees$BoleCankers_I_Lower_YN == 1 & is.na(DataTrees$BoleCanks_ITypes_Lower)))
TreeDataBoleCankLow1 <-TreeDataBoleCankLow1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Lower_YN, BoleCanks_ITypes_Lower, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeBoleCankLow1<- if(nrow(TreeDataBoleCankLow1)>0) {
  write.xlsx(TreeDataBoleCankLow1, file= "WBP_Validation.xlsx", sheetName = "TreeBoleLowCankTrueNoInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where lower bole canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBoleCankLow2 <- subset(DataTrees, (DataTrees$BoleCankers_I_Lower_YN == 0 & !is.na(DataTrees$BoleCanks_ITypes_Lower)))
TreeDataBoleCankLow2 <-TreeDataBoleCankLow2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Lower_YN, BoleCanks_ITypes_Lower, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeBoleCankLow2<- if(nrow(TreeDataBoleCankLow2)>0) {
  write.xlsx(TreeDataBoleCankLow2, file= "WBP_Validation.xlsx", sheetName = "TreeBoleLowCankFalseHasInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where middle bole canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBoleCankMid1 <- subset(DataTrees, (DataTrees$BoleCankers_I_Mid_YN == 1 & is.na(DataTrees$BoleCanks_ITypes_Mid)))
TreeDataBoleCankMid1 <-TreeDataBoleCankMid1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Mid_YN, BoleCanks_ITypes_Mid, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeBoleCankMid1<- if(nrow(TreeDataBoleCankMid1)>0) {
  write.xlsx(TreeDataBoleCankMid1, file= "WBP_Validation.xlsx", sheetName = "TreeBoleMidCankTrueNoInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where middle bole canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBoleCankMid2 <- subset(DataTrees, (DataTrees$BoleCankers_I_Mid_YN == 0 & !is.na(DataTrees$BoleCanks_ITypes_Mid)))
TreeDataBoleCankMid2 <-TreeDataBoleCankMid2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Mid_YN, BoleCanks_ITypes_Mid, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeBoleCankMid2<- if(nrow(TreeDataBoleCankMid2)>0) {
  write.xlsx(TreeDataBoleCankMid2, file= "WBP_Validation.xlsx", sheetName = "TreeBoleMidCankFalseHasInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where upper bole canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBoleCankUpp1 <- subset(DataTrees, (DataTrees$BoleCankers_I_Upper_YN == 1 & is.na(DataTrees$BoleCanks_ITypes_Upper)))
TreeDataBoleCankUpp1 <-TreeDataBoleCankUpp1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Upper_YN, BoleCanks_ITypes_Upper, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeBoleCankUpp1<- if(nrow(TreeDataBoleCankUpp1)>0) {
  write.xlsx(TreeDataBoleCankUpp1, file= "WBP_Validation.xlsx", sheetName = "TreeBoleUpCankTrueNoInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
} 

#Returns tree records where upper bole canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBoleCankUpp2 <- subset(DataTrees, (DataTrees$BoleCankers_I_Upper_YN == 0 & !is.na(DataTrees$BoleCanks_ITypes_Upper)))
TreeDataBoleCankUpp2 <-TreeDataBoleCankUpp2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BoleCankers_I_Upper_YN, BoleCanks_ITypes_Upper, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeBoleCankUpp2<- if(nrow(TreeDataBoleCankUpp2)>0) {
  write.xlsx(TreeDataBoleCankUpp2, file= "WBP_Validation.xlsx", sheetName = "TreeBoleUpCankFalseHasInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
} 


#Returns tree records where lower branch canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBranchCankLow1 <- subset(DataTrees, (DataTrees$BranchCanks_I_Lower_YN == 1 & is.na(DataTrees$BranchCanks_ITypes_Lower)))
TreeDataBranchCankLow1 <-TreeDataBranchCankLow1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Lower_YN, BranchCanks_ITypes_Lower, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeBranchCankLow1<- if(nrow(TreeDataBranchCankLow1)>0) {
  write.xlsx(TreeDataBranchCankLow1, file= "WBP_Validation.xlsx", sheetName = "TreeBranchLowCankTrueNoInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns tree records where lower branch canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBranchCankLow2 <- subset(DataTrees, (DataTrees$BranchCanks_I_Lower_YN == 0 & !is.na(DataTrees$BranchCanks_ITypes_Lower)))
TreeDataBranchCankLow2 <-TreeDataBranchCankLow2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Lower_YN, BranchCanks_ITypes_Lower, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeBranchCankLow2<- if(nrow(TreeDataBranchCankLow2)>0) {
  write.xlsx(TreeDataBranchCankLow2, file= "WBP_Validation.xlsx", sheetName = "TreeBranchLowCankFalseHasInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns tree records where middle branch canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBranchCankMid1 <- subset(DataTrees, (DataTrees$BranchCanks_I_Mid_YN == 1 & is.na(DataTrees$BranchCanks_ITypes_Mid)))
TreeDataBranchCankMid1 <-TreeDataBranchCankMid1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Mid_YN, BranchCanks_ITypes_Mid, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeBranchCankMid1<- if(nrow(TreeDataBranchCankMid1)>0) {
  write.xlsx(TreeDataBranchCankMid1, file= "WBP_Validation.xlsx", sheetName = "TreeBranchMidCankTrueNoInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
}



#Returns tree records where middle branch canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBranchCankMid2 <- subset(DataTrees, (DataTrees$BranchCanks_I_Mid_YN == 0 & !is.na(DataTrees$BranchCanks_ITypes_Mid)))
TreeDataBranchCankMid2 <-TreeDataBranchCankMid2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Mid_YN, BranchCanks_ITypes_Mid, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)

TestTreeBranchCankMid2<- if(nrow(TreeDataBranchCankMid2)>0) {
  write.xlsx(TreeDataBranchCankMid2, file= "WBP_Validation.xlsx", sheetName = "TreeBranchMidCankFalseHasInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
}



#Returns tree records where upper branch canks infestation checkbox = true but infestation type is null- if records exist, write to xlsx 
TreeDataBranchUppCank1 <- subset(DataTrees, (DataTrees$BranchCanks_I_Upper_YN == 1 & is.na(DataTrees$BranchCanks_ITypes_Upper)))
TreeDataBranchUppCank1 <-TreeDataBranchUppCank1%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Upper_YN, BranchCanks_ITypes_Upper, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeBranchUppCank1<- if(nrow(TreeDataBranchUppCank1)>0) {
  write.xlsx(TreeDataBranchUppCank1, file= "WBP_Validation.xlsx", sheetName = "TreeBranchUpCankTrueNoInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns tree records where upper branch canks infestation checkbox = false but infestation type is populated- if records exist, write to xlsx 
TreeDataBranchCankUpp2 <- subset(DataTrees, (DataTrees$BranchCanks_I_Upper_YN == 0 & !is.na(DataTrees$BranchCanks_ITypes_Upper)))
TreeDataBranchCankUpp2 <-TreeDataBranchCankUpp2%>%
  select(Unit_Code, PlotID_Number, TreeID_Number, TreeData_SubPlot_StripID, BranchCanks_I_Upper_YN, BranchCanks_ITypes_Upper, Start_Date, TreeData_ID, Clump_Number, Stem_Letter, Species_Code, TreeHeight_m, Tree_Status, Crown_Health, TreeData_Notes)


TestTreeBranchCankUpp2 <- if(nrow(TreeDataBranchCankUpp2)>0) {
  write.xlsx(TreeDataBranchCankUpp2, file= "WBP_Validation.xlsx", sheetName = "TreeBranchUpCankFalseHasInfestData", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns tree records where Branch Canks upper field values don't match domain values
InfestBranchCanksUp1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BranchCanks_ITypes_Upper)
InfestBranchCanksUp2 <- subset(InfestBranchCanksUp1, (!is.na(DataTrees$BranchCanks_ITypes_Upper)))
InfestBranchCanksUp3 <- InfestBranchCanksUp2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBranchCanksUp2$BranchCanks_ITypes_Upper),]



TestInfestBranchCanksUp <- if(nrow(InfestBranchCanksUp3)>0) {
    write.xlsx(InfestBranchCanksUp3, file= "WBP_Validation.xlsx", sheetName = "TreeBranchUpCankInfestDomains", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns tree records where Branch Canks middle field values don't match domain values
InfestBranchCanksMid1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BranchCanks_ITypes_Mid)
InfestBranchCanksMid2 <- subset(InfestBranchCanksMid1, (!is.na(DataTrees$BranchCanks_ITypes_Mid)))
InfestBranchCanksMid3 <- InfestBranchCanksMid2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBranchCanksMid2$BranchCanks_ITypes_Mid),]
TestInfestBranchCanksMid <- if(nrow(InfestBranchCanksMid3)>0) {
  write.xlsx(InfestBranchCanksMid3, file= "WBP_Validation.xlsx", sheetName = "TreeBranchMidCankInfestDomains", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns tree records where Branch Canks lower field values don't match domain values
InfestBranchCanksLow1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BranchCanks_ITypes_Lower)
InfestBranchCanksLow2 <- subset(InfestBranchCanksLow1, (!is.na(DataTrees$BranchCanks_ITypes_Lower)))
InfestBranchCanksLow3 <- InfestBranchCanksLow2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBranchCanksLow2$BranchCanks_ITypes_Lower),]
TestInfestBranchCanksLow <- if(nrow(InfestBranchCanksLow3)>0) {
  write.xlsx(InfestBranchCanksLow3, file= "WBP_Validation.xlsx", sheetName = "TreeBranchLowCankInfestDomains", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#Returns tree records where Bole Canks upper field values don't match domain values
InfestBoleCanksUp1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BoleCanks_ITypes_Upper)
InfestBoleCanksUp2 <- subset(InfestBoleCanksUp1, (!is.na(DataTrees$BoleCanks_ITypes_Upper)))
InfestBoleCanksUp3 <- InfestBoleCanksUp2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBoleCanksUp2$BoleCanks_ITypes_Upper),]
TestInfestBoleCanksUp <- if(nrow(InfestBoleCanksUp3)>0) {
  write.xlsx(InfestBoleCanksUp3, file= "WBP_Validation.xlsx", sheetName = "TreeBoleUpCankInfestDomains", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns tree records where Bole Canks middle field values don't match domain values
InfestBoleCanksMid1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BoleCanks_ITypes_Mid)
InfestBoleCanksMid2 <- subset(InfestBoleCanksMid1, (!is.na(DataTrees$BoleCanks_ITypes_Mid)))
InfestBoleCanksMid3 <- InfestBoleCanksMid2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBoleCanksMid2$BoleCanks_ITypes_Mid),]
TestInfestBoleCanksMid <- if(nrow(InfestBoleCanksMid3)>0) {
  write.xlsx(InfestBoleCanksMid3, file= "WBP_Validation.xlsx", sheetName = "TreeBoleMidCankInfestDomains", append = TRUE, row.names = FALSE, showNA = FALSE)
}

#Returns tree records where Bole Canks lower field values don't match domain values
InfestBoleCanksLow1 <-DataTrees%>%
  select(Unit_Code, PlotID_Number, BoleCanks_ITypes_Lower)
InfestBoleCanksLow2 <- subset(InfestBoleCanksLow1, (!is.na(DataTrees$BoleCanks_ITypes_Lower)))
InfestBoleCanksLow3 <- InfestBoleCanksLow2[grepl("B|D|E|G|H|I|J|K|L|M|N|P|Q|T|U|V|W|X|Y|Z", InfestBoleCanksLow2$BoleCanks_ITypes_Lower),]
TestInfestBoleCanksLow <- if(nrow(InfestBoleCanksLow3)>0) {
  write.xlsx(InfestBoleCanksLow3, file= "WBP_Validation.xlsx", sheetName = "TreeBoleLowCankInfestDomains", append = TRUE, row.names = FALSE, showNA = FALSE)
}


#############CLOSE THE CONNECTION TO THE DATABASE#####################################################################################
odbcCloseAll()

