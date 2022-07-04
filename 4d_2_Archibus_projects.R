#install.packages("tidyverse")
#install.packages("readxl")
#install.packages("xlsx")
#install.packages("cellranger")
#install.packages("lubridate")
library(tidyverse)
library(readxl)
library(dplyr)
library(lubridate)
library("data.table") 
library("stringr")
options(scipen = 999)   
#setwd("/Users/venries/GitHub/Cristal")

Missions4D_file = "./4DMissions_06.05.2022.csv"
Devis4D_file = "./4DDevisCFC_06.05.2022.csv"
Tresoreries4D_file="./4DTresoreries_06.05.2022.csv"
Revue4D_file="./4DRevuesProjets_06.05.2022.csv"

A1 <- function(row, col) {
    #' Convert real-world (integer) coordinates to Excel®-style A1 notation.
    dollar_a1 <-
        cellranger::ra_ref(row_ref = row, col_ref = col) %>%
        cellranger::to_string(fo = "A1")
    str_replace_all(dollar_a1, '[$]', '')
}
stopifnot(A1(20, 27) == "AA20")

A1.range <- function(from_row, from_col, to_row, to_col) {
    paste(
        A1(from_row, from_col),
        A1(to_row, to_col),
        sep = ":")
}

write_archibus <- function(data, filename, table.header, sheet.name = "sheet1") {
    wb <- xlsx::createWorkbook(type = "xlsx")
    sheet <- xlsx::createSheet(wb, sheetName = sheet.name)

    # Add Archibus-style header
    cell <- xlsx::createCell(xlsx::createRow(sheet, rowIndex = 1), colIndex = 1)
    xlsx::setCellValue(cell[[1,1]], paste("#", table.header, sep = ""))
    xlsx::setCellStyle(cell[[1,1]],
                       xlsx::CellStyle(wb) + xlsx::Font(wb, heightInPoints=22, isBold=TRUE))
    csDate <- xlsx::CellStyle(wb) + xlsx::DataFormat("yyyy-mm-dd")
    csNum <- xlsx::CellStyle(wb) + xlsx::DataFormat("0000")
    csDeci <- xlsx::CellStyle(wb) + xlsx::DataFormat("#,##0.00")
    csWrap <- xlsx::CellStyle(wb) + xlsx::Alignment(wrapText = TRUE)
    
    coldeci <- list('3'= csDeci, '4'= csDeci)
    colwrap <- list('4' = csWrap)
    colnum <- list(
      '9' = csNum,
      '10' = csNum
    )
    coldate <- list(
      '13' = csDate,
      '14' = csDate,
      '20' = csDate
    )

    xlsx::addDataFrame(data.frame(data, check.names = FALSE), sheet,
                       startRow = 2, row.names = FALSE,
                       colStyle=c(colwrap,colnum,coldate),
    #                   colStyle=c(colnum),
                       colnamesStyle = xlsx::CellStyle(wb) +
                           xlsx::Font(wb, isBold = TRUE) +
                           xlsx::Border(color = "black", position = c("TOP", "BOTTOM"),
                                        pen = c("BORDER_THIN", "BORDER_THICK")))

    # save the workbook to an Excel file
    xlsx::saveWorkbook(wb, filename)
}


toArchibusStatus <- function(etat) {

    case_when(etat == "Terminé"   ~ "Completed-Verified",
              etat == "Etude"     ~ "Approved-In Design",
              etat == "Exécution" ~ "Issued-In Process",
              etat == "Annulé"    ~ "Approved-Cancelled",
              etat == "Faux"      ~ "Issued-On Hold",
              etat == ""          ~ "Issued-On Hold",
              TRUE                ~ "")
}

mission_import  <- fread(file = Missions4D_file  , encoding = "Latin-1") %>%
  filter (Etat != "") %>% ## Seulement si la Etat est reseigné
  mutate(CFNo = ifelse(is.na(CFNo), 0, CFNo)) %>% ## remplace NA par 0 fabs CFNo
  mutate(BatimentID = replace(`Bât.`, `Bât.` == "ZZ", "ZE")) %>% ## remplace ZZ par ZE dans BatimentID
  mutate(BatimentID = replace(`Bât.`, `Bât.` == "SK", "SKIL")) %>%
  mutate(CFNo = replace(CFNo, CFNo == 0, NA)) %>%
  mutate(Estimatif = replace(Estimatif, Estimatif == 0, NA)) %>%
  mutate(BudgetTotal = replace(BudgetTotal, BudgetTotal == 0, NA)) %>%
  mutate(FullNameDemandeur = tolower(Demandeur))  %>%
  mutate(FullNameCP = tolower(`CP Nom`)) %>%
  mutate(CP = as.numeric(`CP SCIPER`)) %>%
  mutate(Début = as.Date(replace(Début, Début == "00.00.00", NA),"%d.%m.%Y")) %>%
  mutate(Remise = as.Date(replace(Remise, Remise == "00.00.00", NA),"%d.%m.%Y")) %>%
  mutate(Datemission = as.Date(replace(`Date mission`, `Date mission` == "00.00.00", NA),"%d.%m.%Y"))

batiments_import <- read_excel("./Export Bâtiments.xlsx")

dp <- read_excel("./export DP.xlsx")
         
em <- read_excel("./em2.xlsx", skip = 1)  %>%
  mutate(FullName = tolower(paste(name_last, name_first)) ) %>%
#  mutate(sciper = as.numeric(gsub(".*- ","",em_id)))
  mutate(sciper = as.numeric(em_number))

m <- mission_import %>%
  transmute(FullNameDemandeur=FullNameDemandeur,
            FullNameCP = FullNameCP,
            CP = CP)
m2 <- m %>%
  left_join(em, by=c("FullNameDemandeur"= "FullName"),suffix = c("","_d")) %>%
  left_join(em, by=c("FullNameCP" = "FullName"),suffix = c("","c2")) %>%
  left_join(em, by=c("CP" = "sciper"),suffix = c("","cp")) %>%
  transmute (CP=CP,
              FullNameCP=FullNameCP,
              `#em.em_idcp` = `#em.em_idcp`,
              `#em.em_idc2` = `#em.em_idc2`)

mission_archibus <-
  mission_import %>%
  left_join(batiments_import, by=c("BatimentID"="Building Code")) %>%
  left_join(dp, by=c("CFNo"="dp_id")) %>%
  left_join(em, by=c("FullNameDemandeur"= "FullName"),suffix = c("","_d")) %>%
  left_join(em, by=c("CP" = "sciper"),suffix = c("","cp")) %>%
  left_join(em, by=c("FullNameCP" = "FullName"),suffix = c("","c2")) %>%
  transmute("#project.project_id" = `ID Mission`,
            project_name = Intitulé,
            project_type = Priorisation, #attention il faudra mettre les ids
            description = str_replace_all(Description,"###","\n") ,
            status = toArchibusStatus(Etat),
            criticality = "Noncritical", # à définir valeur par défaut
            site_id = SiteCode,
            bl_id = ifelse(is.na(SiteCode), NA, BatimentID), ## if faut que le site existe pour renseigner bl_id
            dv_id = dv_id,
            dp_id = ifelse(is.na(dv_id), NA, CFNo), ## if faut que CF2 existe pour renseigner CF4
            cost_est_baseline = Estimatif,
            cost_budget = BudgetTotal,
            date_start = Début,
            date_end = Remise,
            requestor = `#em.em_id`, ## ifelse(is.na(em_id), Demandeur, em_id),
            proj_mgr =ifelse(is.na(`#em.em_idcp`), `#em.em_idc2`, `#em.em_idcp`),
            scope=`CF infos`,
            benefit=Justificatif,
            contact_id="TBD",
            date_created=Datemission)
  
write_archibus(mission_archibus, "./01_projects.xlsx",
               table.header = "Activity Projects",
               sheet.name = "4d_projects")


#=================================
#  Action Items
#=================================    
devis_import <- fread(file = Devis4D_file  , encoding = "Latin-1") %>%
  mutate(CFC=str_replace_all(CFC,"[ *]|[.]$",""))
#  mutate(as.numeric(`Montant EPFL`)) %>%
#  mutate(`Montant BLL` = format(as.numeric(`Montant BBL`,scientific = FALSE)))


tresoreries_import <- fread(file = Tresoreries4D_file  , encoding = "Latin-1") %>%
  filter(Année != 0) %>%
  group_by(`ID Mission`) %>%
  summarise(Année = first(Année)) 
revue_import <- fread(file = Revue4D_file, encoding = "Latin-1") %>%
  filter (`Revue no`== 1) %>%
  mutate(Tot_Montant123=Montant1+Montant2+Montant3) %>%
  filter (`Tot_Montant123`!= 0) %>%
  ### filter les montant à 0
  transmute("#activity_log.activity_log_id" = "",
            activity_type = "PROJECT - COST",
            cost_est_cap = Tot_Montant123,
            cost_est_design_cap = "",
            csi_id ="999",
            date_scheduled = ifelse(Date=="00.00.00","",format(as.Date(Date,"%d.%m.%Y"), format="%Y-%m-%d")),
            project_id = `ID Mission`,
            status = "PLANNED",
            action_title = "Budget initial du projet repris de 4D",
            source_type = "EPFL",
            ar_is_change_order = 'No',
            cost_act_cap ="") %>%
  mutate(cost_est_cap=as.numeric(cost_est_cap),
         cost_est_design_cap=as.numeric(cost_est_design_cap))


actionItems <- devis_import %>%
  left_join(tresoreries_import, by=c("ID Mission"= "ID Mission")) %>%
  left_join(revue_import,by=c("ID Mission"= "project_id"),suffix = c("","_r")) %>%
  transmute("#activity_log.activity_log_id" = "",
            activity_type = "PROJECT - COST",
            cost_est_cap = ifelse(!is.na(`cost_est_cap` ) , "", ifelse(`Montant BBL`>0,as.numeric(`Montant BBL`),as.numeric(`Montant EPFL`))),
            cost_est_design_cap = format(ifelse(`Montant BBL`>0,`Montant BBL`,`Montant EPFL`),scientific = FALSE),
            csi_id = CFC,
            date_scheduled = ifelse(is.na(Année),"",paste(Année,"-01-01",sep = "")),
            project_id = `ID Mission`,
            status = "PLANNED",
            action_title = paste("Buget initial repris 4D - ", Travaux, sep = ""),
            source_type = ifelse(substr(CFC,1,1) =="3" | substr(CFC,1,1) == "9", "EPFL", "BBL"),
            ar_is_change_order = 'No',
            cost_act_cap ="") %>%
  mutate(cost_est_cap=as.numeric(cost_est_cap),
         cost_est_design_cap=as.numeric(cost_est_design_cap))

tresoreries_import_since_2022_BBL <- fread(file = Tresoreries4D_file  , encoding = "Latin-1") %>%
#  filter(`Année` > 2021) %>%
  filter(`Montant BBL` > 0) %>%
  transmute("#activity_log.activity_log_id" = "",
            activity_type = "PROJECT - COST",
            cost_est_cap = "",
            cost_est_design_cap = "",
            csi_id = "799",
            date_scheduled = ifelse(is.na(`Année`),"",paste(`Année`,"-01-01",sep = "")),
            project_id = `ID Mission`,
            status = "PLANNED",
            action_title = "Atterrissage repris 4D",
            source_type = "BBL",
            ar_is_change_order = 'No',
            cost_act_cap = `Montant BBL`)

tresoreries_import_since_2022_EPFL <- fread(file = Tresoreries4D_file  , encoding = "Latin-1") %>%
#  filter(`Année` > 2021) %>%
  filter(`Montant EPFL` != 0) %>%
  transmute("#activity_log.activity_log_id" = "",
             activity_type = "PROJECT - COST",
             cost_est_cap = "",
             cost_est_design_cap = "",
             csi_id = "999",
             date_scheduled = ifelse(is.na(`Année`),"",paste(`Année`,"-01-01",sep = "")),
             project_id = `ID Mission`,
             status = "PLANNED",
             action_title = "Atterrissage repris 4D",
             source_type = "EPFL",
             ar_is_change_order = 'No',
             cost_act_cap =  `Montant EPFL`)


actionItems_final <- rbind(actionItems, revue_import, tresoreries_import_since_2022_BBL, tresoreries_import_since_2022_EPFL)


write_archibus(actionItems_final, "./04_ActionItems.xlsx",
               table.header = "Action Items",
               sheet.name = "4d_projects")



#=================================
#  Project Funds
#=================================

mission_import_last_credit_BBL <- mission_import %>%
  filter((`Dernier Crédit BBL`!="")) %>%
  transmute("ID Mission" = `ID Mission`,
            "Dernier Credit" = `Dernier Crédit BBL`)

devis_EPFL <-devis_import %>%
  filter(`Montant EPFL` != 0) %>%
  filter(substr(CFC,1,1) =="3" | substr(CFC,1,1) == "9" ) %>%
  group_by(`ID Mission`) %>% 
  summarise(Montant = sum(`Montant EPFL`)) %>%
  left_join(tresoreries_import,by=c("ID Mission"="ID Mission")) %>%
  left_join(mission_import_last_credit_EPFL,by=c("ID Mission"="ID Mission"))
devis_EPFL$type = "EPFL-TBD"

mission_import_last_credit_EPFL <- mission_import %>%
  filter((`Dernier Crédit EPFL`!="")) %>%
  transmute("ID Mission" = `ID Mission`,
            "Dernier Credit" = `ID Mission`)

devis_BBL <-devis_import %>% 
  filter(`Montant BBL` != 0) %>%
  filter(substr(CFC,1,1) !="3" & substr(CFC,1,1) != "9" | is.na(CFC)) %>%
  group_by(`ID Mission`) %>% 
  summarise(Montant = sum(`Montant BBL`))  %>%
  left_join(tresoreries_import,by=c("ID Mission"="ID Mission")) %>%
  left_join(mission_import_last_credit_BBL,by=c("ID Mission"="ID Mission")) 
devis_BBL$type = "BBL-TBD"

devis <- rbind(devis_EPFL,devis_BBL)

projectfunds_final <- devis %>%
  transmute("#projfunds.project_id" = `ID Mission`,
            "fund_id" = type,
            "fiscal_year" = `Année`,
            "projfunds.ar_sap_fund_ref" = `Dernier Credit`,
            "amount_cap" = Montant,
            "description" = "")

write_archibus(projectfunds_final, "./03_ProjectFunds.xlsx",
               table.header = "Project Funds",
               sheet.name = "4d_projects")
