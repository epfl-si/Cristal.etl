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

#setwd("/Users/venries/GitHub/Cristal")

Missions4D_file = "./4DMissions_08.04.2022.csv"

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
    csWrap <- xlsx::CellStyle(wb) + xlsx::Alignment(wrapText = TRUE)
    
    colwrap <- list(
      '4' = csWrap
    )
    colnum <- list(
      '9' = csNum,
      '10' = csNum
    )
    coldate <- list(
      '13' = csDate,
      '14' = csDate
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

mission_import <- fread(file = Missions4D_file  , encoding = "Latin-1") %>%
  filter (Etat != "") %>% ## Seulement si la Etat est reseigné
  mutate(CFNo = ifelse(is.na(CFNo), 0, CFNo)) %>% ## remplace NA par 0 fabs CFNo
  mutate(BatimentID = replace(`Bât.`, `Bât.` == "ZZ", "ZE")) %>% ## remplace ZZ par ZE dans BatimentID
  mutate(BatimentID = replace(`Bât.`, `Bât.` == "SK", "SKIL")) %>%
  mutate(CFNo = replace(CFNo, CFNo == 0, NA)) %>%
  mutate(Estimatif = replace(Estimatif, Estimatif == 0, NA)) %>%
  mutate(BudgetTotal = replace(BudgetTotal, BudgetTotal == 0, NA)) %>%
  mutate(FullNameDemandeur = tolower(Demandeur))  %>%
  mutate(CP = as.numeric(`CP SCIPER`)) %>%
  mutate(Début = replace(Début, Début == "00.00.00", NA)) %>%
  mutate(Remise = replace(Remise, Remise == "00.00.00", NA))

batiments_import <- read_excel("./Export Bâtiments.xlsx")

dp <- read_excel("./export DP.xlsx")
         
em <- read_excel("./em.xlsx", skip = 1)  %>%
  mutate(FullName = tolower(paste(name_last, name_first)) ) %>%
#  mutate(sciper = as.numeric(gsub(".*- ","",em_id)))
  mutate(sciper = as.numeric(em_number))
m <- mission_import %>%
  transmute(FullNameDemandeur=FullNameDemandeur,
            CP = CP)
m2 <- m %>%
  left_join(em, by=c("FullNameDemandeur"= "FullName"),suffix = c("","_d")) %>%
  left_join(em, by=c("CP" = "sciper"),suffix = c("","cp"))

mission_archibus <-
  mission_import %>%
  left_join(batiments_import, by=c("BatimentID"="Building Code")) %>%
  left_join(dp, by=c("CFNo"="dp_id")) %>%
  left_join(em, by=c("FullNameDemandeur"= "FullName"),suffix = c("","_d")) %>%
  left_join(em, by=c("CP" = "sciper"),suffix = c("","cp")) %>%
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
            proj_mgr = `#em.em_idcp`, ## ifelse(is.na(em_idcp), ChefProjet, em_idcp),
            scope=ServiceTraitant,
            benefit=Justificatif,
            contact_id="TBD")
  
write_archibus(mission_archibus, "./01_projects.xlsx",
               table.header = "Activity Projects",
               sheet.name = "4d_projects")



