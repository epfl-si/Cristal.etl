#install.packages("tidyverse")
#install.packages("readxl")
#install.packages("xlsx")
#install.packages("cellranger")
library(tidyverse)
library(readxl)
library(dplyr)

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

    xlsx::addDataFrame(data.frame(data, check.names = FALSE), sheet,
                       startRow = 2, row.names = FALSE,
                       colnamesStyle = xlsx::CellStyle(wb) +
                           xlsx::Font(wb, isBold = TRUE) +
                           xlsx::Border(color = "black", position = c("TOP", "BOTTOM"),
                                        pen = c("BORDER_THIN", "BORDER_THICK")))

    xlsx::addAutoFilter(sheet, A1.range(2, 1, 2, ncol(data)))

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

toArchibusBat <- function(batId) {
    # si bat est ZZ alors transormer en ZE
    case_when(batId == "ZZ" ~ "ZE",
              TRUE          ~ batId)
}

toArchibusCF4 <- function(cf4Id) {
    case_when(cf4Id == 0 ~ "")
    read.csv("./organisation.txt",header = T, sep = ";",check.names = F)
             
}


mission_import <-read_excel("./4DMissions_03.03.22.xlsx")
cf_import <- read.csv("./organisation.txt",header = T, sep = ";",check.names = F, fileEncoding = "latin1") %>%
   rename(CFNo = `No unité`) %>%
   separate(Hiérarchie, sep = " ", fill = "right", into = c(NA, "CF2", NA, NA, NA))

cf <- cf_import %>%
  filter(Au == "")  ## Seulement la dernière identité de l'unité;
                    ## N.B.: cela oblige Archibus à mettre à jour cette donnée
                    ## avant l'import (s'assurer que c'est le cas !)

mission_archibus <-
  mission_import %>%
  left_join(cf) %>%
  transmute("#project.project_id" = IDMission,
            project_name = Intitule,
            project_type = PriorisationType, #attention il faudra mettre les ids
            description = "",
            status = toArchibusStatus(EtatMission),
            criticality = 3, # à définir valeur par défaut
            site_id = "",
            bl_id = toArchibusBat(BatimentID),
            dv_id = CF2,
            dp_id =  Sigle) 
  
write_archibus(mission_archibus, "./missions-archibus.xlsx",
               table.header = "Activity Projects",
               sheet.name = "4d_projects")

