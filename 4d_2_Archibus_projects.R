#install.packages("tidyverse")
#install.packages("readxl")
#install.packages("xlsx")
library(tidyverse)
library(readxl)

write_archibus <- function(data, filename, table.header) {
    wb <- xlsx::createWorkbook(type="xlsx")
    # Create a sheet in that workbook to contain the data table
    sheet <- xlsx::createSheet(wb, sheetName = "4d_projetcts")

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

    # save the workbook to an Excel file
    xlsx::saveWorkbook(wb, filename)
}

toArchibusStatus <- function(etat) {
    # Juste pour montrer comment on fait — En attente des codes de la part de Thomas :
    case_when(etat == "Terminé" ~ 1,
              etat == "Etude"   ~ 2,
              TRUE              ~ 7)
}

mission_import <- read_excel("/Users/venries/Downloads/MISSIONS_Export-direct_table.xlsx")

mission_archibus <-
  mission_import %>%
  transmute("#project.project_id" = IDMission,
            project_name = Intitule,
            project_type = PriorisationType,
            description = "",
            status = toArchibusStatus(EtatMission))  
  
  
write_archibus(mission_archibus, "./missions-archibus.xlsx",
               table.header = "Activity Projects")
