#install.packages("tidyverse")
#install.packages("readxl")
#install.packages("xlsx")
#install.packages("cellranger")
library(tidyverse)
library(readxl)

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
    # Juste pour montrer comment on fait — En attente des codes de la part de Thomas :
    case_when(etat == "Terminé" ~ 1,
              etat == "Etude"   ~ 2,
              TRUE              ~ 7)
}



mission_import <- read_excel("./MISSIONS_Export-direct_table.xlsx")

mission_archibus <-
  mission_import %>%
  transmute("#project.project_id" = IDMission,
            project_name = Intitule,
            project_type = PriorisationType, #attention il faudra mettre les ids
            description = "",
            status = toArchibusStatus(EtatMission))  
  
write_archibus(mission_archibus, "./missions-archibus.xlsx",
               table.header = "Activity Projects",
               sheet.name = "4d_projects")

