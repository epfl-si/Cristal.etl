#install.packages("tidyverse")
#install.packages("readxl")

#install.packages("devtools")
#devtools::install_github("kassambara/r2excel")
library(tidyverse)
library(readxl)
#library(xlsx)
library(r2excel)

write_archibus <- function(data, filename) {
    wb <- createWorkbook(type="xlsx")
    # Create a sheet in that workbook to contain the data table
    sheet <- createSheet(wb, sheetName = "4d_projetcts")

    # Add header. TODO: respect expectations of Archibus import module
    xlsx.addHeader(wb, sheet, value="Add table",level=1, 
               color="black", underline=1)
    xlsx.addLineBreak(sheet, 1)

    # Add paragraph : Author
    author=paste("Author : Alboukadel KASSAMBARA. \n",
                 "@:alboukadel.kassambara@gmail.com.",
                 "\n Website : http://ww.sthda.com", sep="")
    xlsx.addParagraph(wb, sheet,value=author, isItalic=TRUE, colSpan=5, 
                      rowSpan=4, fontColor="darkgray", fontSize=14)
    xlsx.addLineBreak(sheet, 3)

    # Add table : add a data frame
    xlsx.addTable(wb, sheet, data, startCol=1)
    xlsx.addLineBreak(sheet, 2)

    # save the workbook to an Excel file
    saveWorkbook(wb, filename)
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
  
  
# Décommentr la fin pour produire un .xslx (pas nécessaire dans R studio) :
# write_archibus(mission_archibus, "./missions-archibus.xlsx")


