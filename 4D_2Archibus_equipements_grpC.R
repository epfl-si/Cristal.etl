library(tidyverse)
library(readxl)
library(dplyr)
library(lubridate)
library("data.table") 
library("stringr")
options(scipen = 999)   
#setwd("/Users/venries/GitHub/Cristal")

Equi_grpeC_CHA_file = "./CHA_eq_v1.xlsx"
Equi_grpeC_SAN_file = "./SAN_FRI_CHA_eq_v1.xlsx"
Equi_CHA_installation_file = "./CHA_Installations_11.01.2023.csv"
Equi_CHA_accesoires_file = "./CHA_Accessoires_11.01.2023.csv"
Domaine_file = "./Données de référence.xlsx"
ID_grpeC = "./ID_grp_C.xlsx"

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
  csNum <- xlsx::CellStyle(wb) + xlsx::DataFormat("0000")
  csDate <- xlsx::CellStyle(wb) + xlsx::DataFormat("yyyy-mm-dd")
  colnum <- list(
    '9' = csNum
  )
  coldate <- list(
    '18' = csDate
  )
  xlsx::addDataFrame(data.frame(data, check.names = FALSE), sheet,
                     startRow = 2, row.names = FALSE,
                     colStyle=c(colnum,coldate),
                     colnamesStyle = xlsx::CellStyle(wb) +
                       xlsx::Font(wb, isBold = TRUE) +
                       xlsx::Border(color = "black", position = c("TOP", "BOTTOM"),
                                    pen = c("BORDER_THIN", "BORDER_THICK")))
  
  # save the workbook to an Excel file
  xlsx::saveWorkbook(wb, filename)
}


toArchibusStatus <- function(etat) {
  
  case_when(etat == "Faux"   ~ "in",
            etat == "FAUX"   ~ "in",
            etat == "FALSE"  ~ "in",
            etat == "Vrai"    ~ "out",
            etat == "VRAI"    ~ "out",
            etat == ""        ~ "",
            TRUE              ~ "out")
}
# STANDARS EQUIPEMENT

standards_equip <- read_excel(Domaine_file, "Standards équipement") %>%
  #  filter(Statut == "02 Standard/DT validé") %>%
  filter(!is.na(TR)) %>%
  transmute("#eqstd.eq_std" = `Standard ID (32 car)`,
            description = `Standard d'équipement`,
            category = TR,
            descr_tr = `Domaine Technique`)

standards_equip_tr <- standards_equip %>%
  distinct(category, descr_tr) %>%
  rename(`#tr.tr_id`= category) %>%
  rename(description = descr_tr)


write_archibus(standards_equip_tr, "./00.tr.xlsx",
               table.header = "Trades",
               sheet.name = "Trades")

standards_equip <- standards_equip %>%
  select(`#eqstd.eq_std`,description,category)

write_archibus(standards_equip, "./00.eqstd.xlsx",
               table.header = "Equipment Standards",
               sheet.name = "standards_equipements")

site_import <- read_excel("./Export Bâtiments.xlsx")

batiments_import <- read_excel("./rm.xlsx","Sheet1",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
  left_join(site_import, by=c("#rm.bl_id"="Building Code"))

id_grp_C <- read_excel(ID_grpeC,"Sheet1", col_names = TRUE, col_types = NULL, na = "", skip = 1)

##################################
# CHA equipement grp_C
##################################

cha_C_equip0 <- read_excel(Equi_grpeC_CHA_file, "Feuil1",col_names = TRUE, col_types = NULL, na = "", skip = 0) %>%
  mutate(`ID_Fiche_UUID` = ifelse(is.na(UUID),`ID Fiche`,paste(`ID Fiche`,`UUID`,sep = " "))) %>%
  left_join(id_grp_C,by=c("ID_Fiche_UUID" = "ID_Fiche_UUID"))
cha_C_equip0 <- cha_C_equip0 %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("CHAUF-00000-",formatC(seq.int(nrow(cha_C_equip0)) + 110000, width=6, flag=0, format="d"),sep = ""),eq_id))


export_ID <- cha_C_equip0 %>%
  transmute(eq_id  = eq_id,
            ID_Fiche_UUID = ID_Fiche_UUID)

write_archibus(export_ID, ID_grpeC,
               table.header = "ID",
               sheet.name = "Sheet1")

cha_C_install <- fread(file = Equi_CHA_installation_file  , encoding = "Latin-1")

cha_C_accesoires <- fread(file = Equi_CHA_accesoires_file  , encoding = "Latin-1")

cha_C0 <- cha_C_equip0 %>%
  filter (`Niveau` == "1-Equipement") %>%
  left_join(cha_C_install, by=c("ID Fiche"="ID Fiche", "UUID"="UUID"))

cha_C <- cha_C0 %>%
  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
  left_join(standards_equip, by=c("Standard d'équipement/d'attribut"="description")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = ifelse(is.na(`#eqstd.eq_std`), paste("A DEFINIR",`Domaine technique`,`Standard d'équipement`, sep=" "), `#eqstd.eq_std`),
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = Nom,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = "",
            subcomponent_of = "",
            mfr = "",
            asset_id = paste(`ID Fiche`,`UUID`, sep =" "),
            status = toArchibusStatus(`HS?`),
            modelno = "",
            condition = "fair",
            comments = Remarques,
            date_installed = "")
 
  
cha_C_id_parent <- cha_C_equip0 %>%
  filter (`Niveau` == "1-Equipement") %>%
  left_join(cha_C_install, by=c("ID Fiche"="ID Fiche", "UUID"="UUID")) %>%
  select(eq_id,`ID Fiche`,`Local no`)

     
  
cha_C_enfant <- cha_C_equip0 %>%
  filter (`Niveau` == "2-Accessoire" & `Import oui/non` == "Oui")  %>%
  left_join(cha_C_id_parent, by=c("ID Fiche"="ID Fiche"), suffix = c("", ".parent")) %>%
  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
  left_join(cha_C_accesoires, by=c("ID Fiche"="ID Fiche", "UUID"="UUID")) %>%
  left_join(standards_equip, by=c("Standard d'équipement/d'attribut"="description")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = ifelse(is.na(`#eqstd.eq_std`), paste("A DEFINIR",`Domaine technique`,`Standard d'équipement`, sep=" "), `#eqstd.eq_std`),
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = Description,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = Type,
            subcomponent_of = `eq_id.parent`,
            mfr = Marque,
            asset_id = paste(`ID Fiche`,`UUID`, sep =" "),
            status = "out",
            modelno = "",
            condition = "fair",
            comments = Remarques,
            date_installed = "")

cha_C_equip <- rbind(
  cha_C,
  cha_C_enfant)

write_archibus(cha_C_equip, "./01.eq-CHAUF.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")


#######
# attributs
#######

ass_attrib0 <- cha_C_equip0 %>%
  filter (`Niveau` == "2-Accessoire" & `Import oui/non` == "Oui attribut")  %>%
  group_by(`ID Fiche`) %>% 
  mutate(type_attrib = paste(`Standard d'équipement/d'attribut`, row_number(), sep=(" "))) %>%
  left_join(cha_C_accesoires, by=c("ID Fiche"="ID Fiche", "UUID"="UUID")) %>%
  left_join(cha_C_id_parent, by=c("ID Fiche"="ID Fiche")) %>%
  mutate(value_attrib = paste (Marque, Type, Remarques, sep=(", ")))
  
#######
# Type d'attribut
#######

ass_attrib2 <- ass_attrib0 %>%
  ungroup() %>%
  select(type_attrib) %>%
  distinct(type_attrib) %>%
  transmute("#asset_attribute_std.asset_attribute_std" = toupper(gsub(" ", "", type_attrib)),
            title = type_attrib,
            description = "",
            asset_type = "Equipment")

ass_attrib3 <- read.csv(text="#asset_attribute_std.asset_attribute_std,title,description,asset_type", check.names=FALSE)
ass_attrib3[nrow(ass_attrib3)+1,] <- c("LIEU","Lieu","","Equipment")
ass_attrib3[nrow(ass_attrib3)+1,] <- c("IDCONTRAT","ID contrat","","Equipment")

ass_attrib <- rbind(
  ass_attrib2,
  ass_attrib3)


write_archibus(ass_attrib, "./02.asset_attrib.xlsx",
               table.header = "Asset Attribute Standards",
               sheet.name = "Sheet1")


######################
# Valeurs des attributs
######################

eq_ass_attribut_CHA_acc <- ass_attrib0 %>%
  ungroup() %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id.y,
            "asset_attribute_std" = toupper(gsub(" ", "", type_attrib)),
            value = value_attrib)

eq_ass_attribut_CHA_lieu <- cha_C0 %>%
  filter (`Lieu` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "LIEU",
            value = `Lieu`)

eq_ass_attribut_CHA_idcontrat <- cha_C0 %>% 
  filter (`ID Contrat` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "IDCONTRAT",
            value = `ID Contrat`) 

eq_ass_attribut <- rbind(
                    eq_ass_attribut_CHA_acc,
                    eq_ass_attribut_CHA_lieu,
                    eq_ass_attribut_CHA_idcontrat)


write_archibus(eq_ass_attribut, "./03.eq_asset_attrib.xlsx",
               table.header = "Equipment Asset Attributes",
               sheet.name = "Sheet1")

