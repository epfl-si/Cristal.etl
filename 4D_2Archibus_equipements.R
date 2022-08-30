library(tidyverse)
library(readxl)
library(dplyr)
library("data.table")
#library(sqldf)

ID_file = "./2022-07-01 - Liste Equipements Cristal Test.xlsx"
Domaine_file = "./Données de référence.xlsx"
Equi_grpeA_file = "./Equipements_gpeA_final_v2.xlsx"
Equi_grpeB_file = "./Equipements_gpeB_v4.xlsx"
Equi_grpeC_file = "./Import_equipmt_22-05-23_v4.xlsx"
Venti_file = "./VEN_Installations_04.07.2022.csv"
Venti_Acc_file = "./VEN_Accessoires_27.06.2022.csv"
Elec_file  = "./ELE_Installations_04.07.2022.csv"
Mt_file = "./UTILI_Cellules MT_v03.xlsx"
Utils_file = "./UTILI_GE_air_chaleur_v02.xlsx"
FacSV_file = "./Liste_équipement_à_importer.xlsx"
GMAO_file = "./GMAO_4D_Export2021_MAPPING_eqstd.xlsx"
GMAO_Acc_file = "./4D_GMAO_Accessoires_avec UUID_4-3-22.xlsx"
ELA_Ascenseurs_file = "./LEVAG_ascenseurs_22-05-10.xlsx"
#SV_file = "./Maintenance _Equipements_INFRA_SV_4D.xlsx"
Energie_file = "./ENERGIE_compteurs_v3.xlsx"
TCVS_file = "./Import_TCVS_22-08-19.xlsx"
TCVS_Detail_file = "./TCVS_Installations_19.08.2022.xlsx"

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
            etat == "Vrai"    ~ "out",
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

id_archibus <- read_excel(ID_file, range = cell_cols("A:B"))


batiments_import <- read_excel("./06. rm.xlsx","Sheet1",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
  left_join(site_import, by=c("#rm.bl_id"="Building Code"))

##################################
# CHA equipement
##################################

# CHA Parents

#cha_equip0 <- read_excel(GMAO_file, "CHA_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1)
#cha_equip0$eq_id <-paste("CHA-00000-",formatC(seq.int(nrow(cha_equip0)), width=6, flag=0, format="d"),sep = "")

#cha_equip_parent <- cha_equip0 %>%
#  left_join(standards_equip, by=c("STANDARD D'EQUIPEMENT"="description")) %>%
#  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
#  transmute("#eq.eq_id" = eq_id,
#            eq_std = ifelse(is.na(`#eqstd.eq_std`), "A DEFINIR", `#eqstd.eq_std`),
#            bl_id = `#rm.bl_id`,
#            fl_id = fl_id,
#            rm_id = rm_id,
#            site_id = SiteCode,
#            description = Nom,
#            dv_id = 11500,
#            dp_id = "0047",
#            num_serial = "",
#            modelno = "",
#            subcomponent_of ="",
#            mfr = "",
#            asset_id = `ID Fiche`,
#            status = "in",
#            condition = "fair",
#            comments = Remarques)

# CHA Accesoires

#cha_equip_acc0 <- read_excel(GMAO_Acc_file, "C_acc",col_names = TRUE, col_types = NULL, na = "", )
#cha_equip_acc0$eq_id <-paste("CHA-00000-",formatC(seq.int(nrow(cha_equip_acc0)) + nrow(cha_equip0), width=6, flag=0, format="d"), sep = "")

#cha_equip_acc <- cha_equip_acc0 %>%
#  left_join(cha_equip_parent, by=c("IDFiche"="asset_id"),) %>%
#  transmute(subcomponent_of = `#eq.eq_id`,
#            "#eq.eq_id" = eq_id,
#            eq_std = "ACCESSOIRE",
#            bl_id = bl_id,
#            fl_id = fl_id,
#            rm_id = rm_id,
#            site_id = site_id,
#            description = AccessoireDescription,
#            dv_id = 11500,
#            dp_id = "0047",
#            num_serial = IDNo,
#            modelno = Numero,
#            mfr = Marque,
#            asset_id = paste("CHA-",UUID, sep =""),
#            status = "in",
#            condition = "fair",
#            comments = Remarques)


#cha_equip <- rbind(cha_equip_parent, cha_equip_acc)

#write_archibus(cha_equip, "./01.eq-CHA.xlsx",
#               table.header = "Equipment",
#               sheet.name = "Equipment")

##################################
# VENTI equipement
##################################


ven_valideA <- read_excel(Equi_grpeA_file, "VENTIL",col_names = TRUE, col_types = NULL, na = "") %>%
  left_join(id_archibus,by=c("ID Fiche"="asset_id")) 
ven_valideA <- ven_valideA %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("VENTI-00000-",formatC(seq.int(nrow(ven_valideA)) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

#ven_valideA$eq_id <-paste("VENTI-00000-",formatC(seq.int(nrow(ven_valideA)), width=6, flag=0, format="d"),sep = "")
ven_valideA$UUID <- NA
ven_valideA$ID_Fiche_UUID <- NA

ven_valideBVenti <- read_excel(Equi_grpeB_file, "Feuil1",col_names = TRUE, col_types = NULL, na = "") %>%
  filter (`Domaine technique` == "VENTIL") %>%
  mutate(`Domaine technique` = recode(`Domaine technique`, VENTIL = 'Ventilation' )) %>%
  mutate(`ID_Fiche_UUID` = paste(`ID Fiche`,`UUID`,sep = " ")) %>%
  left_join(id_archibus,by=c("ID_Fiche_UUID" = "asset_id"))
ven_valideBVenti <- ven_valideBVenti %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("VENTI-00000-",formatC(seq.int(nrow(ven_valideBVenti)) + nrow(ven_valideA) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

ven_valideBMobil <- read_excel(Equi_grpeB_file, "Feuil1",col_names = TRUE, col_types = NULL, na = "") %>%
  filter (`Domaine technique` == "MOBIL") %>%
  filter (`Standard d'équipement` != "") %>%
  mutate(`Domaine technique` = recode(`Domaine technique`, VENTIL = 'Mobilier labo' )) %>%
  mutate(`ID_Fiche_UUID` = paste(`ID Fiche`,`UUID`,sep = " ")) %>%
  left_join(id_archibus,by=c("ID_Fiche_UUID" = "asset_id"))
ven_valideBMobil <- ven_valideBMobil %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("MOBIL-00000-",formatC(seq.int(nrow(ven_valideBMobil)) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

#ven_valideBMobil$eq_id <-paste("MOBIL-00000-",formatC(seq.int(nrow(ven_valideBMobil)), width=6, flag=0, format="d"),sep = "")

ven_valideB <- rbind(ven_valideBVenti , ven_valideBMobil)

ven_valideBbis <- ven_valideB %>%
  left_join(ven_valideB %>%
              filter(is.na(UUID)) %>%
              transmute(`ID Fiche`, parent_eq_idB = eq_id),
            by=c("ID Fiche"="ID Fiche" )) %>%
  mutate(parent_eq_idB = ifelse(parent_eq_idB == eq_id, NA, parent_eq_idB)) %>%
  left_join(ven_valideA %>%
              transmute(`ID Fiche`, parent_eq_idA = eq_id), by=c("ID Fiche"="ID Fiche" )) %>%
  mutate(parent_eq_id = ifelse(is.na(parent_eq_idA),parent_eq_idB,parent_eq_idA)) 
ven_valideBbis <- ven_valideBbis[-c(8:9)]
  
ven_valideA$`# local`<- NA
ven_valideA$parent_eq_id <- NA

#ven_valideCVenti <- read_excel(Equi_grpeC_file, "VEN_Accessoires",col_names = TRUE, col_types = NULL, na = "") %>%
#  left_join(standards_equip,by=c("Standard d'équipement"="#eqstd.eq_std")) %>%
#  mutate(`Domaine technique` = recode(`category`, VENTILATION = 'Ventilation' ))
#ven_valideCVenti$category <- NULL
#ven_valideCVenti$"Standard d'équipement" <- NULL
#ven_valideCVenti$eq_id <-paste("VENTI-00000-",formatC(seq.int(nrow(ven_valideCVenti)) + nrow(ven_valideA) + nrow(ven_valideBVenti), width=6, flag=0, format="d"),sep = "")
#ven_valideCVenti <- ven_valideCVenti %>%
#  left_join(ven_valideB %>%
#              transmute(`ID Fiche`, parent_eq_idA = eq_id), by=c("ID Fiche"="ID Fiche" )) %>%
#  mutate(parent_eq_id = parent_eq_idA) 



ven_valide_A_B <- rbind(ven_valideA , ven_valideBbis)


ven_valideCVenti <- read_excel(Equi_grpeC_file, "VEN_Accessoires",col_names = TRUE, col_types = NULL, na = "") %>%
  left_join(standards_equip,by=c("Standard d'équipement"="#eqstd.eq_std")) %>%
  mutate(`Domaine technique` = recode(`category`, VENTILATION = 'Ventilation' ))
ven_valideCVenti$category <- NULL
ven_valideCVenti$"Standard d'équipement" <- NULL
ven_valideCVenti <- ven_valideCVenti %>% rename("Standard d'équipement"=description) %>%
  mutate(`ID_Fiche_UUID` = paste(`ID Fiche`,`UUID`,sep = " ")) %>%
  left_join(id_archibus,by=c("ID_Fiche_UUID" = "asset_id"))
ven_valideCVenti <- ven_valideCVenti %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("VENTI-00000-",formatC(seq.int(nrow(ven_valideCVenti)) + nrow(ven_valideA) + nrow(ven_valideBVenti) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

#ven_valideCVenti$eq_id <-paste("VENTI-00000-",formatC(seq.int(nrow(ven_valideCVenti)) + nrow(ven_valideA) + nrow(ven_valideBVenti), width=6, flag=0, format="d"),sep = "")
ven_valideCVenti <- ven_valideCVenti %>%
  left_join(ven_valide_A_B %>%
              filter(is.na(parent_eq_id)) %>%
              transmute(`ID Fiche`, parent_eq_id = eq_id), by=c("ID Fiche"="ID Fiche" ))
ven_valideCVenti$`# local`<- NA

ven_valide <- rbind(ven_valide_A_B , ven_valideCVenti)


#ven_valide <- ven_valide %>%
#  left_join(ven_valide, by=c("Standard d'équipement"="description"))
ven_equip0 <- fread(file = Venti_file  , encoding = "Latin-1") %>%
  rename("Débit d'air0" = "Débit d'air")

ven_equip1 <- fread(Venti_Acc_file , encoding = "Latin-1")


ven_equip_valide <- ven_valide %>%
  left_join(ven_equip0, by=c("ID Fiche"="ID Fiche")) %>%
  left_join(ven_equip1, by=c("ID Fiche"="ID Fiche", "UUID.x"="UUID"))
  
ven_equip_parent <- ven_equip_valide %>%
  left_join(standards_equip, by=c("Standard d'équipement"="description")) %>%
  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
#  left_join(batiments_import, by=c("# local"=gsub("", "","c_porte")),suffix = c("","_2")) %>%
  left_join(batiments_import, by=c("# local"="c_porte"),suffix = c("","_2")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = ifelse(is.na(`#eqstd.eq_std`), "A DEFINIR", `#eqstd.eq_std`),
            bl_id = ifelse(is.na(`# local`),`#rm.bl_id`, `#rm.bl_id_2`),
            fl_id = ifelse(is.na(`# local`),`fl_id`, `fl_id_2`),
            rm_id = ifelse(is.na(`# local`),`rm_id`, `rm_id_2`),
            site_id = SiteCode,
            description = ifelse(is.na(parent_eq_id),Nom,Description),
            dv_id = 11500,
            dp_id = "0047",
            num_serial = "",
            subcomponent_of = parent_eq_id,
            mfr = ifelse(is.na(parent_eq_id),Marque,""),
            asset_id = ifelse(is.na(`UUID.x`),`ID Fiche`,paste(`ID Fiche`,`UUID.x`, sep =" ")),
            status = toArchibusStatus(`HS?`),
            modelno = ifelse(is.na(parent_eq_id),`Monobloc No`,""),
            condition = "fair",
            comments = ifelse(is.na(parent_eq_id),Remarques.x,paste(Remarques.y,"\n",Intervention,sep="")),
            date_installed = ifelse(is.na(parent_eq_id),`Mise en service.x`,`Mise en service.y`)) %>%
  mutate(date_installed = as.Date(replace(date_installed, date_installed == "00.00.00", NA),"%d.%m.%Y"))

#write_archibus(ven_equip_valide, "./000.tmp.xlsx",
#               table.header = "Equipment",
#               sheet.name = "Equipment")
write_archibus(ven_equip_parent, "./01.eq-VENTI.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")

##################################
# SAN equipement
##################################

# SAN Parents

#san_equip0 <- read_excel(GMAO_file, "SAN_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1)
#san_equip0$eq_id <-paste("SAN-00000-",formatC(seq.int(nrow(san_equip0)), width=6, flag=0, format="d"),sep = "")

#san_equip_parent <- san_equip0 %>%
#  left_join(standards_equip, by=c("STANDARD D'EQUIPEMENT"="description")) %>%
#  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
#  transmute("#eq.eq_id" = eq_id,
#            eq_std = ifelse(is.na(`#eqstd.eq_std`), "A DEFINIR", `#eqstd.eq_std`),
#            bl_id = `#rm.bl_id`,
#            fl_id = fl_id,
#            rm_id = rm_id,
#            site_id = SiteCode,
#            description = Nom,
#            dv_id = 11500,
#            dp_id = "0047",
#            num_serial = "",
#            subcomponent_of ="",
#            mfr = "",
#            asset_id = `ID Fiche`,
#            status = "in",
#            condition = "fair",
            #            date_installed = `Mise en service`,
#            comments = Remarques)

# SAN Accesoires
#san_equip_acc0 <- read_excel(GMAO_Acc_file, "S_acc",col_names = TRUE, col_types = NULL, na = "", )
#san_equip_acc0$eq_id <-paste("SAN-00000-",formatC(seq.int(nrow(san_equip_acc0)) + nrow(san_equip0), width=6, flag=0, format="d"), sep = "")

#san_equip_acc <- san_equip_acc0 %>%
#  left_join(san_equip_parent, by=c("IDFiche"="asset_id")) %>%
#  transmute(subcomponent_of = `#eq.eq_id`,
#            "#eq.eq_id" = eq_id,
#            eq_std = "ACCESSOIRE",
#            bl_id = bl_id,
#            fl_id = fl_id,
#            rm_id = rm_id,
#            site_id = site_id,
#            description = AccessoireDescription,
#            dv_id = 11500,
#            dp_id = "0047",
#            num_serial = Numero,
#            mfr = Marque,
#            asset_id = paste("SAN-",UUID, sep =""),
#            status = "in",
#            condition = "fair",
#            #           date_installed = "",
#            comments = Remarques)


#san_equip <- rbind(san_equip_parent, san_equip_acc)

#write_archibus(san_equip, "./01.eq-SAN.xlsx",
#               table.header = "Equipment",
#               sheet.name = "Equipment")


#############################################

# LEVAGE Ascenseurs

#############################################


leva_equip0 <- read_excel(ELA_Ascenseurs_file, "Inventaire ascenseurs",col_names = TRUE, col_types = NULL, na = "", skip = 1)  %>%
  mutate(`ID_Fiche_UUID` = paste(`ID Fiche (table Electricité)`,`UUID (table Acsenseurs)`,sep = " ")) %>%
  left_join(id_archibus,by=c("ID_Fiche_UUID" = "asset_id"))
leva_equip0 <- leva_equip0 %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("LEVAG-00000-",formatC(seq.int(nrow(leva_equip0)) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))


#leva_equip0$eq_id <-paste("LEVAG-00000-",formatC(seq.int(nrow(leva_equip0)), width=6, flag=0, format="d"),sep = "")

leva_equip <- leva_equip0 %>%
  left_join(standards_equip, by=c("Standard d'équipement"="description")) %>%
  left_join(batiments_import, by=c("# Local"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = `#eqstd.eq_std`,
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = Description,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = "",
            modelno = `N° du Modèle`,
            subcomponent_of ="",
            mfr = Fabricant,
            asset_id = paste(`ID Fiche (table Electricité)`,`UUID (table Acsenseurs)`,sep =" - "),
            status = "in",
            condition = "fair",
            comments = "",
            date_installed = `Date d’Installation`) 
#  mutate(date_installed = as.Date(replace(date_installed, date_installed == "00.00.00", NA),"%Y-%m.%d"))

############################################
# LEVAGE AUTRE
############################################


levag_valide <- read_excel(Equi_grpeA_file, "ELECT-LEVAG",col_names = TRUE, col_types = NULL, na = "") %>%
  filter (`Domaine technique` == "Levage") %>%
  left_join(id_archibus,by=c("ID Fiche"="asset_id"))
levag_valide <- levag_valide %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("LEVAG-00000-",formatC(seq.int(nrow(levag_valide)) + nrow(leva_equip0) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))



#levag_valide$eq_id <-paste("LEVAG-00000-",formatC(seq.int(nrow(levag_valide)) + nrow(leva_equip0), width=6, flag=0, format="d"),sep = "")
Elec_equip0 <- fread(file = Elec_file  , encoding = "Latin-1") 
  
levag_equip_valide <- levag_valide %>%
  left_join(Elec_equip0, by=c("ID Fiche"="ID Fiche"))
levag_equip2 <- levag_equip_valide %>%
  left_join(standards_equip, by=c("Standard d'équipement"="description")) %>%
  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = ifelse(is.na(`#eqstd.eq_std`), "A DEFINIR", `#eqstd.eq_std`),
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = Nom,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = `No de série`,
            modelno = `Installation no`,
            subcomponent_of ="",
            mfr = Fournisseur,
            asset_id = `ID Fiche`,
            status = toArchibusStatus(`HS?`),
            condition = "fair",
            comments = Remarques,
            date_installed = `Mise en service`) %>%
  mutate(date_installed = as.Date(replace(date_installed, date_installed == "00.00.00", NA),"%d.%m.%Y"))

levag <- rbind(leva_equip, levag_equip2)

write_archibus(levag, "./01.eq-LEVAG.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")

##################################
# ELE equipement
##################################

# ELE Parents

elect_valide <- read_excel(Equi_grpeA_file, "ELECT-LEVAG",col_names = TRUE, col_types = NULL, na = "") %>%
  filter (`Domaine technique` != "Levage") %>%
  left_join(id_archibus,by=c("ID Fiche"="asset_id")) %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("ELECT-00000-",formatC(seq.int(nrow(elect_valide)) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

  
#elect_valide$eq_id <-paste("ELECT-00000-",formatC(seq.int(nrow(elect_valide)), width=6, flag=0, format="d"),sep = "")


ele_equip_valide <- elect_valide %>%
  left_join(Elec_equip0, by=c("ID Fiche"="ID Fiche"))

ele_equip_parent <- ele_equip_valide %>%
  left_join(standards_equip, by=c("Standard d'équipement"="description")) %>%
  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = ifelse(is.na(`#eqstd.eq_std`), "A DEFINIR", `#eqstd.eq_std`),
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = Nom,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = `No de série`,
            modelno = `Modèle`,
            subcomponent_of ="",
            mfr = `Fournisseur`,
            asset_id = `ID Fiche`,
            status = toArchibusStatus(`HS?`),
            condition = "fair",
            comments = Remarques,
            date_installed = `Mise en service`) %>%
  mutate(date_installed = as.Date(replace(date_installed, date_installed == "00.00.00", NA),"%d.%m.%Y"))


elect2 <- read_excel(Equi_grpeC_file, "ELE_Installations",col_names = TRUE, col_types = NULL, na = "") %>%
  filter (`Standard d'équipement` != "TGBT") %>%
  mutate("Composant de :" = as.character(`Composant de :`)) %>%
  mutate(`ID_Fiche_UUID` = paste(`ID Fiche`,`UUID`,sep = " ")) %>%
  left_join(id_archibus,by=c("ID_Fiche_UUID" = "asset_id"))
elect2 <- elect2 %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("ELECT-00000-",formatC(seq.int(nrow(elect2)) + nrow(elect_valide) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

elect2$category <- NULL
#elect2$eq_id <-paste("ELECT-00000-",formatC(seq.int(nrow(elect2)) + nrow(elect_valide), width=6, flag=0, format="d"),sep = "")
elect2 <- elect2 %>%
  left_join(elect2 %>%
              transmute(`ID Fiche`,parent_eq_id = eq_id), by=c("Composant de :"="ID Fiche"))


ele_equip_elect2 <- elect2 %>%
  left_join(Elec_equip0, by=c("ID Fiche"="ID Fiche")) %>%
  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = `Standard d'équipement`,
            bl_id = ifelse(is.na(`# Bâtiment`),`#rm.bl_id`,`# Bâtiment`),
            fl_id = ifelse(is.na(`# Bâtiment`),fl_id,""),
            rm_id = ifelse(is.na(`# Bâtiment`),rm_id,""),
            site_id = ifelse(is.na(`# Bâtiment`),SiteCode,"E"),
            description = Nom,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = `No de série`,
            modelno = `Modèle`,
            subcomponent_of = parent_eq_id,
            mfr = `Fournisseur`,
            asset_id = `ID Fiche`,
            status = toArchibusStatus(`HS?`),
            condition = "fair",
            comments = Remarques,
            date_installed ="")

ele_equip <- rbind(ele_equip_parent, ele_equip_elect2)

write_archibus(ele_equip, "./01.eq-ELECT.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")



#################################################3
# ELE Accesoires

#ele_equip_acc0 <- read_excel(GMAO_Acc_file, "E_acc",col_names = TRUE, col_types = NULL, na = "", ) %>%
#  inner_join(ele_equip_parent, by=c("IDFiche"="asset_id"),) 
#ele_equip_acc0$eq_id2 <-paste("ELE-00000-",formatC(seq.int(nrow(ele_equip_acc0)) + nrow(ele_ass_equip) + nrow(ele_equip0), width=6, flag=0, format="d"), sep = "")

#ele_equip_acc <- ele_equip_acc0 %>%
#  transmute(subcomponent_of = `#eq.eq_id`,
#            "#eq.eq_id" = eq_id2,
#            eq_std = "ACCESSOIRE",
#            bl_id = bl_id,
#            fl_id = fl_id,
#            rm_id = rm_id,
#            site_id = site_id,
#            description = AccessoireDescription,
#            dv_id = 11500,
#            dp_id = "0047",
#            num_serial = NoDeSerie,
#            modelno = "Numero",
#            mfr = Fournisseur,
#            asset_id = paste("ELE-",UUID, sep =""),
#            status = "in",
#            condition = "fair",
#            comments = Remarques)


#ele_equip <- rbind(ele_equip_parent, ele_equip_acc)

#write_archibus(ele_equip, "./01.eq-ELE.xlsx",
#               table.header = "Equipment",
#               sheet.name = "Equipment")




#############################################

# UTILS

#############################################


mt_equip0 <- read_excel(Mt_file, "Cellules MT v3",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
  left_join(id_archibus,by=c("ID"="asset_id"))
mt_equip0 <- mt_equip0 %>% 
  mutate(eq_id = ifelse(is.na(eq_id),paste("UTILI-00000-",formatC(seq.int(nrow(mt_equip0)) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

#mt_equip0$eq_id <-paste("UTILI-00000-",formatC(seq.int(nrow(mt_equip0)), width=6, flag=0, format="d"),sep = "")

mt_equip0bis <- mt_equip0 %>%
  left_join(mt_equip0 %>%
              transmute(`ID`, parent_mt = eq_id),
            by=c("Composant de l’équipement :"="ID" )) %>%
  mutate(parent_mt = ifelse(parent_mt == eq_id, NA, parent_mt)) 


mt_equip <- mt_equip0bis %>%
  left_join(standards_equip, by=c("Standard d'équipement"="description")) %>%
  left_join(batiments_import, by=c("Code porte"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = ifelse(is.na(`#eqstd.eq_std`), "A DEFINIR", `#eqstd.eq_std`),
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = `Description de l'équipement`,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = `Numéro de série`,
            modelno = "",
            subcomponent_of = parent_mt,
            mfr = Fabricant,
            asset_id = ID,
            status = "in",
            condition = "fair",
            comments = "",
            date_installed = `Date d’Installation`)


util1 <- read_excel(Utils_file, "GES, citerne",col_names = TRUE, col_types = NULL, na = "", skip = 2) %>%
  left_join(id_archibus,by=c("ID"="asset_id"))
util1 <- util1 %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("UTILI-00000-",formatC(seq.int(nrow(mt_equip0)) + nrow(mt_equip0) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

#util1$eq_id <-paste("UTILI-00000-",formatC(seq.int(nrow(util1)) + nrow(mt_equip0), width=6, flag=0, format="d"),sep = "")
util1 <- util1 %>%
  left_join(util1 %>%
              transmute(`ID`,parent_eq_id = eq_id), by=c("Composant de l'équipement :"="ID"))


util1_equip <- util1 %>%
  left_join(batiments_import, by=c("# Local"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = `Standard d'équipement`,
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = `Description de l’équipement`,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = `Numéro de série`,
            modelno = "",
            subcomponent_of = parent_eq_id,
            mfr = Fabricant,
            asset_id = ID,
            status = "in",
            condition = "fair",
            comments = "",
            date_installed = `Date d’Installation`)

util2 <- read_excel(Utils_file, "Compresseur air",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
  left_join(id_archibus,by=c("ID / UUID"="asset_id"))
util2 <- util2 %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("UTILI-00000-",formatC(seq.int(nrow(mt_equip0)) + nrow(mt_equip0) + nrow(util1) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))


#util2$eq_id <-paste("UTILI-00000-",formatC(seq.int(nrow(util2)) + nrow(mt_equip0) + nrow(util1), width=6, flag=0, format="d"),sep = "")



util2_equip <- util2 %>%
  left_join(batiments_import, by=c("# Local"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = `Standard d'équipement`,
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = `Modèle`,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = `Numéro de série`,
            modelno = "",
            subcomponent_of = "",
            mfr = Fabricant,
            asset_id = `ID / UUID`,
            status = "in",
            condition = "fair",
            comments = "",
            date_installed = `Date d’Installation`)

util3 <- read_excel(Utils_file, "Chaudière gaz",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
  left_join(id_archibus,by=c("ID / UUID"="asset_id"))
util3 <- util3 %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("UTILI-00000-",formatC(seq.int(nrow(mt_equip0)) + nrow(mt_equip0) + nrow(util1) + nrow(util2) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

#util3$eq_id <-paste("UTILI-00000-",formatC(seq.int(nrow(util3)) + nrow(mt_equip0) + nrow(util1) + nrow(util2), width=6, flag=0, format="d"),sep = "")

util3_equip <- util3 %>%
  left_join(batiments_import, by=c("# Local"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = `Standard d'équipement`,
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = `Modèle`,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = `Numéro de série`,
            modelno = "",
            subcomponent_of = "",
            mfr = Fabricant,
            asset_id = `ID / UUID`,
            status = "in",
            condition = "fair",
            comments = "",
            date_installed = `Date d’Installation`)

util4 <- read_excel(Equi_grpeC_file, "ELE_Installations",col_names = TRUE, col_types = NULL, na = "") %>%
  filter (`Standard d'équipement` == "TGBT") %>%
  mutate("Composant de :" = as.character(`Composant de :`)) %>%
  mutate(`ID_Fiche_UUID` = paste(`ID Fiche`,`UUID`,sep = " ")) %>%
  left_join(id_archibus,by=c("ID_Fiche_UUID" = "asset_id"))
util4 <- util4 %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste("UTILI-00000-",formatC(seq.int(nrow(util4)) + nrow(mt_equip0) + nrow(util1) + nrow(util2) + nrow(util4) + 10000, width=6, flag=0, format="d"),sep = ""),eq_id))

util4$category <- "NULL"
#util4$eq_id <-paste("UTILI-00000-",formatC(seq.int(nrow(util4)) + nrow(mt_equip0) + nrow(util1) + nrow(util2) + nrow(util4), width=6, flag=0, format="d"),sep = "")


util4_equip <- util4 %>%
  left_join(Elec_equip0, by=c("ID Fiche"="ID Fiche")) %>%
  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = `Standard d'équipement`,
            bl_id = ifelse(is.na(`# Bâtiment`),`#rm.bl_id`,`# Bâtiment`),
            fl_id = ifelse(is.na(`# Bâtiment`),fl_id,""),
            rm_id = ifelse(is.na(`# Bâtiment`),rm_id,""),
            site_id = ifelse(is.na(`# Bâtiment`),SiteCode,"E"),
            description = Nom,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = `No de série`,
            modelno = `Modèle`,
            subcomponent_of = "",
            mfr = "",
            asset_id = `ID Fiche`,
            status = toArchibusStatus(`HS?`),
            condition = "fair",
            comments = Remarques,
            date_installed = `Mise en service`) %>%
  mutate(date_installed = as.Date(replace(date_installed, date_installed == "00.00.00", NA),"%d.%m.%Y"))



util=rbind(mt_equip, util1_equip, util2_equip, util3_equip, util4_equip)

write_archibus(util, "./01.eq-UTILI.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")



##################################
# MOB equipement
##################################

# MOB Parents

#mob_equip0 <- read_excel(GMAO_file, "MobLABO_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
#  filter(grepl('ok',Remarques,ignore.case=TRUE))

#mob_equip0$eq_id <-paste("MOB-00000-",formatC(seq.int(nrow(mob_equip0)), width=6, flag=0, format="d"),sep = "")

#mob_equip_parent <- mob_equip0 %>%
#  left_join(standards_equip, by=c("STANDARD D'EQUIPEMENT"="description")) %>%
#  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
#  transmute("#eq.eq_id" = eq_id,
#            eq_std = ifelse(is.na(`#eqstd.eq_std`), "A DEFINIR", `#eqstd.eq_std`),
#            bl_id = `#rm.bl_id`,
#            fl_id = fl_id,
#            rm_id = rm_id,
#            site_id = SiteCode,
#            description = Nom,
#            dv_id = 11500,
#            dp_id = "0047",
#            num_serial = `No de série`,
#            subcomponent_of ="",
#            mfr = Fournisseur,
#            asset_id = `ID Fiche`,
#            status = "in",
#            condition = "fair",
#            #            date_installed = `Mise en service`,
# comments = Remarques)

# MOB Accesoires
#mob_equip_acc0 <- read_excel(GMAO_Acc_file, "LaboM_Acc",col_names = TRUE, col_types = NULL, na = "", )
#mob_equip_acc0$eq_id <-paste("MOB-00000-",formatC(seq.int(nrow(mob_equip_acc0)) + nrow(mob_equip0), width=6, flag=0, format="d"), sep = "")

#mob_equip_acc <- mob_equip_acc0 %>%
#  left_join(mob_equip_parent, by=c("IDFiche"="asset_id")) %>%
#  transmute(subcomponent_of = `#eq.eq_id`,
#            "#eq.eq_id" = eq_id,
#            eq_std = "ACCESSOIRE",
#            bl_id = bl_id,
#            fl_id = fl_id,
#            rm_id = rm_id,
#            site_id = site_id,
#            description = AccessoireDescription,
#            dv_id = 11500,
#            dp_id = "0047",
#            num_serial = "",
#            mfr = "",
#            asset_id = paste("MOB-",UUID, sep =""),
#            status = "in",
#            condition = "fair",
            #           date_installed = "",
#            comments = Remarques)


#mob_equip <- rbind(mob_equip_parent, mob_equip_acc)

#write_archibus(mob_equip, "./01.eq-MOB.xlsx",
#               table.header = "Equipment",
#               sheet.name = "Equipment")

##################################
# SV equipement
##################################

sv_4d <-read_excel(FacSV_file, "ARCHIBUS 4D",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
  mutate(CF4 = as.double(CF4))

sv_equip <- read_excel(FacSV_file, "ARCHIBUS FSV",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
#  mutate(`Code équipement` = sub("FSV", "FSVIE", `Code équipement`)) %>%
#  mutate(`Composant équipement` = sub("FSV", "FSVIE", `Composant équipement`)) %>%
  rename("Contact du Labo" = "Contact 1 du Labo") %>% 
  mutate(CF4 = as.double(CF4))

sti_equip <- read_excel(FacSV_file, "ARCHIBUS FSTI",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
#  mutate(`Code équipement` = sub("FSV", "FSTI", `Code équipement`)) %>%
  mutate(CF4 = as.double(CF4))

sv_sti_equip = rbind(sv_4d,sv_equip,sti_equip)

#dp <- read_excel("./export DP.xlsx")

sv_equip_parent <- sv_sti_equip %>%
#  left_join(standards_equip, by=c("Standard ID"="#eqstd.eq_std")) %>%
#  left_join(batiments_import, by=c("Local"="c_porte")) %>%
#  left_join(dp, by=c(CF4="dp_id")) %>%
  transmute("#eq.eq_id" = `Code équipement`,
            eq_std = `Standard ID`,
            bl_id = `# Bâtiment`,
            fl_id = `# Etage`,
            rm_id = `# Local`,
            site_id = "E",
            description = `Description de l'équipement`,
            dv_id = CF2,
            dp_id = CF4,
            num_serial = `Numéro de série`,
            subcomponent_of = `Composant équipement`,
            mfr = Fabricant,
            asset_id = `N°du modèle`,
            modelno = `N°du modèle`,  
            status = ifelse(is.na(`Statut de l'Equipement (champ à rajouter)*`), "out", "in"),
            condition = "fair",
            #            date_installed = `Mise en service`,
            comments = ifelse( is.na(`Commentaires 1`) ,"", ifelse(is.na(`Commentaires 2`) ,`Commentaires 1`,paste(`Commentaires 1`,`Commentaires 2`,sep =" - "))),
            cost_purchase =  ifelse(is.na(`Prix d'Achat`),"",round(as.numeric(`Prix d'Achat`), digits = 2)),
            date_purchased = ifelse(is.na(`Date d'achat`),"",format(as.Date(`Date d'achat`,"%d.%m.%Y"), format="%Y-%m-%d")))

write_archibus(sv_equip_parent, "./01.eq-FACSV.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")


##################################
# Energie
##################################

energie <-read_excel(Energie_file, "Table compteurs_v3",col_names = TRUE, col_types = NULL, na = "", skip = 1)
  
energie_equip <- energie %>%
  left_join(batiments_import, by=c("Local"="c_porte")) %>%
  transmute("#eq.eq_id" = `ID_ARCHIBUS`,
            eq_std = `Standard d'équipement`,
            bl_id = ifelse(is.na(`#rm.bl_id`),`Bâtiment`,`#rm.bl_id`),
            fl_id = ifelse(is.na(`fl_id`),"",fl_id),
            rm_id = ifelse(is.na(`rm_id`),"",rm_id),
            site_id = ifelse(is.na(`SiteCode`),"E",SiteCode),
            description = `Description de l’équipement`,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = `Numéro de série`,
            subcomponent_of = "",
            mfr = Fabricant,
            asset_id = `ID_ARCHIBUS`,
            modelno = "",  
            status = "in",
            condition = "fair",
            #            date_installed = `Mise en service`,
            comments = ifelse(is.na(`Commentaires`) ,"", Commentaires),
            cost_purchase =  "",
            date_purchased = "")

write_archibus(energie_equip, "./01.eq-ENERG.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")

###################################
# TCVS
###################################

prefixe_domaine <- function(domaine) {
  case_when(domaine == "VENTILATION"  ~ "VENTI",
            domaine == "CHAUFFAGE"    ~ "CHAUF",
            domaine == "ELECTRICITE"  ~ "ELECT",
            domaine == "UTILITES"     ~ "UTILI")
}

tcvs <- read_excel(TCVS_file, "TCVS_Installations",col_names = TRUE, col_types = NULL, na = "")  %>%
  mutate(`ID_Fiche_UUID` = paste(`ID Fiche`,`UUID`,sep = " ")) %>%
  left_join(id_archibus,by=c("ID_Fiche_UUID" = "asset_id"))

tcvs <- tcvs %>%
  mutate(eq_id = ifelse(is.na(eq_id),paste(prefixe_domaine(`Domaine technique`),"-00000-",formatC(seq.int(nrow(tcvs)) + 19000, width=6, flag=0, format="d"),sep = ""),eq_id))

tcvs_detail <- read_excel(TCVS_Detail_file, "TCVS_Installations_19.08.2022",col_names = TRUE, col_types = NULL, na = "") %>%
  left_join(tcvs, by=c("ID Fiche"="ID Fiche", "UUID"="UUID"))

#leva_equip0$eq_id <-paste("LEVAG-00000-",formatC(seq.int(nrow(leva_equip0)), width=6, flag=0, format="d"),sep = "")

tcvs_equip <- tcvs_detail %>%
  left_join(batiments_import, by=c("Local no"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = `Standard d'équipement`,
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = Nom,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = "",
            modelno = "",
            subcomponent_of ="",
            mfr = "",
            asset_id = paste(`ID Fiche`,`UUID`,sep =" - "),
            status = toArchibusStatus(`HS?`),
            condition = "fair",
            comments = Remarques,
            date_installed = `Mise en service`) 

write_archibus(tcvs_equip, "./01.eq-TCVS.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")

##################################
# Attributs
##################################


ass_attrib <- read.csv(text="#asset_attribute_std.asset_attribute_std,title,description,asset_type", check.names=FALSE)
ass_attrib[nrow(ass_attrib)+1,] <- c("LIEU","Lieu","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TYPE","Type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIAMETRE","Diametre","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VALEURHORAIRE","Valeur horaire","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("INSTALLATION","Installation","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MONOBLOCTYPE","Monobloc type","","Equipment")
#ass_attrib[nrow(ass_attrib)+1,] <- c("MONOBLOCNO","Monobloc no","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DEBITAIR","Débit d'air","m3/h","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PRESSIONEXTRACTION","Pression extraction","bar","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PRESSIONPULSION","Pression pulsion","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MOTEURTYPE","Moteur type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MOTEURTENSION","Moteur Tension","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PUISSANCE","Puissance","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MOTEURNBTOURS","Moteur nb tours","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MOTEURNBVITESSE","Moteur nb vitesse","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MOTEURAMPERAGE","Moteur ampérage","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VENTILATEURTYPE","Ventilateur type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PRESSIONSTATIQUE","Pression statique","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSMISSIONTYPE","Transmission type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("POSITION","Position","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MATIERE","Matière","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("COURROIEFORME","Courroie forme","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("COURROIETYPE","Courroie type","","Equipment")
#ass_attrib[nrow(ass_attrib)+1,] <- c("PRISEAIREDERNIERNET","Prise Aire Dernier nettoyage","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("CLIMATISEURARMOIRE","Climatiseur Armoire type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RACCORDEMENT","Raccordement","","Equipment")
#ass_attrib[nrow(ass_attrib)+1,] <- c("FICHETECHNIQUE","Fiche technique","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MODELE","Modèle","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("IDCONTRAT","ID contrat","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("INFOSCONTRAT","Infos contrat","","Equipment")
#ass_attrib[nrow(ass_attrib)+1,] <- c("FREQUENCE", "Fréquence","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NOESTI","No ESTI","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MARQUE","Marque","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NOSURPLAN","No sur plan","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("INFOSMES","Infos MES","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIMENSIONSLPH","Dimensions L-P-H","cm","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIMENSION","Dimension","cm","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIAMETRERACCORDEMENT","Diametre raccordement","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VITESSEAIRPREVUE","Vitesse air prévue","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VITESSAIRMESUREE","Vitesse air mesurée","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MANOMETREPOSEOUI","Manometre pose oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MANOMETREPRESSION","Manometre pression","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("GUILLOTINECHAPELLE","Guillotine chapelle","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("HORLOGEPVGVOUI","Horloge PVGV oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("FORSAGEGVOUI","Forsage GV oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ELEMPVOUI","Elem PV oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ELEMGVOUI","Elem GV oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ELEMAUTOOUI","Elem AUTO oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ELEMHORSOUI","Elem HORS oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ELEMPANNEOUI","Elem Panne oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ELEMNOPLATINE","Elem No platine","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIMENSION2","Dimension2","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIMENSION3","Dimension3","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ACONTROLEROUI","A controler oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("IDNO","ID No","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("QUANTITE","Quantité","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("STOCK","Stock","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ETAT","Etat","","Equipment")
#ass_attrib[nrow(ass_attrib)+1,] <- c("INTERVENTION","Intervention","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MAJPAR","MAJ par","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("CTRL1","Ctrl1","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("CHAPELLEOUI","Chapelle Oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("IDDEPANNGE","ID Dépannge","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DEBITDIFFERENCEOUI","Débit difference oui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DEBITAIRMESURE","Débit air mesuré","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DEBITNOM","Débit nom","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DEBITEFFECTIF","Débit effectif","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("POSITION","Position","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NBREHEURE","NbreHeure","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NBREPERSONNES","NbrePersonnes","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ANNEEENSERVICE","AnneeEnService","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ANNEECONSTRUCTION","AnneeConstruction","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIFOURNISSEUR","RelaiFournisseur","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("AUTOTRANSFOOUI","AutotransfoOui","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFORESEAU","TransfoReseau","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOCELLULE","TransfoCellule","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAITYPE","RelaiType","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("UNOMINALE","UNominale","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOUPRIMAIRE","TransfoUPrimaire","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOUSECONDAIREREGLEE","TransfoUSecondaireReglee","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOREMPLISSAGE","TransfoRemplissage","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOCOUPLAGE","TransfoCouplage","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOUCC","TransfoUcc","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("UPRIMAIRE","UPrimaire","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOUSECONDAIRE1","TransfoUSecondaire1","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOUSECONDAIRE2","TransfoUSecondaire2","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOUSECONDAIRE3","TransfoUSecondaire3","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOUSECONDAIRE4","TransfoUSecondaire4","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOUSECONDAIRE5","TransfoUSecondaire5","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOPERTEAVIDE","TransfoPerteAVide","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOPERTEENCHARGE","TransfoPerteEnCharge","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOMASSETOTALE","TransfoMasseTotale","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOMADDEHUILE","TransfoMaddeHuile","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOPREALARME","TransfoPreAlarme","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFOALARME","TransfoAlarme","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DISJONCTEURINOMINAL","DisjoncteurINominal","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DISJONCTEURPOUVOIRDECOUPURE","DisjoncteurPouvoirDeCoupure","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIITHERNMIQUEREGLE","RelaiIThernmiqueRegle","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIDECLANCHEINSTANTANE","RelaiDeclancheInstantane","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAICONSTANTEDETEMPS","RelaiConstanteDeTemps","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIIREPONSE","RelaiIReponse","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAITEMPORISATION","RelaiTemporisation","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIINOMINAL","RelaiINominal","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("CHARGEUTILE","Charge utile","t","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("CHARGEUTILEKG","Charge utile","kg","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TELASCENSEUR","No téléphone ascenseur","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("SYSTEMEURGENCE","Système d'appel d'urgence","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("IMAGE","Image Source","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ORIGINALE","Originale","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ORIGINALE","Originale","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PRIX","Prix","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIAMETRERACCORDEMENT","Diamètre raccordement","mm","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MANOMETREPOSE?","Manomètre Pose?","","Equipment")
#ass_attrib[nrow(ass_attrib)+1,] <- c("INTERVENTION","Intervention","","Equipment")

ass_attrib[nrow(ass_attrib)+1,] <- c("NIVEAUSECU","Niveau Sécurité","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("FOURNISSEUR","Fournisseur","","Equipment")
#ass_attrib[nrow(ass_attrib)+1,] <- c("SIGLELABO","Sigle du Labo","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("CONTACTLABO","Contact du Labo","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("SF6","SF6 (kg)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NOMBREDEMANŒUVRESMAX","Nombre de manœuvres max","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANFORMATEURINTENSITE","Tranformateur d'intensité","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PUISSANCE","Puissance (VA)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("CLASSEDEPRECISION","Classe de précision","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSFORMATEURPOTENTIEL","Transformateur de potentiel","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ALPHAE","Alpha E oui/non","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ALPHAEITRIP","Alpha E I trip (A)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ALPHAETRESET","Alpha E t reset (H)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ALPHAERELAY","Alpha E Relay","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIS-SEUILI","Relais - Seuil I>","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIS-TDECL","Relais - Tdécl","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIS-SEUILI","Relais - Seuil I>>","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIS-TDECL2","Relais - Tdécl2","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIS-SEUILITH","Relais - Seuil Ith","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RELAIS-TDECL3","Relais - Tdécl3","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PUISSANCE","Puissance","kW","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("COUPLAGE","Couplage","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RAPPORTTRANSFORMATION","Rapport de transformation","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PLOT1","Plot 1","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PLOT2","Plot 2","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PLOT3","Plot 3","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PLOT4","Plot 4","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PLOT5","Plot 5","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NATUREDESENROULEMENTS","Nature des enroulements","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("COURANTNOMINALBT","Courant nominal BT (A)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TENSIONDECOURTCIRCUIT","Tension de court circuit (%)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MASSETOTALE","Masse totale (t)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MASSEHUILE","Masse huile (t)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PERTESENCHARGE","Pertes en charge (W)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PERTESAVIDE","Pertes à vide (W)","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TEMP.ALARME","Temp. alarme","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TEMP.DECLENCHEMENT","Temp. déclenchement","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NOSAP","N°SAP","","Equipment")
#ass_attrib[nrow(ass_attrib)+1,] <- c("DOCUMENT","Document","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NOPOD","Numéro POD compteur","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DNEAU","DN compteur eau","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("CODEUREAU","Codeur compteur eau","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("IPCOMPTEUR","Adresse IP compteur","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MODELECOMPTEUR","Modèle compteur","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("REVLEVECOMPTEUR","Relevé compteur","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TYPECOMPTEUR","Type Compteur","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VERSIONFIRMWARE","Version firmware","","Equipment")


write_archibus(ass_attrib, "./02.asset_attrib.xlsx",
               table.header = "Asset Attribute Standards",
               sheet.name = "Sheet1")

# Attributs CHA

#eq_ass_attribut_lieu <- cha_equip0 %>%
#  filter (Lieu != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "LIEU",
#            value = Lieu)

# Attributs CHA Accessoire

#eq_ass_attribut_cha_type <- cha_equip_acc0 %>%
#  filter (Type != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TYPE",
#            value = Type)

#eq_ass_attribut_diametre <- cha_equip_acc0 %>%
#  filter (Diametre != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DIAMETRE",
#            value = Diametre)

#eq_ass_attribut_valeurhoraire <- cha_equip_acc0 %>%
#  filter (ValeurHoraire != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "VALEURHORAIRE",
#            value = ValeurHoraire)

# Attributs VEN

eq_ass_attribut_installation <- ven_equip_valide %>%
  filter (`Installation` != "") %>%
  filter (is.na(`parent_eq_id`)) %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "INSTALLATION",
            value = `Installation`)

eq_ass_attribut_monobloctype <- ven_equip_valide %>%
  filter (`Monobloc Type` != "") %>%
  filter (is.na(`parent_eq_id`)) %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MONOBLOCTYPE",
            value = `Monobloc Type`)

#eq_ass_attribut_monoblocno <- ven_equip_valide %>%
#  filter (`Monobloc No` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "MONOBLOCNO",
#            value = `Monobloc No`) 

eq_ass_attribut_debitair <- ven_equip_valide %>%
  filter (`Débit d'air` != "") %>%
  filter (`Débit d'air` != "0") %>%
  filter (is.na(`parent_eq_id`)) %>%
  #filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DEBITAIR",
            value = `Débit d'air`) 

eq_ass_attribut_debitair0 <- ven_equip_valide %>%
  filter (`Débit d'air0` != "") %>%
  filter (`Débit d'air0` != "0") %>%
  filter (is.na(`parent_eq_id`)) %>%
  #filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DEBITAIR",
            value = `Débit d'air0`) 

eq_ass_attribut_ven_type <- ven_equip_valide %>%
  filter (`Type` != "") %>%
  #filter (!is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = `Type`)

eq_ass_attribut_ven_dimension <- ven_equip_valide %>%
  filter (`Dimension1` != "") %>%
  filter (`Dimension1` != "0") %>%
  #filter (!is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIMENSION",
            value = `Dimension1`)

eq_ass_attribut_manometrepose <- ven_equip_valide %>% 
  filter ( `Manomètre Pose?` != "") %>%
  filter ( `Standard d'équipement` == "Chapelle") %>%
  #filter (!is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MANOMETREPOSE?",
            value = `Manomètre Pose?`)

eq_ass_attribut_pressionextraction <- ven_equip_valide %>%
  filter (`Pression extraction` != "0") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PRESSIONEXTRACTION",
            value = `Pression extraction`) 

eq_ass_attribut_pressionpulsion <- ven_equip_valide %>%
  filter (`Pression pulsion` != "") %>%
  filter (`Pression pulsion` != "0") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PRESSIONPULSION",
            value = `Pression pulsion`) 

eq_ass_attribut_moteurtype <- ven_equip_valide %>%
  filter (`Moteur Type` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MOTEURTYPE",
            value = `Moteur Type`) 

eq_ass_attribut_moteurtension <- ven_equip_valide %>%
  filter (`Tension` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MOTEURTENSION",
            value = `Tension`) 

eq_ass_attribut_puissance <- ven_equip_valide %>%
  filter (`Puissance` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PUISSANCE",
            value = `Puissance`) 

eq_ass_attribut_moteurnbtours <- ven_equip_valide %>%
  filter (`Nb tours` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MOTEURNBTOURS",
            value = `Nb tours`) 

eq_ass_attribut_moteurnbvitesse <- ven_equip_valide %>%
  filter (`Nb vitesse` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MOTEURNBVITESSE",
            value = `Nb vitesse`) 

eq_ass_attribut_moteuramperage <- ven_equip_valide %>%
  filter (`Ampérage nominale` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MOTEURAMPERAGE",
            value = `Ampérage nominale`) 

eq_ass_attribut_ventiallateurtype <- ven_equip_valide %>%
  filter (`Ventilateur type` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "VENTILATEURTYPE",
            value = `Ventilateur type`) 

eq_ass_attribut_pressionstatique <-  ven_equip_valide %>%
  filter (`Pression statique` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PRESSIONSTATIQUE",
            value = `Pression statique`) 

eq_ass_attribut_transmissiontype <- ven_equip_valide %>%
  filter (`Transmission type` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSMISSIONTYPE",
            value = `Transmission type`) 

eq_ass_attribut_position <- ven_equip_valide %>%
  filter (`Position` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "POSITION",
            value = `Position`) 

eq_ass_attribut_matiere <- ven_equip_valide %>%
  filter (`Matière` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MATIERE",
            value = `Matière`) 

eq_ass_attribut_courroieforme <- ven_equip_valide %>%
  filter (`Courroie forme` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "COURROIEFORME",
            value = `Courroie forme`) 

eq_ass_attribut_courroietype <- ven_equip_valide %>%
  filter (`Courroie Type` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "COURROIETYPE",
            value = `Courroie Type`) 

#eq_ass_attribut_priseairdeniernet <- ven_equip_valide %>%
#  filter (`Prise Aire Dernier nettoyage` != "") %>%
#  filter (`Prise Aire Dernier nettoyage` != "00.00.00") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "PRISEAIREDERNIERNET",
#            value = `Prise Aire Dernier nettoyage`) 

eq_ass_attribut_climatisteurarmoire <- ven_equip_valide %>%
  filter (`Climatiseur Armoire Marque type` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "CLIMATISEURARMOIRE",
            value = `Climatiseur Armoire Marque type`) 

eq_ass_attribut_raccordement <- ven_equip_valide %>%
  filter (`Raccordement` != "") %>%
  filter (is.na(`parent_eq_id`)) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RACCORDEMENT",
            value = `Raccordement`) 

#eq_ass_attribut_fichetechnique <- ven_equip_valide %>%
#  filter (`Fiche technique` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "FICHETECHNIQUE",
#            value = `Fiche technique`) 

eq_ass_attribut_diametreraccordement <- ven_equip_valide %>%
  filter ( `Diamètre raccordement` != "") %>%
  filter ( `Diamètre raccordement` != 0) %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIAMETRERACCORDEMENT", 
            value = `Diamètre raccordement`)

###### MOBILIER



#eq_ass_attribut_intervention <- ven_equip_valide %>%
#  filter ( `Intervention` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "INTERVENTION",
#            value = `Intervention`)

# Attributs VEN Accessoire

#eq_ass_attribut_idcontrat <- ven_equip_acc0 %>% 
#  filter (`IDContrat` != "") %>% 
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "IDCONTRAT",
#            value = `IDContrat`) 

#eq_ass_attribut_ven_type <- ven_equip_acc0 %>%
#  filter (AccessoireType != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TYPE",
#            value = `AccessoireType`)

#eq_ass_attribut_ven_dimension <- ven_equip_acc0 %>%
#  filter (`Dimension` != "") %>%
#  filter (`Dimension` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DIMENSION",
#            value = `Dimension`)

#eq_ass_attribut_debitair <- ven_equip_acc0 %>%
#  filter (`DebitAir` != "") %>%
#  filter (`DebitAir` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DEBITAIR",
#            value = `DebitAir`) 

#eq_ass_attribut_diametreraccordement <- ven_equip_acc0 %>%
#  filter (`DiametreRaccordement` != "") %>%
#  filter (`DiametreRaccordement` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DIAMETRERACCORDEMENT",
#            value = `DiametreRaccordement`) 

#eq_ass_attribut_vitesseairprevue <- ven_equip_acc0 %>%
#  filter (`VitesseAirPrevue` != "") %>%
#  filter (`VitesseAirPrevue` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "VITESSEAIRPREVUE",
#            value = `VitesseAirPrevue`) 

#eq_ass_attribut_vitesseairmesuree <- ven_equip_acc0 %>% 
#  filter (`VitesseAirMesuree` != "") %>%
#  filter (`VitesseAirMesuree` != "0") %>% 
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "VITESSEAIRMESUREE",
#            value = `VitesseAirMesuree`) 

#eq_ass_attribut_manometreposeoui <- ven_equip_acc0 %>% 
#  filter (`ManometrePoseOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "MANOMETREPOSEOUI",
#            value = `ManometrePoseOui`)

#eq_ass_attribut_manometrepression <- ven_equip_acc0 %>% 
#  filter (`ManometrePression` != "") %>%
#  filter (`ManometrePression` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "MANOMETREPRESSION",
#            value = `ManometrePression`) 

#eq_ass_attribut_guillotinechapelle <- ven_equip_acc0 %>% 
#  filter (`GuillotineChapelle` != "") %>%
#  filter (`GuillotineChapelle` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "GUILLOTINECHAPELLE", 
#            value = `GuillotineChapelle`) 

#eq_ass_attribut_horlogepvgvoui <- ven_equip_acc0 %>% 
#  filter (`HorlogePVGVOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "HORLOGEPVGVOUI",
#            value = `HorlogePVGVOui`)

#eq_ass_attribut_forsagegvoui <- ven_equip_acc0 %>%
#  filter (`ForsageGVOui` != "") %>% 
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "FORSAGEGVOUI",
#            value = `ForsageGVOui`) 

#eq_ass_attribut_elempvoui <- ven_equip_acc0 %>% 
#  filter (`ElemPVOui` != "") %>% 
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ELEMPVOUI", 
#            value = `ElemPVOui`)

#eq_ass_attribut_elemgvoui <- ven_equip_acc0 %>%
#  filter (`ElemGVOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ELEMGVOUI",
#            value = `ElemGVOui`)

#eq_ass_attribut_elemautooui <- ven_equip_acc0 %>%
#  filter (`ElemAUTOOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ELEMAUTOOUI",
#            value = `ElemAUTOOui`)

#eq_ass_attribut_elemhorsoui <- ven_equip_acc0 %>%
#  filter (`ElemHORSOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ELEMHORSOUI",
#            value = `ElemHORSOui`)

#eq_ass_attribut_elempanneoui <- ven_equip_acc0 %>%
#  filter (`ElemPanneOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ELEMPANNEOUI",
#            value = `ElemPanneOui`)

#eq_ass_attribut_elemnoplatine <- ven_equip_acc0 %>%
#  filter (`ElemNoPlatine` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ELEMNOPLATINE",
#            value = `ElemNoPlatine`)

#eq_ass_attribut_dimension2 <- ven_equip_acc0 %>%
#  filter (`Dimension2` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DIMENSION2",
#            value = `Dimension2`)

#eq_ass_attribut_dimension3 <- ven_equip_acc0 %>%
#  filter (`Dimension3` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DIMENSION3",
#            value = `Dimension3`)

#eq_ass_attribut_acontroleroui <- ven_equip_acc0 %>%
#  filter (`AControlerOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ACONTROLEROUI",
#            value = `AControlerOui`)

#eq_ass_attribut_idno <- ven_equip_acc0 %>%
#  filter (`IDNo` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "IDNO",
#            value = `IDNo`)

#eq_ass_attribut_ven_quantite <- ven_equip_acc0 %>%
#  filter (`Quantite` != "") %>%
#  filter (`Quantite` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "QUANTITE",
#            value = `Quantite`)

#eq_ass_attribut_stock <- ven_equip_acc0 %>%
#  filter (`Stock` != "") %>%
#  filter (`Stock` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "STOCK",
#            value = `Stock`)

#eq_ass_attribut_etat <- ven_equip_acc0 %>%
#  filter (`Etat` != "") %>%
# filter (`Etat` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ETAT",
#            value = `Etat`)

#eq_ass_attribut_intervention <- ven_equip_acc0 %>%
#  filter (`Intervention` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "INTERVENTION",
#            value = `Intervention`)

#eq_ass_attribut_majpar <- ven_equip_acc0 %>%
#  filter (`MAJPar` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "MAJPAR",
#            value = `MAJPar`)

#eq_ass_attribut_ctrl1 <- ven_equip_acc0 %>%
#  filter (`Ctrl1` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "CTRL1",
#            value = `Ctrl1`)

#eq_ass_attribut_chapelleoui <- ven_equip_acc0 %>%
#  filter (`ChapelleOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "CHAPELLEOUI",
#            value = `ChapelleOui`)

#eq_ass_attribut_iddepannge <- ven_equip_acc0 %>%
#  filter (`IDDepannge` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "IDDEPANNGE",
#            value = `IDDepannge`)

#eq_ass_attribut_debitdifferenceoui <- ven_equip_acc0 %>%
#  filter (`DebitDifferenceOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DEBITDIFFERENCEOUI",
#            value = `DebitDifferenceOui`)

#eq_ass_attribut_debitairmesure <- ven_equip_acc0 %>%
#  filter (`DebitAirMesure` != "") %>%
#  filter (`DebitAirMesure` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DEBITAIRMESURE",
#            value = `DebitAirMesure`)

# Attributs SAN Accessoire

#eq_ass_attribut_san_type <- san_equip_acc0 %>%
#  filter (Type != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TYPE",
#            value = Type)

#eq_ass_attribut_valeurhoraire <- san_equip_acc0 %>%
#  filter (ValeurHoraire != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "VALEURHORAIRE",
#            value = ValeurHoraire)

#eq_ass_attribut_diametre <- san_equip_acc0 %>%
#  filter (Diametre != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DIAMETRE",
#            value = Diametre)

#eq_ass_attribut_debitnom <- san_equip_acc0 %>%
#  filter (`DebitNom` != "") %>%
#  filter (`DebitNom` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DEBITNOM",
#            value = `DebitNom`) 

#eq_ass_attribut_debiteffectif <- san_equip_acc0 %>%
#  filter (`DebitEffectif` != "") %>%
#  filter (`DebitEffectif` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DEBITEFFECTIF",
#            value = `DebitEffectif`) 

#eq_ass_attribut_position <- san_equip_acc0 %>%
#  filter (`Position` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "POSITION",
#            value = `Position`) 

# Attributs LEVAG
################################

#eq_ass_attribut_levag_frequence <- levag_equip_valide %>% 
#  filter (`Fréquence` != "") %>% 
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "FREQUENCE",
#            value = `Fréquence`) 

eq_ass_attribut_levag_modele <- levag_equip_valide %>% 
  filter (`Modèle` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MODELE",
            value = `Modèle`) 

#eq_ass_attribut_levag_lieu <- cha_equip0 %>%
#  filter (Lieu != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "LIEU",
#            value = Lieu)

eq_ass_attribut_levag_idcontrat <- levag_equip_valide %>% 
  filter (`ID Contrat` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "IDCONTRAT",
            value = `ID Contrat`) 

eq_ass_attribut_levag_noesti <- levag_equip_valide %>% 
  filter (`No ESTI` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NOESTI",
            value = `No ESTI`) 

#eq_ass_attribut_levag_infoscontrat <- levag_equip_valide %>% 
#  filter (`Infos contrat` != "") %>% 
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "INFOSCONTRAT",
#            value = `Infos contrat`) 

eq_ass_attribut_levag_type <- levag_equip_valide %>%
  filter (`Type` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = `Type`)

eq_ass_attribut_levag_nosurplan <- levag_equip_valide %>%
  filter (`No sur plan` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NOSURPLAN",
            value = `No sur plan`)

#eq_ass_attribut_levag_infosmes <- levag_equip_valide %>%
#  filter (`Infos MES` != "") %>%	
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "INFOSMES",
#            value = `Infos MES`)

eq_ass_attribut_levag_marque <- levag_equip_valide %>%
  filter (`Marque` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MARQUE",
            value = `Marque`)

eq_ass_attribut_levag_charge <- levag_equip_valide %>%
  filter (`Charge utile` != "") %>%
transmute("#eq_asset_attribute.eq_id" = eq_id,
"asset_attribute_std" = "CHARGEUTILE",
            value = `Charge utile`)


# Attributs ELE
################################

#eq_ass_attribut_ele_frequence <- ele_equip_valide %>% 
#  filter (`Fréquence` != "") %>% 
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "FREQUENCE",
#            value = `Fréquence`) 

eq_ass_attribut_ele_idcontrat <- ele_equip_valide %>% 
  filter (`ID Contrat` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "IDCONTRAT",
            value = `ID Contrat`) 

eq_ass_attribut_ele_noesti <- ele_equip_valide %>% 
  filter (`No ESTI` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NOESTI",
            value = `No ESTI`) 

#eq_ass_attribut_ele_infoscontrat <- ele_equip_valide %>% 
#  filter (`Infos contrat` != "") %>% 
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "INFOSCONTRAT",
#            value = `Infos contrat`) 

eq_ass_attribut_ele_type <- ele_equip_valide %>%
  filter (`Type` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = `Type`)

eq_ass_attribut_ele_nosurplan <- ele_equip_valide %>%
  filter (`No sur plan` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NOSURPLAN",
            value = `No sur plan`)

#eq_ass_attribut_ele_infosmes <- ele_equip_valide %>%
#  filter (`Infos MES` != "") %>%	
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "INFOSMES",
#            value = `Infos MES`)

eq_ass_attribut_ele_lieu <- ele_equip_valide %>%
  filter (`Lieu` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "LIEU",
            value = `Lieu`)

eq_ass_attribut_ele_marque <- ele_equip_valide %>%
  filter (`Marque` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MARQUE",
            value = `Marque`)

eq_ass_attribut_ele_type <- ele_equip_valide %>%
  filter (`Type` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = `Type`)

#eq_ass_attribut_nbreheure <- ele_equip_acc0 %>%
#  filter (`NbreHeure` != "") %>%
#  filter (`NbreHeure` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "NBREHEURE",
#            value = `NbreHeure`)

#eq_ass_attribut_ele_quantite <- ele_equip_acc0 %>%
#  filter (`Quantite` != "") %>%
#  filter (`Quantite` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "QUANTITE",
#            value = `Quantite`)

#eq_ass_attribut_nbrepersonnes <- ele_equip_acc0 %>%
#  filter (`NbrePersonnes` != "") %>%
#  filter (`NbrePersonnes` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "NBREPERSONNES",
#            value = `NbrePersonnes`)

#eq_ass_attribut_puissance <- ele_equip_acc0 %>%
#  filter (`Puissance` != "") %>%
#  filter (`Puissance` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "PUISSANCE",
#            value = `Puissance`)

#eq_ass_attribut_anneeenservice <- ele_equip_acc0 %>%
#  filter (`AnneeEnService` != "") %>%
#  filter (`AnneeEnService` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#"asset_attribute_std" = "ANNEEENSERVICE",
#            value = `AnneeEnService`)

#eq_ass_attribut_anneeconstruction <- ele_equip_acc0 %>%
#  filter (`AnneeConstruction` != "") %>%
#  filter (`AnneeConstruction` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ANNEECONSTRUCTION",
#            value = `AnneeConstruction`)

#eq_ass_attribut_relaifournisseur <- ele_equip_acc0 %>%
#  filter (`RelaiFournisseur` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "RELAIFOURNISSEUR",
#            value = `RelaiFournisseur`)

#eq_ass_attribut_autotransfooui <- ele_equip_acc0 %>%
#  filter (`AutotransfoOui` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#           "asset_attribute_std" = "AUTOTRANSFOOUI",
#            value = `AutotransfoOui`)

#eq_ass_attribut_transforeseau <- ele_equip_acc0 %>%
#  filter (`TransfoReseau` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFORESEAU",
#            value = `TransfoReseau`)

#eq_ass_attribut_transfocellule <- ele_equip_acc0 %>%
#  filter (`TransfoCellule` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOCELLULE",
#            value = `TransfoCellule`)

#eq_ass_attribut_relaitype <- ele_equip_acc0 %>%
#  filter (`RelaiType` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "RELAITYPE",
#            value = `RelaiType`)

#eq_ass_attribut_unominale <- ele_equip_acc0 %>%
#  filter (`UNominale` != "") %>%
#  filter (`UNominale` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "UNOMINALE",
#            value = `UNominale`)

#eq_ass_attribut_transfouprimaire <- ele_equip_acc0 %>%
#  filter (`TransfoUPrimaire` != "") %>%
#  filter (`TransfoUPrimaire` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOUPRIMAIRE",
#            value = `TransfoUPrimaire`)

#eq_ass_attribut_transfousecondairereglee <- ele_equip_acc0 %>%
#  filter (`TransfoUSecondaireReglee` != "") %>%
#  filter (`TransfoUPrimaire` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOUSECONDAIREREGLEE",
#            value = `TransfoUSecondaireReglee`)

#eq_ass_attribut_transforemplissage <- ele_equip_acc0 %>%
#  filter (`TransfoRemplissage` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOREMPLISSAGE",
#            value = `TransfoRemplissage`)

#eq_ass_attribut_transfocouplage <- ele_equip_acc0 %>%
#  filter (`TransfoCouplage` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOCOUPLAGE",
#            value = `TransfoCouplage`)

#eq_ass_attribut_transfoucc <- ele_equip_acc0 %>%
#  filter (`TransfoUcc` != "") %>%
#  filter (`TransfoUcc` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOUCC",
#            value = `TransfoUcc`)

#eq_ass_attribut_uprimaire <- ele_equip_acc0 %>%
#  filter (`UPrimaire` != "") %>%
#  filter (`UPrimaire` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "UPRIMAIRE",
#            value = `UPrimaire`)

#eq_ass_attribut_transfousecondaire1 <- ele_equip_acc0 %>%
#  filter (`TransfoUSecondaire1` != "") %>%
#  filter (`TransfoUSecondaire1` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOUSECONDAIRE1",
#            value = `TransfoUSecondaire1`)

#eq_ass_attribut_transfousecondaire2 <- ele_equip_acc0 %>%
#  filter (`TransfoUSecondaire2` != "") %>%
#  filter (`TransfoUSecondaire2` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOUSECONDAIRE2",
#            value = `TransfoUSecondaire2`)

#eq_ass_attribut_transfousecondaire3 <- ele_equip_acc0 %>%
#  filter (`TransfoUSecondaire3` != "") %>%
#  filter (`TransfoUSecondaire3` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOUSECONDAIRE3",
#            value = `TransfoUSecondaire3`)

#eq_ass_attribut_transfousecondaire4 <- ele_equip_acc0 %>%
#  filter (`TransfoUSecondaire4` != "") %>%
#  filter (`TransfoUSecondaire4` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOUSECONDAIRE4",
#            value = `TransfoUSecondaire4`)

#eq_ass_attribut_transfousecondaire5 <- ele_equip_acc0 %>%
#  filter (`TransfoUSecondaire5` != "") %>%
#  filter (`TransfoUSecondaire5` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOUSECONDAIRE5",
#            value = `TransfoUSecondaire5`)

#eq_ass_attribut_transfoperteavide <- ele_equip_acc0 %>%
#  filter (`TransfoPerteAVide` != "") %>%
#  filter (`TransfoPerteAVide` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOPERTEAVIDE",
#            value = `TransfoPerteAVide`)

#eq_ass_attribut_transfoperteencharge <- ele_equip_acc0 %>%
#  filter (`TransfoPerteEnCharge` != "") %>%
#  filter (`TransfoPerteEnCharge` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOPERTEENCHARGE",
#            value = `TransfoPerteEnCharge`)

#eq_ass_attribut_transfomassetotale <- ele_equip_acc0 %>%
#  filter (`TransfoMasseTotale` != "") %>%
#  filter (`TransfoMasseTotale` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOMASSETOTALE",
#            value = `TransfoMasseTotale`)

#eq_ass_attribut_transfomaddehuile <- ele_equip_acc0 %>%
#  filter (`TransfoMaddeHuile` != "") %>%
#  filter (`TransfoMaddeHuile` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOMADDEHUILE",
#            value = `TransfoMaddeHuile`)

#eq_ass_attribut_transfoprealarme <- ele_equip_acc0 %>%
#  filter (`TransfoPreAlarme` != "") %>%
#  filter (`TransfoPreAlarme` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOPREALARME",
#            value = `TransfoPreAlarme`)

#eq_ass_attribut_transfoalarme <- ele_equip_acc0 %>%
#  filter (`TransfoAlarme` != "") %>%
#  filter (`TransfoAlarme` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TRANSFOALARME",
#            value = `TransfoAlarme`)

#eq_ass_attribut_disjoncteurinominal <- ele_equip_acc0 %>%
#  filter (`DisjoncteurINominal` != "") %>%
#  filter (`DisjoncteurINominal` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DISJONCTEURINOMINAL",
#            value = `DisjoncteurINominal`)

#eq_ass_attribut_disjoncteurpouvoirdecoupure <- ele_equip_acc0 %>%
#  filter (`DisjoncteurPouvoirDeCoupure` != "") %>%
#  filter (`DisjoncteurPouvoirDeCoupure` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DISJONCTEURPOUVOIRDECOUPURE",
#            value = `DisjoncteurPouvoirDeCoupure`)

#eq_ass_attribut_relaiithernmiqueregle <- ele_equip_acc0 %>%
#  filter (`RelaiIThernmiqueRegle` != "") %>%
#  filter (`RelaiIThernmiqueRegle` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "RELAIITHERNMIQUEREGLE",
#            value = `RelaiIThernmiqueRegle`)

#eq_ass_attribut_relaideclancheinstantane <- ele_equip_acc0 %>%
#  filter (`RelaiDeclancheInstantane` != "") %>%
#  filter (`RelaiDeclancheInstantane` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "RELAIDECLANCHEINSTANTANE",
#            value = `RelaiDeclancheInstantane`)

#eq_ass_attribut_relaiconstantedetemps <- ele_equip_acc0 %>%
#  filter (`RelaiConstanteDeTemps` != "") %>%
#  filter (`RelaiConstanteDeTemps` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "RELAICONSTANTEDETEMPS",
#            value = `RelaiConstanteDeTemps`)

#eq_ass_attribut_relaiireponse <- ele_equip_acc0 %>%
#  filter (`RelaiIReponse` != "") %>%
#  filter (`RelaiIReponse` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "RELAIIREPONSE",
#            value = `RelaiIReponse`)

#eq_ass_attribut_relaitemporisation <- ele_equip_acc0 %>%
#  filter (`RelaiTemporisation` != "") %>%
#  filter (`RelaiTemporisation` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "RELAITEMPORISATION",
#            value = `RelaiTemporisation`)

#eq_ass_attribut_relaiinominal <- ele_equip_acc0 %>%
#  filter (`RelaiINominal` != "") %>%
#  filter (`RelaiINominal` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "RELAIINOMINAL",
#            value = `RelaiINominal`)
# Assenceurs

eq_ass_attribut_modele <- leva_equip0 %>% 
  filter (`Modèle` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MODELE",
            value = `Modèle`) 

eq_ass_attribut_dimensions <- leva_equip0 %>%
  filter (`Dimensions L-P-H (cm)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIMENSIONSLPH",
            value = `Dimensions L-P-H (cm)`)

eq_ass_attribut_charge <- leva_equip0 %>%
  filter (`Charge utile (kg)` != "") %>%
transmute("#eq_asset_attribute.eq_id" = eq_id,
"asset_attribute_std" = "CHARGEUTILEKG",
            value = `Charge utile (kg)`)

eq_ass_attribut_telascenseur <- leva_equip0 %>%
  filter (`No téléphone ascenseur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TELASCENSEUR",
            value = `No téléphone ascenseur`)

eq_ass_attribut_sytemeurgence <- leva_equip0 %>%
  filter (`Système d'appel d'urgence` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "SYSTEMEURGENCE",
            value = `Système d'appel d'urgence`)

eq_ass_attribut_idcontrat <- leva_equip0 %>% 
  filter (`Contrat` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "IDCONTRAT",
            value = `Contrat`) 


# Mobiliers

#eq_ass_attribut_mob_type <- mob_equip0 %>%
#  filter (`Type` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "TYPE",
#            value = `Type`)

#eq_ass_attribut_image <- mob_equip0 %>%
#  filter (`Image source` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "IMAGE",
#            value = `Image source`)

#eq_ass_attribut_originale <- mob_equip0 %>%
#  filter (`Originale?` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "ORIGINALE",
#            value = `Originale?`)

# Mobilier Accesoires



#eq_ass_attribut_prix <- mob_equip_acc0 %>%
#  filter (`PrixUnitaire` != "") %>%
#  filter (`PrixUnitaire` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "PRIX",
#            value = `PrixUnitaire`)

#eq_ass_attribut_mob_quantite <- mob_equip_acc0 %>%
#  filter (`Quantite` != "") %>%
#  filter (`Quantite` != "0") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "QUANTITE",
#            value = `Quantite`)

# SV



eq_ass_sv_attribut_fournisseur <- sv_sti_equip %>%
  filter (`Fournisseur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `Code équipement`,
            "asset_attribute_std" = "FOURNISSEUR",
            value = `Fournisseur`)

eq_ass_sv_attribut_niveausecu <- sv_sti_equip %>%
  filter (`BioSafety Level` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `Code équipement`,
            "asset_attribute_std" = "NIVEAUSECU",
            value = `BioSafety Level`)

#eq_ass_sv_attribut_siglelabo <- sv_sti_equip %>%
#  filter (`Sigle du Labo` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = `Code équipement`,
#            "asset_attribute_std" = "SIGLELABO",
#            value = `Sigle du Labo`)

eq_ass_sv_attribut_contactlabo <- sv_sti_equip %>%
  filter (`Contact du Labo` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `Code équipement`,
            "asset_attribute_std" = "CONTACTLABO",
            value = `Contact du Labo`)

eq_ass_sv_attribut_nosap <- sv_sti_equip %>%
  filter (`N°SAP` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `Code équipement`,
            "asset_attribute_std" = "NOSAP",
            value = `N°SAP`)

# Atttibuts Cellules MT

eq_ass_attribut_mt_numesti <- mt_equip0 %>%
  filter ( `Num ESTI` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NOESTI",
            value = `Num ESTI`)

eq_ass_attribut_mt_type <- mt_equip0 %>% 
  filter (`Type` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = `Type`) 

eq_ass_attribut_mt_sf6 <- mt_equip0 %>%
  filter ( `SF6 (kg)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "SF6",
            value = `SF6 (kg)`)

eq_ass_attribut_mt_nombredemanœuvresmax <- mt_equip0 %>%
  filter ( `Nombre de manœuvres max` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id, 
            "asset_attribute_std" = "NOMBREDEMANŒUVRESMAX",
            value = `Nombre de manœuvres max`)

eq_ass_attribut_mt_tranformateurintensite <- mt_equip0 %>%
  filter ( `Tranformateur d\'intensité` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANFORMATEURINTENSITE",
            value = `Tranformateur d\'intensité`)

eq_ass_attribut_mt_puissance <- mt_equip0 %>%
  filter ( `Puissance (VA)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PUISSANCE",
            value = `Puissance (VA)`)

eq_ass_attribut_mt_classedeprecision <- mt_equip0 %>%
  filter ( `Classe de précision` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "CLASSEDEPRECISION",
            value = `Classe de précision`)

eq_ass_attribut_mt_transformateurpotentiel <- mt_equip0 %>%
  filter ( `Transformateur de potentiel` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFORMATEURPOTENTIEL",
            value = `Transformateur de potentiel`)

eq_ass_attribut_mt_alphae <- mt_equip0 %>%
  filter ( `Alpha E oui/non` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ALPHAE",
            value = `Alpha E oui/non`)

eq_ass_attribut_mt_alphaeitrip <- mt_equip0 %>%
  filter ( `Alpha E I trip (A)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ALPHAEITRIP",
            value = `Alpha E I trip (A)`)

eq_ass_attribut_mt_alphaetreset <- mt_equip0 %>%
  filter ( `Alpha E t reset (H)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ALPHAETRESET",
            value = `Alpha E t reset (H)`)

eq_ass_attribut_mt_alphaerelay <- mt_equip0 %>%
  filter ( `Alpha E Relay` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ALPHAERELAY",
            value = `Alpha E Relay`)

eq_ass_attribut_mt_relaisseuili <- mt_equip0 %>%
  filter ( `Relais - Seuil I>` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIS-SEUILI",
            value = `Relais - Seuil I>`)

eq_ass_attribut_mt_relaistdecl <- mt_equip0 %>%
  filter ( `Relais - Tdécl` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIS-TDECL",
            value = `Relais - Tdécl`)

eq_ass_attribut_mt_relaisseuili <- mt_equip0 %>%
  filter ( `Relais - Seuil I>>` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIS-SEUILI",
            value = `Relais - Seuil I>>`)

eq_ass_attribut_mt_relaistdecl2 <- mt_equip0 %>%
  filter ( `Relais - Tdécl2` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIS-TDECL2",
            value = `Relais - Tdécl2`)

eq_ass_attribut_mt_relaisseuilith <- mt_equip0 %>%
  filter ( `Relais - Seuil Ith` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIS-SEUILITH",
            value = `Relais - Seuil Ith`)

eq_ass_attribut_mt_relaistdecl3 <- mt_equip0 %>%
  filter ( `Relais - Tdécl3` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIS-TDECL3",
            value = `Relais - Tdécl3`)

eq_ass_attribut_mt_puissance <- mt_equip0 %>%
  filter ( `Puissance (kVA)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PUISSANCE",
            value = `Puissance (kVA)`)

eq_ass_attribut_mt_couplage <- mt_equip0 %>%
  filter ( `Couplage` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "COUPLAGE",
            value = `Couplage`)

eq_ass_attribut_mt_rapporttransformation <- mt_equip0 %>%
  filter ( `Rapport de transformation` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RAPPORTTRANSFORMATION",
            value = `Rapport de transformation`)

eq_ass_attribut_mt_plot1 <- mt_equip0 %>%
  filter ( `Plot 1` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PLOT1",
            value = `Plot 1`)

eq_ass_attribut_mt_plot2 <- mt_equip0 %>%
  filter ( `Plot 2` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PLOT2",
            value = `Plot 2`)

eq_ass_attribut_mt_plot3 <- mt_equip0 %>%
  filter ( `Plot 3` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PLOT3",
            value = `Plot 3`)

eq_ass_attribut_mt_plot4 <- mt_equip0 %>%
  filter ( `Plot 4` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PLOT4",
            value = `Plot 4`)

eq_ass_attribut_mt_plot5 <- mt_equip0 %>%
  filter ( `Plot 5` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PLOT5",
            value = `Plot 5`)

eq_ass_attribut_mt_naturedesenroulements <- mt_equip0 %>%
  filter ( `Nature des enroulements` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NATUREDESENROULEMENTS",
            value = `Nature des enroulements`)

eq_ass_attribut_mt_courantnominalbt <- mt_equip0 %>%
  filter ( `Courant nominal BT (A)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "COURANTNOMINALBT",
            value = `Courant nominal BT (A)`)

eq_ass_attribut_mt_tensiondecourtcircuit <- mt_equip0 %>%
  filter ( `Tension de court circuit (%)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TENSIONDECOURTCIRCUIT",
            value = `Tension de court circuit (%)`)

eq_ass_attribut_mt_massetotale <- mt_equip0 %>%
  filter ( `Masse totale (t)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MASSETOTALE",
            value = `Masse totale (t)`)

eq_ass_attribut_mt_massehuile <- mt_equip0 %>% 
  filter ( `Masse huile (t)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MASSEHUILE",
            value = `Masse huile (t)`)

eq_ass_attribut_mt_pertesencharge <- mt_equip0 %>%
  filter ( `Pertes en charge (W)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PERTESENCHARGE",
            value = `Pertes en charge (W)`)

eq_ass_attribut_mt_pertesavide <- mt_equip0 %>%
  filter ( `Pertes à vide (W)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PERTESAVIDE",
            value = `Pertes à vide (W)`)

eq_ass_attribut_mt_temp_alarme <- mt_equip0 %>%
  filter ( `Temp. alarme` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TEMP.ALARME",
            value = `Temp. alarme`)

eq_ass_attribut_mt_temp_declenchement <- mt_equip0 %>%
  filter ( `Temp. déclenchement` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TEMP.DECLENCHEMENT",
            value = `Temp. déclenchement`)

#eq_ass_attribut_mt_document <- mt_equip0 %>%
#  filter ( `Document` != "") %>%
#  transmute("#eq_asset_attribute.eq_id" = eq_id,
#            "asset_attribute_std" = "DOCUMENT",
#            value = `Document`)


###
# ENERIGIE
###

eq_ass_attribut_enregie_nopod <- energie %>%
  filter ( `Numéro POD compteur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `ID_ARCHIBUS`,
            "asset_attribute_std" = "NOPOD",
            value = `Numéro POD compteur`)

eq_ass_attribut_enregie_dneau <- energie %>%
  filter ( `DN compteur eau` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `ID_ARCHIBUS`,
            "asset_attribute_std" = "DNEAU",
            value = `DN compteur eau`)

eq_ass_attribut_enregie_codeureau <- energie %>%
  filter ( `Codeur compteur eau` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `ID_ARCHIBUS`,
            "asset_attribute_std" = "CODEUREAU",
            value = `Codeur compteur eau`)

eq_ass_attribut_enregie_ipcompteur <- energie %>%
  filter ( `Adresse IP compteur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `ID_ARCHIBUS`,
            "asset_attribute_std" = "IPCOMPTEUR",
            value = `Adresse IP compteur`)

eq_ass_attribut_enregie_modelecompteur <- energie %>%
  filter ( `Modèle compteur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `ID_ARCHIBUS`,
            "asset_attribute_std" = "MODELECOMPTEUR",
            value = `Modèle compteur`)

eq_ass_attribut_enregie_relevecompteur <- energie %>%
  filter ( `Relevé compteur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `ID_ARCHIBUS`,
            "asset_attribute_std" = "REVLEVECOMPTEUR",
            value = `Relevé compteur`)

eq_ass_attribut_enregie_typecompteur <- energie %>%
  filter ( `Type Compteur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `ID_ARCHIBUS`,
            "asset_attribute_std" = "TYPECOMPTEUR",
            value = `Type Compteur`)

eq_ass_attribut_enregie_versionfirmware <- energie %>%
  filter ( `Version firmware` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = `ID_ARCHIBUS`,
            "asset_attribute_std" = "VERSIONFIRMWARE",
            value = `Version firmware`)



####
# TCVS
####

eq_ass_attribut_tcvs_installation <- tcvs_detail %>%
  filter ( `Installation` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "INSTALLATION",
            value = `Installation`)

eq_ass_attribut_tcvs_lieu <- tcvs_detail %>%
  filter ( `Lieu` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "LIEU",
            value = `Lieu`)


eq_ass_attribut <- rbind(
#                         eq_ass_attribut_lieu, 
#                         eq_ass_attribut_san_type,
#                         eq_ass_attribut_cha_type,
                        eq_ass_attribut_ven_type,
                        eq_ass_attribut_ven_dimension,
#                         eq_ass_attribut_diametre, 
#                         eq_ass_attribut_valeurhoraire,
                         eq_ass_attribut_installation,
                         eq_ass_attribut_monobloctype,
#                         eq_ass_attribut_monoblocno,
                         eq_ass_attribut_debitair,
                         eq_ass_attribut_debitair0,
                         eq_ass_attribut_pressionextraction,
                         eq_ass_attribut_pressionpulsion,
                         eq_ass_attribut_moteurtype,
                         eq_ass_attribut_moteurtension,
                         eq_ass_attribut_puissance,
                         eq_ass_attribut_moteurnbtours,
                         eq_ass_attribut_moteurnbvitesse,
                         eq_ass_attribut_moteuramperage,
                         eq_ass_attribut_ventiallateurtype,
                         eq_ass_attribut_pressionstatique,
                         eq_ass_attribut_transmissiontype,
                         eq_ass_attribut_position,
                         eq_ass_attribut_matiere,
                         eq_ass_attribut_courroieforme,
                         eq_ass_attribut_courroietype,
#                         eq_ass_attribut_priseairdeniernet,
                         eq_ass_attribut_climatisteurarmoire,
                         eq_ass_attribut_raccordement,
#                         eq_ass_attribut_fichetechnique,
                         eq_ass_attribut_diametreraccordement,
                         eq_ass_attribut_manometrepose,
#                         eq_ass_attribut_intervention,
                         eq_ass_attribut_idcontrat,
#                         eq_ass_attribut_dimension,
# LEVAG
#                         eq_ass_attribut_levag_frequence,
                         eq_ass_attribut_levag_modele,
#                         eq_ass_attribut_levag_lieu,
                         eq_ass_attribut_levag_idcontrat,
                         eq_ass_attribut_levag_noesti,
#                         eq_ass_attribut_levag_infoscontrat,
                         eq_ass_attribut_levag_type,
                         eq_ass_attribut_levag_nosurplan,
#                         eq_ass_attribut_levag_infosmes,
                         eq_ass_attribut_levag_marque,
                         eq_ass_attribut_levag_charge,
# ELEC                         
                         eq_ass_attribut_ele_lieu,
                         eq_ass_attribut_ele_marque,
                         eq_ass_attribut_ele_type,
#                         eq_ass_attribut_ele_infosmes,
                         eq_ass_attribut_ele_nosurplan,
                         eq_ass_attribut_ele_type,
#                         eq_ass_attribut_ele_infoscontrat,
                         eq_ass_attribut_ele_noesti,
                         eq_ass_attribut_ele_idcontrat,
                         eq_ass_attribut_dimensions,
#                         eq_ass_attribut_ele_frequence,
#                         eq_ass_attribut_diametreraccordement,
#                         eq_ass_attribut_vitesseairprevue,
#                         eq_ass_attribut_vitesseairmesuree,
#                         eq_ass_attribut_manometreposeoui,
#                         eq_ass_attribut_manometrepression,
#                         eq_ass_attribut_guillotinechapelle,
#                         eq_ass_attribut_horlogepvgvoui,
#                         eq_ass_attribut_forsagegvoui,
#                         eq_ass_attribut_elempvoui,
#                         eq_ass_attribut_elemgvoui,
#                         eq_ass_attribut_elemautooui,
#                         eq_ass_attribut_elemhorsoui,
#                         eq_ass_attribut_elempanneoui,
#                         eq_ass_attribut_elemnoplatine,
#                         eq_ass_attribut_dimension2,
#                         eq_ass_attribut_dimension3,
#                         eq_ass_attribut_acontroleroui,
#                         eq_ass_attribut_idno,
#                         eq_ass_attribut_ven_quantite,
#                         eq_ass_attribut_ele_quantite,
#                         eq_ass_attribut_mob_quantite,
#                         eq_ass_attribut_stock,
#                         eq_ass_attribut_etat,
#                         eq_ass_attribut_intervention,
#                         eq_ass_attribut_majpar,
#                         eq_ass_attribut_ctrl1,
#                         eq_ass_attribut_chapelleoui,
#                         eq_ass_attribut_iddepannge,
#                         eq_ass_attribut_debitdifferenceoui,
#                         eq_ass_attribut_debitairmesure,
#                         eq_ass_attribut_debitnom,
#                         eq_ass_attribut_debiteffectif,
#                         eq_ass_attribut_position,
#                         eq_ass_attribut_nbreheure,
#                         eq_ass_attribut_nbrepersonnes,
#                         eq_ass_attribut_anneeenservice,
#                         eq_ass_attribut_anneeconstruction,
#                         eq_ass_attribut_relaifournisseur,
#                         eq_ass_attribut_autotransfooui,
#                         eq_ass_attribut_transforeseau,
#                         eq_ass_attribut_transfocellule,
#                         eq_ass_attribut_relaitype,
#                         eq_ass_attribut_unominale,
#                         eq_ass_attribut_transfouprimaire,
#                         eq_ass_attribut_transfousecondairereglee,
#                         eq_ass_attribut_transforemplissage,
#                         eq_ass_attribut_transfocouplage,
#                         eq_ass_attribut_transfoucc,
#                         eq_ass_attribut_uprimaire,
#                         eq_ass_attribut_transfousecondaire1,
#                         eq_ass_attribut_transfousecondaire2,
#                         eq_ass_attribut_transfousecondaire3,
#                         eq_ass_attribut_transfousecondaire4,
#                         eq_ass_attribut_transfousecondaire5,
#                         eq_ass_attribut_transfoperteavide,
#                         eq_ass_attribut_transfoperteencharge,
#                         eq_ass_attribut_transfomassetotale,
#                         eq_ass_attribut_transfomaddehuile,
#                         eq_ass_attribut_transfoprealarme,
#                         eq_ass_attribut_transfoalarme,
#                         eq_ass_attribut_disjoncteurinominal,
#                         eq_ass_attribut_disjoncteurpouvoirdecoupure,
#                         eq_ass_attribut_relaiithernmiqueregle,
#                         eq_ass_attribut_relaideclancheinstantane,
#                         eq_ass_attribut_relaiconstantedetemps,
#                         eq_ass_attribut_relaiireponse,
#                         eq_ass_attribut_relaitemporisation,
#                         eq_ass_attribut_relaiinominal,
#                        eq_ass_attribut_dimensions,
                         eq_ass_attribut_modele,
                         eq_ass_attribut_charge,
                         eq_ass_attribut_telascenseur,
                         eq_ass_attribut_sytemeurgence,
#                         eq_ass_attribut_mob_type,
#                         eq_ass_attribut_image,
#                         eq_ass_attribut_originale,
#                         eq_ass_attribut_prix,
                          #eq_ass_attribut_valeur,
                          eq_ass_sv_attribut_fournisseur,
                          eq_ass_sv_attribut_niveausecu,
#                          eq_ass_sv_attribut_siglelabo,
                          eq_ass_sv_attribut_contactlabo,
                          eq_ass_sv_attribut_nosap,
                          eq_ass_attribut_mt_numesti,
eq_ass_attribut_mt_type,
eq_ass_attribut_mt_sf6,
eq_ass_attribut_mt_nombredemanœuvresmax,
eq_ass_attribut_mt_tranformateurintensite,
eq_ass_attribut_mt_puissance,
eq_ass_attribut_mt_classedeprecision,
eq_ass_attribut_mt_transformateurpotentiel,
eq_ass_attribut_mt_alphae,
eq_ass_attribut_mt_alphaeitrip,
eq_ass_attribut_mt_alphaetreset,
eq_ass_attribut_mt_alphaerelay,
eq_ass_attribut_mt_relaisseuili,
eq_ass_attribut_mt_relaistdecl,
eq_ass_attribut_mt_relaisseuili,
eq_ass_attribut_mt_relaistdecl2,
eq_ass_attribut_mt_relaisseuilith,
eq_ass_attribut_mt_relaistdecl3,
eq_ass_attribut_mt_puissance,
eq_ass_attribut_mt_couplage,
eq_ass_attribut_mt_rapporttransformation,
eq_ass_attribut_mt_plot1,
eq_ass_attribut_mt_plot2,
eq_ass_attribut_mt_plot3,
eq_ass_attribut_mt_plot4,
eq_ass_attribut_mt_plot5,
eq_ass_attribut_mt_naturedesenroulements,
eq_ass_attribut_mt_courantnominalbt,
eq_ass_attribut_mt_tensiondecourtcircuit,
eq_ass_attribut_mt_massetotale,
eq_ass_attribut_mt_massehuile,
eq_ass_attribut_mt_pertesencharge,
eq_ass_attribut_mt_pertesavide,
eq_ass_attribut_mt_temp_alarme,
eq_ass_attribut_mt_temp_declenchement,
eq_ass_attribut_enregie_nopod,
eq_ass_attribut_enregie_dneau,
eq_ass_attribut_enregie_codeureau,
eq_ass_attribut_enregie_ipcompteur,
eq_ass_attribut_enregie_modelecompteur,
eq_ass_attribut_enregie_relevecompteur,
eq_ass_attribut_enregie_typecompteur,
eq_ass_attribut_enregie_versionfirmware,
eq_ass_attribut_tcvs_installation,
eq_ass_attribut_tcvs_lieu
#eq_ass_attribut_mt_document
                          )

write_archibus(eq_ass_attribut, "./03.eq_asset_attrib.xlsx",
               table.header = "Equipment Asset Attributes",
               sheet.name = "Sheet1")