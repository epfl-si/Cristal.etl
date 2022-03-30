library(tidyverse)
library(readxl)
library(dplyr)


GMAO_file = "./GMAO_4D_Export2021_MAPPING_eqstd.xlsx"
GMAO_Acc_file = "./4D_GMAO_Accessoires_avec UUID_4-3-22.xlsx"
ELA_Ascenseurs_file = "./ELA_ascenseurs.xlsx"
SV_file = "./Maintenance _Equipements_INFRA_SV_4D.xlsx"

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
  
  # save the workbook to an Excel file
  xlsx::saveWorkbook(wb, filename)
}


toArchibusStatus <- function(etat) {
  
  case_when(etat == "FALSE"   ~ "in",
            etat == "TRUE"    ~ "out",
            etat == ""        ~ "",
            TRUE              ~ "out")
}



standards_equip <- read_excel(GMAO_file, "Standards d'équipements") %>%
  rename(std_eq = "Standard d'équipement", dom_tech = "Domaine technique") %>%
  transmute("#eqstd.eq_std" = std_eq,
            description = Description,
            category = dom_tech)


write_archibus(standards_equip, "./00.eqstd.xlsx",
               table.header = "Equipment Standards",
               sheet.name = "standards_equipements")

site_import <- read_excel("./Export Bâtiments.xlsx")

batiments_import <- read_excel("./06. rm.xlsx","Sheet1",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
  left_join(site_import, by=c("#rm.bl_id"="Building Code"))

##################################
# CHA equipement
##################################

# CHA Parents

cha_equip0 <- read_excel(GMAO_file, "CHA_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1)
cha_equip0$eq_id <-paste("CHA-00000-",formatC(seq.int(nrow(cha_equip0)), width=6, flag=0, format="d"),sep = "")

cha_equip_parent <- cha_equip0 %>%
  left_join(standards_equip, by=c("STANDARD D'EQUIPEMENT"="description")) %>%
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
            num_serial = "",
            modelno = "",
            subcomponent_of ="",
            mfr = "",
            asset_id = `ID Fiche`,
            status = "in",
            comments = Remarques)

# CHA Accesoires

cha_equip_acc0 <- read_excel(GMAO_Acc_file, "C_acc",col_names = TRUE, col_types = NULL, na = "", )
cha_equip_acc0$eq_id <-paste("CHA-00000-",formatC(seq.int(nrow(cha_equip_acc0)) + nrow(cha_equip0), width=6, flag=0, format="d"), sep = "")

cha_equip_acc <- cha_equip_acc0 %>%
  left_join(cha_equip_parent, by=c("IDFiche"="asset_id"),) %>%
  transmute(subcomponent_of = `#eq.eq_id`,
            "#eq.eq_id" = eq_id,
            eq_std = "ACCESSOIRE",
            bl_id = bl_id,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = site_id,
            description = AccessoireDescription,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = IDNo,
            modelno = Numero,
            mfr = Marque,
            asset_id = paste("CHA-",UUID, sep =""),
            status = "in",
            comments = Remarques)


cha_equip <- rbind(cha_equip_parent, cha_equip_acc)

write_archibus(cha_equip, "./01.eq-CHA.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")

##################################
# VEN equipement
##################################

# VEN Parents

ven_equip0 <- read_excel(GMAO_file, "VEN_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1)
ven_equip0$eq_id <-paste("VEN-00000-",formatC(seq.int(nrow(ven_equip0)), width=6, flag=0, format="d"),sep = "")

ven_equip_parent <- ven_equip0 %>%
  left_join(standards_equip, by=c("STANDARD D'EQUIPEMENT"="description")) %>%
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
            num_serial = "",
            subcomponent_of ="",
            mfr = Marque,
            asset_id = `ID Fiche`,
            status = toArchibusStatus(`HS?`),
#            date_in_service = `Mise en service`,
            comments = Remarques)

# VEN Accesoires

ven_equip_acc0 <- read_excel(GMAO_Acc_file, "V_acc",col_names = TRUE, col_types = NULL, na = "", )
ven_equip_acc0$eq_id <-paste("VEN-00000-",formatC(seq.int(nrow(ven_equip_acc0)) + nrow(ven_equip0), width=6, flag=0, format="d"), sep = "")

ven_equip_acc <- ven_equip_acc0 %>%
  left_join(ven_equip_parent, by=c("IDFiche"="asset_id")) %>%
  transmute(subcomponent_of = `#eq.eq_id`,
            "#eq.eq_id" = eq_id,
            eq_std = "ACCESSOIRE",
            bl_id = bl_id,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = site_id,
            description = AccessoireDescription,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = IDNo,
            mfr = "",
            asset_id = paste("VEN-",UUID, sep =""),
            status = "in",
#           date_in_service = "",
            comments = Remarques)


ven_equip <- rbind(ven_equip_parent, ven_equip_acc)

write_archibus(ven_equip, "./01.eq-VEN.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")

##################################
# SAN equipement
##################################

# SAN Parents

san_equip0 <- read_excel(GMAO_file, "SAN_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1)
san_equip0$eq_id <-paste("SAN-00000-",formatC(seq.int(nrow(san_equip0)), width=6, flag=0, format="d"),sep = "")

san_equip_parent <- san_equip0 %>%
  left_join(standards_equip, by=c("STANDARD D'EQUIPEMENT"="description")) %>%
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
            num_serial = "",
            subcomponent_of ="",
            mfr = "",
            asset_id = `ID Fiche`,
            status = "in",
            #            date_in_service = `Mise en service`,
            comments = Remarques)

# SAN Accesoires
san_equip_acc0 <- read_excel(GMAO_Acc_file, "S_acc",col_names = TRUE, col_types = NULL, na = "", )
san_equip_acc0$eq_id <-paste("SAN-00000-",formatC(seq.int(nrow(san_equip_acc0)) + nrow(san_equip0), width=6, flag=0, format="d"), sep = "")

san_equip_acc <- san_equip_acc0 %>%
  left_join(san_equip_parent, by=c("IDFiche"="asset_id")) %>%
  transmute(subcomponent_of = `#eq.eq_id`,
            "#eq.eq_id" = eq_id,
            eq_std = "ACCESSOIRE",
            bl_id = bl_id,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = site_id,
            description = AccessoireDescription,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = Numero,
            mfr = Marque,
            asset_id = paste("SAN-",UUID, sep =""),
            status = "in",
            #           date_in_service = "",
            comments = Remarques)


san_equip <- rbind(san_equip_parent, san_equip_acc)

write_archibus(san_equip, "./01.eq-SAN.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")


##################################
# ELE equipement
##################################

# ELE Parents

ele_equip0 <- read_excel(GMAO_file, "ELE_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
  filter (`ID Famille` != "Stations MT") %>%
  filter (`ID Famille` != "Ascenseurs/Monte-charges") %>%
  filter (`ID Famille` != "Téléphones ascenseurs")
ele_equip0$eq_id <-paste("ELE-00000-",formatC(seq.int(nrow(ele_equip0)), width=6, flag=0, format="d"),sep = "")

ele_equip_parent <- ele_equip0 %>%
  left_join(standards_equip, by=c("STANDARD D'EQUIPEMENT"="description")) %>%
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
            num_serial = "",
            modelno = "",
            subcomponent_of ="",
            mfr = "",
            asset_id = `ID Fiche`,
            status = "in",
            comments = Remarques)

# ELE Ascenseurs

ele_ass_equip0 <- read_excel(ELA_Ascenseurs_file, "Inventaire ascenseurs",col_names = TRUE, col_types = NULL, na = "", skip = 0) %>%
  filter (`Migrer` == "Oui") 
ele_ass_equip0$eq_id <-paste("ELE-00000-",formatC(seq.int(nrow(ele_ass_equip0)) + nrow(ele_equip0), width=6, flag=0, format="d"),sep = "")

ele_ass_equip <- ele_ass_equip0 %>%
  left_join(standards_equip, by=c("Domaine technique"="description")) %>%
  left_join(batiments_import, by=c("Local"="c_porte")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = ifelse(is.na(`#eqstd.eq_std`), "A DEFINIR", `#eqstd.eq_std`),
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = Type,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = "",
            modelno = 'No installation',
            subcomponent_of ="",
            mfr = Marque,
            asset_id = paste(`ID Fiche (table Electricité)`,`UUID (table Acsenseurs)`,sep =" - "),
            status = "in",
            comments = Remarques)

# ELE Accesoires

ele_equip_acc0 <- read_excel(GMAO_Acc_file, "E_acc",col_names = TRUE, col_types = NULL, na = "", )
ele_equip_acc0$eq_id <-paste("ELE-00000-",formatC(seq.int(nrow(ele_equip_acc0)) + nrow(ele_ass_equip) + nrow(ele_equip0), width=6, flag=0, format="d"), sep = "")

ele_equip_acc <- ele_equip_acc0 %>%
  inner_join(ele_equip_parent, by=c("IDFiche"="asset_id"),) %>%
  transmute(subcomponent_of = `#eq.eq_id`,
            "#eq.eq_id" = eq_id,
            eq_std = "ACCESSOIRE",
            bl_id = bl_id,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = site_id,
            description = AccessoireDescription,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = NoDeSerie,
            modelno = "Numero",
            mfr = Fournisseur,
            asset_id = paste("ELE-",UUID, sep =""),
            status = "in",
            comments = Remarques)


ele_equip <- rbind(ele_equip_parent, ele_ass_equip, ele_equip_acc)

write_archibus(ele_equip, "./01.eq-ELE.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")


##################################
# MOB equipement
##################################

# MOB Parents

mob_equip0 <- read_excel(GMAO_file, "MobLABO_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1) %>%
  filter(grepl('ok',Remarques,ignore.case=TRUE))

mob_equip0$eq_id <-paste("MOB-00000-",formatC(seq.int(nrow(mob_equip0)), width=6, flag=0, format="d"),sep = "")

mob_equip_parent <- mob_equip0 %>%
  left_join(standards_equip, by=c("STANDARD D'EQUIPEMENT"="description")) %>%
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
            subcomponent_of ="",
            mfr = Fournisseur,
            asset_id = `ID Fiche`,
            status = "in",
            #            date_in_service = `Mise en service`,
            comments = Remarques)

# MOB Accesoires
mob_equip_acc0 <- read_excel(GMAO_Acc_file, "LaboM_Acc",col_names = TRUE, col_types = NULL, na = "", )
mob_equip_acc0$eq_id <-paste("MOB-00000-",formatC(seq.int(nrow(mob_equip_acc0)) + nrow(mob_equip0), width=6, flag=0, format="d"), sep = "")

mob_equip_acc <- mob_equip_acc0 %>%
  left_join(mob_equip_parent, by=c("IDFiche"="asset_id")) %>%
  transmute(subcomponent_of = `#eq.eq_id`,
            "#eq.eq_id" = eq_id,
            eq_std = "ACCESSOIRE",
            bl_id = bl_id,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = site_id,
            description = AccessoireDescription,
            dv_id = 11500,
            dp_id = "0047",
            num_serial = "",
            mfr = "",
            asset_id = paste("MOB-",UUID, sep =""),
            status = "in",
            #           date_in_service = "",
            comments = Remarques)


mob_equip <- rbind(mob_equip_parent, mob_equip_acc)

write_archibus(mob_equip, "./01.eq-MOB.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")

##################################
# SV equipement
##################################

sv_equip <- read_excel(SV_file, "EQUIPEMENTS INFRA Actifs",col_names = TRUE, col_types = NULL, na = "", skip = 2) %>%
  mutate(CF4 = strtoi(sub("C", "", `CC Responsable`)))

sv_equip$eq_id <-paste("FSV-00000-",formatC(seq.int(nrow(sv_equip)), width=6, flag=0, format="d"),sep = "")

dp <- read_excel("./export DP.xlsx")

sv_equip_parent <- sv_equip %>%
  left_join(standards_equip, by=c("Désignation / description du standard d'équipement"="description"))%>%
  left_join(batiments_import, by=c("Local"="c_porte")) %>%
  left_join(dp, by=c(CF4="dp_id")) %>%
  transmute("#eq.eq_id" = eq_id,
            eq_std = ifelse(is.na(`#eqstd.eq_std`), "A DEFINIR", `#eqstd.eq_std`),
            bl_id = `#rm.bl_id`,
            fl_id = fl_id,
            rm_id = rm_id,
            site_id = SiteCode,
            description = `Modèle`,
            dv_id = dv_id,
            dp_id = CF4,
            num_serial = `N° série`,
            subcomponent_of ="",
            mfr = Fabricant,
            asset_id = `N° d'équipment`,
            status = ifelse(Statut != "Non conforme", "in", "out"), 
            #            date_in_service = `Mise en service`,
            comments = "")

write_archibus(sv_equip, "./01.eq-FSV.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")


##################################
# Attributs
##################################

ass_attrib <- read.csv(text="#asset_attribute_std.asset_attribute_std,title,description,asset_type", check.names=FALSE)
ass_attrib[nrow(ass_attrib)+1,] <- c("LIEU","Lieu","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TYPE","type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIAMETRE","Diametre","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VALEURHORAIRE","Valeur horaire","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("INSTALLATION","Installation","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MONOBLOCTYPE","Monobloc type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MONOBLOCNO","Monobloc no","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DEBITAIR","Débit d'air","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PRESSIONEXTRACTION","Pression extraction","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PRESSIONPULSION","Pression pulsion","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MOTEURTYPE","Moteur type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TENSION","Tension","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PUISSANCE","Puissance","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NBTOURS","Nb tours","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NBVITESSE","Nb vitesse","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("AMPERAGENOMINAL","Ampérage nominale","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VENTILATEURTYPE","Ventilateur type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PRESSIONSTATIQUE","Pression statique","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TRANSMISSIONTYPE","Transmission type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("POSITION","Position","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("MATIERE","Matière","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("COURROIEFORME","Courroie forme","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("COURROIETYPE","Courroie type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PRISEAIREDERNIERNET","Prise Aire Dernier nettoyage","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("CLIMATISEURARMOIRE","Climatiseur Armoire Marque type","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("RACCORDEMENT","Raccordement","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("FICHETECHNIQUE","Fiche technique","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("IDCONTRAT","ID contrat","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIMENSION","Dimension","L-P-H (cm)","Equipment")
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
ass_attrib[nrow(ass_attrib)+1,] <- c("INTERVENTION","Intervention","","Equipment")
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
ass_attrib[nrow(ass_attrib)+1,] <- c("CHARGEUTILE","Charge utile","Kg","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("TELASCENSEUR","No téléphone ascenseur","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("SYSTEMEURGENCE","Système d'appel d'urgence","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("IMAGE","Image Source","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ORIGINALE","Originale","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("ORIGINALE","Originale","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("PRIX","Prix","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VALEUR","Valeur d'acquisition","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("NIVEAUSECU","Niveau Sécurité","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("FOURNISSEUR","Fournisseur","","Equipment")

write_archibus(ass_attrib, "./02.asset_attrib.xlsx",
               table.header = "Asset Attribute Standards",
               sheet.name = "Sheet1")

# Attributs CHA

eq_ass_attribut_lieu <- cha_equip0 %>%
  filter (Lieu != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "LIEU",
            value = Lieu)

# Attributs CHA Accessoire

eq_ass_attribut_cha_type <- cha_equip_acc0 %>%
  filter (Type != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = Type)

eq_ass_attribut_diametre <- cha_equip_acc0 %>%
  filter (Diametre != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIAMETRE",
            value = Diametre)

eq_ass_attribut_valeurhoraire <- cha_equip_acc0 %>%
  filter (ValeurHoraire != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "VALEURHORAIRE",
            value = ValeurHoraire)

# Attributs VEN

eq_ass_attribut_installation <- ven_equip0 %>%
  filter (`Installation` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "INSTALLATION",
            value = `Installation`)

eq_ass_attribut_monobloctype <- ven_equip0 %>%
  filter (`Monobloc Type` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MONOBLOCTYPE",
            value = `Monobloc Type`)

eq_ass_attribut_monoblocno <- ven_equip0 %>%
  filter (`Monobloc no` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MONOBLOCNO",
            value = `Monobloc no`) 

eq_ass_attribut_debitair <- ven_equip0 %>%
  filter (`Débit d'air` != "") %>%
  filter (`Débit d'air` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DEBITAIR",
            value = `Débit d'air`) 

eq_ass_attribut_pressionextraction <- ven_equip0 %>%
  filter (`Pression extraction` != "") %>%
  filter (`Pression extraction` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PRESSIONEXTRACTION",
            value = `Pression extraction`) 

eq_ass_attribut_pressionpulsion <- ven_equip0 %>%
  filter (`Pression pulsion` != "") %>%
  filter (`Pression pulsion` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PRESSIONPULSION",
            value = `Pression pulsion`) 

eq_ass_attribut_moteurtype <- ven_equip0 %>%
  filter (`Moteur type` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MOTEURTYPE",
            value = `Moteur type`) 

eq_ass_attribut_tension <- ven_equip0 %>%
  filter (`Tension` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TENSION",
            value = `Tension`) 

eq_ass_attribut_puissance <- ven_equip0 %>%
  filter (`Puissance` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PUISSANCE",
            value = `Puissance`) 

eq_ass_attribut_nbtours <- ven_equip0 %>%
  filter (`Nb tours` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NBTOURS",
            value = `Nb tours`) 

eq_ass_attribut_nbvitesse <- ven_equip0 %>%
  filter (`Nb vitesse` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NBVITESSE",
            value = `Nb vitesse`) 

eq_ass_attribut_ampreagenominal <- ven_equip0 %>%
  filter (`Ampérage nominale` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "AMPERAGENOMINAL",
            value = `Ampérage nominale`) 

eq_ass_attribut_ventiallateurtype <- ven_equip0 %>%
  filter (`Ventilateur type` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "VENTILATEURTYPE",
            value = `Ventilateur type`) 

eq_ass_attribut_pressionstatique <- ven_equip0 %>%
  filter (`Pression statique` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PRESSIONSTATIQUE",
            value = `Pression statique`) 

eq_ass_attribut_transmissiontype <- ven_equip0 %>%
  filter (`Transmission type` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSMISSIONTYPE",
            value = `Transmission type`) 

eq_ass_attribut_position <- ven_equip0 %>%
  filter (`Position` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "POSITION",
            value = `Position`) 

eq_ass_attribut_matiere <- ven_equip0 %>%
  filter (`Matière` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MATIERE",
            value = `Matière`) 

eq_ass_attribut_courroieforme <- ven_equip0 %>%
  filter (`Courroie forme` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "COURROIEFORME",
            value = `Courroie forme`) 

eq_ass_attribut_courroietype <- ven_equip0 %>%
  filter (`Courroie type` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "COURROIETYPE",
            value = `Courroie type`) 

eq_ass_attribut_priseairdeniernet <- ven_equip0 %>%
  filter (`Prise Aire Dernier nettoyage` != "") %>%
  filter (`Prise Aire Dernier nettoyage` != "00.00.00") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PRISEAIREDERNIERNET",
            value = `Prise Aire Dernier nettoyage`) 

eq_ass_attribut_climatisteurarmoire <- ven_equip0 %>%
  filter (`Climatiseur Armoire Marque type` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "CLIMATISEURARMOIRE",
            value = `Climatiseur Armoire Marque type`) 

eq_ass_attribut_raccordement <- ven_equip0 %>%
  filter (`Raccordement` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RACCORDEMENT",
            value = `Raccordement`) 

eq_ass_attribut_fichetechnique <- ven_equip0 %>%
  filter (`Fiche technique` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "FICHETECHNIQUE",
            value = `Fiche technique`) 

# Attributs VEN Accessoire

eq_ass_attribut_idcontrat <- ven_equip_acc0 %>% 
  filter (`IDContrat` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "IDCONTRAT",
            value = `IDContrat`) 

eq_ass_attribut_ven_type <- ven_equip_acc0 %>%
  filter (AccessoireType != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = `AccessoireType`)

eq_ass_attribut_dimension <- ven_equip_acc0 %>%
  filter (`Dimension` != "") %>%
  filter (`Dimension` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIMENSION",
            value = `Dimension`)

eq_ass_attribut_debitair <- ven_equip_acc0 %>%
  filter (`DebitAir` != "") %>%
  filter (`DebitAir` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DEBITAIR",
            value = `DebitAir`) 

eq_ass_attribut_diametreraccordement <- ven_equip_acc0 %>%
  filter (`DiametreRaccordement` != "") %>%
  filter (`DiametreRaccordement` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIAMETRERACCORDEMENT",
            value = `DiametreRaccordement`) 

eq_ass_attribut_vitesseairprevue <- ven_equip_acc0 %>%
  filter (`VitesseAirPrevue` != "") %>%
  filter (`VitesseAirPrevue` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "VITESSEAIRPREVUE",
            value = `VitesseAirPrevue`) 

eq_ass_attribut_vitesseairmesuree <- ven_equip_acc0 %>% 
  filter (`VitesseAirMesuree` != "") %>%
  filter (`VitesseAirMesuree` != "0") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "VITESSEAIRMESUREE",
            value = `VitesseAirMesuree`) 

eq_ass_attribut_manometreposeoui <- ven_equip_acc0 %>% 
  filter (`ManometrePoseOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MANOMETREPOSEOUI",
            value = `ManometrePoseOui`)

eq_ass_attribut_manometrepression <- ven_equip_acc0 %>% 
  filter (`ManometrePression` != "") %>%
  filter (`ManometrePression` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MANOMETREPRESSION",
            value = `ManometrePression`) 

eq_ass_attribut_guillotinechapelle <- ven_equip_acc0 %>% 
  filter (`GuillotineChapelle` != "") %>%
  filter (`GuillotineChapelle` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "GUILLOTINECHAPELLE", 
            value = `GuillotineChapelle`) 

eq_ass_attribut_horlogepvgvoui <- ven_equip_acc0 %>% 
  filter (`HorlogePVGVOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "HORLOGEPVGVOUI",
            value = `HorlogePVGVOui`)

eq_ass_attribut_forsagegvoui <- ven_equip_acc0 %>%
  filter (`ForsageGVOui` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "FORSAGEGVOUI",
            value = `ForsageGVOui`) 

eq_ass_attribut_elempvoui <- ven_equip_acc0 %>% 
  filter (`ElemPVOui` != "") %>% 
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ELEMPVOUI", 
            value = `ElemPVOui`)

eq_ass_attribut_elemgvoui <- ven_equip_acc0 %>%
  filter (`ElemGVOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ELEMGVOUI",
            value = `ElemGVOui`)

eq_ass_attribut_elemautooui <- ven_equip_acc0 %>%
  filter (`ElemAUTOOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ELEMAUTOOUI",
            value = `ElemAUTOOui`)

eq_ass_attribut_elemhorsoui <- ven_equip_acc0 %>%
  filter (`ElemHORSOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ELEMHORSOUI",
            value = `ElemHORSOui`)

eq_ass_attribut_elempanneoui <- ven_equip_acc0 %>%
  filter (`ElemPanneOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ELEMPANNEOUI",
            value = `ElemPanneOui`)

eq_ass_attribut_elemnoplatine <- ven_equip_acc0 %>%
  filter (`ElemNoPlatine` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ELEMNOPLATINE",
            value = `ElemNoPlatine`)

eq_ass_attribut_dimension2 <- ven_equip_acc0 %>%
  filter (`Dimension2` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIMENSION2",
            value = `Dimension2`)

eq_ass_attribut_dimension3 <- ven_equip_acc0 %>%
  filter (`Dimension3` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIMENSION3",
            value = `Dimension3`)

eq_ass_attribut_acontroleroui <- ven_equip_acc0 %>%
  filter (`AControlerOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ACONTROLEROUI",
            value = `AControlerOui`)

eq_ass_attribut_idno <- ven_equip_acc0 %>%
  filter (`IDNo` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "IDNO",
            value = `IDNo`)

eq_ass_attribut_ven_quantite <- ven_equip_acc0 %>%
  filter (`Quantite` != "") %>%
  filter (`Quantite` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "QUANTITE",
            value = `Quantite`)

eq_ass_attribut_stock <- ven_equip_acc0 %>%
  filter (`Stock` != "") %>%
  filter (`Stock` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "STOCK",
            value = `Stock`)

eq_ass_attribut_etat <- ven_equip_acc0 %>%
  filter (`Etat` != "") %>%
  filter (`Etat` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ETAT",
            value = `Etat`)

eq_ass_attribut_intervention <- ven_equip_acc0 %>%
  filter (`Intervention` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "INTERVENTION",
            value = `Intervention`)

eq_ass_attribut_majpar <- ven_equip_acc0 %>%
  filter (`MAJPar` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "MAJPAR",
            value = `MAJPar`)

eq_ass_attribut_ctrl1 <- ven_equip_acc0 %>%
  filter (`Ctrl1` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "CTRL1",
            value = `Ctrl1`)

eq_ass_attribut_chapelleoui <- ven_equip_acc0 %>%
  filter (`ChapelleOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "CHAPELLEOUI",
            value = `ChapelleOui`)

eq_ass_attribut_iddepannge <- ven_equip_acc0 %>%
  filter (`IDDepannge` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "IDDEPANNGE",
            value = `IDDepannge`)

eq_ass_attribut_debitdifferenceoui <- ven_equip_acc0 %>%
  filter (`DebitDifferenceOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DEBITDIFFERENCEOUI",
            value = `DebitDifferenceOui`)

eq_ass_attribut_debitairmesure <- ven_equip_acc0 %>%
  filter (`DebitAirMesure` != "") %>%
  filter (`DebitAirMesure` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DEBITAIRMESURE",
            value = `DebitAirMesure`)

# Attributs SAN Accessoire

eq_ass_attribut_san_type <- san_equip_acc0 %>%
  filter (Type != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = Type)

eq_ass_attribut_valeurhoraire <- san_equip_acc0 %>%
  filter (ValeurHoraire != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "VALEURHORAIRE",
            value = ValeurHoraire)

eq_ass_attribut_diametre <- san_equip_acc0 %>%
  filter (Diametre != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIAMETRE",
            value = Diametre)

eq_ass_attribut_debitnom <- san_equip_acc0 %>%
  filter (`DebitNom` != "") %>%
  filter (`DebitNom` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DEBITNOM",
            value = `DebitNom`) 

eq_ass_attribut_debiteffectif <- san_equip_acc0 %>%
  filter (`DebitEffectif` != "") %>%
  filter (`DebitEffectif` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DEBITEFFECTIF",
            value = `DebitEffectif`) 

eq_ass_attribut_position <- san_equip_acc0 %>%
  filter (`Position` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "POSITION",
            value = `Position`) 

# Attributs ELE

eq_ass_attribut_ele_type <- ele_equip_acc0 %>%
  filter (`` != "") %>%	
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = `Type`)

eq_ass_attribut_nbreheure <- ele_equip_acc0 %>%
  filter (`NbreHeure` != "") %>%
  filter (`NbreHeure` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NBREHEURE",
            value = `NbreHeure`)

eq_ass_attribut_ele_quantite <- ele_equip_acc0 %>%
  filter (`Quantite` != "") %>%
  filter (`Quantite` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "QUANTITE",
            value = `Quantite`)

eq_ass_attribut_nbrepersonnes <- ele_equip_acc0 %>%
  filter (`NbrePersonnes` != "") %>%
  filter (`NbrePersonnes` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "NBREPERSONNES",
            value = `NbrePersonnes`)

eq_ass_attribut_puissance <- ele_equip_acc0 %>%
  filter (`Puissance` != "") %>%
  filter (`Puissance` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PUISSANCE",
            value = `Puissance`)

eq_ass_attribut_anneeenservice <- ele_equip_acc0 %>%
  filter (`AnneeEnService` != "") %>%
  filter (`AnneeEnService` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ANNEEENSERVICE",
            value = `AnneeEnService`)

eq_ass_attribut_anneeconstruction <- ele_equip_acc0 %>%
  filter (`AnneeConstruction` != "") %>%
  filter (`AnneeConstruction` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ANNEECONSTRUCTION",
            value = `AnneeConstruction`)

eq_ass_attribut_relaifournisseur <- ele_equip_acc0 %>%
  filter (`RelaiFournisseur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIFOURNISSEUR",
            value = `RelaiFournisseur`)

eq_ass_attribut_autotransfooui <- ele_equip_acc0 %>%
  filter (`AutotransfoOui` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "AUTOTRANSFOOUI",
            value = `AutotransfoOui`)

eq_ass_attribut_transforeseau <- ele_equip_acc0 %>%
  filter (`TransfoReseau` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFORESEAU",
            value = `TransfoReseau`)

eq_ass_attribut_transfocellule <- ele_equip_acc0 %>%
  filter (`TransfoCellule` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOCELLULE",
            value = `TransfoCellule`)

eq_ass_attribut_relaitype <- ele_equip_acc0 %>%
  filter (`RelaiType` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAITYPE",
            value = `RelaiType`)

eq_ass_attribut_unominale <- ele_equip_acc0 %>%
  filter (`UNominale` != "") %>%
  filter (`UNominale` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "UNOMINALE",
            value = `UNominale`)

eq_ass_attribut_transfouprimaire <- ele_equip_acc0 %>%
  filter (`TransfoUPrimaire` != "") %>%
  filter (`TransfoUPrimaire` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOUPRIMAIRE",
            value = `TransfoUPrimaire`)

eq_ass_attribut_transfousecondairereglee <- ele_equip_acc0 %>%
  filter (`TransfoUSecondaireReglee` != "") %>%
  filter (`TransfoUPrimaire` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOUSECONDAIREREGLEE",
            value = `TransfoUSecondaireReglee`)

eq_ass_attribut_transforemplissage <- ele_equip_acc0 %>%
  filter (`TransfoRemplissage` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOREMPLISSAGE",
            value = `TransfoRemplissage`)

eq_ass_attribut_transfocouplage <- ele_equip_acc0 %>%
  filter (`TransfoCouplage` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOCOUPLAGE",
            value = `TransfoCouplage`)

eq_ass_attribut_transfoucc <- ele_equip_acc0 %>%
  filter (`TransfoUcc` != "") %>%
  filter (`TransfoUcc` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOUCC",
            value = `TransfoUcc`)

eq_ass_attribut_uprimaire <- ele_equip_acc0 %>%
  filter (`UPrimaire` != "") %>%
  filter (`UPrimaire` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "UPRIMAIRE",
            value = `UPrimaire`)

eq_ass_attribut_transfousecondaire1 <- ele_equip_acc0 %>%
  filter (`TransfoUSecondaire1` != "") %>%
  filter (`TransfoUSecondaire1` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOUSECONDAIRE1",
            value = `TransfoUSecondaire1`)

eq_ass_attribut_transfousecondaire2 <- ele_equip_acc0 %>%
  filter (`TransfoUSecondaire2` != "") %>%
  filter (`TransfoUSecondaire2` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOUSECONDAIRE2",
            value = `TransfoUSecondaire2`)

eq_ass_attribut_transfousecondaire3 <- ele_equip_acc0 %>%
  filter (`TransfoUSecondaire3` != "") %>%
  filter (`TransfoUSecondaire3` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOUSECONDAIRE3",
            value = `TransfoUSecondaire3`)

eq_ass_attribut_transfousecondaire4 <- ele_equip_acc0 %>%
  filter (`TransfoUSecondaire4` != "") %>%
  filter (`TransfoUSecondaire4` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOUSECONDAIRE4",
            value = `TransfoUSecondaire4`)

eq_ass_attribut_transfousecondaire5 <- ele_equip_acc0 %>%
  filter (`TransfoUSecondaire5` != "") %>%
  filter (`TransfoUSecondaire5` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOUSECONDAIRE5",
            value = `TransfoUSecondaire5`)

eq_ass_attribut_transfoperteavide <- ele_equip_acc0 %>%
  filter (`TransfoPerteAVide` != "") %>%
  filter (`TransfoPerteAVide` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOPERTEAVIDE",
            value = `TransfoPerteAVide`)

eq_ass_attribut_transfoperteencharge <- ele_equip_acc0 %>%
  filter (`TransfoPerteEnCharge` != "") %>%
  filter (`TransfoPerteEnCharge` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOPERTEENCHARGE",
            value = `TransfoPerteEnCharge`)

eq_ass_attribut_transfomassetotale <- ele_equip_acc0 %>%
  filter (`TransfoMasseTotale` != "") %>%
  filter (`TransfoMasseTotale` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOMASSETOTALE",
            value = `TransfoMasseTotale`)

eq_ass_attribut_transfomaddehuile <- ele_equip_acc0 %>%
  filter (`TransfoMaddeHuile` != "") %>%
  filter (`TransfoMaddeHuile` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOMADDEHUILE",
            value = `TransfoMaddeHuile`)

eq_ass_attribut_transfoprealarme <- ele_equip_acc0 %>%
  filter (`TransfoPreAlarme` != "") %>%
  filter (`TransfoPreAlarme` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOPREALARME",
            value = `TransfoPreAlarme`)

eq_ass_attribut_transfoalarme <- ele_equip_acc0 %>%
  filter (`TransfoAlarme` != "") %>%
  filter (`TransfoAlarme` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TRANSFOALARME",
            value = `TransfoAlarme`)

eq_ass_attribut_disjoncteurinominal <- ele_equip_acc0 %>%
  filter (`DisjoncteurINominal` != "") %>%
  filter (`DisjoncteurINominal` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DISJONCTEURINOMINAL",
            value = `DisjoncteurINominal`)

eq_ass_attribut_disjoncteurpouvoirdecoupure <- ele_equip_acc0 %>%
  filter (`DisjoncteurPouvoirDeCoupure` != "") %>%
  filter (`DisjoncteurPouvoirDeCoupure` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DISJONCTEURPOUVOIRDECOUPURE",
            value = `DisjoncteurPouvoirDeCoupure`)

eq_ass_attribut_relaiithernmiqueregle <- ele_equip_acc0 %>%
  filter (`RelaiIThernmiqueRegle` != "") %>%
  filter (`RelaiIThernmiqueRegle` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIITHERNMIQUEREGLE",
            value = `RelaiIThernmiqueRegle`)

eq_ass_attribut_relaideclancheinstantane <- ele_equip_acc0 %>%
  filter (`RelaiDeclancheInstantane` != "") %>%
  filter (`RelaiDeclancheInstantane` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIDECLANCHEINSTANTANE",
            value = `RelaiDeclancheInstantane`)

eq_ass_attribut_relaiconstantedetemps <- ele_equip_acc0 %>%
  filter (`RelaiConstanteDeTemps` != "") %>%
  filter (`RelaiConstanteDeTemps` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAICONSTANTEDETEMPS",
            value = `RelaiConstanteDeTemps`)

eq_ass_attribut_relaiireponse <- ele_equip_acc0 %>%
  filter (`RelaiIReponse` != "") %>%
  filter (`RelaiIReponse` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIIREPONSE",
            value = `RelaiIReponse`)

eq_ass_attribut_relaitemporisation <- ele_equip_acc0 %>%
  filter (`RelaiTemporisation` != "") %>%
  filter (`RelaiTemporisation` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAITEMPORISATION",
            value = `RelaiTemporisation`)

eq_ass_attribut_relaiinominal <- ele_equip_acc0 %>%
  filter (`RelaiINominal` != "") %>%
  filter (`RelaiINominal` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "RELAIINOMINAL",
            value = `RelaiINominal`)
# Assenceurs

eq_ass_attribut_dimensions <- ele_ass_equip0 %>%
  filter (`Dimensions L-P-H (cm)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DIMENSIONS",
            value = `Dimensions L-P-H (cm)`)

eq_ass_attribut_charge <- ele_ass_equip0 %>%
  filter (`Charge utile (kg)` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "CHARGEUTILE",
            value = `Charge utile (kg)`)

eq_ass_attribut_telascenseur <- ele_ass_equip0 %>%
  filter (`No téléphone ascenseur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TELASCENSEUR",
            value = `No téléphone ascenseur`)

eq_ass_attribut_sytemeurgence <- ele_ass_equip0 %>%
  filter (`Système d'appel d'urgence` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "SYSTEMEURGENCE",
            value = `Système d'appel d'urgence`)



# Mobiliers

eq_ass_attribut_mob_type <- ele_mob_equip0 %>%
  filter (`Type` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "TYPE",
            value = `Type`)

eq_ass_attribut_image <- ele_mob_equip0 %>%
  filter (`Image source` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "IMAGE",
            value = `Image source`)

eq_ass_attribut_originale <- ele_mob_equip0 %>%
  filter (`Originale?` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "ORIGINALE",
            value = `Originale?`)

# Mobilier Accesoires

eq_ass_attribut_prix <- mob_equip_acc0 %>%
  filter (`PrixUnitaire` != "") %>%
  filter (`PrixUnitaire` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "PRIX",
            value = `PrixUnitaire`)

eq_ass_attribut_mob_quantite <- mob_equip_acc0 %>%
  filter (`Quantite` != "") %>%
  filter (`Quantite` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "QUANTITE",
            value = `Quantite`)

# SV

eq_ass_attribut_valeur <- sv_equip_acc0 %>%
  filter (`Valeur d'acquisition` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "VALEUR",
            value = `Valeur d'acquisition`)

eq_ass_attribut_valeur <- sv_equip_acc0 %>%
  filter (`Valeur d'acquisition` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "VALEUR",
            value = `Valeur d'acquisition`)

eq_ass_attribut_fournisseur <- sv_equip_acc0 %>%
  filter (`Fournisseur` != "") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "FOURNISSEUR",
            value = `Fournisseur`)


eq_ass_attribut <- rbind(eq_ass_attribut_lieu, 
                         eq_ass_attribut_san_type,
                         eq_ass_attribut_ele_type,
                         eq_ass_attribut_cha_type,
                         eq_ass_attribut_ven_type,
                         eq_ass_attribut_diametre, 
                         eq_ass_attribut_valeurhoraire,
                         eq_ass_attribut_installation,
                         eq_ass_attribut_monobloctype,
                         eq_ass_attribut_monoblocno,
                         eq_ass_attribut_debitair,
                         eq_ass_attribut_pressionextraction,
                         eq_ass_attribut_pressionpulsion,
                         eq_ass_attribut_moteurtype,
                         eq_ass_attribut_tension,
                         eq_ass_attribut_puissance,
                         eq_ass_attribut_nbtours,
                         eq_ass_attribut_nbvitesse,
                         eq_ass_attribut_ampreagenominal,
                         eq_ass_attribut_ventiallateurtype,
                         eq_ass_attribut_pressionstatique,
                         eq_ass_attribut_transmissiontype,
                         eq_ass_attribut_position,
                         eq_ass_attribut_matiere,
                         eq_ass_attribut_courroieforme,
                         eq_ass_attribut_courroietype,
                         eq_ass_attribut_priseairdeniernet,
                         eq_ass_attribut_climatisteurarmoire,
                         eq_ass_attribut_raccordement,
                         eq_ass_attribut_fichetechnique,
                         eq_ass_attribut_idcontrat,
                         eq_ass_attribut_dimension,
                         eq_ass_attribut_diametreraccordement,
                         eq_ass_attribut_vitesseairprevue,
                         eq_ass_attribut_vitesseairmesuree,
                         eq_ass_attribut_manometreposeoui,
                         eq_ass_attribut_manometrepression,
                         eq_ass_attribut_guillotinechapelle,
                         eq_ass_attribut_horlogepvgvoui,
                         eq_ass_attribut_forsagegvoui,
                         eq_ass_attribut_elempvoui,
                         eq_ass_attribut_elemgvoui,
                         eq_ass_attribut_elemautooui,
                         eq_ass_attribut_elemhorsoui,
                         eq_ass_attribut_elempanneoui,
                         eq_ass_attribut_elemnoplatine,
                         eq_ass_attribut_dimension2,
                         eq_ass_attribut_dimension3,
                         eq_ass_attribut_acontroleroui,
                         eq_ass_attribut_idno,
                         eq_ass_attribut_ven_quantite,
                         eq_ass_attribut_ele_quantite,
                         eq_ass_attribut_mob_quantite,
                         eq_ass_attribut_stock,
                         eq_ass_attribut_etat,
                         eq_ass_attribut_intervention,
                         eq_ass_attribut_majpar,
                         eq_ass_attribut_ctrl1,
                         eq_ass_attribut_chapelleoui,
                         eq_ass_attribut_iddepannge,
                         eq_ass_attribut_debitdifferenceoui,
                         eq_ass_attribut_debitairmesure,
                         eq_ass_attribut_debitnom,
                         eq_ass_attribut_debiteffectif,
                         eq_ass_attribut_position,
                         eq_ass_attribut_nbreheure,
                         eq_ass_attribut_nbrepersonnes,
                         eq_ass_attribut_anneeenservice,
                         eq_ass_attribut_anneeconstruction,
                         eq_ass_attribut_relaifournisseur,
                         eq_ass_attribut_autotransfooui,
                         eq_ass_attribut_transforeseau,
                         eq_ass_attribut_transfocellule,
                         eq_ass_attribut_relaitype,
                         eq_ass_attribut_unominale,
                         eq_ass_attribut_transfouprimaire,
                         eq_ass_attribut_transfousecondairereglee,
                         eq_ass_attribut_transforemplissage,
                         eq_ass_attribut_transfocouplage,
                         eq_ass_attribut_transfoucc,
                         eq_ass_attribut_uprimaire,
                         eq_ass_attribut_transfousecondaire1,
                         eq_ass_attribut_transfousecondaire2,
                         eq_ass_attribut_transfousecondaire3,
                         eq_ass_attribut_transfousecondaire4,
                         eq_ass_attribut_transfousecondaire5,
                         eq_ass_attribut_transfoperteavide,
                         eq_ass_attribut_transfoperteencharge,
                         eq_ass_attribut_transfomassetotale,
                         eq_ass_attribut_transfomaddehuile,
                         eq_ass_attribut_transfoprealarme,
                         eq_ass_attribut_transfoalarme,
                         eq_ass_attribut_disjoncteurinominal,
                         eq_ass_attribut_disjoncteurpouvoirdecoupure,
                         eq_ass_attribut_relaiithernmiqueregle,
                         eq_ass_attribut_relaideclancheinstantane,
                         eq_ass_attribut_relaiconstantedetemps,
                         eq_ass_attribut_relaiireponse,
                         eq_ass_attribut_relaitemporisation,
                         eq_ass_attribut_relaiinominal,
                         eq_ass_attribut_charge,
                         eq_ass_attribut_telascenseur,
                         eq_ass_attribut_sytemeurgence,
                         eq_ass_attribut_mob_type,
                         eq_ass_attribut_image,
                         eq_ass_attribut_originale,
                         eq_ass_attribut_prix,
                         eq_ass_attribut_valeur.
                         eq_ass_attribut_niveausecu,
                         eq_ass_attribut_fournisseur)

write_archibus(eq_ass_attribut, "./03.eq_asset_attrib.xlsx",
               table.header = "Equipment Asset Attributes",
               sheet.name = "Sheet1")