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
  
  # save the workbook to an Excel file
  xlsx::saveWorkbook(wb, filename)
}


toArchibusStatus <- function(etat) {
  
  case_when(etat == "FAUX"   ~ "in",
            etat == "VRAI"   ~ "out",
            etat == ""       ~ "",
            TRUE             ~ "out")
}



standards_equip <- read_excel("./GMAO_4D_Export2021_MAPPING_eqstd.xlsx", "Standards d'équipements") %>%
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

cha_equip0 <- read_excel("./GMAO_4D_Export2021_MAPPING_eqstd.xlsx", "CHA_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1)
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

cha_equip_acc0 <- read_excel("./4D_GMAO_Accessoires_avec UUID_4-3-22.xlsx", "C_acc",col_names = TRUE, col_types = NULL, na = "", )
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
            status = "",
            comments = Remarques)


cha_equip <- rbind(cha_equip_parent, cha_equip_acc)

write_archibus(cha_equip, "./01.eq-CHA.xlsx",
               table.header = "Equipment",
               sheet.name = "Equipment")

##################################
# VEN equipement
##################################

# VEN Parents

ven_equip0 <- read_excel("./GMAO_4D_Export2021_MAPPING_eqstd.xlsx", "VEN_tout",col_names = TRUE, col_types = NULL, na = "", skip = 1)
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

# CHA Accesoires

ven_equip_acc0 <- read_excel("./4D_GMAO_Accessoires_avec UUID_4-3-22.xlsx", "V_acc",col_names = TRUE, col_types = NULL, na = "", )
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
            status = "",
#           date_in_service = "",
            comments = Remarques)


ven_equip <- rbind(ven_equip_parent, ven_equip_acc)

write_archibus(ven_equip, "./01.eq-VEN.xlsx",
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
ass_attrib[nrow(ass_attrib)+1,] <- c("DIMENSION","Dimension","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("DIAMETRERACCORDEMENT","Diametre raccordement","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VitesseAirPrevue","Vitesse air prévue","","Equipment")
ass_attrib[nrow(ass_attrib)+1,] <- c("VitesseAirMesuree","Vitesse air mesurée","","Equipment")
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

eq_ass_attribut_type <- cha_equip_acc0 %>%
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

eq_ass_attribut_type <- ven_equip_acc0 %>%
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
  filter (`Débit d'air` != "") %>%
  filter (`Débit d'air` != "0") %>%
  transmute("#eq_asset_attribute.eq_id" = eq_id,
            "asset_attribute_std" = "DEBITAIR",
            value = `Débit d'air`) 

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

eq_ass_attribut_quantite <- ven_equip_acc0 %>%
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


eq_ass_attribut <- rbind(eq_ass_attribut_lieu, 
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
                         eq_ass_attribut_quantite,
                         eq_ass_attribut_stock,
                         eq_ass_attribut_etat,
                         eq_ass_attribut_intervention,
                         eq_ass_attribut_majpar,
                         eq_ass_attribut_ctrl1,
                         eq_ass_attribut_chapelleoui,
                         eq_ass_attribut_iddepannge,
                         eq_ass_attribut_debitdifferenceoui,
                         eq_ass_attribut_debitairmesure)

write_archibus(eq_ass_attribut, "./03.eq_asset_attrib.xlsx",
               table.header = "Equipment Asset Attributes",
               sheet.name = "Sheet1")