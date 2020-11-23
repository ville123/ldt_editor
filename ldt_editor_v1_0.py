# Valonjakotiedostojen muokkausohjelma.
# Ohjelmalla muokataan ldt-tiedostopäätteellä olevia valonjakotiedostoja.
# LDT-tiedoston rivien merkitys esim: https://docs.agi32.com/PhotometricToolbox/Content/Open_Tool/eulumdat_file_format.htm

# Ohjelma tehty keväällä 2020 OAMK:n yrityslähtöisenä projektityönä.

import math, os, re, shutil, threading
import tkinter as tk
from datetime import datetime
from shutil import copyfile
from tkinter import filedialog, messagebox, StringVar, IntVar, Checkbutton
from tkinter import *
from tkinter.simpledialog import askstring
from tkinter.messagebox import showinfo
from tkinter import font as tkFont

from PIL import Image, ImageTk

today = datetime.today().strftime('%Y-%m-%d')
file_extensions = ('.ldt', '.LDT')

# Tiedoston versionumeron muutos automaattisesti.
# Saa parametrinä vanhan versionumeron (vX.Y tai XvY), palauttaa uuden numeron (vX.Y).
def changeVersion(oldNumber):
    temp = re.findall(r'\d+', oldNumber)
    newY = int(temp[1]) + 1
    if newY > 9:
        newY = 0
        newX = int(temp[0]) + 1
    else:
        newX = temp[0]
    return ('v' + str(newX) + '.' + str(newY))
   
# hakemiston ja alihakemistojen tiedostojen zippaus taustasäikeessä
class AsyncZip(threading.Thread):
    def __init__(self, zip_in, zip_out):
        threading.Thread.__init__(self)
        self.zip_in = zip_in
        self.zip_out = zip_out

    def run(self):
        try:
            shutil.make_archive(self.zip_out, 'zip', self.zip_in)
        except:
            print("FILE NOT FOUND. CANNOT ZIP THE FOLDER.")
            pass

class Frames():
    # callback funktio joka tarkistaa ettei syöte ole liian pitkä (max 78 merkkiä)
    def input_max_78(self, text):
        if text == "": return True
        else:
            if len(text) < 79: return True
            else: return False

    # Syötekenttien syötteiden tarkistuksia:

    # käyttäjä-syötekentän tarkistus
    # (hyväksytään 65 merkkiä, jotta 13 merkkiä jää date:lle kun date ja user kentät yhdistetään 
    # ldt-tiedoston samalle riville)
    def input_max_65(self, text):
        if text == "": return True
        else: 
            if len(text) < 66: return True
            else: return False

    # syöte max 24 merkkiä
    def input_max_24(self, text):
        if text == "": return True
        else: 
            if len(text) < 25: return True
            else: return False

    # syöte int, max 6 merkkiä
    def int_max_6(self, text):
        if text == "": return True
        else:
            if len(text) < 7:
                try:
                    value = int(text)
                except ValueError:
                    return False # syötteessä vääriä merkkejä
                return 0 <= value # syöte ok
            else: return False # liian pitkä syöte

    # syöte int, max 16 merkkiä
    def int_max_16(self, text): 
        if text == "": return True
        else:
            if len(text) < 17:
                try:
                    value = int(text)
                except ValueError:
                    return False # syötteessä vääriä merkkejä
                return 0 <= value # syöte ok
            else: return False # liian pitkä syöte

    # syöte float, max 4 char, kerroin-optio
    # (merkitään ylärajaksi 5 jotta saadaan desim.piste mukaan)
    def light_output_ratio_ok(self, text):
        if text == "": return True
        else:
            if re.search('^((?![a-öA-Ö]).)*$', text): # kirjaimia ei hyväksytä
                if re.search('^[^\*]*$', text): # ei sisällä *:aa
                    if len(text) < 7: # max 6 merkkiä
                        try: value = float(text)
                        except ValueError: return False # syötettiin muu kuin float
                        return 0 <= value < 1000 # syöte ok
                    else: return False # syöte liian pitkä
                else: # syötteessä *
                    if len(text) > 1: # syötteessä muutakin kuin *
                        if len(text) < 7:
                            try: value = float(text[1:])
                            except ValueError: return False
                            return True # *:n jälkeen kelvollinen luku
                        else: return False # syöte liian pitkä
                    else: return True # syötetty pelkkä *
            else: return False # syötettiin kirjain

    # luminous flux oikeellisuus, syöte float max 12 char, kerroinmahdollisuus
    def lum_flux_ok(self, text):
        if text == "": return True
        else:
            if re.search('^((?![a-öA-Ö]).)*$', text): # kirjaimia ei hyväksytä
                if re.search('^[^\*]*$', text): # ei sisällä *:aa
                    if len(text) < 13: # max 12 merkkiä
                        try: value = float(text)
                        except ValueError: return False # syötettiin muu kuin float
                        return 0 <= value # syöte ok
                    else: return False # syöte liian pitkä
                else: # syötteessä *
                    if len(text) > 1: # syötteessä muutakin kuin *
                        if len(text) < 13:
                            try: value = float(text[1:])
                            except ValueError: return False
                            return True # *:n jälkeen kelvollinen luku
                        else: return False # syöte liian pitkä
                    else: return True # syötetty pelkkä *
            else: return False # syötettiin kirjain

    # tarkistetaan symmetria-kentän syötteen oikeellisuus, oltava 0,1,2,3 tai 4
    def symmetry_is_ok(self, text):
        if text == "" or text == "0" or text == "1" or text == "2" or text == "3" or text == "4": return True
        else: return False

    # teho-kentän syötteen oikeellisuus (float, kerroin)
    def watt_is_ok(self, text):
        if text == "": return True
        else:
            if re.search('^((?![a-öA-Ö]).)*$', text): # kirjaimia ei hyväksytä
                if re.search('^[^\*]*$', text): # ei sisällä *:aa
                    if len(text) < 9: # max 8 merkkiä
                        try: value = float(text)
                        except ValueError: return False # syötettiin muu kuin float
                        return 0 <= value # syöte ok
                    else: return False # syöte liian pitkä
                else: # syötteessä *
                    if len(text) > 1: # syötteessä muutakin kuin *
                        if len(text) < 9:
                            try: value = float(text[1:])
                            except ValueError: return False
                            return True # *:n jälkeen kelvollinen luku
                        else: return False # syöte liian pitkä
                    else: return True # syötetty pelkkä *
            else: return False # syötettiin kirjain

    # callback valaisimen ja valaistun alueen pituuden, leveyden, korkeuden syötteen oikeellisuuden tarkastamiseksi
    # syötteen pitää olla int ja ei-negatiivinen.
    def length_is_ok(self, text):
        if text == "": return True
        else:
            if len(text) < 5:
                try:
                    value = int(text)
                except ValueError:
                    return False # syöte sisältää vääriä merkkejä
                return 0 <= value # syöte ok
            else: return False # liian pitkä syöte


    # switchcasen korvaaja jolla saadaan tiettyä entry-kenttää vastaava rivinumero LDT-tiedostossa.
    # (entry_listin index : rivinumero ldt-tiedostossa)
    # HUOM: jos entry-kenttien järjestys muuttuu tai kenttiä lisätään tai poistetaan, niin on päivitettävä keys_for_entries:iä!
    def keys_for_entries(self, argument):
        switcher = {
            0: 0,   # rivi 1 (valaisin - valmistaja)
            1: 8,   # rivi 9 (valaisin - nimi)
            2: 11,  # rivi 12 (mittaus - pvm ja käyttäjä)
            3: 12,  # rivi 13 (valaisimen pituus tai halkaisija) 
            4: 13,  # rivi 14 (valaisimen leveys)
            5: 14,  # rivi 15 (valaisimen korkeus)
            6: 15,  # rivi 16 (valaistun alueen pituus tai halkaisija)
            7: 16,  # rivi 17 (valaistun alueen leveys)
            8: 17,  # rivi 18 (valaistun alueen korkeus)

            9: 2,   # rivi 3 (symmetry indicator)

            10: 22, # rivi 23 (light output ratio luminaire)
            11: 23, # rivi 24 (conversion factor for luminous intensities)
            12: 27, # rivi 28 (tuotekoodi)
            13: 28, # rivi 29 (valovirta)
            14: 29, # rivi 30 (värilämpötila)
            15: 30, # rivi 31 (värintoistoindeksi) 
            16: 31, # rivi 32 (teho)
        }
        return switcher.get(argument, "")

    # luetaan alkup.tiedosto, muutetaan rivit, kirjoitetaan uuteen tiedostoon
    def editFile(self, oldPath, newPath):
        # luetaan alkup.tiedosto, jos se on ldt-tiedosto
        if oldPath.endswith(file_extensions):
            try:
                with open(oldPath, 'r') as original_ldt:
                    original_lines = original_ldt.readlines()
            except:
                print("EXCEPTION in editFiles: Could not read files in folder")               
            numberOfLinesInOriginalFile = len(original_lines)

        entry_list = [child for child in self.newval.winfo_children()
                        if isinstance(child, Entry)] # kaikki Entryt

        user_input_entries = {} # dictionary johon tulee käyttäjän syöttämät entryt
        for entry_object in entry_list:
            # lisätään dictionaryyn ne entryt joihin on kirjoitettu jotain
            if entry_object.get():
                entry_index = entry_list.index(entry_object) # kyseistä entryä vastaava indeksi
                # kerroin entryssä: jos syöte alkaa *:lla ja ei sisällä a-zA-Z merkkejä.
                if re.search("^\*[^a-öA-Ö]", entry_object.get()):
                    # regex: tiputetaan * pois, tallennetaan muuttujaan entry_multiplier
                    entry_multiplier = re.sub("\*", "", entry_object.get())
                    entry_old = original_lines[self.keys_for_entries(entry_index)]

                    try:
                        # entry_new on vanha rivi kerrottuna multiplierillä, talletetaan dictionaryyn
                        entry_new = float(entry_old) * float(entry_multiplier)
                        user_input_entries.update({self.keys_for_entries(entry_index) : str(entry_new)})
                    except:
                        showinfo('ERROR', 'Possible error detected on multiplying old value.')
                else:
                    if entry_object.get()[0] == '*': # Jos käyttäjä syöttänyt vain *-merkin -> virhetilanne
                        if root.flag_multiplier_only:
                            showinfo("Wrong input", "Only * (no value) inserted to one or more entry fields.\nEditing done for valid values only.")
                            root.flag_multiplier_only = False # näytetään msgbox vain kerran per muokkauskerta, eikä jokaiselle tiedostolle erikseen 
                    else: user_input_entries.update({self.keys_for_entries(entry_index) : entry_object.get()})

        if oldPath.endswith(file_extensions):
            # luodaan uuden ldt:n sisältö: muokatut rivit user_input_entrystä ja loput original_linesistä
            rivi_index = 0  # vastaa alkuperäisen tiedoston rivimäärää
            edited_ldt_list = []
            # käydään alkup tiedoston kaikki rivit läpi
            while rivi_index < numberOfLinesInOriginalFile:
                    # jos user_input_entries key sisältää rivi_indeksin, lisätään sen value edited_ldt_list:hen
                    # (eli käyttäjä syöttänyt arvon kyseistä riviä vastaavaan entry-kenttään)
                    if rivi_index in user_input_entries.keys():
                        if rivi_index == 11: # syöte käyttäjä-kenttään
                            # yhdistetään alkuperäinen date ja syötetty user yhdeksi riviksi
                            original_date = re.split('\/', original_lines[rivi_index])
                            edited_ldt_list.append(original_date[0] + '/ ' + user_input_entries[rivi_index] + '\n')
                        elif rivi_index == 17: # syötettiin valaistun alueen korkeus
                            # sama syöte ldt:n riveille 18-21 (kaikkiin c-plane vaihtoehtoihin)
                            for x in range(4):
                                edited_ldt_list.append(user_input_entries[17] + '\n')
                                if x < 3: rivi_index += 1 # viimeisellä kerralla ei kasvateta indeksiä ettei hyppää rivin yli
                        else: # syötettiin muu kuin käyttäjä tai valaistun alueen korkeus
                            edited_ldt_list.append(user_input_entries[rivi_index] + '\n')
                    # muuten lisätään edited_ldt_list:hen rivi_indeksiä vastaava rivi original_linesistä
                    else: # ldt:n rivit joihin ei tehdä muutoksia
                        edited_ldt_list.append(original_lines[rivi_index])
                    rivi_index += 1

            # ldt-tiedoston rivien 8 ja 10 jäätävä tyhjäksi
            edited_ldt_list[7] = '\n'
            edited_ldt_list[9] = '\n'

            # Varmistetaan ettei kerroinoption sisältävien rivien syötteen merkkimäärä ylitä suurinta sallittua.
            # Jos luvussa on liikaa desimaaleja, pyöristetään luku siten ettei rivin sallittu merkkimäärä ylity.

            # rivi 32 - Wattage ( max char 8 )
            if len(edited_ldt_list[31]) > 9: # rivillä saa olla 8 merkkiä sisältäen desim-piste, linebreakista lisää yksi merkki
                line_after_rounding = self.rounding_method(7, edited_ldt_list[31]) # 7 on riville sallittu merkkien maksimimäärä miinus desim.pisteelle varattu merkki
                edited_ldt_list[31] = str(line_after_rounding) + '\n'

            # rivi 23 - light output ratio luminaire: 
            # eulumdat-listan mukaan number of chars = 4, mutta 
            # ldt-tiedostoissa on 5 merkkiä esim 100.00 -> käytetään viittä merkkiä
            if len(edited_ldt_list[22]) > 6:
                line_after_rounding = self.rounding_method(4, edited_ldt_list[22])
                edited_ldt_list[22] = str(line_after_rounding) + '\n'

            # rivi 24 - conversion factor for luminous intensities (max char 6)
            if len(edited_ldt_list[23]) > 7:
                line_after_rounding = self.rounding_method(5, edited_ldt_list[23])
                edited_ldt_list[23] = str(line_after_rounding) + '\n'

            # rivi 29 - total luminous flux of lamps ( max char 12 )
            if len(edited_ldt_list[28]) > 13:
                line_after_rounding = self.rounding_method(11, edited_ldt_list[28])
                edited_ldt_list[28] = str(line_after_rounding) + '\n'

            # rivi 11 - File name (muutetaan vanhan tiedoston versionumero, joka näkyy photoviewissä kohdassa Mittaus: Koodi)
            old_version_number = re.findall('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', edited_ldt_list[10])
            # try-rakenne tilanteisiin joissa rivillä ei ole validia versionumeroa
            try:
                if self.autoversion_var.get(): # tiedostojen nimien automaattinen versionumerointi valittu
                    new_line_11 = re.sub('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', changeVersion(str(old_version_number)), edited_ldt_list[10])
                else: # käyttäjän syöttämä versionumero
                    new_line_11 = re.sub('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', 'v' + str(self.version_number_from_user), edited_ldt_list[10])
                    if not old_version_number: raise Exception
                edited_ldt_list[10] = new_line_11
            except:
                # Tiedoston rivillä 11 väärä formaatti versionumerossa, eikä sitä voida muuttaa. Käyttäjä tekee muutoksen manuaalisesti.
                if root.flag_line_11:
                    showinfo('Warning (11_1)', 'Version number not in correct format in original LDT-file (line 11) and cannot be edited')
                    root.flag_line_11 = False
                
    
            # rivi 9 - Luminaire name 
            old_version_number = re.findall('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', original_lines[8]) # versionumero vanhasta rivistä
            # try-rakenne tilanteisiin joissa rivillä ei ole validia versionumeroa
            try:
                if self.autoversion_var.get():
                    new_version_number = changeVersion(str(old_version_number)) # uusi versionumero
                else:
                    new_version_number =  'v' + str(self.version_number_from_user)
                    if not old_version_number: raise Exception
                try: 
                    if 8 in user_input_entries: # syötettiin luminaire name = dictionary sisältää keyn 8
                        # lisätään new_versio_number edited_ldt_list[8]:aan
                        edited_ldt_list[8] = edited_ldt_list[8].rstrip('\n') + ' ' + new_version_number + '\n'
                    else: # ei syötetty luminaire namea
                        # korvataan vanha versionumero uudella (pitää korvata ettei vanha numero jää riville)
                        edited_ldt_list[8] = re.sub('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', new_version_number, edited_ldt_list[8])
                except:
                    showinfo('warning (9_0)', 'unknown error when modifying luminaire name')
            except:
                if root.flag_line_9:
                    showinfo('Warning (9_1)', 'Version number not in correct format in original LDT-file (line 9) and cannot be edited')
                    root.flag_line_9 = False
                

            # kirjoitetaan tiedostoon kaikki rivit (muokatut ja ne jotka säilyvät ennallaan)
            try:
                with open(newPath, 'w+') as edited_file:
                    for line in edited_ldt_list:
                        edited_file.write(line)
            except:
                if root.flag_writing_error:
                    showinfo('Error in writing to file', 'Error in writing to file.\nPlease try again.')
                    root.flag_writing_error = False

    # pyöristysfunktio (ettei kertoimen sisältävien rivien suurin sallittu merkkimäärä ylity).
    # funktio saa rivin merkkien maksimimäärän ja riville kirjoitettavaksi tarkoitetun merkkijonon (float-luvun).
    # funktio palauttaa merkkijonon pyöristettynä oikean pituiseksi.
    def rounding_method(self, maxchar, round_this):
        how_many_decimals = int(maxchar) - int(len(str(math.trunc(float(round_this)))))
        return round(float(round_this), how_many_decimals)

    # luo muokattavan hakemiston sisälle uuden hakemiston kun ok-nappia painettu
    # kopioi muokattavat tiedostot uuteen hakemistoon ja alihakemistoihin ja muuttaatiedostojen versionumeron
    def makedir(self, mypath):

        # Kysytään uusi versionumero jos automaattista numerointia ei ole valittu
        # (oltava float väliltä 0.0 ja 99.99)
        if self.autoversion_var.get() == 0:
            self.version_number_from_user = tk.simpledialog.askfloat('Version Number', 'Please insert new version number (e.g. 1.0)', initialvalue=0.0, minvalue=0.0, maxvalue=99.99)
            
        # luodaan uusi hakemisto muokattavan hakemiston rinnalle
        # try-rakenne jos kansion luonti ei onnistu
        
        try:
            os.mkdir(mypath)
        except:
            showinfo('Warning (mkdir)', 'Could not create directory for edited files')
        # alihakemistojen luonti, jos se on valittuna valuesFrame:ssa 
        if self.subfolder_var.get():
            for rootdir, dirs, files in os.walk(root.directory):

                for name in dirs:
                    if dir:
                        # muodostetaan alihakemiston polku
                        temp_dir = self.createSubfolderPath(mypath, rootdir, name)
                        try:
                            os.mkdir(temp_dir)
                        except:
                            pass
                # kopioidaan tiedostot uusiin hakemistoihin ja muutetaan versionro
                for name in files:
                    if name.endswith(file_extensions):
                        if self.autoversion_var.get(): # automaattinen versionumerointi valittu valuesFramessa
                            # otetaan vanha versionro ja korvataan uudella kutsumalla changeVersionia
                            old_version_number = re.findall('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', name)
                            # tarkistus onko tied.nimessä validi versionumero
                            if old_version_number:
                                new_file_path = re.sub('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', changeVersion(str(old_version_number)), name)
                            else:
                                new_file_path = name # tiedoston nimessä ei validia versionumeroa, käytetään vanhaa tiedostonimeä
                        else: # käyttäjä antaa versionumeron
                            new_file_path = re.sub('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', 'v' + str(self.version_number_from_user), name)

                        self.editFile(os.path.join(rootdir, name), self.createFilePath(rootdir, self.new_folder_name, new_file_path))
        else:
            # alihakemistojen tiedostojen käsittelyä EI valittu
            for name in os.listdir(root.directory):
                if name.endswith(file_extensions):
                    if self.autoversion_var.get(): # tiedostojen nimien automaattinen versionumerointi
                        old_version_number = re.findall('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', name)
                        # tarkistus onko tied.nimessä validi versionumero
                        if old_version_number:
                            new_file_path = re.sub('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', changeVersion(str(old_version_number)), name)
                        else:
                            new_file_path = name # tiedoston nimessä ei validia versionumeroa, käytetään vanhaa tiedostonimeä
                    else: # käyttäjä antaa versionumeron
                        new_file_path = re.sub('v[0-9]+.[0-9]+|[0-9]+v[0-9]+', 'v' + str(self.version_number_from_user), name)
                    self.editFile(os.path.join(root.directory, name), os.path.join(mypath, new_file_path))               
        return True

    # Tietojensyöttöikkuna
    def valuesFrame(self):
        newval = tk.Toplevel(root)
        # Tallennetaan viittaus ikkunaan sulkemista varten
        self.newval = newval
        newval.title('Insert new values')
        newval.geometry('900x800')
        newval.resizable(0, 0)

        # callback-funktion rekisteröinti. Tehdään Tcl wrapper pythonin funktion ympärille.
        # %P on arvo jonka teksti saa, jos muutos oli kelvollinen.
        validate_color_temp = (newval.register(self.int_max_16), "%P")
        validate_cri = (newval.register(self.int_max_6), "%P")
        validate_input_string = (newval.register(self.input_max_78), "%P")
        validate_lamp_type = (newval.register(self.input_max_24), "%P")
        validate_length = (newval.register(self.length_is_ok), "%P")
        validate_light_output_ratio = (newval.register(self.light_output_ratio_ok), "%P")
        validate_lum_flux = (newval.register(self.lum_flux_ok), "%P")
        validate_symmetry = (newval.register(self.symmetry_is_ok), "%P")
        validate_user = (newval.register(self.input_max_65), "%P")
        validate_wattage = (newval.register(self.watt_is_ok), "%P")

        # fontti buttonien tekstiä varten
        helv16 = tkFont.Font(family='Helvetica', size=16, weight='bold')

        # muuttujat joihin tallennetaan ldt:n eri rivien tiedot
        self.query_companyIdentification = StringVar()  # rivi 1
        self.query_symmetryIndicator = StringVar()      # rivi 3
        self.query_luminaireName = StringVar()          # rivi 9
        self.query_measurementUser = StringVar()        # rivi 12
        self.query_luminaireLength = StringVar()        # rivi 13
        self.query_luminaireWidth = StringVar()         # rivi 14
        self.query_luminaireHeight = StringVar()        # rivi 15
        self.query_illuminatedAreaLength = StringVar()  # rivi 16
        self.query_illuminatedAreaWidth = StringVar()   # rivi 17
        self.query_illuminatedAreaHeight = StringVar()  # rivi 18

        self.query_lightOutputRatio = StringVar()       # rivi 23
        self.query_conversionFactorLumInt = StringVar() # rivi 24
        self.query_lampType = StringVar()               # rivi 28
        self.query_luminousFlux = StringVar()           # rivi 29
        self.query_colorTemp = StringVar()              # rivi 30
        self.query_colorRendering = StringVar()         # rivi 31
        self.query_wattage = StringVar()                # rivi 32


        # valitun hakemistopolun tulostus
        label_folder = tk.Label(newval, text = ('Selected folder: ' + root.directory))
        label_folder.pack(side=TOP)
        button_reset = tk.Button(newval, text='Reset all', command = self.reset_all, bg='#FF276F')
        button_reset['font'] = helv16
        button_reset.pack(side=BOTTOM, fill=X)
        button_makechanges = tk.Button(newval, text='Make changes', command = self.confirmFrame, bg='light green')
        button_makechanges['font'] = helv16
        button_makechanges.pack(side=BOTTOM, fill=X)
        
        # checkbuttonit alikansioiden valintaan ja automaattiseen versionumerointiin, default: checked
        self.subfolder_var = IntVar(value=1)
        self.autoversion_var = IntVar(value=1)
        self.checkbtn_subfolder = Checkbutton(newval, text = 'Edit also subfolders', variable = self.subfolder_var, \
                                    onvalue = 1, offvalue = 0)
        self.checkbtn_subfolder.select()
        self.checkbtn_autoversion = Checkbutton(newval, text = 'Use automatic version numbering', variable = self.autoversion_var, \
                                        onvalue = 1, offvalue = 0)         
        self.checkbtn_autoversion.pack(side=BOTTOM)
        self.checkbtn_subfolder.pack(side=BOTTOM)
                            
        # entry labelit uusien arvojen syöttöön
        # self mukana että saadaan luettua tekstikenttä cgetillä myöhemmin
        self.label_companyIdentification = tk.Label(newval, text='Company identification')
        self.label_luminaireName = tk.Label(newval, text='Luminaire name')
        self.label_measurementUser = tk.Label(newval, text='User')
        self.label_luminaireLength = tk.Label(newval, text='Length/diameter of luminaire [mm]')
        self.label_luminaireWidth = tk.Label(newval, text='Width of luminaire [mm]')
        self.label_luminaireHeight = tk.Label(newval, text='Height of luminaire [mm]')
        self.label_illuminatedAreaLength = tk.Label(newval, text='Length/diameter of luminous area [mm]')
        self.label_illuminatedAreaWidth = tk.Label(newval, text='Width of luminous area [mm]')
        self.label_illuminatedAreaHeight = tk.Label(newval, text='Height of luminous area [mm]')

        self.label_symmetryIndicator = tk.Label(newval, text='Symmetry indicator\n(0=no symmetry, 1=vert.axis, 2=C0-C180,\n3=C90-C270, 4=C0-180 and C90-270)')        
        self.label_lightOutputRatio = tk.Label(newval, text='Light output ratio luminaire (%)')
        self.label_conversionFactorLumInt = tk.Label(newval, text='Conversion factor for luminous intensities')
        self.label_lampType = tk.Label(newval, text='Type of lamps (led type)')
        self.label_luminousFlux = tk.Label(newval, text='Total luminous flux [lm]')
        self.label_colorTemp = tk.Label(newval, text='Color temperature [K]')
        self.label_colorRendering = tk.Label(newval, text='Color rendering index (CRI)')
        self.label_wattage = tk.Label(newval, text='Wattage, incl ballast [W]')
        

        entry_companyIdenticifation = tk.Entry(newval, textvariable = self.query_companyIdentification, validate = "key", validatecommand = validate_input_string)
        entry_luminaireName = tk.Entry(newval, textvariable = self.query_luminaireName, validate = "key", validatecommand = validate_input_string)
        entry_measurementUser = tk.Entry(newval, textvariable = self.query_measurementUser, validate = "key", validatecommand = validate_user)
        entry_luminaireLength = tk.Entry(newval, textvariable = self.query_luminaireLength, validate = "key", validatecommand = validate_length)
        entry_luminaireWidth = tk.Entry(newval, textvariable = self.query_luminaireWidth, validate = "key", validatecommand = validate_length)
        entry_luminaireHeight = tk.Entry(newval, textvariable = self.query_luminaireHeight, validate = "key", validatecommand = validate_length)
        entry_illuminatedAreaLength = tk.Entry(newval, textvariable = self.query_illuminatedAreaLength, validate = "key", validatecommand = validate_length)
        entry_illuminatedAreaWidth = tk.Entry(newval, textvariable = self.query_illuminatedAreaWidth, validate = "key", validatecommand = validate_length)
        entry_illuminatedAreaHeight = tk.Entry(newval, textvariable = self.query_illuminatedAreaHeight, validate = "key", validatecommand = validate_length)

        entry_symmetryIndicator = tk.Entry(newval, textvariable = self.query_symmetryIndicator, validate = "key", validatecommand = validate_symmetry)
        entry_lightOutputRatio = tk.Entry(newval, textvariable = self.query_lightOutputRatio, validate = "key", validatecommand = validate_light_output_ratio)
        entry_conversionFactorLumInt = tk.Entry(newval, textvariable = self.query_conversionFactorLumInt, validate = "key", validatecommand = validate_light_output_ratio)
        entry_lampType = tk.Entry(newval, textvariable = self.query_lampType, validate = "key", validatecommand = validate_lamp_type)
        entry_luminousFlux = tk.Entry(newval, textvariable = self.query_luminousFlux, validate = "key", validatecommand = validate_lum_flux)
        entry_colorTemp = tk.Entry(newval, textvariable = self.query_colorTemp, validate = "key", validatecommand = validate_color_temp)
        entry_colorRendering = tk.Entry(newval, textvariable = self.query_colorRendering, validate = "key", validatecommand = validate_cri)
        entry_wattage = tk.Entry(newval, textvariable = self.query_wattage, validate = "key", validatecommand = validate_wattage)
        

        self.label_companyIdentification.place(x=5, y=25, width=400, height=25)
        self.label_luminaireName.place(x=5, y=50, width=400, height=25)
        self.label_measurementUser.place(x=5, y=75, width=400, height=25)

        self.label_luminaireLength.place(x=5, y=125, width=400, height=25)
        self.label_luminaireWidth.place(x=5, y=150, width=400, height=25)
        self.label_luminaireHeight.place(x=5, y=175, width=400, height=25)
        self.label_illuminatedAreaLength.place(x=5, y=200, width=400, height=25)
        self.label_illuminatedAreaWidth.place(x=5, y=225, width=400, height=25)
        self.label_illuminatedAreaHeight.place(x=5, y=250, width=400, height=25)

        self.label_symmetryIndicator.place(x=5, y=285, width=400, height=75)

        self.label_lightOutputRatio.place(x=5, y=375, width=400, height=25)
        self.label_conversionFactorLumInt.place(x=5, y=400, width=400, height=25)
        self.label_lampType.place(x=5, y=425, width=425, height=25)
        self.label_luminousFlux.place(x=5, y=450, width=400, height=25)
        self.label_colorTemp.place(x=5, y=475, width=400, height=25)
        self.label_colorRendering.place(x=5, y=500, width=400, height=25)
        self.label_wattage.place(x=5, y=525, width=400, height=25)

        entry_companyIdenticifation.place(x=350, y=25, width=400, height=25)
        entry_luminaireName.place(x=350, y=50, width=400, height=25)
        entry_measurementUser.place(x=350, y=75, width=400, height=25)

        entry_luminaireLength.place(x=350, y=125, width=100, height=25)
        entry_luminaireWidth.place(x=350, y=150, width=100, height=25)
        entry_luminaireHeight.place(x=350, y=175, width=100, height=25)
        entry_illuminatedAreaLength.place(x=350, y=200, width=100, height=25)
        entry_illuminatedAreaWidth.place(x=350, y=225, width=100, height=25)
        entry_illuminatedAreaHeight.place(x=350, y=250, width=100, height=25)

        entry_symmetryIndicator.place(x=350, y=310, width=100, height=25)

        entry_lightOutputRatio.place(x=350, y=375, width=100, height=25)
        entry_conversionFactorLumInt.place(x=350, y=400, width=100, height=25)
        entry_lampType.place(x=350, y=425, width=100, height=25)
        entry_luminousFlux.place(x=350, y=450, width=100, height=25)
        entry_colorTemp.place(x=350, y=475, width=100, height=25)
        entry_colorRendering.place(x=350, y=500, width=100, height=25)
        entry_wattage.place(x=350, y=525, width=100, height=25)

        return True

    # tietojensyöttöikkunassa painetaan Reset All -nappia.
    # messageboxissa valitaan halutaanko vaihtaa valittua kansiota.
    # Jos 'yes': avataan kansionvalinta-ikkuna.
    # Jos 'no': tyhjennetään entry-kentät.
    def reset_all(self):
        msg_change_folder = tk.messagebox.askquestion('Change folder?', 'Do you want to change the selected folder?')
        if msg_change_folder == 'yes':
            self.newval.destroy()
            self.button_edit['state'] = 'active'
            self.selectFolder()
        else:
            entry_list = [child for child in self.newval.winfo_children()
                            if isinstance(child, Entry)]
            for entry_object in entry_list:
                entry_object.delete(0,'end')

    # Hakemistopolun muodostus tiedostoille
    def createFilePath(self, root_dir, user_given_folder_name, filename):
        # root_dir = hakemistopolku kansioon tai alikansioon, jossa käsiteltävä tiedosto on
        # user_given_folder_name = kansion nimi jonka käyttäjä syötti kysyttäessä
        # filename = tiedoston nimi jossa versionumero on jo muutettu

        norm_root_dir = os.path.normpath(root_dir)

        splitted_rootdirectory = root.directory.split('/')
        splitted_root_dir = norm_root_dir.split('\\')

        replaced_index = len(splitted_rootdirectory) - 1
        splitted_root_dir[replaced_index] = today + ' - ' + user_given_folder_name

        modified_root_dir = '/'.join(splitted_root_dir)

        returned_path = os.path.join(modified_root_dir, filename)
        
        return returned_path



    # Hakemistopolun muodostus alikansioille makedir-metodia varten
    def createSubfolderPath(self, my_path, root_dir, subfolder):
        # my_path = uusi kansio jossa viimeinen osa on 'today - käyttäjän_antama_nimi'
        # root_dir = makedir-metodin os.walkissa käsiteltävänä olevan kansion polku
        # subfolder = makedir-metodin os.walkissa käsiteltävä kansio

        # viimeisen osan indeksi my_pathista
        splitted_my_path = my_path.split('/')
        replaced_index = len(splitted_my_path) - 1

        root_dir_normpath = os.path.normpath(root_dir)

        splitted_root_dir = root_dir_normpath.split('\\')

        splitted_root_dir[replaced_index] = splitted_my_path[replaced_index]
        modified_sub_path = '/'.join(splitted_root_dir)
        returned_path = os.path.join(modified_sub_path, subfolder)
        return returned_path


    # Hakemistopolkujen muodostus
    def createNewPath(self, selected_folder_path, user_given_folder_name):
        rsplitted_thing = selected_folder_path.rsplit('/', 1)

        rsplitted_thing[1] = today + ' - ' + user_given_folder_name
        modified_folder_path = '/'.join(rsplitted_thing)

        return modified_folder_path


    # Varmistetaan haluaako käyttäjä tehdä muutokset
    def confirmFrame(self):
        # flagilla kontrolloidaan monestiko näytetään varoitus-msgbox ongelmatilanteissa. 
        # Ilman flagia varoitus voi tulla muokattavan hakemiston jokaiselle tiedostolle erikseen.
        root.flag_multiplier_only = True
        root.flag_line_9 = True
        root.flag_line_11 = True
        root.flag_writing_error = True

        # enabloidaan Edit-nappi
        self.button_edit['state'] = 'active'
        # listat kaikista labeleista ja entryistä 
        label_list = [child for child in self.newval.winfo_children()
                        if isinstance(child, Label) and child.cget('text') != ('Selected folder: ' + root.directory)] # kaikki labelit paitsi kansion nimi
        entry_list = [child for child in self.newval.winfo_children()
                        if isinstance(child, Entry)] # kaikki Entryt
        
        dict_of_changes = {}
        for entry_object in entry_list:
            if entry_object.get():
                # listan itemin indeksi
                my_index = entry_list.index(entry_object)
                # dictionary muutoksista label, entry
                dict_of_changes.update({str(label_list[my_index].cget('text')) : entry_object.get()})
        # string muutoksista msgbox-tulostusta varten
        print_changes = ""
        for (key, value) in dict_of_changes.items():
            print_changes += str(key + ": " + value + "\n")

        msgbox = tk.messagebox.askquestion('Confirm changes', ('Make these changes to LDT-files?\n') + print_changes)
        if msgbox == 'yes':
            # kysytään uuden kansion nimi, johon tulee muokatut tiedostot. Jos nimenä tyhjä rivi tai pelkkiä välilyöntejä, kysytään uudelleen.
            flag_folder_not_set = True
            while flag_folder_not_set:
                try:
                    self.new_folder_name = ''
                    while not self.new_folder_name.strip():
                        self.new_folder_name = askstring('Folder name', 'New folder name for edited files (date will be added automatically)', initialvalue='mydirectory')
                    flag_folder_not_set = False
                except: 
                    pass

            # luodaan kansiorakenne, tehdään tiedostojen muokkaus (toteutus eri metodeissa)
            mypath = self.createNewPath(root.directory, self.new_folder_name)
            
            # tarkistetaan onko kansio olemassa ennenkuin kutsutaan makedir funktiota
            if os.path.exists(mypath):
                if os.path.isdir(mypath):
                    showinfo('Folder exists already', 'Foldername exists, please give another name.\nNO CHANGES WILL BE MADE')
            else:
                showinfo('NAME INFO', 'Edited files should be in folder: {}'.format(today + ' - ' + self.new_folder_name))
                self.makedir(mypath)
                # kysytään zipataanko muokatut tiedostot
                msgbox_zip = tk.messagebox.askquestion('Zip selection', 'Create ZIP-archive from edited files?')
                if msgbox_zip == 'yes':
                    zip_this = mypath
                    zip_result = mypath
                    background = AsyncZip(zip_this, zip_result)
                    background.start()
                    background.join()

                tk.messagebox.showinfo('Edit done', 'LDT-files are now edited')
                root.flag_multiplier_only = True
                root.flag_line_9 = True
                root.flag_line_11 = True
                root.flag_writing_error = True
            # tuhotaan tietojensyöttöikkuna
            self.newval.destroy()
        else:
            # Muutosten tekoa ei valittu messageboxissa. Tuhotaan tietojensyöttöikkuna.
            tk.messagebox.showinfo('No edit done', 'Edit cancelled. No changes made to LDT-files')
            self.newval.destroy()
    
    # Muokattavan hakemiston valinta
    def selectFolder(self):
        root.directory = filedialog.askdirectory()
        # cancelilla palataan pääikkunaan, ei välitetä tyhjää stringiä eteenpäin
        # tarkistetaan onko kirjoitusoikeutta valittua kansiota pykälää ylemmälle tasolle
        if root.directory and (os.access(os.path.dirname(root.directory), os.W_OK)):
            # mainFramen Edit-napin disablointi, ettei voida avata toista kansionvalintaa päällekkäin
            self.button_edit['state'] = 'disabled'
            self.command = self.valuesFrame()

    # Aloitusikkuna
    def mainFrame(self, root):
        root.title('LDT-editor')
        root.geometry('500x300')
        root.resizable(0, 0) # ikkunan koko ei käyttäjän muutettavissa

        # fontti buttonien tekstiä varten
        helv16 = tkFont.Font(family='Helvetica', size=16, weight='bold')

        button_info = tk.Button(root, text='Help', command = Frames.help)
        button_info['font'] = helv16

        self.button_edit = tk.Button(root, text='Edit', command = self.selectFolder)
        self.button_edit['font'] = helv16

        button_quit = tk.Button(root, text='Quit', command = Frames.quit)
        button_quit['font'] = helv16

        button_info.pack(fill=BOTH)
        self.button_edit.pack(fill=BOTH)
        button_quit.pack(fill=BOTH)

        # logo aloitusnäkymään
        try:
            # kuvan osoite on exeä ajettaessa muodostettavan temp-hakemiston
            # root-path ja kuvan nimi
            base_path = getattr(sys, '_MEIPASS', '.')+'/'
            img = Image.open(base_path + 'logo_greenled.png')
            img.resize((200,100), Image.ANTIALIAS)
            img = ImageTk.PhotoImage(img)
            panel = Label(root, image = img)    
            panel.image = img
            panel.pack(side=BOTTOM)
        except:
            pass

    # Help-napin metodi
    @staticmethod
    def help():
        tk.messagebox.showinfo('INFO', 'Steps of LDT-modification:\n\n\
1) Click \'Edit\'\n\
2) Select folder you want to modify.\n\
3) Insert new values to those entry fields you want to edit.\n\
    Use decimal point for numbers.\n\
    Inserted value is used as a multiplier if entry is started with *\n\
    (e.g. *0.5)\n\
4) If the files in the possible subfolders should NOT be edited,\n\
    unselect checkbox \'Edit also subfolders\'\n\
5) For implementing the edits -> Click \'Make changes\'\n\
    Clear entry fields or change folder selection -> \'Reset all\'\n\
6) Type the name of the folder which will be created\n\
     for edited LDT-files\n\
    (pop-up window for typing the name)\n\
7) Choose if the folder and edited files will be\n\
    archived as a ZIP.')

    # Ohjelman lopetus Quit-napista
    @staticmethod
    def quit():
        quitBox = tk.messagebox.askquestion('QUIT', 'Do you want to exit program?')
        if quitBox == 'yes':
            root.destroy()

# root is a main window object
root = tk.Tk()
app = Frames()
app.mainFrame(root)
root.mainloop()
