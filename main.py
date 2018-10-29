# -*- coding: utf-8 -*-
"""
Created on Thu Jun 19 14:25:11 2014

@author: Testulf
"""

from classes import *
from Tkinter import *
from sql2csv import *
from dicom2sql import *
from sql2pdf import *
from xls2sql import *
import datetime, webbrowser, tkFileDialog, os, shutil

def ensure_dir(f):
    d = os.path.dirname(f)
    if not os.path.exists(d):
        os.makedirs(d)

if __name__ == "__main__":

    
    # weekly backup... :)
    ###########################
    
    week_number = datetime.date.today().isocalendar()[1]
    year = datetime.date.today().year
    backup_string = "{year}-w{week}".format(year=year, week=week_number)
    backup_folder = "backup\\dump\\"
    backup_fn = 'sqlite_xray_back-up_{}.db'.format(backup_string)
    ensure_dir(backup_folder)
    if not backup_fn in os.listdir(backup_folder):
        try:
            print "No backups this week, trying to save for {}...".format(backup_string)
            shutil.copy2('sqlite_xray.db', '{}{}'.format(backup_folder, backup_fn))
        except IOError:
            print "Cannot copy file..."

    # Run GUI
    ####################
    
    root = Tk()
    mainmenu = MainMenu(root)
    root.mainloop()
    state = mainmenu.getState()
    # possible states:
        # createReport
        # addQA
        # deleteQA
        # runSQL
        # addModalityDICOM
        # addModalityManual
        # editModality
        # cancel
    
    # Check state     
    if state == "createReport":
        print "Lager rapporter."
        root = Tk()
        choosereports = ChooseReports(root, level='report', action='find', reportOptions = True)
        root.mainloop()
        reports_to_create_list = choosereports.getReports()
        reportOptions = choosereports.getReportOptions()
        shortReport = reportOptions['shortReport'] 
        smallTableOnLeeds = reportOptions['smallTableOnLeeds']
                       
        sql2pdf(reports_to_create_list, shortReport, smallTableOnLeeds)
        
    elif state == 'addQA':
        print "Legg til QA"
        root = Tk()
        root.withdraw()
        input_files_tk = tkFileDialog.askopenfilenames(title=u"Velg excel-filer (basert på målingsark.xlsm)", multiple=1)
        input_files = root.splitlist(input_files_tk)
        xls2sql(input_files)
        
    elif state == 'deleteQA':
        print "Sletter rapporter."
        root = Tk()
        choosereports = ChooseReports(root, level='report', action='del')
        root.mainloop()
        reports_to_delete = choosereports.getReports()
        cur = db.cursor()
        cur.execute('pragma foreign_keys=ON')
        
        for report_to_delete in reports_to_delete:
            cur.execute("DELETE FROM qa WHERE qa_id = ?", (report_to_delete,))
            
    elif state == "runSQL":
        print u"Lager SQL-rapport."
        root = Tk()
        root.withdraw()
        input_file = tkFileDialog.askopenfilename(title=u"Velg SQL-fil som skal kjøres", initialdir='sql_macros\\')
        
        input_file = input_file.decode('utf-8')
        cur = db.cursor()
        header, result = exec_sql_file_oneline(cur, input_file)
      
        
        print input_file.split(".")[-2]
        
        fileName = '{}_output.csv'.format(input_file.split(".")[-2])
        
        with open(fileName, 'w') as outputFile:
            outputFile.write("{}\n".format(";".join([str(x) for x in header])))
            for line in result:
                outputFile.write("{}\n".format(";".join([str(x) if x else "" for x in line])))
                
        # open file afterwards
        webbrowser.open(fileName)
            
    elif state == 'addModalityManual':
        print "Legger til maskin."
        root = Tk()
        modality = EditModality(root, action='add', station_name=False)
        root.mainloop()
        
        # all parameters given in window
        parameters = modality.getParameters()
        
        # SQL IDs + parameters for hospital, company & people
        sql_ids = modality.getSQLIDs()        
        
        # find IDs
        if 'FK_mod_hos' in parameters.keys():
            hos_id = int(parameters['FK_mod_hos']) # stored as string in parameters
            hos_name = sql_ids['FK_mod_hos'][hos_id]
            
            
        if 'FK_mod_com' in parameters.keys():
            com_id = int(parameters['FK_mod_com'])
            com_name = sql_ids['FK_mod_com'][com_id]
            
        if 'FK_mod_ppl' in parameters.keys():
            ppl_id = int(parameters['FK_mod_ppl'])
            ppl_name = sql_ids['FK_mod_ppl'][ppl_id]
            
        # SQL magic
        cur = db.cursor()
        cur.execute('pragma foreign_keys=ON')

        # COMPANY INSERT
        #######################        
        cur.execute("SELECT com_id FROM company WHERE com_id = ?", (com_id,))
        if len(cur.fetchall()) < 1:
            # new company, let's add
            cur.execute("INSERT INTO company (com_id, com_name) VALUES (?, ?)", (com_id, com_name))
        
        # Hospital insert
        #####################
        cur.execute("SELECT hos_id FROM hospital WHERE hos_id = ?", (hos_id,))
        if len(cur.fetchall()) < 1:
            # new hospital, let's add
            contact = modality.getContact()
            if not contact:
                cur.execute("INSERT INTO hospital (hos_id, hos_name) VALUES (?, ?)", (hos_id, hos_name))
            else:
                cur.execute("INSERT INTO hospital (hos_id, hos_name, hos_contact) VALUES (?, ?, ?)", (hos_id, hos_name, contact))
            
        # Person insert
        #######################
        cur.execute("SELECT ppl_id FROM people WHERE ppl_id = ?", (ppl_id,))
        if len(cur.fetchall()) < 1:
            # new person, let's add
            job = modality.getJob()
            if not job:
                cur.execute("INSERT INTO people (ppl_id, ppl_name) VALUES (?, ?)", (ppl_id, ppl_name))
            else:
                cur.execute("INSERT INTO people (ppl_id, ppl_name, ppl_job) VALUES (?, ?, ?)", (ppl_id, ppl_name, job))

        
        # Big database insert!
        ##########################
        
        # Some sanitation
        #####################
        
        for k, v in parameters.items():
            if v == "None":
                parameters[k] = None # 'None' -> None
            if "FK" in k:
                parameters[k] = int(v) # SQL ID '35' -> 35
        
        parameters_to_database = {
            'comment' : parameters['comment'] == 'None' and None or parameters['comment']
        
        }
        insertQuery = "INSERT INTO modality ({}) VALUES ({})".format(
                    ", ".join(parameters.keys()),
                    ", ".join(["?"]*len(parameters.keys()))
                    )
        cur.execute(insertQuery, (parameters.values()))
            
        db.commit()
    
    elif state == 'addModalityDICOM':
        print "Legger til maskin med DICOM-informasjon."
        root = Tk()
        root.withdraw()
        input_files_tk = tkFileDialog.askopenfilenames(title="Velg DICOM-filer", multiple=1)
        input_files = root.splitlist(input_files_tk)
        dicom2sql(input_files)
    
    elif state == 'editModality':
        print "Redigerer maskin."
        root = Tk()
        choosereports = ChooseReports(root, level='modality', action='find')
        root.mainloop()
        out = choosereports.getModalities()
        if out:
            modality_to_edit = out[0]
            root = Tk()
            editmodality = EditModality(root, action='edit', station_name = modality_to_edit)
            root.mainloop()
            getState = editmodality.getState()
            if getState:
                cur = db.cursor()
                cur.execute('pragma foreign_keys=OFF')
                getParameters = editmodality.getParameters()
                machineId = editmodality.getId()
                keys = getParameters.keys()
                values = getParameters.values()
                for key, value in getParameters.items():
                    query = """update modality
                                set {} = :newValue
                                where mod_id = :machineId""".format(key)
                    # cannot use :-type placeholder for column name
                    cur.execute(query, {"newValue":value, "machineId":machineId})
                db.commit()
        
    elif state == 'delModality':
        print "Sletter maskin."
        root = Tk()
        choosereports = ChooseReports(root, level='modality', action='del')
        root.mainloop()
        modality_to_del = choosereports.getModalities()
        print modality_to_del
        # Do SQL magic

    elif state == "csv2sql":
        print "Legger til informasjon CSV-filer i databasen."
        csv2sql()
        
    elif state == 'sql2csv':
        print "Lagerer informasjon om database til CSV-filer."
        sql2csv()
        
    elif state == 'createSQL':
        print "Lager database fra grunnen av!"
        createSQL()

    elif not state or state == "cancel":
         print "Ok, ingenting valgt."       
    
    else:
        print "Finner ikke valg: %s" % state