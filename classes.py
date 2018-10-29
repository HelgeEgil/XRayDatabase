# -*- coding: utf-8 -*-
"""
Created on Mon Jun 23 14:30:40 2014

@author: Testulf
"""

import os
from Tkinter import *
from sql_connection import *

class ThisShouldntHappen(Exception):
    def __init__(self, value):
        self.value = value
        
    def __str__(self):
        return repr(self.value)

class Modality:
    def __init__(self):
        self.station_name = None
        self.manufacturer = None
        self.department = None
        self.institution = None
        self.modality = None
        self.model = None
        self.software = None
        self.lab = None
        self.delivery_date = None
        self.contact = None
        self.serial = None
        self.detector = None
        self.has_dap = None
        self.has_aek = None
        self.discard_date = None
        self.comment = None
        self.qa_list = []
        self.responsibility = 1
        self.mobile = None
        
    def addMobile(self, x):
        self.mobile = x
        
    def getMobile(self):
        return self.mobile
    
    def addComment(self, x):
        self.comment = x
        
    def getComment(self):
        return self.comment
    
    def addResponsibility(self, x):
        self.responsibility = x
        
    def getResponsibility(self):
        return self.responsibility
    
    def addQA(self, date, reportnumber = None, reportcomment = None):
        self.qa_list.append([date, reportnumber, reportcomment])
        
    def getQA(self):
        return self.qa_list
        
    def addHasAek(self, x):
        self.has_aek = x

    def getHasAek(self):
        return self.has_aek
        
    def addHasDap(self, x):
        self.has_dap = x

    def getHasDap(self):
        return self.has_dap
        
    def addSerial(self, x):
        self.serial = x
        
    def getSerial(self):
        return self.serial
        
    def addDetector(self, x):
        self.detector = x
        
    def getDetector(self):
        return self.detector
        
    def addContact(self, x):
        self.contact = x
        
    def getContact(self):
        return self.contact
        
    def addDeliveryDate(self, x):
        self.delivery_date = x
        
    def getDeliveryDate(self):
        return self.delivery_date
        
    def addDiscardDate(self, x):
        self.discard_date = x
        
    def getDiscardDate(self):
        return self.discard_date
    
    def addStationName(self, x):
        self.station_name = x
    
    def getStationName(self):
        return self.station_name
    
    def addManufacturer(self, x):
        self.manufacturer = x
        
    def getManufacturer(self):
        return self.manufacturer
    
    def addInstitution(self, x):
        self.institution = x
        
    def getInstitution(self):
        return self.institution
    
    def addDepartment(self, x):
        self.department = x
        
    def getDepartment(self):
        return self.department
        
    def addModality(self, x):
        self.modality = x
        
    def getModality(self):
        return self.modality
        
    def addModel(self, x):
        self.model = x
        
    def getModel(self):
        return self.model
        
    def addSoftware(self, x):
        self.software = x
        
    def getSoftware(self):
        return self.software
        
    def addLab(self, x):
        self.lab = x
        
    def getLab(self):
        return self.lab

class ChooseReports(Frame):

    def __init__(self, parent, level, action, reportOptions = False, maxresults = 9999):
            
        Frame.__init__(self, parent)

        cur = db.cursor()
        cur.execute('pragma foreign_keys=ON')
        
        self.parent = parent        
        self.parent.protocol("WM_DELETE_WINDOW", self.myQuit)
        self.parent.title("Choose reports")
        
        self.maxresults = maxresults
        self.reportOptions = reportOptions
        
        if level == "report":        
            self.reportlevel = True
        else:
            self.reportlevel = False
        
        if self.reportlevel:
            cur.execute("SELECT h.hos_name, m.department, m.station_name, m.model, m.serial, c.com_name, \
            date(qa.study_date, '%Y'), qa.doc_number, m.modality_name , date(m.deliverydate, '%Y'), m.lab \
            FROM modality m \
            INNER JOIN hospital h ON h.hos_id = m.FK_mod_hos \
            INNER JOIN company c ON c.com_id = m.FK_mod_com \
            INNER JOIN qa ON qa.FK_qa_mod = m.mod_id \
            WHERE qa_data = 1")
        else:
            # 0 hospital
            # 1 department
            # 2 stationname
            # 3 model
            # 4 serial
            # 5 company
            # 6 modality name
            # 7 delivery year
            # 8 lab
            # 9 detektor
            cur.execute("SELECT h.hos_name, m.department, m.station_name, m.model, m.serial, c.com_name, \
                        m.modality_name , date(m.deliverydate, '%Y'), m.lab, m.detector FROM modality m \
                        INNER JOIN hospital h ON h.hos_id = m.FK_mod_hos \
                        INNER JOIN company c ON c.com_id = m.FK_mod_com")
#           Was:   m.modality_name , YEAR(m.deliverydate), m.lab, m.detector FROM modality m \
#           With sqlite YEAR(datestring) becomes date(datestring, '%Y')
        
        results = cur.fetchall()
        
        # Create dictionary list with {hospital : [{department: [{machine: [{year: }]}]}]}
        
        self.dict = {}
        self.modelname = {}
        self.create_report_list = []
        self.modality_list = []        
        
        self.possible_years = []        
        
        if self.reportlevel:
            for result in results:
                year = result[6]
                if not year in self.possible_years:
                    self.possible_years.append(year)
        
        
        if reportOptions:
            self.shortReportVar = IntVar()
            self.smallLeedsTableVar = IntVar()
            self.shortReportVar.set(0)
            self.smallLeedsTableVar.set(0)
        
        button_width = 25        
        
        #Find unique hospitals
        for result in results:
            hospital = result[0]
            if not hospital in self.dict.keys():
                self.dict[hospital] = {}
        
        for result in results:
            hospital = result[0]
            department = result[1]
            if not department in self.dict[hospital].keys():
                self.dict[hospital][department] = {}
        
        if not self.reportlevel:
            len_hospital = 0
            len_department = 0
            len_stationname = 0
            len_companymodel = 0
            len_modalityname = 0
            len_lab = 0
            len_detector = 0
            
            for result in results:
                hospital = result[0]
                department = result[1]
                stationname = result[2]
                model = result[3]
                company = result[5]
                modalityname = result[6]
                modalityyear = result[7]
                lab = result[8]
                detector = result[9]
                
                if not detector:
                    detector = ""
                    
                if not lab:
                    lab = ""

                
                len_hospital = max(len_hospital, len(hospital))
                len_department = max(len_department, len(department))
                len_stationname = max(len_stationname, len(stationname))
                len_companymodel = max(len_companymodel, len("%s %s" % (company, model)))
                len_modalityname = max(len_modalityname, len(modalityname))
                len_lab = max(len_lab, len(lab))
        
        for result in results:
            if self.reportlevel:
                hospital = result[0]
                department = result[1]
                stationname = result[2]
                model = result[3]
                serial = result[4]
                company = result[5]
                modalityname = result[8]
                modalityyear = result[9]
                lab = result[10]
            else:
                hospital = result[0]
                department = result[1]
                stationname = result[2]
                model = result[3]
                company = result[5]
                modalityname = result[6]
                modalityyear = result[7]
                lab = result[8]
                detector = result[9]
                
            if not lab:
                lab = ""
            
            companymodel = "%s %s" % (company, model)            
            
            if not stationname in self.dict[hospital][department].keys():
                self.dict[hospital][department][stationname] = []
                #self.modelname[stationname] = "%s | %s | %s | %s | %s" % (lab.ljust(len_lab), 
                #modalityyear, modalityname.ljust(len_modalityname), stationname.ljust(len_stationname), 
                #companymodel.ljust(len_companymodel))
                self.modelname[stationname] = " | ".join(filter(None, [lab, modalityyear, modalityname,
                        stationname, companymodel]))
            
        if self.reportlevel:
            for result in results:
                hospital = result[0]
                department = result[1]
                stationname = result[2]
                year = result[6]
                reportnumber = result[7]
            
                if not reportnumber in self.dict[hospital][department][stationname]:
                    self.dict[hospital][department][stationname].append(reportnumber)                     
       
        
        self.titleframe = Frame(self, borderwidth=5, relief=RIDGE, height=40)
        self.boxframe = Frame(self, borderwidth=25)
        
        self.variable_hospital = StringVar(self)
        self.variable_department = StringVar(self)
        self.variable_modality = StringVar(self)
        if self.reportlevel:
            self.variable_report = StringVar(self)
            self.variable_year = StringVar(self)
        
        self.variable_hospital.trace('w', self.updateoptions_hospital)
        self.variable_department.trace('w', self.updateoptions_department)
        self.variable_modality.trace('w', self.updateoptions_modality)
        if self.reportlevel:
            self.variable_report.trace('w', self.updateoptions_report)
            self.variable_year.trace('w', self.updateoptions_year)
        
        self.action_text = "Velg"        
        
        if action == 'del' and level == 'report':
            title_text = u"Velg rapport(er) som skal slettes."
        elif action == 'del' and level == 'modality':
            self.reportlevel = False
            title_text = u'Velg maskin(er) som skal slettes.'
        elif action == 'find' and level == 'modality':
            self.reportlevel = False
            self.maxresults = True
            title_text = u'Velg maskin som skal redigeres (maksimum 1).'
        elif action == 'find' and level == 'report':
            title_text = u'Velg rapport(er) som skal lages.'
            self.action_text = "Lag"
        
        self.label_title = Label(self.titleframe, text=title_text)        
        
        if self.reportOptions:
            self.reportOptionsFrame = Frame(self, borderwidth=5)
            self.shortReportCheckbutton = Checkbutton(self.reportOptionsFrame, text="Kort rapport", variable=self.shortReportVar)
            self.smallLeedsTableCheckbutton = Checkbutton(self.reportOptionsFrame, text="Krymp IQ-tabell", variable=self.smallLeedsTableVar)

            
        self.label_hospital = Label(self.boxframe, text="Velg sykehus")
        self.label_department = Label(self.boxframe, text="Velg avdeling")
        self.label_modality = Label(self.boxframe, text="Velg maskin")
        
        if self.reportlevel:
            self.label_report = Label(self.boxframe, text="%s rapport" % self.action_text)
        
            self.label_year = Label(self.boxframe, text=u"Begrens år")

        self.button_report = Button(self.boxframe, command=self.createreport)

        if level == 'modality':
            self.item_text = "maskin"
        else:
            self.item_text = "rapport"
            
        self.button_report.configure(text="%s %s" % (self.action_text, self.item_text), width=button_width)       
        
        self.optionmenu_hospital = OptionMenu(self.boxframe, self.variable_hospital, *self.dict.keys())
        self.optionmenu_department = OptionMenu(self.boxframe, self.variable_department, '')
        self.optionmenu_modality = OptionMenu(self.boxframe, self.variable_modality, '')
        
        if self.reportlevel:        
            self.optionmenu_report = OptionMenu(self.boxframe, self.variable_report, '') 
        
            self.optionmenu_year = OptionMenu(self.boxframe, self.variable_year, *self.possible_years)        
        
        self.variable_hospital.set(self.dict.keys()[0])
        
        self.titleframe.pack()

        if self.reportOptions:
            self.reportOptionsFrame.pack()
            self.shortReportCheckbutton.pack(side=LEFT)
            self.smallLeedsTableCheckbutton.pack(side=LEFT)
        
        self.boxframe.pack()
        self.label_title.pack()
        
        self.label_hospital.pack()
        self.optionmenu_hospital.pack()
        self.label_department.pack()
        self.optionmenu_department.pack()
        self.label_modality.pack()
        self.optionmenu_modality.pack()
        if self.reportlevel:
            self.label_report.pack()
            self.optionmenu_report.pack()
            self.label_year.pack()
            self.optionmenu_year.pack()
        
        self.button_report.pack()
        
        self.pack()
        
    def createreport(self):
        self.parent.destroy()
        
    def myQuit(self):
        self.create_report_list = []
        self.modality_list = []
        self.parent.destroy()        
        
    def updateoptions_hospital(self, *args):
        departments = self.dict[self.variable_hospital.get()]
#        self.variable_department.set("Alle")
        menu = self.optionmenu_department['menu']
        menu.delete(0, 'end')
        self.create_report_list = []
        self.modality_list = []
        for department in departments.keys():
            menu.add_command(label=department, command=lambda department=department: self.variable_department.set(department))

            if self.reportlevel:    
                for kmod in departments[department]:
                    for rep in self.dict[self.variable_hospital.get()][department][kmod]:
                        self.create_report_list.append(rep)
            else:
                for modality in departments[department].keys():
                    self.modality_list.append(modality)
        if self.reportlevel:                    
            n = len(self.create_report_list)
        else:
            n = len(self.modality_list)

        self.button_report.configure(text="%s %d %s%s" % (self.action_text, n, self.item_text, n>1 and "er" or ""))
        if n> self.maxresults:
            self.button_report.configure(state=DISABLED)            
        else:
            self.button_report.configure(state=ACTIVE)    
            
    def updateoptions_department(self, *args):
        modalities = self.dict[self.variable_hospital.get()][self.variable_department.get()]
#        self.variable_modality.set("Alle")
        menu = self.optionmenu_modality['menu']
        menu.delete(0, 'end')
        self.create_report_list = []
        self.modality_list = []
        for modality in modalities.keys():
            menu.add_command(label=self.modelname[modality], command=lambda modality=modality: self.variable_modality.set(modality))
            
            if self.reportlevel:
                for rep in modalities[modality]:
                    self.create_report_list.append(rep)
            else:
                self.modality_list.append(modality)
                    
        if self.reportlevel:
            n = len(self.create_report_list)
        else:
            n = len(self.modality_list)
        self.button_report.configure(text="%s %d %s%s" % (self.action_text, n, self.item_text, n>1 and "er" or ""))
        if n> self.maxresults:
            self.button_report.configure(state=DISABLED)            
        else:
            self.button_report.configure(state=ACTIVE)
        
    def updateoptions_modality(self, *args):
        reports = self.dict[self.variable_hospital.get()][self.variable_department.get()][self.variable_modality.get()]
#        self.variable_report.set("Alle")
        if self.reportlevel:
            menu = self.optionmenu_report['menu']
            menu.delete(0, 'end')
            self.create_report_list = []
            
            for report in reports:
                menu.add_command(label=report, command=lambda report=report: self.variable_report.set(report))
                self.create_report_list.append(report)
            n = len(self.create_report_list)
        else:
            modality = self.variable_modality.get()
            self.modality_list = [modality]
            n = len(self.modality_list)
        
        self.button_report.configure(text="%s %d %s%s" % (self.action_text, n, self.item_text, n>1 and "er" or ""))
        if n> self.maxresults:
            self.button_report.configure(state=DISABLED)            
        else:
            self.button_report.configure(state=ACTIVE)
            
    def updateoptions_report(self, *args):
        report = self.variable_report.get()
        self.create_report_list = [report]
        self.button_report.configure(text="%s 1 rapport" % (self.action_text))
        self.button_report.configure(state=ACTIVE)
    def updateoptions_year(self, *args):
        self.create_report_list = [x for x in self.create_report_list if x[:2] == str(self.variable_year.get())[-2:]]
        if self.reportlevel:
            n = len(self.create_report_list)
        else:
            n = len(self.modality_list)
        self.button_report.configure(text="%s %d %s%s" % (self.action_text, n, self.item_text, n>1 and "er" or ""))
        if n> self.maxresults:
            self.button_report.configure(state=DISABLED)            
        else:
            self.button_report.configure(state=ACTIVE)
    def getReports(self):
        return self.create_report_list
        
    def getReportOptions(self):
        return {'smallTableOnLeeds' : self.smallLeedsTableVar.get(),
                'shortReport' : self.shortReportVar.get()}

    def getModalities(self):
        return self.modality_list

class Question(Frame):
    def __init__(self, parent, text):
        
        Frame.__init__(self, parent)
        
        self.text = text
        self.answer = False
        self.parent = parent
        
        self.upperContainer = Frame(self)
        self.lowerContainer = Frame(self)
        
        self.label_question = Label(self.upperContainer, text=self.text)
        self.button_yes = Button(self.lowerContainer, command=self.yes)
        self.button_no = Button(self.lowerContainer, command=self.no)
        
        self.button_yes.configure(text = "Ja")
        self.button_no.configure(text = "No")
        
        self.upperContainer.pack(side=TOP)
        self.lowerContainer.pack(side=TOP)
        
        self.label_question.pack(side=TOP)
        
        self.button_yes.pack(side=LEFT)
        self.button_no.pack(side=LEFT)
        
        self.pack()
        
    def yes(self):
        self.answer = 1
        self.parent.destroy()
    
    def no(self):
        self.answer = 0
        self.parent.destroy()
        
    def getAnswer(self):
        return self.answer

class DeleteReport(Frame):
    def __init__(self, parent, reportnum):
            
        Frame.__init__(self, parent)
        
        self.upperContainer = Frame(self)
        self.lowerContainer = Frame(self)
        
        cur = db.cursor()
        cur.execute('pragma foreign_keys=ON')
        
        self.parent = parent        
        self.reportnum = reportnum
        
        cur.execute("SELECT h.hos_name, m.department, m.station_name, m.model, m.serial, c.com_name, \
        date(qa.study_date, '%Y'), qa.doc_number, m.modality_name, qa.study_date FROM modality m \
        LEFT OUTER JOIN qa ON qa.FK_qa_mod = m.mod_id \
        LEFT OUTER JOIN company c ON c.com_id = m.FK_mod_com \
        LEFT OUTER JOIN hospital h ON h.hos_id = m.FK_mod_hos \
        WHERE qa.doc_number = %s", self.reportnum)
        
        results = cur.fetchall()

        n_rep = len(results)
        
        print "Fant %d rapport%s" % (n_rep, n_rep>1 and "er. Noe er feil!" or ".")
        assert n_rep == 1
        
        fr = results[0]
        hospital = fr[0]
        department = fr[1]
        station_name = fr[2]
        model = fr[3]
        company = fr[5]
        modality_type = fr[8]
        date = fr[9]
        
        self.label_report = Label(self.upperContainer, text=u"Er du sikker på at du vil slette rapport %s?" % self.reportnum)
        self.label_reporttext = Label(self.upperContainer, text=u"%s | %s | %s | %s | %s | %s" % (hospital, department, date, 
                                                                         station_name, company, model))

        self.button_yes = Button(self.lowerContainer, command=self.deletereport)        
        self.button_no  = Button(self.lowerContainer, command=self.donothing)
        
        self.button_yes.configure(text = "Ja")
        self.button_no.configure(text = "Nei")
        
        self.upperContainer.pack(side=TOP)
        self.lowerContainer.pack(side=TOP)        
        
        self.label_report.pack(side=TOP)
        self.label_reporttext.pack(side=TOP)
        
        self.button_yes.pack(side=LEFT)
        self.button_no.pack(side=LEFT)
        
        self.pack()
        
    def deletereport(self):
        print "Sletter rapport %s." % self.reportnum
        cur.execute("DELETE FROM qa WHERE qa_id = ?", (self.reportnum,))
        self.parent.destroy()
        
    def donothing(self):
        self.parent.destroy()        

class MainMenu(Frame):
    def __init__(self, parent):
            
        Frame.__init__(self, parent)
        cur = db.cursor()
        cur.execute('pragma foreign_keys=ON')
#        try:
#            cur = db.cursor()
#        except:
#            db = db_nodb
#            cur = db.cursor()
        try:
            cur.execute("SELECT COUNT(*) FROM MODALITY")
            result = cur.fetchall()[0][0]
#        except sqlite3.OperationalError:
        except:
            result = 0
            
        db.commit()
        if result == 0:
            self.noMachines = True
        else:
            self.noMachines = False
                    
        self.parent = parent
        self.parent.protocol("WM_DELETE_WINDOW", self.cancel)
        self.parent.title("Grensesnitt mot SQL")
        
        self.leftContainer = Frame(self)
        self.upperLeftContainer = Frame(self.leftContainer, borderwidth=5, relief=RIDGE, height=40)

        self.lowerLeftContainer = Frame(self.leftContainer, borderwidth=25)
        
        self.rightContainer = Frame(self)
        self.upperRightContainer = Frame(self.rightContainer, borderwidth=5, relief=RIDGE, height=40)
        self.lowerRightContainer = Frame(self.rightContainer, borderwidth=25)        
        
        self.label_title = Label(self.upperLeftContainer, text="Hovedgrensesnitt mot rapportdatabase")
        self.state = False
        self.addModalityPressed = False
        self.editSQLPressed = False
        
        button_width = 25
        
        self.button_createreport = Button(self.lowerLeftContainer, text="Lag QA-rapport", command=self.createReport, width=button_width)
        self.button_addQA = Button(self.lowerLeftContainer, text=u"Legg til måling", command=self.addQA, width=button_width)
        self.button_deleteQA = Button(self.lowerLeftContainer, text=u"Slett måling", command=self.deleteQA, width=button_width)
#        
#        self.button_runSQLfile = Button(self.lowerLeftContainer, text="Lag CSV-rapport -->", command=self.runSQLfile, width=button_width)
#        self.listOfSQLFiles = ["{}\\{}".format(sql_macro_path, f) for f in os.listdir(sql_macro_path) if '.sql' in f]
#        self.SQLFileButtons = []
#        for SQLFile in self.listOfSQLFiles:
#            self.SQLFileButtons.append(Button(self.lowerLeftContainer, text=SQLFile.split("\\")[-1].replace(".sql",""), command))
        
        self.button_runSQLFile = Button(self.lowerLeftContainer, text=u'SQL-fil til CSV-rapport', command=self.runSQLFile, width=button_width)
        
        self.button_addModality = Button(self.lowerLeftContainer, text=u"Legg til maskin -->", command=self.addModalityChoice, width=button_width)
        self.button_addModalityDICOM = Button(self.lowerRightContainer, text=u"Fra DICOM-filer", command=self.addModalityDICOM, width=button_width)
        self.button_addModalityManual = Button(self.lowerRightContainer, text=u"Manuelt", command=self.addModalityManual, width=button_width)
        self.button_editModality = Button(self.lowerLeftContainer, text=u"Gjør endringer på maskin", command=self.editModality, width=button_width)
        
        self.button_editSQL = Button(self.lowerLeftContainer, text=u"Databasebehandling -->", command=self.editSQL, width=button_width)
        self.button_createSQL = Button(self.lowerRightContainer, text=u"Lag databasestruktur", command=self.createSQL, width=button_width)
        self.button_sql2csv = Button(self.lowerRightContainer, text=u"Lagre database som CSV-filer", command=self.sql2csv, width=button_width)
        self.button_csv2sql = Button(self.lowerRightContainer, text=u"Lag database fra CSV-filer", command=self.csv2sql, width=button_width)
        
        self.button_cancel = Button(self.lowerLeftContainer, text=u"Avbryt", command=self.cancel, width=button_width)
     
        self.addModalityLabel = Label(self.upperRightContainer, text="Legg til maskin")
        self.editSQLLabel = Label(self.upperRightContainer, text="Databasebehandling")
     
        if self.noMachines:
            self.button_createreport.configure(state=DISABLED)
            self.button_addQA.configure(state=DISABLED)
            self.button_deleteQA.configure(state=DISABLED)
            self.button_editModality.configure(state=DISABLED)
            self.button_runSQLFile.configure(state=DISABLED)
#            self.button_addModality.configure(state=DISABLED)
     
        self.leftContainer.pack(side=LEFT)                               
        self.upperLeftContainer.pack(side=TOP)
        self.lowerLeftContainer.pack(side=TOP)

        self.label_title.pack()
        self.button_createreport.pack()
        self.button_addQA.pack()
        self.button_deleteQA.pack()
        self.button_runSQLFile.pack()
        self.button_addModality.pack()
        self.button_editModality.pack()
        self.button_editSQL.pack()
        self.button_cancel.pack()
        
        self.pack()
        
    def createReport(self):
        self.state = 'createReport'
        # MYSQL-MAGI
        self.parent.destroy()
        
    def addQA(self):
        self.state = 'addQA'
        self.parent.destroy()
        
    def deleteQA(self):
        self.state = 'deleteQA'
        self.parent.destroy()
#
    def runSQLFile(self):
        self.state = 'runSQL'
        self.parent.destroy()
        
    def addModalityChoice(self):
        
        if not self.addModalityPressed:
            self.addModalityLabel.pack()        
            self.rightContainer.pack(side=LEFT)
            self.upperRightContainer.pack()
            self.lowerRightContainer.pack()
            self.button_addModalityDICOM.pack()
            self.button_addModalityManual.pack()
            self.button_addModality.configure(relief=SUNKEN)
            self.addModalityPressed = True
        else:
            self.addModalityLabel.pack_forget()        
            self.rightContainer.pack_forget()
            self.upperRightContainer.pack_forget()
            self.lowerRightContainer.pack_forget()
            self.button_addModalityDICOM.pack_forget()
            self.button_addModalityManual.pack_forget()
            self.button_addModality.configure(relief=RAISED)
            self.addModalityPressed = False
        
    def editSQL(self):
        if not self.editSQLPressed:
            self.editSQLLabel.pack()
            self.rightContainer.pack(side=LEFT)
            self.upperRightContainer.pack()
            self.lowerRightContainer.pack()
            self.button_sql2csv.pack()
            self.button_csv2sql.pack()
            self.button_createSQL.pack()
            self.button_editSQL.configure(relief=SUNKEN)
            if self.noMachines:
                self.button_sql2csv.configure(state=DISABLED)
#                self.button_csv2sql.configure(state=DISABLED)
            self.editSQLPressed = True
        else:
            self.editSQLLabel.pack_forget()
            self.rightContainer.pack_forget()
            self.upperRightContainer.pack_forget()
            self.lowerRightContainer.pack_forget()
            self.button_sql2csv.pack_forget()
            self.button_csv2sql.pack_forget()
            self.button_createSQL.pack_forget()
            self.button_editSQL.configure(relief=RAISED)
            self.editSQLPressed = False
        
    def createSQL(self):
        
        # Are You Sure-dialog
        
        self.top = Toplevel()
        self.top.wm_attributes("-topmost", 1)
        self.top.focus_force()
        
#        self.top.bind('<Return>', self.add_ok)
        
        self.top.title(u"Spørsmål")
        
        self.top_topContainer = Frame(self.top)
        self.top_bottomContainer = Frame(self.top)
        
        self.top_topContainer.pack()
        self.top_bottomContainer.pack()
        
        self.add_label = Label(self.top_topContainer, text=u"Er du sikker på at du vil fortsette?\nAll informasjon i databasen vil bli overskrevet.")
        self.add_label.pack(side=LEFT)
        
        self.ok_button = Button(self.top_bottomContainer, text="JA", command=self.sql_yes)
        self.cancel_button = Button(self.top_bottomContainer, text="NEI", command=self.sql_no)
        
        self.ok_button.pack(side=LEFT)
        self.cancel_button.pack(side=LEFT)
            
            
    def sql_yes(self, *args):
        self.state = 'createSQL'
        self.top.destroy()
        self.parent.destroy()
    
    def sql_no(self, *args):
        self.top.destroy()
        
    def sql2csv(self):
        self.state = 'sql2csv'
        self.parent.destroy()
        
    def csv2sql(self):
        self.state = "csv2sql"
        self.parent.destroy()
        
    def addModalityManual(self):
        self.state = 'addModalityManual'
        self.parent.destroy()

    def addModalityDICOM(self):
        self.state = 'addModalityDICOM'
        self.parent.destroy()        
        
    def editModality(self):
        self.state = 'editModality'
        self.parent.destroy()
        
    def cancel(self):
        self.state = 'cancel'
        self.parent.destroy()
        
    def getState(self):
        return self.state

class CorrectListOfTags(Frame):
    def __init__(self, parent, dict_of_tags, type_of_tag):
        
        # dict of form { value : station name of example machine }
        
        Frame.__init__(self, parent)
        
        self.parent = parent
        
        self.parent.protocol("WM_DELETE_WINDOW", self.cancel)
        
        self.titleContainer = Frame(self, borderwidth=5, relief=RIDGE, height=40)
        self.topContainer = Frame(self, borderwidth=25)
        self.lowerContainer = Frame(self, borderwidth=25)
    
        self.labelContainer = Frame(self.topContainer, borderwidth=5)
        self.oldContainer = Frame(self.topContainer, borderwidth=5)
        self.newContainer = Frame(self.topContainer, borderwidth=5)
        
        self.label_title = Label(self.titleContainer, text=u"Se over listen over eksisterende (unike) {}-verdier og gjør dem så like så mulig".format(type_of_tag))
        
        self.machines = dict_of_tags.values()
        self.tags = dict_of_tags.keys()
        
        n = len(self.tags)
        self.updatedTags= [0]*n
        
        self.labelname = [0]*n
        self.oldtext = [0]*n
        self.newtext = [0]*n
        self.label_entries = [0]*n
        self.old_entries = [0]*n
        self.new_entries= [0]*n     
        
        self.title_label_container = Frame(self.labelContainer, borderwidth=5)
        self.title_old_container = Frame(self.oldContainer, borderwidth=5)
        self.title_new_container = Frame(self.newContainer, borderwidth=5)

        self.body_label_container = Frame(self.labelContainer, borderwidth=5)
        self.body_old_container = Frame(self.oldContainer, borderwidth=5)
        self.body_new_container = Frame(self.newContainer, borderwidth=5)
        
        self.title_label = Label(self.title_label_container, text=u"Stasjonsnavn")
        self.old_label = Label(self.title_old_container, text=u"Fra DICOM")
        self.new_label = Label(self.title_new_container, text=u"Ny verdi")
        
        for enum, (machine, tag) in enumerate( zip(self.machines, self.tags) ):
            self.labelname[enum] = StringVar(self, value=machine)
            self.oldtext[enum] = StringVar(self, value=tag)
            self.newtext[enum] = StringVar(self, value=tag)
            self.label_entries[enum] = Entry(self.body_label_container, textvariable=self.labelname[enum], state=DISABLED, selectborderwidth=2)
            self.old_entries[enum] = Entry(self.body_old_container, textvariable=self.oldtext[enum], state=DISABLED, selectborderwidth=2, width=50)
            self.new_entries[enum] = Entry(self.body_new_container, textvariable=self.newtext[enum], selectborderwidth=2, width=50)
        
        buttonwidth = 40
        self.ok_button = Button(self.lowerContainer, text="OK", command=self.ok, width=buttonwidth)
        self.cancel_button = Button(self.lowerContainer, text="Cancel", command=self.cancel, width=buttonwidth)
        
        self.titleContainer.pack(side=TOP)
        self.label_title.pack()        
        
        self.topContainer.pack(side=TOP)
        
        self.labelContainer.pack(side=LEFT)
        self.title_label_container.pack()
        self.title_label.pack()
        self.body_label_container.pack()
        
        for p in self.label_entries:
            p.pack()
        
        self.oldContainer.pack(side=LEFT)
        self.title_old_container.pack()    
        self.old_label.pack()
        self.body_old_container.pack()
        
        for p in self.old_entries:        
            p.pack()

        self.newContainer.pack(side=LEFT)
        self.title_new_container.pack()
        self.new_label.pack()
        self.body_new_container.pack()

        for p in self.new_entries:        
            p.pack()

        self.lowerContainer.pack()
        self.ok_button.pack(side=LEFT)
        self.cancel_button.pack(side=LEFT)
        
        self.pack()


    def ok(self):
        # Set all newtext{} values to the inputted strings
        for enum, valu in enumerate(self.newtext):
            if self.newtext[enum].get() in ["None", ""]:
                self.newtext[enum].set(None)
                
            self.updatedTags[enum] = self.newtext[enum].get()
        self.edited = True
        self.parent.destroy()
        self.quit()
        
    def cancel(self):
        self.edited = False
        self.parent.destroy()
        self.quit()
        
    def getState(self):
        return self.edited
        
    def getParameters(self):
        return self.updatedTags

class EditModality(Frame):
    def __init__(self, parent, action, station_name):
            
        Frame.__init__(self, parent)
        self.parent = parent
        self.station_name = station_name
        self.action = action
        self.edited = False
        
        self.parent.protocol("WM_DELETE_WINDOW", self.myQuit)
        self.parent.wm_attributes("-topmost", 1)
        
        cur = db.cursor()
        cur.execute('pragma foreign_keys=ON')
        
        self.titleContainer = Frame(self, borderwidth=5, relief=RIDGE, height=40)
        self.topContainer = Frame(self, borderwidth=25)
        self.lowerContainer = Frame(self, borderwidth=25)
    
        self.labelContainer = Frame(self.topContainer, borderwidth=5)
        self.oldContainer = Frame(self.topContainer, borderwidth=5)
        self.newContainer = Frame(self.topContainer, borderwidth=5)
    
        if self.action == 'edit':
            label_text = "Rediger modalitet"
            self.parent.title(u"Gjør endringer på modalitet")
        else:
            label_text = "Legg til modalitet"
            self.parent.title(u"Legg til modalitet")
            
        self.label_title = Label(self.titleContainer, text=label_text)
        
        button_width = 25
    
        self.dict_of_keys = {'station_name':'Stasjonsnavn', 'department':'Avdeling', 'model':'Modell',
                        'serial':'Meridanr.', 'deliverydate':'Leveringsdato', 'FK_mod_hos':'Sykehus', 'FK_mod_com':u'Leverandør',
                        'lab':'Lab', 'FK_mod_ppl':'Kontaktperson', 'comment':'Kommentar', 'has_dap':'DAP?',
                        'modality_name':'Modalitet', 'discarddate':'Kassasjonsdato?', 'has_aek':'AEK?',
                        'responsibility':'Ansvar?', 'mobile':'Mobil?', 'detector':'Detektor'}
        self.keys_in_order = ['station_name', 'department', 'model', 'serial', 'deliverydate', 'FK_mod_hos', 'FK_mod_com',
                     'lab', 'FK_mod_ppl', 'comment', 'has_dap', 'modality_name', 'discarddate', 'has_aek',
                     'responsibility', 'mobile', 'detector']
        
        self.old_dictionary = {}
        self.new_dictionary = {}
        self.name_list = {}

        # A list to find person / hospital / company name from ID
        table_dict = {'FK_mod_ppl': 'people', 'FK_mod_hos':'hospital', 'FK_mod_com':'company'}
        table_key_dict = {'people': ['ppl_id', 'ppl_name'],
                          'hospital': ['hos_id', 'hos_name'],
                          'company': ['com_id', 'com_name'] }

        name_no = {'people' : 'kontaktperson', 'hospital' : 'sykehus', 'company' : u'leverandør'}
                          
        for key, table in table_dict.items():
            cur.execute("SELECT {} FROM {}".format( ", ".join(table_key_dict[table]), table))
            result = cur.fetchall()
            if not key in self.name_list.keys():
                self.name_list[key] = {}
            for line in result:
                self.name_list[key][line[0]] = line[1]
                
            self.name_list[key]['-1'] = ""
            self.name_list[key]['-2'] = u"Legg til {}".format(name_no[table])
        
        if self.action == 'edit':
            cur.execute("SELECT m.station_name, m.department, m.model, m.serial, m.deliverydate, \
                        m.FK_mod_hos, m.FK_mod_com, m.lab, m.FK_mod_ppl, m.comment, m.has_dap, \
                        m.modality_name, m.discarddate, m.has_aek, m.responsibility, m.mobile, m.detector, m.mod_id \
                        FROM modality m \
                        WHERE m.station_name = ?", (self.station_name,))
            result = cur.fetchall()
            result = list(result[0])

            self.machineId = result[-1]

            for enum, key in enumerate(self.keys_in_order):
                if key in ['has_aek', 'has_dap', 'responsibility', 'mobile']:
                    bi = result[enum]                    
                    if bi == 1:
                        result[enum] = "Ja"
                    elif bi == 0:
                        result[enum] = "Nei"

                self.old_dictionary[key] = result[enum]
                self.new_dictionary[key] = result[enum]
        
        else:
            for key in self.keys_in_order:
                self.old_dictionary[key] = ""
                self.new_dictionary[key] = ""
        
        self.label_entries = {}
        self.old_entries = {}
        self.new_entries = {}
        self.oldtext = {}
        self.newtext = {}
        self.labelname = {}
        
        self.var_job = StringVar()
        self.var_contact = StringVar()

        self.title_label_container = Frame(self.labelContainer, borderwidth=5)
        self.title_old_container = Frame(self.oldContainer, borderwidth=5)
        self.title_new_container = Frame(self.newContainer, borderwidth=5)

        self.body_label_container = Frame(self.labelContainer, borderwidth=5)
        self.body_old_container = Frame(self.oldContainer, borderwidth=5)
        self.body_new_container = Frame(self.newContainer, borderwidth=5)
        
        self.title_label = Label(self.title_label_container, text=u"Nøkler")
        self.old_label = Label(self.title_old_container, text=u"Gamle verdier")
        self.new_label = Label(self.title_new_container, text="Nye verdier")
        
        for key, name in self.dict_of_keys.items():
            self.labelname[key] = StringVar(self, value=name)
            if not key in ['FK_mod_hos', 'FK_mod_com', 'FK_mod_ppl']:
                self.oldtext[key] = StringVar(self, value=self.old_dictionary[key])
                self.newtext[key] = StringVar(self, value=self.new_dictionary[key])
                self.label_entries[key] = Entry(self.body_label_container, textvariable=self.labelname.get(key), state=DISABLED, selectborderwidth=2)
                self.new_entries[key] = Entry(self.body_new_container, textvariable=self.newtext.get(key), selectborderwidth=2, width=50)
                self.old_entries[key] = Entry(self.body_old_container, textvariable=self.oldtext.get(key), state=DISABLED, selectborderwidth=2, width=50)
            else:
                self.oldtext[key] = StringVar(self, value=self.name_list.get(key).get(self.old_dictionary.get(key)))
                self.old_entries[key] = Entry(self.body_old_container, textvariable=self.oldtext.get(key), state=DISABLED, selectborderwidth=2, width=50)
                self.newtext[key] = StringVar(self, value=self.name_list.get(key).get(self.new_dictionary.get(key)))
                self.label_entries[key] = Entry(self.body_label_container, textvariable=self.labelname.get(key), state=DISABLED, selectborderwidth=2)
                self.new_entries[key] = OptionMenu(self.body_new_container, self.newtext.get(key), *self.name_list.get(key).values())
        
        # add new hospital / company / contact person
        self.newtext['FK_mod_hos'].trace('w', self.updateoptions)
        self.newtext['FK_mod_com'].trace('w', self.updateoptions)
        self.newtext['FK_mod_ppl'].trace('w', self.updateoptions)
        
        buttonwidth = 40
        self.ok_button = Button(self.lowerContainer, text="OK", command=self.okcommand, width=buttonwidth)
        self.cancel_button = Button(self.lowerContainer, text="Cancel", command=self.cancelcommand, width=buttonwidth)
        
        self.titleContainer.pack(side=TOP)
        self.label_title.pack()        
        
        self.topContainer.pack(side=TOP)
        
        self.labelContainer.pack(side=LEFT)
        self.title_label_container.pack()
        self.title_label.pack()
        self.body_label_container.pack()
        
        self.pady = 5.5
        
        for key in self.keys_in_order:
            if not key in ['FK_mod_hos', 'FK_mod_com', 'FK_mod_ppl']:
                self.label_entries[key].pack()
            else:
                self.label_entries[key].pack(pady=self.pady)
                
        if not action == "add":
            self.oldContainer.pack(side=LEFT)
            self.title_old_container.pack()    
            self.old_label.pack()
            self.body_old_container.pack()
            
            for key in self.keys_in_order:
                if not key in ['FK_mod_hos', 'FK_mod_com', 'FK_mod_ppl']:
                    self.old_entries[key].pack()
                else:
                    self.old_entries[key].pack(pady=self.pady)
        
        self.newContainer.pack(side=LEFT)
        self.title_new_container.pack()
        self.new_label.pack()
        self.body_new_container.pack()

        first = True
        for key in self.keys_in_order:
            self.new_entries[key].pack()
            if first:
                self.new_entries[key].focus()
                first = False
            
        for key in self.keys_in_order:
            self.new_entries[key].lift()

        self.lowerContainer.pack()
        self.ok_button.pack(side=LEFT)
        self.cancel_button.pack(side=LEFT)
        
        self.pack()
        
        self.parent.focus_force()

    
    def updateoptions(self, *args):

        # Has the user chosen the "add *" option for institution, company or contact person?
        hos = self.newtext['FK_mod_hos'].get().lower() == "legg til sykehus"
        com = self.newtext['FK_mod_com'].get().lower() == u"legg til leverandør"
        ppl = self.newtext['FK_mod_ppl'].get().lower() == "legg til kontaktperson"

        if hos or com or ppl: # only continue if one of the flags are active
        
            if hos:
                title_text = "Nytt sykehus"
                msg_text = u"Navn på sykehus: "
                contact_text = u"Kontaktperson: "
            
            elif com:
                title_text = u"Ny leverandør"
                msg_text = u"Navn på leverandør: "
                
            elif ppl:
                title_text = "Ny kontaktperson"
                msg_text = u"Navn på kontaktperson: "
                job_text = u"Arbeidstittel: "
                
            else:
                raise ThisShouldntHappen('Only one ADD flag can be active. hos: {}, com: {}, ppl: {}.'.format(hos, com, ppl))
            
            self.var_add = StringVar()
            
            self.top = Toplevel()
            self.top.wm_attributes("-topmost", 1)
            self.top.focus_force()
            
            self.top.bind('<Return>', self.add_ok)
            
            self.top.title(title_text)
            
            self.top_topContainer = Frame(self.top)
            self.top_middleContainer = Frame(self.top)
            self.top_bottomContainer = Frame(self.top)
            
            self.top_topContainer.pack()
            
            if hos or ppl:
                self.top_middleContainer.pack()
            
            self.top_bottomContainer.pack()
            
            
            self.add_label = Label(self.top_topContainer, text=msg_text)
            self.add_label.pack(side=LEFT)
            
            self.add_entry = Entry(self.top_topContainer, textvariable=self.var_add)
            self.add_entry.pack(side=LEFT)
            self.add_entry.focus()
            
            if hos:
                self.contact_label = Label(self.top_middleContainer, text=contact_text)
                self.contact_label.pack(side=LEFT)
                self.contact_entry = Entry(self.top_middleContainer, textvariable=self.var_contact)
                self.contact_entry.pack(side=LEFT)
                
            if ppl:
                self.job_label = Label(self.top_middleContainer, text=job_text)
                self.job_label.pack(side=LEFT)
                self.job_entry = Entry(self.top_middleContainer, textvariable=self.var_job)
                self.job_entry.pack(side=LEFT)
            
            self.ok_button = Button(self.top_bottomContainer, text="OK", command=self.add_ok)
            self.cancel_button = Button(self.top_bottomContainer, text="Cancel", command=self.add_cancel)
            
            self.ok_button.pack(side=LEFT)
            self.cancel_button.pack(side=LEFT)
        
            # add Entry + "ADD" button
            # "ADD" button will add to database + list + remove Entry/ADD + change OptionMenu to new choice.
    
    def add_ok(self, *args):
        
        self.top.destroy()
        
        hos = self.newtext['FK_mod_hos'].get().lower() == "legg til sykehus"
        com = self.newtext['FK_mod_com'].get().lower() == u"legg til leverandør"
        ppl = self.newtext['FK_mod_ppl'].get().lower() == "legg til kontaktperson"

        if hos:
            key = 'FK_mod_hos'
            kname = "sykehus"
        elif com:
            key = 'FK_mod_com'
            kname = u"leverandører"
        elif ppl:
            key = 'FK_mod_ppl'
            kname = "kontaktpersoner"
        else:
            raise ThisShouldntHappen('Only one ADD flag can be active. hos: {}, com: {}, ppl: {}.'.format(hos, com, ppl))

        if not self.var_add.get().lower() in [x.lower() for x in self.name_list[key].values()]:

            # the SQL ID is virtual, i.e. needs to be added into SQL
            SQLIDs = [int(x) for x in self.name_list.get(key).keys() if int(x) > 0]
            if SQLIDs:
                newSQLID = max( SQLIDs ) + 1
            else:
                newSQLID = 0
            
            print "New SQL ID: {}".format(newSQLID)
    
            self.name_list[key][newSQLID] = self.var_add.get()
            self.newtext[key].set(self.var_add.get())
            
            m = self.new_entries[key].children['menu']
            m.delete(0,END)
            
            # We want the EMPTY and ADD NEW options to be at the bottom of the list
            newvalues = [v for k, v in self.name_list[key].items() if int(k) > 0] + \
                        [v for k, v in self.name_list[key].items() if int(k) < 0]
                        
            for val in newvalues:
                m.add_command(label=val, command=lambda v=self.newtext[key],l=val:v.set(l))
        else:
            self.top = Toplevel()
            self.top.title("Beskjed")
            self.top.wm_attributes("-topmost", 1)
            self.top.focus_force()
            
            for k, v in self.name_list[key].items():
                if v.lower() == self.var_add.get().lower():
                    CaseSensitiveName = v
                    break

            self.newtext[key].set(CaseSensitiveName)

            msg = Message(self.top, text="{} finnes allerede i listen over {}.".format(CaseSensitiveName, kname))
            msg.pack()
            
            button = Button(self.top, text="OK", command=self.top.destroy)
            button.pack()
            
            self.top.bind('<Return>', self.close_top)
            
    def close_top(self, *args):
        self.top.destroy()
    
    def add_cancel(self):
        self.var_add.set(None)

        hos = self.newtext['FK_mod_hos'].get().lower() == "legg til sykehus"
        com = self.newtext['FK_mod_com'].get().lower() == u"legg til leverandør"
        ppl = self.newtext['FK_mod_ppl'].get().lower() == "legg til kontaktperson"

        if hos:
            key = 'FK_mod_hos'
        elif com:
            key = 'FK_mod_com'
        elif ppl:
            key = 'FK_mod_ppl'
        else:
            raise ThisShouldntHappen('Only one ADD flag can be active. hos: {}, com: {}, ppl: {}.'.format(hos, com, ppl))

        self.newtext[key].set("")
        
        self.top.destroy()
    
    def okcommand(self):
        # Set all newtext{} values to the inputted strings
        for key in self.newtext.keys():

            # Fix readability. We set 1 -> Ja and 0 -> No. To get back into SQL
            # we need to revert this change            
            if key in ['has_aek', 'has_dap', 'responsibility', 'mobile']:
                bi = self.newtext[key].get()
                if bi.lower() in ['ja', 'j', 'yes', 'y', 1]:
                    self.newtext[key].set(1)
                elif bi.lower() in ['nei', 'n', 0]:
                    self.newtext[key].set(0)
                    
            if key in ['FK_mod_ppl', 'FK_mod_hos', 'FK_mod_com']:
                # we need ID from name
                for idx, vv in self.name_list[key].items():
                    if vv == self.newtext[key].get():
                        self.newtext[key].set(idx)
#                        print "substituted name {} with idx {}".format(vv, idx)
        
            if self.newtext[key].get() in ["None", ""]:
                self.newtext[key].set(None)
                
            self.new_dictionary[key] = self.newtext[key].get()
        self.edited = True
        
        self.parent.destroy()
        
    def getJob(self):
        return self.var_job.get().lower()

    def getContact(self):
        return self.var_contact.get().lower()
        
    def cancelcommand(self):
        self.edited = False
        self.parent.destroy()
        
    def getState(self):
        return self.edited
        
    def getParameters(self):
        return self.new_dictionary
        
    def getSQLIDs(self):
        return {'FK_mod_hos' : self.name_list['FK_mod_hos'],
                'FK_mod_com' : self.name_list['FK_mod_com'],
                'FK_mod_ppl' : self.name_list['FK_mod_ppl']}
    
    def myQuit(self):
        self.edited = False
        self.parent.destroy()
        
    def getId(self):
        return self.machineId