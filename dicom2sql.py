# -*- coding: utf-8 -*-
"""
Created on Mon Jun 17 12:08:50 2013

@author: Testulf
"""

import dicom, os, numpy
from sql_connection import *
from classes import *
from Tkinter import *

do_dicom = True

def enumdict(listed):
    """Return dictionary with indexes."""
    myDict = {}
    for i, x in enumerate(listed):
        myDict[x] = i
        
    return myDict

def dicom2sql(files):
    cur = db.cursor()
    cur.execute('pragma foreign_keys=ON')
    machines = []
    
    contact = None
    delivery = None
    lab = None
    serial = None
    qa = None
    discard = None
    comment = None
    dap = -1
    aek = -1
    responsibility = 1
    mobile = 0
    
    if do_dicom: # read input data from files
                    # WHY CANNOT EVERYONE JUST USE UNICODE? I SHOULD CONVERT TO PYTHON 3!
        for fileName in files:
            ds = dicom.ReadFile(fileName, stop_before_pixels=True)
            station_name = unicode(ds[0x8, 0x1010].value)
            manufacturer = unicode(ds[0x8, 0x70].value)
            series_date = unicode(ds[0x8,0x20].value)
            institution = unicode(ds[0x8,0x80].value)
            modality = unicode(ds[0x8,0x60].value)
            try:
                model = unicode(ds[0x8, 0x1090].value)
            except:
                model = None
            try:    
                department = unicode(ds[0x8,0x1040].value)
            except KeyError:
                department = "N/A"
            try:
                software = unicode(ds[0x18, 0x1020].value)
            except:
                software = None
            
            try:
                detector = unicode(ds[0x18, 0x7006].value) # Detector description
            except:
                detector = None
        
            alreadyExists = False    
                
            if machines > 0:
                for machine in machines:
                    if machine.getStationName() == station_name:
                        alreadyExists = True
            if alreadyExists:
        #        print "Modality of filename", fileName, "already in list"        
                continue
            
            else:
                machines.append(Modality())
                
                if institution:
                    machines[-1].addInstitution(institution)
                if department:
                    machines[-1].addDepartment(department)
                if station_name:
                    machines[-1].addStationName(station_name)
                if manufacturer:
                    machines[-1].addManufacturer(manufacturer)
                if model:
                    machines[-1].addModel(model)
                if detector:
                    machines[-1].addDetector(detector)
                if software:
                    machines[-1].addSoftware(software)
                if modality:
                    machines[-1].addModality(modality)
                if serial:
                    machines[-1].addSerial(serial)
                if delivery:
                    machines[-1].addDeliveryDate(delivery)
                if contact:
                    machines[-1].addContact(contact)
        
                if lab:
                    machines[-1].addLab(lab)
                if not responsibility:
                    machines[-1].addResponsibility(responsibility)
                machines[-1].addMobile(mobile)
                if comment:
                    machines[-1].addComment(comment)
                
                if qa:
                    if numpy.shape(qa) == (2,):
                        machines[-1].addQA(qa[0], qa[1])
                    else:
                        for eachQA in qa:
                            machines[-1].addQA(eachQA[0], eachQA[1])
                if dap > -1:
                    machines[-1].addHasDap(dap)
                if aek > -1:
                    machines[-1].addHasAek(aek)
                    
                if discard:
                    machines[-1].addDiscardDate(discard)
                
            contact = None
            lab = None
            delivery = None
            serial = None
            qa = None
            discard = None
            aek = -1
            dap = -1
            model = None
            responsibility = 1
            comment = None
            mobile = 0
    
    ############################################################################    
    # We are done with input
    # Insert into SQL
    
    # sanity check with user input
    
    all_hospitals = {}
#    all_people = []
    all_departments = {}
    for modality in machines:
        if not modality.getInstitution() in all_hospitals.keys():
            all_hospitals[modality.getInstitution()] = modality.getStationName()
        if not modality.getDepartment() in all_departments.keys():
            all_departments[modality.getDepartment()] = modality.getStationName()
    
#    print u"før første root = Tk()"
    root = Tk()
#    print u"Før første CorrectListOfTags"
    checkHospitalTags = CorrectListOfTags(root, all_hospitals, "Sykehus")
#    print u"Før root.mainloop()"
    root.mainloop()
    
#    print u"Før getState"
    if checkHospitalTags.getState():
        updatedHospitalList = checkHospitalTags.getParameters()
#    print "Etter getstate"
    root = Tk()
    checkDepartmentTags = CorrectListOfTags(root, all_departments, "Avdeling")
    root = mainloop()
    if checkDepartmentTags.getState():
        updatedDepartmentList = checkDepartmentTags.getParameters()

        
#    print "Gammel liste vs oppdatert liste for avd.:", all_hospitals.keys(), updatedHospitalList
    
    hospitals_to_update = zip(all_hospitals.keys(), updatedHospitalList)
    departments_to_update = zip(all_departments.keys(), updatedDepartmentList)
    
    for oldHospital, updatedHospital in hospitals_to_update:
        for modality in machines:
            if modality.getInstitution() == oldHospital:
                modality.addInstitution(updatedHospital) # add = set
    
    for oldDepartment, updatedDepartment in departments_to_update:
        for modality in machines:
            if modality.getDepartment() == oldDepartment:
                modality.addDepartment(updatedDepartment) # add = set
    
    for modality in machines:
    
        # start SQL insert
        FK_mod_hos = None
        FK_mod_com = None
        FK_mod_ppl = None
        FK_mod_sw = None
        
        # get contact ID, insert if not found
                     
        if modality.getContact():
            cur.execute("SELECT ppl_id, ppl_job, ppl_phone FROM people WHERE ppl_name = ?", (modality.getContact(),))
            result = cur.fetchall()
            if result:
                assert len(result) == 1
                FK_mod_ppl = result[0][0]
                
                if (not result[0][1] or not result[0][2]) and modality.getContact() in phone_numbers.keys():
                    cur.execute("UPDATE people SET ppl_job = :job, ppl_phone = :phone WHERE ppl_name = :name",
                                {"job" : phone_numbers[modality.getContact()][0],
                                 "phone" : phone_numbers[modality.getContact()][1],
                                 "name" : modality.getContact()})
                                 
            if not result:
                print modality.getContact(), "not found, adding into database."
    
                cur.execute("INSERT INTO people (ppl_name, ppl_job, ppl_phone) VALUES (?, ?, ?)", \
                        (modality.getContact(), phone_numbers[modality.getContact()][0], phone_numbers[modality.getContact()][1]))
                
                #cur.execute("SELECT LAST_INSERT_ID()")
                FK_mod_ppl = cur.lastrowid # fetchall()[0][0]
                
        # get hospital ID, insert if not found
        if modality.getInstitution():
            cur.execute("SELECT hos_id FROM hospital WHERE hos_name = ?", (modality.getInstitution(),))
            result = cur.fetchall()
            if result:        
                assert len(result) == 1
                FK_mod_hos = result[0][0]
            if not result:
                print modality.getInstitution(), "not found, adding into database."
                cur.execute("INSERT INTO hospital (hos_name) VALUES(?)", (modality.getInstitution(),))
#                cur.execute("SELECT LAST_INSERT_ID()")
                FK_mod_hos = cur.lastrowid # fetchall()[0][0]
                
        # get company ID, insert if not found
        if modality.getManufacturer():
            cur.execute(u"SELECT com_id FROM company WHERE com_name = ?", (modality.getManufacturer(),))
            result = cur.fetchall()
            if result:
                assert len(result) == 1
                FK_mod_com = result[0][0]
            if not result:
                print modality.getManufacturer(), "not found, adding into database."
                cur.execute(u"INSERT INTO company (com_name) VALUES(?)", (modality.getManufacturer(),))
                #        cur.execute("SELECT LAST_INSERT_ID()")
                FK_mod_com = cur.lastrowid # fetchall()[0][0]

        # insert modality into database
        cur.execute(u"INSERT INTO modality (station_name, department, model, deliverydate, discarddate, \
                                            lab, serial, FK_mod_hos, FK_mod_com, FK_mod_ppl, modality_name, \
                                            has_dap, comment, has_aek, responsibility, mobile, detector) \
                                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?, ?,?)", \
                            (modality.getStationName(), modality.getDepartment(), modality.getModel(), 
                               modality.getDeliveryDate(), modality.getDiscardDate(), modality.getLab(), modality.getSerial(),
                                FK_mod_hos, FK_mod_com, FK_mod_ppl, modality.getModality(), modality.getHasDap(), modality.getComment(),
                                modality.getHasAek(), modality.getResponsibility(), modality.getMobile(), modality.getDetector()))
                                # get modality ID
    
#        cur.execute("SELECT LAST_INSERT_ID()")
        mod_id = cur.lastrowid # fetchall()[0][0]
        
        if modality.getSoftware():
            try:
                cur.execute("SELECT sw_id FROM software WHERE sw_name = ?", str(modality.getSoftware(),))
            except:
                print modality.getSoftware()
            result = cur.fetchall()
            if result:
                assert len(result) == 1
                FK_mod_sw = result[0][0]
            if not result:
                print modality.getSoftware(), "not found, adding into database."
                cur.execute("INSERT INTO software (sw_name, FK_sw_mod) VALUES(?, ?)", (unicode(modality.getSoftware()), mod_id))
        
        if modality.getQA():
            qa = modality.getQA()
            cur.execute("SELECT study_date FROM qa WHERE FK_qa_mod = ?", (mod_id,))
            study_dates = cur.fetchall()
    #        if numpy.shape(qa) == (1,2):
    #            if not qa[0][0] in study_dates:
    #                cur.execute("INSERT INTO qa (study_date, doc_number, FK_qa_mod) VALUES (%s, %s, %s)", 
    #                            (qa[0][0], qa[1], mod_id))
    #                
    #        elif numpy.shape(qa) == (2,2):
            for eachQA in qa:
                if not eachQA[0] in study_dates:
                    cur.execute("INSERT INTO qa (study_date, doc_number, qa_comment, FK_qa_mod) VALUES (?, ?, ?, ?)", 
                                (eachQA[0], eachQA[1], eachQA[2], mod_id))
        
        FK_mod_hos = None
        FK_mod_com = None
        FK_mod_ppl = None
        FK_mod_sw = None
        
#    cur.execute("ALTER TABLE qa MODIFY COLUMN qa_comment TEXT CHARACTER SET utf8 COLLATE utf8_general_ci")
#    cur.execute("ALTER TABLE dap MODIFY COLUMN dap_comment TEXT CHARACTER SET utf8 COLLATE utf8_general_ci")
#    cur.execute("ALTER TABLE fluoro_tube MODIFY COLUMN ftkp_comment TEXT CHARACTER SET utf8 COLLATE utf8_general_ci")
#    cur.execute("ALTER TABLE fluoro_tube MODIFY COLUMN ftdxd_comment TEXT CHARACTER SET utf8 COLLATE utf8_general_ci")
#    cur.execute("ALTER TABLE fluoro_tube MODIFY COLUMN ftpxd_comment TEXT CHARACTER SET utf8 COLLATE utf8_general_ci")
#    cur.execute("ALTER TABLE tube_qa MODIFY COLUMN tqa_comment TEXT CHARACTER SET utf8 COLLATE utf8_general_ci")
#    cur.execute("ALTER TABLE fluoro_iq MODIFY COLUMN fiq_comment TEXT CHARACTER SET utf8 COLLATE utf8_general_ci")
#    cur.execute("ALTER TABLE aek MODIFY COLUMN aek_comment TEXT CHARACTER SET utf8 COLLATE utf8_general_ci")    
        
    db.commit()
    list_of_added_modalities = []
    for modality in machines:
        list_of_added_modalities.append(modality.getStationName())
    return u"Har lagt til følgende maskiner: ", ",".join(list_of_added_modalities)
