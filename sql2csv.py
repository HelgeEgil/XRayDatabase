# -*- coding: utf-8 -*-
"""
Created on Mon Jun 17 12:08:50 2013

@author: Testulf
"""
import dicom, os, numpy, re
from sql_connection import *
from classes import *
import subprocess

do_dicom = True

def enumdict(listed):
    """Return dictionary with indexes."""
    myDict = {}
    for i, x in enumerate(listed):
        myDict[x] = i
        
    return myDict

def num(s):
    try:
        return int(s)
    except ValueError:
        return float(s)

def sql2csv():
    cur = db.cursor()
    cur.execute('pragma foreign_keys=ON')
    
    tables = ['helseforetak', 'hospital', 'software', 'company',
              'people', 'modality', 'qapeople',
              'qa', 'aek', 'aek_each',
              'aek_calibration', 'aek_calibration_each',
              'tube_qa', 'tube_qa_large', 'tube_qa_small',
              'fluoro_tube', 'fluoro_tube_kvprec_large',
              'fluoro_tube_kvprec_small', 'fluoro_tube_dxdose',
              'fluoro_tube_pxdose', 'dap', 'dap_each', 
              'fluoro_iq', 'fluoro_iq_each']

    for tableName in tables:
        outputStringList = []
        dbString = "SELECT * FROM {}".format(tableName)
        cur.execute(dbString)
        field_names = [k[0] for k in cur.description]
        columns = cur.fetchall()
        
        outputStringList.append(";".join(field_names))
        for tpl in columns:
            line = [unicode(x) for x in tpl]
            outputStringList.append(";".join(line))
        
        outputString = "\n".join(outputStringList)
        
        fileName = "csv\%s.csv" % tableName
        outputFile = open(fileName, "w")
        
        outputFile.write(outputString)
        
    abspath = os.path.abspath(os.path.join(__file__, os.path.pardir, "csv"))
#    print abspath
#    subprocess.Popen(r'explorer /select,"{}"'.format(abspath))
    os.startfile(abspath)
        
def csv2sql():
    cur = db.cursor()
    cur.execute('pragma foreign_keys=OFF')
    
#    cur.execute("SET FOREIGN_KEY_CHECKS=0")
    
    tables = ['software', 'modality', 'helseforetak', 'hospital', 'company',
          'people', 'qapeople',
          'qa', 'aek', 'aek_each',
          'aek_calibration', 'aek_calibration_each',
          'tube_qa', 'tube_qa_large', 'tube_qa_small',
          'fluoro_tube', 'fluoro_tube_kvprec_large',
          'fluoro_tube_kvprec_small', 'fluoro_tube_dxdose',
          'fluoro_tube_pxdose', 'dap', 'dap_each', 
          'fluoro_iq', 'fluoro_iq_each']

    for tableName in tables:
        fileName = "csv\%s.csv" % tableName
        
        inputFile = open(fileName, "r")
        headers = inputFile.readline().replace("\n", "").split(";")
        
        # We need to clear table before inserting new values...
        # Remember we put set foreign_key_checks = 0, otherwise we would be in trouble
        dbString = "DELETE FROM {}".format(tableName)
        cur.execute(dbString)
        
        for line in inputFile.readlines():
            lineElements = line.replace("\n","").split(";")
            
            lineElements = [unicode(k) for k in lineElements]
            
            for k in range(len(lineElements)):
                if "[" in lineElements[k] and "]" in lineElements[k]:
                    lineElements[k] = str(lineElements[k])
                if 'None' in lineElements[k]:
                    lineElements[k] = None
                    
            # now headers      is [header1, header2, header3, ...]            
            # now lineElements is [column1, column2, column3, ...]
             
#            bind_variables = ", ".join(["%s" for _ in lineElements]) # before sqlite3
            bind_variables = u", ".join([u"?" for _ in lineElements])
            # It is NOT possible to use ? variable insertion as table or variable names, only values.
            # Therefore we use the .format method, where we put in the table name, bind_variables = ? * number of variables
            # and couple the ?,?,?,... to the variables using lineElements
            query = u"INSERT INTO {} ({}) VALUES({})".format(tableName, ", ".join(headers), bind_variables)
#            print "Query: {}".format(query), lineElements
            cur.execute(query, lineElements)

    cur.execute("pragma foreign_keys=ON")
    db.commit()

def createSQL():
    if isSQLite3(database_file): # defined in sql_connection. default = 'sqlite_xray.db'
        # returns true there is a file named database_file, with correct sqlite3 headers
        root = Tk()
        question = Question(root, u"Fant allerede en database i filen `{}`. Vil du slette den?".format(database_file))
        root.mainloop()
        deleteExistingDatabase = question.getAnswer()
        noDataBase = False
    else:
        deleteExistingDatabase = False
        noDataBase = True
        db = sqlite3.connect(database_file)

    if deleteExistingDatabase:
        db = sqlite3.connect(database_file)
        db.close() # if connection was established in new_sql_connection.py
        os.remove(database_file)
        db = sqlite3.connect(database_file)
            
#    if deleteExistingDatabase or noDataBase: # a) no existing database b) former, deleted database
    
    if True:
        try:
            cur = db.cursor()
        except:
            db = db_nodb
            cur = db.cursor()
        
        cur.execute('pragma foreign_keys=ON')
        exec_sql_file(cur, "create_sqlite_db_2014-11-24.sql")
     
if __name__ == "__main__":
    csv2sql()