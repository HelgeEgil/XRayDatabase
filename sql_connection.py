# -*- coding: utf-8 -*-
"""
Created on Mon Jun 23 14:21:48 2014

@author: Testulf
"""
import sqlite3, re
from sqlite3 import OperationalError, ProgrammingError

def exec_sql_file(cursor, sql_file):
    print "\n[INFO] Executing SQL script file: '%s'" % (sql_file)
    statement = ""

    for line in open(sql_file):
        if re.match(r'--', line):  # ignore sql comment lines
            continue
        if not re.search(r'[^-;]+;', line):  # keep appending lines that don't end in ';'
            statement = statement + line
        else:  # when you get a line ending in ';' then exec statement and reset for next statement
            statement = statement + line
            #print "\n\n[DEBUG] Executing SQL statement:\n%s" % (statement)
            try:
                cursor.execute(statement)
            except (OperationalError, ProgrammingError) as e:
                print "\n[WARN] MySQLError during execute statement \n\tArgs: '%s'" % (str(e.args))
            statement = ""
            
def exec_sql_file_oneline(cursor, sql_file):
    print "\n[INFO] Executing SQL script file: '%s'" % (sql_file)
    statement = ""

    for line in open(sql_file):
        if re.match(r'--', line):  # ignore sql comment lines
            continue
        if not re.search(r'[^-;]+;', line):  # keep appending lines that don't end in ';'
            statement = statement + line
        else:  # when you get a line ending in ';' then exec statement and reset for next statement
            statement = statement + line
            #print "\n\n[DEBUG] Executing SQL statement:\n%s" % (statement)
    try:
        cursor.execute(statement)
    except (OperationalError, ProgrammingError) as e:
        print "\n[WARN] MySQLError during execute statement \n\tArgs: '%s'" % (str(e.args))
    return list(map(lambda x: x[0].decode('utf-8'), cursor.description)), cursor.fetchall() 

def isSQLite3(filename):
    from os.path import isfile, getsize

    if not isfile(filename):
        return False
    if getsize(filename) < 100: # SQLite database file header is 100 bytes
        return False
    else:
        fd = open(filename, 'rb')
        Header = fd.read(100)
        fd.close()

        if Header[0:16] == 'SQLite format 3\000':
            return True
        else:
            return False

database_file = 'sqlite_xray.db'
db_nodb = False

try:                               
    db = sqlite3.connect(database_file)
                      
except Exception as e:
    print "Cannot load database mydb due to {}".format(e)