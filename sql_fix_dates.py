# -*- coding: utf-8 -*-
"""
Created on Wed Nov 05 14:08:36 2014

@author: rttn
"""

from sql_connection import *

cur = db.cursor()
cur.execute('pragma foreign_keys=ON')

cur.execute('select mod_id from modality')
mod_ids = cur.fetchall()

for mod_id in mod_ids:
    mod_id = mod_id[0]
    cur.execute('select deliverydate, discarddate from modality where mod_id = ?', (mod_id,))
    results = cur.fetchall()[0]
    delivery = results[0]
    discard = results[1]
    
    if delivery and "." in delivery:
        d = delivery.split(".")
        correct_delivery = "{yyyy}-{mm}-{dd}".format(dd = d[0], mm = d[1], yyyy = d[2])
        cur.execute("update modality set deliverydate = ? where mod_id = ?", (correct_delivery, mod_id))    
        
    if discard and "." in discard:
        d = discard.split(".")
        
        correct_discard = "{yyyy}-{mm}-{dd}".format(dd = d[0], mm = d[1], yyyy = d[2])
        cur.execute("update modality set discarddate = ? where mod_id = ?", (correct_discard, mod_id))
        
cur.execute('select * from modality')
print cur.fetchall()

db.commit()