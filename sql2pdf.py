# -*- coding: utf-8 -*- 
"""
Created on Fri Oct 04 08:58:16 2013

@author: rttn
"""
import sys, os

# Add reportlab to path
CURRENT_DIR = os.path.dirname(os.path.abspath("./modules/pdf/"))
sys.path.append(os.path.dirname(CURRENT_DIR + "\pdf"))

# Import reportlab
import pdf
from pdf.theme import colors, DefaultTheme

from datetime import datetime
from sql_connection import *
import numpy as np
from reportlab.lib.units import inch
import webbrowser

author = 'Helge Pettersen'
logo_filename = 'HB_logo.png'

def first_letters(string):
    output = ""
    if 'haukeland' in string.lower():
        return "HUS"
    for i in string.upper().split():
        output += i[0]
    return output

def sql2pdf(reports_to_create_list, shortReport = False, smallTableOnLeeds = False):
    cur = db.cursor()
    cur.execute('pragma foreign_keys=ON')
    
    # Naming
    month_no = ["januar", "februar", "mars", "april", "mai", "juni", "juli", "august", "september", "oktober", "november", "desember"]
    
    # Header sizes (smaller if shortReport == True)
    if shortReport:
        H2 = pdf.H4
        H1 = pdf.H3
    else:
        H2 = pdf.H2
        H1 = pdf.H1
    
    # We begin creating the reports
    ##############################
    for report_to_create in reports_to_create_list:
        # sqlite format: cur.execute("SELECT * FROM table WHERE value = ?", (list, of, unpacked, arguments,))
        # If there is only one value (one ? in text) to be inserted, the argument must still be a list: "(value, )" not "value"
        
        # Note: We find the QA based on input doc_number
        cur.execute("SELECT study_date, FK_qa_mod, qa_id, qa_comment FROM qa WHERE doc_number = ?", (report_to_create,))
        result = cur.fetchall()
        study_date = result[0][0]
        study_date_parse = datetime.strptime(study_date, '%Y-%m-%d')
        mod_id = result[0][1]
        qa_id = result[0][2]
        qa_comment = "%s" % result[0][3]
        
        cur.execute("SELECT m.model, m.lab, c.com_name, h.hos_name, m.department, m.modality_name \
                        FROM modality m \
                        INNER JOIN company c ON c.com_id = m.FK_mod_com \
                        INNER JOIN hospital h ON h.hos_id = m.FK_mod_hos \
                        WHERE m.mod_id = ?", (mod_id,)) # this mod_id from the last SELECT
        result = cur.fetchall()
        model_name = result[0][0]
        lab = result[0][1]
        company = result[0][2]
        hospital = result[0][3]
        department = result[0][4]
        modality = result[0][5]
                    
        # Find people in QA
        # was left inner join
        cur.execute("SELECT p.ppl_name, p.ppl_job, p.ppl_id FROM qapeople qp \
                    INNER JOIN people p ON p.ppl_id = qp.FK_qappl_ppl \
                    INNER JOIN qa q ON q.qa_id = qp.FK_qappl_qa \
                    WHERE q.qa_id = ?", (qa_id,))
        result = cur.fetchall()
        people = []
        for row in result:
            # name, job
            if not row[1]:
                job = raw_input("Job of %s? " % row[0])
                cur.execute("UPDATE people SET ppl_job = ? WHERE ppl_id = ?", (job, row[2]))
            else:
                job = row[1]
            people.append([row[0], job])
        
        # Concatenate job title and name
        people_job = ["%s %s" % (x[1], x[0]) for x in people]
        
        # Make string from people_job
        if len(people_job) > 2:
            people_string = ", ".join(people_job[:-1]) + " og " + people_job[-1]
        elif len(people_job) == 2:
            people_string = people_job[0] + " og " + people_job[1]
        else:
            people_string = people_job[0]
        
        # Is DAP test done?
        cur.execute("SELECT COUNT(*) FROM DAP WHERE FK_dap_qa = ?", (qa_id,))
        result = cur.fetchall()
        if result[0][0] > 0:
            is_dap = True
        else:
            is_dap = False
        
        
        # is Automatic Exposure Control (AEK) test done?    
        cur.execute("SELECT count(*) FROM aek WHERE FK_aek_qa = ?", (qa_id,))
        result = cur.fetchall()
        if result[0][0] > 0:
            is_aek = True
        else:
            is_aek = False
        
        # tube QA
        # Is the tube QA (kVp accuracy, HVL, ...) done for CONVENTIONAL labs?
        cur.execute("SELECT COUNT(*) FROM tube_qa WHERE FK_tqa_qa = ?", (qa_id,))
        result = cur.fetchall()
        if result[0][0] > 0:
            is_tqa = True
        else:
            is_tqa = False
            
        # fluoro dose measurements (to skin + to detector)
        cur.execute("SELECT ft_id FROM fluoro_tube WHERE FK_ft_qa = ?", (qa_id,))
        result = cur.fetchall()[0][0]
        if result > 0:
            is_ft = True
        else:
            is_ft = False
        ft_id = result
        
        if is_ft:
            # Fluoro dose to skin (pxd = patient dose)
            cur.execute("SELECT ftpxd_id FROM fluoro_tube_pxdose WHERE FK_ftpxd_ft = ?", (ft_id,))
            result = cur.fetchall()
            if len(result) > 0:
                is_ftpxd = True
            else:
                is_ftpxd = False
                
            # Fluoro dose to detector (dxd = detector dose)
            cur.execute("SELECT ftdxd_id FROM fluoro_tube_dxdose WHERE FK_ftdxd_ft = ?", (ft_id,))
            result = cur.fetchall()
            if len(result) > 0:
                is_ftdxd = True
            else:
                is_ftdxd = False
        
        # fluoro tube kV precision
        cur.execute("SELECT COUNT(*) FROM fluoro_tube_kvprec_large WHERE FK_ftkpl_ft = ?", (ft_id,))
        result_large = cur.fetchall()[0][0]
        cur.execute("SELECT COUNT(*) FROM fluoro_tube_kvprec_small WHERE FK_ftkps_ft = ?", (ft_id,))
        result_small = cur.fetchall()[0][0]
        if result_large + result_small > 0:
            is_ftqa = True
        else:
            is_ftqa = False
        
        if result_large > 0:
            is_ftqa_large = True
        else:
            is_ftqa_large = False
        
        if result_small > 0:
            is_ftqa_small = True
        else:
            is_ftqa_small = False
        
        # Leeds IQ
        # This function SHOULD accomodate the other possible IQ plates available in Excel
        # file, since it only gets the contrast % and resolution lp/mm from Excel file.
        cur.execute("SELECT COUNT(*) FROM fluoro_iq WHERE FK_fiq_qa = ?", (qa_id,))
        result = cur.fetchall()[0][0]
        if result > 0:
            is_fiq = True
        else:
            is_fiq = False
        
        #TABLE_WIDTH = 540
        TABLE_WIDTH = 300
        
        class MyTheme(DefaultTheme):
            doc = {
                'leftMargin' : 60,
                'rightMargin' : 60,
                'topMargin' : 30,
                'bottomMargin' : 25,
                'allowSplitting' : False}
        
        class ShortTheme(DefaultTheme): # If "short report" is chosen in window, only one page
            doc = {
                'leftMargin' : 5,
                'rightMargin' : 5,
                'topMargin' : 5,
                'bottomMargin' : 5,
                'allowSplitting' : False}
                
            spacer_height = 0.1 * inch
        
                
        doc = pdf.Pdf('Rapportnr. %s' % report_to_create, author)
        if shortReport:
            doc.set_theme(ShortTheme)
        else:
            doc.set_theme(MyTheme)
        
        # logo + spacer
        
        if not shortReport:
            logo_path = logo_filename
            rescaleFactor = 1.4 # logo rescale factor
            doc.add_image(logo_path, 169*rescaleFactor, 46*rescaleFactor, pdf.LEFT)
        
        if not shortReport:
            doc.add_spacer()
        
        doc.add_header(u'Statuskontroll ved %s - %s' % (hospital, department), H1)
        
        doc.add_paragraph(u"<b>Rapportnr. %s.</b>" % report_to_create)
        doc.add_spacer()
        
        if shortReport:
            breaks = " "
        else:
            breaks = "<br></br><br></br>"
        if lab:
            lab_txt = u" på {}".format(lab)
        else:
            lab_txt = ""
        doc.add_paragraph(u"Det ble utført kvalitetsskontroll på en %s%s ved %s, %s den %d. %s %d.%sKontrollen ble utført av %s." % (
                    company + " " + model_name, lab_txt, department, hospital, study_date_parse.day, month_no[study_date_parse.month-1], 
                    study_date_parse.year, breaks, people_string))
        
        if qa_comment != "None":
            doc.add_paragraph(u"<b>Kommentar for kvalitetskontrollen:</b> %s" % qa_comment)
        
        if not shortReport:
            doc.add_spacer()
        
        #doc.add_header("Alle kontroller på modaliteten")
        
        #doc.add_paragraph("Hva med litt informasjon om FLUORO IQ?")
        
        numName = [u"én", "to", "tre", "fire", "fem", "seks", "syv", u"åtte", "ni", "ti", "elleve", "tolv", "tretten", "fjorten"]
        
        if not shortReport:
            
            dataFormatted = [["Dato", "Rapportnr."]]
            cur.execute("SELECT study_date, doc_number FROM qa WHERE FK_qa_mod = ? ORDER BY study_date DESC", (mod_id,))
            data = cur.fetchall()
            numControl = len(data)
            
            if numControl == 1:
                doc.add_paragraph(u"Det har ikke blitt registrert tidligere kontroller ved dette utstyret.")
            
            else:
                isMultiple = numControl > 1 and "er" or ""
                doc.add_paragraph(u"Det er registrert %s kvalitetskontroll%s ved dette utstyret:" % \
                                (numName[numControl-1], isMultiple))
        
            doc.add_spacer()
            i = 0
            for row in data:
                i += 1
                if row[1] == report_to_create:
                    row_name_for_bold = i # the current QA is shown in green, maybe a bad solution (green could be = OK QA)
        #        dataFormatted.append([str(row[0]), row[1], u"%d år" % (time.localtime().tm_year - row[0].year)])
                dataFormatted.append([str(row[0]), row[1]])
                
            tableStyle_dates = [("BACKGROUND", (0,row_name_for_bold), (-1,row_name_for_bold), colors.lightgreen)]
            
            doc.add_table(dataFormatted, TABLE_WIDTH / 1.5, align="LEFT", extra_style = tableStyle_dates)
        
        if not shortReport:
            doc.add_spacer()
        
        
        # QA begins    
        tqa_string = ""
        
        
        if is_tqa:
            cur.execute("SELECT tqa_id, tqa_fka_cm, tqa_filter_mmcu, tqa_comment FROM tube_qa WHERE \
                            FK_tqa_qa = ?", (qa_id,))
            tubedata = cur.fetchall()[0]
            tqa_id, tqa_fka_cm, tqa_comment = tubedata[0], tubedata[1], u"%s" % tubedata[3]
            
            cur.execute("SELECT COUNT(*) FROM tube_qa_large WHERE FK_tqal_tqa = ? AND tqal_meas_kvp IS NOT NULL", (tqa_id,))
            nTQAL = cur.fetchall()[0][0]
            cur.execute("SELECT COUNT(*) FROM tube_qa_small WHERE FK_tqas_tqa = ? AND tqas_meas_kvp IS NOT NULL", (tqa_id,))
            nTQAS = cur.fetchall()[0][0]
            
        #    print "Det ble funnet %d eksponeringer ved stort fokus, og %d ved lite." % (nTQAL, nTQAS)
            
            if nTQAL > 0 or nTQAS > 0:
                if not shortReport:
                    doc.add_header(u"Kontroll av røntgenrøret", H1)
        
                if not shortReport:
                    if tqa_fka_cm > 0:
                        doc.add_paragraph(u"Røntgenrøret blir målt ved hjelp av et Unfors Xi målekammer. Avstanden mellom fokuspunkt og målekammer \
                                        er %d cm, og det blir ikke benyttet tilleggsfiltrering. Strålefeltet blendes inn til målekammeret." % tqa_fka_cm)
    #                    tqa_string += u"Røntgenrøret blir målt ved hjelp av et Unfors Xi målekammer. Avstanden mellom fokuspunkt og målekammer \
    #                    er %d cm, og det blir ikke benyttet tilleggsfiltrering. Strålefeltet blendes inn til målekammeret." % tqa_fka_cm                
                    else:
                             doc.add_paragraph(u"Røntgenrøret blir målt ved hjelp av et Unfors Xi målekammer. Det blir ikke benyttet tilleggsfiltrering. \
                                             Strålefeltet blendes inn til målekammeret.")
    #                    tqa_string
                if tqa_comment != "None":
                    doc.add_paragraph(u"<b>Kommentar: </b>%s" % tqa_comment)
            if nTQAL > 0:
                # there are large focus tube QA measurements
        
                cur.execute("SELECT tqal.tqal_set_kvp, tqal.tqal_set_mas, tqal.tqal_meas_kvp, \
                            tqal.tqal_meas_microgy, tqal.tqal_meas_hvl_mmal FROM tube_qa_large tqal \
                            WHERE tqal.FK_tqal_tqa = ? AND tqal.tqal_meas_kvp IS NOT NULL", (tqa_id,))
                tqal_data = list(cur.fetchall())
                tqal_table = [["Set kVp", "Set mAs", u"Målt kVp", u"kVp diff."]]
                TQAL_ok = True        
                max_deviation = [0, 0]
                kv_list = [0]
        
                for row in tqal_data:
                    if not row[2] or row[2] == "False": continue
                    if int(row[0]) in kv_list: continue            
                    kv_list.append(int(row[0]))
                    
                    deviation = (float(row[2]) / float(row[0]) - 1) * 100
                    if abs(deviation) > 10:                                    ## Endret fra 5 til 10
                        TQAL_ok = False
                    if abs(deviation) > max_deviation[0]:
                        max_deviation = [abs(deviation), row[0]]
                    tqal_table.append(["%d" % row[0], "%.1f" % row[1], "%.1f" % row[2], "%.1f %%" % deviation])
                
                doc.add_header(u"Nøyaktighet rørspenning (stort fokus)", H2)
                doc.add_paragraph(u"Nøyaktighet i rørspenning ble målt med stort fokus og økende kV-verdier.<br></br>\
                                <b>Krav: </b> Bør være innenfor 5 % av nominell verdi, skal være innenfor 10 %.<br></br>")
                if TQAL_ok:
                    doc.add_paragraph(u"<b>Konklusjon: </b> Godkjent. <br></br> \
                        <b>Resultat: </b> Avviket mellom målt og nominell kV ligger innenfor grenseverdier \
                                        for alle målinger, se tabellen under.")
                else:
                   doc.add_paragraph(u"<b>Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>. <br></br> \
                            <b>Resultat: </b> Ved enkelte kV-verdier ble det målt et større avvik enn kravet tilsier. \
                            Det største avviket var på %.1f %% ved %d kVp. Se tabellen under." % (max_deviation[0], max_deviation[1]))
                if not shortReport:
                    doc.add_spacer()
                    doc.add_table(tqal_table, TABLE_WIDTH/1.1, align="LEFT")
                    doc.add_spacer()
                
            if nTQAS > 0:
                # there are large focus tube QA measurements
        
                cur.execute("SELECT tqas.tqas_set_kvp, tqas.tqas_set_mas, tqas.tqas_meas_kvp, \
                            tqas.tqas_meas_microgy, tqas.tqas_meas_hvl_mmal FROM tube_qa_small tqas \
                            WHERE tqas.FK_tqas_tqa = ? AND tqas.tqas_meas_kvp IS NOT NULL", (tqa_id,))
                tqas_data = list(cur.fetchall())
                tqas_table = [["Set kVp", "Set mAs", u"Målt kVp", u"kVp diff."]]
                TQAS_ok = True        
                max_deviation = [0, 0]
                kv_list = [0]
        
                for row in tqas_data:
                    if not row[2]: continue
                    if int(row[0]) in kv_list: continue       
                    kv_list.append(int(row[0]))
                    
                    deviation = (float(row[2]) / float(row[0]) - 1) * 100
                    if abs(deviation) > 10:                                    ## Endret fra 5 til 10
                        TQAS_ok = False
                    if abs(deviation) > max_deviation[0]:
                        max_deviation = [abs(deviation), row[0]]
                    tqas_table.append(["%d" % row[0], "%.1f" % row[1], "%.1f" % row[2], "%.1f %%" % deviation])
                
                doc.add_header(u"Nøyaktighet rørspenning (lite fokus)", H2)
                doc.add_paragraph(u"Nøyaktighet i rørspenning ble målt med lite fokus og økende kV-verdier.<br></br>\
                                <b>Krav: </b> Innenfor 10 % av nominell verdi.<br></br>")
                if TQAS_ok:
                    doc.add_paragraph(u"<b>Konklusjon: </b> Godkjent. <br></br> \
                        <b>Resultat: </b> Avviket mellom målt og nominell kV ligger innenfor grenseverdier \
                                        for alle målinger, se tabellen under.")
                else:
                   doc.add_paragraph(u"<b>Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>. <br></br> \
                            <b>Resultat: </b> Ved enkelte kV-verdier ble det målt et større avvik enn kravet tilsier. \
                            Det største avviket var på %.1f %% ved %d kVp. Se tabellen under." % (max_deviation[0], max_deviation[1]))
                if not shortReport:
                    doc.add_spacer()
                    doc.add_table(tqas_table, TABLE_WIDTH/1.1, align="LEFT")
                    doc.add_spacer()
                
            if nTQAL > 0:
                # reproducibility tube voltage
                # only with large focus, maybe implement small focus as well?
                
                reproducibility_data = list(tqal_data[:])
                # first sort by kV - 80 (since we want repeat ~80 kV measurements)
                reproducibility_data.sort(key=lambda tup: abs(tup[0]-80))
                for i in range(len(reproducibility_data)):
                    if not reproducibility_data[i][0] == reproducibility_data[i+1][0] or \
                        not reproducibility_data[i+1][1] *0.95 < reproducibility_data[i][1] < 1.05 * reproducibility_data[i+1][1]:
                            # i is the last ~80 kV same mAs measurement                    
                            break
                reproducibility_data = reproducibility_data[:i+1]
                # reproducibility_data only includes ~80 kV same mAs lines
                
                # Find average. np.average(array, axis to make average tuple of, 2nd item in list (average of array[2]))
                # Crashes when reproducibility_data is empty!
                
                print reproducibility_data
                
                kvp_column = np.array(reproducibility_data)[:,2]
                
                kvp_average = np.average(kvp_column)
                kvp_largest_deviation = max([abs(x[2]/kvp_average-1)*100 for x in reproducibility_data])  # already in %
        #        kvp_cov = stats.variation([x[2] for x in reproducibility_data])
                if kvp_largest_deviation < 5:                                  
                    kvp_deviation_ok = True
                else:
                    kvp_deviation_ok = False
                
                dose_average = np.average(reproducibility_data, 0)[3]
                dose_largest_deviation = max([abs(x[3]/dose_average-1)*100 for x in reproducibility_data]) 
        #        dose_cov = stats.variation([x[3] for x in reproducibility_data])        
                if dose_largest_deviation < 5:                                 
                    dose_deviation_ok = True
                else:
                    dose_deviation_ok = False
                
                number_of_measurements_in_reproducibility = numName[len(reproducibility_data)-1]
                kvp_used_in_reproducibility = reproducibility_data[0][0]
                mas_used_in_reproducibility = reproducibility_data[0][1]
                
                doc.add_header(u"Reproduserbarhet rørspenning", H2)
                if not shortReport:
                    doc.add_paragraph(u"Reproduserbarheten i kVp og dose ble funnet ved å måle kVp og dose ved %d kVp og %.1f mAs %s ganger. \
                                    Deretter beregnes målingens gjennomsnittsverdi og variasjonskoeffisienter. Stort fokus benyttes." % \
                                    (kvp_used_in_reproducibility, mas_used_in_reproducibility, number_of_measurements_in_reproducibility))
                
                doc.add_paragraph(u"<b>Krav: </b> Største avvik fra målingens gjennomsnittsverdi skal være innenfor 5 % for kVp og dose.<br></br>")
                if kvp_deviation_ok and dose_deviation_ok:
                    doc.add_paragraph(u"<b>Konklusjon: </b> Godkjent.")
                else:
                    doc.add_paragraph(u"<b>Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>.")
                doc.add_paragraph(u"<b>Resultat: <br></br></b> Gjennomsnittlig rørspenning: %.1f kVp<br></br> \
                                    Største avvik fra gjennomsnittlig rørspenning: %.2f %% <br></br>\
                                    Gjennomsnittlig dose:  %.2f µGy <br></br>\
                                    Største avvik fra gjennomsnittlig dose:  %.2f %% <br></br>" 
                                    % (kvp_average, kvp_largest_deviation, dose_average, dose_largest_deviation))
                if not shortReport:
                    doc.add_spacer()
                
                # Radiation output µGy / mAs
                # we use the table from reproducibility measurements
                
                rad_out_table = [["Set kVp", "Set mAs", u"Målt µGy", "µGy/mAs"]]
                rad_out_ok = True
                for row in reproducibility_data: # kVp, mAs, meas kVp, µGy, HVL
                    rad_out_table.append([row[0], row[1], "%.1f" % row[3], "%.1f" % (row[3] / row[1])])
                    if row[3] / row[2] > 100:
                        rad_out_ok = False
            
                rad_out_list = []
                for row in rad_out_table[1:]:
                    rad_out_list.append(eval(row[-1]))
                
            
                average_rad_out = np.average(rad_out_list)
                
                doc.add_header(u"Stråleutbytte", H2)
                doc.add_paragraph(u"Stråleutbytte måles ved forskjellige rørstrøm- og mAs-innstillinger ved stort fokus. <br></br>\
                                <b>Krav: </b> Anbefalt stråleutbytte ved 80 kVp i en meters avstand ved fast mAs skal ligge \
                                mellom 50 og 100 µGy/mAs. Nye apparater med digital detektor ligger ofte under 50 µGy/mAs.")
                if rad_out_ok:
                    doc.add_paragraph(u"<b>Konklusjon: </b> Godkjent. <br></br> \
                                <b>Resultat: </b> Stråleutbyttet ble målt ved %d kVp og %.1f mAs, og var ved disse innstillingene \
                                gjennomsnittlig %.1f µGy/mAs ved stort fokus. Se tabellen under." %
                                (rad_out_table[1][0], rad_out_table[1][1], average_rad_out))
                else:
                    doc.add_paragraph(u"<b>Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>. <br></br> \
                                <b>Resultat: </b> Stråleutbyttet ble målt ved %d kVp og %.1f mAs, og var ved disse innstillingene \
                                gjennomsnittlig %.1f µGy/mAs ved stort fokus. Se tabellen under." %
                                (rad_out_table[1][0], rad_out_table[1][1], average_rad_out))
                
                if not shortReport:
                    doc.add_spacer()
                    doc.add_table(rad_out_table, align="LEFT")
                    doc.add_spacer()
                
                # HALF VALUE LAYER
                # Data from RP 162
                min_hvl_dict = {50:1.8, 60:2.2, 70:2.5, 80:2.9, 90:3.2, 110:3.9, 125:4.5, 150:4.5}
                
                hvl_data = {}
                for row in tqal_data:
                    kvp = row[0]
                    hvl = row[4]
                    if not kvp in hvl_data.keys() and hvl:
                        hvl_data[kvp] = [hvl,]
                    else:
                        if hvl:
                            hvl_data[kvp].append(hvl)
        
                hvl_table_head = [["kVp", "HVL [mmAl]", "Minimum HVL [mmAl]", "Konklusjon"]]
                hvl_table = []
                
                hvl_ok = True
                hvl_calc_table = []
                
                hvl_kvp_ikke_ok_list = []        
                
                for k, v in hvl_data.items():
                    if k in min_hvl_dict.keys():
                        min_hvl = min_hvl_dict[k]
                    else: # her kommer 40 kV inn
                        min_hvl = 0.0356 * k + 0.0381 # linear interpolation of minimum HVL values from IPEM / IEC
                    average_hvl = np.average(v)
                    if average_hvl < min_hvl:
                        if k>40: hvl_ok = False                           
                        #else: hvl_ok = False       #Sikker på denne linjen skal vekk?
                        hvl_kvp_ikke_ok_list.append(str(k))
                    hvl_table.append([k, "%.1f" % average_hvl, "%.1f" % min_hvl, average_hvl>min_hvl and "Godkjent" or "Ikke godkjent"])
        
                doc.add_header(u"Halvverdilag", H2)
                doc.add_paragraph(u"Halvverdilaget (HVL) er hvor mange mm Al som trengs for halvere intensiteten.")
                doc.add_paragraph(u"<b>Krav: </b>HVL skal være høyere enn minimum HVL ved hver rørspenning.")
        
                
                if not hvl_ok:
                    if len(hvl_kvp_ikke_ok_list) > 1:
                        hvl_kvp_ikke_ok_string = ", ".join(hvl_kvp_ikke_ok_list[:-1]) + " og %s" % hvl_kvp_ikke_ok_list[-1]
                    else:
                        hvl_kvp_ikke_ok_string = hvl_kvp_ikke_ok_list[0]
                
                if hvl_ok:
                    doc.add_paragraph(u"<b>Konklusjon: </b>Godkjent.")
                    for k, v in hvl_data.items():                              # OK, men TUNGVINT
                        min_hvl = 0.0356 * 40 + 0.0381
                        average_hvl = np.average(v)
                        if average_hvl < min_hvl:
                            if k==40: doc.add_paragraph(u"<b>Kommentar: </b>Det skjer iblant at halvverdilaget er lavere enn minimum for 40 kV, og dette kan skyldes høy usikkerhet ved måling av lave kV-verdier.")
                    #if average_hvl[40] < min_hvl[40]:
                else:
                    if '40' in hvl_kvp_ikke_ok_string:
                        om_40 = u" Dette skjer noen ganger ved 40 kV, og kan skyldes høy usikkerhet ved måling av lave kV-verdier."
                    else:
                        om_40 = ""
                    doc.add_paragraph(u"<b>Konklusjon: </b>Ikke godkjent.<br></br><b>Resultat: </b>HVL ved {} kV lavere enn minstekravet.{}".format(hvl_kvp_ikke_ok_string, om_40))
#                    doc.add_paragraph(u"<b>Konklusjon: </b><font color=\"red\">Ikke godkjent</font>.<br></br><b>Resultat: </b>HVL ved {} kV lavere enn minstekravet.{}".format(hvl_kvp_ikke_ok_string, om_40))
                hvl_table.sort()
                hvl_table = hvl_table_head + hvl_table
                
                if not shortReport:        
                    doc.add_spacer()
                    doc.add_table(hvl_table, align="LEFT")
                    doc.add_spacer()
        
        
        # FLUORO TUBE kV PRECISION
        
        if is_ftqa:
            # exists EITHER small focus or large focus fluoro tube QA
            
            cur.execute("SELECT ft_id, ftpxd_fka_cm, ftdxd_fka_cm, ftpxd_fda_cm, ftpxd_fha_cm, ftdxd_fka_cm, ftkp_fka_cm, ftkp_comment FROM fluoro_tube WHERE \
                            FK_ft_qa = ?", (qa_id,))
            ftdata = cur.fetchall()[0]
            assert len(ftdata) == 8
            ft_id, ftpxd_fka_cm, ftdxd_fka_cm, ftpxd_fda_cm, ftpxd_fha_cm, ftdxd_fka_cm, ftkp_fka_cm, ftkp_comment = ftdata
            ftkp_comment = u"%s" % ftkp_comment
            
            cur.execute("SELECT COUNT(*) FROM fluoro_tube_kvprec_large WHERE FK_ftkpl_ft = ? AND ftkpl_meas_kv IS NOT NULL", (ft_id,))
            nFTQAL = cur.fetchall()[0][0]
            cur.execute("SELECT COUNT(*) FROM fluoro_tube_kvprec_small WHERE FK_ftkps_ft = ? AND ftkps_meas_kv IS NOT NULL", (ft_id,))
            nFTQAS = cur.fetchall()[0][0]
                
            if nFTQAL > 0 or nFTQAS > 0:
                if not shortReport:
                    doc.add_header(u"Kontroll av røntgenrøret ved fluoroskopi", H1)
            
                if not shortReport:
                    if ftkp_fka_cm:
                        doc.add_paragraph(u"Røntgenrøret blir målt ved hjelp av et Unfors Xi målekammer. Avstanden mellom fokuspunkt og målekammer \
                                        er %d cm, og det blir ikke benyttet tilleggsfiltrering. Strålefeltet blendes inn til målekammeret.\
                                        Slike målinger er utført der utstyret tillater manuelle innstillinger av kVp og mAs." % ftkp_fka_cm)
                    else:
                         doc.add_paragraph(u"Røntgenrøret blir målt ved hjelp av et Unfors Xi målekammer. Det blir ikke benyttet tilleggsfiltrering. \
                                             Strålefeltet blendes inn til målekammeret. \
                                             Slike målinger er utført der utstyret tillater manuelle innstillinger av kVp og mAs.")    
                if ftkp_comment != "None":
                    doc.add_paragraph(u"<b>Kommentar: </b>%s" % (ftkp_comment))
            if nFTQAL > 0:
        #         there are large focus tube QA measurements
        #                    
                cur.execute("SELECT ftkpl.ftkpl_set_kv, ftkpl.ftkpl_meas_kv, ftkpl.ftkpl_meas_hvl_mmal \
                            FROM fluoro_tube_kvprec_large ftkpl \
                            WHERE ftkpl.FK_ftkpl_ft = ? AND ftkpl.ftkpl_meas_kv IS NOT NULL", (ft_id,))
                            
                ftkpl_data = list(cur.fetchall())
                ftkpl_table = [["Set kVp", u"Målt kVp", u"kVp diff."]]
                FTKPL_ok = True        
                max_deviation = [0, 0]
                kv_list = [0]
        #
                for row in ftkpl_data:
                    if not row[1]: continue
                    if int(row[0]) in kv_list: continue            
                    kv_list.append(int(row[0]))
                    
                    deviation = (float(row[1]) / float(row[0]) - 1) * 100
                    if abs(deviation) > 10:
                        FTKPL_ok = False
                    if abs(deviation) > max_deviation[0]:
                        max_deviation = [abs(deviation), row[0]]
                    ftkpl_table.append(["%d" % row[0], "%.1f" % row[1], "%.1f %%" % deviation])
        #        
                doc.add_header(u"Nøyaktighet rørspenning (stort fokus)", H2)
                doc.add_paragraph(u"Nøyaktighet i rørspenning ble målt med stort fokus og økende kV-verdier.<br></br>\
                                <b>Krav: </b> Bør være innenfor 5 % av nominell verdi, skal være innenfor 10 %.<br></br>")
                if FTKPL_ok:
                    doc.add_paragraph(u"<b>Konklusjon: </b> Godkjent. <br></br> \
                        <b>Resultat: </b> Avviket mellom målt og nominell kV ligger innenfor grenseverdier \
                                        for alle målinger, se tabellen under.")
                else:
                   doc.add_paragraph(u"<b>Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>. <br></br> \
                            <b>Resultat: </b> Ved enkelte kV-verdier ble det målt et større avvik enn kravet tilsier. \
                            Det største avviket var på %.1f %% ved %d kVp. Se tabellen under." % (max_deviation[0], max_deviation[1]))
                
                if not shortReport:
                    doc.add_spacer()
                    doc.add_table(ftkpl_table, TABLE_WIDTH/1.1, align="LEFT")
                    doc.add_spacer()
                
            if nFTQAS > 0:
        #         there are small focus tube QA measurements
        #                    
                cur.execute("SELECT ftkps.ftkps_set_kv, ftkps.ftkps_meas_kv, ftkps.ftkps_meas_hvl_mmal \
                            FROM fluoro_tube_kvprec_small ftkps \
                            WHERE ftkps.FK_ftkps_ft = ? AND ftkps.ftkps_meas_kv IS NOT NULL", (ft_id,))
                            
                ftkps_data = list(cur.fetchall())
                ftkps_table = [["Set kVp", u"Målt kVp", u"kVp diff."]]
                FTKPS_ok = True        
                max_deviation = [0, 0]
                kv_list = [0]
                hvl_ok = True
        #
                for row in ftkps_data:
                    if not row[1]: continue
                    if int(row[0]) in kv_list: continue            
                    kv_list.append(int(row[0]))
                    if not row[2]: # if no HVL measurements
                        hvl_ok = False
                    
                    deviation = (float(row[1]) / float(row[0]) - 1) * 100
                    if abs(deviation) > 10:
                        FTKPS_ok = False
                    if abs(deviation) > max_deviation[0]:
                        max_deviation = [abs(deviation), row[0]]
                    ftkps_table.append(["%d" % row[0], "%.1f" % row[1], "%.1f %%" % deviation])
        #        
                doc.add_header(u"Nøyaktighet rørspenning (lite fokus)", H2)
                doc.add_paragraph(u"Nøyaktighet i rørspenning ble målt med lite fokus og økende kV-verdier.<br></br>\
                                <b>Krav: </b> Bør være innenfor 5 % av nominell verdi, skal være innenfor 10 %.<br></br>")
                if FTKPL_ok:
                    doc.add_paragraph(u"<b>Konklusjon: </b> Godkjent. <br></br> \
                        <b>Resultat: </b> Avviket mellom målt og nominell kV ligger innenfor grenseverdier \
                                        for alle målinger, se tabellen under.")
                else:
                   doc.add_paragraph(u"<b>Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>. <br></br> \
                            <b>Resultat: </b> Ved enkelte kV-verdier ble det målt et større avvik enn kravet tilsier. \
                            Det største avviket var på %.1f %% ved %d kVp. Se tabellen under." % (max_deviation[0], max_deviation[1]))
                if not shortReport:
                    doc.add_spacer()
                    doc.add_table(ftkps_table, TABLE_WIDTH/1.1, align="LEFT")
                    doc.add_spacer()
                
            if nFTQAL > 0:
                # reproducibility tube voltage
                # only with large focus, maybe implement small focus as well?
                
                ftkpl_nohvl = [row[:-1] for row in ftkpl_data]
                reproducibility_data = list(ftkpl_nohvl[:])
                # first sort by kV - 80 (since we want repeat ~80 kV measurements)
                reproducibility_data.sort(key=lambda tup: abs(tup[0]-80))
#                print str(reproducibility_data)
                for i in range(len(reproducibility_data)):
                    
                    if not reproducibility_data[i][0] == reproducibility_data[i+1][0]:
                            # i is the last ~80 kV same mAs measurement                    
                            break
                reproducibility_data = reproducibility_data[:i+1]
                if len(reproducibility_data) > 1:
                    # reproducibility_data only includes ~80 kV same mAs lines
                    
                    # Find average. np.average(array, axis to make average tuple of, 2nd item in list (average of array[2]))
                    
                    kvp_average = np.average(reproducibility_data, 0)[1]
                    kvp_largest_deviation = max([abs(x[1]/kvp_average-1)*100 for x in reproducibility_data]) 
            #        kvp_cov = stats.variation([x[2] for x in reproducibility_data])
                    if kvp_largest_deviation < 5:
                        kvp_deviation_ok = True
                    else:
                        kvp_deviation_ok = False
                    
                    number_of_measurements_in_reproducibility = numName[len(reproducibility_data)-1]
                    kvp_used_in_reproducibility = reproducibility_data[0][0]
                    
                    doc.add_header(u"Reproduserbarhet rørspenning", H2)
                    doc.add_paragraph(u"Reproduserbarheten i kVp og dose ble funnet ved å måle rørspenning ved %d kVp %s ganger. \
                                    Deretter beregnes målingens gjennomsnittsverdi og variasjonskoeffisienter. Stort fokus benyttes." % \
                                    (kvp_used_in_reproducibility, number_of_measurements_in_reproducibility))
                    
                    # should have a table here!!!!
                    
                    doc.add_paragraph(u"<b>Krav: </b> Største avvik fra målingens gjennomsnittsverdi skal være innenfor 5 % for kVp.<br></br>")
                    if kvp_deviation_ok:
                        doc.add_paragraph(u"<b>Konklusjon: </b> Godkjent.")
                    else:
                        doc.add_paragraph(u"<b>Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>.")
                    doc.add_paragraph(u"<b>Resultat: <br></br></b> Gjennomsnittlig rørspenning: %.1f kVp<br></br> \
                                        Største avvik fra gjennomsnittlig rørspenning: %.2f %% <br></br>"\
                                        % (kvp_average, kvp_largest_deviation))
        
                    if not shortReport:
                        doc.add_spacer()
        
                # HALF VALUE LAYER
                min_hvl_dict = {50:1.8, 60:2.2, 70:2.5, 80:2.9, 90:3.2, 110:3.9, 125:4.5, 150:4.5}
                
                hvl_data = {}
                hvl_data_ok = False
                
                for row in ftkpl_data:
                    kvp = row[0]
                    hvl = row[2]
                    
                    if not hvl: # no HVL but we need the table to display kVp accuracy
                        hvl = "N/A"
                    else:
                        hvl_data_ok = True
                        
                    if not kvp in hvl_data.keys():
                        hvl_data[kvp] = [hvl,]
                    else:
                        hvl_data[kvp].append(hvl)
                        
                # spacer from last paragraph
                if hvl_data_ok:
        
                    hvl_table_head = [["kVp", "HVL [mmAl]", "Minimum HVL [mmAl]", "Konklusjon"]]
                    hvl_table = []
                    
                    hvl_ok = True
                    hvl_calc_table = []
                    hvl_kvp_ikke_ok_list = []        
                    
                    for k, v in hvl_data.items():
                        if not "N/A" in v: # display kvp + HVL data
#                            print "OK data: ", k, v
                            if k in min_hvl_dict.keys():
                                min_hvl = min_hvl_dict[k]
                                average_hvl = np.average(v) #Lagt inn 18.11.2014. Skal vel være med her?
                                
                            else:
                                min_hvl = 0.0356 * k + 0.0381 # linear interpolation of minimum HVL values from IPEM / IEC
                                average_hvl = np.average(v)
            
                            if average_hvl < min_hvl:
                                hvl_ok = False
                                hvl_kvp_ikke_ok_list.append(str(k))
                            hvl_table.append([k, "%.1f" % average_hvl, "%.1f" % min_hvl, average_hvl>min_hvl and "Godkjent" or "<font color=\"red\">Ikke godkjent</font>"])
    #                    else: # display kVp and "N/A"
    #                        "NOT OK data: ", k, v
    #                        min_hvl = "N/A"
    #                        hvl_table.append([k, "N/A", "N/A", "N/A"])
                    
                    hvl_table.sort()
                    hvl_table = hvl_table_head + hvl_table
                    
                    doc.add_header(u"Halvverdilag", H2)
                    doc.add_paragraph(u"Halvverdilaget (HVL) er hvor mange mm Al som trengs for halvere intensiteten.")
                    doc.add_paragraph(u"<b>Krav: </b>HVL skal være høyere enn minimum HVL ved hver rørspenning.")
            
                    if not hvl_ok:
                        if len(hvl_kvp_ikke_ok_list) > 1:
                            hvl_kvp_ikke_ok_string = ", ".join(hvl_kvp_ikke_ok_list[:-1]) + " og %s" % hvl_kvp_ikke_ok_list[-1]
                        else:
                            hvl_kvp_ikke_ok_string = hvl_kvp_ikke_ok_list[0]
                    
                    if hvl_ok:
                        doc.add_paragraph(u"<b>Konklusjon: </b>Godkjent.")
                    else:
                        doc.add_paragraph(u"<b>Konklusjon: </b><font color=\"red\">Ikke godkjent</font>.<br></br><b>Resultat: </b>HVL ved %s kV lavere enn minstekravet." % hvl_kvp_ikke_ok_string)
                    
                    if not shortReport:
                        doc.add_spacer()
                        doc.add_table(hvl_table, align="LEFT")
                        doc.add_spacer()
        
        # DOSE TO PATIENT
        
        if is_ftdxd or is_ftpxd:
            # need to test for fluoro_test_pxdose and fluoro_test_dxdose
            cur.execute("SELECT ft.ftpxd_comment, ft.ftpxd_fka_cm, ft.ftpxd_fda_cm, ft.ftpxd_fha_cm, \
                                ft.ftpxd_internalfiltration, ft.ftdxd_focus, ft.ftdxd_fka_cm, ft.ftdxd_fda_cm, \
                                ft.ftdxd_internalfiltration, ft.ftdxd_comment FROM fluoro_tube ft WHERE ft_id = ?", (ft_id,))
                                
            result = cur.fetchall()[0]
            # should be one row
            # pxd = dose to patient
            # dxd = dose to detector
            pxd_comment = u"%s" % result[0]
            pxd_fka = float(result[1]) # focus - Unfors distance
            pxd_fda = float(result[2]) # focus - detector distance
            pxd_fha = float(result[3]) # least focus - skin distance
            pxd_if = result[4] # internal filtration
            dxd_focus = result[5]
            dxd_fka = result[6]
            dxd_fda = result[7]
            dxd_if = result[8]
            dxd_comment = u"%s" % result[9]
            
            #print ft_id, "ft_id"    
            
            if not shortReport:    
                doc.add_header(u"Doser gitt under fluoroskopi", H1)
                doc.add_paragraph(u"Det måles doser til pasient og detektor. Det legges 1 mm Cu plater mellom rør og målekammer for å måle dose til detektor. \
                        Ved å dekke detektor / bildeforsterker med Cu plater, for så å legge målekammer over, måler man huddose til pasient.")
            
            if is_ftdxd:
                 # dxdose
                cur.execute("SELECT ftdxd_program, ftdxd_pps, ftdxd_dosemode, ftdxd_fieldsize_cm, ftdxd_panel_kv, \
                                ftdxd_panel_ma, ftdxd_panel_ms, ftdxd_meas_mgy_s, ftdxd_corrected_mgy_s, ftdxd_is_exposure \
                                FROM fluoro_tube_dxdose WHERE FK_ftdxd_ft = ?", (ft_id,))
                result = cur.fetchall()
                
                
                max_dxdose = 0
                max_program = False
                max_pps = False
                
                table_print = [False]*6
                for row in result:
                    if row[0]: table_print[0] = True
                    if row[1]: table_print[1] = True
                    if row[2]: table_print[2] = True
                    if row[9]: table_print[3] = True
                    if row[3]: table_print[4] = True
                    if row[7]: table_print[5] = True
                                                    
                dxdose_table = ["Program", "Puls/sekund", "Dosemodus", "Eksponering?", u"Feltstørrelse", u"Målt dose"]
                
                # only print columns with numbers in them (at least one row)
                dxdose_table_print = zip(dxdose_table, table_print)
                dxdose_table = [[_x[0] for _x in dxdose_table_print if _x[1]]]
    
                # make detector dose table
                for row in result:
                    program = u"%s" % row[0]
                    pps = u"%s" % row[1]
                    dosemode = u"%s" % row[2]
                    fieldsize = row[3]
                    panel_kv = row[4]
                    panel_ma = row[5]
                    panel_ms = row[6]
                    meas_ugy_s = row[7]
                    corrected_ugy_s = row[8]
                    is_exposure = row[9]
                    
                    if corrected_ugy_s > max_dxdose and not is_exposure:
                        max_dxdose = corrected_ugy_s
                        max_program = program
                        max_pps = pps
                    
                    if fieldsize:
                        fieldsize_str = "%s cm" % fieldsize
                    else:
                        fieldsize_str = ""
                        
                        
                    dxdose_line = [program, pps, dosemode, is_exposure and "Ja" or "Nei", fieldsize_str, "%.2f µGy/s" % corrected_ugy_s]
                    dxdose_line_print = zip(dxdose_line, table_print)
                    dxdose_table.append([_x[0] for _x in dxdose_line_print if _x[1]])
                    
    #                dxdose_table.append([program, pps, dosemode, is_exposure and "Ja" or "Nei", fieldsize_str, "%.2f µGy/s" % corrected_ugy_s])        
            
                # done with table, move on to text
                
                doc.add_header(u"Dose til detektor", H2)
            
              # hjelpetekst
                if dxd_fka and dxd_fda and not shortReport:
                    dist_string = u" Målingen er utført med fokus-målekammer-avstand på %d cm, og fokus-detektor/bildeforsterker-avstand på %d cm. Den oppgitte måleverdien er korrigert for avstanden, og representerer dose ved detektor." % (dxd_fka, dxd_fda)
                else:
                    dist_string = ""
                
                if dxd_focus and not shortReport:
                    if dxd_focus.lower() == "gf":
                        focus_string = u" Det er brukt stort fokus."
                    elif dxd_focus.lower() == "ff":
                        focus_string = u" Det er brukt lite fokus."
                    else:
                        focus_string = ""
                else:
                    focus_string = ""
                
                doc.add_paragraph(u"<b>Krav: </b>Dose til detektor / bildeforsterker skal ikke overstige 1 µGy/s.%s%s" % (dist_string, focus_string))
                
                if max_dxdose < 1:
                    doc.add_paragraph(u"<b>Konklusjon: </b> Godkjent.")
                    doc.add_paragraph(u"<b>Resultat: </b> Ved alle testede programmer var dosen lavere enn grenseverdien, og høyeste verdi målt var %.1f µGy/s. Se tabellen under." % max_dxdose)
                else:
                    doc.add_paragraph(u"<b>Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>.")
                    doc.add_paragraph(u"<b>Resultat: </b> Dosen ved programmet <i>%s</i> ved pulshastighet %s ble målt til %.1f µGy/s, som er over grenseverdien. Se tabellen under." % (max_program, max_pps, max_dxdose))
                
                if dxd_comment != "None":
                    doc.add_paragraph(u"<b>Kommentar: </b>%s" % dxd_comment)
                
                if not shortReport:
                    doc.add_spacer()
                    doc.add_table(dxdose_table, align="LEFT")
                    doc.add_spacer()
                    
            if is_ftpxd:
            # Dose to pasient
                cur.execute("SELECT ftpxd_program, ftpxd_pps, ftpxd_dosemode, ftpxd_fieldsize_cm, ftpxd_panel_kv, \
                                ftpxd_panel_ma, ftpxd_panel_ms, ftpxd_meas_mgy_s, ftpxd_meas_pps, ftpxd_is_exposure \
                                FROM fluoro_tube_pxdose WHERE FK_ftpxd_ft = ?", (ft_id,))
                result = cur.fetchall()
                
    
                table_print = [False]*6
                for row in result:
                    if row[0]: table_print[0] = True
                    if row[1]: table_print[1] = True
                    if row[2]: table_print[2] = True
                    if row[9]: table_print[3] = True
                    if row[3]: table_print[4] = True
                    if row[7]: table_print[5] = True
                
                pxdose_table = ["Program", "Puls/sekund", "Dosemodus", "Eksponering?", u"Feltstørrelse", u"Målt dose"]
                pxdose_table_print = zip(pxdose_table, table_print)
                
                pxdose_table = [[_x[0] for _x in pxdose_table_print if _x[1]]]
                
                max_pxdose = 0
                max_pxdose_exposure = 0
                max_program = False
                max_pps = False
                max_program_exposure = False
                max_pps_exposure = False
                
                for row in result:
                    program = u"%s" % row[0]
                    pps = u"%s" % row[1]
                    dosemode = u"%s" % row[2]
                    fieldsize = row[3]
                    panel_kv = row[4]
                    panel_ma = row[5]
                    panel_ms = row[6]
                    meas_ugy_s = row[7]
                    meas_pps = row[8]
                    is_exposure = row[9]
                    
                    try:
                        corrected_ugy_s = meas_ugy_s * ( pxd_fka**2 / pxd_fha**2)
                    except:
                        print "Finner ikke fokus-hud-avstand eller fokus-kammer-avstand."
                    
                    if corrected_ugy_s > max_pxdose and not is_exposure:
                        max_pxdose = corrected_ugy_s
                        max_program = program
                        max_pps = pps
    
    #               Need some more information in Excel before this can be calculated correctly:
    #                   pulse per second (1/exposure time) for exposure mode
    #                   is _usually_ found in the lower (dose to patient) box
    
    #                if meas_ugy_s > max_pxdose_exposure and is_exposure:
    #                    max_pxdose_exposure = meas_ugy_s
    #                    max_program_exposure = program
    #                    max_pps_exposure = pps
                    
                  
                    
                    if fieldsize:
                        fieldsize_str = "%d cm" % fieldsize
                    else:
                        fieldsize_str = ""            
                    
                    pxdose_line = [program, pps, dosemode, is_exposure and "Ja" or "Nei", fieldsize_str, "%.2f µGy/s" % corrected_ugy_s]
                    pxdose_line_print = zip(pxdose_line, table_print)
                    pxdose_table.append([_x[0] for _x in pxdose_line_print if _x[1]])
                
                doc.add_header(u"Huddose til pasient", H2)
    
              # hjelpetekst
                if pxd_fka and pxd_fda and not shortReport:
                    dist_string = u" Målingen er utført med fokus-målekammer-avstand på %d cm, og fokus-detektor/bildeforsterker-avstand på %d cm. Den oppgitte måleverdien er korrigert for avstanden, og representerer dose ved minste fokus-hud-avstand." % (pxd_fka, pxd_fda)
                else:
                    dist_string = ""
                
                if dxd_focus and not shortReport:
                    if dxd_focus.lower() == "gf":
                        focus_string = u" Det er brukt stort fokus."
                    elif dxd_focus.lower() == "ff":
                        focus_string = u" Det er brukt lite fokus."
                    else:
                        focus_string = ""
                else:
                    focus_string = ""
                
                doc.add_paragraph(u"<b>Krav: </b>Huddose til pasient skal ikke overstige 1600 µGy/s ved fluoroskopi, og 2 mGy/bilde ved eksponering. %s%s" % (dist_string, focus_string))
                
                if max_pxdose < 1600:
                    doc.add_paragraph(u"<b>Konklusjon: </b> Godkjent.")
                    doc.add_paragraph(u"<b>Resultat: </b> Ved alle testede programmer var dosen lavere enn grenseverdien, og høyeste verdi målt var %.1f µGy/s. Se tabellen under." % max_pxdose)
                else:
                    doc.add_paragraph(u"<b>Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>.")
                    doc.add_paragraph(u"<b>Resultat: </b> Dosen ved programmet <i>%s</i> ved pulshastighet %s ble målt til %.1f µGy/s, som er over grenseverdien. Se tabellen under." % (max_program, max_pps, max_pxdose))
                
                if pxd_comment != "None":
                    doc.add_paragraph(u"<b>Kommentar: </b>%s" % (pxd_comment))
                
                if not shortReport:
                    doc.add_spacer()
                    doc.add_table(pxdose_table, align="LEFT")
                    doc.add_spacer()
                
        # Leeds IQ
        if is_fiq:
            cur.execute("SELECT fiq.fiq_id, fiq.fiq_lab_monitor, fiq.fiq_control_monitor, fiq.fiq_comment FROM fluoro_iq fiq WHERE fiq.FK_fiq_qa = ?", (qa_id,))
            result = cur.fetchall()[0]
            fiq_id = result[0]
            fiq_lab_monitor = result[1]
            fiq_control_monitor = result[2]
            
            #print fiq_id, "IQ"
            fiq_comment = u"%s" % result[3]
    
            cur.execute("SELECT * from fluoro_iq_each WHERE FK_fiqe_fiq = ?", (fiq_id,))
            fiq_data = cur.fetchall()
        
            fiq_table_lab = []
            fiq_table_control = []
        
            # We check if these values are present in any of the measurements
            # If they aren't, the result table should be shortened down
            
            is_control_live_contrast = False
            is_control_live_lpmm = False
            is_control_lih_contrast = False
            is_control_lih_lpmm = False
            is_lab_live_contrast = False
            is_lab_live_lpmm = False
            is_lab_lih_contrast = False
            is_lab_lih_lpmm = False
            is_pulsespeed = False
            is_dosemode = False
            is_fieldsize = False
            
            lowest_resolution = 100
            highest_contrast = 0
            
            bad_contrast_program = False
            bad_resolution_program = False    
            
            for row in fiq_data:
                _id = row[0]
                _pulsespeed = row[1]
                _dosemode = row[2]
                _fieldsize = row[3]
                _control_live_contrast = row[4]
                _control_live_lpmm = row[5]
                _control_lih_contrast = row[6]
                _control_lih_lpmm = row[7]
                _lab_live_contrast = row[8]
                _lab_live_lpmm = row[9]
                _lab_lih_contrast = row[10]
                _lab_lih_lpmm = row[11]
                        
                # Any of the flags should be True if at least one measurement is present
                if _pulsespeed: is_pulsespeed = True
                if _dosemode: is_dosemode = True
                if _fieldsize: is_fieldsize = True        
                if _control_live_contrast: is_control_live_contrast = True
                if _control_live_lpmm: is_control_live_lpmm = True
                if _control_lih_contrast: is_control_lih_contrast = True
                if _control_lih_lpmm: is_control_lih_lpmm = True
                if _lab_live_contrast: is_lab_live_contrast = True
                if _lab_live_lpmm: is_lab_live_lpmm = True
                if _lab_lih_contrast: is_lab_lih_contrast = True
                if _lab_lih_lpmm: is_lab_lih_lpmm = True
                
                if _lab_live_lpmm and _lab_live_lpmm < lowest_resolution:
                    lowest_resolution = _lab_live_lpmm
                    
                if _lab_lih_lpmm and _lab_lih_lpmm < lowest_resolution:
                    lowest_resolution = _lab_lih_lpmm
                    bad_resolution_program = _dosemode
               
                if _control_lih_lpmm and _control_lih_lpmm < lowest_resolution:
                    lowest_resolution = _control_lih_lpmm
                    bad_resolution_program = _dosemode
                        
                if _control_live_lpmm and _control_live_lpmm < lowest_resolution:
                    lowest_resolution = _control_live_lpmm
                    bad_resolution_program = _dosemode
        
                if _lab_live_contrast and _lab_live_contrast > highest_contrast:
                    highest_contrast = _lab_live_contrast
                    bad_contrast_program = _dosemode
        
                if _lab_lih_contrast and _lab_lih_contrast > highest_contrast:
                    highest_contrast = _lab_lih_contrast
                    bad_contrast_program = _dosemode
        
                if _control_live_contrast and _control_live_contrast > highest_contrast:
                    highest_contrast = _control_live_contrast
                    bad_contrast_program = _dosemode
        
                if _control_lih_contrast and _control_lih_contrast > highest_contrast:
                    highest_contrast = _control_lih_contrast
                    bad_contrast_program = _dosemode
                    
            is_control = is_control_live_contrast + is_control_live_lpmm + is_control_lih_contrast + is_control_lih_lpmm
            is_lab = is_lab_live_contrast + is_lab_live_lpmm + is_lab_lih_contrast + is_lab_lih_lpmm
                    
            if not shortReport:
                doc.add_header(u"Bildekvalitet målt med Leeds-fantomene", H1)
                    
            if not shortReport:
                doc.add_paragraph(u"Leeds fantom TOR CDR ble eksponert forskjellige protokoller, pulshastigheter og dosemodi. Det ble brukt \
                        1 mm Cu som filtrering i strålefeltet. Det kan måles både ved live fluoroskopi og last image hold (LiH).")
            
            if fiq_comment != "None":
                doc.add_paragraph(u"<b>Kommentar: </b>%s" % fiq_comment)
            
            if modality in ("DX", "CR", "DR"):
                res_limit = 2.4
                contrast_limit = 0.028
            elif modality in ("RF", "XA"):
                res_limit = 1.0
                contrast_limit = 0.04
            else:
                print("Unknown modality type: {}\nUsing DX limits for image quality.".format(modality))
                contrast_limit = 0.028
                res_limit = 2.4
            
            contrast_ok = highest_contrast <= contrast_limit and True or False
            resolution_ok = lowest_resolution >= res_limit and True or False    
            
            #doc.add_spacer()
            doc.add_header(u"Høykontrast-oppløsning", H2)
            doc.add_paragraph(u"<b>Krav: </b>%.1f linjepar per mm (lp/mm) skal være detekterbare. For å avlese linjepar kan det brukes \
                        forstørrelse, da det er systemets evne til å skille linjeparene som er viktig." % res_limit)
            if resolution_ok:
                doc.add_paragraph(u"<b> Konklusjon: </b>Godkjent.")
                doc.add_paragraph(u"<b> Resultat: </b> Laveste målte oppløsning var %.1f lp/mm, som er godkjent. Se tabellen under. Resultatene er \
                referanse for framtidige målinger." % float(lowest_resolution))
            else:
                doc.add_paragraph(u"<b> Konklusjon: </b> <font color=\"red\">Ikke godkjent</font>.")
                doc.add_paragraph(u"<b> Resultat: </b> Laveste målte oppløsning var %.1f lp/mm ved programmet <i>%s</i>, som er under grensen. Se tabellen under. Resultatene er \
                referanse for framtidige målinger." % (float(lowest_resolution), bad_resolution_program))
                
            doc.add_header(u"Lavkontrast", H2)
            doc.add_paragraph(u"<b>Krav: </b>Minst %.1f %% kontrastforskjell skal være synlig. Det målte tallet skal altså være <i>lavere</i> enn %.1f %%." % (contrast_limit*100, contrast_limit*100))
            if contrast_ok:
                doc.add_paragraph(u"<b>Konklusjon: </b>Godkjent.")
                doc.add_paragraph(u"<b>Resultat: </b>Den høyeste synlige kontrastforskjellen var %.1f %%, som er bedre enn grenseverdien. \
                Se tabellen under. Resultatene er referanse for framtidige målinger." % (highest_contrast*100))
            else:
                doc.add_paragraph(u"<b>Konklusjon: </b><font color=\"red\">Ikke godkjent</font>.")
                doc.add_paragraph(u"<b>Resultat: </b>Den høyeste synlige kontrastforskjellen var %.1f %% ved programmet <i>%s</i>, som er dårligere enn grenseverdien. \
                Se tabellen under. Resultatene er referanse for framtidige målinger." % (highest_contrast*100, bad_contrast_program))    
                
            if is_control:
                    
                pls = is_pulsespeed and ["Pulshastighet"] or []
                dos = is_dosemode and ["Dosemodus"] or []
                flt = is_fieldsize and ["Feltstørrelse"] or []
                clvc = is_control_live_contrast and ["Kontrast\n live"] or []
                clvl = is_control_live_lpmm and [u"Oppløsning\n live"] or []
                clic = is_control_lih_contrast and ["Kontrast\n LiH"] or []
                clil = is_control_lih_lpmm and [u"Oppløsning\n LiH"] or []
            
                fiq_table_control_header = [pls + dos + flt + clvl + clil + clvc + clic]    
        
                for row in fiq_data:
                    _id = row[0]
                    _pulsespeed = is_pulsespeed and [u"%s" % row[1]] or []
                    _dosemode = is_dosemode and [u"%s" % row[2]]  or []
                    _fieldsize = is_fieldsize and [u"%s" % row[3]] or []
                    
                    if row[4] and is_control_live_contrast:
                        _control_live_contrast = ["%.1f %%" % (100*float(row[4]))]
                    elif is_control_live_contrast: 
                        _control_live_contrast = [""]
                    else:
                        _control_live_contrast = []
                        
                    if row[5] and is_control_live_lpmm:
                        _control_live_lpmm = ["%.1f lp/mm" % row[5]]
                    elif is_control_live_lpmm:
                        _control_live_lpmm = [""]
                    else:
                        _control_live_lpmm = []
                        
                    if row[6] and is_control_lih_contrast:
                        _control_lih_contrast = ["%.1f %%" % (100*float(row[6]))]
                    elif is_control_lih_contrast:
                        _control_lih_contrast = [""]
                    else:
                        _control_lih_contrast = []
                        
                    if row[7] and is_control_lih_lpmm:
                        _control_lih_lpmm = ["%.1f lp/mm" % row[7]]
                    elif is_control_lih_lpmm:
                        _control_lih_lpmm = [""]
                    else:
                        _control_lih_lpmm = []
                                
                    fiq_table_control.append(_pulsespeed + _dosemode + _fieldsize + _control_live_lpmm +
                                             _control_lih_lpmm + _control_live_contrast + _control_lih_contrast)
                                             
                fiq_table_control = fiq_table_control_header + fiq_table_control
                if not shortReport:     
                    doc.add_spacer()
                    doc.add_paragraph(u"Tabell for målinger utført på kontrollrom:")
                    doc.add_table(fiq_table_control, align="LEFT")        
                    doc.add_spacer()
            
            if is_lab:
                pls = is_pulsespeed and ["Pulshastighet"] or []
                dos = is_dosemode and ["Dosemodus"] or []
                flt = is_fieldsize and ["Feltstørrelse"] or []
                llvc = is_lab_live_contrast and ["Kontrast\n  live"] or []
                llvl = is_lab_live_lpmm and [u"Oppløsning\n  live"] or []
                llic = is_lab_lih_contrast and ["Kontrast\n  LiH"] or []
                llil = is_lab_lih_lpmm and [u"Oppløsning\n  LiH"] or []
                
                fiq_table_lab_header = [pls + dos + flt + llvl + llil + llvc + llic]
            
                for row in fiq_data:
                    _id = row[0]
                    # _pulsespeed, _dosemode and _fieldsize may be UTF-8
                    _pulsespeed = is_pulsespeed and [u"%s" % row[1]] or []
                    _dosemode = is_dosemode and [u"%s" % row[2]]  or []
                    _fieldsize = is_fieldsize and [u"%s" % row[3]] or []
        
                    if row[8] and is_lab_live_contrast:
                        _lab_live_contrast =  ["%.1f %%" % (100.*float(row[8]))]
                    elif is_lab_live_contrast:
                        _lab_live_contrast = [""]
                    else:
                        _lab_live_contrast = []
                        
                    if row[9] and is_lab_live_lpmm:
                        _lab_live_lpmm = ["%.1f lp/mm" % row[9]]
                    elif is_lab_live_lpmm:
                        _lab_live_lpmm = [""]
                    else:
                        _lab_live_lpmm = []
                        
                    if row[10] and is_lab_lih_contrast:    
                        _lab_lih_contrast = ["%.1f %%" % (100.*float(row[10]))]
                    elif is_lab_lih_contrast:
                        _lab_lih_contrast = [""]
                    else:
                        _lab_lih_contrast = []
                        
                    if row[11] and is_lab_lih_lpmm:
                        _lab_lih_lpmm = ["%.1f lp/mm" % row[11]]
                    elif is_lab_lih_lpmm:
                        _lab_lih_lpmm = [""]
                    else:
                        _lab_lih_lpmm = []
                    
                    fiq_table_lab.append(_pulsespeed + _dosemode + _fieldsize  + _lab_live_lpmm + 
                                _lab_lih_lpmm + _lab_live_contrast + _lab_lih_contrast)
            
                fiq_table_lab = fiq_table_lab_header + fiq_table_lab
                
                if smallTableOnLeeds:
                    leedsTableStyle = [('FONTSIZE', (0,0), (-1,-1), 8)]
                else:
                    leedsTableStyle = []
                
                if not shortReport:
                    doc.add_spacer()
                    doc.add_paragraph(u"Tabell for målinger utført på modalitet:")
                    doc.add_table(fiq_table_lab, align="LEFT", extra_style = leedsTableStyle)
                    doc.add_spacer()        
                    
        if is_dap:
            cur.execute("SELECT de.dap_each_kv, de.dap_each_modality_mean, de.dap_each_meter_mean_kv, \
                        (de.dap_each_modality_mean / de.dap_each_meter_mean_kv) - 1 \
                        FROM DAP_each de, DAP d \
                        WHERE de.FK_de_d = d.dap_id AND d.FK_dap_qa = ? \
                        ORDER BY de.dap_each_kv", (qa_id,))
                        
            dap_data = cur.fetchall()
            dap_table = [["kVp     ", "Modalitet\n[mGycm2]", "Ref. DAP-meter\n[mGycm2]", "Avvik"]]
            
            for row in dap_data:
                deviation = float(row[3]) * 100
                dap_table.append(["%d" % row[0], "%.1f" % row[1], "%.1f" % row[2], "%.1f %%" % deviation])
        
            cur.execute("SELECT dap_comment FROM dap WHERE FK_dap_qa = ?", (qa_id,))
                
            dap_comment = u"%s" % cur.fetchall()[0][0]
                
            if modality in ("DX", "CR"):
                dap_limit = 25
            elif modality in ("RF", "XA"):
                dap_limit = 35
            else:
                print "DAP limit not found for modality %s" % modality
            
            dap_max = 0
            for line in dap_data:
                if abs(line[3]) > dap_max:
                    dap_max = abs(line[3])
            
            if not shortReport:    
                doc.add_header(u"Kontroll av DAP-meter", H1)
            else:
                doc.add_header(u"Kontroll av DAP-meter", H2)
        
            doc.add_paragraph(u"Utstyrets dose-areal-produkt (DAP)-meter ble kontrollert ved bruk av et referanse DAP-meter.")
            doc.add_paragraph(u"<b>Krav: </b> Avviket mellom DAP-verdien til utstyret og målt verdi på referanse DAP-meteret skal være mindre enn %d %%." % dap_limit)
        
            if dap_max < dap_limit:
                doc.add_paragraph(u"<b>Konklusjon: </b>Godkjent.")
            else:
                doc.add_paragraph(u"<b>Konklusjon: </b><font color=\"red\">Ikke godkjent</font>.")
                
            doc.add_paragraph(u"<b>Resultat: </b> Det største avviket ble målt til %.1f %%." % (100*dap_max))
            
            if dap_comment != "None":
                doc.add_paragraph(u"<b>Kommentar</b>: %s" % dap_comment)
        
            if not shortReport:    
                doc.add_spacer()
                doc.add_table(dap_table, TABLE_WIDTH/1.1, align="LEFT")
                doc.add_spacer()
        
        if is_aek:
            
            cur.execute("SELECT aek.aek_position, aeke.aeke_mmcu, aeke.aeke_chamber_num, aeke.aeke_calc_dose_1, aeke.aeke_calc_dose_2 \
                        FROM aek_each aeke, aek \
                        WHERE aek.aek_id = aeke.FK_aeke_aek AND aek.FK_aek_qa = ? and aeke_calc_dose_1 IS NOT NULL and aeke_calc_dose_2 IS NOT NULL", (qa_id,))
            aek_data = cur.fetchall()
            aek_reproducibility_max_table = 0
            aek_reproducibility_max_wall = 0
            
            nTable = 0
            nWall = 0        
            
            for row in aek_data:
                position = row[0]
                mmcu = row[1]
                chamber_num = row[2]
                calc_dose_1 = float(row[3])
                calc_dose_2 = float(row[4])
                
                if position == "table":
                    nTable += 1
                    aek_reproducibility_mean = (calc_dose_1 + calc_dose_2) / 2.
                    aek_reproducibility_deviation = max(abs(calc_dose_1 - aek_reproducibility_mean)/aek_reproducibility_mean,
                                                    abs(calc_dose_2 - aek_reproducibility_mean)/aek_reproducibility_mean)
                    
                    aek_reproducibility_max_table = max(aek_reproducibility_max_table, aek_reproducibility_deviation)
                
                elif position == "wall":
                    nWall += 1
                    aek_reproducibility_mean = (calc_dose_1 + calc_dose_2) / 2.
                    aek_reproducibility_deviation = max(abs(calc_dose_1 - aek_reproducibility_mean)/aek_reproducibility_mean,
                                                    abs(calc_dose_2 - aek_reproducibility_mean)/aek_reproducibility_mean)
                    
                    aek_reproducibility_max_wall = max(aek_reproducibility_deviation, aek_reproducibility_max_wall)
        
            aek_wall_comment = aek_table_comment = False    
            cur.execute("SELECT aek_position, aek_comment FROM aek WHERE FK_aek_qa = ?", (qa_id,))
            result = cur.fetchall()
            for row in result:
                if row[0] == "wall":
                    aek_wall_comment = u"%s" % row[1]
                if row[0] == "table":
                    aek_table_comment = u"%s" % row[1]
                    
            if not shortReport:    
                doc.add_header(u"Kontroll av automatisk eksponeringskammer", H1)
            if not shortReport:
                doc.add_paragraph(u"Ved test av automatisk eksponeringskontroll i røntgenbord eller vekkbucky, testes samsvar mellom kamre, \
                                    reproduserbarhet for hvert kammer, samt tykkelseskompensasjon. Typisk eksponeringsprotokoll brukes.<br></br><br></br>")
                doc.add_paragraph(u"Målingen utføres ved å først finne sammenhengen mellom innstilt mAs og egenmålt dose ved lite felt. Denne målingen utføres ved både lite (1 mm) og mye (2 mm) Cu-filtrering. Så blendes feltet ut og automatikk slåes på, og målinger utføres med kun Cu-filtrering i feltet. mAs-verdien noteres og regnes om til dose vha. sammenhengen som er funnet.")
        #    doc.add_paragraph(u"Sammenhengen mellom mAs og dose ble målt ved %s, og er X = %d µGy/mAs * Y mAs + %d µGy ved 1 mm og X = %d µGy/mAs * Y mAs + %d µGy ved 2 mm." % (measured_microgy, slope_1mm, offset_1mm, slope_2mm, offset_2mm))
            if nTable > 0:
                is_aek_table = True
            else:
                is_aek_table = False
                
            if nWall > 0:
                is_aek_wall = True
            else:
                is_aek_wall = False
            
            if is_aek_table:
                if aek_reproducibility_max_table > 0.2:
                    reproducibility_table_ok = False
                else:
                    reproducibility_table_ok = True
                
            if is_aek_wall:
                if aek_reproducibility_max_wall > 0.2:
                    reproducibility_wall_ok = False
                else:
                    reproducibility_wall_ok = True
        
            if not shortReport:
                doc.add_spacer()
            doc.add_header(u"Reproduserbarhet", H2)
        
            doc.add_paragraph(u"<b>Krav: </b>Alle målinger med ett kammer skal ligge innenfor 20 % fra gjennomsnittet i en måleserie.")
            
            if is_aek_table:
                if reproducibility_table_ok:
                    doc.add_paragraph(u"<b>Konklusjon for bordbucky: </b> Godkjent")
                    doc.add_paragraph(u"<b>Resultat for bordbucky: </b>Alle kamrene var innenfor %.1f %% fra gjennomsnittet i en måleserie." % (aek_reproducibility_max_table*100))
                else:
                    doc.add_paragraph(u"<b>Konklusjon for bordbucky: </b> <font color=\"red\">Ikke godkjent</font>.")
                    doc.add_paragraph(u"<b>Resultat for bordbucky: </b>Kamrene var utenfor grenseverdiene, med et maksimalt utslag på %.1f %% fra gjennomsnittet i en måleserie." % (aek_reproducibility_max_table*100))
                if aek_table_comment != "None":
                    doc.add_paragraph(u"<b>Kommentar for bordbucky: </b>%s" % (aek_table_comment))
            
            if is_aek_wall:
                if reproducibility_wall_ok:
                    doc.add_paragraph(u"<b>Konklusjon for veggbucky: </b> Godkjent")
                    doc.add_paragraph(u"<b>Resultat for veggbucky: </b>Alle kamrene var innenfor %.1f %% fra gjennomsnittet i en måleserie." % (aek_reproducibility_max_wall*100))
                else:
                    doc.add_paragraph(u"<b>Konklusjon for veggbucky: </b> <font color=\"red\">Ikke godkjent</font>.")
                    doc.add_paragraph(u"<b>Resultat for veggbucky: </b>Kamrene var utenfor grenseverdiene, med et maksimalt utslag på %.1f %% fra gjennomsnittet i en måleserie." % (aek_reproducibility_max_wall*100))
                if aek_wall_comment != "None":
                    doc.add_paragraph(u"<b>Kommentar for veggbucky: </b>%s" % (aek_wall_comment))
        
            
            if nTable > 2 or nWall > 2: # 2 measurements for each position, 1 mmCu and 2mmCu
                if not shortReport:
                    doc.add_spacer()
                    
                doc.add_header(u"Samsvar mellom kamre", H2)
                aek_consistency_table = [0]*3
                aek_consistency_wall = [0]*3
        
                for row in aek_data:    
                    if is_aek_table:        
                        if row[0] == "table" and row[1] == 1:
                            aek_consistency_table[row[2]-1] = np.mean((row[3], row[4]))
            
                    if is_aek_wall:      
                        if row[0] == "wall" and row[1] == 1:
                            aek_consistency_wall[row[2]-1] = np.mean((row[3], row[4]))
            
                aek_table_mean = np.mean(aek_consistency_table)
                aek_wall_mean = np.mean(aek_consistency_wall)
        
                aek_max_consistency_table = max([0] + [ abs(x / aek_table_mean - 1) for x in aek_consistency_table if not x == aek_table_mean])
                aek_max_consistency_wall = max([0] + [ abs(x / aek_wall_mean - 1) for x in aek_consistency_wall if not x == aek_wall_mean])
            
                if is_aek_table:
                    if aek_max_consistency_table > 0.2:
                        consistency_table_ok = False
                    else:
                        consistency_table_ok = True
                
                if is_aek_wall:
                    if aek_max_consistency_wall > 0.2:
                        consistency_wall_ok = False
                    else:
                        consistency_wall_ok = True
                
                doc.add_paragraph(u"<b>Krav: </b>Gjennomsnittet for ett kammer skal ligge innen 20 % fra gjennomsnittet over alle kamre.")
                
                if is_aek_table:
                    if consistency_table_ok:
                        doc.add_paragraph(u"<b>Konklusjon for bordbucky: </b> Godkjent.")
                        doc.add_paragraph(u"<b>Resultat for bordbucky: </b>Hvert kammer var innenfor %.1f %% fra gjennomsnittet over alle kamrene." % (aek_max_consistency_table*100))
                    else:
                        doc.add_paragraph(u"<b>Konklusjon for bordbucky: </b> <font color=\"red\">Ikke godkjent</font>.")
                        doc.add_paragraph(u"<b>Resultat for bordbucky: </b>Kamrene var opptil %.1f %% unna gjennomsnittet over alle kamrene." % (aek_max_consistency_table*100))
                        
                if is_aek_wall:
                    if consistency_wall_ok:
                        doc.add_paragraph(u"<b>Konklusjon for veggbucky: </b> Godkjent.")
                        doc.add_paragraph(u"<b>Resultat for veggbcky: </b>Hvert kammer var innenfor %.2f %% fra gjennomsnittet over alle kamrene." % (aek_max_consistency_wall*100))
                    else:
                        doc.add_paragraph(u"<b>Konklusjon for veggbucky: </b> <font color=\"red\">Ikke godkjent</font>.")
                        doc.add_paragraph(u"<b>Resultat for veggbucky: </b>Kamrene var opptil %.1f %% unna gjennomsnittet over alle kamrene." % (aek_max_consistency_wall*100))
        
            if not shortReport:
                doc.add_spacer()
            doc.add_header(u"Tykkelseskompenasjon", H2)
            
            aek_thickness_table_1mm = [0]*3
            aek_thickness_table_2mm = [0]*3
            aek_thickness_wall_1mm = [0]*3
            aek_thickness_wall_2mm = [0]*3
        
            for row in aek_data:
                if row[0] == "table" and row[1] == 1:
                    aek_thickness_table_1mm[row[2]-1] = np.mean([row[3], row[4]])
                if row[0] == "table" and row[1] == 2:
                    aek_thickness_table_2mm[row[2]-1] = np.mean([row[3], row[4]])
                if row[0] == "wall" and row[1] == 1:
                    aek_thickness_wall_1mm[row[2]-1] = np.mean([row[3], row[4]])
                if row[0] == "wall" and row[1] == 2:
                    aek_thickness_wall_2mm[row[2]-1] = np.mean([row[3], row[4]])
        
            if is_aek_table:
                aek_thickness_table = [0]*3
                for i in range(3):
                    if aek_thickness_table_1mm[i] and aek_thickness_table_2mm[i]:
                        aek_thickness_table[i] = aek_thickness_table_2mm[i] / aek_thickness_table_1mm[i] - 1
                max_aek_thickness_table = max( [ abs(x) for x in aek_thickness_table ] )
            
            if is_aek_wall:    
                aek_thickness_wall = [0]*3
                for i in range(3):
                    if aek_thickness_wall_1mm[i] and aek_thickness_wall_2mm[i]:
                        aek_thickness_wall[i] = aek_thickness_wall_2mm[i] / aek_thickness_wall_1mm[i] - 1
                max_aek_thickness_wall = max( [ abs(x) for x in aek_thickness_wall ] )
        
        
            if is_aek_table:
                if max_aek_thickness_table > 0.4:
                    aek_thickness_table_ok = False
                else:
                    aek_thickness_table_ok = True
            
            if is_aek_wall:
                if max_aek_thickness_wall > 0.4:
                    aek_thickness_wall_ok = False
                else:
                    aek_thickness_wall_ok = True
                
            doc.add_paragraph(u"<b>Krav: </b>Ved bruk av ett av kamrene bør gjennomsnittet for hver tykkelse (1 mm og 2 mm Cu) ligge innenfor 20 % fra \
                    gjennomsnittet over alle målingene ved dette kammeret, og skal ligge innenfor 40 %.")
            
            if is_aek_table:
                if aek_thickness_table_ok:
                    doc.add_paragraph(u"<b>Konklusjon for bordbucky: </b>Godkjent.")
                    doc.add_paragraph(u"<b>Resultat for bordbucky: </b>Forskjellen mellom dosen ved 1 mm og 2 mm Cu var %.1f %%, som er innenfor grenseverdien." % (max_aek_thickness_table*100))
                    
                else:
                    doc.add_paragraph(u"<b>Konklusjon for bordbucky: </b><font color=\"red\">Ikke godkjent</font>.")
                    doc.add_paragraph(u"<b>Resultat for bordbucky: </b>Forskjellen mellom dosen ved 1 mm og 2 mm Cu var %.1f %%, som er utenfor grenseverdien" % (max_aek_thickness_table*100))
            
            if is_aek_wall:
                if aek_thickness_wall_ok:
                    doc.add_paragraph(u"<b>Konklusjon for veggbucky: </b>Godkjent.")
                    doc.add_paragraph(u"<b>Resultat for veggbucky: </b>Forskjellen mellom dosen ved 1 mm og 2 mm Cu var %.1f %%, som er innenfor grenseverdien." % (max_aek_thickness_wall*100))
                    
                else:
                    doc.add_paragraph(u"<b>Konklusjon for veggbucky: </b><font color=\"red\">Ikke godkjent</font>.")
                    doc.add_paragraph(u"<b>Resultat for veggbucky: </b>Forskjellen mellom dosen ved 1 mm og 2 mm Cu var %.1f %%, som er utenfor grenseverdien" % (max_aek_thickness_wall*100))
            
        #    if is_aek_table and not is_aek_wall:
        #        doc.add_paragraph(u"<b>Resultat: </b>For bord-buckyen var det største avviket %.1f %%." % (max_aek_thickness_table*100))
        #    if is_aek_wall and not is_aek_table:
        #        doc.add_paragraph(u"<b>Resultat: </b>For vegg-buckyen var det største avviket %.1f %%." % (max_aek_thickness_wall*100))
        #    if is_aek_wall and is_aek_table:
        #        doc.add_paragraph(u"<b>Resultat: </b>For bord-buckyen var det største avviket %.1f %%, mens for vegg-buckyen \
        #            var det største avviket %.1f %%." % (max_aek_thickness_table*100, max_aek_thickness_wall*100))
        
        
            # tabell over dose i kamre vs mmCu
            if is_aek_wall and is_aek_table:
                aek_table = [["Kammer", u"Dose [µGy]\n1mm  2mm ", "", "Avvik", "              ", "Kammer", u"Dose [µGy]\n1mm  2mm ", "", "Avvik"]]
                aek_tableStyle = [('SPAN', (1,0), (2,0)),
                                  ('SPAN', (6,0), (7,0)),
                                  ('SPAN', (0,-1),(3,-1)),
                                  ('SPAN', (5,-1), (-1,-1)),
                                  ('ALIGN', (0,-1), (-1,-1), "CENTER"),
                                  ('BACKGROUND', (4, 0), (4,-1), colors.white),
                                    ('GRID', (4,0), (4,-1),2, colors.white),
                                  ('LINEABOVE', (0,-1), (3,-1), 1, colors.black),
                                ('LINEABOVE', (5,-1), (-1,-1), 1, colors.black)]
                chamberName = [u"Venstre", u"Høyre", "Midtre"]
                for i in range(3):
                    t1 = aek_thickness_table_1mm[i]
                    t2 = aek_thickness_table_2mm[i]
                    w1 = aek_thickness_wall_1mm[i]
                    w2 = aek_thickness_wall_2mm[i]        
                    
                    aek_table.append([chamberName[i], "%.1f" % t1, "%.1f" % t2, "%.1f %%" % (100*(t2/t1-1)) , "",
                                      chamberName[i], "%.1f" % w1, "%.1f" % w2, "%.1f %%" % (100*(w2/w1-1))])
            
                aek_table.append(["Bord-bucky", "", "", "", "", "Vegg-bucky", "", "", ""])
            elif is_aek_table and not is_aek_wall:
                aek_table = [["Kammer", u"Dose [µGy]\n1mm  2mm ", "", "Avvik", ]]
                aek_tableStyle = [('SPAN', (1,0), (2,0)),
                                  ('SPAN', (0,-1),(3,-1)),
                                  ('ALIGN', (0,-1), (-1,-1), "CENTER"),
                                  ('BACKGROUND', (4, 0), (4,-1), colors.white),
                                    ('GRID', (4,0), (4,-1),2, colors.white),
                                  ('LINEABOVE', (0,-1), (3,-1), 1, colors.black),
                                ('LINEABOVE', (3,-1), (-1,-1), 1, colors.black)]
                chamberName = [u"Venstre", u"Høyre", "Midtre"]
                for i in range(3):
                    t1 = aek_thickness_table_1mm[i]
                    t2 = aek_thickness_table_2mm[i]    
                    if t1 and t2:
                        aek_table.append([chamberName[i], "%.1f" % t1, "%.1f" % t2, "%.1f %%" % (100*(t2/t1-1))])
            
                aek_table.append(["Bord-bucky", "", "", "",])
                
            elif not is_aek_table and is_aek_wall:
                aek_table = [["Kammer", u"Dose [µGy]\n1mm  2mm ", "", "Avvik", ]]
                aek_tableStyle = [('SPAN', (1,0), (2,0)),
                                  ('SPAN', (0,-1),(3,-1)),
                                  ('ALIGN', (0,-1), (-1,-1), "CENTER"),
                                  ('BACKGROUND', (4, 0), (4,-1), colors.white),
                                    ('GRID', (4,0), (4,-1),2, colors.white),
                                  ('LINEABOVE', (0,-1), (3,-1), 1, colors.black),
                                ('LINEABOVE', (3,-1), (-1,-1), 1, colors.black)]
                chamberName = [u"Venstre", u"Høyre", "Midtre"]
                for i in range(3):
                    t1 = aek_thickness_wall_1mm[i]
                    t2 = aek_thickness_wall_2mm[i]    
                    
                    aek_table.append([chamberName[i], "%.1f" % t1, "%.1f" % t2, "%.1f %%" % (100*(t2/t1-1))])
            
                aek_table.append(["vegg-bucky", "", "", "",])
            
            if not is_aek_table and not is_aek_wall:
                print report_to_create, " is aek but not table nor wall"
                
            if not shortReport:
                doc.add_spacer()
                doc.add_table(aek_table, extra_style = aek_tableStyle, align="LEFT")
        machine_string = '{hospital_abbreviation}_{department_abbreviation}_{lab}_{manufacturer}_{model}'.format(
                        hospital_abbreviation=first_letters(hospital),
                        department_abbreviation=first_letters(department),
                        lab=lab, manufacturer=company, model=model_name)
        if shortReport:
            _fn = "%s_%s_%s (kort).pdf"  % (study_date, machine_string, report_to_create)
        else:
            _fn = "%s_%s_%s.pdf" % (study_date, machine_string, report_to_create)
        _fn = "Rapporter\\" + _fn
        doc.save(_fn)
        print "Lagret rapport: {}".format(os.path.abspath(_fn))
        webbrowser.open(_fn)
    
    db.commit()