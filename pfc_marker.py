#!/usr/bin/env python
# -*- coding: utf-8 -*-
# pfc_marker released under GPLv3

import sys, time, re, os, string, itertools, fileinput, shutil, subprocess, inspect, urllib.request
from lxml import etree
import lxml.etree as etree
from timecode import Timecode
from urllib.request import urlopen, Request
import requests
from PyQt5 import QtCore, QtSvg
from PyQt5.QtCore import QDir, Qt
from PyQt5.QtWidgets import QMessageBox, QApplication, QMainWindow, QInputDialog, QLineEdit, QFileDialog, QLabel
from pfc_marker_ui import *


version = "0.151203"
########################################################################

os.environ["REQUESTS_CA_BUNDLE"] = os.path.join(os.getcwd(), "cacert.pem")

programdata = os.getenv('ALLUSERSPROFILE')
print(programdata)
programdata_pfc_marker = os.path.join(programdata,'pfc_marker')
print(programdata_pfc_marker)
if not os.path.isdir(programdata_pfc_marker):
    os.makedirs(programdata_pfc_marker)

initcfg = os.path.join(programdata_pfc_marker, 'init.cfg')
xml_formatted_markers = os.path.join(programdata_pfc_marker, 'xml_formatted_markers.txt')


if os.path.isfile(initcfg):
    with open(initcfg, "r") as file:
        lines = file.readlines()
        recentprjfld = lines[0]
        checkupdateinitcfg = lines[1]
else:
    file = open(initcfg, "w")
    file.write("select pfrp\ncheckupdates=1")
    file.close()
    with open(initcfg, "r") as file:
        lines = file.readlines()
        recentprjfld = lines[0]
        checkupdateinitcfg = lines[1]

prjdic ={}


def about():

    aboutcontent =  "<p>nifty tool to inject cmx3600-edl events or dvs clipster markers</p>" \
                    "<p>as clip-markers into a sequence of your pfclean-project.</p>" \
                    "<p>written in python3.4.3, gui pyqt5, x64</p>" \
                    "<p><a href=http://suminus.github.io/pfc_marker style=\"color: black;\" ><b>open project page</a></p>" \
                   #"<p><a href='mailto:support@mariohartz.de' style=\"color: black;\" ><b>© mario hartz</a> </b></p>"

    reply = QMessageBox.information(None, "About", aboutcontent)
    if reply == QMessageBox.Ok:
        pass      
    else:
        pass

def selprj():
    ui.combobox_selectseq.setEnabled(False)
    openFilesPath = ''
    fileName, _ = QFileDialog.getOpenFileName(None, "select pfclean project", recentprjfld, "PFClean Prj (*.pfrp);;All Files (*.*)")
    if fileName:

        ui.combobox_selectseq.clear()
        ui.combobox_selectseq.setCurrentIndex(0)
        ui.combobox_selectseq.addItems(['select sequence'])
        ui.combobox_selectseq.setItemText(0, "select sequence")
        ui.combobox_selectseq.model().item(0).setEnabled(False)

        prj = fileName
        print(prj)
        prjdic['prj'] = prj
        recentprjfldnew = prj.rsplit('/', 1)[0]
        recentprjfldnew = recentprjfldnew + '/'
        prjdic['prjfolder'] = recentprjfldnew
        
        with open(initcfg, "r") as file:
            lines = file.readlines()
            lines[0] = '%s\n' % recentprjfldnew
        with open(initcfg, "w") as file:
            for line in lines:
                file.write(line)


        prjfile = prj.split('/')[-1:]
        prjfile = prjfile[0]
        print (prjfile)
        ui.label_msgs.setText('selected project: ' + prjfile)

        prjseqfld = prj.rsplit('/', 1)[0] + '/sequences'

        for prjseqid in os.listdir(prjseqfld):
            print('prjseqid: ' + prjseqid)
            prjseqpath = prjseqfld + '/' + prjseqid
            prjseqpath = str(prjseqpath.replace("\n", ""))
            print('prjseqpath: ' + prjseqpath)
            with open(prjseqpath, "r") as prjseqfile:
                for line in itertools.islice(prjseqfile, 3, 4):
                    prjseqfile = line.replace("\n", "")
                    prjseqname = re.compile('<name>(.*?)</name>', re.DOTALL |  re.IGNORECASE).findall(prjseqfile)
                    print('prjseqname[0]: ' + prjseqname[0])
                    pattern = '<endFrame>0</endFrame>'

                    if str(prjseqname[0]) == 'Rough edit 1':
                        if pattern in open(prjseqpath).read():
                            print('0 lenght ---- boom!!')
                            ui.label_msgs.append('filtered "Rough edit 1" as pfclean internal zero-length seq')
                        else:
                            ui.combobox_selectseq.addItems(prjseqname)
                            prjseqname = ' '.join(prjseqname)
                            prjdic[prjseqname] = prjseqid
                            print(prjseqname + '   aka   ' + prjseqid)
                    elif str(prjseqname[0]) == 'Sequence 1':
                        if pattern in open(prjseqpath).read():
                            print('0 lenght ---- boom!!')
                            ui.label_msgs.append('filtered "Sequence 1" as pfclean internal zero-length seq')
                        else:
                            ui.combobox_selectseq.addItems(prjseqname)
                            prjseqname = ' '.join(prjseqname)
                            prjdic[prjseqname] = prjseqid
                            print(prjseqname + '   aka   ' + prjseqid)
                    else:
                        ui.combobox_selectseq.addItems(prjseqname)
                        prjseqname = ' '.join(prjseqname)
                        prjdic[prjseqname] = prjseqid
                        print(prjseqname + '   aka   ' + prjseqid)
                        
                    
        prjdic['prjseqfolder'] = prjseqfld
        print(prjdic)  
              
        ui.combobox_selectseq.setEnabled(True)
        ui.combobox_selectseq.currentIndexChanged['QString'].connect(enablebtn)
        
    else:
        print('no project selected!')
        ui.label_msgs.setText('no pfclean-project selected!')
    ui.checkBox_srctc.setEnabled(False)
    ui.checkBox_rectc.setEnabled(False)

def enablebtn():
    selected_prjseqname = str(ui.combobox_selectseq.currentText())
    print("seleceted prjseqname in combobox_selectseq: " + selected_prjseqname)
    if str(selected_prjseqname) != 'select sequence':
        if (selected_prjseqname) is not '':
            ui.checkBox_srctc.setEnabled(True)
            ui.checkBox_rectc.setEnabled(True)
            ui.btn_seledl.setEnabled(True)
            ui.btn_selcp.setEnabled(True)
            ui.btn_selcsv.setEnabled(True)
            selected_prjseqname = str(ui.combobox_selectseq.currentText())
            print('seleceted prjseqname in combobox_selectseq -> filtered: ' + selected_prjseqname)

            selected_prjseqid = prjdic.get(selected_prjseqname)
            print('dict result selected_prjseqid: ' + str(selected_prjseqid))
            prjseqfld = prjdic.get('prjseqfolder')
            print('dict result prjseqfld: ' + str(prjseqfld))

            tree = etree.parse(str(prjseqfld + '/' + selected_prjseqid))
            root = tree.getroot()
            prjseqname = root.findtext("name")
            prjdic['prjseqname'] = prjseqname
            prjseqid = root.findtext("uniqueId")
            prjdic['prjseqid'] = prjseqid
            prjseqstart = root.findtext("startFrame")
            prjdic['prjseqstart'] = prjseqstart
            prjseqend = root.findtext("endFrame")
            prjdic['prjseqend'] = prjseqend
            prjseqoff = root.findtext("startFrameOffset")
            if prjseqoff is None:
                prjdic['prjseqoff'] = 0
            else:
                prjdic['prjseqoff'] = prjseqoff
            prjseqfps = root.findtext("frameRate")
            prjdic['prjseqfps'] = prjseqfps
            prjseqdf = root.findtext("dropFrame")
            prjdic['prjseqdf'] = str(prjseqdf)
            ui.label_msgs.append('seclected sequence <b>' + prjseqname + '</b> is <b>' + prjseqfps + 'fps.')
        else:
            pass
    else:
        ui.btn_seledl.setEnabled(False)
        ui.btn_import.setEnabled(False) 
        ui.btn_xml.setEnabled(False)
        ui.checkBox_srctc.setEnabled(False)
        ui.checkBox_rectc.setEnabled(False)

def checkboxrectc():
    if ui.checkBox_rectc.isChecked():
        ui.checkBox_srctc.setChecked(False)
    else:
        ui.checkBox_srctc.setChecked(True)

def checkboxsrctc():
    if ui.checkBox_srctc.isChecked():
        ui.checkBox_rectc.setChecked(False)
    else:
        ui.checkBox_rectc.setChecked(True)

def seledl():
    selected_prjseqname = str(ui.combobox_selectseq.currentText())

    if ui.checkBox_rectc.isChecked():
        conformmode = 6
        conformmodemsg = 'and <b> record tc.'
    else:
        pass
    if ui.checkBox_srctc.isChecked():
        conformmode = 4
        conformmodemsg = 'and <b> source tc.'
    else:
        pass

    fileName, _ = QFileDialog.getOpenFileName(None, "select cmx 3600 edl with pfclean project framerate", '' , "EDL's (*.edl);;All Files (*.*)")
    if fileName:
        edlpath = fileName
        edl = edlpath.split('/')[-1:]
        edl = ' '.join(edl)
        tclist = []
        with open(edlpath) as f:
            for line in f:
                check = line
                if check[0].isdigit():
                    linecons =' '.join(line.split())
                    linecons = linecons.split(' ')
                    if linecons[3] == 'D':
                        print('--> dissolve detected!!')
                    else:
                        tclist.append(linecons[conformmode])  # conformmode srctc = 4     rectc = 6
                else:
                    print('-> skipped line: ' + check + 'not a valid event entry!' )
                    pass
            print(tclist)
            print('edl-events: ' + str(len(tclist)))

        prjseqfps = prjdic.get('prjseqfps')
        tclistframes = []
        for event in tclist:
            tcfr = Timecode(prjseqfps, event)
            eventframes = tcfr.frames
            eventframes = eventframes - 1 # remove one frame offset
            tclistframes.append(eventframes)
        
        prjseqoff = prjdic.get('prjseqoff')
        prjseqid = prjdic.get('prjseqid')
        prjseqstart = prjdic.get('prjseqstart')
        prjseqend = prjdic.get('prjseqend')
        file = open(xml_formatted_markers, "w")
        file.write('\t<clipFrameMarkers>\n\t\t<clipFrameMarker>\n\t\t\t<identifier>%s</identifier>\n\t\t\t<counter>%s</counter>\n\t\t\t<frameMarkers>' % (prjseqid, prjseqoff))
        count = 0
        for event in tclistframes:
            if  event >= int(prjseqstart) and event <= int(prjseqend):
                count +=1
                insert1 ="\n\t\t\t\t<frameMarker>"
                insert2 ="\n\t\t\t\t\t<frame>%s</frame>" % event #\n
                insert3 ="\n\t\t\t\t\t<name>%s</name>" % edl
                insert4 ="\n\t\t\t\t\t<notes>%s</notes>" % ('based on: ' + edl + ' parsed at: ' + time.strftime("%Y-%d-%m-%H-%M-%S"))
                insert5 ="\n\t\t\t\t</frameMarker>"
                file = open(xml_formatted_markers, "a+")
                file.write(str(insert1))
                file.write(str(insert2))
                file.write(str(insert3))
                file.write(str(insert4))
                file.write(str(insert5))
                file.close()
        file = open(xml_formatted_markers, "a+")
        file.write('\n\t\t\t</frameMarkers>\n\t\t</clipFrameMarker>\n\t</clipFrameMarkers>')
        file.close()

        ui.label_msgs.append(edl + ' contains ' + str(len(tclist)) + ' entries.')
        ui.label_msgs.append('parsed with <b>' + prjseqfps + 'fps </b>' + conformmodemsg )
        ui.label_msgs.append('edl events in sequence range: <b>' + str(count))

        if count == 0:
            ui.btn_import.setEnabled(False)
            ui.btn_xml.setEnabled(False)
        else:
            ui.btn_import.setEnabled(True)
            ui.btn_xml.setEnabled(True)
        ui.checkBox_srctc.setEnabled(False)
        ui.checkBox_rectc.setEnabled(False)
    else:
        ui.label_msgs.append('no edl selected')

def selcp():
    openFilesPath = ''
    fileName, _ = QFileDialog.getOpenFileName(None, "select Clipster Project", recentprjfld , "Clipster Project (*.cp);;All Files (*.*)")
    if fileName:
        cprjpath = fileName
        print(cprjpath)
        prjdic['cprjpath'] = cprjpath
        cprjfile = cprjpath.split('/')[-1:]
        cprjfile = cprjfile[0]
        print (cprjfile)
        ui.label_msgs.append('selected clipster project:<b> ' + cprjfile)

        prjseqoff = prjdic.get('prjseqoff')
        prjseqid = prjdic.get('prjseqid')
        prjseqstart = prjdic.get('prjseqstart')
        prjseqend = prjdic.get('prjseqend')
        prjseqfps = prjdic.get('prjseqfps')
        
        with open(cprjpath) as f:
            for line in f:
                if line.startswith('<TIMECODE OFFSET='):
                    cptcoffset = re.findall('"([^"]*)"', line )
                    cptcoffset = cptcoffset[0]
                    tcfr = Timecode(prjseqfps, cptcoffset)
                    cptcoffset = tcfr.frames -1
                    prjdic['cptcoffset'] = cptcoffset
                    break       # break on first match, is the right one
                else:
                    pass

        cptclist = []
        cptclistframes = []

        with open(cprjpath) as f:
            file = open(xml_formatted_markers, "w")
            file.write('\t<clipFrameMarkers>\n\t\t<clipFrameMarker>\n\t\t\t<identifier>%s</identifier>\n\t\t\t<counter>%s</counter>\n\t\t\t<frameMarkers>' % (prjseqid, prjseqoff))
            count = 0

            for line in f:
                linecons =' '.join(line.split())
                linecons = linecons.split(' ')
                if linecons[0] == '<MARKER':
                    event = re.findall('"([^"]*)"', str(linecons[1]))
                    #comment = re.findall('\"(.+?)\"', str(linecons))
                    comment = re.findall(r'"([^"]*)"', str(line))
                    del comment[0]
                    comment = ' '.join(comment)
                    cptclist.append(event)
                    eventfr = int(int(event[0]) / 40000000)
                    print(eventfr)
                    eventfroff = eventfr + cptcoffset
                    cptclistframes.append(eventfroff)
                    print (cptclistframes)

                    if  eventfroff >= int(prjseqstart) and eventfroff <= int(prjseqend):
                        count +=1
                        insert1 ="\n\t\t\t\t<frameMarker>"
                        insert2 ="\n\t\t\t\t\t<frame>%s</frame>" % eventfroff #\n
                        insert3 ="\n\t\t\t\t\t<name>%s</name>" % cprjfile
                        insert4 ="\n\t\t\t\t\t<notes>%s</notes>" % comment
                        insert5 ="\n\t\t\t\t</frameMarker>"
                        file = open(xml_formatted_markers, "a+")
                        file.write(str(insert1))
                        file.write(str(insert2))
                        file.write(str(insert3))
                        file.write(str(insert4))
                        file.write(str(insert5))
                        file.close()
                    else:
                        pass
                else:
                    pass

            file = open(xml_formatted_markers, "a+")
            file.write('\n\t\t\t</frameMarkers>\n\t\t</clipFrameMarker>\n\t</clipFrameMarkers>')
            file.close()


        ui.label_msgs.append(cprjfile + ' contains ' + str(len(cptclist)) + ' timeline-marker entries.')
        ui.label_msgs.append(cprjfile + ' timeline offset is: ' + str(cptcoffset))
        ui.label_msgs.append('parsed with <b>' + prjseqfps + 'fps </b>')
        ui.label_msgs.append('clipster events in sequence range: <b>' + str(count))

        if count == 0:
            ui.btn_import.setEnabled(False)
            ui.btn_xml.setEnabled(False)
        else:
            ui.btn_import.setEnabled(True)
            ui.btn_xml.setEnabled(True)
        
        ui.checkBox_srctc.setEnabled(False)
        ui.checkBox_rectc.setEnabled(False)

    else:
        ui.label_msgs.append('no clipster-project selected!')

def selcsv():
    openFilesPath = ''
    fileName, _ = QFileDialog.getOpenFileName(None, "select csv file with new line with new event", '' , "csv textfile (*.txt);;All Files (*.*)")
    if fileName:
        prjseqfps = prjdic.get('prjseqfps')
        prjseqoff = prjdic.get('prjseqoff')
        prjseqid = prjdic.get('prjseqid')
        prjseqstart = prjdic.get('prjseqstart')
        prjseqend = prjdic.get('prjseqend')

        csvpath = fileName
        csvfile = csvpath.split('/')[-1:]
        csvfile = ' '.join(csvfile)
        ui.label_msgs.append('selected csv:<b> ' + csvfile)
        csvtclist = []
        with open(csvpath) as f:
            file = open(xml_formatted_markers, "w")
            file.write('\t<clipFrameMarkers>\n\t\t<clipFrameMarker>\n\t\t\t<identifier>%s</identifier>\n\t\t\t<counter>%s</counter>\n\t\t\t<frameMarkers>' % (prjseqid, prjseqoff))
            count = 0

            for line in f:
                if line[0].isdigit():
                    linecons =' '.join(line.split())
                    linecons = linecons.split(',')
                    csvevent = linecons[0]
                    csvnotes = linecons[1]
                    csvtclist.append(csvevent)
                    print(linecons[0])
                    print(linecons[1])

                    csvtclistframes = []
                    csvtcfr = Timecode(prjseqfps, csvevent)
                    eventframes = csvtcfr.frames
                    eventframes = eventframes - 1 # remove one frame offset
                    csvtclistframes.append(eventframes)
                    print(csvtclistframes)
                    if  eventframes >= int(prjseqstart) and eventframes <= int(prjseqend):
                        count +=1
                        insert1 ="\n\t\t\t\t<frameMarker>"
                        insert2 ="\n\t\t\t\t\t<frame>%s</frame>" % eventframes #\n
                        insert3 ="\n\t\t\t\t\t<name>%s</name>" % csvfile
                        insert4 ="\n\t\t\t\t\t<notes>%s</notes>" % csvnotes
                        insert5 ="\n\t\t\t\t</frameMarker>"
                        file = open(xml_formatted_markers, "a+")
                        file.write(str(insert1))
                        file.write(str(insert2))
                        file.write(str(insert3))
                        file.write(str(insert4))
                        file.write(str(insert5))
                        file.close()
                    file = open(xml_formatted_markers, "a+")
                    file.write('\n\t\t\t</frameMarkers>\n\t\t</clipFrameMarker>\n\t</clipFrameMarkers>')
                    file.close()
                    ui.label_msgs.append(csvfile + ' contains ' + str(len(csvtclist)) + ' entries.')
                    ui.label_msgs.append('csv events in sequence range: <b>' + str(count))
                    print(csvtclist)
                    print('csv-events: ' + str(len(csvtclist)))
                else:
                    print('-> skipped line: ' + line[0] + 'not a valid event entry!' )
                    pass
        if count == 0:
            ui.btn_import.setEnabled(False)
            ui.btn_xml.setEnabled(False)
        else:
            ui.btn_import.setEnabled(True)
            ui.btn_xml.setEnabled(True)
        
        ui.checkBox_srctc.setEnabled(False)
        ui.checkBox_rectc.setEnabled(False)

def savexmlmarker():
    filename, _ = QFileDialog.getSaveFileName(None, 'Save XML-formated Markers to Text-File', str(prjdic.get('prjfolder') + prjdic.get('prj') +'_xml_markers.txt') , "Text Files (*.txt);; All Files (*)")
    if filename:
        with open(xml_formatted_markers, 'rb') as f:
            data = f.read()
        with open(filename, 'wb') as f:
            f.write(data)
        ui.label_msgs.append('file saved as: \n' + str(filename))
        os.system(filename)
    else:
        pass

def inject():

    prjbkp = str(prjdic.get('prj')) + time.strftime("%Y-%d-%m-%H-%M-%S") + '.bkp'
    shutil.copyfile(str(prjdic.get('prj')), prjbkp)

    temp = os.path.join(programdata_pfc_marker, 'temp')
    injectmarkers = os.path.join(programdata_pfc_marker, 'xml_formatted_markers.txt')
    injectmarkers = etree.parse(injectmarkers)
    injectmarkers = injectmarkers.find('clipFrameMarker')

    prjseqid = prjdic.get('prjseqid')
    prj = prjdic.get('prj')
    tree = etree.parse(prj)

    allmarkers = tree.find('clipFrameMarkers')
    print('orig-prj-markers: ' + str(allmarkers))
    check = allmarkers.text

    if check is not None:
        clipFrameMarker = allmarkers.findall('clipFrameMarker')
        for clipFrameMarkerSeq in clipFrameMarker:
            identifiers = clipFrameMarkerSeq.findall('identifier')
            for identifier in identifiers:
                if identifier.text == prjseqid:
                    print('found prjseqid in orig: ' + prjseqid)
                    origclipframemarker = identifier.getparent()
                    print('parent-name in orig: ' + origclipframemarker.tag)
                    origclipframemarker.clear()
                    origclipframemarker.insert(0, injectmarkers)

        f1 = open(temp, 'wb')
        f1.write(etree.tostring(tree, pretty_print=False))
        f1.close()

        f1 = open(temp, 'r')
        f2 = open(temp + 'new1', 'w')
        for line in f1:
            f2.write(line.replace('<clipFrameMarker><clipFrameMarker>', '<clipFrameMarker>'))
        f1.close()
        f2.close()

        f1 = open(temp + 'new1', 'r')
        f2 = open(temp + 'new2', 'w')
        for line in f1:
            f2.write(line.replace('</clipFrameMarker><clipFrameMarker>', '\t<clipFrameMarker>'))
        f1.close()
        f2.close()
            
        f1 = open(temp + 'new2', 'r')
        f2 = open(prj, 'w')
        for line in f1:
            f2.write(line.replace('</clipFrameMarker></clipFrameMarkers>', '</clipFrameMarkers>'))
        f1.close()
        f2.close()

    else:
        # no markers at all
        allmarkers.insert(0, injectmarkers)

        f1 = open(temp, 'wb')
        f1.write(etree.tostring(tree, pretty_print=False))
        f1.close()

        f1 = open(temp, 'r')
        f2 = open(prj, 'w')
        for line in f1:
            f2.write(line.replace('<clipFrameMarkers><clipFrameMarker>', '<clipFrameMarkers>\n\t\t<clipFrameMarker>'))
        f1.close()
        f2.close()
            
    ui.label_msgs.append('import successful. <b> moooh!')

    """
    # <clipFrameMarkers/> # only present, when no markers in whole project
    """

def exit():
    sys.exit(0)

def blink():
    ui.logo.setPixmap(QtGui.QPixmap(":/pfc_marker_logo.svg"))

def logochange():
    try:
        ui.logo.setPixmap(QtGui.QPixmap(":/pfc_marker_logo_closed.svg"))
        timer.singleShot(200, blink)
    finally:
        pass

def check_internet():
    try:
        network = urllib.request.urlopen('http://www.google.com', timeout=1)
        return True
    except urllib.request.URLError:
        return False

def checkupdateconf():
    if ui.actionCheckUpdates.isChecked():
        with open(initcfg, "r") as file:
            lines = file.readlines()
            lines[1] = 'checkupdates=1'
        with open(initcfg, "w") as file:
            for line in lines:
                file.write(line)
    else:
        with open(initcfg, "r") as file:
            lines = file.readlines()
            lines[1] = 'checkupdates=0'
        with open(initcfg, "w") as file:
            for line in lines:
                file.write(line)


def checkupdate():

    if check_internet():
        buildmsi = version
        msi_update = 'https://raw.githubusercontent.com/suminus/pfc_marker/master/msi/current_version'
        
        check = requests.head(msi_update, verify=find_data_file('cacert.pem')).status_code

        if check == 200:
            onlinemsi = urlopen(msi_update).read().decode('utf_8')
            onlinemsi = str(onlinemsi)
            onlinemsi = onlinemsi.replace("\n", "")
            print('update msi textfile: ' + onlinemsi)

            if onlinemsi > buildmsi:
                ui.menuUpdate.setTitle("Update available!")
                ui.actionUpdate.setEnabled(True)
                prjdic['onlinemsi'] = onlinemsi
        else:
            print(check)
            ui.menuUpdate.setTitle("")
    else:
        ui.menuUpdate.setTitle("")
        ui.label_msgs.append('no internet access')

def update():
    updatecontent = "<p>do you want to download the latest msi-package?</p>" \
                    "<p>pfc_marker will close and start the update automatically.</p>" 

    reply = QMessageBox.question(None, "Update", updatecontent, QMessageBox.Yes | QMessageBox.No )
    if reply == QMessageBox.Yes:
        onlinemsi = prjdic.get('onlinemsi')
        ui.label_msgs.append('downloading update. please wait ...')
        ui.statusbar.showMessage('downloading update. please wait ...')
        msi_updateurl = 'https://github.com/suminus/pfc_marker/blob/master/msi/pfc_marker-' + onlinemsi + '-amd64.msi?raw=true'
        print('msi_updateurl: ' + str(msi_updateurl))
        msi_update = 'pfc_marker-' + onlinemsi + '-amd64.msi'
        print('msi: ' + str(msi_update))

        check = requests.head(msi_updateurl, verify=find_data_file('cacert.pem')).status_code
        print('   RET:   ' + str(check))
        if check == 302:
            with urllib.request.urlopen(msi_updateurl) as response, open(os.path.join(os.getenv('TEMP'), msi_update), 'wb') as out_file:
                data = response.read()
                out_file.write(data)
            os.startfile(os.path.join(os.getenv('TEMP'), msi_update))
            exit()
        else:
            ui.label_msgs.append('updatefile not found!')
            ui.statusbar.showMessage('updatefile not found!')

    elif reply == QMessageBox.No:
        print("No")
    else:
        print("Cancel")

def find_data_file(filename):
    if getattr(sys, 'frozen', False):
        # The application is frozen
        datadir = os.path.dirname(sys.executable)
    else:
        # The application is not frozen
        # Change this bit to match where you store your data files:
        datadir = os.path.dirname(__file__)

    return os.path.join(datadir, filename)



if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.setWindowTitle("pfclean marker %s x64" % version)

    timer=QtCore.QTimer()
    timer.timeout.connect(logochange)
    timer.start(6000)

    ui.actionExit.triggered.connect(exit)
    ui.actionAbout.triggered.connect(about)
    ui.btn_selprj.clicked.connect(selprj)
    
    ui.btn_seledl.setEnabled(False)
    ui.btn_seledl.clicked.connect(seledl)
    
    ui.btn_selcp.clicked.connect(selcp)
    ui.btn_selcp.setEnabled(False)

    ui.btn_selcsv.clicked.connect(selcsv)
    ui.btn_selcsv.setEnabled(False)

    ui.menuUpdate.setTitle("")
    ui.btn_import.setEnabled(False)
    ui.checkBox_srctc.setEnabled(False)
    ui.checkBox_rectc.setEnabled(False)

    ui.btn_xml.clicked.connect(savexmlmarker)
    ui.checkBox_rectc.stateChanged.connect(checkboxrectc)
    ui.checkBox_srctc.stateChanged.connect(checkboxsrctc)
    ui.btn_import.clicked.connect(inject)
    ui.label_msgs.setText('')
    ui.combobox_selectseq.setEnabled(False)
    ui.combobox_selectseq.setItemText(0, "select sequence")
    ui.combobox_selectseq.model().item(0).setEnabled(False)

    ui.actionUpdate.triggered.connect(update)
    ui.actionCheckUpdates.triggered.connect(checkupdateconf)

    if checkupdateinitcfg == 'checkupdates=1':
        ui.actionCheckUpdates.setChecked(True)
        timer.singleShot(10000, checkupdate)
    else:
        ui.actionCheckUpdates.setChecked(False)
    

    MainWindow.show()
    sys.exit(app.exec_())
