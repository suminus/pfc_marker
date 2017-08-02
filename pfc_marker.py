#!/usr/bin/env python
# -*- coding: utf-8 -*-
# pfc_marker released under GPLv3

import sys, time, re, os, string, itertools, fileinput, shutil, subprocess, inspect, urllib.request, requests
import lxml.etree as etree
from lxml import etree
from socket import timeout
from timecode import Timecode
from urllib.request import urlopen, Request, urlretrieve
from PyQt5 import QtCore, QtSvg
from PyQt5.QtCore import QDir, Qt
from PyQt5.QtWidgets import QMessageBox, QApplication, QMainWindow, QInputDialog, QLineEdit, QFileDialog, QLabel
from pfc_marker_ui import *


version = "0.170802"
########################################################################

os.environ["REQUESTS_CA_BUNDLE"] = os.path.join(os.getcwd(), "cacert.pem")

#remove old versions programdata init, due to user-rights issues
programdata = os.getenv('ALLUSERSPROFILE')
print(programdata)
programdata_pfc_marker = os.path.join(programdata,'pfc_marker')
print(programdata_pfc_marker)
if os.path.isdir(programdata_pfc_marker):
    shutil.rmtree(programdata_pfc_marker)

userhome = os.getenv('USERPROFILE')
appdata_pfc_marker = os.path.join(userhome,'AppData', 'Roaming', 'pfc_marker')
if not os.path.isdir(appdata_pfc_marker): 
    os.makedirs(appdata_pfc_marker)
print(appdata_pfc_marker)

initcfg = os.path.join(appdata_pfc_marker, 'init.cfg')
xml_formatted_markers = os.path.join(appdata_pfc_marker, 'xml_formatted_markers.txt')

if os.path.isfile(initcfg):
    with open(initcfg, "r") as file:
        lines = file.readlines()
        recentprjfld = lines[0]
        checkupdateinitcfg = lines[1]
else:
    file = open(initcfg, "w")
    file.write("select projectfile\ncheckupdates=1")
    file.close()
    with open(initcfg, "r") as file:
        lines = file.readlines()
        recentprjfld = lines[0]
        checkupdateinitcfg = lines[1]

prjdic = {}


def about():

    aboutcontent =  "<p>nifty tool to inject "\
                    "cmx3600-edl events, a csv-list, dvs clipster timeline markers <br>" \
                    "or avid mediacomposer markers as clip-markers into an existing sequence <br>" \
                    "in your pfclean-project.<br><br>" \
                    "written in python3.4.3, gui pyqt5, win 7/8/10 x64</p>" \
                    "<p><a href=http://suminus.github.io/pfc_marker style=\"color: black;\" ><b>visit project page</a></p>" \


    reply = QMessageBox.information(None, "About", aboutcontent)
    if reply == QMessageBox.Ok:
        pass      
    else:
        pass

def selprj():
    disablebtn()
    openFilesPath = ''

    fileName, _ = QFileDialog.getOpenFileName(None, "select pfclean project", recentprjfld, "PFClean Prj (*.pfrp);;All Files (*.*)")
    if fileName:
        print('fileName: {}' .format(fileName))
        ui.combobox_selectseq.clear()
        ui.combobox_selectseq.setCurrentIndex(0)
        ui.combobox_selectseq.addItems(['select sequence'])
        ui.combobox_selectseq.setItemText(0, "select sequence")
        ui.combobox_selectseq.model().item(0).setEnabled(False)

        prj = fileName
        
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
        prjdic['prjfile'] = prjfile
        print (prjfile)
        ui.label_msgs.setText('selected project: {}' .format(prjfile))

        prjseqfld = prj.rsplit('/', 1)[0] + '/sequences'
        prjseqname_list = []
        for prjseqid in os.listdir(prjseqfld):
            print('prjseqid: {}' .format(prjseqid))
            prjseqpath = '{}/{}' .format(prjseqfld, prjseqid)
            print('prjseqpath: {}' .format(prjseqpath))
            with open(prjseqpath, "r") as prjseqfile:
                for line in itertools.islice(prjseqfile, 3, 4):
                    prjseqfile = line.replace("\n", "")
                    prjseqname = re.compile('<name>(.*?)</name>', re.DOTALL |  re.IGNORECASE).findall(prjseqfile)
                    print('prjseqname[0]: {}' .format(prjseqname[0]))
                    pattern = '<endFrame>0</endFrame>'

                    if str(prjseqname[0]) == 'Rough edit 1' or 'Sequence 1':
                        if pattern in open(prjseqpath).read():
                            print(prjseqname[0] + ' with 0 lenght ---- boom!!')
                            ui.label_msgs.append('filtered %s as pfclean zero-length seq' %(str(prjseqname[0])))
                        else:
                            prjseqname_list.append(str(prjseqname))
                            prjseqname = ' '.join(prjseqname)
                            prjdic[prjseqname] = prjseqid
                            print(prjseqname + '   aka   ' + prjseqid)
                    
        prjdic['prjseqfolder'] = prjseqfld
        print(prjdic)
        for item in sorted(prjseqname_list):
            item = item.strip('[]').strip("''")
            ui.combobox_selectseq.addItem(item)
        ui.combobox_selectseq.setEnabled(True)
        ui.combobox_selectseq.currentIndexChanged['QString'].connect(enablebtn)
        
    else:
        print('no project selected!')
        ui.label_msgs.setText('no pfclean-project selected!')
        ui.combobox_selectseq.clear()
        ui.combobox_selectseq.setCurrentIndex(0)
        ui.combobox_selectseq.addItems(['select sequence'])
        ui.combobox_selectseq.setItemText(0, "select sequence")
        ui.combobox_selectseq.model().item(0).setEnabled(False)
        disablebtn()

def enablebtn():
    selected_prjseqname = str(ui.combobox_selectseq.currentText())
    if str(selected_prjseqname) != 'select sequence':
        if (selected_prjseqname) is not '':
            ui.checkBox_srctc.setEnabled(True)
            ui.checkBox_rectc.setEnabled(True)
            ui.btn_seledl.setEnabled(True)
            ui.btn_selcp.setEnabled(True)
            ui.btn_selcsv.setEnabled(True)
            ui.btn_selavid.setEnabled(True)
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
            ui.label_msgs.append('seclected sequence <b> {} </b> is <b> {} fps.' .format(prjseqname, prjseqfps))
            ui.label_msgs.append('startframe: <b> {} </b>endframe: <b> {}' .format (prjseqstart, prjseqend))

    else:
        disablebtn()
        pass

def disablebtn():
    try: ui.combobox_selectseq.disconnect() # avoid multiple times same output in messagebox after selecting a sequence
    except Exception: pass

    prjdic.clear()
    ui.combobox_selectseq.setEnabled(False)
    ui.btn_seledl.setEnabled(False)
    ui.btn_selcp.setEnabled(False)
    ui.btn_selcsv.setEnabled(False)
    ui.btn_selavid.setEnabled(False)
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
        conformmode = 6 # column to read from edl
        conformmodemsg = 'and <b> record tc.'
    else:
        pass
    if ui.checkBox_srctc.isChecked():
        conformmode = 4  # column to read from edl
        conformmodemsg = 'and <b> source tc.'
    else:
        pass

    fileName, _ = QFileDialog.getOpenFileName(None, "select cmx 3600 edl with pfclean project framerate", '' , "EDL's (*.edl);;All Files (*.*)")
    if fileName:
        prjseqoff = prjdic.get('prjseqoff')
        prjseqid = prjdic.get('prjseqid')
        prjseqstart = prjdic.get('prjseqstart')
        prjseqend = prjdic.get('prjseqend')
        prjseqfps = prjdic.get('prjseqfps')
        edlpath = fileName
        edl = edlpath.split('/')[-1:]
        edl = ' '.join(edl)
        tclist = []
        count = 0
        file = open(xml_formatted_markers, "w")
        file.write('\t<clipFrameMarkers>\n\t\t<clipFrameMarker>\n\t\t\t<identifier>{}</identifier>\n\t\t\t<counter>{}</counter>\n\t\t\t<frameMarkers>' .format(prjseqid, prjseqoff))

        with open(edlpath) as f:
            for line in f:
                if line[0].isdigit():
                    linecons =' '.join(line.split())
                    linecons = linecons.split(' ')
                    if linecons[3] == 'D':
                        print('--> dissolve detected!!')
                    else:
                        tclist.append(linecons[conformmode])
                        tcfr = Timecode(prjseqfps, linecons[conformmode])
                        eventframes = tcfr.frames
                        eventframes = eventframes - 1
                        if  eventframes >= int(prjseqstart) and eventframes <= int(prjseqend):
                            count +=1
                            insert1 ="\n\t\t\t\t<frameMarker>"
                            insert2 ="\n\t\t\t\t\t<frame>%s</frame>" % eventframes #\n
                            insert3 ="\n\t\t\t\t\t<name>%s</name>" % str('edl: ' + edl)
                            insert4 ="\n\t\t\t\t\t<notes>%s</notes>" % str('tape: ' + linecons[1]) # usually tapename
                            insert5 ="\n\t\t\t\t</frameMarker>"
                            file = open(xml_formatted_markers, "a+")
                            file.write(str(insert1))
                            file.write(str(insert2))
                            file.write(str(insert3))
                            file.write(str(insert4))
                            file.write(str(insert5))
                            file.close()
                else:
                    print('-> skipped line: ' + line[0] + 'not a valid event entry!' )
                    pass
            print('edl-events: ' + str(len(tclist)))
        file = open(xml_formatted_markers, "a+")
        file.write('\n\t\t\t</frameMarkers>\n\t\t</clipFrameMarker>\n\t</clipFrameMarkers>')
        file.close()

        ui.label_msgs.append(edl + ' contains <b> {}</b> entries.' .format(len(tclist)))
        ui.label_msgs.append('edl events in sequence range: <b> {}' .format(count))
        ui.label_msgs.append('parsed with <b>{} fps </b>{}' .format(prjseqfps, conformmodemsg))

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
        ui.label_msgs.append('selected clipster project:<b> {}' .format(cprjfile))

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







        
        ## extract all MARKERLIST-part for clean up \n
        with open(cprjpath) as f:
            xmldata = f.read()
            markerlist = re.findall(r'<MARKERLIST>(.*?)</MARKERLIST>',xmldata,re.DOTALL)
            markerlist = "".join(markerlist).replace('\n',' ')
            markerlist = "".join(markerlist).replace('/>','/>\n')
            xml_markerlist_cleanup = os.path.join(appdata_pfc_marker, 'xml_markerlist_cleanup.txt')
            file = open(xml_markerlist_cleanup, "w")
            file.write(markerlist)
            file.close()

        with open(xml_markerlist_cleanup) as f:
            file = open(xml_formatted_markers, "w")
            file.write('\t<clipFrameMarkers>\n\t\t<clipFrameMarker>\n\t\t\t<identifier>%s</identifier>\n\t\t\t<counter>%s</counter>\n\t\t\t<frameMarkers>' % (prjseqid, prjseqoff))
            count = 0
            
            markername = cprjfile  # set generic 170802


            for line in f:
                linecons =' '.join(line.split())
                linecons = linecons.split(' ')
                if linecons[0] == '<MARKER':
                    event = re.findall('"([^"]*)"', str(linecons[1]))
                    marknamecomment = re.findall(r'"([^"]*)"', str(line))
                    del marknamecomment[0]
                    if len(marknamecomment) == 0:
                    	   markername = "no-name"
                    	   comment = "no-comment"

                    if len(marknamecomment) == 2:
                        markername = marknamecomment[0]
                        comment = marknamecomment[1]
                        print('markername: {}' .format(markername))
                        print('comment: {}' .format(comment))

                    if len(marknamecomment) == 1:
                        comment = marknamecomment[0]
                        markername = cprjfile
                        print('comment: {}' .format(comment))
                        print('no markername using cprjfile: {}' .format(markername))

                    cptclist.append(event)
                    nanotofr = int(1000000000 / int(prjseqfps))
                    eventfr = int(int(event[0]) / nanotofr)    #40000000 @ 25fps
                    print(eventfr)
                    eventfroff = eventfr + cptcoffset
                    cptclistframes.append(eventfroff)
                    print (cptclistframes)

                    if  eventfroff >= int(prjseqstart) and eventfroff <= int(prjseqend):
                        count +=1
                        insert1 ="\n\t\t\t\t<frameMarker>"
                        insert2 ="\n\t\t\t\t\t<frame>%s</frame>" % eventfroff #\n
                        insert3 ="\n\t\t\t\t\t<name>%s</name>" % str('cp: ' + markername)
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

        ui.label_msgs.append('{} contains <b> {} </b> timeline-marker entries.' .format(cprjfile, len(cptclist)))
        ui.label_msgs.append('{} timeline offset is: <b>{}</b>' .format(cprjfile, cptcoffset))
        ui.label_msgs.append('parsed with <b> {} fps </b>' .format(prjseqfps))
        ui.label_msgs.append('clipster events in sequence range: <b>{}' .format(count))

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
                    csvnoteslen = int(len(linecons))
                    csvnoteslist = []
                    for i in range(csvnoteslen):
                        if i > 0:
                            csvnoteslist.append(linecons[i])
                    print(csvnoteslist)
                    csvtclist.append(csvevent)
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
                        insert3 ="\n\t\t\t\t\t<name>%s</name>" % str('csv: ' + csvfile)
                        insert4 ="\n\t\t\t\t\t<notes>%s</notes>" % (','.join(csvnoteslist))
                        insert5 ="\n\t\t\t\t</frameMarker>"
                        file = open(xml_formatted_markers, "a+")
                        file.write(str(insert1))
                        file.write(str(insert2))
                        file.write(str(insert3))
                        file.write(str(insert4))
                        file.write(str(insert5))
                        file.close()

                    print(csvtclist)
                    print('csv-events: ' + str(len(csvtclist)))
                else:
                    print('-> skipped line: ' + line[0] + 'not a valid event entry!' )
                    pass
            file = open(xml_formatted_markers, "a+")
            file.write('\n\t\t\t</frameMarkers>\n\t\t</clipFrameMarker>\n\t</clipFrameMarkers>')
            file.close()
            
        if count == 0:
            ui.label_msgs.append(csvfile + ' contains ' + str(len(csvtclist)) + ' entries.')
            ui.label_msgs.append('csv events in sequence range: <b>' + str(count))
            ui.btn_import.setEnabled(False)
            ui.btn_xml.setEnabled(False)
        else:
            ui.label_msgs.append(csvfile + ' contains ' + str(len(csvtclist)) + ' entries.')
            ui.label_msgs.append('csv events in sequence range: <b>' + str(count))
            ui.btn_import.setEnabled(True)
            ui.btn_xml.setEnabled(True)
        
        ui.checkBox_srctc.setEnabled(False)
        ui.checkBox_rectc.setEnabled(False)


def selavid():
    openFilesPath = ''
    fileName, _ = QFileDialog.getOpenFileName(None, "select avid markers exported as txt", '' , "avid marker textfile (*.txt);;All Files (*.*)")
    if fileName:
        prjseqfps = prjdic.get('prjseqfps')
        prjseqoff = prjdic.get('prjseqoff')
        prjseqid = prjdic.get('prjseqid')
        prjseqstart = prjdic.get('prjseqstart')
        prjseqend = prjdic.get('prjseqend')

        print(prjseqstart + ' ' + prjseqend)

        avidpath = fileName
        avidfile = avidpath.split('/')[-1:]
        avidfile = ' '.join(avidfile)
        ui.label_msgs.append('selected avid-marker-textfile:<b> ' + avidfile)

        avidtclist = []
        with open(avidpath) as f:
            file = open(xml_formatted_markers, "w")
            file.write('\t<clipFrameMarkers>\n\t\t<clipFrameMarker>\n\t\t\t<identifier>%s</identifier>\n\t\t\t<counter>%s</counter>\n\t\t\t<frameMarkers>' % (prjseqid, prjseqoff))
            count = 0

            for line in f:
                if line[0]:
                    linecons = re.split(r'\t+', line.rstrip('\n'))
                    print('line: ' + line)
                    avidevent = linecons[1]
                    print('avidevent tc: ' + avidevent)
                    avidname = linecons[4]
                    print('avidname: ' + avidname)
                    avidnoteslen = int(len(linecons))
                    avidnoteslist = []
                    for i in range(avidnoteslen):
                        if i > 0:
                            avidnoteslist.append(linecons[i])
                    print(avidnoteslist)

                    avidtclist.append(avidevent)
                    avidtclistframes = []
                    avidtcfr = Timecode(prjseqfps, avidevent)
                    eventframes = avidtcfr.frames
                    eventframes = eventframes - 1 # remove one frame offset
                    avidtclistframes.append(eventframes)
                    print(avidtclistframes)
                    if  eventframes >= int(prjseqstart) and eventframes <= int(prjseqend):
                        count +=1
                        insert1 ="\n\t\t\t\t<frameMarker>"
                        insert2 ="\n\t\t\t\t\t<frame>%s</frame>" % eventframes #\n
                        insert3 ="\n\t\t\t\t\t<name>%s</name>" % str('avid: ' + avidname)
                        insert4 ="\n\t\t\t\t\t<notes>%s</notes>" % (','.join(avidnoteslist))
                        insert5 ="\n\t\t\t\t</frameMarker>"
                        file = open(xml_formatted_markers, "a+")
                        file.write(str(insert1))
                        file.write(str(insert2))
                        file.write(str(insert3))
                        file.write(str(insert4))
                        file.write(str(insert5))
                        file.close()
                else:
                    print('-> skipped line: ' + line[0] + 'not a valid event entry!' )
                    pass

            file = open(xml_formatted_markers, "a+")
            file.write('\n\t\t\t</frameMarkers>\n\t\t</clipFrameMarker>\n\t</clipFrameMarkers>')
            file.close()
            
        if count == 0:
            ui.btn_import.setEnabled(False)
            ui.btn_xml.setEnabled(False)
            ui.label_msgs.append('avid marker events in sequence range: <b>' + str(count))
        else:
            ui.label_msgs.append(avidfile + ' contains ' + str(len(avidtclist)) + ' entries.')
            ui.label_msgs.append('avid marker events in sequence range: <b>' + str(count))
            print(avidtclist)
            print('avid-events: ' + str(len(avidtclist)))
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
        os.startfile(filename)
    else:
        pass

def inject():
    #make backup of untouched file
    timestamp = time.strftime("%Y-%d-%m-%H-%M-%S")
    prjbkp = str(prjdic.get('prj')) + '-' + timestamp + '.bkp'
    shutil.copyfile(str(prjdic.get('prj')), prjbkp)
    ui.label_msgs.append('backup-file: ' + (str(prjdic.get('prjfile')) + '-' + timestamp + '.bkp'))
    temp = os.path.join(appdata_pfc_marker, 'temp')
    injectmarkers = os.path.join(appdata_pfc_marker, 'xml_formatted_markers.txt')
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
        print(clipFrameMarker)
        for clipFrameMarkerSeq in clipFrameMarker:
            identifiers = clipFrameMarkerSeq.findall('identifier')
            for identifier in identifiers:
                origclipframemarker = identifier.getparent()
                if identifier.text == prjseqid: # if markers present, search our current sequence
                    print('found prjseqid in orig: ' + prjseqid)
                    print('var origclipframemarker ... parent-name in orig: ' + origclipframemarker.tag)
                    origclipframemarker.clear()
                    origclipframemarker.insert(0, injectmarkers)
                else:
                    origclipframemarkersection = origclipframemarker.getparent() # if markers present, but not the current sequence -> add new child object
                    print('var origclipframemarkersection: ' + str(origclipframemarkersection))
                    origclipframemarkersection.append(injectmarkers)
                    pass
                
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
    disablebtn()
    """
    # <clipFrameMarkers/> # only present, when no markers in whole project
    """

def blink():
    ui.logo.setPixmap(QtGui.QPixmap(":/pfc_marker_logo.svg"))

def logochange():
    try:
        ui.logo.setPixmap(QtGui.QPixmap(":/pfc_marker_logo_closed.svg"))
        timer.singleShot(166, blink)
    finally:
        pass

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

def check_internet():
    url = 'http://www.google.com'
    try:
        network = urllib.request.urlopen(url, timeout=1)
    except urllib.request.URLError:
        print('data of %s not retrieved.' %(url))
        return False
    except timeout:
        print('socket timed out for %s!' %(url))
        return False
    else:
        print('internet available')
        return True

def checkupdate():
    if check_internet():
        buildmsi = version
        msi_update = 'https://raw.githubusercontent.com/suminus/pfc_marker/master/msi/current_version'
        
        check = requests.head(msi_update, verify=find_data_file('cacert.pem')).status_code

        if check == 200:
            onlinemsi = urlopen(msi_update).read().decode('utf_8')
            onlinemsi = str(onlinemsi)
            onlinemsi = onlinemsi.replace("\n", "")
            print('found online msi textfile: ' + onlinemsi)

            if onlinemsi > buildmsi:
                ui.menuUpdate.setTitle("Update available!")
                ui.actionUpdate.setEnabled(True)
                ui.label_msgs.setText('checking for updates ... update found.')
                prjdic['onlinemsi'] = onlinemsi
            else:
                ui.label_msgs.setText('checking for updates ... no update found.')
        else:
            print(check)
            ui.menuUpdate.setTitle('')
    else:
        ui.menuUpdate.setTitle('')
        ui.label_msgs.setText('checking for updates ... no internet access.')
        

def update():
    updatecontent = "<p>do you want to download the latest msi-package?</p>" \
                    "<p>pfc_marker will close and start the update automatically.</p>" 

    reply = QMessageBox.question(None, "Update", updatecontent, QMessageBox.Yes | QMessageBox.No )
    if reply == QMessageBox.Yes:

        onlinemsi = prjdic.get('onlinemsi')
        ui.label_msgs.append('downloading update. please wait ...')
        ui.statusbar.showMessage('downloading update. please wait ...')
        msi_updateurl = 'https://github.com/suminus/pfc_marker/blob/master/msi/pfc_marker-{}-amd64.msi?raw=true' .format(onlinemsi)
        print('msi_updateurl: {}' .format(msi_updateurl))
        msi_update = 'pfc_marker-{}-amd64.msi' .format(onlinemsi)
        print('msi: ' + str(msi_update))

        check = requests.head(msi_updateurl, verify=find_data_file('cacert.pem')).status_code
        print('   RET:   ' + str(check))
        if check == 302:
            savepath = str(os.path.join(os.getenv('TEMP'))) + '\\' + msi_update
            ui.progressBar.show()
            ui.statusbar.addPermanentWidget(ui.progressBar)
            ui.progressBar.setValue(0)

            try:
                urllib.request.urlretrieve(msi_updateurl, savepath, downprogress)
                ui.label_msgs.append('download finished.')
                time.sleep(1)
                os.startfile(os.path.join(os.getenv('TEMP'), msi_update))
                exit()
            except Exception:
                ui.label_msgs.append('download of update failed!')
                ui.progressBar.hide()

        else:
            ui.label_msgs.append('updatefile not found!')
            ui.statusbar.showMessage('updatefile not found!')

    elif reply == QMessageBox.No:
        print("No")
    else:
        print("Cancel")

def downprogress(blocknum, blocksize, totalsize):
    readsofar = blocknum * blocksize
    if totalsize > 0:
        percent = readsofar * 100 / totalsize
        ui.progressBar.setValue(int(percent))
        return int(percent)

def find_data_file(filename):
    if getattr(sys, 'frozen', False):
        # The application is frozen
        datadir = os.path.dirname(sys.executable)
    else:
        # The application is not frozen
        # Change this bit to match where you store your data files:
        datadir = os.path.dirname(__file__)
    return os.path.join(datadir, filename)

def exit():
    sys.exit(0)


if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.setWindowTitle('pfc_marker {}' .format(version))


    timer=QtCore.QTimer()
    timer.timeout.connect(logochange)
    timer.start(5000)

    disablebtn()

    ui.actionExit.triggered.connect(exit)
    ui.actionAbout.triggered.connect(about)
    ui.btn_selprj.clicked.connect(selprj)
    ui.btn_seledl.clicked.connect(seledl)
    ui.btn_selcp.clicked.connect(selcp)
    ui.btn_selcsv.clicked.connect(selcsv)
    ui.btn_selavid.clicked.connect(selavid)

    ui.btn_xml.clicked.connect(savexmlmarker)
    ui.checkBox_rectc.stateChanged.connect(checkboxrectc)
    ui.checkBox_srctc.stateChanged.connect(checkboxsrctc)
    ui.btn_import.clicked.connect(inject)
    ui.label_msgs.setText('')

    ui.menuUpdate.setTitle("")
    ui.actionUpdate.triggered.connect(update)
    ui.actionCheckUpdates.triggered.connect(checkupdateconf)
    ui.progressBar.hide()

    if checkupdateinitcfg == 'checkupdates=1':
        ui.label_msgs.setText('checking for updates ...')
        ui.actionCheckUpdates.setChecked(True)
        timer.singleShot(100, checkupdate) # wait some ms to build ui

    else:
        ui.actionCheckUpdates.setChecked(False)
    
    MainWindow.show()
    sys.exit(app.exec_())