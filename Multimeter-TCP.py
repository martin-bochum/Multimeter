#!/usr/bin python3
# -*- coding: utf-8 -*-
'''
    author: Martin Müller Bochum

    DMM Control Software
    Copyright (C)  2021/2022  Martin Müller.

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program (GPL.txt).  If not, see <https://www.gnu.org/licenses/>.
'''
import sys
from sys import argv
import pyvisa as visa
import sched, time
from time import sleep
from datetime import datetime
from numpy import*
import os
from PyQt5 import QtCore, QtGui, QtWidgets, uic
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import socket, errno
from pyqtgraph import PlotWidget
import pyqtgraph as pg
import xlsxwriter
import configparser

'''
# if you want audio high and low beep "Limit ON"
# you get many errors on the console but it works

import pyaudio
import numpy as np
import struct

FS = 44100  #  frames per second, samples per second or sample rate

def play_sound(type, frequency, volume, duration):
   generate_sound(type, frequency, volume, duration)

def generate_sound(type, frequency, volume, duration):
    outbuf = np.random.normal(loc=0, scale=1, size=int(float(duration / 1000.0)*FS))

    if type == "sine":
        dur = int(FS * float(duration / 1000.0))
        theta = 0.0
        incr_theta = frequency * 2 * math.pi / FS # frequency increment normalized for sample rate
        for i in range(dur):
            outbuf[i] = volume * math.sin(theta)
            theta += incr_theta

    p = pyaudio.PyAudio()
    stream = p.open(format=pyaudio.paFloat32, channels=1, rate=FS, output=True)
    data = b''.join(struct.pack('f', samp) for samp in outbuf) # must pack the binary data
    stream.write(data)
    stream.stop_stream()
    stream.close()
    p.terminate()
'''

global ts, rad, einheit, durch, maxi, mini, reading, messungen, x, y, daten, line, p1, otto, otto1, einheit, pen, plot_label, HOST, PORT, SCREEN, LOCAL, ASRL, RS232_PORT, RS232_SPEED, SN_SHOW, s, socket_verbunden, dual_flag, dual, limit_switch, low_fail, up_fail, otto2, otto21, einheit2, reading2, dual_index, dual_index_alt, usb_switch, otto_alt, save_flag, save_timer, save_intervall, led_index, fileName, hugo, nk, wb_row, wb_col, wb_col2, format_date, format_time, workbook, worksheet, func_1, func_2, cold_boot, dmm4095, lower_val, upper_val, save_start, DEBUG, offset_bak, db_bak, rs_flag, messungen_alt, max_graph, xy_counter, resources, oscilloscope, instrument
save_start = int(round(time.time()))
rs_flag = "0"
max_graph = 300
xy_counter = 0
messungen_alt = 0
lower_val = 0.0
upper_val = 0.0
dmm4095 = False
multiplikator = 1
ts = 1
rad = 0
limit_switch= 0
low_fail = 0
up_fail = 0
display_c_1 = '#aaff00'       # light green
display_c_2 = '#00ff7f'       # dark green
display_c_1_b = '#507fab'     # light blue
display_c_2_b = '#395b7a'     # dark blue
dual_flag = 0
dual_index = ['NONe','VOLT','VOLT AC','CURR','CURR AC','FREQ','PER']
dual_index_alt = 0
save_flag = 0
save_timer = 29
save_intervall = 29
led_index = 0
fileName = "dmmLogFile_dummy.txt"
nk = ["{0:+07.1f}", "{0:+07.2f}", "{0:+07.3f}", "{0:+07.4f}", "{0:.1f}", "{0:.2f}", "{0:.3f}", "{0:.4f}", "{0:}"]
wb_row = 0
wb_col = 0
func_1 = ''
func_2 = ''
dual_switch = 0
cold_boot = 1
DEBUG = 0
SCREEN = "default.ui"
LOCAL = 1
SN_SHOW = 1
poll_intervall = 250        # 250ms SOCKET, 250ms USB
offset_bak = 0.0
db_bak = 0
DB_DBM_REF_409x = [50, 75, 93, 110, 124, 125, 135, 150, 250, 300, 500, 600, 800, 900, 1000, 1200, 8000]

# Temp. not working, FIRST switch to VDC (1-10) or Ω (11-14), than switch to Temp.
TEMP_RDT_TYPE = ["KITS90", "NITS90", "EITS90", "JITS90", "TITS90", "SITS90", "RITS90", "BITS90", "W5_26", "W3_25", "PT100", "PT10", "Cu100", "Cu50"]

CAP_409x = ["2nF", "20nF", "200nF", "2µF", "20µF", "200µF", "10mF"]

VDC_4095 = ["600mV", "6V", "60V", "600V", "750V"]
VAC_4095 = ["600mV", "6V", "60V", "600V", "750V"]
ADC_4095 = ["600µA", "6mA", "60mA", "600mA", "6A", "10A"]
AAC_4095 = ["60mA", "600mA", "6A", "10A"]

VDC_4096 = ["200mV", "2V", "20V", "200V", "750V"]
VAC_4096 = ["200mV", "2V", "20V", "200V", "750V"]
ADC_4096 = ["200µA", "2mA", "20mA", "200mA", "2A", "10A"]
AAC_4096 = ["20mA", "200mA", "2A", "10A"]

RES_4095 = ["600Ω", "6KΩ", "60KΩ", "600KΩ", "6MΩ", "60MΩ", "100MΩ"]
RES_4096 = ["200Ω", "2KΩ", "20KΩ", "200KΩ", "2MΩ", "10MΩ", "100MΩ"]

print('for a RS232 connection use \"python3 Multimeter-RS232.py\"')

try:
    output = 'IP='+argv[1]+' on PORT='+str(int(argv[2]))+' DEBUG='+str(int(argv[3]))
    print(output)
    HOST = argv[1]
    PORT = int(argv[2])
    DEBUG = int(argv[3])
except IndexError:
    config = configparser.ConfigParser()
    config.read('multimeter.ini')
    HOST = config['hw_settings']['HOST']
    PORT = config['hw_settings']['PORT']
    SCREEN = config['hw_settings']['SCREEN']
    LOCAL = config['hw_settings']['LOCAL']
    SN_SHOW = config['hw_settings']['SN_SHOW']
    print("multimeter.ini File: HOST=" + HOST + ", PORT=" +  str(PORT) + "\n")

if 'blue' in SCREEN:
    display_c_1 = display_c_1_b
    display_c_2 = display_c_2_b

s = None
socket_verbunden = 0
usb_switch = 0
resources = visa.ResourceManager('@py')
instruments = array(resources.list_resources())
#    visa.log_to_screen() # enable pyvisa debug output to console
for instrument in instruments:
    print("Checking: ", instrument)
    if 'ASRL' not in instrument:
        oscilloscope = resources.open_resource(instrument, write_termination='\n', read_termination='\n', query_delay=0.01)
        oscilloscope.timeout = 5000 # ms
        try:
            identity = oscilloscope.query('*IDN?')
            print("Resource: '" + instrument + "' is \"" + identity + '\"\n')
            if "P4095" in identity or "3041" in identity or "7060" in identity:
                dmm4095 = True
                usb_switch = 1
            if "P4096" in identity or "3051" in identity or "7200" in identity:
                usb_switch = 1
                dmm4095 = False
            poll_intervall = 250
        except visa.VisaIOError:
            oscilloscope = resources.open_resource(instrument, write_termination='\n', read_termination='\n', query_delay=0.01)
            print('No connection to DMM: ' + instrument)
            usb_switch = 0

if usb_switch == 0:
    instrument = 'TCPIP::'+HOST+'::'+str(PORT)+'::SOCKET'
    oscilloscope = resources.open_resource(instrument, write_termination='\n', read_termination='\n', query_delay=0.01)
    oscilloscope.timeout = 5000 # ms
    try:
        identity = oscilloscope.query('*IDN?')
        print("Resource: '" + instrument + "' is " + identity + '\n')
        usb_switch = 2
        poll_intervall = 250
        if "P4095" in identity or "3041" in identity or "7060" in identity:
            dmm4095 = True
        if "P4096" in identity or "3051" in identity or "7200" in identity:
            dmm4095 = False
    except visa.VisaIOError:
        oscilloscope = resources.open_resource(instrument, write_termination='\n', read_termination='\n', query_delay=0.01)
        print('No connection to: ' + instrument)
        usb_switch = 0

# usb_switch = 0        # TEST forced TELNET connection HOST PORT
# dmm4095 = False       # TEST forced P4096
radmax = 10
einheit = ""
x = [0.0]
y = [0]
daten = [[x], [y]]

def ende():
    global save_flag, save_timer, save_intervall, fileName, wb_row, wb_col, format_date, format_time, workbook, worksheet, resources, oscilloscope
    leer = oscilloscope.write('SYSTem:LOCal')
    print ("Window Ende")
    if save_flag == 1:
        save_flag = 0
        workbook.close()

class Ui(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs):
        super(Ui, self).__init__(*args, **kwargs)
        uic.loadUi(SCREEN, self)        # Load the .ui file
        leer = self.tcpip('*IDN?')
        if SN_SHOW == '0':
            idn_text = []
            idn_text = leer.split(',')
            leer = leer.replace(str(idn_text[2]), "xxxxxxx")
        leer = leer.replace(",", " ")
        self.setFixedSize(745, 240)
        if usb_switch == 0:
            self.setWindowTitle(leer + " - " + HOST + ":" + str(PORT) + " TELNET")
        elif usb_switch == 1:
            self.setWindowTitle(leer + " " + instrument)
        elif usb_switch == 2:
            self.setWindowTitle(leer + " " + instrument)
        leer = self.tcpip('SYSTem:VERSion?')
        if dmm4095 == True:
            print ("DMM 4½ Digits 60000 Counts")
            if 'blue' in SCREEN:
                self.setFixedSize(745, 258)
                self.Logo.setProperty("text","PeakTech P4095")
        elif dmm4095 == False:
            print ("DMM 5½ Digits 200000 Counts")
            if 'blue' in SCREEN:
                self.setFixedSize(745, 258)
                self.Logo.setProperty("text","PeakTech P4096")
        print ("SCPI version: " + leer)
        leer = self.tcpip('RATE S')
        print ("RATE: " + str(leer))
        if int(LOCAL) == 0:
            leer = self.tcpip('SYSTem:REMote')
            print ("SYSTem:REMote: " + str(leer))
        leer = self.tcpip('SYSTem:DATE?')
        print ("DMM DATE: " + str(leer))
        leer = self.tcpip('SYSTem:TIME?')
        print ("DMM TIME: " + str(leer))

        self.comboBox.addItem('- OFF -')
        self.comboBox.addItem('Volt DC')
        self.comboBox.addItem('Volt AC')
        self.comboBox.addItem('Curr. DC')
        self.comboBox.addItem('Curr. AC')
        self.comboBox.addItem('Frequency')
        self.comboBox.addItem('Period')
        self.comboBox.setProperty("enabled", "1")
        self.comboBox.currentIndexChanged.connect(self.dualdisplaychange)
        
        self.lineEdit_9.setVisible(False)
        self.DebugText.setVisible(False)
        self.offsetText.setVisible(False)
        self.comboBox_2.addItem('1')
        self.comboBox_2.addItem('2')
        self.comboBox_2.addItem('4')
        self.comboBox_2.addItem('8')
        self.comboBox_2.addItem('15')
        self.comboBox_2.addItem('30')
        self.comboBox_2.addItem('60')
        self.comboBox_2.addItem('120')
        self.comboBox_2.addItem('300')
        self.comboBox_2.setProperty("enabled", "1")
        self.comboBox_2.setCurrentIndex(6)
        self.comboBox_2.currentIndexChanged.connect(self.save_change)
        self.save_widget.setVisible(False)

        for i in range(len(DB_DBM_REF_409x)):
            self.combobox_db.addItem(str(DB_DBM_REF_409x[i]))
        self.combobox_db.setCurrentIndex(0)
        self.combobox_db.currentIndexChanged.connect(self.db_change)
        self.db_widget.setVisible(False)

        self.temp_widget.setVisible(False)
        self.d_widget.setVisible(False)
        self.s_widget.setVisible(False)

        self.frame.setVisible(False)
        
        self.pushButton_1.clicked.connect(self.stat)
        self.pushButton_1.setProperty("text","Statistics ON")
        self.pushButton_6.clicked.connect(self.spannung)
        self.pushButton_12.clicked.connect(self.spannungac)
        self.pushButton_7.clicked.connect(self.strom)
        self.pushButton_13.clicked.connect(self.stromac)
        self.pushButton_8.clicked.connect(self.ohm)
        self.pushButton_14.clicked.connect(self.buzz)
        self.pushButton_15.clicked.connect(self.diode)
        self.pushButton_9.clicked.connect(self.cap)
        self.pushButton_10.clicked.connect(self.freq)
        self.pushButton_16.clicked.connect(self.per)
        self.pushButton_11.clicked.connect(self.hot)
        self.pushButton.clicked.connect(self.limit)
        self.pushButton_2.clicked.connect(self.t_c)
        self.pushButton_3.clicked.connect(self.t_f)
        self.pushButton_4.clicked.connect(self.t_k)
        self.pushButton_5.clicked.connect(self.toggle_dual)
        self.pushButton_5.setEnabled(False)
        self.pushButton_17.clicked.connect(self.rec_start)
        self.dial.valueChanged.connect(self.rad)
        self.dbText.setVisible(False)
        self.actionMesswerte_speichern.triggered.connect(self.save)

        self.actionEnde.triggered.connect(self.quit)
        self.actionMeasure_STOP.triggered.connect(self.run_stop)
        self.actionDebug.triggered.connect(self.debugger)
        if DEBUG == 1:
            self.actionDebug.setChecked(True)
        self.actionRel_Offset.triggered.connect(self.offset_toggle)
        self.actionRESET.triggered.connect(self.reset_on_error)

        self.lineEdit_14.setVisible(False)
        self.lineEdit_15.setVisible(False)
        self.lcdNumber.setFrame(False)
        self.led=[QPixmap('led-off.png'),QPixmap('led-green-on.png'),QPixmap('led-red-on.png')]
        
        self.actionmV_DC_Impedance_20G.triggered.connect(self.imp_toggle)
        
        self.action_4_Wire.triggered.connect(self.w_toggle)
        
        self.actionCurrent_DC_Filter.triggered.connect(self.f_toggle)
        
        self.actionAbout.triggered.connect(self.about)

        offset = self.tcpip('CALCulate:NULL:OFFSet?')
        if offset == '0.000000':
            self.actionRel_Offset.setChecked(False)
        elif offset != '0.000000':
            self.actionRel_Offset.setChecked(True)
        offset_bak = round(float(offset),6)

        rs_flag = leer = self.tcpip('SHOW?')
        if rs_flag == "1":
            self.actionMeasure_STOP.setChecked(False)
            self.timer_led=QTimer()
            self.timer_led.singleShot(50, self.led_off)
            self.timer_single=QTimer()
            self.timer_single.start(500)
            self.timer_single.timeout.connect(self.update)
        elif rs_flag == "0":
            self.actionMeasure_STOP.setChecked(True)
            self.timer_led=QTimer()
            self.timer_led.singleShot(50, self.led_off)
            self.timer_single=QTimer()
            self.timer_single.start(500)
            self.timer_single.timeout.connect(self.update)
            self.label.setPixmap(self.led[2])
        self.show()

    def about(self):
        QMessageBox.about(self, "About", "Owon \t XDM3041, XDM3051\nPeakTech \t P4095, P4096\nVOLTCRAFT \t VC-7060BT, VC-7200BT\n\nUSB/TCP Control Software\nVersion: 1.00\n\nDevelopment and bug reports:\nmartin@martin-bochum.de\n\nCopyright (C)  2021/2022  Martin Müller\nThis program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.\n\nThis program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.\nSee the GNU General Public License for more details.\n\nYou should have received a copy of the GNU General Public\nLicense along with this program (GPL.txt).\nIf not, see <https://www.gnu.org/licenses/>.")

    def f_toggle(self):
        einheitdc = self.tcpip('CURRent:DC:FILTer:STATe?')
#        einheit = self.tcpip('FUNC1?')
        if einheit != "VOLT" and einheit != "CURR":
            return
        if einheitdc == "1":
            leer = self.tcpip('FUNCtion1 \"CURRent:DC\"')
            sleep(.5)
            leer = self.tcpip('CURRent:DC:FILTer:STATe OFF')
            sleep(.5)
            if einheit == "VOLT":
                leer = self.tcpip('FUNCtion1 \"VOLTage:DC\"')
            elif einheit == "CURR":
                leer = self.tcpip('FUNCtion1 \"CURRent:DC\"')
            self.actionCurrent_DC_Filter.setChecked(False)
        elif einheitdc == "0":
            leer = self.tcpip('FUNCtion1 \"CURRent:DC\"')
            sleep(.5)
            leer = self.tcpip('CURRent:DC:FILTer:STATe ON')
            sleep(.5)
            if einheit == "VOLT":
                leer = self.tcpip('FUNCtion1 \"VOLTage:DC\"')
            elif einheit == "CURR":
                leer = self.tcpip('FUNCtion1 \"CURRent:DC\"')
            self.actionCurrent_DC_Filter.setChecked(True)

    def run_stop(self):
        rs_flag = self.tcpip('SHOW?')
        if rs_flag == "1":
            leer = self.tcpip('SHOW OFF')
            self.actionMeasure_STOP.setChecked(True)
            self.led_on(2)
        if rs_flag == "0":
            leer = self.tcpip('SHOW ON')
            self.actionMeasure_STOP.setChecked(False)
            self.timer_single.setInterval(poll_intervall)
            self.timer_single.start()
            self.led_on(2)
            self.timer_led.singleShot(1000, self.led_off)

    def w_toggle(self):
        einheit = self.tcpip('FUNC1?')
        if "FRES" in einheit:
            leer = self.tcpip('CONFigure:RESistance')
            self.action_4_Wire.setChecked(False)
        elif "RES" in einheit:
            leer = self.tcpip('CONFigure:FRESistance')
            self.action_4_Wire.setChecked(True)

    def imp_toggle(self):
        impedance = self.tcpip('VOLTage:DC:IMPedance:AUTO?')
        if int(impedance) == 1:
            self.tcpip('VOLTage:DC:IMPedance:AUTO OFF')
            self.actionmV_DC_Impedance_20G.setChecked(False)
        elif int(impedance) == 0:
            self.tcpip('VOLTage:DC:IMPedance:AUTO ON')
            self.actionmV_DC_Impedance_20G.setChecked(True)

    def db_change(self, dx):
        self.combobox_db.setCurrentIndex(dx)
        self.tcpip('CALCulate:DBM:REFerence ' + str(self.combobox_db.currentText()))
        
    def offset_toggle(self):
      self.limit_off()
      if ts != 0:
        off = self.tcpip('CALCulate:NULL:OFFSet?')
        if off == '0.000000':
            off_txt = self.tcpip('MEAS1?')
            self.tcpip('CALCulate:NULL:OFFSet ' + str(round(float(off_txt),6)))
            self.tcpip('CALCulate:FUNCtion NULL')
        elif off != '0.000000':
            self.tcpip('CALCulate:NULL:OFFSet 0.00')
            self.tcpip('CALCulate:STATe OFF')
      if ts == 0:
          self.stat_off()
          self.offset_toggle()

    def reset_on_error(self):
        global resources, oscilloscope
        del resources
        del oscilloscope
        resources = visa.ResourceManager('@py')
        oscilloscope = resources.open_resource(instrument, write_termination='\n', read_termination='\n', query_delay=0.01)
        del resources
        del oscilloscope
        resources = visa.ResourceManager('@py')
        oscilloscope = resources.open_resource(instrument, write_termination='\n', read_termination='\n', query_delay=0.01)
        self.tcpip('SHOW ON')
        self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.tcpip('CALCulate:STATe OFF')
        self.debugger('Software Reset !')
        leer = self.tcpip('RATE S')
        self.limit_off()
        self.stat_off()
#        leer = self.tcpip('*IDN?')
#        print (leer)
        self.update()
        
    def rec_start(self):
        global save_flag, save_timer, save_intervall, fileName, wb_row, wb_col, format_date, format_time, workbook, worksheet, save_start
        save_flag = 1
        save_timer = int(round(time.time()))
        save_intervall = int(self.comboBox_2.currentText())
        save_start = int(round(time.time())) + save_intervall
        wb_row = 1
        format_date = workbook.add_format({'num_format': 'dd.mm.yyyy'})
        format_time = workbook.add_format({'num_format': 'hh:mm:ss'})
        self.label_2.setPixmap(self.led[2])
        if os.path.isfile(fileName):
            os.remove(fileName)
        self.pushButton_17.setStyleSheet("background:rgb(255,0,0)")
        self.pushButton_17.setProperty("text","Recording...")
        self.pushButton_17.setEnabled(False)
        self.comboBox_2.setEnabled(False)

    def debugger(self, debug_text):
        if self.actionDebug.isChecked() == True:
            self.DebugText.setVisible(True)
            t_string = datetime.now().strftime("%H:%M:%S")
#            debug_text = t_string + " Debug: " + str(debug_text)
            debug_text = t_string + " " + str(debug_text)
            self.DebugText.setProperty("text", debug_text)
        elif self.actionDebug.isChecked() == False:
            self.DebugText.setVisible(False)

    def save_change(self):
        global save_flag, save_timer, save_intervall, fileName, save_start
        save_intervall = int(self.comboBox_2.currentText()) - 1

    def buttons_off(self):
        self.pushButton_6.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_12.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_7.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_13.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_8.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_14.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_15.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_9.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_10.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_16.setStyleSheet("background:rgb(255,255,255)")
        self.pushButton_11.setStyleSheet("background:rgb(255,255,255)")
        
    def save(self):
        global save_flag, save_timer, save_intervall, fileName, wb_row, wb_col, format_date, format_time, workbook, worksheet
        fileName = ""
        if save_flag == 0 and self.actionMesswerte_speichern.isChecked() == True:
            options = QFileDialog.DontUseNativeDialog
            fileName, _ = QFileDialog.getSaveFileName(self,"Log-Datei speichern","dmmLogFile.xlsx","Alle Files (*);;Text Files (*.xlsx)", options=options)
            if fileName:
                self.debugger(fileName)
                workbook = xlsxwriter.Workbook(fileName)
                worksheet = workbook.add_worksheet()
                worksheet.set_column('A:K', 12)
                worksheet.write('A1', 'Datum')
                worksheet.write('B1', 'Zeit')
                worksheet.write('C1', 'V DC (V)')
                worksheet.write('D1', 'V AC (V)')
                worksheet.write('E1', 'A DC (A)')
                worksheet.write('F1', 'A AC (A)')
                worksheet.write('G1', 'Freq. (Hz)')
                worksheet.write('H1', 'Per. (s)')
                worksheet.write('I1', 'Res. (Ω)')
                worksheet.write('J1', 'Cap. (F)')
                worksheet.write('K1', 'Temp. (°C)')
                self.actionMesswerte_speichern.setChecked(True)
                self.save_widget.setVisible(True)
            elif fileName == '':
                self.actionMesswerte_speichern.setChecked(False)
                save_flag = 0
                return 0
        elif save_flag == 1 and self.actionMesswerte_speichern.isChecked() == False:
            save_flag = 0
            workbook.close()
            self.label_2.setPixmap(self.led[0])
            self.save_widget.setVisible(False)
            self.pushButton_17.setStyleSheet("background:rgb(255,255,255)")
            self.pushButton_17.setProperty("text","Rec. Start")
            self.pushButton_17.setEnabled(True)
            self.comboBox_2.setEnabled(True)
        elif save_flag == 0 and self.actionMesswerte_speichern.isChecked() == False:
            save_flag = 0
            workbook.close()
            self.label_2.setPixmap(self.led[0])
            self.save_widget.setVisible(False)
            self.pushButton_17.setStyleSheet("background:rgb(255,255,255)")
            self.pushButton_17.setProperty("text","Rec. Start")
            self.pushButton_17.setEnabled(True)
            self.comboBox_2.setEnabled(True)

    def quit(self):
        global save_flag, save_timer, save_intervall, fileName, wb_row, wb_col, format_date, format_time, workbook, worksheet
        print ("Menü Ende")
        if save_flag == 1:
            save_flag = 0
            workbook.close()
            self.label_2.setPixmap(self.led[0])
            self.save_widget.setVisible(False)
        leer = self.tcpip('SYSTem:LOCal')
        self.close()

    def limit(self):
        global limit_switch, low_fail, up_fail, multiplikator, otto, reading, lower_val, upper_val
        kill_txt = ['V', 'A', 'F', 'Hz', 's', '♪', ' ', 'Ω', '↓', '↑', 'C', 'F', 'K', '°']
        if limit_switch == 0:
          if self.lineEdit_11.text() == "" and self.lineEdit_12.text() == "":
            self.lineEdit_11.setProperty("text", str(round(reading+(reading/100*0.025),4)) + otto)
            self.lineEdit_12.setProperty("text", str(round(reading-(reading/100*0.025),4)) + otto)
#            self.lineEdit_11.setProperty("text", str(round(reading,4)) + otto)
#            self.lineEdit_12.setProperty("text", str(round(reading,4)) + otto)
          if self.lineEdit_11.text() != "" and self.lineEdit_12.text() != "":
            self.lineEdit_14.setVisible(True)
            self.lineEdit_15.setVisible(True)
            self.pushButton.setProperty("text","Limit OFF")
            upper = self.lineEdit_11.text()
            lower = self.lineEdit_12.text()
            upper = upper.replace(',', '.')
            lower = lower.replace(',', '.')
            for i in range(len(kill_txt)):
                upper = upper.replace(kill_txt[i], '')
            for i in range(len(kill_txt)):
                lower = lower.replace(kill_txt[i], '')
            if "G" not in lower and "M" not in lower and "k" not in lower and "m" not in lower and "µ" not in lower and "u" not in lower and "n" not in lower and "p" not in lower:
                lower_val = round(float(lower),4)
            if "G" not in upper and "M" not in upper and "k" not in upper and "m" not in upper and "µ" not in upper and "u" not in upper and "n" not in upper and "p" not in upper:
                upper_val = round(float(upper),4)
            if "G" in lower:
                lower = lower.replace('G', '')
                lower_val = round(float(lower)*1e+9,4)
            if "G" in upper:
                upper = upper.replace('G', '')
                upper_val = round(float(upper)*1e+9,4)
            if "M" in lower:
                lower = lower.replace('M', '')
                lower_val = round(float(lower)*1e+6,4)
            if "M" in upper:
                upper = upper.replace('M', '')
                upper_val = round(float(upper)*1e+6,4)
            if "k" in lower:
                lower = lower.replace('k', '')
                lower_val = round(float(lower)*1e+3,7)
            if "k" in upper:
                upper = upper.replace('k', '')
                upper_val = round(float(upper)*1e+3,7)
            if "m" in lower:
                lower = lower.replace('m', '')
                lower_val = round(float(lower)*1e-3,5)
            if "m" in upper:
                upper = upper.replace('m', '')
                upper_val = round(float(upper)*1e-3,5)
            if "µ" in lower or "u" in lower:
                lower = lower.replace('u', '')
                lower = lower.replace('µ', '')
                lower_val = round(float(lower)*1e-6,7)
            if "µ" in upper or "u" in upper:
                upper = upper.replace('u', '')
                upper = upper.replace('µ', '')
                upper_val = round(float(upper)*1e-6,7)
            if "n" in lower:
                lower = lower.replace('n', '')
                lower_val = round(float(lower)*1e-9,10)
            if "n" in upper:
                upper = upper.replace('n', '')
                upper_val = round(float(upper)*1e-9,10)
            if "p" in lower:
                lower = lower.replace('p', '')
                lower_val = round(float(lower)*1e-12,13)
            if "p" in upper:
                upper = upper.replace('p', '')
                upper_val = round(float(upper)*1e-12,13)
            leer = self.tcpip('CALCulate:LIMit:LOWer '+ str(lower_val))
            leer = self.tcpip('CALCulate:LIMit:UPPer '+ str(upper_val))
            limit_switch = 1
            if ts == 0:
                self.stat_off()
        elif limit_switch == 1:
            self.lineEdit_14.setVisible(False)
            self.lineEdit_15.setVisible(False)
            self.lineEdit_11.setStyleSheet("background-color: #a2a2a2; border: 1px solid black;")
            self.lineEdit_12.setStyleSheet("background-color: #a2a2a2; border: 1px solid black;")
            leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
            leer = self.tcpip('CALCulate:STATe OFF')
            limit_switch = 0
            low_fail = 0
            up_fail = 0
            self.pushButton.setProperty("text","Limit ON")
            self.lineEdit_11.setProperty("text", "")
            self.lineEdit_12.setProperty("text", "")    
            self.lineEdit_11.setProperty("placeholderText", "Upper Limit")
            self.lineEdit_12.setProperty("placeholderText", "Lower Limit")

    def limit_off(self):
        global limit_switch, low_fail, up_fail
        self.lineEdit_14.setVisible(False)
        self.lineEdit_15.setVisible(False)
        self.lineEdit_11.setStyleSheet("background-color: #a2a2a2; border: 1px solid black;")
        self.lineEdit_12.setStyleSheet("background-color: #a2a2a2; border: 1px solid black;")
        leer = self.tcpip('CALCulate:STATe OFF')
        limit_switch = 0
        low_fail = 0
        up_fail = 0
        self.pushButton.setProperty("text","Limit ON")
        self.lineEdit_11.setProperty("text", "")
        self.lineEdit_12.setProperty("text", "")    
        self.lineEdit_11.setProperty("placeholderText", "Upper Limit")
        self.lineEdit_12.setProperty("placeholderText", "Lower Limit")

    def toggle_dual(self, ix):
        global dual_flag, dual_index_alt, dual_index, einheit, einheit2, func_1, func_2, dual_switch, cold_boot
        self.stat_off()
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        func_1 = self.tcpip('FUNCtion1?')
        func_2 = self.tcpip('FUNCtion2?')
        if dual_flag == 1:
            if func_1 == "VOLT":
                func_1 = "VOLT:DC"
            if func_2 == "VOLT":
                func_2 = "VOLT:DC"
            if func_1 == "CURR":
                func_1 = "CURR:DC"
            if func_2 == "CURR":
                func_2 = "CURR:DC"
            func_1 = 'FUNCtion2 \"'+func_1.replace(' ', ':')+'\"'
            func_2 = 'CONFigure:'+func_2.replace(' ', ':')
#            print (func_2+' <===> '+func_1)
            self.debugger(func_2+' <===> '+func_1)
            leer = self.tcpip('FUNCtion2 \"NONe\"')
            sleep (0.2)
            leer = self.tcpip(func_2)
            sleep (0.2)
            leer = self.tcpip(func_1)
            dual_switch = 1

    def dualdisplaychange(self, ix):
        global dual_flag, dual_index_alt, dual_index, einheit2, func_1, func_2, dual_switch, cold_boot
        self.stat_off()
        self.limit_off()
        d1 = "Dual_Switch: " + str(dual_switch)
        d2 = " - Dual_Flag: "+ str(dual_flag)
        d3 = " - Index Alt : " + str(dual_index_alt)
        d4 = " - Index ix: " + str(ix)
        self.debugger(d1 + d2 + d3 + d4)
        if ix == 0 and dual_switch == 0:
            leer = self.tcpip('FUNCtion2 \"NONe\"')
            self.lineEdit_9.setProperty("text", "DUAL OFF")
            self.lineEdit_9.setVisible(False)
            dual_flag = 0
            self.timer_single.setInterval(1000)
        if dual_index_alt != ix:
            self.debugger("Dual: " + einheit2 + " -> " + self.comboBox.currentText())
            if ix == 0:
                leer = self.tcpip('FUNCtion2 \"NONe\"')
                dual_flag = 0
                self.timer_single.setInterval(1000)
            if ix == 1:
                self.tcpip('FUNCtion2 \"VOLT:DC\"')
                dual_flag = 1
            elif ix == 2:
                self.tcpip('FUNCtion2 \"VOLT:AC\"')
                dual_flag = 1
            elif ix == 3:
                self.tcpip('FUNCtion2 \"CURR:DC\"')
                dual_flag = 1
            elif ix == 4:
                self.tcpip('FUNCtion2 \"CURR:AC\"')
                dual_flag = 1
            elif ix == 5:
                self.tcpip('FUNCtion2 \"FREQuency\"')
                dual_flag = 1
            elif ix == 6:
                self.tcpip('FUNCtion2 \"PERiod\"')
                dual_flag = 1

    def tcpip(self, text_out):
        global HOST, PORT, s, usb_switch, oscilloscope, resources, socket_verbunden
        data = '0'
        if usb_switch == 0:
            for res in socket.getaddrinfo(HOST, PORT, socket.AF_UNSPEC, socket.SOCK_STREAM):
                af, socktype, proto, canonname, sa = res
                try:
                    s = socket.socket(af, socktype, proto)
                except OSError as msg:
                    s = None
                    continue
                try:
                    s.connect(sa)
                except OSError as msg:
                    s.close()
                    s = None
                    continue
                break
            if s is None:
                print('could not open socket')
                sys.exit(1)
            with s:
                data = "sende..."
                if "?" in text_out:
                    s.sendall(str.encode(text_out))
                    data = s.recv(8196)
                    data = data.decode('utf-8')
                    data = data.replace('"', '')
                    data = data.strip('\n')
                    data = data.rstrip()
                elif " " in text_out or "CONT" in text_out or "DIOD" in text_out or "TEMP" in text_out or "AUTO" in text_out or "RESet" in text_out or "RATE" in text_out or "CONF" in text_out or "SYST" in text_out:
                    text_out = text_out
                    s.sendall(str.encode(text_out))
            return (data)
        elif usb_switch == 1 or usb_switch == 2:
            if "?" in text_out:
                try:
                    data = oscilloscope.query(text_out)
                    data = data.replace('"', '')
                    data = data.strip('\n')
                except visa.errors.VisaIOError:
                    self.timer_single.stop()
                    self.debugger('No connection to: ' + instrument)
#                    self.reset_on_error()
                    del resources
                    del oscilloscope
                    resources = visa.ResourceManager('@py')
                    oscilloscope = resources.open_resource(instrument, write_termination='\n', read_termination='\n', query_delay=0.01)
                    self.debugger('RESTART query '+text_out)
                    data = oscilloscope.query(text_out)
                    data = data.replace('"', '')
                    data = data.strip('\n')
                    return (data)
#                    self.update()

            elif " " in text_out or "CONT" in text_out or "DIOD" in text_out or "TEMP" in text_out or "AUTO" in text_out or "RESet" in text_out or "RATE" in text_out or "CONF" in text_out or "SYST" in text_out:
                try:
                    data = ''
                    data = oscilloscope.write(text_out)
                    val = (byte(data)) & 1
#                    print (val)
                except visa.errors.VisaIOError:
                    self.timer_single.stop()
                    self.debugger('No connection to: ' + instrument)
#                    self.reset_on_error()
                    del resources
                    del oscilloscope
                    resources = visa.ResourceManager('@py')
                    oscilloscope = resources.open_resource(instrument, write_termination='\n', read_termination='\n', query_delay=0.01)
                    self.debugger('RESTART write '+text_out)
                    data = ''
                    data = oscilloscope.query(text_out)
                    return (data)
#                    self.update()
            return (data)
    
    def rad(self):
        global rad, einheit
        self.dial.setNotchesVisible(True)
        rad = self.dial.value()
        if int(rad) == 0:
            leer = self.tcpip('AUTO')
        elif int(rad >= 1):
            leer = self.tcpip('RANGE1 ' + str(rad))
        if "TEMP" in einheit:
            if int(rad) <= 10:
                leer = self.tcpip('CONFigure:VOLTage:DC:RANG AUTO')
                text_out = "V ThermoCouple"
            elif int(rad) > 10:
                leer = self.tcpip('CONFigure:RESistance:RANG AUTO')
                text_out = "Ω ThermoResistor"
            sleep(0.01)
            leer = self.tcpip('CONFigure:TEMPerature')
            sleep(0.01)
            temptype = self.tcpip('TEMP:RTD:TYP?')
            if temptype != '':
                temp_type_val = TEMP_RDT_TYPE.index(temptype) + 1
                self.dial.setValue(temp_type_val)
                leer = self.tcpip('CONFigure:TEMPerature' + temptype)
                leer = self.tcpip('TEMPerature:RTD:TYPe ' + temptype)
            self.debugger(str(rad) + ' - ' + temptype + ' - '  + text_out)

    def stat(self):
        global ts, x, y, einheit, pen, plot_label, otto, otto1, otto_alt, max_graph, xy_counter, messungen_alt
        if int(ts) == 1:
          if "CONT" not in einheit and "DIOD" not in einheit:
            if limit_switch == 1:
              self.limit_off()
            self.setFixedSize(745, 610)
            if 'blue' in SCREEN:
                self.setFixedSize(745, 628)
            off = self.tcpip('CALCulate:NULL:OFFSet?')
            if off != '0.000000':
                self.tcpip('CALCulate:STATe OFF')
                self.tcpip('CALCulate:NULL:OFFSet 0.00')
                self.tcpip('CALCulate:FUNCtion AVERage')
            elif off == '0.000000':
                self.tcpip('CALCulate:FUNCtion AVERage')
            
            self.pushButton_1.setProperty("text","Statistics OFF")
            messungen = self.tcpip('CALCulate:AVERage:COUNt?')
            messungen = round(float(messungen),0)
            messungen_alt = messungen
            self.graphWidget.clear()
            xy_counter = 1
            del y[0:-1]
            del x[0:-1]
            y[0]=int(messungen)
            x[0]=reading
#            y.append(0)
#            x.append(reading)
            pen = pg.mkPen(color=(0, 0, 0), width=2)
            self.graphWidget.setBackground((170, 255, 0))
            self.graphWidget.setTitle(plot_label, color="b", size="10pt")
            styles = {"color": "#000", "font-size": "10px"}
            if 'black' in SCREEN:
                pen = pg.mkPen(color=(255, 255, 0), width=2)
                self.graphWidget.setBackground((0, 0, 0))
                self.graphWidget.setTitle(plot_label, color="w", size="10pt")
                styles = {"color": "#fff", "font-size": "10px"}
            if 'blue' in SCREEN:
                pen = pg.mkPen(color=(255, 255, 0), width=2)
                self.graphWidget.setBackground((80, 127, 171))
                self.graphWidget.setTitle(plot_label, color="w", size="10pt")
                styles = {"color": "#fff", "font-size": "12px"}
#            labelStyle = {'color': '#000')
            self.graphWidget.setLabel("left", otto + " " + otto1, **styles)
            self.graphWidget.setLabel("bottom", "Measurements", **styles)
#            self.graphWidget.addLegend()
            self.graphWidget.showGrid(x=True, y=True, alpha=1.0)
            self.graphWidget.setXRange(0, 1, padding=0)
#           self.graphWidget.setYRange(20, 55, padding=0)
            self.graphWidget.enableAutoRange()
            self.graphWidget.hideButtons()
            otto_alt = otto
            ts = 0
#            if messungen > max_graph:
#                self.stat_off()
        elif int(ts) == 0:
            self.setFixedSize(745, 240)
            if 'blue' in SCREEN:
                self.setFixedSize(745, 258)
            leer = self.tcpip('CALCulate:STATe OFF')
            self.pushButton_1.setProperty("text","Statistics ON")
            self.lineEdit_3.setProperty("text", "")
            self.lineEdit_5.setProperty("text", "")
            self.lineEdit_7.setProperty("text", "")
            self.lineEdit_8.setProperty("text", "")
            self.lineEdit_4.setProperty("text", "")
            self.lineEdit_6.setProperty("text", "")
            ts = 1
            messungen_alt = 0

    def stat_off(self):
        global ts, messungen_alt
        self.setFixedSize(745, 240)
        if 'blue' in SCREEN:
            self.setFixedSize(745, 258)
        leer = self.tcpip('CALCulate:STATe OFF')
        self.pushButton_1.setProperty("text","Statistics ON")
        self.lineEdit_3.setProperty("text", "")
        self.lineEdit_5.setProperty("text", "")
        self.lineEdit_7.setProperty("text", "")
        self.lineEdit_8.setProperty("text", "")
        self.lineEdit_4.setProperty("text", "")
        self.lineEdit_6.setProperty("text", "")
        ts = 1
        messungen_alt = 0

    def freq(self):
        leer = self.tcpip('FUNCtion2 \"NONe\"')
        self.dbText.setVisible(False)
        self.db_widget.setVisible(False)
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.dial.setValue(2)
        self.limit_off()
        self.stat_off()
#        leer = self.tcpip('CONFigure:FREQuency:RANG AUTO')
        leer = self.tcpip('CONFigure:FREQuency')
        leer = self.tcpip('RANGE1 2')

    def per(self):
        leer = self.tcpip('FUNCtion2 \"NONe\"')
        self.dbText.setVisible(False)
        self.db_widget.setVisible(False)
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.dial.setValue(2)
        self.limit_off()
        self.stat_off()
#        leer = self.tcpip('CONFigure:PERiod:RANG AUTO')
        leer = self.tcpip('CONFigure:PERiod')
        leer = self.tcpip('RANGE1 2')

    def hot(self):
        self.dbText.setVisible(False)
        self.db_widget.setVisible(False)
#        self.dial.setValue(0)
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.limit_off()
        self.stat_off()
        leer = self.tcpip('CONFigure:TEMPerature')
        leer = self.tcpip('TEMPerature:RTD:SHOW ALL')

    def t_c(self):
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.dial.setValue(1)
        self.limit_off()
        self.stat_off()
        leer = self.tcpip('TEMPerature:RTD:UNIT C')

    def t_f(self):
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.dial.setValue(1)
        self.limit_off()
        self.stat_off()
        leer = self.tcpip('TEMPerature:RTD:UNIT F')

    def t_k(self):
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.dial.setValue(1)
        self.limit_off()
        self.stat_off()
        leer = self.tcpip('TEMPerature:RTD:UNIT K')

    def spannung(self):
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        leer = self.tcpip('FUNCtion2 \"NONe\"')
        self.dial.setValue(0)
        self.limit_off()
        self.stat_off()
#        leer = self.tcpip('CONFigure:VOLTage:DC:RANG AUTO')
        leer = self.tcpip('CONFigure:VOLTage:DC')
        
    def spannungac(self):
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        leer = self.tcpip('FUNCtion2 \"NONe\"')
        self.dial.setValue(0)
        self.limit_off()
        self.stat_off()
#        leer = self.tcpip('CONFigure:VOLTage:AC:RANG AUTO')
        leer = self.tcpip('CONFigure:VOLTage:AC')
#        self.reset_on_error()

    def strom(self):
        self.dbText.setVisible(False)
        self.db_widget.setVisible(False)
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        leer = self.tcpip('FUNCtion2 \"NONe\"')
        self.dial.setValue(0)
        self.limit_off()
        self.stat_off()
#        leer = self.tcpip('CONFigure:CURRent:DC:RANG AUTO')
        leer = self.tcpip('CONFigure:CURRent:DC')

    def stromac(self):
        self.dbText.setVisible(False)
        self.db_widget.setVisible(False)
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        leer = self.tcpip('FUNCtion2 \"NONe\"')
        self.dial.setValue(0)
        self.limit_off()
        self.stat_off()
#        leer = self.tcpip('CONFigure:CURRent:AC:RANG AUTO')
        leer = self.tcpip('CONFigure:CURRent:AC')

    def cap(self):
        self.dbText.setVisible(False)
        self.db_widget.setVisible(False)
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.limit_off()
        self.stat_off()
#        leer = self.tcpip('CONFigure:CAPacitance:RANG AUTO')
        leer = self.tcpip('CONFigure:CAPacitance')

    def ohm(self):
        self.dbText.setVisible(False)
        self.db_widget.setVisible(False)
        self.dial.setValue(0)
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        leer = self.tcpip('CONFigure:RESistance')
        self.limit_off()
        self.stat_off()
#        leer = self.tcpip('CONFigure:RESistance:RANG AUTO')

    def buzz(self):
        self.dbText.setVisible(False)
        self.db_widget.setVisible(False)
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.dial.setValue(4)
        self.limit_off()
        self.stat_off()
        leer = self.tcpip('CONFigure:CONTinuity')

    def diode(self):
        self.dbText.setVisible(False)
        self.db_widget.setVisible(False)
        leer = self.tcpip('CALCulate:NULL:OFFSet 0.00')
        self.dial.setValue(0)
        self.limit_off()
        self.stat_off()
        leer = self.tcpip('CONFigure:DIODe')

    def led_off(self):
#        self.frame.setVisible(True)
        self.label.setPixmap(self.led[0])

    def led_on(self, color):
#        self.frame.setVisible(True)
        self.label.setPixmap(self.led[color])

    def update(self):
      global ts, rad, einheit, durch, maxi, mini, reading, messungen, x, y, daten, line, p1, otto, otto1, einheit, pen, plot_label, HOST, PORT, s, dual_flag, dual, limit_switch, low_fail, up_fail, otto2, otto21, einheit2, reading2, dual_index, dual_index_alt, usb_switch, otto_alt, save_flag, save_timer, save_intervall, led_index, fileName, hugo, nk, wb_row, wb_col, wb_col2, format_date, format_time, workbook, worksheet, func_1, func_2, cold_boot, dmm4095, lower_val, upper_val, save_start, offset_bak, db_bak, messungen_alt, max_graph, xy_counter
      try:
        self.buttons_off()
        self.timer_single.stop()
        self.frame.setVisible(True)
        self.led_on(1)
#        self.db_widget.setVisible(False)
        multiplikator = 1
        readinglines = self.tcpip('MEAS1?')
        self.led_on(1)
        einheit = self.tcpip('FUNC1?')
        self.led_on(1)
        bereich = self.tcpip('AUTO1?')
        self.dial.setNotchesVisible(True)
        offset = self.tcpip('CALCulate:NULL:OFFSet?')
        nachkomma = 1
        if "TEMP" in einheit:
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_11.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(False)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(True)
            self.comboBox.setProperty("enabled", "0")
            plot_label = "Temperature"
            self.dial.setRange(1,14)
            self.dial.setProperty("toolTip", "Sensor Selection " + str(TEMP_RDT_TYPE))
            self.dial.setNotchesVisible(True)
            temptype = self.tcpip('TEMP:RTD:TYP?')
            self.small_1.setProperty("text", temptype)
            otto1 = "Sensor "+temptype
            temp_type_val = TEMP_RDT_TYPE.index(temptype) + 1
            self.dial.setValue(temp_type_val)
            otto = self.tcpip('TEMP:RTD:UNIT?')
            if "K" not in otto:
                otto = "°" + otto
            nachkomma = 6
            reading = round((float(readinglines)),2)
            if round(float(offset),6) != 0.000000:
                self.offsetText.setVisible(True)
                self.offsetText.setProperty("text", "Rel. 0 Offset = " + str(round(float(offset)*multiplikator,6)) + " " + otto)
                reading = round((float(readinglines)-round(float(offset)*multiplikator,6)),2)
            wb_col = 10
            if save_flag == 1:
                worksheet.write('K1', 'Temp. ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('K1', 'Temp. ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')
        elif "VOLT AC" in einheit:
            self.pushButton_12.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(True)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "1")
            plot_label = "Voltage AC"
            self.dial.setRange(0,5)
            if dmm4095 == True:
                self.dial.setProperty("toolTip", "Range Auto " + str(VAC_4095))
            if dmm4095 == False:
                self.dial.setProperty("toolTip", "Range Auto " + str(VAC_4096))
            self.dial.setNotchesVisible(True)
            otto1 = "MANUAL"
            otto = self.tcpip('RANGE1?')
            if bereich != "1":
                if dmm4095 == True:
                    self.small_1.setProperty("text", VAC_4095[int(otto)-1])
                elif dmm4095 == False:
                    self.small_1.setProperty("text", VAC_4096[int(otto)-1])
                self.dial.setValue(int(otto))
            if bereich == "1":
                otto1 = "AUTO"
                self.small_1.setProperty("text", "AUTO")
                self.dial.setValue(0)
            if int(otto) == 1:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "mV"
                multiplikator = 1000
                nachkomma = 6
            elif int(otto) == 2:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "V"
                multiplikator = 1
                nachkomma = 8
            elif int(otto) == 3:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "V"
                multiplikator = 1
                nachkomma = 7
            elif int(otto) == 4:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "V"
                multiplikator = 1
                nachkomma = 6
            elif int(otto) == 5:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "V"
                multiplikator = 1
                nachkomma = 6
            reading = round(float(readinglines)*multiplikator,nachkomma)
            wb_col = 3
            if save_flag == 1:
                worksheet.write('D1', 'V AC ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('D1', 'V AC ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')
            if self.actiondB_dBm.isChecked() == False:
                self.dbText.setVisible(False)
                self.db_widget.setVisible(False)
            if self.actiondB_dBm.isChecked() == True:
                ref_ohm = self.tcpip('CALCulate:DBM:REFerence?')
                set_index = self.combobox_db.findText(str(int(ref_ohm)))
                zw = ((reading/multiplikator)**2)/(float(ref_ohm)*0.001)
                if db_bak != set_index:
                    self.combobox_db.setCurrentIndex(set_index)
                    db_bak = set_index
                if zw != 0.0:
                    self.db_widget.setVisible(True)
                    zw = ((reading/multiplikator)**2)/(float(ref_ohm)*0.001)
                    db = round(10*log10(zw),2)
                    self.dbText.setVisible(True)
                    self.dbText.setProperty("text", str(db) + 'dBm ' + str(int(ref_ohm)) + 'Ω')
                elif zw == 0.0:
                    self.dbText.setVisible(False)
            if reading >= 1e09:
                reading = 888888888
        elif "VOLT" in einheit:
            self.pushButton_6.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(True)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "1")
            plot_label = "Voltage DC"
            rad = 0
            self.dial.setRange(0,5)
            if dmm4095 == True:
                self.dial.setProperty("toolTip", "Range Auto " + str(VDC_4095))
            if dmm4095 == False:
                self.dial.setProperty("toolTip", "Range Auto " + str(VDC_4096))
            self.dial.setNotchesVisible(True)
            impedance = self.tcpip('VOLTage:DC:IMPedance:AUTO?')
            filter = self.tcpip('CURRent:DC:FILTer:STATe?')
            if filter == "1":
                self.actionCurrent_DC_Filter.setChecked(True)
            elif filter == "0":
                self.actionCurrent_DC_Filter.setChecked(False)
            otto1 = "MANUAL"
            otto = self.tcpip('RANGE1?')
            if bereich != "1":
                if dmm4095 == True:
                    self.small_1.setProperty("text", VDC_4095[int(otto)-1])
                elif dmm4095 == False:
                    self.small_1.setProperty("text", VDC_4096[int(otto)-1])
                self.dial.setValue(int(otto))
            if bereich == "1":
                otto1 = "AUTO"
                self.small_1.setProperty("text", "AUTO")
                self.dial.setValue(0)
            if int(otto) == 1:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VDC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VDC_4096[int(otto)-1]
                otto = "mV"
                multiplikator = 1000
                if impedance == '0':
                    imp_str = "10MΩ"
                    self.actionmV_DC_Impedance_20G.setChecked(False)
                elif impedance == '1':
                    imp_str = "10GΩ"
                    self.actionmV_DC_Impedance_20G.setChecked(True)
                otto1 = imp_str + " " + otto1
                nachkomma = 2
            elif int(otto) == 2:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VDC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VDC_4096[int(otto)-1]
                otto = "V"
                multiplikator = 1
                nachkomma = 4
            elif int(otto) == 3:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VDC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VDC_4096[int(otto)-1]
                otto = "V"
                multiplikator = 1
                nachkomma = 3
            elif int(otto) == 4:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VDC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VDC_4096[int(otto)-1]
                otto = "V"
                multiplikator = 1
                nachkomma = 2
            elif int(otto) == 5:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VDC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VDC_4096[int(otto)-1]
                otto = "V"
                multiplikator = 1
                nachkomma = 1
            reading = round(float(readinglines)*multiplikator,nachkomma)
            wb_col = 2
            if save_flag == 1:
                worksheet.write('C1', 'V DC ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('C1', 'V DC ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')
            if self.actiondB_dBm.isChecked() == False:
                self.dbText.setVisible(False)
                self.db_widget.setVisible(False)
            if self.actiondB_dBm.isChecked() == True:
                ref_ohm = self.tcpip('CALCulate:DBM:REFerence?')
                set_index = self.combobox_db.findText(str(int(ref_ohm)))
                zw = ((reading/multiplikator)**2)/(float(ref_ohm)*0.001)
                if db_bak != set_index:
                    self.combobox_db.setCurrentIndex(set_index)
                    db_bak = set_index
                if zw != 0.0:
                    self.db_widget.setVisible(True)
                    zw = ((reading/multiplikator)**2)/(float(ref_ohm)*0.001)
                    db = round(10*log10(zw),2)
                    self.dbText.setVisible(True)
                    self.dbText.setProperty("text", str(db) + 'dBm ' + str(int(ref_ohm)) + 'Ω')
                elif zw == 0.0:
                    self.dbText.setVisible(False)
        elif "CURR AC" in einheit:
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_13.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(True)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "1")
            plot_label = "Current AC"
            self.dial.setRange(0,4)
            if dmm4095 == True:
                self.dial.setProperty("toolTip", "Range Auto " + str(AAC_4095))
            if dmm4095 == False:
                self.dial.setProperty("toolTip", "Range Auto " + str(AAC_4096))
            self.dial.setNotchesVisible(True)
            otto1 = "MANUAL"
            otto = self.tcpip('RANGE1?')
            if bereich != "1":
                if dmm4095 == True:
                    self.small_1.setProperty("text", AAC_4095[int(otto)-1])
                elif dmm4095 == False:
                    self.small_1.setProperty("text", AAC_4096[int(otto)-1])
                self.dial.setValue(int(otto))
            if bereich == "1":
                otto1 = "AUTO"
                self.small_1.setProperty("text", "AUTO")
                self.dial.setValue(0)
            if int(otto) == 1:
                if dmm4095 == True:
                    otto1 = otto1 + " " + AAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + AAC_4096[int(otto)-1]
                otto = "mA"
                multiplikator = 1000
                nachkomma = 7
            elif int(otto) == 2:
                if dmm4095 == True:
                    otto1 = otto1 + " " + AAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + AAC_4096[int(otto)-1]
                otto = "mA"
                multiplikator = 1000
                nachkomma = 6
            elif int(otto) == 3:
                if dmm4095 == True:
                    otto1 = otto1 + " " + AAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + AAC_4096[int(otto)-1]
                otto = "A"
                multiplikator = 1
                nachkomma = 8
            elif int(otto) == 4:
                if dmm4095 == True:
                    otto1 = otto1 + " " + AAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + AAC_4096[int(otto)-1]
                otto = "A"
                multiplikator = 1
                nachkomma = 7
            reading = round(float(readinglines)*multiplikator,nachkomma)
            wb_col = 5
            if save_flag == 1:
                worksheet.write('F1', 'A AC ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('F1', 'A AC ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')
        elif "CURR" in einheit:
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_7.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(True)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "1")
            plot_label = "Current DC"
            self.dial.setRange(0,6)
            if dmm4095 == True:
                self.dial.setProperty("toolTip", "Range Auto " + str(ADC_4095))
            if dmm4095 == False:
                self.dial.setProperty("toolTip", "Range Auto " + str(ADC_4096))
            self.dial.setNotchesVisible(True)
            filter = self.tcpip('CURRent:FILTer:STATe?')
            if filter == "1":
                self.actionCurrent_DC_Filter.setChecked(True)
            elif filter == "0":
                self.actionCurrent_DC_Filter.setChecked(False)
            otto1 = "MANUAL"
            otto = self.tcpip('RANGE1?')
            if bereich != "1":
                if dmm4095 == True:
                    self.small_1.setProperty("text", ADC_4095[int(otto)-1])
                elif dmm4095 == False:
                    self.small_1.setProperty("text", ADC_4096[int(otto)-1])
                self.dial.setValue(int(otto))
            if bereich == "1":
                otto1 = "AUTO"
                self.small_1.setProperty("text", "AUTO")
                self.dial.setValue(0)
            if int(otto) == 1:
                if dmm4095 == True:
                    otto1 = otto1 + " " + ADC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + ADC_4096[int(otto)-1]
                otto = "µA"
                multiplikator = 1000000
                nachkomma = 2
            elif int(otto) == 2:
                if dmm4095 == True:
                    otto1 = otto1 + " " + ADC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + ADC_4096[int(otto)-1]
                otto = "mA"
                multiplikator = 1000
                nachkomma = 4
            elif int(otto) == 3:
                if dmm4095 == True:
                    otto1 = otto1 + " " + ADC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + ADC_4096[int(otto)-1]
                otto = "mA"
                multiplikator = 1000
                nachkomma = 3
            elif int(otto) == 4:
                if dmm4095 == True:
                    otto1 = otto1 + " " + ADC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + ADC_4096[int(otto)-1]
                otto = "mA"
                multiplikator = 1000
                nachkomma = 2
            elif int(otto) == 5:
                if dmm4095 == True:
                    otto1 = otto1 + " " + ADC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + ADC_4096[int(otto)-1]
                otto = "A"
                multiplikator = 1
                nachkomma = 4
            elif int(otto) == 6:
                if dmm4095 == True:
                    otto1 = otto1 + " " + ADC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + ADC_4096[int(otto)-1]
                otto = "A"
                multiplikator = 1
                nachkomma = 3
            reading = round(float(readinglines)*multiplikator,nachkomma)
            wb_col = 4
            if save_flag == 1:
                worksheet.write('E1', 'A DC ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('E1', 'A DC ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')
        elif "FRES" in einheit:
            self.action_4_Wire.setChecked(True)
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_8.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(False)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "0")
            plot_label = "Resistor"
            self.dial.setRange(0,7)
            if dmm4095 == True:
                self.dial.setProperty("toolTip", "Range Auto " + str(RES_4095))
            if dmm4095 == False:
                self.dial.setProperty("toolTip", "Range Auto " + str(RES_4096))
            self.dial.setNotchesVisible(True)
            otto1 = "4W MANUAL"
            otto = self.tcpip('RANGE1?')
            if bereich != "1":
                if dmm4095 == True:
                    self.small_1.setProperty("text", RES_4095[int(otto)-1])
                elif dmm4095 == False:
                    self.small_1.setProperty("text", RES_4096[int(otto)-1])
                self.dial.setValue(int(otto))
            if bereich == "1":
                otto1 = "4W AUTO"
                self.small_1.setProperty("text", "AUTO")
                self.dial.setValue(0)
            if int(otto) == 1:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "Ω"
                multiplikator = 1
                nachkomma = 6
            elif int(otto) == 2:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "kΩ"
                multiplikator = 0.001
                nachkomma = 8
            elif int(otto) == 3:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "kΩ"
                multiplikator = 0.001
                nachkomma = 7
            elif int(otto) == 4:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "kΩ"
                multiplikator = 0.001
                nachkomma = 6
            elif int(otto) == 5:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "MΩ"
                multiplikator = 0.000001
                nachkomma = 8
            elif int(otto) == 6:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "MΩ"
                multiplikator = 0.000001
                nachkomma = 7
            elif int(otto) == 7:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "MΩ"
                multiplikator = 0.000001
                nachkomma = 6
            check = round(float(readinglines)*multiplikator,3)
            reading = round(float(readinglines)*multiplikator,nachkomma)
            wb_col = 8
            if save_flag == 1:
                worksheet.write('I1', 'Res. ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('I1', 'Res. ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')
            if check >= 1000:
                reading = 88888888
#                otto1 = "Eingang offen"
        elif "RES" in einheit:
            self.action_4_Wire.setChecked(False)
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_8.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(False)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "0")
            plot_label = "Resistor"
            self.dial.setRange(0,7)
            if dmm4095 == True:
                self.dial.setProperty("toolTip", "Range Auto " + str(RES_4095))
            if dmm4095 == False:
                self.dial.setProperty("toolTip", "Range Auto " + str(RES_4096))
            self.dial.setNotchesVisible(True)
            otto1 = "2W MANUAL"
            otto = self.tcpip('RANGE1?')
            if bereich != "1":
                if dmm4095 == True:
                    self.small_1.setProperty("text", RES_4095[int(otto)-1])
                elif dmm4095 == False:
                    self.small_1.setProperty("text", RES_4096[int(otto)-1])
                self.dial.setValue(int(otto))
            if bereich == "1":
                otto1 = "2W AUTO"
                self.small_1.setProperty("text", "AUTO")
                self.dial.setValue(0)
            if int(otto) == 1:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "Ω"
                multiplikator = 1
                nachkomma = 6
            elif int(otto) == 2:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "kΩ"
                multiplikator = 0.001
                nachkomma = 8
            elif int(otto) == 3:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "kΩ"
                multiplikator = 0.001
                nachkomma = 7
            elif int(otto) == 4:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "kΩ"
                multiplikator = 0.001
                nachkomma = 6
            elif int(otto) == 5:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "MΩ"
                multiplikator = 0.000001
                nachkomma = 8
            elif int(otto) == 6:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "MΩ"
                multiplikator = 0.000001
                nachkomma = 7
            elif int(otto) == 7:
                if dmm4095 == True:
                    otto1 = otto1 + " " + RES_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + RES_4096[int(otto)-1]
                otto = "MΩ"
                multiplikator = 0.000001
                nachkomma = 6
            check = round(float(readinglines)*multiplikator,3)
            reading = round(float(readinglines)*multiplikator,nachkomma)
            wb_col = 8
            if save_flag == 1:
                worksheet.write('I1', 'Res. ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('I1', 'Res. ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')
            if check >= 1000:
                reading = 88888888
#                otto1 = "Eingang offen"
        elif "CONT" in einheit:
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_14.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(False)
            self.s_widget.setVisible(False)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "0")
#            leer = self.tcpip('SYSTem:BEEPer:STATe ON')
            self.dial.setRange(0,20)
            self.dial.setProperty("toolTip", "Ω Threshold")
            self.dial.setNotchesVisible(True)
            otto1 = "Continuity < "
#            if cold_boot == 1:
#                cold_boot = 0
#                self.dial.setValue(4)
            rad = self.dial.sliderPosition()
            otto1 = otto1 + str(rad) + "Ω"
            self.tcpip('CONT:THRE ' + str(rad))
            self.small_1.setProperty("text", str(rad) + " Ω")
            multiplikator = 1
            reading = round(float(readinglines)*multiplikator,2)
            if int(rad) >= reading:
                otto = "Ω ♪"
            elif int(rad) < reading:
                otto = "Ω"
            if reading == 1e09:
                reading = 88888888
            nachkomma = 6
        elif "DIOD" in einheit:
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_15.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(False)
            self.s_widget.setVisible(False)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "0")
            leer = self.tcpip('SYSTem:BEEPer:STATe ON')
            self.dial.setNotchesVisible(False)
            self.dial.setProperty("toolTip", "OFF")
            self.small_1.setProperty("text", "---")
            otto1 = "Diode"
            otto = "V ♪"
            multiplikator = 1
            reading = round(float(readinglines)*multiplikator,4)
            if reading == 1e09:
                reading = 88888888
            nachkomma = 4
        elif "CAP" in einheit:
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_9.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(False)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "0")
            plot_label = "Capacitor"
            self.dial.setRange(0,7)
            self.dial.setProperty("toolTip", "Range Auto " + str(CAP_409x))
            self.dial.setNotchesVisible(True)
            otto1 = "MANUAL"
            otto = self.tcpip('RANGE1?')
            if bereich != "1":
                if dmm4095 == True:
                    self.small_1.setProperty("text", CAP_409x[int(otto)-1])
                elif dmm4095 == False:
                    self.small_1.setProperty("text", CAP_409x[int(otto)-1])
                self.dial.setValue(int(otto))
            if bereich == "1":
                otto1 = "AUTO"
                self.small_1.setProperty("text", "AUTO")
                self.dial.setValue(0)
            if int(otto) == 1:
                otto1 = otto1 + " " + CAP_409x[int(otto)-1]
                otto = "nF"
                multiplikator = 1000000000
                nachkomma = 7
            elif int(otto) == 2:
                otto1 = otto1 + " " + CAP_409x[int(otto)-1]
                otto = "nF"
                multiplikator = 1000000000
                nachkomma = 6
            elif int(otto) == 3:
                otto1 = otto1 + " " + CAP_409x[int(otto)-1]
                otto = "nF"
                multiplikator = 1000000000
                nachkomma = 5
            elif int(otto) == 4:
                otto1 = otto1 + " " + CAP_409x[int(otto)-1]
                otto = "µF"
                multiplikator = 1000000
                nachkomma = 7
            elif int(otto) == 5:
                otto1 = otto1 + " " + CAP_409x[int(otto)-1]
                otto = "µF"
                multiplikator = 1000000
                nachkomma = 6
            elif int(otto) == 6:
                otto1 = otto1 + " " + CAP_409x[int(otto)-1]
                otto = "µF"
                multiplikator = 1000000
                nachkomma = 5
            elif int(otto) == 7:
                otto1 = otto1 + " " + CAP_409x[int(otto)-1]
                otto = "mF"
                multiplikator = 1000
                nachkomma = 7
            reading = round(float(readinglines)*multiplikator,nachkomma)
            wb_col = 9
            if save_flag == 1:
                worksheet.write('J1', 'Cap. ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('J1', 'Cap. ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')
            if reading >= 1000000000:
                reading = 88888888
        elif "FREQ" in einheit:
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_10.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(True)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "1")
            plot_label = "Frequency"
            self.dial.setRange(0,5)
            if dmm4095 == True:
                self.dial.setProperty("toolTip", "Range Auto " + str(VAC_4095))
            if dmm4095 == False:
                self.dial.setProperty("toolTip", "Range Auto " + str(VAC_4096))
            self.dial.setNotchesVisible(True)
            otto1 = "MANUAL"
            otto = self.tcpip('RANGE1?')
            if bereich != "1":
                if dmm4095 == True:
                    self.small_1.setProperty("text", VAC_4095[int(otto)-1])
                elif dmm4095 == False:
                    self.small_1.setProperty("text", VAC_4096[int(otto)-1])
                self.dial.setValue(int(otto))
            if bereich == "1":
                otto1 = "AUTO"
                self.small_1.setProperty("text", "AUTO")
                self.dial.setValue(0)
#            temptype = "Hz" + otto
            if int(otto) == 1:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "Hz"
                multiplikator = 1
                nachkomma = 8
            elif int(otto) == 2:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "Hz"
                multiplikator = 1
                nachkomma = 8
            elif int(otto) == 3:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "Hz"
                multiplikator = 1
                nachkomma = 8
            elif int(otto) == 4:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "Hz"
                multiplikator = 1
                nachkomma = 8
            elif int(otto) == 5:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "Hz"
                multiplikator = 1
                nachkomma = 8
            check = round(float(readinglines)*multiplikator,3)
            if check >= 1000000.0:
                multiplikator = 0.000001
                otto = "MHz"
                nachkomma = 8
            elif check >= 1000.0:
                multiplikator = 0.001
                otto = "kHz"
                nachkomma = 8
            reading = round(float(readinglines)*multiplikator,4)
            wb_col = 6
            if save_flag == 1:
                worksheet.write('G1', 'Freq. ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('G1', 'Freq. ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')
        elif "PER" in einheit:
            self.actiondB_dBm.setChecked(False)
            self.dbText.setVisible(False)
            self.pushButton_16.setStyleSheet("background-color:" + display_c_1 + ";")
            self.d_widget.setVisible(True)
            self.s_widget.setVisible(True)
            self.temp_widget.setVisible(False)
            self.comboBox.setProperty("enabled", "1")
            plot_label = "Cycle Duration"
            self.dial.setRange(0,5)
            if dmm4095 == True:
                self.dial.setProperty("toolTip", "Range Auto " + str(VAC_4095))
            if dmm4095 == False:
                self.dial.setProperty("toolTip", "Range Auto " + str(VAC_4096))
            self.dial.setNotchesVisible(True)
            otto1 = "MANUAL"
            otto = self.tcpip('RANGE1?')
            if bereich != "1":
                if dmm4095 == True:
                    self.small_1.setProperty("text", VAC_4095[int(otto)-1])
                elif dmm4095 == False:
                    self.small_1.setProperty("text", VAC_4096[int(otto)-1])
                self.dial.setValue(int(otto))
            if bereich == "1":
                otto1 = "AUTO"
                self.small_1.setProperty("text", "AUTO")
                self.dial.setValue(0)
#            temptype = "µs" + otto
            if int(otto) == 1:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "µs"
                multiplikator = 1000000
                nachkomma = 8
            elif int(otto) == 2:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "µs"
                multiplikator = 1000000
                nachkomma = 8
            elif int(otto) == 3:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "µs"
                multiplikator = 1000000
                nachkomma = 8
            elif int(otto) == 4:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "µs"
                multiplikator = 1000000
                nachkomma = 8
            elif int(otto) == 5:
                if dmm4095 == True:
                    otto1 = otto1 + " " + VAC_4095[int(otto)-1]
                elif dmm4095 == False:
                    otto1 = otto1 + " " + VAC_4096[int(otto)-1]
                otto = "µs"
                multiplikator = 1000000
                nachkomma = 6
            check = round(float(readinglines)*multiplikator,4)
            if check >= 1000.0:
                multiplikator = 1000
                otto = "ms"
            elif check >= 1000000.0:
                multiplikator = 1
                otto = "s"
            reading = round(float(readinglines)*multiplikator,nachkomma)
            wb_col = 7
            if save_flag == 1:
                worksheet.write('H1', 'Per. ('+ otto + ')')
                if round(float(offset)*multiplikator,6) != 0.000000:
                    worksheet.write('H1', 'Per. ( Rel. 0 Offset = ' + str(round(float(offset)*multiplikator,6)) + ' ' + otto + ' )')

        reading_bak = reading
        if int(ts) == 0:
          self.led_on(2)
          if "CONT" not in einheit and "DIOD" not in einheit:
            messungen = self.tcpip('CALCulate:AVERage:COUNt?')
            messungen = round(float(messungen),0)
            self.lineEdit_8.setProperty("text", "Samples: " + str(int(messungen)))
            if int(messungen) >= 2 and int(messungen_alt) != int(messungen):
                durch  = self.tcpip('CALCulate:AVERage:AVERage?')
                durch = round(float(durch)*multiplikator,4)
                self.lineEdit_5.setProperty("text", "Ø: "+ str(durch) + otto)
                maxi  = self.tcpip('CALCulate:AVERage:MAXimum?')
                maxi = round(float(maxi)*multiplikator,4)
                self.lineEdit_7.setProperty("text", "Max: " + str(maxi) + otto)
                mini  = self.tcpip('CALCulate:AVERage:MINimum?')
                mini = round(float(mini)*multiplikator,4)
                self.lineEdit_3.setProperty("text", "Min: " + str(mini) + otto)
                spann = round(float(maxi - mini),4)
                abw = array([[reading], [durch]])
                abweich = round(float(abw.std(ddof=1)),4)
                self.lineEdit_4.setProperty("text", "Spann: " + str(spann) + otto)
                self.lineEdit_6.setProperty("text", "Std Dev: " + str(abweich) + otto)
                if int(xy_counter) >= max_graph:
                    for i in range(len(x)-1):
                        x[i] = x[i+1]
                        y[i] = y[i+1]
                    self.graphWidget.clear()
                    y[i+1] = int(messungen)
                    x[i+1] = reading
                elif int(xy_counter) < max_graph:
                    xy_counter = xy_counter + 1
                    y.append(int(messungen))
                    x.append(reading)
                self.graphWidget.plot(y,x, pen=pen)

            if int(messungen) < int(messungen_alt):
                ts = 1
                self.stat_off()
                self.update()
            messungen_alt = messungen
            
        if reading == 888888888:
            self.lcdNumber.setFont(QFont('DejaVu Sans Mono', 32))
            self.lcdNumber.setProperty("text", " - RANGE ? - ")
            self.lcdNumber.setAlignment(QtCore.Qt.AlignCenter)
        elif reading == 88888888:
            self.lcdNumber.setFont(QFont('DejaVu Sans Mono', 32))
            self.lcdNumber.setProperty("text", "Input Open ?")
            self.lcdNumber.setAlignment(QtCore.Qt.AlignCenter)
        elif reading != 88888888:
            self.lcdNumber.setFont(QFont('DejaVu Sans Mono', 60))
            fo_string = nk[nachkomma-1].format(reading)
            self.lcdNumber.setProperty("text", fo_string)
            self.lcdNumber.setAlignment(QtCore.Qt.AlignRight)
            if limit_switch == 0:
                self.lineEdit_11.setProperty("placeholderText", str(round(reading+(reading/100*0.025),2)) + otto + '↑')
                self.lineEdit_12.setProperty("placeholderText", str(round(reading-(reading/100*0.025),2)) + otto + '↓')

            if round(float(offset)*multiplikator,6) != 0.000000:
                self.actionRel_Offset.setChecked(True)
                self.offsetText.setVisible(True)
                self.offsetText.setProperty("text", "Rel. 0 Offset = " + str(round(float(offset),6)) + " " + otto)
                self.lcdNumber.setFrame(True)
            elif round(float(offset)*multiplikator,6) == 0.0:
                self.actionRel_Offset.setChecked(False)
                self.offsetText.setVisible(False)
                self.lcdNumber.setFrame(False)

            if round(float(offset)*multiplikator,6) != 0.000000:
                self.actionRel_Offset.setChecked(True)
                self.offsetText.setVisible(True)
                self.offsetText.setProperty("text", "Rel. 0 Offset = " + str(round(float(offset),6)) + " " + otto)
                self.lcdNumber.setFrame(True)
            elif round(float(offset)*multiplikator,6) == 0.0:
                self.actionRel_Offset.setChecked(False)
                self.offsetText.setVisible(False)
                self.lcdNumber.setFrame(False)
            
        self.lineEdit.setProperty("text", otto)
        self.lineEdit_2.setProperty("text", otto1)

        offset_bak = round(float(offset),6)

        if limit_switch == 1:
            self.lineEdit_11.setStyleSheet("background-color: #a2a2a2; border: 1px solid black;")
            self.lineEdit_12.setStyleSheet("background-color: #a2a2a2; border: 1px solid black;")
            if reading > upper_val*multiplikator:
#                play_sound("sine", 880, 0.5, 200)  #  duration in milliseconds
                up_fail = up_fail + 1
                self.lineEdit_11.setStyleSheet("background-color: #ff0000; border: 1px solid black;")
            elif reading < lower_val*multiplikator:
#                play_sound("sine", 220, 0.5, 200)  #  duration in milliseconds
                low_fail = low_fail + 1
                self.lineEdit_12.setStyleSheet("background-color: #ff0000; border: 1px solid black;")
            self.lineEdit_11.setProperty("text", str(round(float(upper_val*multiplikator), 4)) + otto + '↑')
            self.lineEdit_14.setProperty("text", str(up_fail))
            self.lineEdit_12.setProperty("text", str(round(float(lower_val*multiplikator), 4)) + otto + '↓')
            self.lineEdit_15.setProperty("text", str(low_fail))

        einheit2 = self.tcpip('FUNC2?')
        if einheit2 not in ['NONe', 'VOLT', 'VOLT AC', 'CURR', 'CURR AC', 'FREQ', 'PER']:
            einheit2 = 'NONe'
        elif einheit2 == 'NONe':
            dual_flag = 0
            self.comboBox.setStyleSheet("background:rgb(127,127,127)")
            self.pushButton_5.setEnabled(False)
            self.lineEdit_9.setVisible(False)
            self.pushButton_5.setText("DUAL Display")
        index = dual_index.index(einheit2)
        if index != 0 and dual_index_alt == index:
            dual_flag = 1
            self.comboBox.setStyleSheet("background-color:" + display_c_2 + ";")
            self.label.setPixmap(self.led[2])
            if self.actionMeasure_STOP.isChecked() == False:
                self.led_on(0)
                self.dual_d()
                self.label.setPixmap(self.led[2])
                self.lineEdit_9.setVisible(True)
                self.comboBox.setProperty("currentIndex", index)
                self.pushButton_5.setText("<Dual toggle>")
                self.pushButton_5.setEnabled(True)
        elif index == 0:
            self.comboBox.setStyleSheet("background:rgb(127,127,127)")
            dual_flag = 0
#            self.timer_led.singleShot(30, self.led_off)
            self.pushButton_5.setEnabled(False)
            self.lineEdit_9.setVisible(False)
            self.comboBox.setProperty("currentIndex", index)
            self.pushButton_5.setText("DUAL Display")
        dual_index_alt = index

        if save_flag == 1:
            if save_timer >= save_start:
                save_timer = int(round(time.time()))
                save_intervall = int(self.comboBox_2.currentText())
                save_start = int(round(time.time())) + save_intervall
                self.label_2.setPixmap(self.led[2])
                now = datetime.now()
                worksheet.write(wb_row, 0, now, format_date)
                worksheet.write(wb_row, 1, now, format_time)
                worksheet.write(wb_row, wb_col, reading_bak)
                if dual_flag == 1:
                    worksheet.write(wb_row, wb_col2, reading2)
                wb_row = wb_row + 1
            elif save_start > save_timer:
                save_timer = int(round(time.time()))
                self.label_2.setPixmap(self.led[0])

        if self.actionMeasure_STOP.isChecked():
            if cold_boot == 0:
                leer = self.tcpip('SHOW OFF')
                self.actionMeasure_STOP.setChecked(True)
                self.timer_single.stop()
                self.buttons_off()
                self.label.setPixmap(self.led[2])
                self.comboBox.setStyleSheet("background:rgb(127,127,127)")
                self.small_1.setProperty("text", "STOP")
            elif cold_boot >= 1:
                self.timer_single.setInterval(poll_intervall)
                self.timer_single.start()
                cold_boot = cold_boot - 1
        if self.actionMeasure_STOP.isChecked() == False:
            self.timer_led.singleShot(30, self.led_off)
            self.timer_single.setInterval(poll_intervall)
            self.timer_single.start()
        
      except ValueError as msg:
        self.update()

    def dual_d(self):
      global ts, rad, einheit, durch, maxi, mini, reading, messungen, x, y, daten, line, p1, otto, otto1, einheit, pen, plot_label, HOST, PORT, s, dual_flag, dual, limit_switch, low_fail, up_fail, otto2, otto21, einheit2, reading2, dual_index, dual_index_alt, usb_switch, otto_alt, save_flag, save_timer, save_intervall, led_index, fileName, hugo, nk, wb_row, wb_col, wb_col2, format_date, format_time, workbook, worksheet, func_1, func_2, cold_boot, dmm4095, lower_val, upper_val, save_start
      try:
        multiplikator = 1
        nachkomma = 1
        readinglines2 = self.tcpip('MEAS2?')
        bereich2 = self.tcpip('AUTO2?')
        einheit2 = self.tcpip('FUNC2?')
        if "VOLT AC" in einheit2:
            self.pushButton_12.setStyleSheet("background-color:" + display_c_2 + ";")
            self.comboBox.setProperty("enabled", "1")
            plot_label2 = "Wechselspannung"
            otto21 = "MANUAL"
            otto2 = self.tcpip('RANGE2?')
            if bereich2 == "1":
                otto21 = "AUTO"
            if int(otto2) == 1:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "mV"
                multiplikator = 1000
                nachkomma = 8
            elif int(otto2) == 2:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "V"
                multiplikator = 1
                nachkomma = 8
            elif int(otto2) == 3:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "V"
                multiplikator = 1
                nachkomma = 7
            elif int(otto2) == 4:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "V"
                multiplikator = 1
                nachkomma = 6
            elif int(otto2) == 5:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "V"
                multiplikator = 1
                nachkomma = 6
            reading2 = round(float(readinglines2)*multiplikator,nachkomma)
            wb_col2 = 3
            if save_flag == 1:
                worksheet.write('D1', 'V AC ('+ otto2 + ')')
            if reading2 >= 1e09:
                reading2 = 888888888
        elif "VOLT" in einheit2:
            self.pushButton_6.setStyleSheet("background-color:" + display_c_2 + ";")
            self.comboBox.setProperty("enabled", "1")
            plot_label2 = "Gleichspannung"
            rad = 0
            otto21 = "MANUAL"
            otto2 = self.tcpip('RANGE2?')
            if bereich2 == "1":
                otto21 = "AUTO"
            if int(otto2) == 1:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VDC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VDC_4096[int(otto2)-1]
                otto2 = "mV"
                multiplikator = 1000
                nachkomma = 2
            elif int(otto2) == 2:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VDC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VDC_4096[int(otto2)-1]
                otto2 = "V"
                multiplikator = 1
                nachkomma = 4
            elif int(otto2) == 3:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VDC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VDC_4096[int(otto2)-1]
                otto2 = "V"
                multiplikator = 1
                nachkomma = 3
            elif int(otto2) == 4:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VDC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VDC_4096[int(otto2)-1]
                otto2 = "V"
                multiplikator = 1
                nachkomma = 2
            elif int(otto2) == 5:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VDC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VDC_4096[int(otto2)-1]
                otto2 = "V"
                multiplikator = 1
                nachkomma = 3
            reading2 = round(float(readinglines2)*multiplikator,nachkomma)
            wb_col2 = 2
            if save_flag == 1:
                worksheet.write('C1', 'V DC ('+ otto2 + ')')
        elif "CURR AC" in einheit2:
            self.pushButton_13.setStyleSheet("background-color:" + display_c_2 + ";")
            self.comboBox.setProperty("enabled", "1")
            plot_label2 = "Wechselstrom"
            otto21 = "MANUAL"
            otto2 = self.tcpip('RANGE2?')
            if bereich2 == "1":
                otto21 = "AUTO"
            if int(otto2) == 1:
                if dmm4095 == True:
                    otto21 = otto21 + " " + AAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + AAC_4096[int(otto2)-1]
                otto2 = "mA"
                multiplikator = 1000
                nachkomma = 7
            elif int(otto2) == 2:
                if dmm4095 == True:
                    otto21 = otto21 + " " + AAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + AAC_4096[int(otto2)-1]
                otto2 = "mA"
                multiplikator = 1000
                nachkomma = 6
            elif int(otto2) == 3:
                if dmm4095 == True:
                    otto21 = otto21 + " " + AAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + AAC_4096[int(otto2)-1]
                otto2 = "A"
                multiplikator = 1
                nachkomma = 8
            elif int(otto2) == 4:
                if dmm4095 == True:
                    otto21 = otto21 + " " + AAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + AAC_4096[int(otto2)-1]
                otto2 = "A"
                multiplikator = 1
                nachkomma = 7
            reading2 = round(float(readinglines2)*multiplikator,nachkomma)
            wb_col2 = 5
            if save_flag == 1:
                worksheet.write('F1', 'A AC ('+ otto2 + ')')
        elif "CURR" in einheit2:
            self.pushButton_7.setStyleSheet("background-color:" + display_c_2 + ";")
            self.comboBox.setProperty("enabled", "1")
            plot_label2 = "Gleichstrom"
            otto21 = "MANUAL"
            otto2 = self.tcpip('RANGE2?')
            if bereich2 == "1":
                otto21 = "AUTO"
            if int(otto2) == 1:
                if dmm4095 == True:
                    otto21 = otto21 + " " + ADC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + ADC_4096[int(otto2)-1]
                otto2 = "µA"
                multiplikator = 1000000
                nachkomma = 2
            elif int(otto2) == 2:
                if dmm4095 == True:
                    otto21 = otto21 + " " + ADC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + ADC_4096[int(otto2)-1]
                otto2 = "mA"
                multiplikator = 1000
                nachkomma = 4
            elif int(otto2) == 3:
                if dmm4095 == True:
                    otto21 = otto21 + " " + ADC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + ADC_4096[int(otto2)-1]
                otto2 = "mA"
                multiplikator = 1000
                nachkomma = 3
            elif int(otto2) == 4:
                if dmm4095 == True:
                    otto21 = otto21 + " " + ADC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + ADC_4096[int(otto2)-1]
                otto2 = "mA"
                multiplikator = 1000
                nachkomma = 2
            elif int(otto2) == 5:
                if dmm4095 == True:
                    otto21 = otto21 + " " + ADC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + ADC_4096[int(otto2)-1]
                otto2 = "A"
                multiplikator = 1
                nachkomma = 4
            elif int(otto2) == 6:
                if dmm4095 == True:
                    otto21 = otto21 + " " + ADC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + ADC_4096[int(otto2)-1]
                otto2 = "A"
                multiplikator = 1
                nachkomma = 3
            reading2 = round(float(readinglines2)*multiplikator,nachkomma)
            wb_col2 = 4
            if save_flag == 1:
                worksheet.write('E1', 'A DC ('+ otto2 + ')')
        elif "FREQ" in einheit2:
            self.pushButton_10.setStyleSheet("background-color:" + display_c_2 + ";")
            self.comboBox.setProperty("enabled", "1")
            plot_label2 = "Frequenz"
            otto21 = "MANUAL"
            otto2 = self.tcpip('RANGE2?')
            if bereich2 == "1":
                otto21 = "AUTO"
#            temptype = "Hz" + otto2
            if int(otto2) == 1:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "Hz"
                multiplikator = 1
                nachkomma = 8
            elif int(otto2) == 2:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "Hz"
                multiplikator = 1
                nachkomma = 8
            elif int(otto2) == 3:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "Hz"
                multiplikator = 1
                nachkomma = 8
            elif int(otto2) == 4:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "Hz"
                multiplikator = 1
                nachkomma = 8
            elif int(otto2) == 5:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "Hz"
                multiplikator = 1
                nachkomma = 8
            check = round(float(readinglines2)*multiplikator,4)
            if check >= 1000000.0:
                multiplikator = 0.000001
                otto2 = "MHz"
            elif check >= 1000.0:
                multiplikator = 0.001
                otto2 = "kHz"
            reading2 = round(float(readinglines2)*multiplikator,nachkomma)
            wb_col2 = 6
            if save_flag == 1:
                worksheet.write('G1', 'Freq. ('+ otto2 + ')')
            
        elif "PER" in einheit2:
            self.pushButton_16.setStyleSheet("background-color:" + display_c_2 + ";")
            self.comboBox.setProperty("enabled", "1")
            plot_label2 = "Periodendauer"
            otto21 = "MANUAL"
            otto2 = self.tcpip('RANGE2?')
            if bereich2 == "1":
                otto21 = "AUTO"
#            temptype = "µs" + otto2
            if int(otto2) == 1:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "µs"
                multiplikator = 1000000
                nachkomma = 8
            elif int(otto2) == 2:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "µs"
                multiplikator = 1000000
                nachkomma = 8
            elif int(otto2) == 3:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "µs"
                multiplikator = 1000000
                nachkomma = 8
            elif int(otto2) == 4:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "µs"
                multiplikator = 1000000
                nachkomma = 8
            elif int(otto2) == 5:
                if dmm4095 == True:
                    otto21 = otto21 + " " + VAC_4095[int(otto2)-1]
                elif dmm4095 == False:
                    otto21 = otto21 + " " + VAC_4096[int(otto2)-1]
                otto2 = "µs"
                multiplikator = 1000000
                nachkomma = 6
            check = round(float(readinglines2)*multiplikator,4)
            if check >= 1000.0:
                multiplikator = 1000
                otto2 = "ms"
            elif check >= 1000000.0:
                multiplikator = 1
                otto2 = "s"
            reading2 = round(float(readinglines2)*multiplikator,4)
            wb_col2 = 7
            if save_flag == 1:
                worksheet.write('H1', 'Per. ('+ otto2 + ')')

        if reading2 == 888888888:
            self.lineEdit_9.setStyleSheet("background-color:" + display_c_2 + ";")
            self.lineEdit_9.setProperty("text", " - RANGE ? - " + otto2 + " " + otto21)
        elif reading2 != 888888888:
            fo_string = nk[nachkomma-1].format(reading2)
            self.lineEdit_9.setStyleSheet("background-color:" + display_c_2 + ";")
            self.lineEdit_9.setProperty("text", fo_string + " " + otto2 + " - " + otto21)

        self.timer_led.singleShot(30, self.led_off)
#        self.label.setPixmap(self.led[0])
      except ValueError as msg:
        self.update()

EXIT_CODE_REBOOT = -11231351

def main():
    exitCode = 0
    while True:
        try: app = QtWidgets.QApplication(sys.argv)
        except RuntimeError: app = QApplication.instance()
        app.aboutToQuit.connect(ende)
        window = Ui()
        exitCode = app.exec_()
        if exitCode != EXIT_CODE_REBOOT: break
    return exitCode

main()
