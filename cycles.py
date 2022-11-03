# -*- coding: utf-8 -*-
"""
Created on Thu Aug 22 14:25:59 2019

@author: wiktor
"""

import sys
from openpyxl import *
from datetime import datetime

#read file name from command line argument and load
print('Loading components file...')
try:
    file = sys.argv[1]
except IndexError:
    print("Usage: python3 cycles.py report_components.xlsx report_ac.xlsx")
    quit()
wb_components = load_workbook(filename = file, data_only=True)
ws_comp = wb_components.active
print(str(ws_comp.max_row - 1) + ' controls loaded.')

print('Loading A/C flights file...')
try:
    file = sys.argv[2]
except IndexError:
    print("Usage: python3 cycles.py report_components.xlsx report_ac.xlsx")
    quit()
wb_ac = load_workbook(filename = file, data_only=True)
ws_ac = wb_ac.active 
print(str(ws_ac.max_row - 1) + ' flight logs loaded.')

#read hours anc cycles of an ac from the report_ac file
ac_hours = ws_ac['H' + str(ws_ac.max_row)].value
ac_cycles = ws_ac['J' + str(ws_ac.max_row)].value

#formatting of the output file
ws_comp['L1'] = "TOTAL_HRS"
ws_comp['M1'] = "TOTAL_CYC"
ws_comp['N1'] = "TOTAL_DAYS"

#main loop
for row in range(2, ws_comp.max_row + 1):
    date_installed = ws_comp['K' + str(row)].value
    #if part has no installation date, take z date
    if date_installed == None:
        date_installed = ws_ac['B2'].value
    #if part installation date is smaller than z date, take z date
    elif date_installed < ws_ac['B2'].value:
        date_installed = ws_ac['B2'].value
        
    #find the first date larger than itself
    for line in range(2, ws_ac.max_row + 1):
        if ws_ac['B' + str(line)].value >= date_installed:
            hrs_before = ws_ac['H' + str(line - 1)].value
            cyc_before = ws_ac['J' + str(line - 1)].value
            break
    
    #if the part has been installed on z date, it has 0 hrs_before installation 
    if date_installed == ws_ac['B2'].value:
        hrs_before = 0
        cyc_before = 0
        
    #calculate difference between ac_hours, ac_cycles and before
    hrs_now = ws_comp['G' + str(row)].value + (ac_hours - int(hrs_before))
    cyc_now = ws_comp['H' + str(row)].value + (ac_cycles - int(cyc_before))
    days_delta = datetime.now() - date_installed
    days_now = ws_comp['I' + str(row)].value + days_delta.days
    
    #write in new data
    ws_comp['L' + str(row)] = hrs_now
    ws_comp['M' + str(row)] = cyc_now
    ws_comp['N' + str(row)] = days_now
    
    #give feedback
    print('PN: ' + str(ws_comp['A' + str(row)].value) + ' SN: ' + str(ws_comp['B' + str(row)].value) + ' Control: ' + str(ws_comp['C' + str(row)].value) + ' has been computed.')
    
wb_components.save(filename='output_' + str(ws_ac['A2'].value) + '.xlsx')
print('Done.')