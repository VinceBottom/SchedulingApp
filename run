# -*- coding: utf-8 -*-
##
##
##
##import modules statements- requires python-docx, datetime, timedelta, pyinstaller (only if exe desired)
#import docx2txt
import docx
#import sys
from docx import document
from docx.enum.text import WD_COLOR_INDEX
import docx.enum.text
##
#import calendar
#import string
import datetime
from datetime import timedelta
#import time
##while loop for answer verification will repeat unless user confirms with 'y'
x = False
while x == False:
##input date, hours, 
    date1 = input("Enter date of daytime test session (YY/MM/DD): ")
    y = False
    while y == False:
        hrsleep = input("Enter 7, 8.5, or 10 (for hours of sleep): ")
        if hrsleep == "7" or hrsleep == "8.5" or hrsleep == "10":
            y = True
        else:
            print("Please enter a 7, 8.5, or 10.")
    rise1 = input("Enter rise time (xx:xx): ")
    name1 = input ("Enter subject name: ")
    ##convert to datetime format
    date2 = datetime.datetime.strptime(date1, '%y/%m/%d').date().strftime('%A')
    date1 = datetime.datetime.strptime(date1, '%y/%m/%d').date()
    ##eliminate user entered AM by mistake just in case
    rise1 = rise1.strip("AaMmPp")
    rise1 = rise1
    ##verify response sequence
    print("")##aesthetics only
    print("Is this correct?")
    print("Subject name: " + name1)
    print("Date of daytime test: " + str(date1) + "  " + str(date2) + " session")
    print("Sleep time: " + hrsleep + "h")
    print("Rise time: " + rise1 +"AM")
    verify = input("Please enter 'Y' or 'N': ")
    verify = verify.lower()##changes to all lowercase for ease of checking
    if verify.find('y')!=-1:
        x = True
    elif verify.find('y')==-1: 
        print("")
        print("Ok, please reenter values.")
print("Ok, document is saved to K drive.")
##.weekday() will return a datetime date object as an integer representing the number of the day of the week
##
##make datetime object starting from daytime session
##time handling sequence using timedelta to subtract
##convert to string to make full datetime object of datetime object, to subtract days from there
fulldt = str(date1) + " " + rise1
fulldt = datetime.datetime.strptime(fulldt, '%Y-%m-%d %H:%M')
##calculate bedtime
##altered sleep variables first (use time and date)
hrsleep = float(hrsleep)
bedtime_dtobj = fulldt - timedelta(hours=hrsleep)
firstseq_bedtime = fulldt - timedelta(hours=8.5)
first_bdt = firstseq_bedtime.strftime('%I:%M%p')
bedtime = bedtime_dtobj.strftime('%I:%M%p')
bdt0 = fulldt - timedelta(days=1)
bdt1 = bdt0 - timedelta(days=1)
bdt2 = bdt0 - timedelta(days=2)
bdt3 = bdt0 - timedelta(days=3)
bdt4 = bdt0 - timedelta(days=4)
bdt5 = bdt0 - timedelta(days=5)
bdt6 = bdt0 - timedelta(days=6)
##normal sleep schedule variables (use date only)- use fulldt in case bedtime is am vs pm
nbdt7 = bdt0 - timedelta(days=7)
nbdt8 = bdt0 - timedelta(days=8)
nbdt9 = bdt0 - timedelta(days=9)
nbdt10 = bdt0 - timedelta(days=10)
nbdt11 = bdt0 - timedelta(days=11)
nbdt12 = bdt0 - timedelta(days=12)
nbdt13 = bdt0 - timedelta(days=13)
nbdt14 = bdt0 - timedelta(days=14)
#nbdt15 = bdt0 - timedelta(days=15)
ra_arrivaltime = bedtime_dtobj - timedelta(hours=2)##sets research assistant arrival time to two hours prior to bedtime
ra_arrivaltime = ra_arrivaltime.strftime('%I:%M%p')##strips date so arrival time can be applied to two different nights
sbj_arrivaltime = fulldt + timedelta(hours=1.5)
sbj_arrivaltime = sbj_arrivaltime.strftime('%I:%M%p')
wake_time = fulldt.strftime("%I:%M%p")
##trim 0's from front of sbj arrival time and rise time for aesthetic purposes, if they are present
if sbj_arrivaltime[0] == '0':
    sbj_arrivaltime = sbj_arrivaltime[1:]
if wake_time[0] == '0':
    wake_time = wake_time[1:]
###############################
def day(datetime_object):##shortcut to day of the week
    return datetime_object.strftime('%A')
##document writing
if hrsleep == "7" or hrsleep == "10":
    hrsleep = int(hrsleep)
doc = docx.Document()
doc_para = doc.add_paragraph('')
doc_para.add_run("No naps during the day").italic=True
doc_para = doc.add_paragraph('Name:  ' + name1 + "\t" + "\t")
doc_para.add_run(str(hrsleep)).bold=True 
doc_para.add_run('  Hours in bed')
doc_para = doc.add_paragraph()
#paragraph_format = doc_para.paragraph_format
doc_para.add_run('Date:').underline=True
##sequence per night- starting at 14
doc_para = doc.add_paragraph(nbdt14.strftime('%a') + "\t" + nbdt14.strftime('%d') + '-' + nbdt14.strftime('%b') + "\t" + "\t" + "\t" + ' Put on watch.')
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t" + "Go to bed at your typical " + nbdt14.strftime('%A') + " bed time.")
##13
doc_para = doc.add_paragraph(nbdt13.strftime('%a') + "\t" + nbdt13.strftime('%d') + '-' + nbdt13.strftime('%b') + "\t" + "\t" + "\t" + ' Wake up at your typical ' + nbdt13.strftime('%A') + " wake time.")
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t" + "Go to bed at your typical " + nbdt13.strftime('%A') + " bed time.")
##12
doc_para = doc.add_paragraph(nbdt12.strftime('%a') + "\t" + nbdt12.strftime('%d') + '-' + nbdt12.strftime('%b') + "\t" + "\t" + "\t" + ' Wake up at your typical ' + nbdt12.strftime('%A') + " wake time.")
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t" + "Go to bed at your typical " + nbdt12.strftime('%A') + " bed time.")
##11
doc_para = doc.add_paragraph(nbdt11.strftime('%a') + "\t" + nbdt11.strftime('%d') + '-' + nbdt11.strftime('%b') + "\t" + "\t" + "\t" + ' Wake up at your typical ' + nbdt11.strftime('%A') + " wake time.")
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t" + "Go to bed at your typical " + nbdt11.strftime('%A') + " bed time.")
##10
doc_para = doc.add_paragraph(nbdt10.strftime('%a') + "\t" + nbdt10.strftime('%d') + '-' + nbdt10.strftime('%b') + "\t" + "\t" + "\t" + ' Wake up at your typical ' + nbdt10.strftime('%A') + " wake time.")
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t" + "Go to bed at your typical " + nbdt10.strftime('%A') + " bed time.")
##9
doc_para = doc.add_paragraph(nbdt9.strftime('%a') + "\t" + nbdt9.strftime('%d') + '-' + nbdt9.strftime('%b') + "\t" + "\t" + "\t" + ' Wake up at your typical ' + nbdt9.strftime('%A') + " wake time.")
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t" + "Go to bed at your typical " + nbdt9.strftime('%A') + " bed time.")
##8
doc_para = doc.add_paragraph(nbdt8.strftime('%a') + "\t" + nbdt8.strftime('%d') + '-' + nbdt8.strftime('%b') + "\t" + "\t" + "\t" + ' Wake up at your typical ' + nbdt8.strftime('%A') + " wake time.")
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t" + "Go to bed at your typical " + nbdt8.strftime('%A') + " bed time.")
##7
doc_para = doc.add_paragraph(nbdt7.strftime('%a') + "\t" + nbdt7.strftime('%d') + '-' + nbdt7.strftime('%b') + "\t" + "\t" + "\t" + ' Wake up at your typical ' + nbdt7.strftime('%A') + " wake time.")
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t" + "Go to bed at your typical " + nbdt7.strftime('%A') + " bed time.")
##6
doc_para = doc.add_paragraph(bdt6.strftime('%a') + "\t" + bdt6.strftime('%d') + '-' + bdt6.strftime('%b') + "\t" + "\t" + "\t" + ' Wake up at your typical ' + bdt6.strftime('%A') + " wake time.")
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t")
font = doc_para.add_run(first_bdt).font
font.highlight_color = WD_COLOR_INDEX.GRAY_25
font.highlight_color
doc_para.add_run("  Go to bed")
##5
doc_para = doc.add_paragraph(bdt5.strftime('%a') + "\t" + bdt5.strftime('%d') + '-' + bdt5.strftime('%b') + "\t" + "\t" + "\t" + wake_time + '  Wake up')
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t")
font = doc_para.add_run(first_bdt).font
font.highlight_color = WD_COLOR_INDEX.GRAY_25
font.highlight_color
doc_para.add_run("  Go to bed")
##4
doc_para = doc.add_paragraph(bdt4.strftime('%a') + "\t" + bdt4.strftime('%d') + '-' + bdt4.strftime('%b') + "\t" + "\t" + "\t" + wake_time + '  Wake up')
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t")
font = doc_para.add_run(first_bdt).font
font.highlight_color = WD_COLOR_INDEX.GRAY_25
font.highlight_color
doc_para.add_run("  Go to bed")
doc_para = doc.add_paragraph()
##3
doc_para = doc.add_paragraph(bdt3.strftime('%a') + "\t" + bdt3.strftime('%d') + '-' + bdt3.strftime('%b') + "\t" + "\t" + "\t" + wake_time + '  Wake up')
doc_para.add_run("  ***RESTRICTION BEGIN***").bold=True
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t")
font = doc_para.add_run(bedtime).font
font.bold=True
font.highlight_color = WD_COLOR_INDEX.GRAY_25
font.highlight_color
doc_para.add_run("  Go to bed")
#table = doc.add_table(rows=1, cols=1)
#cell = table.cell(0, 0)
#cell.text = (bedtime)
#table.style = 'Table Grid'
##2
doc_para = doc.add_paragraph(bdt2.strftime('%a') + "\t" + bdt2.strftime('%d') + '-' + bdt2.strftime('%b') + "\t" + "\t" + "\t" + wake_time + '  Wake up')
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t")
doc_para.add_run (ra_arrivaltime + "  Arrival time for electrode night").bold=True
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t")
font = doc_para.add_run(bedtime).font
font.bold=True
font.highlight_color = WD_COLOR_INDEX.GRAY_25
font.highlight_color
doc_para.add_run("  Go to bed")
##1
doc_para = doc.add_paragraph(bdt1.strftime('%a') + "\t" + bdt1.strftime('%d') + '-' + bdt1.strftime('%b') + "\t" + "\t" + "\t" + wake_time + '  Wake up')
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t")
font = doc_para.add_run(bedtime).font
font.bold=True
font.highlight_color = WD_COLOR_INDEX.GRAY_25
font.highlight_color
doc_para.add_run("  Go to bed")
##0
doc_para = doc.add_paragraph(bdt0.strftime('%a') + "\t" + bdt0.strftime('%d') + '-' + bdt0.strftime('%b') + "\t" + "\t" + "\t" + wake_time + '  Wake up')
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t")
doc_para.add_run (ra_arrivaltime + "  Arrival time for electrode night").bold=True
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t")
font = doc_para.add_run(bedtime).font
font.bold=True
font.highlight_color = WD_COLOR_INDEX.GRAY_25
font.highlight_color
doc_para.add_run("  Go to bed")
#doc_para = doc.add_paragraph(longspace + shortspace + "(Arrival time for electrode night) @ " + ra_arrivaltime)
##fulldt
doc_para = doc.add_paragraph(fulldt.strftime('%a') + "\t" + fulldt.strftime('%d') + '-' + fulldt.strftime('%b') + "\t" + "\t" + "\t" + wake_time + '  Wake up')
doc_para = doc.add_paragraph("\t" + "\t" + "\t" + "\t" + "8:30AM  Arrive at the lab with electrodes still on head")
##newdocname = 'K:\ ' + name1 + str(date1) + '.docx'
newdocname = name1 + str(date1) + '.docx'
doc.save(newdocname)
#my_text = docx2txt.process(newdocname)
#print(my_text)
##
##transfer to exe sequence
##pyinstaller yourprogram.py

