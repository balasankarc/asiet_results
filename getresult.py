#! /usr/bin/python
# -*- coding: UTF-8 -*-
#
# Copyright 2015 Balasankar C <balasankarc@autistici.org>
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# .
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
# .
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.

import requests
import sys
import pdftableextract as pdf
from pyPdf import PdfFileReader
import argparse
from argparse import RawTextHelpFormatter as rt
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import json
import statistics


def download(url, start, end):
    '''Using the specified url this function downloads the results of register
    numbers from 'start' to 'end'.'''
    for i in range(start, end + 1):
        print "Roll Number #", i
        payload = dict(exam=28, prn=i, Submit2='Submit')
        r = requests.post(url, payload)
        if r.status_code == 200:
            with open('result' + str(i) + '.pdf', 'wb') as resultfile:
                for chunk in r.iter_content():
                    resultfile.write(chunk)


def process(start, end):
    '''This method processes the specified results and populate necessary data
    structures.'''
    global result, passcount, failcount, absentcount, numberofstudents
    global college, branch, exam
    outputfile = open('marklist.csv', 'w')
    numberofstudents = end - start + 1
    for count in range(start, end + 1):
        try:
            print "Roll Number #", count
            string = ""
            pages = ["1"]
            f = open("result" + str(count) + ".pdf", "rb")
            PdfFileReader(f)          # Checking if valid pdf file
            f.close()
            cells = [pdf.process_page("result" + str(count) + ".pdf", p)
                     for p in pages]
            cells = [item for sublist in cells for item in sublist]
            li = pdf.table_to_list(cells, pages)[1]
            for i in li:
                if 'Branch' in i[0]:
                    collegepos = i[0].index('College : ')
                    branchpos = i[0].index('Branch : ')
                    namepos = i[0].index('Name : ')
                    registerpos = i[0].index('Register No : ')
                    exampos = i[0].index('Exam Name : ')
                    college = i[0][collegepos:branchpos][9:].strip().title()
                    branch = i[0][branchpos:namepos][9:].strip().title()
                    name = i[0][namepos:registerpos][6:].strip().title()
                    exam = i[0][exampos:][11:].strip().title()
                    register = i[0][registerpos:exampos][13:].strip()
                    string = branch + "," + name + "," + register
                    if college not in result:
                        result[college] = {}
                    if branch not in result[college]:
                        result[college][branch] = {}
                elif 'Mahatma' in i[0]:
                    pass
                elif 'Sl. No' in i[0]:
                    pass
                elif 'Semester Result' in i[1]:
                    pass
                else:
                    subject = [i][0][1]
                    internal = i[2]
                    external = i[3]
                    if internal == '-':
                        internal = 0
                    else:
                        internal = int(internal)
                    if external == '-':
                        external = 0
                    else:
                        external = int(external)
                    total = internal + external
                    res = i[5]
                    if subject not in result[college][branch]:
                        result[college][branch][subject] = {}
                    result[college][branch][subject][register] = \
                        [total, res]
            string = string + "\n"
            outputfile.write(string)
        except:
            f.close()
            print "Invalid result file for Roll Number #", count
            numberofstudents -= 1
            continue
    jsonout = json.dumps(result)
    outfile = open('output.json', 'w')
    outfile.write(jsonout)
    outfile.close()
    outputfile.close()


def getsummary():
    '''This method generates summary pdf from the results of result processor.
    '''
    global result, passcount, failcount, absentcount, numberofstudents
    global college, branch, exam, totalsum

    doc = SimpleDocTemplate("report.pdf", pagesize=A4,
                            rightMargin=72, leftMargin=72,
                            topMargin=50, bottomMargin=30)
    Story = []
    doc.title = "Exam Result Summary"
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='Center1', alignment=1, fontSize=18))
    styles.add(ParagraphStyle(name='Center2', alignment=1, fontSize=13))
    styles.add(ParagraphStyle(name='Normal2', bulletIndent=20))
    styles.add(ParagraphStyle(name='Normal3', fontSize=12))
    for college in result:
        Story.append(Paragraph(college, styles["Center1"]))
        Story.append(Spacer(1, 0.25 * inch))
        Story.append(Paragraph(exam, styles["Center2"]))
        Story.append(Spacer(1, 12))
        for branch in result[college]:
            numberofstudents = len(
                result[college].itervalues().next().itervalues().next())
            Story.append(Paragraph(branch, styles["Center2"]))
            Story.append(Spacer(1, 0.25 * inch))
            Story.append(Paragraph("Total Number of Students : %d" %
                         numberofstudents, styles["Normal2"]))
            Story.append(Spacer(1, 12))
            for subject in result[college][branch]:
                marklist = [int(result[college][branch][subject][x][0])
                            for x in result[college][branch][subject]]
                print marklist
                average = statistics.mean(marklist)
                stdev = statistics.pstdev(marklist)
                print "Subject : ", stdev
                # numberofstudents = len(result[college][branch][subject])
                passlist = {x for x in result[college][branch][
                    subject] if 'P' in result[college][branch][subject][x]}
                faillist = {x for x in result[college][branch][
                    subject] if 'F' in result[college][branch][subject][x]}
                absentlist = {x for x in result[college][branch][
                    subject] if 'AB' in result[college][branch][subject][x]}
                passcount = len(passlist)
                failcount = len(faillist)
                absentcount = len(absentlist)
                percentage = float(passcount) / numberofstudents
                # total = 0
                # for regno in result[college][branch][subject]:
                #    total += result[college][branch][subject][regno][0]
                # average = total / numberofstudents
                subjectname = "<b>%s</b>" % subject
                passed = "<bullet>&bull;</bullet>Students Passed : %d" \
                    % passcount
                failed = " <bullet>&bull;</bullet>Students Failed : %d" \
                    % failcount
                absent = " <bullet>&bull;</bullet>Students Absent : %d" \
                    % absentcount
                percentage = " <bullet>&bull;</bullet>Pass Percentage : %.2f"\
                    % percentage
                average = " <bullet>&bull;</bullet>Average Marks : %.2f" \
                    % average
                stdev = "<bullet>&bull;</bullet>Standard Deviation : %.2f" \
                    % stdev
                Story.append(Paragraph(subjectname, styles["Normal"]))
                Story.append(Spacer(1, 12))
                Story.append(Paragraph(passed, styles["Normal2"]))
                Story.append(Spacer(1, 12))
                Story.append(Paragraph(failed, styles["Normal2"]))
                Story.append(Spacer(1, 12))
                Story.append(Paragraph(absent, styles["Normal2"]))
                Story.append(Spacer(1, 12))
                Story.append(Paragraph(percentage, styles["Normal2"]))
                Story.append(Spacer(1, 12))
                Story.append(Paragraph(average, styles["Normal2"]))
                Story.append(Spacer(1, 12))
                Story.append(Paragraph(stdev, styles["Normal2"]))
                Story.append(Spacer(1, 12))
        Story.append(PageBreak())
    doc.build(Story)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Download and Generate Result Summaries',
        formatter_class=rt)
    parser.add_argument(
        "operation", help="Specify the operation\n\
        1. Download Results \n\
        2. Generate Summary")
    parser.add_argument(
        "start_number", help="Specify the starting register number")
    parser.add_argument(
        "end_number", help="Specify the ending register number")
    args = parser.parse_args()

    url = 'http://projects.mgu.ac.in/bTech/btechresult/index.php?module=public'
    url = url + '&attrib=result&page=result'
    start = int(args.start_number)
    end = int(args.end_number)
    operation = int(args.operation)
    if operation == 1:
        print "###############################################"
        print "Downloading Results"
        print "###############################################"
        download(url, start, end)
    elif operation == 2:
        result = {}
        totalsum = {}
        passcount = {}
        failcount = {}
        absentcount = {}
        numberofstudents = 0
        college = ''
        branch = ''
        exam = ''
        print "###############################################"
        print "Processing Results"
        print "###############################################"
        process(start, end)
        getsummary()
    else:
        print "Wrong option"
        sys.exit(0)
