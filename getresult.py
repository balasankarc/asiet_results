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

import argparse
import json
import statistics
import sys
import textwrap
from argparse import RawTextHelpFormatter as rt

import pdftableextract as pdf
import requests
from lxml import etree
from pyPdf import PdfFileReader
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer


def getexamlist(url):
    '''This method prints a list of exams whose results are available'''
    try:
        page = requests.get(url)
        pagecontenthtml = page.text
        tree = etree.HTML(pagecontenthtml)
        code = tree.xpath('//option/@value')
        text = tree.xpath('//option')
        print "|--------|" + "-----------" * 6 + "----|"
        print "|  Code  |\t\t\t\t   Exam\t\t\t\t\t|"
        print "|--------|" + "-----------" * 6 + "----|"
        for i in range(1, len(code)):
            examname = [x.ljust(60 - len(x))
                        for x in textwrap.wrap(text[i].text, 60)]
            print "|  ", code[i], "\t |\t", "\t|\n|    \t |\t".join(examname),
            print "\t\t|"
            print "|--------|" + "-----------" * 6 + "----|"
    except:
            print "There are some issues with the connectivity.",
            print "May be due to heavy load. Please try again later"
            sys.exit(0)


def download(url, examcode, start, end):
    '''Using the specified url this function downloads the results of register
    numbers from 'start' to 'end'.'''
    try:
        for i in range(start, end + 1):
            print "Roll Number #", i
            payload = dict(exam=examcode, prn=i, Submit2='Submit')
            r = requests.post(url, payload)
            if r.status_code == 200:
                with open('result' + str(i) + '.pdf', 'wb') as resultfile:
                    for chunk in r.iter_content():
                        resultfile.write(chunk)
    except:
        print "There are some issues with the connectivity.",
        print "May be due to heavy load. Please try again later"
        sys.exit(0)


def process(start, end):
    '''This method processes the specified results and populate necessary data
    structures.'''
    global result, passcount, failcount, absentcount, numberofstudents
    global college, branch, exam
    numberofstudents = end - start + 1
    for count in range(start, end + 1):
        try:
            print "Roll Number #", count
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
                    exam = i[0][exampos:][11:].strip().title()
                    register = i[0][registerpos:exampos][13:].strip()
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
        except:
            f.close()
            print "Invalid result file for Roll Number #", count
            numberofstudents -= 1
            continue
    jsonout = json.dumps(result)
    outfile = open('output.json', 'w')
    outfile.write(jsonout)
    outfile.close()


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
        for branch in result[college]:
            Story.append(Paragraph(college, styles["Center1"]))
            Story.append(Spacer(1, 0.25 * inch))
            Story.append(Paragraph(exam, styles["Center2"]))
            Story.append(Spacer(1, 12))
            numberofstudents = len(result[college][branch].itervalues().next())
            Story.append(Paragraph(branch, styles["Center2"]))
            Story.append(Spacer(1, 0.25 * inch))
            Story.append(Paragraph("Total Number of Students : %d" %
                         numberofstudents, styles["Normal2"]))
            Story.append(Spacer(1, 12))
            for subject in result[college][branch]:
                marklist = [int(result[college][branch][subject][x][0])
                            for x in result[college][branch][subject]]
                average = statistics.mean(marklist)
                stdev = statistics.pstdev(marklist)
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
        1. Get list and codes of exams\n\
        2. Download Results \n\
        3. Generate Summary")
    parser.add_argument(
        "--start", help="Starting register number", default=-1)
    parser.add_argument(
        "--end", help="Ending register number", default=-1)
    parser.add_argument("--exam", help="Exam code", default=-1)
    args = parser.parse_args()

    url = 'http://projects.mgu.ac.in/bTech/btechresult/index.php?module=public'
    url = url + '&attrib=result&page=result'
    start = int(args.start)
    end = int(args.end)
    exam = int(args.exam)
    operation = int(args.operation)
    if operation == 1:
        getexamlist(url)
    elif operation == 2:
        if start == -1:
            print "Starting register number missing. Use --start option"
            sys.exit(0)
        elif end == -1:
            print "Ending register number missing . Use --end option"
            sys.exit(0)
        elif exam == -1:
            print "Exam code missing. Use operation #1 to get list of exams.",
            print "Use --exam option to specify code."
            sys.exit(0)
        print "###############################################"
        print "Downloading Results"
        print "###############################################"
        download(url, exam, start, end)
    elif operation == 3:
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
