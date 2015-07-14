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


def download(url, start, end):
    for i in range(start, end + 1):
        print "Roll Number #", i
        payload = dict(exam=28, prn=i, Submit2='Submit')
        r = requests.post(url, payload)
        if r.status_code == 200:
            with open('result' + str(i) + '.pdf', 'wb') as resultfile:
                for chunk in r.iter_content():
                    resultfile.write(chunk)


def process(start, end):
    global result, passcount, failcount, absentcount, numberofstudents
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
                if 'ADI SHANKARA' in i[0]:
                    branchpos = i[0].index('Branch : ')
                    namepos = i[0].index('Name : ')
                    registerpos = i[0].index('Register No : ')
                    exampos = i[0].index('Exam Name : ')
                    branch = i[0][branchpos:namepos][9:].strip()
                    name = i[0][namepos:registerpos][6:].strip()
                    register = i[0][registerpos:exampos][13:].strip()
                    string = branch + "," + name + "," + register
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
                    if subject not in passcount:
                        passcount[subject] = 0
                    if subject not in failcount:
                        failcount[subject] = 0
                    if res == 'P':
                        passcount[subject] += 1
                    if res == 'F':
                        failcount[subject] += 1
                    string += "," + subject + "," + str(internal) + \
                        "," + str(external)
                    if subject in result:
                        result[subject] += total
                    else:
                        result[subject] = total
            string = string + "\n"
            outputfile.write(string)
        except:
            f.close()
            print "Invalid result file for Roll Number #", count
            numberofstudents -= 1
            continue

    outputfile.close()


def getsummary():
    global result, passcount, failcount, absentcount, numberofstudents
    for key in result:
        print key
        print "\t Total Students :", numberofstudents
        print "\t Students Passed :", passcount[key]
        print "\t Students Failed :", failcount[key]
        print "\t Pass Percentage :",
        print float(passcount[key] * 100) / numberofstudents
        avg = float(result[key]) / numberofstudents
        print "\t Average Marks :", avg

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Download and Generate Result Summaries')
    parser.add_argument(
        "operation", help="Specify the operation\n\
        1. Download Results \n\
        2. Generate Summary")
    parser.add_argument(
        "start_number", help="Specify the starting register number")
    parser.add_argument(
        "end_number", help="Specify the starting register number")
    args = parser.parse_args()

    if len(sys.argv) < 4:
        print "Usage : python result.py [option] <start reg no> <end reg no>"
        sys.exit(0)

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
        passcount = {}
        failcount = {}
        absentcount = {}
        numberofstudents = 0
        print "###############################################"
        print "Processing Results"
        print "###############################################"
        process(start, end)
        print "###############################################"
        print "Results Summary"
        print "###############################################"
        getsummary()
    else:
        print "Wrong option"
        sys.exit(0)
