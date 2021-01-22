"""
Copyright 2020, Ceanna Ruehle, All rights reserved.
"""

# TODO
# Make comments an optional popup
# cancel current request in case of typos?


# Which column should grades be written to? Recall column A= 0, column B= 1...
columnToWriteGradesTo = 1
columnToWriteCommentsTo = 2

import datetime
import xlsxwriter
import xlrd
import re
import os # import miscellaneous operating system interfaces (to find cwd later on)
now = datetime.datetime.now()
date = now.strftime("%m.%d.%Y")
hours, minutes = now.strftime("%H:%M").split(":")
hours = int(hours)
if hours > 12:
    hours -= 12
standardTime = "{}:{}".format(hours, minutes)
militaryTime = now.strftime("%H:%M")
PacPath = os.getcwd() #Determine the current working directory 
rosterLocation=PacPath+'\\roster.xlsx' # assumes this .py and the .xlsx are in the same folder 
rosterWB = xlrd.open_workbook(rosterLocation) 
rosterSheet = rosterWB.sheet_by_index(0) 

print()
print("  Welcome to MNS320 Grader.")
print()
print("  Press 'Enter' twice to begin, enter 'a' for attendance mode, or enter 'i' for instructions.")
category = input("  ")
# Instructions print out:
if category == 'i' or category == 'I':
    print("  Enter the ID of the student or students, then press Enter. This is the acronym in the cell adjacent to the student's name in the roster sheet.")
    print("  You may now enter the grade. Unentered students default to an empty red cell to indicate an absence.")
    print("  If you would like to mark a presenter, you may enter their grade as 'P' or 'p'. The cell will turn blue.")
    print("  You will be notified of invalid ID's without the program crashing, and you my reenter values to override previous entries if you made a typo.")
    print("  To close the program, type 'q' in any field and press Enter.")
    print("  This program depends on the place you stored the roster and where you want to save the new .xlsx file. To change these, you have to go into the code and change them there. They're near the top." )
    print("  One last thing: make sure you don't have a .xlsx file open with the same name you're trying to save to or you're gonna have a bad time.")

# Attendance mode activated: just enter student ids of present students. An excel file will be saved, and absent students will be printed out.
elif category == 'a':
    newWorkbookNamePath = 'C:\\Users\\ceann\\Documents\\School\\Semester 9\\MNS320'
    newWorkbookName = '{}\\attendance_{}2.xlsx'.format(newWorkbookNamePath, date)    
    # Student name, ID, and grades will be stored as class instances of MNS320
    class MNS320():    
        def __init__(self, **kwargs):
            if 'name' in kwargs:
                lastFirstName = kwargs['name']
                self.name = lastFirstName
            if 'id' in kwargs:
                ID = kwargs['id']
                self.ID = ID
            self.presence = ''
    
    # Generating class instances for all the students
    idlist=[]
    for i in range(rosterSheet.nrows): 
        kwargs = {'name': rosterSheet.cell_value(i, 0), 'id': rosterSheet.cell_value(i, 1)}
        idlist.append(rosterSheet.cell_value(i, 1))
        locals()[rosterSheet.cell_value(i, 1)] = MNS320(**kwargs)
    
    # Input will modify the student's presence status
    while True:
        ID = input('Student: ')
        if ID == 'q':
            break
        if ID in idlist:
            locals()[ID].presence = 'P'
        else: print('{} is not a valid student ID. Please try again.'.format(ID))
    
    # Save names and grades to a new excel file
    workbook = xlsxwriter.Workbook(newWorkbookName)
    gradeSheet = workbook.add_worksheet()
    (row, column) = (0, 0)
    for i in range(rosterSheet.nrows):
        # Note that locals()[rosterSheet.cell_value(i, 1)] is just the 
        # student's ID pulled from the roster sheet 
        gradeSheet.write(row, column,                locals()[rosterSheet.cell_value(i, 1)].name)
        gradeSheet.write(row, columnToWriteGradesTo, locals()[rosterSheet.cell_value(i, 1)].presence)
        row += 1
    workbook.close()
    
    attendanceWB = xlrd.open_workbook(newWorkbookName) 
    attendanceSheet = attendanceWB.sheet_by_index(0) 
    
    absences=[]
    
    for i in range(attendanceSheet.nrows): 
        if attendanceSheet.cell_value(i, 1) == '':
            last, first = attendanceSheet.cell_value(i, 0).split('; ')
            absences.append(first+' '+last)
    print('I noted who was able to make the zoom meeting today at {}. There were {} students who were not logged into Zoom at that time:'.format(standardTime, len(absences)))
    
    for i in absences:
        print(i)

# Grading mode is the default
else:    
    name = input(" What do you want to save your file as? Don't include the path or file type.\n\n")
    newWorkbookNamePath = PacPath 
    newWorkbookName = '{}\\{}.xlsx'.format(newWorkbookNamePath, name)
    
    # Student name, ID, and grades will be stored as class instances of MNS320
    class MNS320():    
        def __init__(self, **kwargs):
            if 'name' in kwargs:
                lastFirstName = kwargs['name']
                self.name = lastFirstName
            if 'id' in kwargs:
                ID = kwargs['id']
                self.ID = ID
            self.grade = ''
            self.comments = ''
    
    # Generating class instances for all the students
    idlist=[]
    for i in range(rosterSheet.nrows): 
        kwargs = {'name': rosterSheet.cell_value(i, 0), 'id': rosterSheet.cell_value(i, 1)}
        idlist.append(rosterSheet.cell_value(i, 1))
        locals()[rosterSheet.cell_value(i, 1)] = MNS320(**kwargs)

    # Input will modify the student's grade
    while True:
        # Split input into discrete student entries (for multiple student ID entries)
        ID = re.split(r'[;,\s]\s*',input('Student(s): '))
        check, valid = True, True
        if ID == ['q'] or ID == ['']:
            break
        for eachID in ID:
            if eachID not in idlist:
                print('{} is not a valid student ID. Please try again.'.format(eachID))
                valid = False
        if valid == True:
            for eachID in ID:
                    if locals()[eachID].grade!='':
                        print("Overriding ID "+eachID)
            grade = input('Grade: ')
            if grade == 'q':
                break
            comments = input('Comments?   ')
            if comments == 'q':
                break
            for eachID in ID:
                if eachID in idlist:
                    locals()[eachID].grade = grade
                    locals()[eachID].comments = comments

    # Save names and grades to a new excel file
    workbook = xlsxwriter.Workbook(newWorkbookName)
    gradeSheet = workbook.add_worksheet()
    (row, column) = (0, 0)
    for i in range(rosterSheet.nrows):
        # Note that locals()[rosterSheet.cell_value(i, 1)] is just the 
        # student's ID pulled from the roster sheet 
        gradeSheet.write(row, column, locals()[rosterSheet.cell_value(i, 1)].name)
        # If the grade is not marked, the cell is red #f4cccc to mark an absence
        if locals()[rosterSheet.cell_value(i, 1)].grade == '':
            cell_format = workbook.add_format()        
            cell_format.set_bg_color('#f4cccc')
            gradeSheet.write(row, columnToWriteGradesTo, locals()[rosterSheet.cell_value(i, 1)].grade, cell_format)
        # If the grade is marked as 'p' or 'P', the cell is blue #c9daf8 to mark a presenter
        elif locals()[rosterSheet.cell_value(i, 1)].grade == 'p' or locals()[rosterSheet.cell_value(i, 1)].grade == 'P':
            cell_format = workbook.add_format()        
            cell_format.set_bg_color('#c9daf8')
            gradeSheet.write(row, columnToWriteGradesTo,'', cell_format)
        else:
            gradeSheet.write(row, columnToWriteGradesTo,   locals()[rosterSheet.cell_value(i, 1)].grade)
            gradeSheet.write(row, columnToWriteCommentsTo, locals()[rosterSheet.cell_value(i, 1)].comments)
        row += 1
    workbook.close()
    print("Grade sheet saved to '{}'".format(newWorkbookName))
