# Ceanna Ruehle

# Relies on .csv file named 'roster.csv' in the same file as this .py
# This .csv should be of the following format to work here without edits:
# column A: LastName; FirstName
# column D: UT EID email

# Also this script won't work if your class is going on during midnight, but
# for the sake of anyone involved with that, hopefully that won't matter.


#%% Imported modules

import pandas as pd
import os
import re
import datetime

#%% Parameters which may change with each check

# Select which Zoom report you want to analyze.
zoom_report_file_name = 'C:/Users/ceann/Downloads/zoomus_meeting_report_94986659321.csv'

# Note datetime follows 24-hr time. So if class starts at 3:30pm and you want a 
# 5 minute buffer for late arrivals, you would specify datetime.time(15, 35).
class_start           = datetime.time(15, 35)
class_end             = datetime.time(16, 30)

# How much time is significant for loss of connectivity?
significant_departure = 45 # [s]

#%%  Read data from the zoom meeting report to put in lists

zoom_report           = pd.read_csv(zoom_report_file_name)
zoom_report.columns   = ['Name','Email','Join','Leave','Duration','Alt']
zoom_report_name      = zoom_report.Name.to_list()
zoom_report_email     = zoom_report.Email.to_list()
join                  = zoom_report.Join.to_list()
leave                 = zoom_report.Leave.to_list()

#%% Read data from the roster to put in lists

current_working_dir   = os.getcwd() #Determine the current working directory 
roster_location       = current_working_dir+'\\roster.csv' # assumes this .py and the .csv are in the same folder 
roster                = pd.read_csv(roster_location)
roster.columns        = ['Name','ID','Grade','Email','Notes']
roster_email          = roster.Email.to_list()
name                  = roster.Name.to_list()

#%% Functions

def time_in_seconds(time):     # takes datetime argument and returns time in seconds as int
    time_in_seconds = int(time.strftime('%H'))*3600+int(time.strftime('%M'))*60+int(time.strftime('%S'))
    return time_in_seconds


def zoom_to_datetime(z):     # takes date and time from Zoom format and returns it in datetime
    jointime  = datetime.datetime(int(z[2]), int(z[0]), int(z[1]), int(z[3]), int(z[4]), int(z[5]))
    return jointime


def zoom_to_time(z):          # takes time from Zoom format and returns it in datetime
    jointime  = datetime.time(int(z[3]), int(z[4]), int(z[5]))
    return jointime


#%% Class

# Student name, ID, and attendance will be stored as class instances of StudentData
class StudentData():    
    def __init__(self, **kwargs):
        if 'name' in kwargs:
            kwname = kwargs['name']
            self.name = kwname
        if 'id' in kwargs:
            ID = kwargs['id']
            self.ID = ID
        if 'Email' in kwargs:
            email = kwargs['Email']
            self.email = email
        self.joined = []
        self.left = []
        
#TODO: __lt__ should enable .sort() for sorting entries by time rather
#      than alphabetically as with the roster. 
    def __lt__(self, other):
         return self.joined[0] < other.joined[0]


#%% Create class instances named for every email in the roster.
#   Each instance is named as the email without punctuation. For example,
#   Matthew McConaughey's email 'mdm659@utexas.edu' would give instance 'mdm659utexasedu'.
#   For the University of Texas system, the eids are unique to each member, so 
#   redundacy should not be an issue.

registered_students = [] 
instances = []

for i in range(len(roster_email)):
    kwargs = {'name': name[i], 'email': roster_email[i], 'id': ''.join(re.split(r'\W+', str(roster_email[i])))}
    locals()[''.join(re.split(r'\W+', str(roster_email[i])))] = StudentData(**kwargs)
    
    # Make a list of the class instances as strings for later reference
    registered_students.append(''.join(re.split(r'\W+', str(roster_email[i])))) 

    # And as full instances
    instances.append(locals()[''.join(re.split(r'\W+', str(roster_email[i])))]) 

#%% Create a list of the attendees as strings (without repetition) including 
#   unregistered people such as the professor or TAs.

attendees = []
[attendees.append(''.join(re.split(r'\W+', str(zoom_report_email[i])))) for i in range(len(zoom_report_email)) if ''.join(re.split(r'\W+', str(zoom_report_email[i]))) not in attendees]



#TODO: The list of present students should help for sorting entries by time rather
#      than alphabetically as with the roster. 

#   Create a list of the attendees as instances (without repetition) excluding
#   unregistered people such as the professor or TAs.

present_students = []
for attendee in attendees:
    if attendee in registered_students:
        present_students.append(locals()[attendee])


#%% Edit the attributes of the class instances based on the zoom report on
#   when the students arrived and left the Zoom meeting.

for i in range(len(zoom_report_email)): 
        
    # Ignore the email if it's not in the list of students.
    if ''.join(re.split(r'\W+', str(zoom_report_email[i]))) in registered_students:
    
        # Temporary variable 'ClassInstance' indicates the class instance to avoid cumbersome references
        ClassInstance = locals()[''.join(re.split(r'\W+', str(zoom_report_email[i])))]
        
        # Temporary variables identify the times of login and logoff 
        jointime  = zoom_to_time(re.split(r'\W+',  join[i]))
        leavetime = zoom_to_time(re.split(r'\W+', leave[i])) # What goes in must come out...
            
        # If the student is joining the meeting for the first time, fill out their record.
        if  ClassInstance.joined == []:
            ClassInstance.joined.append( jointime  )
            ClassInstance.left.append(   leavetime )
            
        # If the student is not joining the meeting for the first time, then the 
        # existing record is adjusted.
        else:
            minidur = time_in_seconds(jointime) - time_in_seconds(locals()[''.join(re.split(r'\W+', str(zoom_report_email[i])))].left[-1])
            
            # If the difference between this join time and the last leave time < 30 seconds, 
            # then count it as a single duration. This is because Zoom counts multiple 
            # join times / leave times even if the break in connectivity is negligible. 
            if abs(minidur) < significant_departure:  
                # Get rid of last leave time in favor of the new one
                ClassInstance.left.pop(-1)
                ClassInstance.left.append( leavetime )
                # Keep last join
    
            # If the difference between sessions is significant, then the break 
            # is recorded. 
            else:
                ClassInstance.joined.append( jointime  )
                ClassInstance.left.append(   leavetime )

            
#%% Check who was fully absent. Then print to console.
                
#   Putting this for loop before checking tardiness means the absentees will
#   be printed out at the top of the console.

for i in range(len(roster_email)): 
    if ''.join(re.split(r'\W+', str(roster_email[i]))) not in attendees:
        
        # Output will be the student's name
        last, first = name[i].split('; ')
        fullname = first+' '+last
        
        print('Absent: '+str(fullname))
print()

#%% Check who had arrival(s) after the (buffered) start time and departure(s) 
#   before the (buffered) end time. Then print to console.

for i in range(len(roster_email)): # for each person in the roster

    redundant = False
    
    # Temporary variable 'ClassInstance' indicates the class instance to avoid cumbersome references
    ClassInstance = locals()[''.join(re.split(r'\W+', str(roster_email[i])))]
    
    # Output will be the student's name
    last, first = name[i].split('; ')
    fullname = first+' '+last    
    
    # Make sure the student in question wasn't already marked as absent.
    if ''.join(re.split(r'\W+', str(roster_email[i]))) in attendees:        
    
        # Check if the first login time is after class started.
        if time_in_seconds(ClassInstance.joined[0]) > time_in_seconds(class_start):
            print(fullname+' joined at '+ClassInstance.joined[0].strftime("%H:%M:%S"))
        
        # Check if there are multiple logins.
        if len(ClassInstance.joined) > 1:
            for i in range(len(ClassInstance.joined)): 
                print(fullname+' left at '+ClassInstance.left[i].strftime("%H:%M:%S"))
            if time_in_seconds(ClassInstance.left[-1]) < time_in_seconds(class_end):
                redundant = True
        
        # Check if the logoff time was before class ended.
        if time_in_seconds(ClassInstance.left[-1]) < time_in_seconds(class_end) and redundant == False:
            print(fullname+' left at '+ClassInstance.joined[0].strftime("%H:%M:%S"))
        
     




        
        
        
        
        
        
        
        
        
        