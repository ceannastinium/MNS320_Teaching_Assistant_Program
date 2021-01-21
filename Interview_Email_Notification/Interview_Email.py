
# -*- coding: utf-8 -*-
"""
Created on Fri Feb  7 12:42:04 2020
@author: Ceanna Ruehle
"""

order = 14
FirstRun = True
ChangeOrder = False

print("\nWelcome to MNS320 Auto Email Generator.\n")

while True:
        
    import pyperclip # import module to copy text to clipboard; note this one requires pip if it's not intalled on the computer already
    import xlrd # import excel spreadsheet reading capabilities
    import os # import miscellaneous operating system interfaces (to find cwd later on)
    import datetime 
    # now = datetime.datetime.now()
    # date = now.strftime("%m.%d.%Y")
    
    
    PacPath = os.getcwd() #Determine the current working directory 
    loc=PacPath+'\\Interview_Tracking.xlsx' # assumes this .py and the .xlsx are in the same folder 
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_name('Groups')
    
    
    # Determine relationship between order and rows in spreadsheet
    Group_list = []
    for c in range(1,max(range(sheet.nrows))+1):
        if sheet.cell_value(c, 2) != '':
            Group_list.append(int(c))
    i = Group_list[order-1]
    
    # Team number
    team=int(sheet.cell_value(i, 2))
        
    # Pull interviewee's information from Excel and format to full and short versions (eg, 'Dr. Dong-Ha Min' and 'Dr. Min')
    dialogistFull = sheet.cell_value(i, 4)+" "+sheet.cell_value(i, 5)+" "+sheet.cell_value(i, 6)
    dialogistShort = sheet.cell_value(i, 4)+" "+sheet.cell_value(i, 6)
    
    # Pull date information from Excel and translate 
    year, monthnum, daynum = xlrd.xldate_as_tuple(sheet.cell_value(i, 3),1)[:3]
    month = datetime.date(year, monthnum, daynum).strftime('%B')
    day = datetime.date(year, monthnum, daynum).strftime('%A')
    date = day+', '+month+' '+str(daynum)
    
    if FirstRun == True:
    
        print("Enter 'm' for the text of the email, 'r' for the recipients of the email, 'c' for the cc'd, "
              "'s' for the subject line, 'o' to change the order, or 'q' to quit\n")
        print('The order is set to {}: Team {} presenting on {}'.format(order,team, date))
        FirstRun = False
        
    if ChangeOrder == True:
        print('\nThe order is set to {}: Team {} presenting on {}'.format(order,team, date))
    
    
    category = input("  ")
    
    # NOTE - add case for if there's only one TA, maybe separate excel sheet w TA information? incl OH
    # If you do this, remember the "we" and "us" references in the text or risk sounding like a psycho
    
    # How to sign 
    author = "Ceanna"
    # coTA sample, <coTA = "Sojin"> for one other TA or <coTa =  "Victoria, Kyle,"> for multiple coTAs
    coTA = "Victoria, Kyle,"
        
    if category.lower() == 'm':    
        
        pyperclip.copy('Hey Team {}!\n\n'
        'You are set to interview {} on {}. I am emailing you so you can begin working on the project '
        'and setting up a time to meet with {} or me.\n\n'
        'Just a bit of helpful guidance to get you started. I recommend that you check out Canvas->Modules->Group Project: '
        'Expert Interviews->NPR radio interview samples and some suggestions. This should give you an idea of general logistics '
        'while preparing the project.\n\n'
        'I also recommend that everyone reads and familiarizes yourself with {}\'s research. You can start delegating who will '
        'give an introduction, debrief, transcribe, etc. Also delegate where everyone wants to focus their attention and what '
        'types of questions you want to ask. Once you have decided on general roles, read the papers you have decided to study. '
        'Choose RECENT papers in which {} is listed as the first author, or listed early on (links to the papers should be '
        'provided on CV or website). Form a Google Doc to draft questions. This does not need to be extremely thought out before '
        'you meet with us, just a general idea so that we can help guide it in the right direction. It is okay to ask multiple '
        'questions from the same paper.\n\n '
        'We are most available during office hours, but if that doesn\'t work for you, just send us an email with your availability.\n'
        'Our office hours are:\n'
        'Victoria: Tue/Thu 11am-Noon\n'
        'Kyle: Thu 3-5pm\n'
        'Ceanna: Tue 1-2pm & 3-4pm\n'
        '\nChat with each other and see if there is a time that could work for everyone (or almost everyone) and let us know.\n\n'
        'Any additional questions, don\'t hesitate to email me back!\n\n'
        'Best regards,\n'
        '{}'.format(team,dialogistFull,date,coTA,dialogistShort, dialogistShort, author))
    
        print('\nMessage to team {} copied to clipboard!\n'.format(team))
        
        #break
    
    elif category.lower() == 'r':
        print('\nTeam {} (presenting on {}) email addresses:\n'.format(team, date))
        print(sheet.cell_value(i,1),'\n',sheet.cell_value(i+1,1),'\n',sheet.cell_value(i+2,1),'\n',sheet.cell_value(i+3,1))
        pyperclip.copy(str(sheet.cell_value(i,1)+' '+sheet.cell_value(i+1,1)+' '+sheet.cell_value(i+2,1)+' '+sheet.cell_value(i+3,1)))
        
        #break
    
    elif category.lower() == 'c':
        print('\nTA email addresses:')
        print()
        print('vmcongdon@utexas.edu, kyle.capistrantfossa@utexas.edu\n')
        pyperclip.copy('vmcongdon@utexas.edu, kyle.capistrantfossa@utexas.edu')
        
        #break
        
    elif category.lower() == 's':
        print('\nMNS320 Interview with {}\n'.format(dialogistFull))
        pyperclip.copy('MNS320 Interview with {}'.format(dialogistFull))
        
    elif category.lower() == 'q':
        print('\nclosing...')
        
        break
        
    elif category.lower() == 'o':
        order = int(input('Which order?\n\n  '))
        ChangeOrder = True
        
    else:
        print('\nPlease enter a valid command\n')
