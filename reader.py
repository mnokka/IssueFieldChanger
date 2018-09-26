#encoding=utf8

# POC tool to change issue fields to existing Jira issues
#Excel provides info which issues (key) issue field is to be changed
# Forked and modified from https://github.com/mnokka/ExcelReader
#
# Author mika.nokka1@gmail.com 
#

#
#from __future__ import unicode_literals

import openpyxl 
import sys, logging
import argparse
#import re
from collections import defaultdict
from ChangeIssue import Authenticate  # no need to use as external command
from ChangeIssue import DoJIRAStuff

import glob
import re
import os
import time
import unidecode


start = time.clock()
__version__ = u"0.1"

# should pass via parameters
#ENV="demo"
ENV=u"PROD"

logging.basicConfig(level=logging.DEBUG) # IF calling from Groovy, this must be set logging level DEBUG in Groovy side order these to be written out



def main(argv):
    
    JIRASERVICE=u""
    JIRAPROJECT=u""
    PSWD=u''
    USER=u''
  
    logging.debug (u"--Python starting checking excel to change Jira issue field value --") 

 
    parser = argparse.ArgumentParser(usage="""
    {1}    Version:{0}     -  mika.nokka1@gmail.com
    
    USAGE:
    -filepath  | -p <Path to Excel file directory>
    -filename   | -n <Excel filename>

    """.format(__version__,sys.argv[0]))

    # parser.add_argument('-f','--filepath', help='<Path to attachment directory>')
    parser.add_argument('-q','--excelfilepath', help='<Path to excel directory>')
    parser.add_argument('-n','--filename', help='<Issues to be changed Excel filename>')
    # parser.add_argument('-m','--subfilename', help='<Subtasks Excel filename>')
    parser.add_argument('-v','--version', help='<Version>', action='store_true')
    
    parser.add_argument('-w','--password', help='<JIRA password>')
    parser.add_argument('-u','--user', help='<JIRA user>')
    parser.add_argument('-s','--service', help='<JIRA service>')
    parser.add_argument('-p','--project', help='<JIRA project>')
    #parser.add_argument('-z','--rename', help='<rename files>') #adhoc operation activation
    #parser.add_argument('-x','--ascii', help='<ascii file names>') #adhoc operation activation
        
    args = parser.parse_args()
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
           
    #filepath = args.filepath or ''
    excelfilepath = args.excelfilepath or ''
    filename = args.filename or ''
    #subfilename=args.subfilename or ''
    
    JIRASERVICE = args.service or ''
    JIRAPROJECT = args.project or ''
    PSWD= args.password or ''
    USER= args.user or ''
    #RENAME= args.rename or ''
    #ASCII=args.ascii or ''
    
    # quick old-school way to check needed parameters
    if (JIRASERVICE=='' or  JIRAPROJECT==''  or PSWD=='' or USER=='' or  excelfilepath=='' or filename==''):
        parser.print_help()
        sys.exit(2)
        
     
    Authenticate(JIRASERVICE,PSWD,USER)
    jira=DoJIRAStuff(USER,PSWD,JIRASERVICE)
    
   # Parse(JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME,subfilename,excelfilepath,filename,ENV)



############################################################################################################################################
# Parse attachment files and add to matching Jira issue
#

#NOTE: Uses hardcoded sheet/column value

def Parse(JIRASERVICE,JIRAPROJECT,PSWD,USER,RENAME,subfilename,excelfilepath,filename,ENV):

    files=excelfilepath+"/"+filename
    logging.debug ("File:{0}".format(files))
   
    Issues=defaultdict(dict) 
   
    #main excel definitions
    MainSheet="general_report" 
    wb= openpyxl.load_workbook(files)
    #types=type(wb)
    #logging.debug ("Type:{0}".format(types))
    #sheets=wb.get_sheet_names()
    #logging.debug ("Sheets:{0}".format(sheets))
    CurrentSheet=wb[MainSheet] 
    #logging.debug ("CurrentSheet:{0}".format(CurrentSheet))
    #logging.debug ("First row:{0}".format(CurrentSheet['A4'].value))

   
    #subtasks excel definitions
    logging.debug ("ExcelFilepath: %s     ExcelFilename:%s" %(excelfilepath ,subfilename))
    subfiles=excelfilepath+"/"+subfilename
    logging.debug ("SubFiles:{0}".format(subfiles))
   
    
    SubMainSheet="general_report" 
    subwb= openpyxl.load_workbook(subfiles)
    #types=type(wb)
    #logging.debug ("Type:{0}".format(types))
    #sheets=wb.get_sheet_names()
    #logging.debug ("Sheets:{0}".format(sheets))
    SubCurrentSheet=subwb[SubMainSheet] 
    #logging.debug ("CurrentSheet:{0}".format(CurrentSheet))
    #logging.debug ("First row:{0}".format(CurrentSheet['A4'].value))
   
   
   
    

    
    
    
    ########################################
    #CONFIGURATIONS AND EXCEL COLUMN MAPPINGS, both main and subtask excel
    DATASTARTSROW=5 # data section starting line MAIN TASKS EXCEL
    DATASTARTSROWSUB=5 # data section starting line SUB TASKS EXCEL
    C=3 #SUMMARY
    D=4 #Issue Type
    E=5 #Status Always "Open"    
    G=7 #ResponsibleNW
    H=8 #Creator
    I=9 #Inspection date --> Original Created date in Jira Changed as Inspection Date
    J=10 # Subtask TASK-ID
    K=11 #system number, subtasks excel 
    M=13 #Shipnumber
    N=14 #system number
    P=16 #PerformerNW
    Q=17 #Performer, subtask excel
    R=18 #Responsible ,subtask excel
    #U=20 #Responsible Phone Number --> Not taken, field just exists in Jira
    S=19 #DepartmentNW
    V=22 #Deck
    W=23 #Block
    X=24 # Firezone
    AA=27 #Subtask DeckNW
   

    
   
    #print Issues.items() 
    
    #key=18503 # check if this key exists
    #if key in Issues:
    #    print "EXISTS"
    #else:
    #    print "NOT THERE"
    #for key, value in Issues.iteritems() :
    #    print key, value

    #
               
    


    ##########################################################################################################################
    
    
    
    Authenticate(JIRASERVICE,PSWD,USER)
    jira=DoJIRAStuff(USER,PSWD,JIRASERVICE)

    
   
        
    ### MAIN EXCEL ###########################################################################################
    #Go through main excel sheet for main issue keys (and contents findings)
    # Create dictionary structure (remarks)
    # NOTE: Uses hardcoded sheet/column values
    # NOTE: As this handles first sheet, using used row/cell reading (buggy, works only for first sheet) 
    #
    i=DATASTARTSROW # brute force row indexing
    for row in CurrentSheet[('B{}:B{}'.format(DATASTARTSROW,CurrentSheet.max_row))]:  # go trough all column B (KEY) rows
        for mycell in row:
            KEY=mycell.value
            logging.debug("ROW:{0} Original ID:{1}".format(i,mycell.value))
            Issues[KEY]={} # add to dictionary as master key (KEY)
            
            #Just hardocode operations, POC is one off
            #LINKED_ISSUES=(CurrentSheet.cell(row=i, column=K).value) #NOTE THIS APPROACH GOES ALWAYS TO THE FIRST SHEET
            #logging.debug("Attachment:{0}".format((CurrentSheet.cell(row=i, column=K).value))) # for the same row, show also column K (LINKED_ISSUES) values
            #Issues[KEY]["LINKED_ISSUES"] = LINKED_ISSUES
            
            SUMMARY=(CurrentSheet.cell(row=i, column=C).value)
            if not SUMMARY:
                SUMMARY="Summary for this task has not been defined"
            Issues[KEY]["SUMMARY"] = SUMMARY
            
            ISSUE_TYPE=(CurrentSheet.cell(row=i, column=D).value)
            Issues[KEY]["ISSUE_TYPE"] = ISSUE_TYPE
            
            STATUS=(CurrentSheet.cell(row=i, column=E).value)
            Issues[KEY]["STATUS"] = STATUS
            
            RESPONSIBLE=(CurrentSheet.cell(row=i, column=G).value)
            Issues[KEY]["RESPONSIBLE"] = RESPONSIBLE.encode('utf-8')
            
            #REPORTER=(CurrentSheet.cell(row=i, column=G).value)
            #Issues[KEY]["REPORTER"] = REPORTER
            
            
            CREATOR=(CurrentSheet.cell(row=i, column=H).value)
            Issues[KEY]["CREATOR"] = CREATOR
            
            CREATED=(CurrentSheet.cell(row=i, column=I).value) #Inspection date
            # ISO 8601 conversion to Exceli time
            time2=CREATED.strftime("%Y-%m-%dT%H:%M:%S.000-0300")  #-0300 is UTC delta to Finland, 000 just keeps Jira happy
            print "CREATED ISOFORMAT TIME2:{0}".format(time2)
            CREATED=time2
            INSPECTED=CREATED # just reusing value
            Issues[KEY]["INSPECTED"] = INSPECTED
            
            
            SHIP=(CurrentSheet.cell(row=i, column=M).value)
            Issues[KEY]["SHIP"] = SHIP
            
            PERFORMER=(CurrentSheet.cell(row=i, column=P).value)
            Issues[KEY]["PERFORMER"] = PERFORMER # .encode('utf-8')
            
              
            #RESPHONE=(CurrentSheet.cell(row=i, column=U).value)
            #Issues[KEY]["RESPHONE"] = RESPHONE
            
            DEPARTMENT=(CurrentSheet.cell(row=i, column=S).value)
            Issues[KEY]["DEPARTMENT"] = DEPARTMENT
            
            DECK=(CurrentSheet.cell(row=i, column=V).value)
            Issues[KEY]["DECK"] = DECK
            
            BLOCK=(CurrentSheet.cell(row=i, column=W).value)
            Issues[KEY]["BLOCK"] = BLOCK
            
            FIREZONE=(CurrentSheet.cell(row=i, column=X).value)
            Issues[KEY]["FIREZONE"] = FIREZONE
            
                
            SYSTEMNUMBER=(CurrentSheet.cell(row=i, column=N).value)
            Issues[KEY]["SYSTEMNUMBER"] = SYSTEMNUMBER
            
            
            
            
            #Create sub dictionary for possible subtasks (to be used later)
            Issues[KEY]["REMARKS"]={}
            
            logging.debug("---------------------------------------------------")
            i=i+1
    #print Issues
    #print Issues.items() 
    
    #key=18503 # check if this key exists
    #if key in Issues:
    #    print "EXISTS"
    #else:
    #    print "NOT THERE"
    #for key, value in Issues.iteritems() :
    #    print key, value
   
        
   
           
       
         
  

    
    
  
        
    end = time.clock()
    totaltime=end-start
    print "Time taken:{0} seconds".format(totaltime)
       
            
    print "*************************************************************************"
        

 
  

    
logging.debug ("--Python exiting--")
if __name__ == "__main__":
    main(sys.argv[1:]) 