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

# should pass via  parameters
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
    
    Parse(JIRASERVICE,JIRAPROJECT,PSWD,USER,excelfilepath,filename,ENV,jira)



############################################################################################################################################
# Parse attachment files and add to matching Jira issue
#

#NOTE: Uses hardcoded sheet/column value

def Parse(JIRASERVICE,JIRAPROJECT,PSWD,USER,excelfilepath,filename,ENV,jira):

    files=excelfilepath+"/"+filename
    logging.debug ("Excel file:{0}".format(files))
   
  
   
    Issues=defaultdict(dict) 
   
    #main excel definitions
    MainSheet="Sheet0" 
    wb= openpyxl.load_workbook(files)
    #types=type(wb)
    #logging.debug ("Type:{0}".format(types))
    #sheets=wb.get_sheet_names()
    #logging.debug ("Sheets:{0}".format(sheets))
    CurrentSheet=wb[MainSheet] 
    logging.debug ("CurrentSheet:{0}".format(CurrentSheet))
    logging.debug ("First key:{0}".format(CurrentSheet['B4'].value))
    logging.debug ("First Area code:{0}".format(CurrentSheet['C4'].value))
   

    ########################################
    #CONFIGURATIONS AND EXCEL COLUMN MAPPINGS
    DATASTARTSROW=4 # data section starting line 
    #A=1 # issue key
    B=2 # issue key
    C=3 # Area code, to be owerwritten if something exists in the issue
    

    #print Issues.items() 
    
    #key=18503 # check if this key exists
    #if key in Issues:
    #    print "EXISTS"
    #else:
    #    print "NOT THERE"
    #for key, value in Issues.iteritems() :
    #    print key, value

        
    ### MAIN EXCEL ###########################################################################################
    #Go through main excel sheet for main issue keys (and contents findings)
    # Create dictionary structure (remarks)
    # NOTE: Uses hardcoded sheet/column values
    # NOTE: As this handles first sheet, using used row/cell reading (buggy, works only for first sheet) 
    #
    i=DATASTARTSROW # brute force row indexing
    for row in CurrentSheet[('B{}:B{}'.format(DATASTARTSROW,CurrentSheet.max_row))]:  # go trough all column A(Issue KEY) rows
        for mycell in row:
            KEY=mycell.value
            logging.debug("ROW:{0}     Issue-key:{1}".format(i,mycell.value))
            AREA=(CurrentSheet.cell(row=i, column=C).value)
            logging.debug("             Area code:{1}".format(i,AREA))
            
            myissue=123
            for issue in jira.search_issues("project=NB1394DM  and issuekey = {0}".format(KEY), maxResults=10):
                #bug: if more than one match will fail
                myissuekey=format(issue.key)
                logging.debug("Jira issue key (from Jira): {0}".format(myissuekey))
                #logging.debug("ISSUE: {0}:".format(issue))
                #logging.debug("ID{0}: ".format(issue.id))
              
                myissueareavalue=issue.fields.customfield_10007
                logging.debug("Current Jira  value: {0}:".format(myissueareavalue))
                if (myissueareavalue is None):
                    logging.debug("*** No previous Area value ****")
                    logging.debug("*** Setting initial value as:{0}".format(AREA))
                else:                      
                    logging.debug("Current Jira Area value: {0}:".format(myissueareavalue[0]))
                    logging.debug("SHOULD overwrite {0} ----> {1}".format(myissueareavalue[0],AREA))
             
                #issue.update(customfield_10007=AREA)
                issue.update(fields={'customfield_10007': [{'value':AREA}]}) #   .format(AREA))
                #if (myissuekey=="NB1400DM-1936"):
                #    logging.debug("Found: {0}".format(myissuekey))
                #    issue.update(customfield_10019=NEW_DRWNMB)
                logging.debug("---------------------------------------------------")
                sys.exit(5)
            i=i+1
            
    #for issue in jira.search_issues('project=NB1400DM  and issuekey = NB1400DM-1165', maxResults=10):
    #    logging.debug("{}: {}".format(issue.key, issue.fields.summary))
           
       
         
  

    
    
  
        
    end = time.clock()
    totaltime=end-start
    print "Time taken:{0} seconds".format(totaltime)
       
            
    print "*************************************************************************"
        

 
  

    
logging.debug ("--Python exiting--")
if __name__ == "__main__":
    main(sys.argv[1:]) 