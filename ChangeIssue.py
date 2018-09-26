# Jira auhentications
#
# 26.9.2018  mika.nokka1@gmail.com
# 
# To be used via importing only
# 

import datetime 
import time
import argparse
import sys
import netrc
import requests, os
from requests.auth import HTTPBasicAuth
# We don't want InsecureRequest warnings:
import requests
requests.packages.urllib3.disable_warnings()
import itertools, re, sys
from jira import JIRA
import random


__version__ = "0.1"
thisFile = __file__

    
def main(argv):

    JIRASERVICE=""
    JIRAPROJECT=""
    JIRASUMMARY=""
    JIRADESCRIPTION=""
    PSWD=''
    USER=''
    
    parser = argparse.ArgumentParser(usage="""
    {1}    Version:{0}     -  mika.nokka1@gmail.com
    

    """.format(__version__,sys.argv[0]))

    parser.add_argument('-p','--project', help='<JIRA project key>')
    parser.add_argument('-j','--jira', help='<Target JIRA address>')
    parser.add_argument('-v','--version', help='<Version>', action='store_true')

    parser.add_argument('-ps','--password', help='<JIRA password>')
    parser.add_argument('-u','--user', help='<JIRA user>')
    
    args = parser.parse_args()
        
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
         

    JIRASERVICE = args.jira or ''
    JIRAPROJECT = args.project or ''

    PSWD= args.password or ''
    USER= args.user or ''
  
    # quick old-school way to check needed parameters
    if (JIRASERVICE=='' or  JIRAPROJECT=='' or  PSWD=='' or USER==''):
        parser.print_help()
        sys.exit(2)

    print "User:{0}, PS:{1}, Service:{2}".format(USER,PSWD,JIRASERVICE)
    user, PASSWORD = Authenticate(JIRASERVICE,PSWD,USER)
    jira= DoJIRAStuff(user,PASSWORD,JIRASERVICE)
   
    
####################################################################################################   
# POC skips .netrc usage
# 
def Authenticate(JIRASERVICE,PSWD,USER):
    host=JIRASERVICE
    user=USER
    PASSWORD=PSWD
    
    
    f = requests.get(host,auth=(user, PASSWORD))
         
    # CHECK WRONG AUTHENTICATION    
    header=str(f.headers)
    HeaderCheck = re.search( r"(.*?)(AUTHENTICATION_DENIED|AUTHENTICATION_FAILED|AUTHENTICATED_FAILED)", header)
    if HeaderCheck:
        CurrentGroups=HeaderCheck.groups()    
        print ("Group 1: %s" % CurrentGroups[0]) 
        print ("Group 2: %s" % CurrentGroups[1]) 
        print ("Header: %s" % header)         
        print "Authentication FAILED - HEADER: {0}".format(header) 
        print "--> ERROR: Apparantly user authentication gone wrong. EXITING!"
        sys.exit(1)
    else:
        print "Authentication OK \nHEADER: {0}".format(header)    
    print "---------------------------------------------------------"
    return user,PASSWORD

###################################################################################    
def DoJIRAStuff(user,PASSWORD,JIRASERVICE):
 jira_server=JIRASERVICE
 try:
     print("Connecting to JIRA: %s" % jira_server)
     jira_options = {'server': jira_server}
     jira = JIRA(options=jira_options,basic_auth=(user,PASSWORD))
     print "JIRA Authorization OK"
 except Exception,e:
    print("Failed to connect to JIRA: %s" % e)
 return jira   
    


############################################################################################'
#

        
if __name__ == "__main__":
        main(sys.argv[1:])
        