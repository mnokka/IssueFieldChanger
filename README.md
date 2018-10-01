# IssueFieldChanger , working POC
Use excel provided Jira task list (with changed value field) to change this field in Jira (to excel defined issue) 

Usage: python reader.py -p PROJECTNAME -s JIRAURL -u USERNAME -w PASSWORD -q EXCELPATH -n  EXCELNAME


Code defines excel columns for Issue-Key, Current custom field value and new value for the current custom field
Branched in order to set new value 
