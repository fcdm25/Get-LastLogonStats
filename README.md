LastLogonUsers
==============
################################################################################################################################################################
# Script accepts 2 parameters from the command line
#
# Office365Username - Mandatory - Administrator login ID for the tenant we are querying
# Office365Password - Mandatory - Administrator login password for the tenant we are querying
#
# To run the script
#
# .\Get-LastLogonStats.ps1 -Office365Username admin@contoso.com -Office365Password Password
#
# NOTE: If you do not pass an input file to the script, it will return the last logon time of ALL mailboxes in the tenant.  Not advisable for tenants with large
# user count (< 3,000) 
#
# Author: 				Alan Byrne
# Version: 				2.0
# Last Modified Date: 	20.09.2014
# Last Modified By: 	Alexander Makarov aka fcdm25 (aa_makarov@guu.ru)
#                       Alexander Zubarev aka strike (av_zubarev@guu.ru) 
################################################################################################################################################################
