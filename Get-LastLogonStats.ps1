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

#Accept input parameters
Param(
	[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $Office365Username,
	[Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $Office365Password
)

#Constant Variables
$OutputFile = "LastLogonDate.csv"   #The CSV Output file that is created, change for your purposes

#Main
Function Main {
    $results = @()

	#Remove all existing Powershell sessions
	Get-PSSession | Remove-PSSession
	
	#Call ConnectTo-ExchangeOnline function with correct credentials
	ConnectTo-ExchangeOnline -Office365AdminUsername $Office365Username -Office365AdminPassword $Office365Password			
	
	#Prepare Output file with headers
	#echo "UserPrincipalName,LastLogonDate" > $OutputFile
	
	#Check if we have been passed an input file path
	if ($userIDFile -ne "")
	{
		#We have an input file, read it into memory
		$objUsers = import-csv -Header "UserPrincipalName" $UserIDFile
	}
	else
	{
		#No input file found, gather all mailboxes from Office 365
		$objUsers = get-mailbox -ResultSize Unlimited | select UserPrincipalName
	}
	
	#Iterate through all users
    $counter = 0 	
    $counter_all=$objUsers.length

	Foreach ($objUser in $objUsers)
	{	
        $counter += 1
		#Connect to the users mailbox
		$objUserMailbox = get-mailboxstatistics -Identity $($objUser.UserPrincipalName) | Select LastLogonTime, ItemCount, DisplayName
		#Prepare UserPrincipalName variable
		$strUserPrincipalName = $objUser.UserPrincipalName
       
        echo "$strUserPrincipalName `t($counter/$counter_all)"
		
		#Check if they have a last logon time. Users who have never logged in do not have this property
		if ($objUserMailbox.LastLogonTime -eq $null)
		{
			#Never logged in, update Last Logon Variable
			$strLastLogonTime = "Never Logged In"
		}
		else
		{
			#Update last logon variable with data from Office 365

			$strLastLogonTime = $objUserMailbox.LastLogonTime
		}
		

        $props = @{
            UserPrincipalName=$strUserPrincipalName
            LastLogonTime=$strLastLogonTime
            ItemCount=$objUserMailbox.ItemCount
            DisplayName=$objUserMailbox.DisplayName
        }
    
  
		$results += New-Object psobject -Property $props

	}
    Write-Host $OutputFile
    $results | select-object -property UserPrincipalName,DisplayName,LastLogonTime,ItemCount | Export-Csv $OutputFile -NoTypeInformation -Delimiter ";" -Encoding UTF8
	
	#Clean up session
	Get-PSSession | Remove-PSSession
}

###############################################################################
#
# Function ConnectTo-ExchangeOnline
#
# PURPOSE
#    Connects to Exchange Online Remote PowerShell using the tenant credentials
#
# INPUT
#    Tenant Admin username and password.
#
# RETURN
#    None.
#
###############################################################################
function ConnectTo-ExchangeOnline 
{    
    Param(  
        [Parameter( 
        Mandatory=$true, 
        Position=0)] 
        [String]$Office365AdminUsername, 
        [Parameter( 
        Mandatory=$true, 
        Position=1)] 
        [String]$Office365AdminPassword 
 
    ) 
         
    #Encrypt password for transmission to Office365 
    $SecureOffice365Password = ConvertTo-SecureString -AsPlainText $Office365AdminPassword -Force     
     
    #Build credentials object 
    $Office365Credentials  = New-Object System.Management.Automation.PSCredential $Office365AdminUsername, $SecureOffice365Password 
     
    #Create remote Powershell session 
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $Office365credentials -Authentication Basic –AllowRedirection         
 
    #Import the session 
    Import-PSSession $Session -AllowClobber | Out-Null 
} 
 
 
# Start script 
. Main