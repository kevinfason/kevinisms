#From https://kevinisms.fason.org
<#    
.SYNOPSIS    
    
  Get all computer objects from Active Directory with "Windows" in the OS name, sort by version and email HTML-formatted results
        
.COMPATABILITY     
     
  Tested on PS v4. 
      
.EXAMPLE  
  PS C:\> Get-OSCount.ps1  
  All options are set as variables in the GLOBALS section so you simply run the script.  
  
.NOTES    
        
  NAME:       Get-OSCount.ps1    
    
  AUTHOR:     Brian D. Arnold    
    
  CREATED:    8/28/14  
    
  LASTEDIT:   9/3/14   
#>

# Import AD module
Import-Module ActiveDirectory

###################
##### GLOBALS #####
###################

# Get domain
$DomainName = (Get-ADDomain).NetBIOSName 

# How many days ago was the lastLogon attribute updated?
$days = 120
$lastLogonDate = (Get-Date).AddDays(-$days).ToFileTime()

# SMTP settings
$smtpServer = "smtpsmarthost.yourdomain.com"
$smtpFrom = "OSCount@yourdomain.com"
$smtpTo = "recipient@yourdomain.com"
$messageSubject = "$DomainName Windows OS Counts - lastLogon within $days days"

# HTML settings 
$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border-width: 1px solic black;border-style: solid;border-color: black;border-collapse: collapse;}"
$style = $style + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black; background: #dddddd;}"
$style = $style + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;}"
$style = $style + "</style>"

################
##### MAIN #####
################

# Query AD
$Computers = @(Get-ADComputer -Properties Name,operatingSystemversion,operatingsystem,lastLogontimeStamp -Filter {(OperatingSystem -like "*Windows*") -AND (lastLogontimeStamp -ge $lastLogonDate)})
#foreach($Computer in $Computers)
{
    $Computer.OperatingSystem = $Computer.OperatingSystem -replace '®' -replace '™' -replace '专业版','Professional (Ch)' -replace 'Professionnel','Professional (Fr)'
}
$Computers | Group-Object operatingSystem,operatingsystemversion | Select Count,Name | Sort Name 

# Send output as email
$message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto
$message.Subject = $messageSubject
$message.IsBodyHTML = $true
$message.Body = $Computers | Group-Object operatingSystem,operatingsystemversion | Select Count,Name | Sort Name | ConvertTo-Html -Head $style 

$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)  
