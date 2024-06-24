#send an email and also update the log file
# Set up log file name with current date
$logFileName = "Logifilename $(Get-Date -Format 'dd-MM-yyyy hhmmss').log"
$logFilePath = "C:\Yourpath\Logs\$logFileName"

# Start logging - Script start time
$startTime = Get-Date
"Script started at: $startTime" | Out-File -FilePath $logFilePath -Append


#exceute some code 
"Write a log in the file " | Out-File -FilePath $logFilePath -Append



#Lets send an email as a notification
# Email configuration
$smtpUsername = "firstname.lastname@domain.com"  # Replace with your email address
$EmailTo = "firstname.lastname2@domain.com"  # Recipient's email address     
$EmailSubject = "Subject Line"
$EmailCc="firstname.lastname3@domain.com" ,"firstname.lastname4@domain.com"  #cc email addresses
	  
	  
	  
# Define the email body
$emailBody = @"
Hello Team,

The following user has been terminated:
UserID: $userID
EMPNAME: $EmpName


"@

# Check if user is in .LAS-M365-License-Project-P3 group and modify email body accordingly
if ($ProjectIsInGroup) {
   $emailBody += @"
Please raise a request to cancel Microsoft Project Subscription as the user is a part of 'Project'.

"@
}
# Check if user is in LAS-M365-License-Visio-P2 group and modify email body accordingly
if ($VisioIsInGroup) {
   $emailBody += @"
Please raise a request to cancel Visio Subscription as the user is a part of 'Visio'.

"@
}

$emailBody += @"

Regards,
Yours faithfully
Signature
"@


        # SMTP server configuration for Office 365
        $smtpServer = "smtp.domain.i" 
        $smtpPort = 657

        # Send email with attachment
         try 
		 {
               Send-MailMessage -From $smtpUsername -To $EmailTo -Subject $EmailSubject -Body $emailBody -SmtpServer $smtpServer -Port $smtpPort -UseSsl `
                     -Attachments $outputExcelFilePath
               
				
			#Add a log for Excel file creation  
	         "Success: Email sent successfully" | Out-File -FilePath $logFilePath -Append
				
         } 
		 catch 
		 {
              #Write-Host "Failed to send email: $_"
			  
			  
				  #Add a failure message in the log file.
		        $errorMessage = "Failure: Failed to send email: $_"
                #write-output $errorMessage
               $errorMessage | Out-File -FilePath $logFilePath -Append
	     }
    
    

