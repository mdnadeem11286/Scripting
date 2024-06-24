# Set up log file name with current date
$logFileName = "Logifilename $(Get-Date -Format 'dd-MM-yyyy hhmmss').log"
$logFilePath = "C:\Yourpath\Logs\$logFileName"

# Start logging - Script start time
$startTime = Get-Date
"Script started at: $startTime" | Out-File -FilePath $logFilePath -Append


#exceute some code 
"Write a log in the file " | Out-File -FilePath $logFilePath -Append