#you can generate a log file.
# Set up log file name with current date
$logFileName = "Logfilename_$(Get-Date -Format 'dd-MM-yyyy hhmmss').log"
$logFilePath = "C:\your path\Logs\$logFileName"

# Start logging - Script start time
$startTime = Get-Date
"Script started at: $startTime" | Out-File -FilePath $logFilePath -Append

#Execute a sql query with example

$server = "yourservername"
#or your IP Address
$server= "10.00.00.234"
$instance = "Instance Name"
$port = "Port Number"


# Invoke SQL query with specified credentials
#$username
#$password


#Conditions are the record in termination register must be added after 1st Jan 2023
#The cessation column must be 1 which means we need to terminate the user and it was not revoked.
#The same EMP from termination will be checked in PhoneDirectory and see if the user has multiple entries for the same employee, the script will check only the most recent emp number.


#you can write your own query

$sqlQuery = @"
SELECT 
    p.Colum1, 
    t.Coolum1,
    p.column2, 
    p.column3, 
    p.column4,
    t.column2, 
    t.column3,
    t.Column4 AS ColumnName
FROM 
    [Databasename].[dbo].[Tablename] AS p
INNER JOIN 
    [Database2].[dbo].[Tablename] AS t 
ON 
    t.Column2 COLLATE SQL_Latin1_General_CP1_CI_AS = p.Column1 COLLATE SQL_Latin1_General_CP1_CI_AS
WHERE p.column2 IS NOT NULL
    AND p.column2 <> ''
    AND p.column4 = (
        SELECT 
            MAX(b.column4) 
        FROM 
            [Database].[dbo].[Table] AS b 
        WHERE 
            b.Column2 = p.Column2
    )
    AND t.Column5 = 1 AND t.column6 >= '2024-01-01 00:00:00'
	AND TRY_CONVERT(date, t.datecolumn, 105) = CONVERT(date, DATEADD(day, +2, GETDATE()))
	
	
"@  

#always execute in  Try Catch block
try 
{
	#using windows authentication
     $sqlData = Invoke-Sqlcmd -ServerInstance "$server,$port\$instance" -Query $sqlQuery -ErrorAction Stop -TrustServerCertificate
	 
	 #or you can use the below if username and password are provided
	 #$sqlData = Invoke-Sqlcmd -ServerInstance "$server,$port\$instance" -Query $sqlQuery -Username $username -Password $password -ErrorAction Stop -TrustServerCertificate
	 
	 
     try
	 {
	 $rowCount = $sqlData.Count
	 
	 #if returned 1 row, there will be no count
	 if($rowCount -eq $null)
	 {
		  $rowCount=1;
	 }
	 }
	 catch
	 {
		  $rowCount=1;
	 }
	
	#Add a SQL Success line in the log file.
	

#you can append the log file if you want.
	"Success: SQL query executed successfully $(Get-Date -Format 'dd-MM-yyyy hh:mm:ss')" | Out-File -FilePath $logFilePath -Append
	"Success: Found $rowCount number of rows" | Out-File -FilePath $logFilePath -Append
 
 }
 
catch 
{
      # Handle the error as needed and add a line to the log file.
	$errorMessage = "Failure: Failed to Execute SQL query Error: $_"
    $errorMessage | Out-File -FilePath $logFilePath -Append
}
