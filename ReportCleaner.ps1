################################################################
##                                                            ##
##   Purpose: This is script was designed to translate        ##
##            and export the Data Protector txt report        ##
##            to a Excel file to be manipulated.              ##
##                                                            ##
##     Created by: Mark "The Automator" Borst                 ##
##     Version: 1.1                                           ##           
##     Last update date : 16/04/2020                          ##
##                                                            ##
##                                                            ##
##     Attention: This was created for Schneider Electric     ##
##                 Internal usage only                        ##
##                                                            ##
##     Changelog:                                             ##
##     1.1 - Added a task to sum MSSQL and FS separated       ##
################################################################


# Select the DP txt file, removes the 4 first rows and reacreate as $data
$input = Get-Content "D:\Tableau Data\ReportJur.txt" | select -Skip 4
$data = $input[1..($input.Length - 1)]

# List of Variables that will have value added
$GBWritten = 0
$Media = 0
$Files = 0
$Duration = 0
$Completed = 0
$InProgress = 0
$Failed = 0
$maxLength = 0
$FSGBs = 0
$SQLGBs = 0 


# Separate the data(txt file) with tab spaces 
$objects = ForEach($record in $data) {
    $split = $record -split "\s{2,}|\t+"
    If($split.Length -gt $maxLength){
        $maxLength = $split.Length
    }
    $props = @{}
    For($i=0; $i -lt $split.Length; $i++) {
        $props.Add([String]($i+1),$split[$i])
    }
    New-Object -TypeName PSObject -Property $props
}

# Creates a header for the table
$headers = [String[]](1..$maxLength)

# Export the table created as a CSV file
$objects | 
Select-Object $headers | 
Export-Csv -NoTypeInformation -Path "D:\Tableau Data\Temporary.csv"

# Need to work on a way to remove the need to create this temporary file

# Imports the temporary csv
$data2= Import-Csv "D:\Tableau Data\Temporary.csv"

# Looks on the Status Column and count the amount of Completed, In Progress and Failed sessions

# Completed
ForEach ($Line in $data2) 
{
    if ($Line.3 -like "*Completed*")
	{
	    $Completed = $Completed + 1
	}
}

# In Progress
ForEach ($Line in $data2) 
{
    if ($Line.3 -like "In Progress")
	{
	    $InProgress = $InProgress + 1
	}
}

$SuccessfulJobs = $InProgress + $Completed

# Failed
ForEach ($Line in $data2) 
{
    if ($Line.3 -like "*Failed*")
	{
	    $Failed = $Failed + 1
	}
}

# Total Sessions

$TotalSessions = $SuccessfulJobs + $Failed

# Sum the GBs Written in all File Servers Sessions
ForEach ($Line in $data2)
{
    if ($Line.2 -like "*MSSQL*")
	{
	    $FSGBs = $FSGBs + $Line.11
	}
}

# Sum the GBs Written in all SQL Server Sessions
ForEach ($Line in $data2)
{
    if ($Line.2 -notlike "*MSSQL*")
	{
	    $SQLGBs = $SQLGBs + $Line.11
	}
}

# Sum the GBs Written in all MS SQL Servers Sessions

# Sum the GBs Written in all sessions
ForEach ($Line in $data2) {$GBWritten = $GBWritten + $Line.11}

# Sum the amount of medias used
ForEach ($Line in $data2) {$Media = $Media + $Line.12}

# Sum the amount of files copied
ForEach ($Line in $data2) {$Files = $Files + $Line.20}

# Sum the duration of each job, excluding the ones that didn't started and divide later for 60 (to be showed as minutes)
ForEach ($Line in $data2) 
{
    if ($Line.8 -gt 0)
	{ 
        $Duration = $Duration + ($Line.8 - $Line.6)
	}
}

$InMinutes=[math]::Round($Duration/60)

# Result Output (Add the daily value to each column)
$object = New-Object PSObject
$object | Add-Member -Name Date -Value (Get-Date) -MemberType NoteProperty
$object | Add-Member -Name "Completed Sessions" -Value $Completed -MemberType NoteProperty
$object | Add-Member -Name "In Progress Sessions" -Value $InProgress -MemberType NoteProperty
$object | Add-Member -Name "Successful Jobs" -Value $SuccessfulJobs -MemberType NoteProperty
$object | Add-Member -Name "Failed Sessions" -Value $Failed -MemberType NoteProperty
$object | Add-Member -Name "Total Sessions" -Value $TotalSessions -MemberType NoteProperty
$object | Add-Member -Name "GBs Written" -Value $GBWritten -MemberType NoteProperty
$object | Add-Member -Name "GB FS Written" -Value $FSGBs -MemberType NoteProperty
$object | Add-Member -Name "GB SQL Written" -Value $SQLGBs -MemberType NoteProperty
$object | Add-Member -Name "Medias Written" -Value $Media -MemberType NoteProperty
$object | Add-Member -Name "Files Copied" -Value $Files -MemberType NoteProperty
$object | Add-Member -Name "Total Duration (In Minutes)" -Value $InMinutes -MemberType NoteProperty

# Export to an Excel file with the rest of the sessions
$object | Export-Excel "D:\Tableau Data\DPReportsSum-Jurubatuba.xlsx" -Append

Remove-Item "D:\Tableau Data\Temporary.csv"
