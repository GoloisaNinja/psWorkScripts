# Outlook COM library
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

# Outlook application deets
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
# Excel application
$excelObj = New-Object -ComObject Excel.Application
# Excel global vars
$ws = $null
$wb = $null
$rowMax = $null
$range = $null

# Folder to scrape
$folder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Date range
# TESTING $startDate = (Get-Date).AddDays(-11)
$dtFormat = "yyyy-MM-dd"
Write-Host "Program will run from start date to today (/) |•,,,,•| (/)"
$dateStr = Read-Host "Enter a start date ($dtFormat): "
$startDate = [DateTime]::ParseExact($dateStr, $dtFormat, $null)
$endDate = Get-Date
Write-Host "Please wait while the warp core is brought online..."
# Subject
$subject = "LFSS Termination Report" 

# Filter them mails bruh
$filteredEmails = $folder.Items | Where-Object {
    $_.Subject -match $subject -and 
    $_.ReceivedTime -ge $startDate -and 
    $_.ReceivedTime -le $endDate
}

# Get a count
$attachmentCount = 0
# Terminated users that need checked count
$terminatedCount = 0
# File name
$cdAsStr = $endDate.ToString("yyyy-MM-ddHHmmss")
# Excel files Output file path
$fp = "C:\Users\Jon Collins\Documents\SAGE_TERMINATION_ARCHIVE\"
# Termination Hash Table
$termHash = @{}
# Latest Users listing file path
$ufp = "C:\Users\Jon Collins\Documents\SAGE_USERS_LOC\SAGEUSERS_LATEST.xlsx"
# Termination text file name prefix
$tfname = "TERMINATION_CHECK" + $cdAsStr + ".txt"
# Termination text file output path
$tfp = "C:\Users\Jon Collins\Documents\SAGE_TERMINATION_OUTPUT\" + $tfname
# Regex Pattern for termination xlsx file date
$pattern = "(?<=LIN HR - LFSS Termination Report )+.{10}"
# Process the filtered emails and save the attachments
foreach ($attachment in $filteredEmails.attachments) {
    if ($attachment.FileName -Match "Termination") {
        $attachment.SaveAsFile($fp + $attachment.FileName)
        $attachmentCount++
    }
}

Write-Host "Thank you for your continued patience human - you will be rewarded with cake...promise..."

# Substring termination names by first name letter and last name as lower case
# This should be switched to employee ID if we ever figure a way to get that into Sage
# Example, if name is John Smith - arraylist value will be jsmith
Get-ChildItem $fp -Filter *.xlsx |
Foreach-Object {
    $report = Select-String -InputObject $_.FullName -Pattern $pattern -AllMatches
    $reportDate = $report.Matches.Value
    $wb = $excelObj.Workbooks.Open($_.FullName)
    $ws = $wb.Sheets.Item(1)
    $rowMax = ($ws.UsedRange.Rows).Count
    for ($x = 5; $x -le $rowMax; $x++) {
        # Column B is Name - Column D is termination date
        $nameSubStr = $ws.Cells.Item($x,"B").Value2.split(",")
        $termDate = [DateTime]::FromOADate($ws.Cells.Item($x,"D").Value2)
        $termDate = $termDate.ToString("yyyy-MM-dd")
        $comboDate = "$termDate|$reportDate"
        $fname = $nameSubStr[1].ToLower().Trim()
        $lname = $nameSubStr[0].ToLower().Trim()
        $termHash[$fname.Substring(0,1) + $lname] = $comboDate
    }
    # Close the terminated workbook after range has be evaluated
    $wb.Close($false)
}

# Set new workbook object to be latest user list
$wb = $excelObj.Workbooks.Open($ufp)
$ws = $wb.Sheets.Item(1)
$rowMax = ($ws.UsedRange.Rows).Count
$termOutputStr = "name|termination date|report date" + "`n"
for ($row = 1; $row -le $rowMax; $row++) {
    # User value always in column A
    $user = $ws.Cells.Item($row, "A").Value2
    if ($user -ne $null) {
        $cu = $user.ToLower().Trim() # Clean User
        if ($termHash.ContainsKey($cu)) {
            $terminatedCount++
            $termOutputStr += $user + "|" + $termHash[$cu] + "`n"
        }
    }
}
# Close the latest users workbook
$wb.Close($false)
if ($terminatedCount -gt 0) {
    $termOutputStr | Out-File -FilePath $tfp
    Write-Host "**ALERT! ($terminatedCount) TERMINATED USERS EXIST - CHECK TERMINATION OUTPUT**"
} else {
    Write-Host "**No terminated users found in Sage Latest Users**"
}
Write-Host "Termination attachments processed: " $attachmentCount

Write-Host "Casting horcruxes..."
Write-Host "Petting puppies..."
Write-Host "Sniffing glue..."
# Quit Excel object an marshall before close
$excelObj.quit()
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelObj)

# Clean up the termination archive files
Get-ChildItem -Path $fp -File | Foreach { $_.Delete() }
Write-Host "Termination Archive has been nuked \_(--)_/"

Write-Host "(⌐■_■)"
Write-Host "(>⌐■-■"
Write-Host "( >⌐■-■"
Write-Host "( —>⌐■-■"
Write-Host "( —_>⌐■-■"
Write-Host "( —_—>⌐■-■"
Write-Host "( —_—)>⌐■-■"
Write-Host "( ❂_❂)>⌐■-■"

Write-Host "TERMINATING - ALL YOUR BASE BELONG TO US"