# Outlook COM library
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

# Outlook application deets
$Outlook = New-Object -ComObject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

# Folder to scrape
$folder = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Date range
Write-Host "Program will run from start date to today"
$dtFormat = "yyyy-MM-dd"
$dateStr = Read-Host "Enter a start date ($dtFormat): "
$startDate = [DateTime]::ParseExact($dateStr, $dtFormat, $null)
$endDate = Get-Date

# Subject
$subject = "SAGE DELETED ORDER" 

# Filter them mails bruh
$filteredEmails = $folder.Items | Where-Object {
    $_.Subject -eq $subject -and 
    $_.ReceivedTime -ge $startDate -and 
    $_.ReceivedTime -le $endDate
}

# Get a count
$recordCount = 0

# Concat desired text into string
$outputStr = ""
# File name
$currDate = Get-Date
$cdAsStr = $currDate.ToString("yyyy-MM-ddHHmmss")
$fn = "SAGE_OD_" + $cdAsStr + ".txt"
# Output file path
$outputFP = "C:\Users\Jon Collins\Documents\SAGE_DO_ARCHIVE\" + $fn
# Process the filtered emails
# Postive lookbehind regex for just the 7 digit order number to be deleted
$pattern = "(?<=SAGE Sales Order Number: )+\d{7}"
Write-Host "Engaging the Event Horizon gravity drive...please wait..."
foreach ($email in $filteredEmails) {
    $string = $email.Body
    $result = Select-String -InputObject $string -Pattern $pattern -AllMatches
    $extracted = $result.Matches.Value
    $outputStr += $extracted + "`n"
    $recordCount++
}
if ($recordCount -gt 0) {
    $outputStr | Out-File -Filepath $outputFP
    Write-Host "**ALERT! $recordCount emails processed**"
    Write-Host "Review the Sage Delete Order Archive!"
} else {
    Write-Host "Well, well, well - how the turntables..."
    Write-Host "Delete orders are as empty as my robot soul..."
    Write-Host "(/) |•,,,,•| (/)"
}


Write-Host "(⌐■_■)"
Write-Host "(>⌐■-■"
Write-Host "( >⌐■-■"
Write-Host "( —>⌐■-■"
Write-Host "( —_>⌐■-■"
Write-Host "( —_—>⌐■-■"
Write-Host "( —_—)>⌐■-■"
Write-Host "( ❂_❂)>⌐■-■"

Write-Host "TERMINATING - ALL YOUR BASE BELONG TO US"