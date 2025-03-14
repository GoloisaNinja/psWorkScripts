<#
    PS script for sending emails to CDC customer service 
    Initial options should include:
    Duplicate Item
    Bad Date
    No New Order Type
#>
$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)

#---PASSWORD STUFF---#
[Reflection.Assembly]::LoadWithPartialName("System.Web")
[System.Web.Security.Membership]::GeneratePassword(15,2)

#---FUNCTIONS---#
function Set-MyTemplate {
    param (
        [Parameter(Mandatory)]
        [string]$tName
    )
    $reason = ""
    Switch($tName) {
        "dupe" {$reason = "duplicate item."}
        "date" {$reason = "bad date/time."}
    }
#---MAIL TEMPLATES---#
#---HEADER/BASE---#
$base = @"
<html>
<style>
    .mnone{margin: 0px;}
    .blue{color: #1b41fa;}
</style>
<body>
    <p>Hello CDC,</p>
    <p class="mnone">An order change for SO # <strong>$so</strong> failed because of a $reason</p>
    <p class="mnone">Please refer to the attached images if needed.</p>
    <p class="mnone">Order ship date: <strong>$shipDate</strong></p>
"@
#----------------#
#-----FOOTER-----#
$footer = @"
<br />
<br />
<h3>For internal use</h3>
<p class="mnone">SAP Order Number: <strong><span class="blue">$sapOrder</span></strong></p>
<p class="mnone">IDOC reference: <strong><span class="blue">$idoc</span></strong></p>
</body>
</html>
"@
#----------------#
#---DUPLICATE----#
$dupe = @"
$base
    <p class="mnone">The duplicate item is listed as:</p>
    <p class="mnone">Item Number: <strong>$itemNum</strong></p>
    <p class="mnone">Item Description: <strong>$itemDesc</strong></p>
    <p class="mnone">Quantity: <strong>$qtyAsString</strong></p>
    <h3>Steps to resolve</h3>
    <p class="mnone">If you see the duplicated item in your order, simply remove it and submit another order change 
    with any other relevant changes. This will become the newest change and should correct the error. If 
    you do not see a duplicate item in your order, then unfortunately it is possible the order file has
    become corrupted in Sage. In this case it is best to delete SO # <strong>$so2</strong> and create
    a new order with the relevant info and changes.</p>
$footer
"@
#---------------#
#---BAD DATE----#
$date = @"
$base
    <p class="mnone">An order must have a valid date and a time in military format with only digits (no characters like am/pm).</p>
    <p class="mnone">An order with a time value of all zeros is allowed, but the date must contain a valid value.</p>
    <h3>Steps to resolve</h3>
    <p class="mnone">This change has failed. Submit another order change with all the relevant information and a good date and time value.
    SO # <strong>$so2</strong> will be updated accordingly.</p>
$footer
"@
#---------------#
#---SAGE CREDENTIAL---#
$sage = @"
<html>
<style>
    .mnone{margin: 0px;}
    .blue{color: #1b41fa;}
</style>
<body>
    <p>Hello $user,</p>
    <p class="mnone">Your Sage Credentials have been <strong>$sageAction</strong>.</p>
    <p class="mnone">Please login to Sage using the following:</p>
    <p class="mnone">Login: <strong><span class="blue">$login</span></strong></p>
    <p class="mnone">Password: <strong><span class="blue">$password</span></strong></p>
    <br />
    <h3>Please Note:</h3>
    <p class="mnone">This is a temporary password. Once logged in to Sage, please choose the menu option File > Change User Password</p>
    <p class="mnone">The included email attachment shows you what this menu looks like, if needed.</p>
    <p class="mnone">Your password must contain upper case and lower case letters, at least one digit, and at least one special character (ex. !).</p>
</body>
</html>
"@
#----------------#
 Switch($tname) {
    "dupe" {$template = $dupe}
    "date" {$template = $date}
    "sage" {$template = $sage}
 }
 $template
}

function Get-MyAttachments {
    if ($attachmentCount -gt 0) {
        $attachments = Get-ChildItem -Path $attachmentPath | Sort-Object LastWriteTime -Descending | Select-Object -First $attachmentCount
    }
    $attachments
}

function Get-MySubjectLine {
    param (
        [Parameter(Mandatory)]
        [string]$tType 
    )
    $returnSub = ""
    $endStr = ""
    Switch($tType) {
        "dupe" {$endStr = "DUPLICATE ITEM"}
        "date" {$endStr = "BAD DATE/TIME"}
    }
    $returnSub = "$company - $so - ORDER CHANGE FAILURE - $endStr"
    $returnSub
}

function Get-MyEmailDetails {
    param (
        [Parameter(Mandatory)]
        [string]$tType
    )
    $dupeDetails = @{
        "to" = "CDCcustomerservice@onelineage.com"
        "cc" = "aramero@onelineage.com;jparker@onelineage.com;bhamlin@onelineage.com"
        "sub" = Get-MySubjectLine -tType "dupe"
    }
    $dateDetails = @{
        "to" = "CDCcustomerservice@onelineage.com"
        "cc" = "aramero@onelineage.com;jparker@onelineage.com;bhamlin@onelineage.com"
        "sub" = Get-MySubjectLine -tType "date"
    }
    $sageDetails = @{
        "to" = $userEmail
        "cc" = $null
        "sub" = "Sage Credentials"
    }
    Switch($tType) {
        "dupe" {return $dupeDetails;Break}
        "date" {return $dateDetails;Break}
        "sage" {return $sageDetails;Break}
    }
}

## USER START
Write-Host 
Write-Host @"
Welcome to the TemplatR!
Select your template type and follow the prompts.
If including attachments - ensure they are placed 
into C:\template_attachments before continuing.
"@

#---GLOBALS---#
$dtformat = "yyyy-MM-dd"
$template = $null
$company = $null
$soStr = $null
$itemNum = $null
$itemDesc = $null
$quantity = $null
$qtyAsString = $null
$shipDate = $null
$sapOrder = $null
$idoc = $null
$attachmentCount = 0
$attachmentPath = "C:\Users\holid\template_attachments"
$attachments = $null
$detailsObj = $null
$to = $null
$cc = $null
$sub = $null

$user = $null
$userEmail = $null
$sageActionInt = 0
$sageAction = $null
$login = $null
$password = $null


$templateChoice = 0
$templateChoiceStr = ""
$quantity = 0
$attachmentCount = -1
$soStr = ""

while ($templateChoice -lt 1 -or $templateChoice -gt 4) {
    [int]$templateChoice = Read-Host "Enter 1 (Duplicate Item) 2 (Bad Date) 3 (No New Order) 4 (Sage Credential)"
}

#---TEMPLATE AGNOSTIC PROMPTS---#
Switch ($templateChoice) {
    1 {$templateChoiceStr = "Duplicate Item TemlateR Selected! Please complete the prompts."; Break}
    2 {$templateChoiceStr = "Bad Date TemplatR Selected! Please complete the prompts."; Break}
    3 {$templateChoiceStr = "No New Order TemplatR Selected! Please complete the prompts."; Break}
    4 {$templateChoiceStr = "Sage Credential TemplatR Selected! Please complete the prompts."; Break}
}
Write-Host $templateChoiceStr
#---SAGE ORDER/CDC TEMPLATE AGNOSTIC PROMPTS---#
If ($templateChoice -lt 4) {
 $company = Read-Host "Enter the redi company (ex. R100)"
 While ($soStr.length -ne 7) {
     $soStr = Read-Host "Enter the Sales Order # (must be 7 digits)"
 }
 $so = [int]$soStr
 $so2 = $so
 $dateStr = Read-Host "Enter the ship date ($dtformat)"
 $dtTime = [DateTime]::ParseExact($dateStr, $dtformat, $null)
 $shipDate = $dtTime.ToString("MM-dd-yyyy").Substring(0,10)
 $sapOrder= Read-Host "Enter the SAP Order Number"
 $idoc = Read-Host "Enter the IDOC number"
 While ($attachmentCount -lt 0 -or $attachmentCount -gt 5) {
     [int]$attachmentCount = Read-Host "How many attachments to include? (enter 0 if none, max of 5)"
 }
}

if ($templateChoice -eq 1) {
    $itemNum = Read-Host "Enter the duplicate item number"
    $itemDesc = Read-Host "Enter the item description"
    While ($quantity -le 0 -or $quantity -gt 2) {
        [int]$quantity = Read-Host "Enter 1 (qty was same) 2 (qty differed)"
    }
    If ($quantity -eq 1) {
        $qtyAsString = "Same Quantity"
    } else {
        $qtyAsString = "Different Quantity"
    }
    $detailsObj = Get-MyEmailDetails -tType "dupe"
    $attachments = Get-MyAttachments
    $template = Set-MyTemplate -tName "dupe"
}

if ($templateChoice -eq 2) {
    $detailsObj = Get-MyEmailDetails -tType "date"
    $attachments = Get-MyAttachments
    $template = Set-MyTemplate -tName "date"
}

if ($templateChoice -eq 4) {
    $user = Read-Host "Enter the users first name"
    $userEmail = Read-Host "Enter the users email address"
    while ($sageActionInt -lt 1 -or $sageActionInt -gt 2) {
        $sageActionInt = Read-Host "Credential Action - Enter 1 (created) 2 (reset)"
    }
    Switch($sageActionInt) {
        1 {$sageAction = "created";Break}
        2 {$sageAction = "reset";Break}
    }
    $login = Read-Host "Enter a login for the user (this is typically the users Emp ID)"
    #---PASSWORD SHOULD HAVE AT LEAST LOWER,UPPER,DIGIT,AND ! OR @---#
    do {
        $password = [System.Web.Security.Membership]::GeneratePassword(15,2)
    } until ($password -match '(?=.*[a-z])(?=.*[A-Z])(?=.*\d)(?=.*[\!\@]).+')
    #---OVERRIDE ATTACHMENT LOCATION AND COUNT FOR STATIC INCLUSION OF SAGE IMAGE---#
    $attachmentPath = "C:\Users\holid\credential_attachment"
    $attachmentCount = 1
    $detailsObj = Get-MyEmailDetails -tType "sage"
    $attachments = Get-MyAttachments
    $template = Set-MyTemplate -tName "sage"
}

$to = $detailsObj["to"]
$cc = $detailsObj["cc"]
$sub = $detailsObj["sub"]

$Mail.To = ("$to")
$Mail.Cc = ("$cc")
$Mail.Subject = ("$sub")
if ($attachmentCount -gt 0) {
    foreach ($attach in $attachments) {
        [void]$Mail.Attachments.Add($attach.FullName)
    }
}
$Mail.HTMLBody = ("$template")
$Mail.Display()
Write-Host
Write-Host
Write-Host @"
Powering down quantum arrays...
Cooling memory cores...
Exiting No Exits...
Thank you for using TemplatR meatbag...have a nice day
"@