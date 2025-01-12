<#
    PS script for sending emails to CDC customer service 
    Initial options should include:
    Duplicate Item
    Bad Date
    No New Order Type
#>
$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)

#---FUNCTIONS---#
function Set-MyTemplate {
    param (
        [Parameter(Mandatory)]
        [string]$tName
    )
#---MAIL TEMPLATES---#
$dupe = @"
<html>
<style>
    .mnone{margin: 0px;}
    .blue{color: #1b41fa;}
</style>
<body>
    <p>Hello CDC,</p>
    <p class="mnone">An order change for SO # <strong>$so</strong> failed because of a duplicate item.</p>
    <p class="mnone">Please refer to the attached images if needed.</p>
    <p class="mnone">The duplicate item is listed as:</p>
    <p class="mnone">Item Number: <strong>$itemNum</strong></p>
    <p class="mnone">Item Description: <strong>$itemDesc</strong></p>
    <p class="mnone">Quantity: <strong>$qtyAsString</strong></p>
    <p class="mnone">Order ship date: <strong>$shipDate</strong></p>
    <h3>Steps to resolve</h3>
    <p class="mnone">If you see the duplicated item in your order, simply remove it and submit another order change 
    with any other relevant changes. This will become the newest change and should correct the error. If 
    you do not see a duplicate item in your order, then unfortunately it is possible the order file has
    become corrupted in Sage. In this case it is best to delete SO # <strong>$so2</strong> and create
    a new order with the relevant info and changes. 
    <br />
    <br />
    <h3>For internal use</h3>
    <p class="mnone">SAP Order Number: <strong><span class="blue">$sapOrder</span></strong></p>
    <p class="mnone">IDOC reference: <strong><span class="blue">$idoc</span></strong></p>
</body>
</html>
"@
    If ($tname -eq "dupe") {
        $template = $dupe
    }
    $template
}

## USER START
Write-Host 
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
$recips = $null
$cc = $null
$subject = $null
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


$templateChoice = 0
$quantity = 0
$attachmentCount = -1
$soStr = ""

while ($templateChoice -lt 1 -or $templateChoice -gt 3) {
    [int]$templateChoice = Read-Host "Enter 1 (Duplicate Item) 2 (Bad Date) 3 (No New Order)"
}

if ($templateChoice -eq 1) {
    Write-Host "Duplicate Item Template Selected! Please complete the prompts."
    $company = Read-Host "Enter the redi company (ex. R100)"
    While ($soStr.length -ne 7) {
        Write-Host "Sales order must be seven digits - try again"
        $soStr = Read-Host "Enter the 7 Digit Sales Order #"
    }
    $so = [int]$soStr
    $so2 = $so
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
    $dateStr = Read-Host "Enter the ship date ($dtformat)"
    $dtTime = [DateTime]::ParseExact($dateStr, $dtformat, $null)
    $shipDate = $dtTime.ToString("MM-dd-yyyy").Substring(0,10)
    $sapOrder= Read-Host "Enter the SAP Order Number"
    $idoc = Read-Host "Enter the IDOC number: "
    While ($attachmentCount -lt 0 -or $attachmentCount -gt 5) {
        [int]$attachmentCount = Read-Host "How many attachments to include? (enter 0 if none, max of 5)"
    }
    if ($attachmentCount -gt 0) {
        $attachments = Get-ChildItem -Path $attachmentPath | Sort-Object LastWriteTime -Descending | Select-Object -First $attachmentCount
    }
    $recips = "CDCcustomerservice@onelineage.com"
    $cc = "aromano@onelineage.com;jparker@onelineage.com;bhamlin@onelineage.com"
    $subject = "$company - $so - ORDER CHANGE FAILURE - DUPLICATE ITEM"
    $template = Set-MyTemplate -tName "dupe"
}

$Mail.To = "($recips)"
$Mail.Cc = ("$cc")
$Mail.Subject = ("$subject")
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