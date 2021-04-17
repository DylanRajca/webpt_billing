# Synopsis - Pull patient data from webPT reports in 'Upload' directory and parse into a billing report, then upload billing report to 'Billing Reports' directory.

# Include functions.ps1
. "betsy\webpt\webpt_billing_functions.ps1"

# File Paths
$analyticsPath = "$env:USERPROFILE\Documents\WebPT Analytics"
$uploadPath = "$analyticsPath\Upload"
$billedPath = "$analyticsPath\Billing Reports"

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
##### Main #####
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# If files exist in 'Upload' directory
if ($files = getFiles) {
    main $files
}


























return
# Create billing_report workbook in Billing Report directory.
$XL = New-Object -comobject Excel.Application
$XL.visible = $false
$billing = $XL.Workbooks.Add()
$billing_sheet = $billing.worksheets.item(1)

# Iterate through webpt reports.
for ($i = 0; $i -lt $files.Count; $i++) {
    $book = $XL.Workbooks.Open($files[$i].FullName)
    $Sheet = $book.worksheets.item(1)
    $identifier = $Sheet.range("I1").text

    $billing_sheet.cells.item(1, ($i + 1)) = $identifier
    # Save & close workbook
    $book.Save()
    $book.close($true)
}


# Save billing_report.xlsx and quit excel
$billing.SaveAs("$env:USERPROFILE\Desktop\webPT-reports\Billing Report\billing_report.xlsx")
$XL.Quit()




return


$n = 1
$array = @()
while (($Sheet.Rows.item(1).cells.item($n).Value2).length -gt 0) {
    $array += , ($Sheet.Rows.item(1).cells.item($n).Value2)
    $n = $n + 1
}

$n = 6
while (($Sheet.Rows.item($n).cells.item(9).Value2).length -gt 5) {
    $blArray += , ($Sheet.Rows.item($n).cells.item(9).Value2)
    $n = $n + 1
}


return $blArray


return $array

for ($i = 1; $i -le 14; $i++) {
    $ex = $Sheet.Rows.item($i).cells.item(9).Value2
    Write-Host $ex
}