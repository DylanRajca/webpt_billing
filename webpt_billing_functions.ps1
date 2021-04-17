function getFiles {
    try {
        return  Get-ChildItem "$uploadPath\*.csv" , "$uploadPath\*.xls" -ErrorAction Stop
    }
    catch {
        Write-Host "Please make sure 'Upload' directory exists in $analyticsPath." 
    }
}

function generateHeader ($Sheet) {
    $n = 1
    while (($Sheet.Rows.item(1).cells.item($n).Value2).length -gt 0) {
        $header += , ($Sheet.Rows.item(1).cells.item($n).text)
        $n = $n + 1
    }
    return $header
}

function parseReport ($header, $book, $Sheet) {

    # Identify which report is passed.
    if ($header -match "Start Time") {
        Write-Host "Scheduled Visits"
    }
    elseif ($header -match "Documenting Therapist Type") {
        Write-Host "Patient Notes"
    }
    elseif ($header -match "CPT code") {
        Write-Host "Billed Units"
    }
    elseif ($header -match "Authorization Number") {
        Write-Host "Authorization"
    }
}

function main ($files) {
    
    # Open Excel
    $XL = New-Object -comobject Excel.Application
    $XL.visible = $false

    # Iterate through WebPt reports.
    for ($i = 0; $i -lt $files.Count; $i++) {
        $openFile = $files[$i]

        # Open workbook
        $openBook = $XL.Workbooks.Open($openFile.FullName)
        $openBookSheet = $openBook.worksheets.item(1)

        # Capture report header
        $header = generateHeader($openBookSheet) 

        # Parse report 
        parseReport $header $openBook $openBookSheet
        
        # Save and close workbooks
        $openBook.Save()
        $openBook.close($true)
    }
     
    # Quit Excel
    $XL.Quit()
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