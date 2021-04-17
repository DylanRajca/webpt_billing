function getFiles {
    try {
        return  Get-ChildItem "$uploadPath\*.csv" , "$uploadPath\*.xls" -ErrorAction Stop
    }
    catch {
        Write-Host "Please make sure 'Upload' directory exists in $analyticsPath." 
    }
}

function generateHeader ($Sheet, $lastColAddress) {
    for ($i = 1; $i -le $lastColAddress; $i++) {
        $header += , ($Sheet.Rows.item(1).cells.item($i).text)
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

        # Save index of last column/row used
        [int]$lastRowAddress = ($openBookSheet.UsedRange.Rows.count + 1) - 1
        [int]$lastColAddress = ($openBookSheet.UsedRange.columns.count + 1) - 1

        # Capture report header
        $header = @(generateHeader $openBookSheet $lastColAddress)

        # Parse report 
        parseReport $header $openBook $openBookSheet
        
        # Save and close workbooks
        $openBook.Save()
        $openBook.close($true)
    }
     
    # Quit Excel
    $XL.Quit()
}
