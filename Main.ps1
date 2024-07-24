#//-------------------------------------------------------TODO--------------------------------------------------------------------------
#Hide the raw sheets once everything is updated
#once everything is in the right table, see if there is a way to autorefresh the powerquery
    #maybe use powerAutomate or some function
    #rebuild the datamodel since it's broken (see if the script is what breaks it, figure out why)  
#//--------------------------------------------------------------------------------------------------------------------------------------

Connect-AzAccount

.\ "$PSScriptRoot\alterDataFunctions.ps1"
.\ "$PSScriptRoot\MasterPull.ps1"

#Pull Objects
$Objects = @(loadSQLObjects), @(loadAppServiceObjects), @(loadvmssObjects)
createObjectCSV($Objects)

#Pull Reservations
#createReservationsCSV

#//------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
$excel = New-Object -ComObject Excel.Application
# Specify the path to your existing workbook
$filePath = "$PSScriptRoot\Reservations_Overview_P2.xlsx"
# Open the workbook
$workbook = $excel.Workbooks.Open($filePath)
#$csvFiles = "C:\Users\espincha\Documents\reservations\ReservationsReport.csv", "C:\Users\espincha\Documents\reservations\reservation_automation_ETL\Objects.csv"
$csvFiles ="$PSScriptRoot\Objects.csv"

foreach ($csv in $csvFiles) {
    # Get the filename without extension
    $sheetName = [System.IO.Path]::GetFileNameWithoutExtension($csv)
    
    # Add a new sheet
    $newWorksheet = $workbook.Worksheets.Add()
    $todayDate = Get-Date -Format "MM_dd_yyyy"
    $newWorksheet.Name = "$sheetName $todayDate"
    
    # Import new collected data from CSV to the notebook
    $TxtConnector = "TEXT;" + $csv
    $Connector = $newWorksheet.QueryTables.Add($TxtConnector, $newWorksheet.Range("A1"))
    $query = $newWorksheet.QueryTables.Item($Connector.Name)
    $query.TextFileOtherDelimiter = ","
    $query.TextFileParseType = 1
    $query.TextFileColumnDataTypes = ,1 * $newWorksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1
    
    # Execute & delete the import query
    $query.Refresh()
    $query.Delete()
}

Write-Host "New Object sheet and Reservation sheet have been moved to the main sheet"

$TodaySheet = $workbook.sheets.item("TodayObjects")
$YesterdaySheet = $workbook.sheets.item("YesterdayObjects")
Write-Host "Remove yesterday sheet"
wipeTable($YesterdaySheet) #wipe yesterday table
Write-Host "Move today to yesterday"
moveTodayToYesterday #Move today data table to yesterday data table
Write-Host "Remove today sheet"
wipeTable($TodaySheet) #Clear Today table to clear up space for new data
Write-Host "Move rawObjects to today's sheet"
moveObjectsToToday($newWorksheet)


#figure out how to refresh the queries in script
#append the updates and the new and removed columns to the static table


# Save & close the Workbook as XLSX
$workbook.Save()
$excel.Quit()

Write-Host "Reservation_Overview_Workbook has been updated"
