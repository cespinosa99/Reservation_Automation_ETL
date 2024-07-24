
# ------------------------------------------------------Functions------------------------------------------------------------------

#wipe yesterday table (YesterdayRawObjects)
function wipeTable($Sheet){
    $RawData = $Sheet.ListObjects(1)
    Write-Host $RawData.Name
    $RawData.DataBodyRange.ClearContents()
}


#---------------------It's really weird with referncing sheets. It works. Don't touch it.
#Move today's (yesterday's) data to yesterday table
function moveTodayToYesterday{

    $SheetFrom = $workbook.sheets.item("TodayObjects")
    $SheetTo = $workbook.sheets.item("YesterdayObjects")
    $SheetFromData = $SheetFrom.UsedRange.Offset(1,0)
    #to
    $SheetToData = $SheetTo.ListObjects(1) #pointing at the second object b/c the first table is invisible in the backend
    #copy
    $SheetFromData.Copy($SheetToData)


    # #add "new" label
    $startCell = $SheetTo.Cells.Item(2,10)
    $lastRow = $SheetTo.Cells($SheetTo.Rows.Count, 1).End(-4162).Row  # -4162 is xlUp
    $endCell = $SheetTo.Cells.Item($lastRow, 10)
    # AutoFill the formula down the column
    $startCell.Formula = "old"
    $range = $SheetTo.Range($startCell, $endCell)
    $startCell.AutoFill($range, 0)  # 0 is xlFillDefault
}


#Move today's raw data to the todayTable
function moveObjectsToToday($rawDataWorksheet){
 
    $SheetFrom = $rawDataWorksheet
    $SheetTo = $workbook.sheets.item("TodayObjects")
    $SheetFromData = $SheetFrom.UsedRange.Offset(1,0)
    #to
    $SheetToData = $SheetTo.ListObjects(1)
    #copy
    $SheetFromData.Copy($SheetToData)
    
    #concat the items
    $startCell = $SheetTo.Cells.Item(2,9)
    $lastRow = $SheetTo.Cells($SheetTo.Rows.Count, 1).End(-4162).Row  # -4162 is xlUp
    $endCell = $SheetTo.Cells.Item($lastRow, 9)
    # AutoFill the formula down the column
    $startCell.Formula = "=CONCAT(A2:G2)"
    $range = $SheetTo.Range($startCell, $endCell)
    $startCell.AutoFill($range, 0)  # 0 is xlFillDefault


    # #add "new" label
    $startCell = $SheetTo.Cells.Item(2,10)
    $lastRow = $SheetTo.Cells($SheetTo.Rows.Count, 1).End(-4162).Row  # -4162 is xlUp
    $endCell = $SheetTo.Cells.Item($lastRow, 10)
    # AutoFill the formula down the column
    $startCell.Formula = "new"
    $range = $SheetTo.Range($startCell, $endCell)
    $startCell.AutoFill($range, 0)  # 0 is xlFillDefault

}


