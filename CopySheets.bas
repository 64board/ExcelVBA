Attribute VB_Name = "CopySheets"
Option Explicit

' Copies sheets ranges (values and formats) from a closed workbook into
' a sheet on the opened workbook.
Private Sub CopySheets(ByVal fileName As String, ByVal sheetName As String, ByVal dstSheetName As String, ByVal rng As String)

    Sheets(dstSheetName).Cells.Clear
    
    Dim srcWB As Workbook
    
    ' Open the source workbook and copy the values
    Set srcWB = Workbooks.Open(fileName)

    srcWB.Sheets(sheetName).Range(rng).Copy

    ThisWorkbook.Activate
    
    ' Paste values and formats
    With Sheets(dstSheetName)
        .Range(rng).PasteSpecial Paste:=xlPasteFormats
        .Range(rng).PasteSpecial Paste:=xlPasteColumnWidths
        .Range(rng).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    End With
    
    ' Get out of the copy mode
    Application.CutCopyMode = False
    
    ' Close the source workbook without saving
    srcWB.Close savechanges:=False

End Sub

Public Sub diff()
    Dim fileName As String
    
    ' Copy 2 sheets from 2 different files
    
    ' The first file
    ' Get the file names from a cell
    fileName = Sheets("Main").Range("B1").Value
    CopySheets fileName, "Summary", "Summary", "A1:M26"
    CopySheets fileName, "Day Positions", "DayPositions", "A1:N32"
    
    ' The second file
    fileName = Sheets("Main").Range("B2").Value
    CopySheets fileName, "Summary", "SummaryNew", "A1:M26"
    CopySheets fileName, "Day Positions", "DayPositionsNew", "A1:N32"
    
    ThisWorkbook.Sheets("Diff").Activate
    ThisWorkbook.Save
    
End Sub

