Attribute VB_Name = "NarrowHelper"
Option Explicit

Public Sub run()
    
    Application.ScreenUpdating = False
    
    Call read_narrow_csv
    Call update_dailytemp_tab
        
    ActiveWorkbook.Save
    
    Call save_new_narrow
    
    Call save_as_pdf
    
    Application.ScreenUpdating = True
    
End Sub

Private Sub update_dailytemp_tab()
Attribute update_dailytemp_tab.VB_ProcData.VB_Invoke_Func = " \n14"
    ' Reads a CSV file into a sheet, then creates a column extracting the symbol portion
    ' from first column of CSV, that column will be used as an index.
    ' Apply a filter using a list of symbols.
    ' Copy filtered data to another sheet.
    
    ' Open .dat file.
    Dim file_name As String
    file_name = ThisWorkbook.Path & "\" & "dailytemp.dat"
    
    ' CSV file has 8 fields, first one is symbol, second is date,
    ' then open, high, low and close price fields, then 2 more fields that are not use.
    Workbooks.OpenText fileName:=file_name, _
        Origin:=xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=True, Space:=False, Other:=False, _
        FieldInfo:=Array(Array(1, xlTextFormat), Array(2, xlMDYFormat), Array(3, xlGeneralFormat), Array(4, xlGeneralFormat), _
        Array(5, xlGeneralFormat), Array(6, xlGeneralFormat), Array(7, xlGeneralFormat), Array(8, xlGeneralFormat)), _
        TrailingMinusNumbers:=False
    
    ' Clean filters and cells in destination sheet.
    Windows("narrow.xlsm").Activate
    Worksheets("dailytemp").Activate
    'ActiveSheet.AutoFilterMode = False
    'Cells.Clear
    
    ' Clean data sheet.
    'Worksheets("data").Activate
    'Cells.Clear
    
    ' Copying data.
    Windows("dailytemp.dat").Activate
    
    ' Copy.
    Columns("A:H").Select
    Selection.Copy
    
    ' Paste.
    Windows("narrow.xlsm").Activate
    Worksheets("dailytemp").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    ' Close .dat file.
    Windows("dailytemp.dat").Activate
    ActiveWindow.Close
    
    ' Formula for filtering based on symbols portion only, no contract part.
    ' Contract portion is at the end and is 3 characters long.
    Dim last_row As Long
    last_row = Range("A:A").SpecialCells(xlCellTypeLastCell).Row
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-8],LEN(RC[-8])-3)"
    Selection.AutoFill Destination:=Range("I1:I" & last_row)
    
    ' Filtering.
    'Cells.Select
    'Selection.AutoFilter
    'ActiveSheet.Range("$A$1:$I$" & last_row).AutoFilter Field:=9, Criteia1:=Array( _
    '    "RCL", "RNG", "RHO", "RNRB", "RSP", "RNQ", "RRT", "USD", "REC", "RBP", "RAD", _
    '    "RGC", "RSI", "DEX", "S", "C", "SB", "KC"), Operator:=xlFilterValues
   
    'Range("A1").Select
    
    ' Copy filtered values only.
    'Range("A1:I" & last_row).SpecialCells(xlCellTypeVisible).Copy
    'Worksheets("data").Activate
    'Cells(1, 1).PasteSpecial
    'Application.CutCopyMode = False
    'Range("A1").Select
    
    Worksheets("narrow").Activate
    
    ' Update date
    Range("W2").Value = Worksheets("dailytemp").Range("B1")
    
End Sub

Private Sub read_narrow_csv()
    
    Worksheets("narrow").Activate
    Range("A3").Activate
    
    Dim file_name As String
    file_name = ActiveWorkbook.Path & "\narrow.csv"
    
    read_csv file_name:=file_name, delimiter:=","
    
    ActiveWorkbook.Save
   
End Sub

Private Sub read_csv(file_name As String, delimiter As String)

    ' Reads a CSV file into a sheet.
    
    Dim line As String
    Dim fields() As String
    Dim f As Variant
    Dim file_number As Integer
    Dim count As Integer
    
    file_number = FreeFile
    Open file_name For Input As file_number
   
    Do Until EOF(1)
        Line Input #file_number, line
        
        ' Columns values.
        fields = Split(line, delimiter)

        ' For jumping back to original column.
        count = 0
        
        For Each f In fields
            ActiveCell = f
            ActiveCell.Offset(0, 1).Select
            count = count + 1
        Next
        
        ' Next row
        ActiveCell.Offset(1, -count).Select
    Loop
   
    Close file_number
   
End Sub

Private Sub save_new_narrow()
    Dim file_name As String
    file_name = ActiveWorkbook.Path & "\narrow_new.csv"
    write_csv file_name:=file_name, rng:=Range("A3:A20,Q3:R20"), delimiter:=","
End Sub

Private Sub write_csv(file_name As String, rng As Range, delimiter As String)

    ' Writes a CSV file from a sheet range. The range can be a multiple selection range.
    ' Assumes same amount of rows in case of multiple selection range.
    
    Dim r As Range
    Dim c As Range
    Dim line As String
    Dim index As Integer
    
    ' Auxiliary array to keep the lines.
    Dim lines() As String
    ReDim lines(rng.Rows.count - 1)
    Dim element As Variant
    For Each element In lines
        element = ""
    Next
    
    ' For the case of more than one area selection.
    Dim area As Range
    For Each area In rng.Areas

        index = 0
        For Each r In area.Rows
        
            line = ""
            
            For Each c In r.Cells
                line = line & c.Value & ","
            Next
        
            lines(index) = lines(index) & line
            
            index = index + 1

        Next
    Next
    
    ' Write CSV file.
    Dim file_number As Integer
    file_number = FreeFile
    Open file_name For Output As #file_number
    
    Dim i As Integer
    For i = LBound(lines) To UBound(lines)
        
        ' Remove last comma from string line.
        line = lines(i)
        line = Left(line, Len(line) - 1)
        
        Print #file_number, line
    Next
    
    Close file_number
   
End Sub

Private Sub save_as_pdf()

    Dim fileName As String
    fileName = ActiveWorkbook.Path & "\narrow_new.pdf"
    
    With ActiveSheet
        .ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=fileName, _
        OpenAfterPublish:=False
    End With

End Sub

