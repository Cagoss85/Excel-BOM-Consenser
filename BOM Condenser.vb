Sub SummariseBOMs()
' Based on hiker95, 08/24/2014, ME801072
' Modified by CJG 6/10/2021

' Define variables
Dim wb As Worksheet
Dim summarySheet As Worksheet
Dim row As Long
Dim lastRow As Long
Dim n As Long
Dim nextRow As Long

Application.ScreenUpdating = True

If Evaluate("ISREF(Summary!A1)") Then
    Worksheets("Summary").Cells.Clear
End If

If Not Evaluate("ISREF(Summary!A1)") Then Worksheets.Add().Name = "Summary"     'If there isnt a summary sheet add one
Set summarySheet = Sheets("Summary")        'Set the current sheet to the summary sheet


'__Create the Summary BOM Format__
With summarySheet
  .UsedRange.Clear
  With .Cells(1, 1).Resize(, 3)     'In the current sheet (1,1), make 3 columns
    .Value = Array("MFG PART NO", "DESCRIPTION", "QTY")     'Specify column names
    .Font.Bold = True       'Bold Font
  End With
End With

'__Stack all existing data in summary sheet__
For Each wb In ThisWorkbook.Worksheets      'Iterate over all worksheets
  If wb.Name <> "Summary" And wb.Name <> "BOM Template" Then        'If the name of a sheet isnt summary
    With wb
      lastRow = .Cells(Rows.Count, 1).End(xlUp).row     'Row corresponding to last data entry in Column B
      nextRow = summarySheet.Cells(summarySheet.Rows.Count, "A").End(xlUp).row + 1      'The row in summary where the data will be inserted
      .Range("B3:D" & lastRow).Copy     'Copy cells from range B2:Bottom Right Corner of Data
      summarySheet.Range("A" & nextRow).PasteSpecial xlPasteValues      'Paste values on summary starting at A2
      
      Application.CutCopyMode = False       'Clear the clipboard
    End With
  End If
Next wb 'Move onto the next sheet

'__Combine Data__
summarySheet.Activate
With summarySheet
  lastRow = .Cells(Rows.Count, 1).End(xlUp).row   'Last row of data
  .Range("A2:C" & lastRow).Sort key1:=.Range("A2"), order1:=1, key2:=.Range("B2"), order2:=2
  With .Range("D2:D" & lastRow)
    .FormulaR1C1 = "=RC[-3]&RC[-2]"
    .Value = .Value
  End With

  'Iterate over each row of data
  For row = 2 To lastRow
    n = Application.CountIf(.Columns(4), .Cells(row, 4).Value)
    If n > 1 Then
      .Range("C" & row).Value = Evaluate("=Sum(C" & row & ":C" & row + n - 1 & ")")
      .Range("A" & row + 1 & ":C" & row + 1 + n - 2).ClearContents
    End If
    row = row + n - 1
  Next row

  summarySheet.Columns(4).ClearContents
  
  'Search for empty rows, If one exists, delete all empty rows
  lastRow = .Cells(.Rows.Count, "A").End(xlUp).row   'Last row of data
  For row = 2 To lastRow
    If IsEmpty(.Range("A" & row).Value) Then        'If theres an empty cell in col A
        summarySheet.Range("A2:C" & lastRow).SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp      'Delete all empty rows and shift data up
    End If
  Next row
  
  summarySheet.Columns(1).Resize(, 3).AutoFit

  '__Format Borders__
  'Set to thin Black Borders
  lastRow = .Cells(Rows.Count, "A").End(xlUp).row
  .Range("A1:C" & lastRow).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
  .Range("A1:C" & lastRow).Borders.Color = RGB(0, 0, 0)
  .Range("A1:C" & lastRow).Borders.Weight = xlThin

  'Thick Borders to separate Rows
  For row = 2 To lastRow
    .Range("A" & row & ":C" & row).Borders(xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
    .Range("A" & row & ":C" & row).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
    .Range("A" & row & ":C" & row).Borders(xlEdgeBottom).Weight = xlMedium
  Next row

  'Thick Border to surround everything
  .Range("A1:C" & lastRow).BorderAround , ColorIndex:=0, Weight:=xlMedium
End With

Application.ScreenUpdating = True
End Sub

