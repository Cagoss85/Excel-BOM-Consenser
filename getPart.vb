Sub InsertPart()

    Dim wb As Worksheet
    Set wb = ActiveWorkbook.ActiveSheet

    Dim PNString As String
    Dim PNLong As Long
    Dim lastRow As Long
    Dim outputArr As Variant
    
    lastRow = Cells(Rows.Count, 2).End(xlUp).row
    PNString = Application.InputBox("Input Part Number:", "Insert Part Tool", _
                                    "100000", , , , , 2)
    If Len(PNString) <> 6 Or IsNumeric(PNString) <> True Then
        MsgBox ("Alakai Int PN has to be Numeric and of Length 6")
    Else
        PNLong = CLng(PNString)
        outputArr = lookUpPN(PNLong)
        With wb
            Cells(lastRow + 1, 1).Value = lastRow - 1
            Cells(lastRow + 1, 2).Value = PNLong
            Cells(lastRow + 1, 3).Value = outputArr(1)
            Cells(lastRow + 1, 6).Value = outputArr(2)
            Cells(lastRow + 1, 7).Value = outputArr(3)
            Cells(lastRow + 1, 8).Value = outputArr(4)
        End With
    End If
End Sub

'For this function to work, dictionaries must be enabled'
Function lookUpPN(PN As Long) As Variant

Dim wkb As Excel.Workbook
Dim wks As Excel.Worksheet

Dim row As Long
Dim lastRow As Long
Dim tempArr As Variant
ReDim tempArr(1 To 4)
Dim descDict As Dictionary
Dim wghtDict As Dictionary
Dim distDict As Dictionary
Dim distPnDict As Dictionary


Set wkb = Excel.Workbooks("Part_Number_Lookup.xlsm")
Set wks = wkb.Worksheets("Lookup_Table")
Set descDict = New Dictionary
Set wghtDict = New Dictionary
Set distDict = New Dictionary
Set distPnDict = New Dictionary

lastRow = wks.Cells(Rows.Count, 2).End(xlUp).row

For row = 3 To lastRow
    descDict(wks.Cells(row, 2).Value) = wks.Cells(row, 3).Value
    wghtDict(wks.Cells(row, 2).Value) = wks.Cells(row, 6).Value
    distDict(wks.Cells(row, 2).Value) = wks.Cells(row, 7).Value
    distPnDict(wks.Cells(row, 2).Value) = wks.Cells(row, 8).Value
Next row

tempArr(1) = descDict(PN)
tempArr(2) = wghtDict(PN)
tempArr(3) = distDict(PN)
tempArr(4) = distPnDict(PN)
    
lookUpPN = tempArr

End Function