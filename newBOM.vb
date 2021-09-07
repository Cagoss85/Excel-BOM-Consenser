Sub NewBOM()

Dim wb As Worksheet

'Updateby Extendoffice
    Dim newSheetName As String
    Dim checkSheetName As String
    
    newSheetName = Application.InputBox("Input Sheet Name:", "New BOM Tool", _
                                    "sheet4", , , , , 2)
    On Error Resume Next
    checkSheetName = Worksheets(newSheetName).Name
    If checkSheetName = "" Then
        Worksheets.Add.Name = newSheetName
        'MsgBox "The sheet named ''" & newSheetName & _ "'' does not exist in this workbook but it has been created now.", _
        vbInformation, "New BOM Tool"
        Set wb = Sheets(newSheetName)        'Set the current sheet to the new sheet
        With wb
            .UsedRange.Clear
            With .Cells(2, 1).Resize(, 8)     'In the current sheet (1,1), make 8 columns
                .Value = Array("ID NO", "INT PART NO", "DESCRIPTION", "QTY", "UOM", "WEIGHT", "DISTRIBUTOR", "DIST PART NO")     'Specify column names
                .Font.Bold = True       'Bold Font
            End With
            wb.Columns(1).Resize(, 8).AutoFit
            
            Dim rng As Range
            Set rng = Range("A1:H1")
            Dim formula As String
            formula = "=CONCATENATE(""BILL OF MATERIAL: "" ,Mid(CELL( ""Filename"", A1), Find( ""]"", CELL( ""Filename"", A1)) + 1, 255))"
            With rng
                .Merge
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True       'Bold Font
            End With
            Range("A1").formula = formula
        End With
    Else
        MsgBox "The sheet named ''" & newSheetName & _
        "''exists in this workbook.", vbInformation, "New BOM Tool"
    End If

End Sub
