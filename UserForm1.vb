Option Explicit On

Dim descDict As Dictionary
Dim wghtDict As Dictionary
Dim distDict As Dictionary
Dim distPnDict As Dictionary

Private Sub Label1_Click()

    'THESE ARE VERY DANGEROUS
    'Debug.Print "clearing"
    'Application.SendKeys "^g ^a {DEL}"

End Sub

Private Sub TextBox1_Change()

    Dim key As Variant, i As Long
    ReDim Items(descDict.Count - 1) As String
    
    Dim matches() As String, j As Long
    Dim val As String
    val = UserForm1.TextBox1.Value

    Dim mlen As Long

    For Each key In descDict.Keys()
        Items(i) = descDict(key)
        If InStr(Items(i), val) > 0 Then
            j = j + 1
            ReDim Preserve matches(1 To j)
            matches(j) = Items(i)
        End If
        i = i + 1
    Next

    mlen = UBound(matches)

    If mlen > 0 Then
        UserForm1.Label2.Caption = matches(1)
    Else
        UserForm1.Label2.Caption = "No Results"
    End If
    If mlen > 1 Then
        UserForm1.Label3.Caption = matches(2)
    Else
        UserForm1.Label3.Caption = ""
    End If
    If mlen > 2 Then
        UserForm1.Label4.Caption = matches(3)
    Else
        UserForm1.Label4.Caption = ""
    End If
    If mlen > 3 Then
        UserForm1.Label5.Caption = matches(4)
    Else
        UserForm1.Label5.Caption = ""
    End If
    If mlen > 4 Then
        UserForm1.Label6.Caption = matches(5)
    Else
        UserForm1.Label6.Caption = ""
    End If
    If mlen > 5 Then
        UserForm1.Label7.Caption = matches(6)
    Else
        UserForm1.Label7.Caption = ""
    End If
    If mlen > 6 Then
        UserForm1.Label8.Caption = matches(7)
    Else
        UserForm1.Label8.Caption = ""
    End If
    If mlen > 7 Then
        UserForm1.Label9.Caption = matches(8)
    Else
        UserForm1.Label9.Caption = ""
    End If
    If mlen > 8 Then
        UserForm1.Label10.Caption = matches(9)
    Else
        UserForm1.Label10.Caption = ""
    End If
    If mlen > 9 Then
        UserForm1.Label11.Caption = matches(10)
    Else
        UserForm1.Label11.Caption = ""
    End If
    If mlen > 10 Then
        UserForm1.Label6.Caption = matches(11)
    Else
        UserForm1.Label2.Caption = ""
    End If

End Sub

Private Sub UserForm_Initialize()

    '    Me.StartUpPosition = 0
    '    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    '    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    '    Me.Show
    Call initDicts()

End Sub

Function initDicts()

    Dim wkb As Excel.Workbook
    Dim wks As Excel.Worksheet

    Dim row As Long
    Dim lastRow As Long
    Dim tempArr As Variant
    ReDim tempArr(1 To 4)

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

End Function
