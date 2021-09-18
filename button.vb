Sub Button1_Click()
    '    UserForm1.StartUpPosition = 0
    '    UserForm1.Top = Application.Top + 122
    '    UserForm1.Left = Application.Left + Application.Width - UserForm1.Width - 12
    '    UserForm1.Show
    With UserForm1
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub
