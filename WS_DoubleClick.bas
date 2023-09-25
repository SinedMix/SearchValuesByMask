Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    'The code described below checks which cell the user double-clicks
    'and if the row and column match the specified ones, the custom form is displayed.
    'Otherwise, the program will enter the normal input mode.
    
    'by Sined - aboutdatum.ru
    Cancel = True
    If Target.Column = 2 And Target.Row = 4 Then
        searchForm.Show 0
    Else:
        Application.SendKeys "{F2}"
    End If
End Sub