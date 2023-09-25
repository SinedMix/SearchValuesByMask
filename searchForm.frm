VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} searchForm 
   Caption         =   "Search"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   OleObjectBlob   =   "searchForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "searchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'by Sined - aboutdatum.ru

Option Explicit
Option Compare Text
Dim x

Private Sub CheckBox1_Click()
    TextBox1_Change
End Sub


Private Sub UserForm_Initialize()
    Dim arr As Variant, lLastRow, i As Long, tbl As Object
    arr = Array()
    'lLastRow = ThisWorkbook.Sheets("List").Cells(Rows.Count, 1).End(xlUp).Row
    Set tbl = Sheets("List").ListObjects("ListOfCities")
    For i = 3 To tbl.ListRows.Count
        ReDim Preserve arr(i - 3)
        arr(i - 3) = tbl.DataBodyRange(i - 2)
    Next
    x = arr
End Sub

Private Sub TextBox1_Change()
Dim i As Long, s As String, txt As String, lt As Long

txt = TextBox1.Text: lt = Len(txt)
If lt = 0 Then Exit Sub

For i = 1 To UBound(x)
    If CheckBox1.Value = True Then
        If txt = Mid(x(i), 1, lt) Then s = s & "~" & x(i) ' first letters search
    Else:
        If InStr(1, x(i), txt) > 0 Then s = s & "~" & x(i) ' full match search
    End If
Next i

Me.ListBox1.List = Split(Mid(s, 2), "~")

End Sub

Private Sub ListBox1_Click()
If ListBox1.ListIndex = -1 Then Exit Sub
    ActiveCell.Value = ListBox1
End Sub

Private Sub CommandButton1_Click(): Unload Me: End Sub


