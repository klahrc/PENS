Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'prevent save prompt
    Me.Saved = True
End Sub

Private Sub Workbook_Open()
    Application.OnTime Now, "Initialise"
End Sub