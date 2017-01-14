VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReleaseNotes 
   Caption         =   "PENS - Release Notes"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   OleObjectBlob   =   "frmReleaseNotes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReleaseNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OKButton_Click()
    gFrmReleaseNotes.Hide
End Sub

Private Sub UserForm_Initialize()

    Me.StartUpPosition = 0
    Me.Top = Application.Top + 100
    Me.Left = Application.Left + Application.Width - Me.Width - 25

    Me.txtReleaseNotes.SetFocus

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If

    Call OKButton_Click


End Sub

