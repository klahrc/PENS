VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ofrmCostDetails 
   Caption         =   "Project Cost Details"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4875
   OleObjectBlob   =   "ofrmCostDetails.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ofrmCostDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    gFrmCostDet.Hide

End Sub



Private Sub UserForm_Initialize()

    Me.StartUpPosition = 0
    Me.Top = Application.Top + 100
    Me.Left = Application.Left + Application.Width - Me.Width - 25
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then
        Cancel = True
    End If

    Call cmdOK_Click
End Sub

