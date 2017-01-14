VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTips 
   Caption         =   "PENS - Tips"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7305
   OleObjectBlob   =   "frmTips.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
    gFrmTips.Hide
End Sub

Private Sub cmdNext_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call cmdNext_click
End Sub

Private Sub cmdPrevious_dblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call cmdPrevious_Click
End Sub

Private Sub cmdNext_click()
    If glPosTip < UBound(gsTipsArray) Then
        glPosTip = glPosTip + 1
    End If
    txtTip.Value = gsTipsArray(glPosTip)
End Sub

Private Sub cmdPrevious_Click()
    If glPosTip > 0 Then
        glPosTip = glPosTip - 1
    End If
    txtTip.Value = gsTipsArray(glPosTip)
End Sub

Private Sub txtTip_Change()

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



