VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHelp 
   Caption         =   "YOU ARE NOTE ALONE!"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   OleObjectBlob   =   "frmHelp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub lblMail_Click()

    On Error GoTo ErrorHandler

    Link = "mailto:president@whitehouse.gov"
    On Error GoTo NoCanDo
    ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
    Unload Me
    Exit Sub
NoCanDo:
    MsgBox "Cannot open " & Link
End Sub

Private Sub lblWeb_Click()

    On Error GoTo ErrorHandler

    Link = "http://www.whitehouse.gov"
    On Error GoTo NoCanDo
    ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
    Unload Me
    Exit Sub
NoCanDo:
    MsgBox "Cannot open " & Link
End Sub

Private Sub OKButton_Click()

    On Error GoTo ErrorHandler

    Unload Me
End Sub


