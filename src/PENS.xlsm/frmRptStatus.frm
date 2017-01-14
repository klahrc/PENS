VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRptStatus 
   Caption         =   "PENS - Reports "
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5340
   OleObjectBlob   =   "frmRptStatus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRptStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    On Error Resume Next
    Windows(gsLastReport).Activate        'Make the new report the active window
    Unload Me
End Sub
