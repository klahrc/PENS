VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDetStatus 
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4335
   OleObjectBlob   =   "frmDetStatus.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDetStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsAnchors As CAnchors

Private Sub CmdFontSizeMinus_Click()
    On Error Resume Next

    If Me.txtDetStatus.Font.Size > 10 Then
        Me.txtDetStatus.Font.Size = Me.txtDetStatus.Font.Size - 1
    End If

    If Me.lblStatusReport.Font.Size > 10 Then
        Me.lblStatusReport.Font.Size = Me.lblStatusReport.Font.Size - 1
    End If
End Sub

Private Sub CmdFontSizePlus_Click()
    On Error Resume Next

    Me.txtDetStatus.Font.Size = Me.txtDetStatus.Font.Size + 1

    If (Me.lblStatusReport.Font.Size < 24) Then
        Me.lblStatusReport.Font.Size = Me.lblStatusReport.Font.Size + 1
    End If
End Sub

Private Sub CmdFontSizePlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CmdFontSizePlus_Click
End Sub

Private Sub CmdFontSizeMinus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CmdFontSizeMinus_Click
End Sub

Private Sub CmdFontSizePlus_()

    Private Sub cmdOK_Click()
        gFrmDetStatus.Hide
    End Sub


    Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

        '''Set m_clsAnchors = Nothing

        If CloseMode = 0 Then
            Cancel = True
        End If

        Call cmdOK_Click


    End Sub


    Private Sub UserForm_Initialize()

        Set m_clsAnchors = New CAnchors

        Set m_clsAnchors.Parent = Me

        ' restrict minimum size of userform
        m_clsAnchors.MinimumWidth = 45
        m_clsAnchors.MinimumHeight = 250

        With m_clsAnchors

            With .Anchor("frmFontSize")
                .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
                .MinimumHeight = 36
            End With


            .Anchor("txtDetStatus").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or _
            enumAnchorStyleBottom Or enumAnchorStyletop
            .Anchor("cmdOK").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
            .Anchor("cmdOK").MinimumHeight = 24

            .Anchor("lblStatusReport").AnchorStyle = enumAnchorStyletop Or enumAnchorStyleRight Or enumAnchorStyleLeft
            .Anchor("lblDiv1").AnchorStyle = enumAnchorStyleRight Or enumAnchorStyleLeft Or enumAnchorStyleBottom

        End With

        ' live updates whilst resizing
        '''CheckBox1.Value = True
        '''ListBox1.RowSource = "B12:B19"

        m_clsAnchors.UpdateWhilstDragging = True

        Me.lblStatusReport.BackColor = RGB(55, 96, 145)

        Me.StartUpPosition = 0
        Me.Top = Application.Top + 100
        Me.Left = Application.Left + Application.Width - Me.Width - 25

        Me.txtDetStatus.SetFocus

    End Sub



    Private Sub UserForm_Terminate()

        Set m_clsAnchors = Nothing

    End Sub




