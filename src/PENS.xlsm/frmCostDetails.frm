VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCostDetails 
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   OleObjectBlob   =   "frmCostDetails.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCostDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_clsAnchors As CAnchors
Private Sub CmdFontSizeMinus_Click()
    On Error Resume Next

    If Me.lsvCostDetails.Font.Size > 10 Then
        Me.lsvCostDetails.Font.Size = Me.lsvCostDetails.Font.Size - 1
    End If

    If Me.lblCostDetails.Font.Size > 10 Then
        Me.lblCostDetails.Font.Size = Me.lblCostDetails.Font.Size - 1
    End If

    If (Me.lsvCostDetails.ColumnHeaders(1).Width - 10) >= 140 Then
        Me.lsvCostDetails.ColumnHeaders(1).Width = Me.lsvCostDetails.ColumnHeaders(1).Width - 10
    End If

    If (Me.lsvCostDetails.ColumnHeaders(2).Width - 5) >= 70 Then
        Me.lsvCostDetails.ColumnHeaders(2).Width = Me.lsvCostDetails.ColumnHeaders(2).Width - 5
    End If
End Sub

Private Sub CmdFontSizePlus_Click()
    On Error Resume Next


    Me.lsvCostDetails.Font.Size = Me.lsvCostDetails.Font.Size + 1

    If (Me.lblCostDetails.Font.Size < 24) Then
        Me.lblCostDetails.Font.Size = Me.lblCostDetails.Font.Size + 1
    End If

    Me.lsvCostDetails.ColumnHeaders(1).Width = Me.lsvCostDetails.ColumnHeaders(1).Width + 10
    Me.lsvCostDetails.ColumnHeaders(2).Width = Me.lsvCostDetails.ColumnHeaders(2).Width + 5


End Sub

Private Sub CmdFontSizePlus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CmdFontSizePlus_Click
End Sub

Private Sub CmdFontSizeMinus_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call CmdFontSizeMinus_Click
End Sub

Private Sub cmdOK_Click()
    gFrmCostDet.Hide

End Sub



Private Sub UserForm_Initialize()
    Dim lvwItem As ListItem

    Set m_clsAnchors = New CAnchors

    Set m_clsAnchors.Parent = Me

    ' restrict minimum size of userform
    m_clsAnchors.MinimumWidth = 45
    m_clsAnchors.MinimumHeight = 160

    With m_clsAnchors

        With .Anchor("frmFontSize")
            .AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleBottom
            .MinimumHeight = 36
        End With


        .Anchor("lsvCostDetails").AnchorStyle = enumAnchorStyleLeft Or enumAnchorStyleRight Or _
        enumAnchorStyleBottom Or enumAnchorStyletop
        .Anchor("cmdOK").AnchorStyle = enumAnchorStyleBottom Or enumAnchorStyleRight
        .Anchor("cmdOK").MinimumHeight = 24

        .Anchor("lblCostDetails").AnchorStyle = enumAnchorStyletop Or enumAnchorStyleRight Or enumAnchorStyleLeft

    End With

    ' live updates whilst resizing
    '''CheckBox1.Value = True
    '''ListBox1.RowSource = "B12:B19"

    m_clsAnchors.UpdateWhilstDragging = True


    With Me.lsvCostDetails
        .FullRowSelect = True
        .View = lvwReport
        .LabelEdit = lvwManual
        .Gridlines = True
        .Font.Name = "Calibri"
        .Font.Size = 10
        .HideColumnHeaders = True



        .ColumnHeaders.Add 1, , "", 140        '250
        .ColumnHeaders.Add 2, , "", 70        '140

        Set lvwItem = .ListItems.Add(, , "In Year Budget:")
        lvwItem.SubItems(1) = "$0"

        Set lvwItem = .ListItems.Add(, , "In Year Revised Baseline:")
        lvwItem.SubItems(1) = "$0"

        Set lvwItem = .ListItems.Add(, , "In Year Dashboard NE:")
        lvwItem.SubItems(1) = "$0"

        Set lvwItem = .ListItems.Add(, , "Reporting NE:")
        lvwItem.SubItems(1) = "$0"

        Set lvwItem = .ListItems.Add(, , "YTD Actuals:")
        lvwItem.SubItems(1) = "$0"

        Set lvwItem = .ListItems.Add(, , "Dashboard NE vs Revised BL:")
        lvwItem.SubItems(1) = "$0"


    End With



    Me.lblCostDetails.BackColor = RGB(55, 96, 145)


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


