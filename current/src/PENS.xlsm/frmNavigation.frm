VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNavigation 
   Caption         =   "The Local Guide"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3600
   OleObjectBlob   =   "frmNavigation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_ProjList As Collection



'---------------------------------------------------------------------------------------
' Method : UpdDetCostForm
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub UpdDetCostForm()
    Dim DashNE As Double
    Dim dRevBL As Double

    Dim lvwItem As ListItem


    With gFrmCostDet
        If Me.lstProjects.ListCount > 0 Then

            .lblCostDetails = "COST DETAILS - " + m_ProjList.Item(Me.lstProjects.ListIndex + 1).ProjName

            .lsvCostDetails.ListItems(1).SubItems(1) = Format((m_ProjList.Item(Me.lstProjects.ListIndex + 1).IYBudget), "$#,##0")
            .lsvCostDetails.ListItems(2).SubItems(1) = Format((m_ProjList.Item(Me.lstProjects.ListIndex + 1).IYRevisedBL), "$#,##0")
            .lsvCostDetails.ListItems(3).SubItems(1) = Format((m_ProjList.Item(Me.lstProjects.ListIndex + 1).IYNE), "$#,##0")
            .lsvCostDetails.ListItems(4).SubItems(1) = Format((m_ProjList.Item(Me.lstProjects.ListIndex + 1).RepNE), "$#,##0")
            .lsvCostDetails.ListItems(5).SubItems(1) = Format((m_ProjList.Item(Me.lstProjects.ListIndex + 1).YTDActuals), "$#,##0")

            DashNE = Val(m_ProjList.Item(Me.lstProjects.ListIndex + 1).IYNE)
            dRevBL = Val(m_ProjList.Item(Me.lstProjects.ListIndex + 1).IYRevisedBL)

            If dRevBL <> 0 Then
                '''.lblNEvsRevBL.Caption = Format((DashNE / dRevBL) - 1, "0.0%")
                .lsvCostDetails.ListItems(6).SubItems(1) = Format((DashNE / dRevBL) - 1, "0.0%")
            Else
                ''' .lblNEvsRevBL.Caption = ""
                .lsvCostDetails.ListItems(6).SubItems(1) = ""
            End If


            '.lblIYBudget.Caption = Format((m_ProjList.item(Me.lstProjects.ListIndex + 1).IYBudget), "$#,##0")
            '.lblIYRB.Caption = Format((m_ProjList.item(Me.lstProjects.ListIndex + 1).IYRevisedBL), "$#,##0")
            '.lblYTDActuals.Caption = Format((m_ProjList.item(Me.lstProjects.ListIndex + 1).YTDActuals), "$#,##0")
            '.lblIYNE.Caption = Format((m_ProjList.item(Me.lstProjects.ListIndex + 1).IYNE), "$#,##0")
        Else

            .lblCostDetails = "COST DETAILS"
            ' .lblIYBudget.Caption = ""
            ' .lblIYRB.Caption = ""
            ' .lblYTDActuals.Caption = ""
            ' .lblIYNE.Caption = ""
            ' .lblNEvsRevBL.Caption = ""
        End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Method : chkIncludesBAU_Click
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub chkIncludesBAU_Click()
    If gbCompletedNavPanelLoad Then
        Call updateNavigationPanel        ' Make it a function to test!!!
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : chkIncludesDelivered_Click
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub chkIncludesDelivered_Click()
    If gbCompletedNavPanelLoad Then
        Call updateNavigationPanel        ' Make it a function to test!!!
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : cmdCost_Click
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub cmdCost_Click()

    Call UpdDetCostForm

    gFrmCostDet.Show False

End Sub

'---------------------------------------------------------------------------------------
' Method : updDetStatus
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub updDetStatus()

    With gFrmDetStatus
        If Me.lstProjects.ListCount > 0 Then


            .lblStatusReport = "STATUS REPORT - " + m_ProjList.Item(Me.lstProjects.ListIndex + 1).ProjName
            .txtDetStatus.Text = m_ProjList.Item(Me.lstProjects.ListIndex + 1).DetStatus
            .txtDetStatus.SetFocus
            .txtDetStatus.CurLine = 0

        Else
            .txtDetStatus.Text = ""
            .lblStatusReport = "STATUS REPORT"
        End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Method : cmdStatus_Click
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub cmdStatus_Click()

    gFrmDetStatus.Show False

End Sub

'---------------------------------------------------------------------------------------
' Method : cmdLeftArrow_Click
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub cmdLeftArrow_Click()
    If Me.lstProjects.ListIndex > 0 Then
        Me.lstProjects.Selected(Me.lstProjects.ListIndex - 1) = True
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : cmdLeftArrow_dblClick
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub cmdLeftArrow_dblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call cmdLeftArrow_Click
End Sub

'---------------------------------------------------------------------------------------
' Method : cmdRightArrow_Click
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub cmdRightArrow_Click()
    If Me.lstProjects.ListIndex < Me.lstProjects.ListCount - 1 Then
        Me.lstProjects.Selected(Me.lstProjects.ListIndex + 1) = True
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : cmdRightArrow_dblClick
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub cmdRightArrow_dblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call cmdRightArrow_Click
End Sub

'---------------------------------------------------------------------------------------
' Method : cmbDeliveryLeader_Change
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub cmbDeliveryLeader_Change()

    If Me.cmbDeliveryLeader.ListIndex > 0 Then
        Me.cmbDeliveryLeader.BackColor = RGB(0, 255, 255)        ' Highlight if search is active
    Else
        Me.cmbDeliveryLeader.BackColor = RGB(242, 242, 242)        ' Back to Grey
    End If

    If gbCompletedNavPanelLoad Then
        Call updateNavigationPanel        ' Make it a function to test!!!
    End If

End Sub

'---------------------------------------------------------------------------------------
' Method : cmbActivationStatus_Change
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub cmbActivationStatus_Change()

    If Me.cmbActivationStatus.ListIndex > 0 Then
        Me.cmbActivationStatus.BackColor = RGB(0, 255, 255)        ' Highlight if search is active
    Else
        Me.cmbActivationStatus.BackColor = RGB(242, 242, 242)        ' Back to Grey
    End If


    If gbCompletedNavPanelLoad Then
        Call updateNavigationPanel        ' Make it a function to test!!!
    End If

End Sub

'---------------------------------------------------------------------------------------
' Method : cmbCategory_Change
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub cmbCategory_Change()

    If Me.cmbCategory.ListIndex > 0 Then
        Me.cmbCategory.BackColor = RGB(0, 255, 255)        ' Highlight if search is active
    Else
        Me.cmbCategory.BackColor = RGB(242, 242, 242)        ' Back to Grey
    End If

    If gbCompletedNavPanelLoad Then
        Call updateNavigationPanel        ' Make it a function to test!!!
    End If

End Sub

'Private Sub Image1_Click()
'    Dim wbNew As Workbook
'    Dim lPos As Long'

'    If (MsgBox("In order to guarantee a smooth navigation experience, PENS requires to remove all filters and frozen panes from the Portfolio Plan. Are you OK to proceed? " + _
     '               vbNewLine, vbQuestion + vbYesNo, "PENS - Navigation") = vbYes) Then

'        Set wbNew = openDashboardFile()        ' Make it a function to test!!!

'        If Not wbNew Is Nothing Then
'            wbNew.Sheets("Portfolio Plan").Select
'            ActiveWindow.FreezePanes = False
'            wbNew.Sheets("Portfolio Plan").AutoFilterMode = False

'            wbNew.Sheets("Portfolio Plan").Range("C4").Select
'            ActiveWindow.FreezePanes = True

'            wbNew.Sheets("Portfolio Plan").Rows(1).RowHeight = 21
'            wbNew.Sheets("Portfolio Plan").Rows(2).RowHeight = 0

' Assuming there is at least one project (goto the last and come back to current point in the Projects listbox)
'            lPos = Me.lstProjects.ListIndex
'            Me.lstProjects.ListIndex = gFrmNavPanel.lstProjects.ListCount - 1
'            Me.lstProjects.ListIndex = lPos
'        Else
'            ' ERRORROR!!!!!
'        End If
'    Else
'        Unload Me
'    End If

'End Sub



Private Sub txtSearch_Change()

    If Me.txtSearch.Value = "" Then
        Me.txtSearch.BackColor = RGB(242, 242, 242)        ' Back to Grey
        Me.lblCount.ForeColor = vbWhite
    Else
        Me.txtSearch.BackColor = RGB(0, 255, 255)        ' Highlight if search is active
        Me.lblCount.ForeColor = vbCyan        ' Cyan font if search is active
    End If

    Call updateNavigationPanel        ' Make it a function to test!!!

End Sub
'---------------------------------------------------------------------------------------
' Method : updateNavigationPanel
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub updateNavigationPanel()
    Dim j As Long
    Dim lRevBL As Long
    Dim lNE As Long


    j = loadProjects(m_ProjList, Me.lstProjects, _
                     Me.cmbDeliveryLeader.List(Me.cmbDeliveryLeader.ListIndex), _
                     Me.cmbActivationStatus.List(Me.cmbActivationStatus.ListIndex), _
                     Me.cmbCategory.List(Me.cmbCategory.ListIndex), _
                     Me.txtSearch.Value)        ''''Make it a function to test!!!!

    If j > 0 Then
        ''''''' DALE Me.Caption = "The Local Guide - 1/" + Str(j)

        lblCount.Caption = "1/" + CStr(j)

        Me.cmdStatus.Caption = " STATUS"
        Me.cmdCost.Caption = "COST"

        Select Case UCase(m_ProjList.Item(Me.lstProjects.ListIndex + 1).Status)
            Case "GREEN"
                Me.cmdStatus.BackColor = vbGreen
            Case "YELLOW"
                Me.cmdStatus.BackColor = vbYellow
            Case "RED"
                Me.cmdStatus.BackColor = vbRed
            Case Else
                Me.cmdStatus.BackColor = vbBlack
        End Select


        lRevBL = m_ProjList.Item(Me.lstProjects.ListIndex + 1).IYRevisedBL
        lNE = m_ProjList.Item(Me.lstProjects.ListIndex + 1).IYNE

        If lRevBL <> 0 Then
            Select Case Abs((lNE / lRevBL) - 1)
                Case Is >= 0.1
                    Me.cmdCost.BackColor = vbRed
                Case Is >= 0.05
                    Me.cmdCost.BackColor = vbYellow
                Case Is >= 0
                    Me.cmdCost.BackColor = vbGreen
            End Select
            Me.cmdCost.ForeColor = vbBlack
        Else
            Me.cmdCost.BackColor = vbBlack        'Back to Grey
            Me.cmdCost.ForeColor = vbWhite
        End If

        Me.cmdImplementation.Caption = m_ProjList.Item(Me.lstProjects.ListIndex + 1).ImpDate
        If Me.cmdImplementation.Caption <> "" Then
            Me.cmdImplementation.BackColor = RGB(242, 242, 242)        ''''&HE0E0E0 (Back to Grey)
        Else
            Me.cmdImplementation.BackColor = vbBlack
        End If

    Else
        '''' DALE Me.Caption = "The Local Guide - 0/0"
        lblCount.Caption = "0/0"

        Me.cmdStatus.Caption = ""
        Me.cmdStatus.BackColor = vbBlack

        Me.cmdCost.Caption = ""
        Me.cmdCost.BackColor = vbBlack

        Me.cmdImplementation.Caption = ""
        Me.cmdImplementation.BackColor = vbBlack
    End If

    Call UpdDetCostForm
    Call updDetStatus


End Sub

'---------------------------------------------------------------------------------------
' Method : lstProjects_Change
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub lstProjects_Change()
    Dim lNE As Long
    Dim lRevBL As Long


    If Not gbComboZapCompleted Then Exit Sub

    On Error Resume Next        ''''''''' Ojo si no encuentra Portfolio Plan tab avisar que no encuentra el dashboard workbook!!!!!!!!!

    Application.Goto Sheets("Portfolio Plan").Cells(guDictProj(lstProjects.Value), 1), True

    ''.MsgBox (colPrjInfo.Item(lstProjects.ListIndex + 1).lRow)

    '''' DALE Me.Caption = "The Local Guide -" + Str(Me.lstProjects.ListIndex + 1) + "/" + Str(Me.lstProjects.ListCount)
    lblCount.Caption = CStr(Me.lstProjects.ListIndex + 1) + "/" + CStr(Me.lstProjects.ListCount)

    Select Case UCase(m_ProjList.Item(Me.lstProjects.ListIndex + 1).Status)
        Case "GREEN"
            Me.cmdStatus.BackColor = vbGreen
        Case "YELLOW"
            Me.cmdStatus.BackColor = vbYellow
        Case "RED"
            Me.cmdStatus.BackColor = vbRed
        Case Else
            Me.cmdStatus.BackColor = vbBlack
    End Select

    lNE = Val(m_ProjList.Item(Me.lstProjects.ListIndex + 1).IYNE)
    lRevBL = Val(m_ProjList.Item(Me.lstProjects.ListIndex + 1).IYRevisedBL)

    If lRevBL <> 0 Then
        Select Case Abs((lNE / lRevBL) - 1)
            Case Is >= 0.1
                Me.cmdCost.BackColor = vbRed
            Case Is >= 0.05
                Me.cmdCost.BackColor = vbYellow
            Case Is >= 0
                Me.cmdCost.BackColor = vbGreen
        End Select
        Me.cmdCost.ForeColor = vbBlack
    Else
        Me.cmdCost.BackColor = vbBlack        'Back to Grey
        Me.cmdCost.ForeColor = vbWhite
    End If

    Me.cmdImplementation.Caption = m_ProjList.Item(Me.lstProjects.ListIndex + 1).ImpDate
    If Me.cmdImplementation.Caption <> "" Then
        Me.cmdImplementation.BackColor = RGB(242, 242, 242)        ''''&HE0E0E0
    Else
        Me.cmdImplementation.BackColor = vbBlack
    End If


    Call UpdDetCostForm
    Call updDetStatus

End Sub

'---------------------------------------------------------------------------------------
' File   : frmNavigation
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

    Call ShowMinimizeButton(Me, False)

    With Me
        .StartUpPosition = 0
        .Top = Application.Top + 25
        .Left = Application.Left + Application.Width - Me.Width - 25
        .BackColor = RGB(55, 96, 145)
        .lblSearch.BackColor = RGB(55, 96, 145)
        .lblCount.BackColor = RGB(55, 96, 145)

        'Me.Label3.BackColor = RGB(55, 96, 145)
        'Me.Label4.BackColor = RGB(55, 96, 145)
        .lblStatus.BackColor = RGB(55, 96, 145)
        .lblYear.BackColor = RGB(55, 96, 145)
        .chkIncludesBAU.BackColor = RGB(55, 96, 145)
        .chkIncludesDelivered.BackColor = RGB(55, 96, 145)

        .chkIncludesBAU.Value = False

        If gbFY17 Then .lblYear = "2 0 1 7"

        .lstProjects.SetFocus

    End With
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Unload gFrmDetStatus
    Unload gFrmCostDet

End Sub

'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    If CloseMode = 0 Then
'        gFrmDetStatus.Hide
'        gFrmCostDet.Hide
'        Me.Hide
'    End If
'End Sub




