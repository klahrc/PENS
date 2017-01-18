VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResMan 
   Caption         =   "The Resource Inspector"
   ClientHeight    =   10515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17490
   OleObjectBlob   =   "frmResMan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmResMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim m_ProjList As Collection
Sub updateGrid()
    Dim bSearchTextMatch As Boolean
    Dim lTotal As Long
    Dim i As Long

    Dim bFirstRow As Long

    bSearchTextMatch = False
    lTotal = 0

    bFirstRow = False

    With Me.iGrid1
        .BeginUpdate

        For i = 1 To .RowCount

            Select Case lSumType
                Case bynone
                    bSearchTextMatch = (InStr(1, UCase(.CellValue(i, .ColIndex("Resource Name"))) + UCase(.CellValue(i, .ColIndex("Project Name"))) + _
                    UCase(.CellValue(i, .ColIndex("Role"))) + UCase(.CellValue(i, .ColIndex("Project Code"))), UCase(Me.txtSearch.Value)) > 0)
                Case byResource
                    bSearchTextMatch = (InStr(1, UCase(.CellValue(i, .ColIndex("Resource Name"))), UCase(Me.txtSearch.Value)) > 0)

                Case byRole
                    bSearchTextMatch = (InStr(1, UCase(.CellValue(i, .ColIndex("Role"))), UCase(Me.txtSearch.Value)) > 0)

                Case byproject
                    bSearchTextMatch = (InStr(1, UCase(.CellValue(i, .ColIndex("Project Name"))) + UCase(.CellValue(i, .ColIndex("Project Code"))), UCase(Me.txtSearch.Value)) > 0)

            End Select

            If (Me.cmbResStatus.List(Me.cmbResStatus.ListIndex) = "ALL" Or .CellValue(i, .ColIndex("F/T or Contract")) = Me.cmbResStatus.List(Me.cmbResStatus.ListIndex)) _
                And (Me.cmbCC.List(Me.cmbCC.ListIndex) = "ALL" Or .CellValue(i, .ColIndex("CC")) = Me.cmbCC.List(Me.cmbCC.ListIndex)) _
                And (Me.cmbPortfolio.List(Me.cmbPortfolio.ListIndex) = "ALL" Or .CellValue(i, .ColIndex("Portfolio")) = Me.cmbPortfolio.List(Me.cmbPortfolio.ListIndex)) _
                And bSearchTextMatch Then

                If Not bFirstRow Then
                    bFirstRow = True
                    glFirstRow = i
                End If

                .RowVisible(i) = True

                lTotal = lTotal + 1
                .CellValue(i, .ColIndex("vRow")) = lTotal
            Else
                .RowVisible(i) = False
            End If
        Next
        glTotRows = lTotal


        For i = 1 To .ColCount
            If .ColTag(i) = "PAST" Then
                If chkHidePast.Value Then
                    .ColVisible(i) = False
                Else
                    .ColVisible(i) = True
                End If
            End If
        Next i


        .EndUpdate

        .CellSelected(glFirstRow, 1) = True
        lblRowNumber.Caption = Str(iGrid1.CellValue(glFirstRow, iGrid1.ColIndex("vRow"))) + "/" + Str(glTotRows)

        lblSelectedCount.Caption = 1
        lblSelectedFTE.Caption = 0
        lblSelectedCost.Caption = Format(0, "$#,##0.0")

        .SetFocus
    End With

End Sub

Private Sub updDetStatus()
    With gFrmDetStatus
        If Me.lstProjects.ListCount > 0 Then
            .txtDetStatus.Text = m_ProjList.Item(Me.lstProjects.ListIndex + 1).DetStatus
            .txtDetStatus.SetFocus
            .txtDetStatus.CurLine = 0
        Else
            .txtDetStatus.Text = ""
        End If
    End With
End Sub

Private Sub cmdStatus_Click()

    gFrmDetStatus.Show False

End Sub

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

Private Sub chkHidePast_Click()
    Dim i As Long


    With iGrid1
        .BeginUpdate

        For i = 1 To .ColCount
            If .ColTag(i) = "PAST" Then
                If chkHidePast.Value Then
                    .ColVisible(i) = False
                Else
                    .ColVisible(i) = True
                End If
            End If
        Next i

        .EndUpdate
    End With



End Sub

Private Sub cmbCC_Change()

    If Me.cmbCC.ListIndex > 0 Then
        Me.cmbCC.BackColor = RGB(0, 255, 255)        ' Highlight if search is active
    Else
        Me.cmbCC.BackColor = RGB(242, 242, 242)        ' Back to Grey
    End If

    If gbCompletedFirstLoad Then Call updateGrid
End Sub

Private Sub cmbResource_Change()
    If Me.cmbResource.ListIndex > 0 Then
        Me.cmbResource.BackColor = RGB(0, 255, 255)        ' Highlight if search is active
    Else
        Me.cmbResource.BackColor = RGB(242, 242, 242)        ' Back to Grey
    End If

    If gbCompletedFirstLoad Then Call updateGrid
End Sub

Private Sub cmbPortfolio_Change()

    If Me.cmbPortfolio.ListIndex > 0 Then
        Me.cmbPortfolio.BackColor = RGB(0, 255, 255)        ' Highlight if search is active
    Else
        Me.cmbPortfolio.BackColor = RGB(242, 242, 242)        ' Back to Grey
    End If

    If gbCompletedFirstLoad Then Call updateGrid
End Sub

Private Sub cmbResStatus_Change()
    If Me.cmbResStatus.ListIndex > 0 Then
        Me.cmbResStatus.BackColor = RGB(0, 255, 255)        ' Highlight if search is active
    Else
        Me.cmbResStatus.BackColor = RGB(242, 242, 242)        ' Back to Grey
    End If

    If gbCompletedFirstLoad Then Call updateGrid
End Sub

Private Sub cmdCollapseAll_Click()
    Me.iGrid1.CollapseAllRows
End Sub

Private Sub cmdExpandALL_Click()
    Me.iGrid1.ExpandAllRows
End Sub

Private Sub iGrid1_BeforeRowCollapseExpand(ByVal lRowIfAny As Long, ByVal bNowExpanded As Boolean, bDoDefault As Boolean)

    iGrid1.CurRow = lRowIfAny

End Sub

Private Sub iGrid1_ColHeaderBeginDrag(ByVal lCol As Long, bCancel As Boolean, ByVal bFrozenAreaNotAllowed As Boolean)
    ''' bCancel = True
End Sub

Private Sub iGrid1_ColHeaderClick(ByVal lCol As Long, bDoDefault As Boolean, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long)

    If lSumType <> bynone Then bDoDefault = False

End Sub

Private Sub iGrid1_HeaderRightClick(ByVal lColIfAny As Long, ByVal Shift As Integer, ByVal x As Long, ByVal y As Long, bDoDefault As Boolean)

    If lSumType <> bynone Then bDoDefault = False

End Sub

Private Sub iGrid1_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    With iGrid1
        If KeyCode = vbKeyRight Then
            If .CurCol = .ColCount - 1 Then        ' The very last column is hidden (vRow)
                bDoDefault = False
            End If

        ElseIf KeyCode = vbKeyLeft Then
            If .CurCol = 1 Then
                bDoDefault = False
            End If
        End If
    End With
End Sub

Private Sub iGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim dSum As Double
    Dim lTotal As Long
    Dim arr() As TSelItemInfo


    arr = iGrid1.SelItems.GetArray

    dSum = 0
    lTotal = 0
    For i = 1 To UBound(arr)
        If iGrid1.CellValue(arr(i).Row, arr(i).Col) <> "" Then
            dSum = dSum + Val(iGrid1.CellValue(arr(i).Row, arr(i).Col))
            lTotal = lTotal + 1
        End If
    Next

    lblSelectedCount.Caption = lTotal
    lblSelectedFTE.Caption = dSum
    lblSelectedCost.Caption = Format(dSum * 37.5 * 122, "$#,##0.0")
    lblRowNumber.Caption = Str(iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("vRow"))) + "/" + Str(glTotRows)
    lblStatus.Caption = iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("F/T or Contract"))
    lblBillable.Caption = iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("Billable"))
    lblPortfolio2.Caption = iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("Portfolio"))
    lblSystem.Caption = iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("System"))

End Sub

Private Sub iGrid1_MouseUp(Button As Integer, Shift As Integer, _
    ByVal x As Single, ByVal y As Single, _
    ByVal lRowIfAny As Long, ByVal lColIfAny As Long, bDoDefault As Boolean)


    Dim i As Long
    Dim dSum As Double
    Dim lTotal As Long
    Dim arr() As TSelItemInfo


    arr = iGrid1.SelItems.GetArray

    dSum = 0
    lTotal = 0
    For i = 1 To UBound(arr)
        If iGrid1.CellValue(arr(i).Row, arr(i).Col) <> "" Then
            dSum = dSum + Val(iGrid1.CellValue(arr(i).Row, arr(i).Col))
            lTotal = lTotal + 1
        End If
    Next

    lblSelectedCount.Caption = lTotal
    lblSelectedFTE.Caption = dSum
    lblSelectedCost.Caption = Format(dSum * 37.5 * 122, "$#,##0.0")



    lblRowNumber.Caption = Str(iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("vRow"))) + "/" + Str(glTotRows)
    lblStatus.Caption = iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("F/T or Contract"))
    lblBillable.Caption = iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("Billable"))
    lblPortfolio2.Caption = iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("Portfolio"))
    lblSystem.Caption = iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("System"))

End Sub


Private Sub Image1_Click()

    Call openResourceFile        ' Make it a function to test!!!

End Sub


Private Sub optSumNone_Click()
    If gbCompletedFirstLoad Then

        lSumType = bynone

        Call populate_Grid(frmResMan, bynone)

        iGrid1.SetFocus

        iGrid1.CellSelected(1, 1) = True


        lblCC.Visible = True
        lblResStatus.Visible = True
        lblPortfolio1.Visible = True

        cmbCC.Visible = True
        cmbResStatus.Visible = True
        cmbPortfolio.Visible = True

        cmbCC.ListIndex = 0
        cmbResStatus.ListIndex = 0
        cmbPortfolio.ListIndex = 0

        txtSearch.Value = ""

        '''lblRowNumber.Caption = Str(iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("vRow"))) + "/" + Str(glTotRows)

    End If
End Sub

Private Sub optSumNone_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call optSumNone_Click
End Sub

Private Sub optSumPerCC_Click()

    lSumType = byRole

    Call populate_Grid(frmResMan, byRole)


    iGrid1.SetFocus
    iGrid1.CellSelected(1, 1) = True


    lblCC.Visible = True
    lblResStatus.Visible = False
    lblPortfolio1.Visible = False

    cmbCC.Visible = True
    cmbResStatus.Visible = False
    cmbPortfolio.Visible = False

    cmbCC.ListIndex = 0
    cmbResStatus.ListIndex = 0
    cmbPortfolio.ListIndex = 0

    txtSearch.Value = ""

    '''lblRowNumber.Caption = Str(iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("vRow"))) + "/" + Str(glTotRows)
End Sub

Private Sub optSumPerCC_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call optSumPerCC_Click
End Sub

Private Sub optSumPerProject_Click()

    lSumType = byproject

    Call populate_Grid(frmResMan, byproject)

    iGrid1.SetFocus
    iGrid1.CellSelected(1, 1) = True

    lblCC.Visible = False
    lblResStatus.Visible = False
    lblPortfolio1.Visible = False

    cmbCC.Visible = False
    cmbResStatus.Visible = False
    cmbPortfolio.Visible = False

    cmbCC.ListIndex = 0
    cmbResStatus.ListIndex = 0
    cmbPortfolio.ListIndex = 0

    txtSearch.Value = ""

    '''lblRowNumber.Caption = Str(iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("vRow"))) + "/" + Str(glTotRows)

End Sub

Private Sub optSumPerProject_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call optSumPerProject_Click
End Sub

Private Sub optSumPerResource_Click()

    lSumType = byResource

    Call populate_Grid(frmResMan, byResource)

    iGrid1.SetFocus
    iGrid1.CellSelected(1, 1) = True


    lblCC.Visible = True
    lblResStatus.Visible = True
    lblPortfolio1.Visible = False

    cmbCC.Visible = True
    cmbResStatus.Visible = True
    cmbPortfolio.Visible = False

    cmbCC.ListIndex = 0
    cmbResStatus.ListIndex = 0
    cmbPortfolio.ListIndex = 0

    txtSearch.Value = ""

    ''' lblRowNumber.Caption = Str(iGrid1.CellValue(iGrid1.CurRow, iGrid1.ColIndex("vRow"))) + "/" + Str(glTotRows)
End Sub

Private Sub optSumPerResource_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call optSumPerResource_Click
End Sub

Private Sub txtSearch_Change()

    If Me.txtSearch.Value = "" Then
        Me.txtSearch.BackColor = RGB(242, 242, 242)        ' Back to Grey
    Else
        Me.txtSearch.BackColor = RGB(0, 255, 255)        ' Highlight if search is active
    End If

    If gbCompletedFirstLoad Then Call updateGrid

    txtSearch.SetFocus

End Sub


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
        Me.Caption = "The Local Guide - 1/" + Str(j)

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
        Me.Caption = "The Local Guide - 0/0"

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

Private Sub lstProjects_Change()
    Dim lNE As Long
    Dim lRevBL As Long


    If Not gbComboZapCompleted Then Exit Sub

    On Error Resume Next        ''''''''' Ojo si no encuentra Portfolio Plan tab avisar que no encuentra el dashboard workbook!!!!!!!!!

    Application.Goto Sheets("Portfolio Plan").Cells(guDictProj(lstProjects.Value), 1), True

    ''.MsgBox (colPrjInfo.Item(lstProjects.ListIndex + 1).lRow)

    Me.Caption = "The Local Guide -" + Str(Me.lstProjects.ListIndex + 1) + "/" + Str(Me.lstProjects.ListCount)

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

Private igrid


Private Sub UserForm_Initialize()
    ''StartUpPosition = 0
    ''Top = Application.Top + 25
    ''Left = Application.Left + Application.Width - Width - 25
    BackColor = RGB(55, 96, 145)
    lblSearch.BackColor = RGB(55, 96, 145)
    chkHidePast.BackColor = RGB(55, 96, 145)

    lblLeft.BackColor = RGB(55, 96, 145)
    lblRight.BackColor = RGB(55, 96, 145)
    lblResStatus.BackColor = RGB(55, 96, 145)
    Label4.BackColor = RGB(55, 96, 145)
    lblSelectedCount.BackColor = RGB(55, 96, 145)
    lblSelectedFTE.BackColor = RGB(55, 96, 145)
    lblSelectedCost.BackColor = RGB(55, 96, 145)
    Label6.BackColor = RGB(55, 96, 145)
    Label7.BackColor = RGB(55, 96, 145)
    Label8.BackColor = RGB(55, 96, 145)


    lblCount.BackColor = RGB(55, 96, 145)
    lblFTE.BackColor = RGB(55, 96, 145)
    lblCost.BackColor = RGB(55, 96, 145)

    lblOnlyWithCapacity.BackColor = RGB(55, 96, 145)
    lblStatus.BackColor = RGB(55, 96, 145)
    lblCC.BackColor = RGB(55, 96, 145)

    lblRow.BackColor = RGB(55, 96, 145)
    lblRowNumber.BackColor = RGB(55, 96, 145)


    lblSystem.BackColor = RGB(55, 96, 145)

    lblPortfolio1.BackColor = RGB(55, 96, 145)
    lblPortfolio2.BackColor = RGB(55, 96, 145)

    lblBillable.BackColor = RGB(55, 96, 145)

    optSumPerResource.BackColor = RGB(55, 96, 145)
    optSumPerProject.BackColor = RGB(55, 96, 145)
    optSumPerCC.BackColor = RGB(55, 96, 145)
    optSumNone.BackColor = RGB(55, 96, 145)

    'optStatusFT.BackColor = RGB(55, 96, 145)
    'optStatusContractor.BackColor = RGB(55, 96, 145)
    'optStatusBoth.BackColor = RGB(55, 96, 145)

    optResAvailability.BackColor = RGB(55, 96, 145)
    optResDemand.BackColor = RGB(55, 96, 145)
    'optResSupply.BackColor = RGB(55, 96, 145)

    Frame1.BackColor = RGB(55, 96, 145)
    Frame2.BackColor = RGB(55, 96, 145)

    ' lstResources.SetFocus
End Sub

