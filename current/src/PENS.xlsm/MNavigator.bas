Attribute VB_Name = "MNavigator"
'---------------------------------------------------------------------------------------
' File   : MNavigator
' Author : cklahr
' Date   : 1/17/2016
' Purpose: Module responsible of generating the Pivot Tables
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------

Const adCmdText = 1        ' Required for early binding
Private Const msMODULE As String = "MNavigator"

'---------------------------------------------------------------------------------------
' Method : loadProjects
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function loadProjects(ByRef colPrjList As Collection, ByRef lstProj As MSForms.ListBox, ByVal sDL As String, ByVal sAS As String, ByVal sCat As String, _
    ByVal sSearch As String) As Long

    Dim i As Long
    Dim pInfo As clsPrjInfo
    Dim colPrj As Collection


    Set colPrj = New Collection

    gbComboZapCompleted = False

    With lstProj

        While .ListCount > 0
            .RemoveItem (.ListCount - 1)
        Wend

        i = 0
        For Each pInfo In gColPrjInfo
            If (Left(sDL, 3) = "ALL" Or sDL = pInfo.DeliveryLeader) And (Left(sAS, 3) = "ALL" Or sAS = pInfo.ActivationStatus) And (Left(sCat, 3) = "ALL" Or sCat = pInfo.Category) And _
                (InStr(1, UCase(pInfo.ProjName) + " - " + UCase(pInfo.ProjCode) + " (" + UCase(pInfo.PM) + ")", UCase(sSearch)) > 0) Then

                If ((pInfo.Category <> "9 - NDM" Or gFrmNavPanel.chkIncludesBAU.Value) And _
                    (pInfo.ActivationStatus <> "Delivered" Or gFrmNavPanel.chkIncludesDelivered.Value)) Then

                    .AddItem (pInfo.ProjName + " - " + pInfo.ProjCode + " (" + pInfo.PM + ")")
                    colPrj.Add pInfo
                    i = i + 1

                End If
            End If
        Next

        gbComboZapCompleted = True

        If .ListCount > 0 Then
            .Selected(0) = True
        End If
    End With


    Set colPrjList = colPrj

    Set colPrj = Nothing

    loadProjects = i

End Function

'---------------------------------------------------------------------------------------
' Method : loadComboFromTable
' Author : cklahr
' Date   : 1/9/2017
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub loadComboFromTable(ByRef cmbList As ComboBox, ByVal sField As String, ByVal cn As Object)
    Dim cmdCommand As Object
    Dim rstRecordset As Object

    ' Set the command text.
    Set cmdCommand = CreateObject("ADODB.Command")
    Set cmdCommand.activeconnection = cn
    With cmdCommand
        '''''@@@@@@@@ .CommandText = "SELECT [Names] FROM tbl_PortfolioPlan WHERE [Roles] = 'Actual - HW Amort'" 'Budget
        ''''.CommandText = "SELECT distinct [Delivery Leader] FROM tbl_PortfolioPlan"
        .CommandText = "SELECT distinct [" + sField + "] FROM tbl_PortfolioPlan"
        .CommandType = adCmdText
        .Execute
    End With

    ' Open the recordset.
    Set rstRecordset = CreateObject("ADODB.Recordset")
    Set rstRecordset.activeconnection = cn
    rstRecordset.Open cmdCommand


    cmbList.AddItem ("ALL " + sField)


    With rstRecordset
        While Not .EOF
            If Len(.Fields(sField)) > 1 Then
                cmbList.AddItem .Fields(sField)
            End If
            rstRecordset.MoveNext
        Wend
    End With


    Set cmdCommand = Nothing
    Set rstRecordset = Nothing


End Sub

'---------------------------------------------------------------------------------------
' Method : loadNavPanel
' Author : cklahr
' Date   : 3/19/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function loadNavPanel() As Boolean


    Dim sPPDB As String
    Dim cnConn As Object
    Dim cmdCommand As Object
    Dim rstRecordset As Object

    Dim sOldProjCode As String        ' Everything is a string!
    Dim sProjName As String
    Dim sProjCode As String
    Dim sPM As String

    Dim i As Long
    Dim prjInfo As clsPrjInfo

    Dim lNextRow As Long

    Dim j As Long


    gbCompletedNavPanelLoad = False

    sPPDB = gsLocal_Folder + "\" + gwsConfig.Range(gsDB_NAME).Value        'PDASH.accdb


    Set cnConn = CreateObject("ADODB.Connection")
    With cnConn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
        .Open sPPDB
    End With


    ' Set the command text.
    Set cmdCommand = CreateObject("ADODB.Command")
    Set cmdCommand.activeconnection = cnConn
    With cmdCommand
        .CommandText = "SELECT * FROM tbl_PortfolioPlan ORDER BY RowID"
        .CommandType = adCmdText
        .Execute
    End With

    ' Open the recordset.
    Set rstRecordset = CreateObject("ADODB.Recordset")
    Set rstRecordset.activeconnection = cnConn
    rstRecordset.Open cmdCommand

    sOldProjCode = ""

    '**********************************************************
    On Error GoTo kk
    '**********************************************************

    With rstRecordset
        While Not .EOF
            ''''.Sheets("Sheet1").Cells(i, 1).Value = CDbl(Trim(rstRecordset.fields("Names")))
            If Len(.Fields("Project Code")) > 1 And .Fields("Project Code") <> sOldProjCode Then        ' And .fields("Delivery Leader") = "80 Gabriel Cincotta" Then
                Set prjInfo = New clsPrjInfo

                sProjName = .Fields("Project Name")
                sProjCode = .Fields("Project Code")

                Debug.Print sProjCode
                Debug.Print .Fields("RowID")

                prjInfo.lRow = Val(.Fields("RowID"))
                prjInfo.ProjCode = .Fields("Project Code")

                If .Fields("Project Name") <> vbNull Then prjInfo.ProjName = .Fields("Project Name")
                If .Fields("Delivery Leader") <> vbNull Then prjInfo.DeliveryLeader = .Fields("Delivery Leader")
                If .Fields("Activation Status") <> vbNull Then prjInfo.ActivationStatus = .Fields("Activation Status")
                If .Fields("CAT") <> vbNull Then prjInfo.Category = .Fields("CAT")

                prjInfo.Status = Trim((FetchValue(cnConn, "Names", "Roles", "Actual - MI", "Project Code", .Fields("Project Code"))))        ' Project Status
                prjInfo.IYNE = Format(FetchValue(cnConn, "Do not remove1", "Roles", "NE (Actual + Rem. Plan)", "Project Code", .Fields("Project Code")), "#0")        ' Dashboard NE
                prjInfo.RepNE = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW Maint", "Project Code", rstRecordset.Fields("Project Code"))), "$#,##0")        ' Reporting NE

                prjInfo.ImpDate = Format(FetchValue(cnConn, "Names", "Roles", "Actual - IG", "Project Code", .Fields("Project Code")), "dd-mmm-yy")        ' Implementation date

                If gbFY17 Then
                    prjInfo.IYBudget = Format((FetchValue(cnConn, "Names", "Roles", "Actual - SW SAAS", "Project Code", rstRecordset.Fields("Project Code"))), "#0")        ' Budget
                Else
                    prjInfo.IYBudget = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW Amort", "Project Code", rstRecordset.Fields("Project Code"))), "#0")        ' Budget
                End If
                prjInfo.IYRevisedBL = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW < $5K", "Project Code", rstRecordset.Fields("Project Code"))), "#0")        ' Revised BL
                prjInfo.YTDActuals = Format(FetchValue(cnConn, "Names", "Roles", "Actual - Travel", "Project Code", rstRecordset.Fields("Project Code")), "#0")        'YTD Actuals
                '''prjInfo.PM = Trim(FetchValue(cnConn, "Names", "Roles", "PM", "Project Code", rstRecordset.fields("Project Code")))        'PM Name

                If .Fields("Project Manager") <> vbNull Then
                    prjInfo.PM = .Fields("Project Manager")
                Else
                    prjInfo.PM = "No PM found"
                End If

                'If prjInfo.PM = "" Then
                '    prjInfo.PM = "No PM found"
                'End If

                lNextRow = Val(.Fields("RowID")) + 2
                prjInfo.DetStatus = FetchValue(cnConn, "Report", "RowID", Trim(Str(lNextRow)), "Project Code", .Fields("Project Code"))        ' Detailed Project Status


                gColPrjInfo.Add prjInfo
                sOldProjCode = .Fields("Project Code")

                guDictProj.Add sProjName + " - " + sProjCode + " (" + prjInfo.PM + ")", Val(.Fields("RowID"))

                Set prjInfo = Nothing

            End If

            '''i = i + 1
            rstRecordset.MoveNext
        Wend
    End With

bosta:

    Call loadComboFromTable(gFrmNavPanel.cmbDeliveryLeader, "Delivery Leader", cnConn)        '''' Make it a function to test!!!!
    Call loadComboFromTable(gFrmNavPanel.cmbActivationStatus, "Activation Status", cnConn)        '''' Make it a function to test!!!!
    Call loadComboFromTable(gFrmNavPanel.cmbCategory, "Cat", cnConn)        '''' Make it a function to test!!!!

    gFrmNavPanel.cmbDeliveryLeader.ListIndex = 0
    gFrmNavPanel.cmbActivationStatus.ListIndex = 0

    gbCompletedNavPanelLoad = True

    ' This will trigger "updateNavigationPanel"!!!
    gFrmNavPanel.cmbCategory.ListIndex = 0


    Set cnConn = Nothing
    Set cmdCommand = Nothing
    Set rstRecordset = Nothing

    loadNavPanel = True

    Exit Function

kk:
    'i = 1
    '''Resume bosta

End Function
