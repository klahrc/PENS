Attribute VB_Name = "MGenReps"
'---------------------------------------------------------------------------------------
' File   : MGenReps
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Generates Data for analysis
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------
Private Const msMODULE As String = "MGenReps"

Const adCmdText = 1                                        ' Required for early binding

'---------------------------------------------------------------------------------------
' Method : ProduceReports
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Produce Pivot tables
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function ProduceReports(ByVal sType As String, ByVal sFilePP As String, ByVal sFolderRS As String, ByVal sFileNameRS As String) As Double
    Const sSOURCE As String = "ProduceReports"

    Dim PC As Object

    '''Dim sdConPP As WorkbookConnection
    ''' Dim sdConRS As WorkbookConnection
    Dim sConnPP As String
    Dim sConnRS As String
    Dim sCommandTextPP As String
    Dim sCommandTextRS As String

    Dim dGrandTotalFTE_NE As Double
    Dim dGrandTotalPROJ_SUMMARY As Double
    Dim dGrandTotalCC_CAPACITY As Double
    Dim dGrandTotalNES As Double

    Dim dGrandTotalPS As Double
    Dim dGrandTotalCS As Double
    Dim dGrandTotalIGLS As Double

    Dim dRoundUp As Double
    Dim dRoundDown As Double

    Dim sRepFilename As String
    Dim wbNew As Workbook

    Dim ufUpdate As frmRptStatus

    Dim bPivot As Boolean

    Set ufUpdate = New frmRptStatus

    If sType <> "CC" Then
        sRepFilename = "PENS_Rep_" + Format(Now(), "mmmddyyyy_hhmmss") + ".xlsx"
        Set wbNew = Workbooks.Add
        ''''wbNew.SaveAs Filename:=gsLocal_Folder & "\" & sRepFilename

        Dim sPPDB As String
        sPPDB = gsLocal_Folder + "\" + gwsConfig.Range(gsDB_NAME).Value

        Dim cnConn As Object
        Set cnConn = CreateObject("ADODB.Connection")
        With cnConn
            .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
            .Open sPPDB
        End With


        ' Set the command text.
        Dim cmdCommand As Object
        Set cmdCommand = CreateObject("ADODB.Command")
        Set cmdCommand.activeconnection = cnConn
        With cmdCommand
            .CommandText = "Select * From tbl_PortfolioPlan"
            .CommandType = adCmdText
            .Execute
        End With


        ' Open the recordset.
        Dim rstRecordset As Object
        Set rstRecordset = CreateObject("ADODB.Recordset")
        Set rstRecordset.activeconnection = cnConn
        rstRecordset.Open cmdCommand


        ' Create a PivotTable cache and report.
        Set PC = wbNew.PivotCaches.Create(SourceType:=xlExternal)
        Set PC.Recordset = rstRecordset
    End If


    With ufUpdate
        .cmdOK.Visible = False
        .lblCompleted.Caption = "Please wait..."

        .Show False

        '--------------------------------------------------------------------------
        If sType = "PT" Then
            .lstStatus.AddItem "FTE NE Summary report in progress..."
            .lstStatus.TopIndex = ufUpdate.lstStatus.ListCount - 1
            DoEvents


            dGrandTotalFTE_NE = GenPT_FTE_NE(wbNew, PC)


            .lstStatus.Selected(ufUpdate.lstStatus.ListCount - 1) = True
            .lstStatus.List(ufUpdate.lstStatus.ListCount - 1) = "FTE NE Summary report complete"
        End If
        '--------------------------------------------------------------------------
        If sType = "PT" Then
            .lstStatus.AddItem "Project Summary report in progress..."
            .lstStatus.TopIndex = ufUpdate.lstStatus.ListCount - 1
            DoEvents

            dGrandTotalPROJ_SUMMARY = GenPT_PROJ_SUMMARY(wbNew, PC)

            .lstStatus.Selected(ufUpdate.lstStatus.ListCount - 1) = True
            .lstStatus.List(ufUpdate.lstStatus.ListCount - 1) = "Project Summary report complete"
        End If

        'CC Capacity tab

        '    sConnRS = "oledb;Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sFileRS & ";" & "Extended Properties=Excel 12.0 Macro;"
        '    sCommandText = "select * from [CC Capacity$B3:N100]"
        '    On Error Resume Next
        '    ThisWorkbook.Connections("conRS").Delete
        '    On Error GoTo 0
        '    Set sdConRS = ActiveWorkbook.Connections.Add(Name:="conRS", Description:="conRS", connectionstring:=sConnRS, CommandText:=sCommandText, lcmdtype:=xlCmdSql)


        '--------------------------------------------------------------------------
        ''''If sType = "ALL" Or sType = "CC" Then
        If sType = "CC" Then                               'No longer included in ALL

            If RefreshTemplates() Then                     ' Ensure the template is available
                .lstStatus.AddItem "CC Capacity vs Demand report in progress..."
                .lstStatus.TopIndex = ufUpdate.lstStatus.ListCount - 1
                DoEvents

                'dGrandTotalCC_CAPACITY = GenPT_CC_CAPACITY(wbNew, PC, sFolderRS, sFileNameRS)

                dGrandTotalCC_CAPACITY = GenPT_CC_CAPACITY(sFolderRS, sFileNameRS)

                .lstStatus.Selected(ufUpdate.lstStatus.ListCount - 1) = True
                .lstStatus.List(ufUpdate.lstStatus.ListCount - 1) = "CC Capacity vs Demand report complete"
            Else
                MsgBox "The Capacity report cannot be generated due to a missing template..." & _
                vbNewLine & vbNewLine & "Please contact Project Support office for assistance", vbOKOnly Or vbExclamation, "Oops!"
            End If
        End If
        '--------------------------------------------------------------------------
        If sType = "PD" Then
            bPivot = (MsgBox("Do you want the Pivot Table too?", vbQuestion + vbYesNo) = vbYes)

            .lstStatus.AddItem "Portfolio Project Data with Pivot report in progress..."
            .lstStatus.TopIndex = ufUpdate.lstStatus.ListCount - 1

            DoEvents

            dGrandTotalNES = Gen_NES(wbNew, sFilePP, bPivot)

            .lstStatus.Selected(ufUpdate.lstStatus.ListCount - 1) = True
            .lstStatus.List(ufUpdate.lstStatus.ListCount - 1) = "Portfolio Project Data with Pivot report complete"
        End If
        '--------------------------------------------------------------------------
        If sType = "PDSGC" Then                            ' Short version is disable yet!!!

            .lstStatus.AddItem "Portfolio Project Data (Short version GC) report in progress..."
            .lstStatus.TopIndex = ufUpdate.lstStatus.ListCount - 1
            DoEvents

            dGrandTotalNES = Gen_NES_Short_GC(wbNew, sFilePP)

            .lstStatus.Selected(ufUpdate.lstStatus.ListCount - 1) = True
            .lstStatus.List(ufUpdate.lstStatus.ListCount - 1) = "Portfolio Project Data (Short version GC) report complete"
        End If
        '--------------------------------------------------------------------------
        If sType = "PS" Then                            ' Short version is disable yet!!!

            .lstStatus.AddItem "Plan Seasonality report in progress..."
            .lstStatus.TopIndex = ufUpdate.lstStatus.ListCount - 1
            DoEvents

            dGrandTotalPS = GenPT_PS(wbNew, PC)

            .lstStatus.Selected(ufUpdate.lstStatus.ListCount - 1) = True
            .lstStatus.List(ufUpdate.lstStatus.ListCount - 1) = "Plan Seasonality report complete"
        End If
        '--------------------------------------------------------------------------
        If sType = "CS" Then                            ' Short version is disable yet!!!

            .lstStatus.AddItem "Consulting Seasonality report in progress..."
            .lstStatus.TopIndex = ufUpdate.lstStatus.ListCount - 1
            DoEvents

            dGrandTotalNES = GenPT_CS(wbNew, PC)

            .lstStatus.Selected(ufUpdate.lstStatus.ListCount - 1) = True
            .lstStatus.List(ufUpdate.lstStatus.ListCount - 1) = "Consulting Seasonality report complete"
        End If
        '--------------------------------------------------------------------------
        If sType = "IGLS" Then                            ' Short version is disable yet!!!

            .lstStatus.AddItem "IG Labour Seasonality report in progress..."
            .lstStatus.TopIndex = ufUpdate.lstStatus.ListCount - 1
            DoEvents

            dGrandTotalNES = GenPT_IGLS(wbNew, PC)

            .lstStatus.Selected(ufUpdate.lstStatus.ListCount - 1) = True
            .lstStatus.List(ufUpdate.lstStatus.ListCount - 1) = "IG Labour Seasonality report complete"
        End If
        '--------------------------------------------------------------------------
        .lblCompleted.Caption = "Click OK to continue..."
        '''''.lblCompleted.Font.Bold = False

        .cmdOK.Visible = True
        .cmdOK.SetFocus
    End With


    ''''Application.ScreenUpdating = True

    dRoundUp = Ceiling(dGrandTotalFTE_NE, 0.01)
    dRoundDown = Ceiling(dGrandTotalFTE_NE, -0.01)

    If (dGrandTotalFTE_NE - dRoundDown) < (dRoundUp - dGrandTotalFTE_NE) Then
        dGrandTotalFTE_NE = dRoundDown
    Else
        dGrandTotalFTE_NE = dRoundUp
    End If

    dRoundUp = Ceiling(dGrandTotalPROJ_SUMMARY, 0.01)
    dRoundDown = Ceiling(dGrandTotalPROJ_SUMMARY, -0.01)

    If (dGrandTotalPROJ_SUMMARY - dRoundDown) < (dRoundUp - dGrandTotalPROJ_SUMMARY) Then
        dGrandTotalPROJ_SUMMARY = dRoundDown
    Else
        dGrandTotalPROJ_SUMMARY = dRoundUp
    End If

    If (dGrandTotalFTE_NE = dGrandTotalPROJ_SUMMARY) Then
        ProduceReports = dGrandTotalFTE_NE
    Else
        ProduceReports = -1
    End If

    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In wbNew.Sheets
        If Left(ws.Name, 5) = "Sheet" Then
            ws.Delete
        End If
    Next
    Application.DisplayAlerts = True

    wbNew.SaveAs FileName:=gsLocal_Folder & "\" & sRepFilename

    gsLastReport = wbNew.Name

    Set wbNew = Nothing

    If sType <> "CC" Then
        cnConn.Close

        Set cnConn = Nothing
        Set cmdCommand = Nothing
        Set rstRecordset = Nothing
    End If

    Set ufUpdate = Nothing


ErrorExit:
    ' Clean up
    Exit Function

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
