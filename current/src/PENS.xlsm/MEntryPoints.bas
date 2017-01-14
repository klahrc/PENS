Attribute VB_Name = "MEntryPoints"
'---------------------------------------------------------------------------------------
' File   : MEntryPoints
' Author : cklahr
' Date   : 1/31/2016
' Purpose: This Module contains all of the entry point procedures into the application
' Arguments:None
' Pending: None
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------
Private Const msMODULE As String = "MEntryPoints"

'---------------------------------------------------------------------------------------
' Method : PENS2007REP_ALL
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub PENS2007REP_ALL(control As Object)

    Call PENS2007Reports("ALL")

    ' STILL MISSING THE ERROR HANDLING!!!!

End Sub

'---------------------------------------------------------------------------------------
' Method : PENS2007REP_CC
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub PENS2007REP_CC(control As Object)

    Call PENS2007Reports("CC")

    ' STILL MISSING THE ERROR HANDLING!!!!

End Sub

'---------------------------------------------------------------------------------------
' Method : PENS2007REP_PD
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub PENS2007REP_PD(control As Object)

    Call PENS2007Reports("PD")

    ' STILL MISSING THE ERROR HANDLING!!!!

End Sub

'---------------------------------------------------------------------------------------
' Method : PENS2007REP_PD_Short_GC
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub PENS2007REP_PD_Short_GC(control As Object)

    Call PENS2007Reports("PDSGC")

    ' STILL MISSING THE ERROR HANDLING!!!!

End Sub
Public Sub PENS2007REP_PT(control As Object)

    Call PENS2007Reports("PT")

    ' STILL MISSING THE ERROR HANDLING!!!!

End Sub

'---------------------------------------------------------------------------------------
' Method : PENS2007REP_PS
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:Plan Seasonality
'---------------------------------------------------------------------------------------
Public Sub PENS2007REP_PS(control As Object)

    Call PENS2007Reports("PS")

    ' STILL MISSING THE ERROR HANDLING!!!!

End Sub

'---------------------------------------------------------------------------------------
' Method : PENS2007REP_CS
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub PENS2007REP_CS(control As Object)

    Call PENS2007Reports("CS")

    ' STILL MISSING THE ERROR HANDLING!!!!

End Sub


'---------------------------------------------------------------------------------------
' Method : PENS2007REP_IGLS
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub PENS2007REP_IGLS(control As Object)

    Call PENS2007Reports("IGLS")

    ' STILL MISSING THE ERROR HANDLING!!!!

End Sub
'---------------------------------------------------------------------------------------
' Method : OpenAndAdjustDash
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function OpenAndAdjustDash(ByVal bDashIsOpen As Boolean) As Workbook
    Dim wbNew As Workbook

    If (MsgBox("In order to guarantee a smooth navigation experience, PENS requires to remove all filters and frozen panes from the Portfolio Plan. Are you OK to proceed? " + _
        vbNewLine, vbQuestion + vbYesNo, "PENS - Navigation") = vbYes) Then

        If Not bDashIsOpen Then
            Set wbNew = openDashboardFile()
        Else
            Set wbNew = Application.Workbooks(gsPP_Filename)
        End If

        If Not wbNew Is Nothing Then
            wbNew.Sheets("Portfolio Plan").Select
            ActiveWindow.FreezePanes = False
            wbNew.Sheets("Portfolio Plan").AutoFilterMode = False

            wbNew.Sheets("Portfolio Plan").Range("C4").Select
            ActiveWindow.FreezePanes = True

            wbNew.Sheets("Portfolio Plan").Rows(1).RowHeight = 21
            wbNew.Sheets("Portfolio Plan").Rows(2).RowHeight = 0
        Else
            ' ERRORROR!!!!!
        End If
    End If

    Set OpenAndAdjustDash = wbNew

End Function

'---------------------------------------------------------------------------------------
' Method : PENS2007Reports
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS2007Reports(ByVal sType As String)          ' Do I need to pass control as a parameter?????????????????????
    Dim sPathPP As String
    Dim sFolderRS As String
    Dim sFileNameRS As String
    Dim iRet As Double
    Dim i As Integer
    Dim lOK2Proceed As Boolean

    Const sSOURCE As String = "PENS2007Reports"

    On Error GoTo PENS2007REPORTS_Error

    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If


    Call RetrieveAppSettings

    ' Automatic check (Manual = false)
    If Not CheckAndUpdate(gbConnected2Network, False) Then

        lOK2Proceed = False

        If GetSetting("PENS", "General", "Version") = "" Then
            lOK2Proceed = True
        ElseIf CDbl(gwsConfig.Range(gsPENS_VERSION)) = CDbl(GetSetting("PENS", "General", "Version")) Then
            lOK2Proceed = True
        End If

        If lOK2Proceed Then

            If Not gbConnected2Network Then
                MsgBox "No access to the network. You can still use Portfolio Ensemble with local files..."

                SaveSetting "PENS", "General", "Local_Folder", "True"
                gbUse_Local_Folder = True                  'GetSetting(, , )
            Else
                Call LogActivity("PENS2007Reports")        ' Make it function to test?
            End If

            ' If the user pressed the Cancel button, raise a custom
            ' user cancel error. This will cause the central error
            ' handler to exit the program without displaying an
            ' error message.
            '#########################If frmMain.UserCancel Then Err.Raise glUSER_CANCEL


            If SourceDataLoad(sPathPP, sFolderRS, sFileNameRS) Then

                ' Saving in case some folder changed during SourceDataLoad
                Call SaveAppSettings                       ' Make it function?

                Call Access_MakeTable(sPathPP, False)      ' Make it a function to test!!!!

                iRet = ProduceReports(sType, sPathPP, sFolderRS, sFileNameRS)

                '!!!If Not ProducePLD(sPathPP) Then
                'Raise error, etc
                '!!!!End If



                '@@@@@@@@@  OJO VER COMO IMPLEMENTO ESTO EN EL frmRptStatus!!!!!!!!!!!!!!!!!!!
                If (iRet > 0) Then
                    '!!!!!!lblChkSum.Caption = "Pivot Checksum OK (" + CStr(iRet) + ")"
                    '@@@@@@@@@@@@@@ i = MsgBox("All Good...", vbOKOnly, "CONGRATS!")
                ElseIf iRet = -1 Then                      'CheckSum failed
                    '!!!!!!lblChkSum.Caption = "WARNING:Pivot Checksum failed!"
                    '@@@@@@@@@i = MsgBox("Oops...Pivot Checksum failed", vbCritical, "ERROR")
                ElseIf iRet = -2 Then                      'Couldn't connect to Source Data (Portfolio Plan)
                    '@@@@@@@@@@@i = MsgBox("Can't connect to source file/s" + Chr(10) + "Contact your provider...", vbCritical, "ERROR")
                    '''Close D4A file!
                    ActiveWorkbook.Close savechanges:=False
                End If

                '!!!!!!!!!!lblChkSum.Caption = "Waiting for Pivot Checksum results..."
            Else


            End If

        Else
            MsgBox "Please restart Excel to complete the update process...", vbCritical, "PENS"
        End If
    End If


    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Close #giDebugFile
    '''Unload gFrmExtPanel
    ''''Set gFrmExtPanel = Nothing

    On Error GoTo 0

    Exit Sub

PENS2007REPORTS_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

'''''''''''' PENS2007EXTRACT NO LONGER IN USE!!!
'---------------------------------------------------------------------------------------
' Method : PENS2007Extract
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS2007Extract(control As Object)              ' Do I need to pass control as a parameter?????????????????????
    Dim i As Long
    Dim lOK2Proceed As Boolean

    Const sSOURCE As String = "PENS2007Extract"

    On Error GoTo PENS2007EXTRACT_Error

    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If


    ' frmExtraction toolbar loading
    '''Set gFrmExtPanel = New frmExtraction
    ''''Load gFrmExtPanel

    ''''gFrmExtPanel.Show (False)                              ' Non-modal form. The logic continues...

    ' Automatic check (Manual = false)
    If Not CheckAndUpdate(gbConnected2Network, False) Then


        lOK2Proceed = False

        If GetSetting("PENS", "General", "Version") = "" Then
            lOK2Proceed = True
        ElseIf CDbl(gwsConfig.Range(gsPENS_VERSION)) = CDbl(GetSetting("PENS", "General", "Version")) Then
            lOK2Proceed = True
        End If

        If lOK2Proceed Then

            If Not gbConnected2Network Then
                MsgBox "No access to the network. You can still use Portfolio Ensemble with local files..."
                '''gwsConfig.Range(gsUSE_LOCAL_DATA).Value = True

                i = SaveSetting("PENS", "General", "Local_Folder", "True")
                gbUse_Local_Folder = GetSetting("PENS", "General", "Use_Local_Folder")
            End If

            ' If the user pressed the Cancel button, raise a custom
            ' user cancel error. This will cause the central error
            ' handler to exit the program without displaying an
            ' error message.
            '#########################If frmMain.UserCancel Then Err.Raise glUSER_CANCEL

        Else
            MsgBox "Please restart Excel to complete the update process...", vbCritical, "PENS"
        End If

    End If
    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Close #giDebugFile
    '''Unload gFrmExtPanel
    ''''Set gFrmExtPanel = Nothing

    On Error GoTo 0

    Exit Sub

PENS2007EXTRACT_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : PENS2007Navigation
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS2007Navigation(control As Object)           ' Do I need to pass control as a parameter?????????????????????
    Dim sPathPP As String
    Dim sFolderRS As String
    Dim sFileNameRS As String
    Dim dictProj As Object
    Dim lOK2Proceed As Boolean
    Dim wbNew As Workbook
    Dim bDashIsOpen As Boolean


    Const sSOURCE As String = "PENS2007Navigation"

    On Error GoTo PENS2007NAVIGATION_Error

    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    Set gColPrjInfo = New Collection
    Set guDictProj = CreateObject("scripting.dictionary")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If

    Call RetrieveAppSettings                               'Make it function to test??


    ' Automatic check (Manual = false)
    If Not CheckAndUpdate(gbConnected2Network, False) Then

        lOK2Proceed = False

        If GetSetting("PENS", "General", "Version") = "" Then
            lOK2Proceed = True
        ElseIf CDbl(gwsConfig.Range(gsPENS_VERSION)) = CDbl(GetSetting("PENS", "General", "Version")) Then
            lOK2Proceed = True
        End If

        If lOK2Proceed Then

            If Not gbConnected2Network Then
                MsgBox "No access to the network. You can still use Portfolio Ensemble with local files..."
                '''gwsConfig.Range(gsUSE_LOCAL_DATA).Value = True

                SaveSetting "PENS", "General", "Use_Local_Folder", "True"
                gbUse_Local_Folder = GetSetting("PENS", "General", "Use_Local_Folder")
            Else
                Call LogActivity("PENS2007Navigation")     ' Make it function to test?
            End If


            If SourceDataLoad(sPathPP, sFolderRS, sFileNameRS) Then

                ' Saving in case some folder changed during SourceDataLoad
                Call SaveAppSettings                       ' Make it function?

                Call Access_MakeTable(sPathPP, True)       ' Make it a function to test!!!!

                ' ddddddddddddddddddddddddddddddddddddd


                bDashIsOpen = IsDashboardOpen()
                If Not bDashIsOpen Then
                    If (MsgBox("Would you like PENS to open the Portfolio Plan now? " + vbNewLine, vbQuestion + vbYesNo, "PENS - Navigation") = vbYes) Then
                        Set wbNew = OpenAndAdjustDash(bDashIsOpen)
                    Else
                        ' Make it a tip...
                        'MsgBox "TIP: You can open the Portfolio Plan at any time clicking on the Local Guide image..." + vbNewLine + vbNewLine & _
                        '       "IMPORTANT: If you have the Porfolio Plan spreadsheet already open, you must remove filters and freezes to ensure a smooth navigation!", vbInformation + vbOKOnly, "PENS - Help"
                    End If
                Else
                    Set wbNew = OpenAndAdjustDash(bDashIsOpen)
                End If

                DoEvents


                ' frmExtraction toolbar loading
                Set gFrmNavPanel = New frmNavigation
                Load gFrmNavPanel

                gFrmNavPanel.Show (False)

                ' frmCostDetails form loading
                Set gFrmCostDet = New frmCostDetails
                Load gFrmCostDet

                Set gFrmDetStatus = New frmDetStatus
                Load gFrmDetStatus

                If loadNavPanel() Then

                    gFrmNavPanel.lstProjects.SetFocus

                Else
                    ' ERRRORRRRR
                End If


                ' Assuming there is at least one project (goto the last and come back to the first in the Projects listbox)
                gFrmNavPanel.lstProjects.ListIndex = gFrmNavPanel.lstProjects.ListCount - 1 ''ojo que esto puede fallar en algun momento
                gFrmNavPanel.lstProjects.ListIndex = 0     ' This time goes always to the first!!!

            End If
        End If                                             'OK2Proceed

        Else                                                   'CheckAndUpdate

        MsgBox "Please restart Excel to complete the update process...", vbCritical, "PENS"
    End If


    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Close #giDebugFile
    Unload gFrmNavPanel
    Set gFrmNavPanel = Nothing
    Set guDictProj = Nothing
    Set gColPrjInfo = Nothing

    On Error GoTo 0

    Exit Sub

PENS2007NAVIGATION_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub


'---------------------------------------------------------------------------------------
' Method : PENS_Inspector
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS_Inspector(control As Object)               ' Do I need to pass control as a parameter?????????????????????
    Dim sPathPP As String
    Dim sFolderRS As String
    Dim sFileNameRS As String
    Dim dictProj As Object
    Dim lOK2Proceed As Boolean
    Dim wbNew As Workbook
    Dim bDashIsOpen As Boolean


    Const sSOURCE As String = "PENS_Inspector"

    On Error GoTo PENS_Inspector_Error


    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If

    Call RetrieveAppSettings                               'Make it function to test??

    ' Automatic check (Manual = false)
    If Not CheckAndUpdate(gbConnected2Network, False) Then

        lOK2Proceed = False

        If GetSetting("PENS", "General", "Version") = "" Then
            lOK2Proceed = True
        ElseIf CDbl(gwsConfig.Range(gsPENS_VERSION)) = CDbl(GetSetting("PENS", "General", "Version")) Then
            lOK2Proceed = True
        End If

        If lOK2Proceed Then

            If Not gbConnected2Network Then
                MsgBox "No access to the network. You can still use Portfolio Ensemble with local files..."
                '''gwsConfig.Range(gsUSE_LOCAL_DATA).Value = True

                SaveSetting "PENS", "General", "Use_Local_Folder", "True"
                gbUse_Local_Folder = True
            Else
                Call LogActivity("PENS_Inspector")         ' Make it function to test?
            End If

            ' If the user pressed the Cancel button, raise a custom
            ' user cancel error. This will cause the central error
            ' handler to exit the program without displaying an
            ' error message.
            '#########################If frmMain.UserCancel Then Err.Raise glUSER_CANCEL

            '            Set gFrmSettings = New frmSettings
            '            Load gFrmSettings
            '
            '            gbInitialized = False                          ' Some activities should wait until the form is initialized
            '            If Not gFrmSettings.Initialize() Then          ' If UserForm initialization failed, raise a custom error.
            '                '    Err.Raise glCANT_INITIALIZE
            '            Else
            '                gbInitialized = True
            '
            '                gFrmSettings.Show (False)                  ' Show form modeless
            '
            '                DoEvents                                   ' Do I need this?????????????????????????????
            '
            '
            '
            '                ' If the user pressed the Cancel button, raise a custom
            '                ' user cancel error. This will cause the central error
            '                ' handler to exit the program without displaying an
            '                ' error message.
            '                '#########################If frmMain.UserCancel Then Err.Raise glUSER_CANCEL
            '
            '
            '                '''''''''gFrmSettings.Show (False)
            '            End If

            If SourceDataLoad(sPathPP, sFolderRS, sFileNameRS) Then

                ' Saving in case some folder changed during SourceDataLoad
                Call SaveAppSettings                       ' Make it function?

                gbCompletedFirstLoad = False

                Call loadResTable(sFolderRS + "\" + sFileNameRS)


                Load frmResMan
                frmResMan.Show False


                Call populate_Grid(frmResMan, bynone)

                Call populate_combos(frmResMan)


                'sSql = "DROP TABLE tbl_Resources;"
                'scn.Execute sSql

                frmResMan.optSumNone.Value = True

                gbCompletedFirstLoad = True

                frmResMan.iGrid1.SetFocus
            End If
        Else
            MsgBox "Please restart Excel to complete the update process...", vbCritical, "PENS"
        End If

    End If

    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Close #giDebugFile
    Unload gFrmNavPanel
    Set gFrmNavPanel = Nothing
    Set guDictProj = Nothing
    Set gColPrjInfo = Nothing

    On Error GoTo 0

    Exit Sub

PENS_Inspector_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub











'---------------------------------------------------------------------------------------
' Method : PENS2007Settings
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS2007Settings(control As Object)             ' Do I need to pass control as a parameter?????????????????????
    Dim i As Variant
    Dim lOK2Proceed As Boolean

    Const sSOURCE As String = "PENS2007Settings"

    On Error GoTo PENS2007SETTINGS_Error

    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If

    Call RetrieveAppSettings                               'Make it function to test??

    ' Automatic check (Manual = false)
    If Not CheckAndUpdate(gbConnected2Network, False) Then

        lOK2Proceed = False

        If GetSetting("PENS", "General", "Version") = "" Then
            lOK2Proceed = True
        ElseIf CDbl(gwsConfig.Range(gsPENS_VERSION)) = CDbl(GetSetting("PENS", "General", "Version")) Then
            lOK2Proceed = True
        End If

        If lOK2Proceed Then

            If Not gbConnected2Network Then
                MsgBox "No access to the network. You can still use Portfolio Ensemble with local files..."
                '''gwsConfig.Range(gsUSE_LOCAL_DATA).Value = True

                SaveSetting "PENS", "General", "Use_Local_Folder", "True"
                gbUse_Local_Folder = True
            Else
                Call LogActivity("PENS2007Settings")       ' Make it function to test?
            End If

            ' If the user pressed the Cancel button, raise a custom
            ' user cancel error. This will cause the central error
            ' handler to exit the program without displaying an
            ' error message.
            '#########################If frmMain.UserCancel Then Err.Raise glUSER_CANCEL

            Set gFrmSettings = New frmSettings
            Load gFrmSettings

            gbInitialized = False                          ' Some activities should wait until the form is initialized
            If Not gFrmSettings.Initialize() Then          ' If UserForm initialization failed, raise a custom error.
                '    Err.Raise glCANT_INITIALIZE
            Else
                gbInitialized = True

                gFrmSettings.Show (False)                  ' Show form modeless

                DoEvents                                   ' Do I need this?????????????????????????????



                ' If the user pressed the Cancel button, raise a custom
                ' user cancel error. This will cause the central error
                ' handler to exit the program without displaying an
                ' error message.
                '#########################If frmMain.UserCancel Then Err.Raise glUSER_CANCEL


                '''''''''gFrmSettings.Show (False)
            End If

        Else
            MsgBox "Please restart Excel to complete the update process...", vbCritical, "PENS"
        End If

    End If

    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Close #giDebugFile
    ''''Unload gFrmExtPanel
    ''''Set gFrmExtPanel = Nothing

    On Error GoTo 0

    Exit Sub

PENS2007SETTINGS_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Sub
'---------------------------------------------------------------------------------------
' Method : PENS2007Feedback
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS2007Feedback(control As Object)             ' Do I need to pass control as a parameter?????????????????????
    Const sSOURCE As String = "PENS2007Feedback"

    On Error GoTo PENS2007FEEDBACK_Error

    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If





    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    strbody = "We look forward for your feedback!" & vbNewLine & vbNewLine & _
    "From new functionality to simple observations about how to improve this tool. " & _
    "Your time is appreciated and we promise to assess and respond ASAP" & vbNewLine & vbNewLine & _
    "Thank you!" & vbNewLine & _
    "The PENS team..."

    On Error Resume Next
    With OutMail
        .To = "MIISPortfolioSupport@mackenzieinvestments.com"
        .CC = ""
        .BCC = ""
        .Subject = "PENS - Feedback / Suggestions"
        .Body = strbody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
        ''''.Send        'or use .Display
        .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing





    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Close #giDebugFile

    On Error GoTo 0

    Exit Sub

PENS2007FEEDBACK_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : PENS2007Issue
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS2007Issue(control As Object)                ' Do I need to pass control as a parameter?????????????????????
    Const sSOURCE As String = "PENS2007Issue"
    Dim spath As String

    On Error GoTo PENS2007ISSUE_Error

    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If


    spath = ThisWorkbook.Path
    If Right$(spath, 1) <> "\" Then spath = spath & "\"



    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    strbody = "We are sorry about the inconvenience!" & vbNewLine & vbNewLine & _
    "Please send us an email with a description of the issue, screenshots, etc. " & _
    "An error log has been already attached to this email." & vbNewLine & vbNewLine & _
    "We'll reply to you ASAP." & vbNewLine & vbNewLine & _
    "Regards," & vbNewLine & _
    "The PENS team..." & vbNewLine & vbNewLine

    On Error Resume Next
    With OutMail
        .To = "MIISPortfolioSupport@mackenzieinvestments.com"
        .CC = ""
        .BCC = ""
        .Subject = "PENS - Reporting an issue!"
        .Body = strbody
        'You can add a file like this
        .Attachments.Add (spath + "error.log")
        ''''.Send        'or use .Display
        .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing





    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Close #giDebugFile

    On Error GoTo 0

    Exit Sub

PENS2007ISSUE_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : PENS2007RelNotes
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS2007RelNotes(control As Object)             ' Do I need to pass control as a parameter?????????????????????
    Const sSOURCE As String = "PENS2007RelNotes"
    Dim sReleaseNotes As String

    On Error GoTo PENS2007RELNOTES_Error

    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If



    Set gFrmReleaseNotes = New frmReleaseNotes
    Load gFrmReleaseNotes


    'ABS en todas las formulas
    'if no Revised BL MI labor change...bring Budget MI labour to portfolio data. Poner en la columna que lo que es diferente va en redish...

    sReleaseNotes = ""
    sReleaseNotes = sReleaseNotes & ">>> Version 3.5"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Added Portfolio Data (Short Version) with less columns and subtotals per LOB."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigation: Cost Details now showing Reporting NE."
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 3.4"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Portfolio Data - Renamed Q2 NE Explanation column to NE Explanation."
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 3.3"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Portfolio Data - Now overwriting Reporting NE with Actuals if Reporting NE is less than Actuals."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Portfolio Data - Now overwriting Revised BL MI labor with Budget MI labor if no Revised BL MI found."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Portfolio Data - Added ""MI Revised BL Labor"" and ""Q2 NE Explanation"" columns."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigation: Now freezing Months row."
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 3.2"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Portfolio Data - Added ""MI Budget Labor"" and ""Budget / CR Notes"" columns."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Added checkbox to exclude BAU initiatives by default."
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 3.1"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Fixed navigation issue related to filters and frozen panes in the Portfolio Plan tab."
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 3.0"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Added possibility to run individual reports."
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 2.9"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Status report and  Cost details popups are now resizable and font size can be adjusted."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: PENS is using the PM info in the PM column instead of the Resource info for searches."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Renamed ""NE Summary"" tab for ""Portfolio Data""."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Changed report name generated by PENS (from D4A_<Timestamp> to PENS_Rep_<Timestamp>)."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: The ""Portolio Data"" tab (former ""NE Summary"") is now sorted by LOB and Project Type."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Added comments in header cells explaning meaning of colors in the columns."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Auto-update: Fixed issue related to updates triggering over and over if users don't restart"
    sReleaseNotes = sReleaseNotes & Chr(10) & "Excel to complete an update (now PENS won't allow users to continue until they restart Excel)"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Auto-update: Added alert that users will have to restart Excel to complete the update"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Auto-update: Added suggestion to review release notes after the update is completed to learn"
    sReleaseNotes = sReleaseNotes & Chr(10) & "about new enhancements and fixes"""
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Auto-update: Fixed issue requiring users to re-enter the local folder information after every update."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Help: Completed setup of MI IS Portfolio Support mailbox for feedback and issues reported by users."
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Help: Enabled submenu to report issues (the tool is automatically attaching the error log to the email)."
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 2.8"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Added posibility to search by PM"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Fixed issue when someone types in a ComboList"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Highlight the search bar when a Search is active"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Added 2 more filters (Activation Status and Category)"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Highlight ComboList selection when there are filters in place"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: The Navigator asks if you want to open the Dashboard file after loading"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Fixed issue when no Delivery Leader is assigned in the Portfolio Plan(eg. PSERIES)"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Fixed queries to point to the new Revised BL location after Dashboard changes"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Fixed listbox refresh slowness when re-populate project list"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Help: Enabled Release Notes and About submenues"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Improved NE Summary queries performance (went down from 25 sec to 15 sec)"
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 2.7"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Fixed navigation issue when importing the Portfolio Dashboard and rows go out of sequence"
    sReleaseNotes = sReleaseNotes & Chr(10) & "  (Added a new column at the right in the Porfoltio Dashboard (Row ID) with forumla = row())"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Set fixed width for columns in Pivot Tables"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Added Search functionality (by Project Name)"
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 2.6"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Combined Proj code and Proj Name in Listbox to allow searchig by Proj Code too"
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 2.5"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Added Status details popup when clicking on the Status button"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Added Cost details popup when clicking on the Cost button"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Completed NE Summary tab (added FTE-NI fields for Neena)"
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 2.4"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Moved Capacity table to the left and Demand table to the right"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Added headers and totals in CC Capacity tab"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Changed Pivot Tables formatting to show '.' if zero"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Removed year from Pivot Table headers"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Navigator: Completed first version of navigator"
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & ">>> Version 2.3"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Completed CC Capacity table as a Pivot Table"
    sReleaseNotes = sReleaseNotes & Chr(10) & "* Reporting: Change Pivot Table resolutions to 1 decimal digit"
    sReleaseNotes = sReleaseNotes & Chr(10) & " -------------------------------------------------------------------------------------------------------------------------------------------------------"
    sReleaseNotes = sReleaseNotes & Chr(10) & "To be continued..."


    ''''''''''''- Make Implementation date color black if no info. OK (2.6)





    With gFrmReleaseNotes

        .txtReleaseNotes.Text = sReleaseNotes
        .txtReleaseNotes.SetFocus
        .txtReleaseNotes.CurLine = 0


    End With


    gFrmReleaseNotes.Show False


    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Close #giDebugFile

    On Error GoTo 0

    Exit Sub

PENS2007RELNOTES_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : PENS2007About
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS2007About(control As Object)                ' Do I need to pass control as a parameter?????????????????????
    Const sSOURCE As String = "PENS2007About"
    Dim sReleaseNotes As String

    On Error GoTo PENS2007ABOUT_Error

    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If



    Set gFrmAbout = New frmAbout
    Load gFrmAbout



    gFrmAbout.Show False


    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Close #giDebugFile

    On Error GoTo 0

    Exit Sub

PENS2007ABOUT_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub



'---------------------------------------------------------------------------------------
' Method : PENS2007Tips
' Author : cklahr
' Date   : 1/31/2016
' Purpose: Entry Point procedure invoked from the ribbon
' Arguments:
' Pending:
' Comments: Developed using Excel 2007
'---------------------------------------------------------------------------------------
Public Sub PENS2007Tips(control As Object)                 ' Do I need to pass control as a parameter?????????????????????
    Const sSOURCE As String = "PENS2007Tips"
    Dim sReleaseNotes As String

    Dim i As Long
    Dim sTips As String

    On Error GoTo PENS2007TIPS_Error

    sTips = ""
    sTips = sTips & "Navigator->Local Guide: Did you know that PENS can open the Portfolio Plan for you at any time clicking on the Local Guide image?" + ";"
    sTips = sTips & "2" + ";"
    sTips = sTips & "3"


    gsTipsArray = Split(sTips, ";")

    Call ShuffleArrayInPlace(gsTipsArray())


    Set gwsConfig = ThisWorkbook.Sheets("Configuration")

    ' Setting Debug mode based on configuration
    gbDebugMode = gwsConfig.Range(gsDEBUG_MODE).Value

    If gbDebugMode Then
        ' Open Debug file for logging purposes
        If Not bOpenDebug() Then Err.Raise glCANT_OPEN_DEBUG_FILE

        If Not bWriteDebug(msMODULE, sSOURCE, "***NEW Run start") Then Err.Raise glCANT_WRITE_DEBUG_FILE
    End If



    Set gFrmTips = New frmTips
    Load gFrmTips


    glPosTip = 0
    ''For i = 0 To UBound(gsTipsArray)
    gFrmTips.txtTip.Value = gsTipsArray(glPosTip)
    'Debug.Print gsTipsArray(i)
    ''Next




    gFrmTips.Show False


    If gbDebugMode Then Close #giDebugFile

    Exit Sub

ErrorExit:
    ' Clean up
    On Error Resume Next

    Unload gFrmTips
    Set gFrmTips = Nothing

    Close #giDebugFile

    On Error GoTo 0

    Exit Sub

PENS2007TIPS_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Sub






