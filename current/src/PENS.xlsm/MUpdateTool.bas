Attribute VB_Name = "MUpdateTool"
'---------------------------------------------------------------------------------------
' File   : MUpdateTool
' Author : cklahr
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------
Private Const msMODULE As String = "MUpdateTool"

Dim mcUpdate As clsUpdate

'---------------------------------------------------------------------------------------
' Method : ManualUpdate
' Author : cklahr
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub ManualUpdate()
    On Error Resume Next
    Application.OnTime Now, "'" & ThisWorkbook.FullName & "'!CheckAndUpdate"

    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Method : CheckAndUpdate
' Author : cklahr
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments: Return true if an update is completed, false otherwise
'---------------------------------------------------------------------------------------
Public Function CheckAndUpdate(ByRef bC2N As Boolean, Optional bManual As Boolean = True) As Boolean
    Const sSOURCE As String = "CheckAndUpdate"
    Dim sError As String
    Dim sUpdLocation As String


    On Error GoTo CheckAndUpdate_Error

    ' Assumption is that there won't be update
    CheckAndUpdate = False

    ' Connected to Network?
    bC2N = False


    Set mcUpdate = New clsUpdate
    If bManual Then
        Application.Cursor = xlWait
    End If

    With mcUpdate
        'Set intial values of class
        'Current build extracted from worksheet (see named ranges)
        .Build = CDbl(gwsConfig.Range(gsPENS_VERSION))
        'Name of this app, probably a global variable, such as appname
        .AppName = "PENS"
        'Get rid of possible old backup copy
        .RemoveOldCopy
        'URL which contains build # of new version

        'Started check automatically or manually?
        .Manual = bManual


        ' Try first to install Beta versions if part of the program and it's an allowed user
        If gbJoin_BETA_Program And bIsAllowedBetaUser(GetUserName()) Then
            sUpdLocation = gwsConfig.Range(gsUPDATE_FOLDER) + "\BETA"

            sUpdLocation = Replace(sUpdLocation, "\", "/")
            .CheckURL = "file:" + sUpdLocation + "/" + "PENS.htm"
            .DownloadName = "file:" + sUpdLocation + "/" + ThisWorkbook.Name

            ' Web example if needed...
            '.DownloadName = "http://www.jkp-ads.com/downloadscript.asp?filename=" & ThisWorkbook.Name
            '.CheckURL = "http://www.jkp-ads.com/Updateanaddinbuild.htm"

            If .PossibleUpdate(sError) Then
                bC2N = True
                .Connected2Network = True

                
                '*************************************************************************************************************************************
                '
                ' Call registerIgridOCX
                '
                '*************************************************************************************************************************************
                
                CheckAndUpdate = .DoUpdate(True)
            End If
        End If

        ' Try always with Release version (if there is a newer version then update!)
        sUpdLocation = gwsConfig.Range(gsUPDATE_FOLDER)
        sUpdLocation = Replace(sUpdLocation, "\", "/")
        .CheckURL = "file:" + sUpdLocation + "/" + "PENS.htm"
        .DownloadName = "file:" + sUpdLocation + "/" + ThisWorkbook.Name


        If Not .PossibleUpdate(sError) Then
            .Connected2Network = False
            If .Manual Then
                MsgBox "Error fetching version information (Possible network problem): " & sError, vbExclamation + vbOKOnly, .AppName
            Else
                '########### RAISE PROGRAMATIC ERROR!!!
            End If
        Else
            bC2N = True
            .Connected2Network = True

            '*************************************************************************************************************************************
            ' Call registerIgridOCX
            '*************************************************************************************************************************************
            
            CheckAndUpdate = .DoUpdate(False)
        End If

        Application.Cursor = xlDefault
        Set mcUpdate = Nothing
    End With

    Exit Function

ErrorExit:
    Application.Cursor = xlDefault
    CheckAndUpdate = False
    Set mcUpdate = Nothing

    On Error GoTo 0

    Exit Function

CheckAndUpdate_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function
