Attribute VB_Name = "MStandardCode"
'---------------------------------------------------------------------------------------
' File   : MStandardCode
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Option Explicit

Const adCmdText = 1                                        ' Required for early binding

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------
Private Const msMODULE As String = "MStandardCode"

'---------------------------------------------------------------------------------------
' Method : ShuffleArrayInPlace
' Author : cklahr
' Date   : 5/3/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub ShuffleArrayInPlace(InArray() As String)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ShuffleArrayInPlace
    ' This shuffles InArray to random order, randomized in place.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim Temp As Variant
    Dim j As Long

    Randomize
    For N = LBound(InArray) To UBound(InArray)
        j = CLng(((UBound(InArray) - N) * Rnd) + N)
        If N <> j Then
            Temp = InArray(N)
            InArray(N) = InArray(j)
            InArray(j) = Temp
        End If
    Next N
End Sub

'---------------------------------------------------------------------------------------
' Method : LogActivity
' Author : cklahr
' Date   : 4/25/2016
' Purpose: Open the log file (PENSLog.xlsx on the update folder, and write out Userid, UserName, Module, and Timestamp). Then close the log file.
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub LogActivity(sModule As String)
    Dim oCn As Object
    Dim sLogText As String

    sLogText = "'" + GetUserName() + "','" + sModule + "'," + "'" + Format(Now(), "ddmmmyyyy_hhmmss") + "','" + CStr(gwsConfig.Range(gsPENS_VERSION)) + _
    "','" + CStr(gbJoin_BETA_Program) + "','" + CStr(iGridOcxRegistered()) + "'"
    '''sLogText = "'" + GetUserName() + "','" + sModule + "'," + "'" + Format(Now(), "ddmmmyyyy_hhmmss") + "','" + sModule + "'"

    On Error Resume Next

    Set oCn = CreateObject("ADODB.Connection")
    oCn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=" + gwsConfig.Range(gsUPDATE_FOLDER).Value + "\" + "PENSLog.xlsx;" & _
    "Extended Properties=Excel 12.0;"

    oCn.Execute "Insert Into [PENSLog$] Values (" + sLogText + ")"
    oCn.Close

    Set oCn = Nothing


End Sub

'---------------------------------------------------------------------------------------
' Method : bIsAllowedBetaUser
' Author : cklahr
' Date   : 10/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:Max 100 Beta users
'---------------------------------------------------------------------------------------
Function bIsAllowedBetaUser(sUserName As String) As Boolean

    Dim cn As Object
    Dim rs As Object
    Dim sSql As String
    Dim sBetaUsersFile As String
    Dim sCon As String


    '''On Error GoTo kk

    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    sBetaUsersFile = gwsConfig.Range(gsUPDATE_FOLDER) + "\BETA\bt.xlsx"

    sCon = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sBetaUsersFile _
    & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"

    cn.Open sCon

    sSql = "SELECT * FROM [Sheet1$A1:A100] WHERE [BetaUsers] =" + """" + "ALL" + """"

    rs.Open sSql, cn


    If rs.EOF Then                                         ' Not everyone is allowed

        rs.Close

        sSql = "SELECT * FROM [Sheet1$A1:A100] WHERE [BetaUsers] =" + """" + sUserName + """"

        rs.Open sSql, cn

        'Debug.Print rs.GetString

        bIsAllowedBetaUser = Not rs.EOF

    Else
        bIsAllowedBetaUser = True                          ' Everyone is allowed!
    End If

kk:

    cn.Close

    Set cn = Nothing
    Set rs = Nothing

End Function

'---------------------------------------------------------------------------------------
' Method : openDashboardFile
' Author : cklahr
' Date   : 10/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function openDashboardFile() As Workbook
    Dim wbNew As Workbook
    Dim sPathPP As String

    If gbUse_Local_Folder Then
        sPathPP = gsLocal_Folder + "\" + gsPP_Filename
    Else
        sPathPP = gsPP_Network_Folder + "\" + gsPP_Filename
    End If

    On Error Resume Next

    '' Set wbNew = Workbooks.Open(sPathPP, , True) ''''''''READONLY!!!
    Application.Calculation = xlCalculationManual

    Set wbNew = Workbooks.Open(sPathPP)

    DoEvents

    Application.Calculation = xlCalculationAutomatic


    Set openDashboardFile = wbNew

End Function


'* Returns 0 if not ok, 1 if OK, -1 if users cancels
'---------------------------------------------------------------------------------------
' Method : ValidateFileFolders
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function ValidateFileFolders(ByVal bCheckRem As Boolean, ByVal bCheckLocal As Boolean) As Integer
    Const sSOURCE As String = "ValidateFileFolders"

    Dim wsConfig As Worksheet
    Dim sFullPathPP As String
    Dim sFullPathRS As String
    Dim i As Integer

    ''''''Set wsConfig = ThisWorkbook.Sheets("Configuration")


    If (Right(gsPP_Network_Folder, 1) = "\") Then
        gsPP_Network_Folder = Left(gsPP_Network_Folder, Len(gsPP_Network_Folder) - 1)
    End If

    If (Right(gsRS_Network_Folder, 1) = "\") Then
        gsRS_Network_Folder = Left(gsRS_Network_Folder, Len(gsRS_Network_Folder) - 1)
    End If

    If (Right(gsPP_Network_Folder, 1) = "\") Then
        gsPP_Network_Folder = Left(gsPP_Network_Folder, Len(gsPP_Network_Folder) - 1)
    End If

    If (Right(gsLocal_Folder, 1) = "\") Then
        gsLocal_Folder = Left(gsLocal_Folder, Len(gsLocal_Folder) - 1)
    End If

    'Validate if remote Portfolio Plan file exists and FileName is not empty!
    If (bCheckRem) Then
        If (Not FileFolderExists(gsPP_Network_Folder + "\" + gsPP_Filename) Or Len(gsPP_Filename) = 0) Then
            i = MsgBox("Portfolio Dashboard file not found", vbCritical, "Oops...")
            ' Give user the chance to pick up a folder
            'changes the folder dialogs title
            Application.FileDialog(msoFileDialogFilePicker).Title = "Please pick the Portfolio Dashboard file"
            'the dialog is displayed to the user
            i = Application.FileDialog(msoFileDialogFilePicker).Show
            'checks if user has canceled the dialog
            If i <> 0 Then
                'dispaly message box
                sFullPathPP = FileUNC(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1))

                gsPP_Network_Folder = Left(sFullPathPP, InStrRev(sFullPathPP, "\"))
                gsPP_Filename = Right(sFullPathPP, Len(sFullPathPP) - InStrRev(sFullPathPP, "\"))
                ''''''''''            sRemPathPP = gwsConfig.Range("C4").value + "\" + txtFileNamePP.Value
                ValidateFileFolders = 0
            Else
                ValidateFileFolders = -1
            End If

            Exit Function
        End If
    End If

    'Validate if remote Resource Spreadsheet file exists
    If (bCheckRem) Then
        If (Not FileFolderExists(gsRS_Network_Folder + "\" + gsRS_Filename) Or Len(gsRS_Filename) = 0) Then
            i = MsgBox("Resource Spreadsheet file not found", vbCritical, "Oops...")
            ' Give user the chance to pick up a folder
            'changes the folder dialogs title
            Application.FileDialog(msoFileDialogFilePicker).Title = "Please pick the Resource Spreadsheet file"
            'the dialog is displayed to the user
            i = Application.FileDialog(msoFileDialogFilePicker).Show
            'checks if user has canceled the dialog
            If i <> 0 Then
                'dispaly message box
                sFullPathRS = FileUNC(Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1))

                gsRS_Network_Folder = Left(sFullPathRS, InStrRev(sFullPathRS, "\"))
                gsRS_Filename = Right(sFullPathRS, Len(sFullPathRS) - InStrRev(sFullPathRS, "\"))

                '''''''''''''''            sRemPathRS = txtRemFolderRS.Value + "\" + txtFileNameRS.Value
                ValidateFileFolders = 0
            Else
                ValidateFileFolders = -1
            End If

            Exit Function
        End If
    End If

    'Validate if local folder exists
    If (bCheckLocal) Then
        If (Not pathExists(gsLocal_Folder)) Then
            i = MsgBox("Local folder not found!" & vbNewLine & vbNewLine & _
            "Please choose a local folder to store reports and cached data...", vbCritical, "PENS")
            ' Give user the chance to pick up a folder
            'changes the folder dialogs title
            Application.FileDialog(msoFileDialogFolderPicker).Title = "Please pick a local folder to store the reports"
            'the dialog is displayed to the user
            i = Application.FileDialog(msoFileDialogFolderPicker).Show
            'checks if user has canceled the dialog
            If i <> 0 Then
                'dispaly message box
                gsLocal_Folder = FileUNC(Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1))
                ValidateFileFolders = 0
            Else
                ValidateFileFolders = -1
            End If

            Exit Function
        End If
    End If


    ValidateFileFolders = 1                                ' Tudo OK!

End Function

'---------------------------------------------------------------------------------------
' Method : FillFirst2RowsWithZeroesPP
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub FillFirst2RowsWithZeroesPP()




End Sub

'---------------------------------------------------------------------------------------
' Method : ad
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub ad()
    Dim iOutput As Integer

    On Error Resume Next

    iOutput = MsgBox("Tired of waiting for Pivot tables to refresh?", vbYesNo, "Please answer the question...")

    If iOutput = vbYes Then
        iOutput = MsgBox("Tired of formatting issues in your Pivot tables?", vbYesNo, "Please answer the question...")
        If iOutput = vbYes Then
            iOutput = MsgBox("Tired of discrepancies among your Pivot tables?", vbYesNo, "Please answer the question...")
            If iOutput = vbYes Then
                iOutput = MsgBox("Tired of THIS???", vbYesNo, "Please answer the question...")
            Else
                iOutput = MsgBox("OK. Good luck then!", vbOKOnly, "It's your choice...")
                Unload Me
                End
            End If
        Else
            iOutput = MsgBox("OK. Good luck then!", vbOKOnly, "It's your choice...")
            Unload Me
            End
        End If
    Else
        iOutput = MsgBox("OK. Good luck then!", vbOKOnly, "It's your choice...")
        Unload Me
        End
    End If

    frmTired.Show

    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Method : bOpenDebug
' Author : cklahr
' Date   : 2/2/2016
' Purpose: Open Debug.log for debugging purposes
' Arguments:
' Pending:
' Comments: Opens debug.log for append in the same directory where this file resides
'---------------------------------------------------------------------------------------
Function bOpenDebug() As Boolean
    Dim spath As String

    ' Get the application directory.
    spath = ThisWorkbook.Path
    If Right$(spath, 1) <> "\" Then spath = spath & "\"

    ' Open the log file for debugging
    giDebugFile = FreeFile()

    On Error Resume Next

    Open spath & gsFILE_DEBUG_LOG For Append As #giDebugFile

    bOpenDebug = (Err.Number = 0)

End Function

'---------------------------------------------------------------------------------------
' Method : bWriteDebug
' Author : cklahr
' Date   : 2/2/2016
' Purpose: Writes a string to the Debug file and return True if succesful
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function bWriteDebug(ByVal sModule As String, _
    ByVal sProc As String, _
    ByVal sLogData As String, _
    Optional ByVal sFile As String) As Boolean

    Dim sFullSource As String
    Dim sLogText As String

    If Len(sFile) = 0 Then sFile = ThisWorkbook.Name

    ' Construct the fully-qualified error source name.
    sFullSource = "[" & sFile & "]" & sModule & "." & sProc

    ' Create the error text to be logged.
    sLogText = "  " & sFullSource & ", DEBUG: " & sLogData

    On Error Resume Next
    Print #giDebugFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText

    bWriteDebug = (Err.Number = 0)

End Function

'---------------------------------------------------------------------------------------
' Method : RefreshTemplates
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments: Templates are copied in the local PENS folder
'---------------------------------------------------------------------------------------
Public Function RefreshTemplates() As Boolean
    Const sSOURCE As String = "RefreshTemplates"

    Dim sLocalPathTemplate As String
    Dim sRemPathTemplate As String

    Dim i As Integer
    Dim bb As Boolean

    On Error GoTo RefreshTemplates_Error

    '    i = 0
    '    Do While (i = 0)
    '        i = ValidateFileFolders(True, True)
    '    Loop
    '
    '    If i = -1 Then
    '        RefreshTemplates = False
    '        Exit Function
    '    End If


    sLocalPathTemplate = gsLocal_Folder + "\" + gsCCC_Filename

    If gbConnected2Network Then
        If gbJoin_BETA_Program Then
            sRemPathTemplate = gwsConfig.Range(gsUPDATE_FOLDER) + "\BETA" + "\" + gsCCC_Filename
        Else
            sRemPathTemplate = gwsConfig.Range(gsUPDATE_FOLDER) + "\" + gsCCC_Filename
        End If


        'If i > 0 Then
        If FileLastModified(sLocalPathTemplate) <> FileLastModified(sRemPathTemplate) Then
            If Not VBCopyFileFolder(sRemPathTemplate, sLocalPathTemplate) Then
                RefreshTemplates = False
                Exit Function
            End If
        End If


        '    Else
        '        Exit Function
        '    End If

        ' Need to fill the first 2 rows of the PP file with zeroes
        ''''''Call FillFirst2RowsWithZeroesPP

        '    If (gwsConfig.Range(gsINFORM_LOCAL_COPY)) Then
        '        MsgBox "Local copy completed succesfully", vbOKOnly, "CONGRATS!"
        '    End If

        ' Save configuration after succesfully accessed and copied the data!
        '''ThisWorkbook.Save
        Call SaveAppSettings                                   'Make it function to test? 'Should I save here???

        RefreshTemplates = True
    Else
        ' Confirm there is a local template already
        RefreshTemplates = (FileLastModified(sLocalPathTemplate) <> "N/A")
    End If

ErrorExit:
    ' Clean up
    Exit Function

RefreshTemplates_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Function

'---------------------------------------------------------------------------------------
' Method : RefreshLocalSources
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function RefreshLocalSources() As Boolean
    Const sSOURCE As String = "RefreshLocalSources"

    Dim sLocalPathPP As String
    Dim sRemPathPP As String
    Dim sLocalPathRS As String
    Dim sRemPathRS As String

    Dim i As Integer
    Dim bb As Boolean

    On Error GoTo RefreshLocalSources_Error

    i = 0
    Do While (i = 0)
        i = ValidateFileFolders(True, True)
    Loop

    If i = -1 Then
        RefreshLocalSources = False
        Exit Function
    End If


    sLocalPathPP = gsLocal_Folder + "\" + gsPP_Filename
    sRemPathPP = gsPP_Network_Folder + "\" + gsPP_Filename

    sLocalPathRS = gsLocal_Folder + "\" + gsRS_Filename
    sRemPathRS = gsRS_Network_Folder + "\" + gsRS_Filename


    If i > 0 Then
        If FileLastModified(sLocalPathPP) <> FileLastModified(sRemPathPP) Then
            If Not VBCopyFileFolder(sRemPathPP, sLocalPathPP) Then
                RefreshLocalSources = False
                Exit Function
            End If
        End If

        If FileLastModified(sLocalPathRS) <> FileLastModified(sRemPathRS) Then
            If Not VBCopyFileFolder(sRemPathRS, sLocalPathRS) Then
                RefreshLocalSources = False
                Exit Function
            End If
        End If
    Else
        Exit Function
    End If

    ' Need to fill the first 2 rows of the PP file with zeroes
    ''''''Call FillFirst2RowsWithZeroesPP

    If (gwsConfig.Range(gsINFORM_LOCAL_COPY)) Then
        MsgBox "Local copy completed succesfully", vbOKOnly, "CONGRATS!"
    End If

    ' Save configuration after succesfully accessed and copied the data!
    '''ThisWorkbook.Save
    Call SaveAppSettings                                   'Make it function to test?

    RefreshLocalSources = True

ErrorExit:
    ' Clean up
    Exit Function

RefreshLocalSources_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Function

'---------------------------------------------------------------------------------------
' Method : SourceDataLoad
' Author : cklahr
' Date   : 3/19/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function SourceDataLoad(ByRef sPathPP As String, ByRef sFolderRS As String, ByRef sFileNameRS As String) As Boolean
    Dim bb As Boolean
    Dim sLastUpdatePP As String
    Dim sLastUpdateRS As String

    Const sSOURCE As String = "SourceDataLoad"

    SourceDataLoad = False

    If (gbUse_Local_Folder) Then
        sLastUpdatePP = FileLastModified(gsLocal_Folder + "\" + gsPP_Filename)

        If sLastUpdatePP = "N/A" Then                      ' Can't find a local file to use!
            MsgBox "No local Data Sources found. Please change your settings. Click OK to continue..."
            Exit Function
        End If

        sLastUpdateRS = FileLastModified(gsLocal_Folder + "\" + gsRS_Filename)

        If sLastUpdateRS = "N/A" Then
            MsgBox "No local Data Sources found. Please change your settings. Click OK to continue..."
            Exit Function
        End If

        If (MsgBox("Are you OK to use local files for this run?" + vbNewLine + vbNewLine + _
            "(if not, please change the Settings to use Live Data)", vbYesNo, "Please confirm...") = vbYes) Then
            ' Include PP and RS timestamps in this message!!!!!!!!!!!!!!
            sPathPP = gsLocal_Folder + "\" + gsPP_Filename
            sFolderRS = gsLocal_Folder
            sFileNameRS = gsRS_Filename
        Else
            Exit Function
        End If
    Else
        'Will refresh local sources
        sPathPP = gsPP_Network_Folder + "\" + gsPP_Filename
        sFolderRS = gsRS_Filename
        sFileNameRS = gsRS_Filename

        If Not RefreshLocalSources Then                    ' MAKE IT A FUNCTION TO TEST!!!!
            SourceDataLoad = False
            Exit Function
        End If
    End If

    '============= I'm using local Data Sources ALWAYS!!!!!========================
    sPathPP = gsLocal_Folder + "\" + gsPP_Filename
    sFolderRS = gsLocal_Folder
    sFileNameRS = gsRS_Filename
    '==================================================================================================

    SourceDataLoad = True

End Function

'---------------------------------------------------------------------------------------
' Method : FetchValue
' Author : cklahr
' Date   : 3/9/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function FetchValue(cn As Object, sRetField As String, sCondField1 As String, sCondition1 As String, sCondField2 As String, sCondition2 As String) As String
    Dim cmdCommand As Object
    Dim rstRecordset As Object

    '******************
    '''On Error GoTo kk
    '******************

    Set cmdCommand = CreateObject("ADODB.Command")

    Set cmdCommand.activeconnection = cn
    With cmdCommand

        If sRetField = "Report" Then
            .CommandText = "SELECT [" + sRetField + "] " + "FROM tbl_PortfolioPlan WHERE [" + sCondField1 + "] = " + sCondition1 + " AND [" + sCondField2 + "] = " + "'" + sCondition2 + "'"
        Else
            .CommandText = "SELECT [" + sRetField + "] " + "FROM tbl_PortfolioPlan WHERE [" + sCondField1 + "] = " + "'" + sCondition1 + "'" + " AND [" + sCondField2 + "] = " + "'" + sCondition2 + "'"
        End If
        .CommandType = adCmdText
        .Execute
    End With

    ' Open the recordset.

    Set rstRecordset = CreateObject("ADODB.Recordset")
    Set rstRecordset.activeconnection = cn
    rstRecordset.Open cmdCommand

    Dim a As String

    If rstRecordset.EOF Then                               ' Didn't find a thing!
        a = ""
    Else
        If (rstRecordset.Fields(sRetField) <> vbNull) Then
            a = rstRecordset.Fields(sRetField)
        Else
            a = ""
        End If
    End If

    FetchValue = a

    Set cmdCommand = Nothing
    Set rstRecordset = Nothing

    Exit Function
kk:
    ''''a = ""
End Function

'---------------------------------------------------------------------------------------
' Method : RetrieveAppSettings
' Author : cklahr
' Date   : 2/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub RetrieveAppSettings()
    Dim vaKeys As Variant

    vaKeys = GetAllSettings("PENS", "General")


    ' Adding a new key Join_BETA_Program!!
    If (GetSetting("PENS", "General", "Join_BETA_Program") = "") Then
        SaveSetting "PENS", "General", "Join_BETA_Program", False
    End If


    If Not IsArray(vaKeys) Then                            ' Settings not in Registry
        'Get initial settings from the spreadsheet and save it in the registry
        SaveSetting "PENS", "General", "PP_Filename", gwsConfig.Range(gsFILENAME_PP).Value
        SaveSetting "PENS", "General", "PP_Network_Folder", gwsConfig.Range(gsREM_FOLDER_PP).Value
        SaveSetting "PENS", "General", "RS_Filename", gwsConfig.Range(gsFILENAME_RS).Value
        SaveSetting "PENS", "General", "RS_Network_Folder", gwsConfig.Range(gsREM_FOLDER_RS).Value
        SaveSetting "PENS", "General", "Local_Folder", gwsConfig.Range(gsLOCAL_FOLDERS).Value
        SaveSetting "PENS", "General", "Use_Local_Folder", gwsConfig.Range(gsUSE_LOCAL_DATA).Value
        SaveSetting "PENS", "General", "Join_BETA_Program", gbJoin_BETA_Program
    End If

    ' Load global variables
    gsPP_Filename = GetSetting("PENS", "General", "PP_Filename")
    gsPP_Network_Folder = GetSetting("PENS", "General", "PP_Network_Folder")
    gsRS_Filename = GetSetting("PENS", "General", "RS_Filename")
    gsRS_Network_Folder = GetSetting("PENS", "General", "RS_Network_Folder")
    gsLocal_Folder = GetSetting("PENS", "General", "Local_Folder")
    gbUse_Local_Folder = GetSetting("PENS", "General", "Use_Local_Folder")
    gbJoin_BETA_Program = GetSetting("PENS", "General", "Join_BETA_Program")

End Sub


'---------------------------------------------------------------------------------------
' Method : SaveAppSettings
' Author : cklahr
' Date   : 2/18/2016
' Purpose:
' Arguments:
' Pending: ERROR HANDLING!!!!!!!
' Comments:
'---------------------------------------------------------------------------------------
Public Sub SaveAppSettings()

    'Save variable changes in the registry
    SaveSetting "PENS", "General", "PP_Filename", gsPP_Filename
    SaveSetting "PENS", "General", "PP_Network_Folder", gsPP_Network_Folder

    SaveSetting "PENS", "General", "RS_Filename", gsRS_Filename
    SaveSetting "PENS", "General", "RS_Network_Folder", gsRS_Network_Folder

    SaveSetting "PENS", "General", "Local_Folder", gsLocal_Folder
    SaveSetting "PENS", "General", "Use_Local_Folder", gbUse_Local_Folder
    SaveSetting "PENS", "General", "Join_BETA_Program", gbJoin_BETA_Program

End Sub

Sub DeleteAllSettings(ByRef vaSettings As Variant)
    Dim nItem As Integer

    If IsArray(vaSettings) Then
        For nItem = 0 To UBound(vaSettings)
            Debug.Print vaSettings(nItem, 0) & ": " & _
            vaSettings(nItem, 1)
            DeleteSetting "PENS", "General", vaSettings(nItem, 0)
        Next
    End If
End Sub


'********************************************************************************************************************************************************************************
'Function Name                     : IsDashboardOpen(ByVal sPathPP As String)
'Function Description             : Function to check whether specified workbook is open
'Data Parameters                  : sPathPP:- Specify name or path to the workbook. eg: "Nucleation.xlsx" or "C:\Users\Kannan.S\Desktop\Nucleation\Nucleation.xlsm"
'Created by                           : Kannan S
'Email                                   : info@nucleation.in
'Creation date                       : 13-Nov-2013
'Website                               : www.nucleation.in
'THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
'LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
'Feel free to use the code as you wish but kindly keep this header section intact.
'Copyright © 2013 Nucleation. All Rights Reserved.
'********************************************************************************************************************************************************************************
Function IsDashboardOpen() As Boolean
    Dim WB As Excel.Workbook
    Dim WBName As String
    Dim WBPath As String
    Dim sPathPP As String
    Dim sPathPPArray() As String


    IsDashboardOpen = False

    If gbUse_Local_Folder Then
        sPathPP = gsLocal_Folder + "\" + gsPP_Filename
    Else
        sPathPP = gsPP_Network_Folder + "\" + gsPP_Filename
    End If


    Err.Clear
    On Error Resume Next
    sPathPPArray = Split(sPathPP, "\")
    Set WB = Application.Workbooks(sPathPPArray(UBound(sPathPPArray)))
    WBName = sPathPPArray(UBound(sPathPPArray))
    WBPath = WB.Path & "\" & WBName
    If Not WB Is Nothing Then
        If UBound(sPathPPArray) > 0 Then
            If LCase(WBPath) = LCase(sPathPP) Then
                IsDashboardOpen = True
                WB.Activate
            End If
        Else
            IsDashboardOpen = True
            WB.Activate
        End If
    End If
    Err.Clear
End Function
