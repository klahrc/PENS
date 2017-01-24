Attribute VB_Name = "modMain"
'-------------------------------------------------------------------------
' Module    : modMain
' Company   : JKP Application Development Services (c) 2005
' Author    : Jan Karel Pieterse, jkpieterse@jkp-ads.com
' Created   : 24-2-2005
' Purpose   : Main module
'
' You may use this module for your own applications,
' but please keep this header intact.
'-------------------------------------------------------------------------

Option Explicit

Dim msAddInLibPath As String        'Holds full path and filename where addin is going to be stored
Dim msCurrentAddInPath As String        'Holds current path and filename where addin is located

Public gbVarsOK As Boolean        'True when variables have been initialised
Public gsPath As String        'Path to place addin in (if empty, librarypath will be used)
Public gsAppName As String        'Name of Application
Public gsFilename As String        'Name of Addin file
Public gsRegKey As String        'RegKey for settings

Public Const gsOCXFilename As String = "iGrid500_10Tec.ocx"

'---------------------------------------------------------------------------------------
' Method : SomeThingWrong
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub SomeThingWrong()
    If Application.OperatingSystem Like "*Win*" Then
        MsgBox prompt:="Something went wrong during copying" & vbNewLine _
                       & "of the add-in to the add-in directory:" _
                       & vbNewLine & vbNewLine & gsPath _
                       & vbNewLine & vbNewLine & "You can install " & gsAppName & " manually by copying the file" _
                       & vbNewLine & gsFilename & " to this directory yourself and registering the addin" _
                       & vbNewLine & "using Tools, Addins from the menu of Excel." _
                       & vbNewLine & vbNewLine & "Don't press OK yet, first do the copying from Windows Explorer." _
                       & vbNewLine & "It gives you the opportunity to ALT-TAB back to Excel" _
                       & vbNewLine & "to read this text.", Buttons:=vbOKOnly, Title:=gsAppName & " Install / Repair"
    Else
        MsgBox prompt:="Something went wrong during copying" & vbNewLine _
                       & "of the add-in to the add-in directory:" _
                       & vbNewLine & vbNewLine & gsPath _
                       & vbNewLine & vbNewLine & "You can install " & gsAppName & " manually by copying the file" _
                       & vbNewLine & gsFilename & " to this directory yourself and installing the addin" _
                       & vbNewLine & "using Tools, Addins from the menu of Excel." _
                       & vbNewLine & vbNewLine & "Don't press OK yet, first do the copying in the Finder." _
                       & vbNewLine & "It gives you the opportunity to Command-TAB back to Excel" _
                       & vbNewLine & "to read this text.", Buttons:=vbOKOnly, Title:=gsAppName & " Install / Repair"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : Initialise
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub Initialise()
    ' Loading configuration from hidden tab (see Named ranges for details)
    gsAppName = ThisWorkbook.Names("AppName").RefersToRange.Value
    gsFilename = ThisWorkbook.Names("FileName").RefersToRange.Value
    gsRegKey = ThisWorkbook.Names("RegKey").RefersToRange.Value
    gsPath = ThisWorkbook.Names("Path").RefersToRange.Value

    'Check if an installation path has been specified
    'If not, use the default all user library path
    If gsPath = "" Then
        gsPath = Application.LibraryPath
    End If

    If gsPath = "All User Addin Library" Then
        gsPath = Application.LibraryPath
    ElseIf gsPath = "Current User Addin Path" Then
        gsPath = Application.UserLibraryPath
    End If

    'Add trailing path separator if needed
    If Right(gsPath, 1) <> Application.PathSeparator Then
        gsPath = gsPath & Application.PathSeparator
    End If

    ThisWorkbook.Worksheets(1).Unprotect
    ThisWorkbook.Worksheets(1).Buttons(1).Caption = "Install / Repair " & gsAppName
    ThisWorkbook.Worksheets(1).Buttons(2).Caption = "Uninstall " & gsAppName
    ThisWorkbook.Worksheets(1).Protect userinterfaceonly:=True
    gbVarsOK = True
End Sub

'---------------------------------------------------------------------------------------
' Method : SaveAndClose
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub SaveAndClose()
    'Save this setup file ensuring the buttons on the front sheet show the macro enable caption
    ThisWorkbook.Worksheets(1).Buttons(1).Caption = "Please enable macro's to make this button work."
    ThisWorkbook.Worksheets(1).Buttons(2).Caption = "Please enable macro's to make this button work."
    ThisWorkbook.Worksheets(1).Activate
    ThisWorkbook.Save
    ThisWorkbook.Close
End Sub

'---------------------------------------------------------------------------------------
' Method : PathExists
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function PathExists(ByVal sPath As String) As Boolean
    Dim sFound As String
    Dim bExists As Boolean
    'Check whether the path sPath exists
    If Right(sPath, 1) = Application.PathSeparator Then
        sPath = Left(sPath, Len(sPath) - 1)        ' & Application.PathSeparator
    End If

    bExists = False

    sFound = Dir(sPath, vbDirectory)
    If sFound <> "" Then
        If (GetAttr(sPath) And vbDirectory) Then
            bExists = True
        End If
    End If
TidyUp:
    PathExists = bExists
End Function

'---------------------------------------------------------------------------------------
' Method : AddPath
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub AddPath(ByVal sPath As String)
    Dim bExists As Boolean
    Dim sTemp As String
    Dim sFound As String
    Dim iPos As Integer
    Dim sCurdir As String
    On Error GoTo LocErr
    sCurdir = CurDir
    If PathExists(sPath) = False Then
        iPos = 3
        If Right(sPath, 1) <> Application.PathSeparator Then
            sPath = sPath & Application.PathSeparator
        End If

        'Build the entire path, checking existence of each (sub)folder
        While iPos > 0

            iPos = InStr(iPos + 1, sPath, Application.PathSeparator)
            sTemp = Left(sPath, iPos)
            If sTemp = "" Then GoTo TidyUp
            If PathExists(sTemp) = False Then
                MkDir sTemp
            Else
                ChDir sTemp
            End If
        Wend
    End If
TidyUp:
    If sCurdir <> CurDir Then
        ChDrive sCurdir
        ChDir sCurdir
    End If
    Exit Sub
LocErr:
    Stop
    If Err.Number = 75 Then
        MsgBox "Path creation failed!!", vbCritical + vbOKOnly, gsAppName
        Resume TidyUp
    Else
        MsgBox "Unexpected error: Error " & Err.Number & vbNewLine & _
               Err.Description, vbCritical + vbOKOnly, gsAppName
        Resume TidyUp
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : testme
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub testme()
    If PathExists("c:\data\temp\test\test") = False Then
        AddPath "c:\data\temp\test\test"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : ClearAddinRegister
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub ClearAddinRegister()
    'Courtesy: Richard Reye
    Dim lCount As Long
    Dim sGoUpandDown As String

    'Turn display alerts off so user is not prompted to remove Addin from list
    Application.DisplayAlerts = False

    Do
        'Get Count of all AddIns
        lCount = Application.AddIns.Count

        'Create string for SendKeys that will move up & down AddIn Manager List
        'Any invalid AddIn listed will be removed
        sGoUpandDown = "{Up " & lCount & "}{DOWN " & lCount & "}"

        Application.SendKeys sGoUpandDown & "~", False
        Application.Dialogs(xlDialogAddinManager).Show

        'Continue process until all invalid AddIns are removed since
        'this code can only remove one at a time.
    Loop While lCount <> Application.AddIns.Count

    Application.DisplayAlerts = True
End Sub

'---------------------------------------------------------------------------------------
' Method : Setup
' Author : cklahr
' Date   : 12/17/2016
' Purpose: Install / Repair PENS - Invoked from button in "Sheet1"
' Arguments:
' Pending:
' Comments: Install / Repair
'           V2 introduces the option to reinstall without having to uninstall first and the install of iGrid500.OCX
'           If addin already exists, reinstall overwriting existing PENS add-in and delete registry entries related to PENS configuration (as if it was a fresh install)
'           If iGrid500.ocx not present, it will copy and register the OCX
'---------------------------------------------------------------------------------------
Sub Setup()
    Dim sOldAddIn As String

    If Not gbVarsOK Then Initialise

    'Ask for confirmation
    If MsgBox("This will install / Repair " & gsAppName _
              & vbNewLine & vbNewLine & "to:" & vbNewLine _
              & vbNewLine & "'" & gsPath & "'" & vbNewLine _
              & vbNewLine & vbNewLine & "Proceed?", vbYesNo, gsAppName & " Install / Repair") = vbYes Then

        On Error Resume Next

        msCurrentAddInPath = ThisWorkbook.Path

        'Check for addin in remote folder (looking for PENS.xlam on the same folder)
        If Dir(msCurrentAddInPath & Application.PathSeparator & gsFilename) = "" Then
            MsgBox "The file '" & gsFilename & "' appears to be missing from this folder:" _
                   & vbNewLine & msCurrentAddInPath & vbNewLine & _
                   "Please copy the PENS add-in to this folder and retry." _
                   , vbCritical + vbOKOnly, gsAppName & " Install / Repair Error"
            Exit Sub
        End If

        Dim bRestart As Boolean
        bRestart = False

        'Close older copy of the addin if it exists
        sOldAddIn = Workbooks(gsFilename).FullName
        Workbooks(gsFilename).Close False
        If sOldAddIn <> "" Then        'Re-install or repair
            Kill sOldAddIn
            bRestart = True
        End If

        'Check if the install path exists, create if not, cancel setup if fails
        If PathExists(gsPath) = False Then
            If MsgBox("The path '" & gsPath & "' does not exist, create it?", vbQuestion + vbYesNo, gsAppName) = vbYes Then
                AddPath (gsPath)
            End If
        End If

        If PathExists(gsPath) = False Then
            MsgBox "Creating of the install path:" & vbNewLine & _
                   "'" & gsPath & "'" & vbNewLine & "has failed or was cancelled, Install / Repair cancelled.", _
                   vbCritical + vbOKOnly, "Install / Repair " & gsAppName & ", error"
            Exit Sub
        End If

'*************************************************************************************************************************************
'*************************************************iGdrid500 section - UNCOMMENT FOR NEXT RELEASE**************************************
'*************************************************************************************************************************************
'        ' Install iGrid500 OCX if not installed already
'        If Not iGridOcxRegistered() Then
'            ' Copy iGrid500 OCX to Addin folder
'            FileCopy msCurrentAddInPath & Application.PathSeparator & gsOCXFilename, gsPath & gsOCXFilename
'            '''''''''''''FileCopy msCurrentAddInPath & Application.PathSeparator & gsOCXFilename, "C:\RefKK\" & gsOCXFilename
'
'            ' Register from Addin folder
'            Call RegisterFile(gsPath & gsOCXFilename)
'            ''''''''''''''Call RegisterFile("C:\RefKK\" & gsOCXFilename)
'        End If
'
'        ' Add Reference to VBE for iGrid500.ocx
'        Call AddReference
'*************************************************************************************************************************************

        Err.Clear
        'Copy addin to install path

        FileCopy msCurrentAddInPath & Application.PathSeparator & gsFilename, gsPath & gsFilename
        If Err.Number <> 0 Then
            SomeThingWrong
            Exit Sub
        End If

        'Now add the addin to the addins list and install the addin
        With AddIns.Add(FileName:=gsPath & gsFilename)
            .Installed = True
        End With

        ' If it's a repair, delete version from Registry as if it was a fresh install!
        If GetSetting(gsRegKey, "General", "Version") <> "" Then        'gsRegKey should be PENS in the configuration tab (hidden)
            DeleteSetting gsRegKey, "General", "Version"
        End If

        'No errors, all seems well.
        If Err.Number = 0 Then
            If bRestart Then
                MsgBox "Successfully installed " & gsAppName & "." & _
                       vbNewLine & vbNewLine & "You must restart Excel now...", vbInformation + vbOKOnly, _
                       "Install / Repair " & gsAppName
            Else
                MsgBox "Successfully installed " & gsAppName & "." & _
                       vbNewLine & "You can close this file.", vbInformation + vbOKOnly, _
                       "Install / Repair " & gsAppName
            End If
        Else
            SomeThingWrong
        End If

    Else
        MsgBox "Install Cancelled", vbInformation + vbOKOnly, gsAppName & " Install / Repair"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Method : Uninstall
' Author : cklahr
' Date   : 12/17/2016
' Purpose: Uninstall PENS - Invoked from button in "Sheet1"
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub Uninstall()
    Dim lReply As Long

    If Not gbVarsOK Then Initialise

    'Confirm removal
    lReply = MsgBox("This will remove " & gsAppName & vbNewLine & _
                    "from your system." & vbNewLine & vbNewLine & "Proceed?", vbYesNo, gsAppName & " Uninstall")
    If lReply = vbYes Then

        'Check if an installation path has been specified
        'If not, use the default all user library path
        If gsPath = "" Then
            gsPath = Application.LibraryPath & Application.PathSeparator
        End If

        On Error Resume Next

        'Close addin
        Workbooks(gsFilename).Close False

        'Delete file
        On Error Resume Next
        Kill gsPath & gsFilename

        'If a registry key has been specified, remove it
        If Not gsRegKey = "" Then
            DeleteSetting gsRegKey
        End If

        ' Unregister addin from Excel
        ClearAddinRegister

'*************************************************************************************************************************************
'*************************************************iGdrid500 section - UNCOMMENT FOR NEXT RELEASE**************************************
'*************************************************************************************************************************************
'        ' Delete Reference from VBE for iGrid500.ocx
'        Call DeleteReference
'
'        ' UnRegister OCX from Addin folder
'        Call UnRegisterFile(gsPath & gsOCXFilename)
'*************************************************************************************************************************************

        ''' Kill gsPath & gsOCXFilename (It won't work because the OCX is open by Excel!?)

        MsgBox gsAppName & " has been removed from your computer."
    End If
End Sub

