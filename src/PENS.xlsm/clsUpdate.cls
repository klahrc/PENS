VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-------------------------------------------------------------------------
' Module    : clsUpdate
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse / Adapted by Cesar Klahr (2-Feb-2016)
' Created   : 19-2-2007
' Purpose   : Class to check for program updates
'-------------------------------------------------------------------------
Option Explicit


Private Declare Function URLDownloadToFile Lib "urlmon" _
Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
ByVal szURL As String, ByVal szFileName As String, _
ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long


Private msAppName As String
Private msBuild As String
Private msCheckURL As String
Private msCurrentAddinName As String
Private msDownloadName As String
Private msTempAddInName As String
Private mbManual As Boolean
Private msNewBuild As String
Private mbConnected2Network As Boolean
'---------------------------------------------------------------------------------------
' Method : DownloadFile
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub DownloadFile(strWebFilename As String, strSaveFileName As String)
    On Error Resume Next

    ' Download the file.
    URLDownloadToFile 0, strWebFilename, strSaveFileName, 0, 0

    On Error GoTo 0
End Sub
'---------------------------------------------------------------------------------------
' Method : GetUpdate
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function GetUpdate() As Boolean

    On Error Resume Next

    'If workbook has been saved readonly, we can safely delete the file!
    If ThisWorkbook.ReadOnly Then
        Err.Clear
        Kill CurrentAddinName
    End If

    ThisWorkbook.SaveAs TempAddInName

    DoEvents
    Kill CurrentAddinName

    On Error GoTo 0

    DownloadFile DownloadName, CurrentAddinName
    '''''If Err = 0 Then GetUpdate = True

    GetUpdate = (Err = 0)

End Function
'-------------------------------------------------------------------------
' Procedure : PossibleUpdate
' Company   : JKP Application Development Services (c)
' Author    : Jan Karel Pieterse (www.jkp-ads.com)
' Created   : 2-6-2009
' Purpose   : Fetches version information from a webpage
'-------------------------------------------------------------------------
Public Function PossibleUpdate(ByRef sError As String) As Boolean
    Dim oHTTP As Object
    Dim sURL As String
    Dim lError As Long
    Dim sTextResponse As String

    On Error Resume Next


    sURL = CheckURL
    Set oHTTP = CreateObject("MSXML2.XMLHTTP")
    If oHTTP Is Nothing Then
        sError = "Cannot create HTTP connection"
        PossibleUpdate = False
        On Error GoTo 0
        Exit Function
    End If


    NewBuild = 0        'Class attribute!
    lError = 0
    Do
        With oHTTP
            .Open "GET", sURL, False
            .Send

            sTextResponse = Trim(.ResponseText)

            'If Right(sTextResponse, 1) = "A" Then        'Still Alpha version, not ready for pilot users

            'If GetUserName() = "cklahr" Then
            '    sTextResponse = Left(sTextResponse, Len(sTextResponse) - 1)
            'Else
            '    PossibleUpdate = False
            '    Exit Do
            'End If

            'End If


            Select Case True
                Case Err.Number <> 0
                    sError = " ##Error " & Err.Number & ": " & Err.Description
                    lError = lError + 1
                    PossibleUpdate = False
                Case InStr(sTextResponse, "404 Not Found")
                    sError = " ##Error## 404 Not Found."
                    lError = lError + 1
                    PossibleUpdate = False
                Case Else
                    NewBuild = Val(sTextResponse)

                    PossibleUpdate = True
                    Exit Do
            End Select

            ' Will try a couple of times...
            If lError > 2 Then Exit Do
        End With
    Loop

    Set oHTTP = Nothing

    On Error GoTo 0

End Function

'---------------------------------------------------------------------------------------
' Method : DoUpdate
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function DoUpdate(ByVal bIsBETA As Boolean) As Boolean
    Dim sError As String
    Dim bReturn As Boolean
    Dim lGoAhead As Long

    On Error Resume Next

    bReturn = False
    If Connected2Network Then        'If it's a new version or has a BETA version and wants to fallback
        If (CDbl(NewBuild) > CDbl(Build)) Or (gwsConfig.Range(gsRELEASE_TYPE) = "BETA" And Not gbJoin_BETA_Program) Then

            If CDbl(NewBuild) <> CDbl(GetSetting("PENS", "General", "Version")) Then        ' Only do the upgrade if the new build is different from the Registry version

                If MsgBox("There is an update for " & AppName & "(Version: " & NewBuild & ")" & vbNewLine & vbNewLine & _
                    "Do you wish to download now?" & vbNewLine & _
                    "(You'll have to restart Excel to complete the update)", vbQuestion + vbYesNo, AppName) = vbYes Then

                    lGoAhead = vbYes

                    If bIsBETA Then        'Double Check if it's a BETA version
                        lGoAhead = MsgBox("You're about to install a BETA version" & vbNewLine & vbNewLine & _
                        "Do you want to continue?" & vbNewLine & _
                        "(You'll have to restart Excel to complete the update)", vbQuestion + vbYesNo, AppName)
                    End If


                    If lGoAhead = vbYes Then
                        If GetUpdate Then
                            Application.Cursor = xlDefault
                            MsgBox "Successfully updated the " & AppName & " Add-In!" & vbNewLine & _
                            "Please refer to the Release notes in the Help menu " & vbNewLine & _
                            "for a summary of new enhancements and fixes." & vbNewLine & vbNewLine & _
                            "You must restart Excel to start using the new version!", vbOKOnly + vbInformation, AppName

                            ' Save new version in Registry
                            SaveSetting "PENS", "General", "Version", NewBuild

                            bReturn = True
                        Else
                            Application.Cursor = xlDefault
                            MsgBox "Updating " & AppName & " has failed, please try again later.", _
                            vbInformation + vbOKOnly, AppName
                        End If
                    End If
                End If
            End If
        ElseIf Manual Then
            Application.Cursor = xlDefault
            MsgBox AppName & " is up to date.", vbInformation + vbOKOnly, AppName
        End If
    End If

    DoUpdate = bReturn

    On Error GoTo 0

End Function

'---------------------------------------------------------------------------------------
' Method : Build
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Get Build() As String
    Build = msBuild
End Property

'---------------------------------------------------------------------------------------
' Method : Build
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Let Build(ByVal sBuild As String)
    msBuild = sBuild
End Property

'---------------------------------------------------------------------------------------
' Method : RemoveOldCopy
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub RemoveOldCopy()
    CurrentAddinName = ThisWorkbook.FullName
    TempAddInName = CurrentAddinName & "(OldVersion)"

    On Error Resume Next
    Kill TempAddInName

    On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------
' Method : CurrentAddinName
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Property Get CurrentAddinName() As String
    CurrentAddinName = msCurrentAddinName
End Property

'---------------------------------------------------------------------------------------
' Method : CurrentAddinName
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Property Let CurrentAddinName(ByVal sCurrentAddinName As String)
    msCurrentAddinName = sCurrentAddinName
End Property

'---------------------------------------------------------------------------------------
' Method : TempAddInName
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Property Get TempAddInName() As String
    TempAddInName = msTempAddInName
End Property

'---------------------------------------------------------------------------------------
' Method : TempAddInName
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Property Let TempAddInName(ByVal sTempAddInName As String)
    msTempAddInName = sTempAddInName
End Property

'---------------------------------------------------------------------------------------
' Method : DownloadName
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Get DownloadName() As String
    DownloadName = msDownloadName
End Property

'---------------------------------------------------------------------------------------
' Method : DownloadName
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Let DownloadName(ByVal sDownloadName As String)
    msDownloadName = sDownloadName
End Property

'---------------------------------------------------------------------------------------
' Method : CheckURL
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Get CheckURL() As String
    CheckURL = msCheckURL
End Property

'---------------------------------------------------------------------------------------
' Method : CheckURL
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Let CheckURL(ByVal sCheckURL As String)
    msCheckURL = sCheckURL
End Property


'---------------------------------------------------------------------------------------
' Method : AppName
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Get AppName() As String
    AppName = msAppName
End Property

'---------------------------------------------------------------------------------------
' Method : AppName
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Let AppName(ByVal sAppName As String)
    msAppName = sAppName
End Property

'---------------------------------------------------------------------------------------
' Method : Manual
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Get Manual() As Boolean
    Manual = mbManual
End Property

'---------------------------------------------------------------------------------------
' Method : Manual
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Let Manual(ByVal bManual As Boolean)
    mbManual = bManual
End Property

'---------------------------------------------------------------------------------------
' Method : NewBuild
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Get NewBuild() As String
    NewBuild = msNewBuild
End Property

'---------------------------------------------------------------------------------------
' Method : NewBuild
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Let NewBuild(ByVal sNewBuild As String)
    msNewBuild = sNewBuild
End Property
'---------------------------------------------------------------------------------------
' Method : Connected2Network
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Get Connected2Network() As Boolean
    Connected2Network = mbConnected2Network
End Property

'---------------------------------------------------------------------------------------
' Method : Connected2Network
' Author : JKP
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Property Let Connected2Network(ByVal bC2N As Boolean)
    mbConnected2Network = bC2N
End Property


