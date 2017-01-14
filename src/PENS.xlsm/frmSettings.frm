VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "PENS - SETTINGS"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10005
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkForceLocal_Click()
    On Error Resume Next
    '''gwsConfig.Range(gsUSE_LOCAL_DATA).Value = optUseLocalSources.Value
    gbUse_Local_Folder = optUseLocalSources.Value
    On Error GoTo 0
End Sub


Private Sub chkJoinBETA_Click()

    gbJoin_BETA_Program = chkJoinBETA.Value

End Sub

Private Sub cmdTestOK_Click()
    Dim i As Long
    Dim j As Long

    i = 0
    Do While (i = 0)        ' Continue verifying
        i = ValidateFileFolders(Not optUseLocalSources.Value, True)        'Should always check LocalFolder is valid (it's used to store outputs)
    Loop

    If i = 1 Then
        Call SaveAppSettings
        Call RereshSettingsForm
        j = MsgBox("Test completed succesfully", vbOKOnly, "CONGRATS!")
    Else
        '''' MsgBox "Test failed"
        j = MsgBox("Oops...Test failed", vbCritical, "ERROR")
    End If

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub optUseLiveData_Click()
    If gbConnected2Network Then
        '''gwsConfig.Range(gsUSE_LOCAL_DATA).Value = False
        gbUse_Local_Folder = False
    Else
        MsgBox ("No network detected. Forcing use of local Data Sources...")
        optUseLocalSources.Value = True
    End If
End Sub

Private Sub optUseLocalSources_Click()
    '''gwsConfig.Range(gsUSE_LOCAL_DATA).Value = True
    gbUse_Local_Folder = True
End Sub

'---------------------------------------------------------------------------------------
' Method : txtFileNamePP_Change
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub txtFileNamePP_Change()
    On Error Resume Next
    '''gwsConfig.Range(gsFILENAME_PP).Value = txtFileNamePP.Value
    gsPP_Filename = txtFileNamePP.Value
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------
' Method : txtFileNameRS_Change
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub txtFileNameRS_Change()
    On Error Resume Next
    '''gwsConfig.Range(gsFILENAME_RS).Value = txtFileNameRS.Value
    gsRS_Filename = txtFileNameRS.Value
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Method : txtLocalFolder_Change
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub txtLocalFolder_Change()
    On Error Resume Next
    '''gwsConfig.Range(gsLOCAL_FOLDER).Value = txtLocalFolder.Value
    gsLocal_Folder = txtLocalFolder.Value
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Method : txtRemFolderPP_Change
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub txtRemFolderPP_Change()
    On Error Resume Next
    '''gwsConfig.Range(gsREM_FOLDER_PP).Value = txtRemFolderPP.Value
    gsPP_Network_Folder = txtRemFolderPP.Value
    On Error GoTo 0

End Sub

'---------------------------------------------------------------------------------------
' Method : txtRemFolderRS_Change
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub txtRemFolderRS_Change()
    On Error Resume Next
    '''gwsConfig.Range(gsREM_FOLDER_RS).Value = txtRemFolderRS.Value
    gsRS_Network_Folder = txtRemFolderRS.Value
    On Error GoTo 0
End Sub
'---------------------------------------------------------------------------------------
' UserForm Method Procedures Follow
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Method : Initialize
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:None
' Pending:
' Comments:Providing a custom Initialize method for your UserForms allows you to link the error handling of
'           the UserForm initialization into the central error handling system. There is easy way to do this in
'           the UserForm Initialize Event procedure, so in any case where UserForm initialization might fail we
'           recommend you create a custom Boolean Initialize function rather than using the Initialize event.
'---------------------------------------------------------------------------------------
Public Function Initialize() As Boolean
    Const sSOURCE As String = "Initialize()"

    Dim bReturn As Boolean

    On Error GoTo Initialize_Error

    '?????? Application.EnableCancelKey = xlDisabled ????????????????????????????????????????

    Call RereshSettingsForm

    '''' Set a variable to force reshresh in every run
    '''''''chkRefreshDataSources.Value = gwsConfig.Range("C11").Value
    ''' If (gwsConfig.Range(gsUSE_LOCAL_DATA).Value) Then
    'If (gbUse_Local_Folder) Then
    '    optUseLocalSources.Value = True
    'Else
    '    optUseLiveData.Value = True
    'End If

    ' lblChkSum will show the grand total for the Pivot Tables if all match
    '!!!!!!!lblChkSum.Caption = "Waiting for Pivot Checksum results..."

    'If required, I can show an ad :-)
    If gwsConfig.Range("C18").Value Then
        ''''''Call ad
        gwsConfig.Range("C18").Value = False        'Only for the first time...every time would be annoying
    End If

    cmdTestOK.Caption = "Test & Save!"

    Initialize = True

    Exit Function

ErrorExit:
    'Clean up
    Initialize = False
    Exit Function

Initialize_Error:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Function

'---------------------------------------------------------------------------------------
' Method : RereshSettingsForm
' Author : cklahr
' Date   : 2/18/2016
' Purpose:
' Arguments:
' Pending: ERROR HANDLING!!!!!!!
' Comments:
'---------------------------------------------------------------------------------------
Private Sub RereshSettingsForm()


    'Refresh form with last settings
    With Me
        .txtFileNamePP.Value = gsPP_Filename
        .txtRemFolderPP.Value = gsPP_Network_Folder

        .txtFileNameRS.Value = gsRS_Filename
        .txtRemFolderRS.Value = gsRS_Network_Folder

        .txtLocalFolder.Value = gsLocal_Folder

        If gbUse_Local_Folder Then
            .optUseLocalSources.Value = True
        Else
            .optUseLiveData.Value = True
        End If

        .chkJoinBETA = gbJoin_BETA_Program

    End With



End Sub


Private Sub UserForm_Click()

End Sub
