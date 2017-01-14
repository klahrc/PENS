Attribute VB_Name = "modiGridOCX"
'******************************************************************************************
'                                                                                         *
'iGrid500 OCX GUID: HKEY_CLASSES_ROOT\TypeLib\{64B45CCC-07E2-4394-8BD1-24D27C18D694}      *
'                                                                                         *
'******************************************************************************************

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' API Constants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const HKEY_CURRENT_USER As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Error Constants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const C_REG_ERR_NO_ERROR = 0
Public Const C_REG_ERR_INVALID_BASE_KEY = vbObjectError + 1
Public Const C_REG_ERR_INVALID_DATA_TYPE = vbObjectError + 2
Public Const C_REG_ERR_KEY_NOT_FOUND = vbObjectError + 3
Public Const C_REG_ERR_VALUE_NOT_FOUND = vbObjectError + 4
Public Const C_REG_ERR_DATA_TYPE_MISMATCH = vbObjectError + 5
Public Const C_REG_ERR_ENTRY_LOCKED = vbObjectError + 6
Public Const C_REG_ERR_INVALID_KEYNAME = vbObjectError + 7
Public Const C_REG_ERR_UNABLE_TO_OPEN_KEY = vbObjectError + 8
Public Const C_REG_ERR_UNABLE_TO_READ_KEY = vbObjectError + 9
Public Const C_REG_ERR_UNABLE_TO_CREATE_KEY = vbObjectError + 10
Public Const C_REG_ERR_UBABLE_TO_READ_VALUE = vbObjectError + 11
Public Const C_REG_ERR_UNABLE_TO_UDPATE_VALUE = vbObjectError + 12
Public Const C_REG_ERR_UNABLE_TO_CREATE_VALUE = vbObjectError + 13
Public Const C_REG_ERR_UNABLE_TO_DELETE_KEY = vbObjectError + 14
Public Const C_REG_ERR_UNABLE_TO_DELETE_VALUE = vbObjectError + 15

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public  Variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public G_Reg_AppErrNum As Long
Public G_Reg_AppErrText As String
Public G_Reg_SysErrNum As Long
Public G_Reg_SysErrText As String

Private Const ERROR_SUCCESS = 0&

Private Const REGSTR_MAX_VALUE_LENGTH As Long = &H100

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" ( _
ByVal HKey As Long, _
ByVal lpSubKey As String, _
phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
ByVal HKey As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long


Private Function IsValidKeyName(KeyName As String) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsValidKeyName
    ' Returns True or False indicating whether KeyName is valid.
    ' An invalid key is one whose name length is greater than
    ' REGSTR_MAX_VALUE_LENGTH or is all spaces or is an empty string.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    IsValidKeyName = (Len(KeyName) <= REGSTR_MAX_VALUE_LENGTH) And (Len(Trim(KeyName)) > 0)
End Function

Private Function GetAppErrText(ErrNum As Long) As String
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' GetAppErrText
    ' This returns the text description of the application error
    ' number in ErrNum.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case ErrNum
        Case C_REG_ERR_NO_ERROR
            GetAppErrText = vbNullString
        Case C_REG_ERR_INVALID_BASE_KEY
            GetAppErrText = "Invalid Base Key Value."
        Case C_REG_ERR_INVALID_DATA_TYPE
            GetAppErrText = "Invalid Data Type."
        Case C_REG_ERR_KEY_NOT_FOUND
            GetAppErrText = "Key Not Found."
        Case C_REG_ERR_VALUE_NOT_FOUND
            GetAppErrText = "Value Not Found."
        Case C_REG_ERR_DATA_TYPE_MISMATCH
            GetAppErrText = "Value Data Type Mismatch."
        Case C_REG_ERR_ENTRY_LOCKED
            GetAppErrText = "Registry Entry Locked."
        Case C_REG_ERR_INVALID_KEYNAME
            GetAppErrText = "The Specified Key Is Invalid."
        Case C_REG_ERR_UNABLE_TO_OPEN_KEY
            GetAppErrText = "Unable To Open Key."
        Case C_REG_ERR_UNABLE_TO_READ_KEY
            GetAppErrText = "Unable To Read Key."
        Case C_REG_ERR_UNABLE_TO_CREATE_KEY
            GetAppErrText = "Unable To Create Key."
        Case C_REG_ERR_UBABLE_TO_READ_VALUE
            GetAppErrText = "Unable To Read Value."
        Case C_REG_ERR_UNABLE_TO_UDPATE_VALUE
            GetAppErrText = "Unable To Update Value."
        Case C_REG_ERR_UNABLE_TO_CREATE_VALUE
            GetAppErrText = "Unable To Create Value."
        Case C_REG_ERR_UNABLE_TO_DELETE_KEY
            GetAppErrText = "Unable To Delete Key."
        Case C_REG_ERR_UNABLE_TO_DELETE_VALUE
            GetAppErrText = "Unable To Delete Value."

        Case Else
            GetAppErrText = "Undefined Error."
    End Select
End Function


Private Function IsValidBaseKey(BaseKey As Long) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' IsValidBaseKey
    ' This returns True of BaseKey is valid base key
    ' (HKEY_CURRENT_USER etc) or False if BaseKey is not
    ' a valid base key.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case BaseKey
        Case HKEY_CURRENT_USER, HKEY_LOCAL_MACHINE, HKEY_CLASSES_ROOT, HKEY_CURRENT_CONFIG, HKEY_DYN_DATA, HKEY_PERFORMANCE_DATA, HKEY_USERS
            IsValidBaseKey = True
        Case Else
            IsValidBaseKey = False
    End Select

End Function

Private Sub ResetErrorVariables()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ResetErrorVariables
    ' This resets the global error values to their default
    ' (no error) values.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    G_Reg_AppErrNum = C_REG_ERR_NO_ERROR
    G_Reg_AppErrText = vbNullString
    G_Reg_SysErrNum = C_REG_ERR_NO_ERROR
    G_Reg_SysErrText = vbNullString
End Sub

Public Function RegistryKeyExists(BaseKey As Long, KeyName As String) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' RegistryKeyExists
    ' Returns True or False indicating whether KeyName exists in BaseKey.
    ' Returns False if an error occurred. See the global error values
    ' for more information.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim HKey As Long
    Dim Res As Long

    ''ResetErrorVariables
    If IsValidBaseKey(BaseKey:=BaseKey) = False Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = GetAppErrText(G_Reg_AppErrNum)
        RegistryKeyExists = False
    End If

    If IsValidKeyName(KeyName:=KeyName) = False Then
        G_Reg_AppErrNum = C_REG_ERR_INVALID_BASE_KEY
        G_Reg_AppErrText = GetAppErrText(G_Reg_AppErrNum)
        RegistryKeyExists = False
    End If

    Res = RegOpenKey(HKey:=BaseKey, lpSubKey:=KeyName, phkResult:=HKey)
    If Res = ERROR_SUCCESS Then
        RegistryKeyExists = True
    Else
        RegistryKeyExists = False
    End If

    RegCloseKey HKey:=HKey

End Function

'---------------------------------------------------------------------------------------
' Method : iGridOcxRegistered
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function iGridOcxRegistered() As Boolean
    iGridOcxRegistered = False

    If (RegistryKeyExists(HKEY_CLASSES_ROOT, "TypeLib\{64B45CCC-07E2-4394-8BD1-24D27C18D694}")) Then
        iGridOcxRegistered = True
    End If

End Function
'---------------------------------------------------------------------------------------
' Method : RegisterFile
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub RegisterFile(ByVal sFileName As String)
    ShellExecute 0, "runas", "cmd", "/c regsvr32 /s " & """" & sFileName & """", "C:\", 0 'SW_HIDE =0
End Sub

'---------------------------------------------------------------------------------------
' Method : AddReference
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub AddReference()
    'Macro purpose:  To add a reference to the project using the GUID for the
    'reference library

    Dim strGUID As String, theRef As Variant, i As Long

    'Update the GUID you need below.
    strGUID = "{64B45CCC-07E2-4394-8BD1-24D27C18D694}"

    'Set to continue in case of error
    On Error Resume Next

    'Remove any missing references
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.isbroken = True Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i

    'Clear any errors so that error trapping for GUID additions can be evaluated
    Err.Clear

    'Add the reference
    ThisWorkbook.VBProject.References.AddFromGuid _
    GUID:=strGUID, Major:=1, Minor:=0

    'If an error was encountered, inform the user
    Select Case Err.Number
        Case Is = 32813
            'Reference already in use.  No action necessary
        Case Is = vbNullString
            'Reference added without issue
        Case Else
            'An unknown error was encountered, so alert the user
            MsgBox "A problem was encountered trying to" & vbNewLine _
            & "add or remove a reference in this file" & vbNewLine & "Please check the " _
            & "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
    End Select
    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Method : registerIgridOCX
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function registerIgridOCX()
    Dim sRemPathOCX As String

    ' Copy and Register OCX if not installed yet!
    If Not iGridOcxRegistered() Then
        If gbJoin_BETA_Program Then
            sRemPathOCX = gwsConfig.Range(gsUPDATE_FOLDER) + "\BETA" + "\" + gsOCXFilename
        Else
            sRemPathOCX = gwsConfig.Range(gsUPDATE_FOLDER) + "\" + gsOCXFilename
        End If

        ' Copy iGrid500 OCX to Addin folder
        Call VBCopyFileFolder(sRemPathOCX, Application.UserLibraryPath & gsOCXFilename) '' Ojo que es function

        ' Register from Addin folder
        Call RegisterFile(Application.UserLibraryPath & "\" & gsOCXFilename)

        ' Add Reference to VBE
        Call AddReference
    End If
End Function












