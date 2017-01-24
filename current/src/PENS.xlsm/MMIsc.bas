Attribute VB_Name = "MMIsc"
'---------------------------------------------------------------------------------------
' File   : MMIsc
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------
Private Const msMODULE As String = "MMisc"
Const NoError = 0                                          'The Function call was successful

'---------------------------------------------------------------------------------------
' Declarations
'---------------------------------------------------------------------------------------
Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

'32-bit API declarations
Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) _
As Long

Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Declare Function GetForegroundWindow Lib "User32.dll" () As Long

Private Declare Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" _
(ByVal hWnd As Long, _
ByVal nIndex As Long) _
As Long

Private Declare Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" _
(ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) _
As Long

Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" _
(ByVal hwndOwner As Long, ByVal nFolder As Long, _
ByVal hToken As Long, ByVal dwFlags As Long, _
ByVal pszPath As String) As Long

Private Const CSIDL_PERSONAL As Long = &H5

Private Const WS_THICKFRAME As Long = &H40000
Private Const GWL_STYLE As Long = -16

Public Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" _
(ByVal lpszLocalName As String, _
ByVal lpszRemoteName As String, _
cbRemoteName As Long) As Long

Private Declare Function PathIsUNC Lib "shlwapi" Alias "PathIsUNCA" (ByVal pszPath As String) As Long

' Declare for call to mpr.dll.
Declare Function WNetGetUser Lib "mpr.dll" _
Alias "WNetGetUserA" (ByVal lpName As String, _
ByVal lpUserName As String, lpnLength As Long) As Long

'---------------------------------------------------------------------------------------
' Declarations for VBCopyFileFolder
'---------------------------------------------------------------------------------------
Private Declare Function SHFileOperation Lib "shell32.dll" _
Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_COPY = &H2

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type

'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
Public Declare Function InternetGetConnectedState _
Lib "wininet.dll" (lpdwFlags As Long, _
ByVal dwReserved As Long) As Boolean

'---------------------------------------------------------------------------------------
' Method : IsConnected
' Author : cklahr
' Date   : 2/2/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function IsConnected() As Boolean
    Dim Stat As Long

    IsConnected = (InternetGetConnectedState(Stat, 0&) <> 0)
End Function
'---------------------------------------------------------------------------------------
' Method : VBCopyFileFolder
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function VBCopyFileFolder(ByRef strSource As String, ByRef strTarget As String) As Boolean
    Dim op As SHFILEOPSTRUCT
    Dim FSO As Object
    Dim bRet As Boolean

    'Set Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    'If File Exists, delete it first
    If FSO.FileExists(strTarget) Then
        FSO.DeleteFile strTarget, True
    End If

    With op
        .wFunc = FO_COPY
        .pTo = strTarget
        .pFrom = strSource
        '''' .fFlags = FOF_SIMPLEPROGRESS
    End With

    '~~> Perform operation
    SHFileOperation op

    ' If copy was succesful I should be able to find strTarget!
    If FSO.FileExists(strTarget) Then
        bRet = True
    Else
        bRet = False
    End If

    Set FSO = Nothing

    VBCopyFileFolder = bRet
End Function

'---------------------------------------------------------------------------------------
' Method : FileLastModified
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function FileLastModified(sFileName As String) As String
    Dim fs As Object
    Dim f As Object
    Dim s As String

    Set fs = CreateObject("Scripting.FileSystemObject")

    If fs.FileExists(sFileName) Then
        Set f = fs.GetFile(sFileName)
        FileLastModified = Format(f.DateLastModified, "mmm-dd-yyyy-h:mm:ss")
    Else
        FileLastModified = "N/A"
    End If

    Set fs = Nothing
    Set f = Nothing
End Function

'---------------------------------------------------------------------------------------
' Method : FileFolderExists
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function FileFolderExists(strFullPath As String) As Boolean

    On Error Resume Next

    If Not Dir(strFullPath, vbDirectory) = vbNullString Then
        FileFolderExists = True
    End If

    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Method : MakeFormResizable
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Private Sub MakeFormResizable()
    Dim lStyle As Long
    Dim hWnd As Long
    Dim RetVal

    On Error Resume Next

    hWnd = GetForegroundWindow

    'Get the basic window style
    lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME

    'Set the basic window styles
    RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)

    On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Method : MyDocuments
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function MyDocuments() As String
    Dim pos As Long
    Dim sBuffer As String

    On Error Resume Next

    sBuffer = Space$(260)
    If SHGetFolderPath(0&, CSIDL_PERSONAL, -1, 0&, sBuffer) = 0 Then
        pos = InStr(1, sBuffer, Chr(0))
        MyDocuments = Left$(sBuffer, pos - 1)
    End If

    On Error GoTo 0
End Function

'Public Function MyDocuments() As String
'    Dim pos As Long
'    Dim sBuffer As String
'
'    On Error Resume Next
'
'    sBuffer = Space$(260)
'    If SHGetFolderPath(0&, CSIDL_PERSONAL, -1, 0&, sBuffer) = 0 Then
'        pos = InStr(1, sBuffer, Chr(0))
'        MyDocuments = Left$(sBuffer, pos - 1)
'    End If
'
'    On Error GoTo 0
'End Function
'---------------------------------------------------------------------------------------
' Method : AlphaNumeric
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function AlphaNumeric(pValue) As Boolean
    Dim lPos As Integer
    Dim LChar As String
    Dim LValid_Values As String

    On Error Resume Next

    'Start at first character in pValue
    lPos = 1

    'Set up values that are considered to be alphanumeric
    LValid_Values = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

    'Test each character in pValue
    While lPos <= Len(pValue)
        'Single character in pValue
        LChar = Mid(pValue, lPos, 1)

        'If character is not alphanumeric, return FALSE
        If InStr(LValid_Values, LChar) = 0 Then
            AlphaNumeric = False
            Exit Function
        End If

        'Increment counter
        lPos = lPos + 1
    Wend

    'Value is alphanumeric, return TRUE
    AlphaNumeric = True

    On Error GoTo 0
End Function
'---------------------------------------------------------------------------------------
' Method : pathExists
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function pathExists(pName) As Boolean

    On Error Resume Next

    If Dir(pName, vbDirectory) = "" Then
        Exit Function
    Else
        pathExists = (GetAttr(pName) And vbDirectory) = vbDirectory
    End If

    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Method : FileUNC
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function FileUNC(ByVal strPath As String) As String
    Dim strNetPath As String

    On Error Resume Next

    strNetPath = String(255, Chr(0))
    WNetGetConnection Left(strPath, 2), strNetPath, 255
    If PathIsUNC(strNetPath) Then
        FileUNC = Left(strNetPath, InStr(1, strNetPath, Chr(0)) - 1) & Right(strPath, Len(strPath) - 2)
    Else
        FileUNC = strPath
    End If

    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Method : Ceiling
' Author : cklahr
' Date   : 1/31/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function Ceiling(ByVal x As Double, Optional ByVal Factor As Double = 1) As Double
    On Error Resume Next

    ' X is the value you want to round
    ' Factor is the multiple to which you want to round

    Ceiling = (Int(x / Factor) - (x / Factor - Int(x / Factor) > 0)) * Factor



    On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Method : GetUserName
' Author : cklahr
' Date   : 2/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Function GetUserName() As String

    Const lpnLength As Integer = 255                       ' Buffer size for the return string.

    Dim Status As Integer                                  ' Get return buffer space.
    Dim lpName, lpUserName As String                       ' For getting user information.

    lpUserName = Space$(lpnLength + 1)                     ' Assign the buffer size constant to lpUserName.

    Status = WNetGetUser(lpName, lpUserName, lpnLength)    ' Get the log-on name of the person using product.

    ' See whether error occurred.
    If Status = NoError Then
        ' This line removes the null character. Strings in C are null-
        ' terminated. Strings in Visual Basic are not null-terminated.
        ' The null character must be removed from the C strings to be used
        ' cleanly in Visual Basic.
        lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    Else
        ' An error occurred.
        GetUserName = "UNKNOWN"
        Exit Function
    End If

    ' Display the name of the person logged on to the machine.
    GetUserName = lpUserName

End Function

