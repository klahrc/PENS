Attribute VB_Name = "MErrorHandler"
'
' Description:  This module contains the central error handler and related constant declarations.
'
' Authors:     Cesar klahr
'
' Comment
' --------------------------------------------------------------
' Initial version
'
Option Explicit
Option Private Module




'---------------------------------------------------------------------------------------
' Global Constant Declarations Follow
'---------------------------------------------------------------------------------------
Public Const glHANDLED_ERROR As Long = 9999        ' Run-time error number for our custom errors.
Public Const glUSER_CANCEL As Long = 18        ' The error number generated when the user cancels program execution.

Public Const glCANT_OPEN_DEBUG_FILE As Long = 9998        ' Can't Open Debug File
Public Const glCANT_WRITE_DEBUG_FILE As Long = 9997        ' Can't Write Debug File
Public Const glCANT_UPDATE_SOURCES_STATUS As Long = 9996        ' Can't update status of Source files
Public Const glCANT_INITIALIZE As Long = 9995
Public Const glUPDATE_NOT_WORKING As Long = 9994



' **************************************************************
' Module Constant Declarations Follow
' **************************************************************
Private Const msSILENT_ERROR As String = "UserCancel"        ' Used by the central error handler to bail out silently on user cancel.
Private Const msFILE_ERROR_LOG As String = "Error.log"        ' The name of the file where error messages will be logged to.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: This is the central error handling procedure for the
'           program. It logs and displays any run-time errors
'           that occur during program execution.
'
' Arguments:    sModule         The module in which the error occured.
'               sProc           The procedure in which the error occured.
'               sFile           (Optional) For multiple-workbook
'                               projects this is the name of the
'                               workbook in which the error occured.
'               bEntryPoint     (Optional) True if this call is
'                               being made from an entry point
'                               procedure. If so, an error message
'                               will be displayed to the user.
'
' Returns:      Boolean         True if the program is in debug
'                               mode, False if it is not.
'
' Date          Developer       Action
' --------------------------------------------------------------
' 9/9/13       Cesar            Initial version
'
Public Function bCentralErrorHandler( _
    ByVal sModule As String, _
    ByVal sProc As String, _
    Optional ByVal sFile As String, _
    Optional ByVal bEntryPoint As Boolean) As Boolean

    Static sErrMsg As String

    Dim iFile As Integer
    Dim lErrNum As Long
    Dim sFullSource As String
    Dim spath As String
    Dim sLogText As String

    ''' Include case options for missing error codes....
    'Public Const glERR_POP_GROWTH100KSTAGING_DATA As Long = 9980 ' Can't Populate 100KCHART Data
    'Public Const glERR_POP_FORMULAS4GP_GROWTH100KSTAGING_SHEET As Long = 9976 'Can't populate Formulas for GrossPerf in 100KGrowth
    'Public Const glERR_POP_RAW_BMK_DATA = 9975 ' Can't populate Raw BMK data to all staging sheets

    Select Case Err.Number
        Case glCANT_UPDATE_SOURCES_STATUS
            sErrMsg = "Can't update status of source files"

    End Select

    ' Grab the error info before it's cleared by
    ' On Error Resume Next below.
    lErrNum = Err.Number
    ' If this is a user cancel, set the silent error flag
    ' message. This will cause the error to be ignored.
    If lErrNum = glUSER_CANCEL Then sErrMsg = msSILENT_ERROR
    ' If this is the originating error, the static error
    ' message variable will be empty. In that case, store
    ' the originating error message in the static variable.
    If Len(sErrMsg) = 0 Then
        sErrMsg = Err.Description
    Else
        ' sErrMsg = sErrMsg & vbCrLf & "Click OK to finish"
    End If

    ' We cannot allow errors in the central error handler.
    On Error Resume Next

    ' Load the default filename if required.
    If Len(sFile) = 0 Then sFile = ThisWorkbook.Name

    ' Get the application directory.
    spath = ThisWorkbook.Path
    If Right$(spath, 1) <> "\" Then spath = spath & "\"

    ' Construct the fully-qualified error source name.
    sFullSource = "[" & sFile & "]" & sModule & "." & sProc

    ' Create the error text to be logged.
    sLogText = "  " & sFullSource & ", Error " & CStr(lErrNum) & ": " & sErrMsg

    ' Open the log file, write out the error information and
    ' close the log file.
    iFile = FreeFile()
    Open spath & msFILE_ERROR_LOG For Append As #iFile
    Print #iFile, Format$(Now(), "mm/dd/yy hh:mm:ss"); sLogText
    If bEntryPoint Then Print #iFile,
    Close #iFile

    ' Do not display or debug silent errors.
    If sErrMsg <> msSILENT_ERROR Then

        ' Show the error message when we reach the entry point
        ' procedure or immediately if we are in debug mode.
        'If bEntryPoint Or gbDebugMode Then
        If gbDebugMode Then
            'Application.ScreenUpdating = True
            MsgBox sErrMsg, vbCritical, gsAPP_NAME
            ' Clear the static error message variable once
            ' we've reached the entry point so that we're ready
            ' to handle the next error.
            sErrMsg = vbNullString
        End If

        ' The return vale is the debug mode status.
        bCentralErrorHandler = gbDebugMode

    Else
        ' If this is a silent error, clear the static error
        ' message variable when we reach the entry point.
        If bEntryPoint Then sErrMsg = vbNullString
        bCentralErrorHandler = False
    End If

End Function


