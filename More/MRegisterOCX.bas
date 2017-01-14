Attribute VB_Name = "MRegisterOCX"
'NOTE: YOU CAN CHANGE THE NAME OF THE .OCX OR .DLL FILE
'IN THE ZIPPED EXAMPLE TO ONE OF YOUR OWN CHOICE.
 
'***************CODE FOR MODULE1***************
 
Option Explicit
 
'---------------------------------------------------------------------------------------
' Method : PutFileInSystem
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub PutFileInSystem()
    Dim FileSysObject As Object
    Dim FileName$, FilesOldPath$, FilesNewPath$
     
    FileName = [D3]
    FilesOldPath = ActiveWorkbook.Path & "\"
    FilesNewPath = "C:\Windows\System\"
     
    Set FileSysObject = CreateObject("Scripting.FileSystemObject")
    If Not FileSysObject.FileExists(FilesOldPath & FileName) Then
        MsgBox "The file " & FilesOldPath & FileName & " was not found...", , _
        "File Is Missing"
    ElseIf Not FileSysObject.FileExists(FilesNewPath & FileName) Then
        'move the file to the new location
        FileSysObject.MoveFile (FilesOldPath & FileName), FilesNewPath & FileName
        MsgBox FilesOldPath & FileName & vbLf & vbNewLine & _
        "has been installed in the location given below:" & vbLf & vbNewLine & _
        FilesNewPath & FileName
    Else
        MsgBox "Sorry, the install cannot be performed. There is" & vbLf & _
        "already a " & FilesNewPath & FileName, , "Destination File Already Exists"
    End If
    RegisterIt
End Sub
 
'---------------------------------------------------------------------------------------
' Method : TakeFileFromSystem
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub TakeFileFromSystem()
    Dim FileSysObject As Object
    Dim FileName$, FilesOldPath$, FilesNewPath$
    FileName = [D3]
    
    FilesOldPath = "C:\Windows\System\"
    FilesNewPath = ActiveWorkbook.Path & "\"
    Set FileSysObject = CreateObject("Scripting.FileSystemObject")
    If Not FileSysObject.FileExists(FilesOldPath & FileName) Then
        MsgBox "The file " & FilesOldPath & FileName & " was not found...", , _
        "File Is Missing"
    ElseIf Not FileSysObject.FileExists(FilesNewPath & FileName) Then
        'move the file to the new location
        On Error GoTo ErrorMsg
        FileSysObject.MoveFile (FilesOldPath & FileName), FilesNewPath & FileName
        MsgBox FilesOldPath & FileName & vbLf & vbNewLine & _
        "has been moved to the location given below:" & vbLf & vbNewLine & _
        FilesNewPath & FileName
    Else
        MsgBox "Sorry, the file removal cannot be performed. There is an existing " & _
        FileName & vbLf & _
        "file in " & FilesNewPath & " please remove it first", , "File In The Way..."
    End If
    DeRegisterIt
    Exit Sub
ErrorMsg:
    MsgBox "This workbook has a reference set to the file you're trying to uninstall, " _
    & vbLf & "you will need to go into the VBE window, select Tools/References and " _
    & vbLf & "uncheck that particular reference before you can uninstall the file."
    End
End Sub
  
'---------------------------------------------------------------------------------------
' Method : RegisterIt
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub RegisterIt()
    Dim Tmp$, FilesName$, Ref As Object
    Dim FileSysObject As Object
    Const FilesPath = "C:\Windows\System\"
    
    FilesName = [D3]
    Set FileSysObject = CreateObject("Scripting.FileSystemObject")
    If Not FileSysObject.FileExists(FilesPath & FilesName) Then
        MsgBox "The file " & FilesPath & FilesName & " was not found...", , _
        "Cannot Be Registered"
        Exit Sub
    End If
    Tmp = Register("c:\windows\system\" & FilesName)
    MsgBox FilesName & " Registered"
End Sub
 
'Note: Different to registering in this respect.
'The file may've already been removed (say,
'manually) and we just want to de-register it
 
'---------------------------------------------------------------------------------------
' Method : DeRegisterIt
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub DeRegisterIt()
    Dim Tmp$, FilesName$, Ref As Object
    FilesName = [D3]
    Tmp = DeRegister("c:\windows\system\" & FilesName)
    MsgBox FilesName & " Deregistered"
End Sub
'**********************************************

