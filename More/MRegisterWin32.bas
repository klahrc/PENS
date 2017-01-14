Attribute VB_Name = "MRegisterWin32"
Option Explicit
 
'All required Win32 SDK functions to register/unregister any ActiveX component
Private Declare Function LoadLibraryRegister Lib "KERNEL32" Alias _
"LoadLibraryA" (ByVal lpLibFileName$) As Long
 
Private Declare Function FreeLibraryRegister Lib "KERNEL32" Alias _
"FreeLibrary" (ByVal hLibModule&) As Long
 
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject&) As Long
Private Declare Function GetProcAddressRegister Lib "KERNEL32" Alias _
"GetProcAddress" (ByVal hModule&, _
ByVal lpProcName$) As Long
 
Private Declare Function CreateThreadForRegister Lib "KERNEL32" Alias _
"CreateThread" (lpThreadAttributes As Any, _
ByVal dwStackSize&, ByVal lpStartAddress&, _
ByVal lpparameter&, ByVal dwCreationFlags&, _
ThreadID&) As Long
 
Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle&, _
ByVal dwMilliseconds&) As Long
 
Private Declare Function GetExitCodeThread Lib "KERNEL32" (ByVal Thread&, _
lpExitCode&) As Long
 
Private Declare Sub ExitThread Lib "KERNEL32" (ByVal ExitCode&)
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Public Const DllRegisterServer = 1
Public Const DllUnRegisterServer = 2
 
'---------------------------------------------------------------------------------------
' Method : Register
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function Register(FileName$) As String
    If Dir(FileName) = Empty Then
        Register = "File not found"
        Exit Function
    Else
        Register = RegisterFile(FileName, DllRegisterServer)
    End If
End Function
 
'---------------------------------------------------------------------------------------
' Method : DeRegister
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function DeRegister(FileName$) As String
    If Dir(FileName) = Empty Then
        DeRegister = "File not found"
        Exit Function
    Else
        DeRegister = RegisterFile(FileName, DllUnRegisterServer)
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
Function RegisterFile(ByVal FileName$, ByVal RegFunction&) As String
    Dim LoadLib&, ProcAddress&, ThreadID&, Successful&, ExitCode&, Thread&
    If FileName = Empty Then Exit Function
    LoadLib = LoadLibraryRegister(FileName)
    If LoadLib = 0 Then
        RegisterFile = "File Can't Be Loaded"              'Couldn't load component
        Exit Function
    End If
    If RegFunction = DllRegisterServer Then
        ProcAddress = GetProcAddressRegister(LoadLib, "DllRegisterServer")
    ElseIf RegFunction = DllUnRegisterServer Then
        ProcAddress = GetProcAddressRegister(LoadLib, "DllUnregisterServer")
    End If
    If ProcAddress = 0 Then
        RegisterFile = "Not ActiveX Component"
        If LoadLib Then FreeLibraryRegister (LoadLib)
        Exit Function
    Else
        Thread = CreateThreadForRegister(ByVal 0&, 0&, ByVal ProcAddress, _
        ByVal 0&, 0&, ThreadID)
        If Thread Then
            Successful = (WaitForSingleObject(Thread, 10000) = WAIT_OBJECT_0)
            If Not Successful Then
                Call GetExitCodeThread(Thread, ExitCode)
                ExitThread (ExitCode)
                RegisterFile = "Component Registration Failed"
                If LoadLib Then FreeLibraryRegister (LoadLib)
                Exit Function
            Else
                If RegFunction = DllRegisterServer Then
                    RegisterFile = Empty                   'registered successfully
                ElseIf RegFunction = DllUnRegisterServer Then
                    RegisterFile = Empty                   'unregistered successfully
                End If
            End If
            CloseHandle (Thread)
            If LoadLib Then FreeLibraryRegister (LoadLib)
        End If
    End If
End Function


