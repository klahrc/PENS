Attribute VB_Name = "modReferences"
'---------------------------------------------------------------------------------------
' Method : DeleteReference
' Author : cklahr
' Date   : 12/18/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub DeleteReference()
    'Macro purpose:  To add a reference to the project using the GUID for the
    'reference library
     
    Dim strGUID As String, theRef As Variant, i As Long
     
    'Update the GUID you need below.
    strGUID = "{64B45CCC-07E2-4394-8BD1-24D27C18D694}"
     
    'Set to continue in case of error
    On Error Resume Next
     
    'Clear any errors so that error trapping for GUID additions can be evaluated
    Err.Clear
     
    'Remove any missing references
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.GUID = strGUID Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i
     
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

