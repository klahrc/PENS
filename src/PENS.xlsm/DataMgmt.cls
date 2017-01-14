VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataMgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : DataMgmt
' Author : cklahr
' Date   : 9/25/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Option Explicit

' Set a reference to the Microsoft ActiveX Data Objects 2.X Library
Private mcnConnect As Object

'
Private Sub Class_Terminate()
    On Error Resume Next
    ' Close and destroy the connection object
    mcnConnect.Close
    Set mcnConnect  =  Nothing
End Sub

'---------------------------------------------------------------------------------------
' Method : Initialize
' Author : cklahr
' Date   : 9/25/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub Initialize(ByRef sConnection As String)
    ' Create the connection to the database.  
    Set mcnConnect = CreateObject("ADODB.Connection")
    mcnConnect.ConnectionString = sConnection
    mcnConnect.Open
End Sub

'---------------------------------------------------------------------------------------
' Method : getCustomerData
' Author : cklahr
' Date   : 9/25/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Public Sub getCustomerData(ByRef rngDestination As Excel.Range)
    Dim rsData As Object
    Dim sSQL As String

    Set rsData = CreateObject("ADODB.Recordset")

    sSql = "SELECT * FROM Customers"
    set rsData = CreateObject("ADODB.Recordset")
    rsData.Open sSQL ,  mcnConnect,  adOpenForwardOnly,  adLockReadOnly

    If Not rsData.EOF Then
        rngDestination.CopyFromRecordset rsData
    End If

    rsData.Close
End Sub









