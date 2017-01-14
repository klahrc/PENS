Attribute VB_Name = "Module1"
Option Explicit
Const adVarWChar As Long = 202
Const adLongVarWChar As Long = 203
Const adDouble As Long = 5
Const adDecimal As Long = 7
Const adColNullable As Long = 2

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------

Sub Access_Resources(ByVal sFile As String, ByVal bAllText As Boolean)
    '==================================================================
    ' Procedure - Access_MakeTable()
    ' Creates a Table in Access using Excel worksheet as Datasource
    '==================================================================

    Dim strQuery As String
    Dim stDB As String
    Dim xlLoc As String
    Dim ErrorMsg As String
    Dim sRange As String
    Dim sSql As String

    Dim oCatalog As Object
    Dim c As Object
    Dim tblnew As Object
    Dim scn As Object

    stDB = gsLocal_Folder + "\" + KKresource.accdb         'gwsConfig.Range(gsDB_NAME).Value

    Set oCatalog = CreateObject("ADOX.Catalog")

    On Error Resume Next
    Kill stDB

    On Error GoTo errHandler

    oCatalog.Create "provider='Microsoft.ACE.OLEDB.12.0';" & "Data Source=" & stDB

    Set scn = oCatalog.activeconnection

    If bAllText Then
        '''' IMPORT EVERYTHING AS STRING!!!!
        xlLoc = "'" & sFile & "'[Excel 12.0;HDR=YES;IMEX=1;]"
        sSql = "SELECT * INTO tbl_PortfolioPlan FROM [Portfolio Plan$A7:BL1088] IN " & xlLoc & ";"        'Remove hardcoding!!!!!!!!!!!!!!!!!!!!!
    Else
        Set tblnew = CreateObject("ADOX.Table")
        Set c = CreateObject("ADOX.Column")

        ' Create a new Table object.
        With tblnew
            .Name = "tbl_PortfolioPlan"
            ' Create fields and append them to the
            ' Columns collection of the new Table object.
            With .Columns
                .Append "Project Name", adVarWChar
                .Append "Project Code", adVarWChar
                .Append "Roles", adVarWChar
                .Append "JAN", adDouble
                .Append "FEB", adDouble
                .Append "MAR", adDouble
                .Append "APR", adDouble
                .Append "MAY", adDouble
                .Append "JUN", adDouble
                .Append "JUL", adDouble
                .Append "AUG", adDouble
                .Append "SEP", adDouble
                .Append "OCT", adDouble
                .Append "NOV", adDouble
                .Append "DEC", adDouble
                .Append "Do not remove1", adVarWChar
                .Append "Do not remove2", adVarWChar
                .Append "Names", adVarWChar
                .Append "Report", adLongVarWChar
                .Append "Delivery Leader", adVarWChar
                .Append "Account Manager", adVarWChar
                .Append "Activation Status", adVarWChar
                .Append "App", adVarWChar
                .Append "Project Type", adVarWChar
                .Append "Cat", adVarWChar
                .Append "Program", adVarWChar
                .Append "LOB", adVarWChar
                .Append "Project Manager", adVarWChar
                .Append "Publish NE", adVarWChar
            End With

            For Each c In .Columns
                c.Attributes = adColNullable
            Next

        End With

        oCatalog.Tables.Append tblnew


        sRange = "[" & "Portfolio Plan" & "$A3:AC10000]"

        sSql = "INSERT INTO tbl_PortfolioPlan([Project Code]) "        ' No need to enumerate the rest of the fields...
        sSql = sSql & "SELECT * FROM [Excel 12.0 Macro;HDR=YES;DATABASE=" & sFile & "]." & sRange

    End If
    scn.Execute sSql
    ' Debug.Print sSql
    scn.Close

finished:

    Set scn = Nothing
    Set oCatalog = Nothing
    Set c = Nothing
    Set tblnew = Nothing
    Set scn = Nothing

    Exit Sub

errHandler:
    ErrorMsg = "An error has occurred." & Chr(10) & Chr(10) & Err.Number & " - -" & Err.Description
    MsgBox ErrorMsg, vbCritical, "Error Message - " & ThisWorkbook.Name
    Resume finished

End Sub




