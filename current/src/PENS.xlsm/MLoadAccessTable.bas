Attribute VB_Name = "MLoadAccessTable"
Option Explicit
Const adVarWChar As Long = 202
Const adLongVarWChar As Long = 203
Const adDouble As Long = 5
Const adDecimal As Long = 7
Const adColNullable As Long = 2

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------

Sub Access_MakeTable(ByVal sFile As String, ByVal bAllText As Boolean)
    '==================================================================
    ' Procedure - Access_MakeTable()
    ' Creates a Table in Access using Excel worksheet as Datasource
    '==================================================================

    Dim stDB As String
    Dim xlLoc As String
    Dim ErrorMsg As String
    Dim sRange As String
    Dim sSql As String

    Dim oCatalog As Object
    Dim c As Object
    Dim tblnew As Object
    Dim scn As Object

    Dim idxPrimary As Object

    ' Specific code for 2017
    gbFY17 = (InStr(1, gsPP_Filename, "2017") > 0)

    stDB = gsLocal_Folder + "\" + gwsConfig.Range(gsDB_NAME).Value

    Set oCatalog = CreateObject("ADOX.Catalog")

    On Error Resume Next
    Kill stDB

    On Error GoTo errHandler

    oCatalog.Create "provider='Microsoft.ACE.OLEDB.12.0';" & "Data Source=" & stDB

    Set scn = oCatalog.activeconnection

    If bAllText Then
        '''' IMPORT EVERYTHING AS STRING!!!!
        xlLoc = "'" & sFile & "'[Excel 12.0;HDR=YES;IMEX=1;]"
        sSql = "SELECT * INTO tbl_PortfolioPlan FROM [Portfolio Plan$A3:AG10000] IN " & xlLoc & ";"        'Remove hardcoding!!!!!!!!!!!!!!!!!!!!!

        scn.Execute sSql
        ' Debug.Print sSql

        ' Now the table exists!
        Set tblnew = oCatalog.tables("tbl_PortfolioPlan")

        ' Here I create the Index AFTER running the SQL statement!
        Set idxPrimary = CreateObject("ADOX.Index")
        With idxPrimary
            .Name = "ProjCode"
            ''''.PrimaryKey = True
            .Unique = False
            ''.Columns.Append "Roles"
            .Columns.Append "Project Code"
            .IndexNulls = 0        ' adIndexNullsAllow
            tblnew.Indexes.Append idxPrimary
            Set idxPrimary = Nothing
        End With


    Else
        Set tblnew = CreateObject("ADOX.Table")
        Set c = CreateObject("ADOX.Column")

        ' Create a new Table object.
        With tblnew
            .Name = "tbl_PortfolioPlan"
            ' Create fields and append them to the
            ' Columns collection of the new Table object.
            With .Columns
                .Append "Project Code", adVarWChar
                .Append "Project Name", adVarWChar
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
                ''''''''.Append "Do not remove1", adVarWChar
                .Append "Do not remove1", adDouble
                .Append "Do not remove2", adVarWChar
                .Append "Names", adVarWChar
                .Append "Report", adLongVarWChar
                .Append "Delivery Leader", adVarWChar
                .Append "Account Manager", adVarWChar
                .Append "Activation Status", adVarWChar
                .Append "App", adVarWChar
                If gbFY17 Then
                    .Append "Budget Category", adVarWChar
                Else
                    .Append "Project Type", adVarWChar
                End If
                .Append "Cat", adVarWChar
                .Append "Program", adVarWChar
                .Append "LOB", adVarWChar
                .Append "Project Manager", adVarWChar
                .Append "Publish NE", adVarWChar
                .Append "RowID", adDouble

                If Not gbFY17 Then
                    .Append "MI Labor Cost - Revised BL", adDouble
                    .Append "NE Explanation", adLongVarWChar
                End If

                If gbFY17 Then
                    .Append "Project Priority", adVarWChar
                End If
            End With

            For Each c In .Columns
                c.Attributes = adColNullable
            Next

        End With

        oCatalog.tables.Append tblnew

        sRange = "[" & "Portfolio Plan" & "$A3:AG10000]"

        sSql = "INSERT INTO tbl_PortfolioPlan([Project Code]) "        ' No need to enumerate the rest of the fields...
        sSql = sSql & "SELECT * FROM [Excel 12.0 Macro;HDR=YES;DATABASE=" & sFile & "]." & sRange


        ' Here I create the Index before running the SQL statement!

        Set idxPrimary = CreateObject("ADOX.Index")
        With idxPrimary
            .Name = "ProjCode"
            ''''.PrimaryKey = True
            .Unique = False
            ''.Columns.Append "Roles"
            .Columns.Append "Project Code"
            .IndexNulls = 0        ' adIndexNullsAllow
            tblnew.Indexes.Append idxPrimary
            Set idxPrimary = Nothing
        End With


        scn.Execute sSql
        ' Debug.Print sSql

    End If

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

'---------------------------------------------------------------------------------------
' Method : loadResTable
' Author : cklahr
' Date   : 12/17/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub loadResTable(ByVal sFile As String)
    Dim strQuery As String
    Dim stDB As String
    Dim xlLoc As String
    Dim ErrorMsg As String
    Dim sRange As String
    Dim sSql As String


    Dim oCatalog As Object
    Dim scn As Object



    Set oCatalog = CreateObject("ADOX.Catalog")

    stDB = gsLocal_Folder + "\ResDB.accdb"
    On Error Resume Next
    Kill stDB

    On Error GoTo errHandler

    oCatalog.Create "provider='Microsoft.ACE.OLEDB.12.0';" & "Data Source=" & stDB
    '''''''''sFile = "C:\users\cklahr\Documents\Data\Resource Tracking 2016.xlsm"


    xlLoc = "'" & sFile & "'[Excel 12.0;HDR=YES;IMEX=1;]"        ' '''' IMPORT EVERYTHING AS STRING!!!!

    Set scn = oCatalog.activeconnection
    scn.cursorlocation = 3        'aduseclient


    sSql = "SELECT * INTO tbl_Resources FROM [QA$A3:BL300] IN " & xlLoc & ";"        'Remove hardcoding!!!!!!!!!!!!!!!!!!!!!
    scn.Execute sSql

    sSql = "INSERT INTO tbl_Resources SELECT * FROM [Excel 12.0 Macro;HDR=YES;DATABASE=" & sFile & "].[SAN$A3:BL500]"
    scn.Execute sSql

    sSql = "INSERT INTO tbl_Resources SELECT * FROM [Excel 12.0 Macro;HDR=YES;DATABASE=" & sFile & "].[SA$A3:BL500]"
    scn.Execute sSql

    sSql = "INSERT INTO tbl_Resources SELECT * FROM [Excel 12.0 Macro;HDR=YES;DATABASE=" & sFile & "].[SD$A3:BL1000]"
    scn.Execute sSql

    sSql = "INSERT INTO tbl_Resources SELECT * FROM [Excel 12.0 Macro;HDR=YES;DATABASE=" & sFile & "].[PM$A3:BL500]"
    scn.Execute sSql

    scn.Close

finished:

    Set scn = Nothing
    Set oCatalog = Nothing

    Set scn = Nothing

    Exit Sub

errHandler:
    ErrorMsg = "An error has occurred." & Chr(10) & Chr(10) & Err.Number & " - -" & Err.Description
    MsgBox ErrorMsg, vbCritical, "Error Message - " & ThisWorkbook.Name
    Resume finished

End Sub

