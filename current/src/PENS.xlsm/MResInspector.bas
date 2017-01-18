Attribute VB_Name = "MResInspector"
Option Explicit
Const adVarWChar As Long = 202
Const adLongVarWChar As Long = 203
Const adDouble As Long = 5
Const adDecimal As Long = 7
Const adColNullable As Long = 2

Const adCmdText = 1        ' Required for early binding

Const CSIDL_PERSONAL As Long = &H5

Enum SumType
    byResource = 1
    byproject = 2
    byRole = 3
    bynone = 4
End Enum

Public lSumType As Long

Public glTotRows As Long

Public glFirstRow As Long

Public cmdCommand As Object
Public r As Object 'Recordset

Public gbCompletedFirstLoad As Boolean

Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" _
() '(ByVal hwndOwner As Long, ByVal nFolder As Long, _
ByVal hToken As Long, ByVal dwFlags As Long, _
ByVal pszPath As String) As Long


Public Function openResourceFile() As Workbook
    Dim wbNew As Workbook
    Dim sPathRS As String

    ' If gbUse_Local_Folder Then
    '     sPathRS = gsLocal_Folder + "\" + gsPP_Filename
    ' Else
    '     sPathRS = gsPP_Network_Folder + "\" + gsPP_Filename
    ' End If

    sPathRS = "C:\users\cklahr\Documents\Data\Resource Tracking 2016.xlsm"

    On Error Resume Next

    '' Set wbNew = Workbooks.Open(sPathPP, , True) ''''''''READONLY!!!
    Set wbNew = Workbooks.Open(sPathRS)

    Set openResourceFile = wbNew

End Function

Sub populate_Grid(frm As frmResMan, ByVal SumType As Long)
    Dim sField As String
    Dim sValue As String

    Dim i As Long
    Dim j As Long

    Dim dSum As Double
    Dim lParentRow As Long
    Dim sNextValue As String
    Dim sOldValue As String

    Dim dLastYear As Date

    Dim scn As Object
    Dim cmdCommand As Object

    Dim a As Double
    Dim sDB As String


    sDB = gsLocal_Folder + "\ResDB.accdb"

    Set scn = CreateObject("ADODB.Connection")
    With scn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
        .Open sDB
    End With


    lSumType = SumType

    ' Set the command text.
    Set cmdCommand = CreateObject("ADODB.Command")
    Set cmdCommand.activeconnection = scn

    With cmdCommand
        Select Case SumType
            Case byResource:
                sField = "Resource Name"
                .CommandText = "SELECT * FROM tbl_Resources ORDER BY [" + sField + "]"
            Case byproject:
                sField = "Project Name"
                .CommandText = "SELECT * FROM tbl_Resources ORDER BY [" + sField + "]"
            Case byRole:
                sField = "Role"
                .CommandText = "SELECT * FROM tbl_Resources ORDER BY [" + sField + "],[Resource Name]"
            Case bynone:
                sField = ""
                .CommandText = "SELECT * FROM tbl_Resources ORDER BY [Resource Name],[Project Name]"
        End Select

        .CommandType = adCmdText
        .Execute
    End With


    ' Open the recordset.
    Set r = CreateObject("ADODB.Recordset")
    Set r.activeconnection = scn
    r.Open cmdCommand

    With frmResMan.iGrid1
        .BeginUpdate
        .Clear (True)

        .MultiSelect = True
        .GridlineColor = RGB(191, 191, 191)
        .Font.Name = "Calibri"
        .Font.Size = 10
        .FrozenCols = 3
        .Editable = False
        .BrowseAllNonFrozenCols = True
        .HighlightBackColor = RGB(200, 215, 240)
        .HighlightForeColor = vbBlue

        .DrawRowText = True

        With .Header.Font
            .Name = "Calibri"
            .Size = 10
            .Bold = True
        End With


        Select Case SumType
            Case byResource, bynone:
                .AddCol skey:="Resource Name", sHeader:="Resource Name", lWidth:=150, vTag:="NONE"
                .AddCol skey:="Project Name", sHeader:="Project Name", lWidth:=150, vTag:="NONE"
                .AddCol skey:="Project Code", sHeader:="Project Code", lWidth:=80, vTag:="NONE"
                .AddCol skey:="Role", sHeader:="Role", lWidth:=100, vTag:="NONE"

            Case byproject:
                .AddCol skey:="Project Name", sHeader:="Project Name", lWidth:=250, vTag:="NONE"
                .AddCol skey:="Project Code", sHeader:="Project Code", lWidth:=80, vTag:="NONE"
                .AddCol skey:="Resource Name", sHeader:="Resource Name", lWidth:=150, vTag:="NONE"
                .AddCol skey:="Role", sHeader:="Role", lWidth:=100, vTag:="NONE"

            Case byRole:
                .AddCol skey:="Role", sHeader:="Role", lWidth:=100, vTag:="NONE"
                .AddCol skey:="Resource Name", sHeader:="Resource Name", lWidth:=150, vTag:="NONE"
                .AddCol skey:="Project Name", sHeader:="Project Name", lWidth:=150, vTag:="NONE"
                .AddCol skey:="Project Code", sHeader:="Project Code", lWidth:=80, vTag:="NONE"
        End Select

        '' .ColumnHeaders.Add 4, , "Portfolio", 50

        For i = 12 To r.Fields.Count - 1
            .AddCol sHeader:=Format(r.Fields(i).Name, "dd-mmm"), lWidth:=50, vTag:="FUTURE"        ', eheaderalignh:=igAlignHCenter
        Next i

        .AddCol skey:="CC", sHeader:="CC", lWidth:=40, vTag:="NONE"
        .AddCol skey:="Billable", sHeader:="Billable", lWidth:=50, vTag:="NONE"
        .AddCol skey:="Portfolio", sHeader:="Portfolio", lWidth:=70, vTag:="NONE"
        .AddCol skey:="Resource ID", sHeader:="Resource ID", lWidth:=70, vTag:="NONE"
        .AddCol skey:="F/T or Contract", sHeader:="F/T or Contract", lWidth:=100, vTag:="NONE"
        .AddCol skey:="Lead", sHeader:="Lead", lWidth:=40, vTag:="NONE"
        .AddCol skey:="System", sHeader:="System", lWidth:=100, vTag:="NONE"

        .AddCol skey:="vRow", sHeader:="vRow", lWidth:=50, vTag:="NONE"
        .ColVisible("vRow") = False



        sOldValue = ""

        Do While Not r.EOF
            If Trim(r.Fields("Resource Name")) <> "-" And Trim(r.Fields("Project Name")) <> "Resource Availability" And _
                Trim(r.Fields("Project Name")) <> "Resource Supply" Then
                If SumType <> bynone Then
                    If sOldValue <> UCase(Trim(r.Fields(sField).Value)) Then
                        sOldValue = UCase(Trim(r.Fields(sField).Value))
                        .AddRow
                        .CellValue(.RowCount, .ColIndex("vRow")) = .RowCount

                        lParentRow = .RowCount
                        .RowTreeButton(.RowCount) = igRowTreeButtonVisible

                        Select Case SumType
                            Case byResource:
                                .CellValue(.RowCount, 1) = r.Fields(sField).Value
                                .CellValue(.RowCount, 4) = r.Fields("Role").Value

                                If (r.Fields("CC").Value) <> "" Then .CellValue(.RowCount, .ColIndex("CC")) = Trim(r.Fields("CC").Value)
                                If (r.Fields("Resource ID").Value) <> "" Then .CellValue(.RowCount, .ColIndex("Resource ID")) = Trim(r.Fields("Resource ID").Value)
                                If (r.Fields("F/T or Cont").Value) <> "" Then .CellValue(.RowCount, .ColIndex("F/T or Contract")) = Trim(r.Fields("F/T or Cont").Value)

                            Case byproject:
                                .CellValue(.RowCount, 1) = r.Fields(sField).Value
                                .CellValue(.RowCount, 2) = r.Fields("Project Code").Value


                            Case byRole:
                                .CellValue(.RowCount, 1) = r.Fields(sField).Value
                                If (r.Fields("CC").Value) <> "" Then .CellValue(.RowCount, .ColIndex("CC")) = Trim(r.Fields("CC").Value)


                        End Select

                        .CellFont(.RowCount, 1).Bold = True
                        .CellFont(.RowCount, 2).Bold = True
                        .CellFont(.RowCount, 3).Bold = True
                        .CellFont(.RowCount, 4).Bold = True


                        .CellBackColor(.RowCount, 1) = RGB(98, 235, 232)        ' RGB(192, 192, 192)
                        .CellBackColor(.RowCount, 2) = RGB(98, 235, 232)
                        .CellBackColor(.RowCount, 3) = RGB(98, 235, 232)
                        .CellBackColor(.RowCount, 4) = RGB(98, 235, 232)

                    End If

                    .AddRow
                    .CellValue(.RowCount, .ColIndex("vRow")) = .RowCount
                    .RowLevel(.RowCount) = 1
                Else
                    .AddRow
                    .CellValue(.RowCount, .ColIndex("vRow")) = .RowCount
                End If


                Select Case SumType
                    Case byResource, bynone:
                        .CellValue(.RowCount, 1) = r.Fields("Resource Name").Value
                        .CellValue(.RowCount, 2) = r.Fields("Project Name").Value        'Populate Date column
                        .CellValue(.RowCount, 3) = r.Fields("Project Code").Value
                        .CellValue(.RowCount, 4) = r.Fields("Role").Value        'Populate Usercolumn


                    Case byproject:
                        .CellValue(.RowCount, 1) = r.Fields("Project Name").Value        'Populate Date column
                        .CellValue(.RowCount, 2) = r.Fields("Project Code").Value
                        .CellValue(.RowCount, 3) = r.Fields("Resource Name").Value
                        .CellValue(.RowCount, 4) = r.Fields("Role").Value        'Populate Usercolumn

                    Case byRole:
                        .CellValue(.RowCount, 1) = r.Fields("Role").Value        'Populate Usercolumn
                        .CellValue(.RowCount, 2) = r.Fields("Resource Name").Value
                        .CellValue(.RowCount, 3) = r.Fields("Project Name").Value        'Populate Date column
                        .CellValue(.RowCount, 4) = r.Fields("Project Code").Value
                End Select

                .CellBackColor(.RowCount, 1) = RGB(220, 250, 250)
                .CellBackColor(.RowCount, 2) = RGB(220, 250, 250)
                .CellBackColor(.RowCount, 3) = RGB(220, 250, 250)
                .CellBackColor(.RowCount, 4) = RGB(220, 250, 250)



                ' Safest place to tag past columns!
                For i = 12 To r.Fields.Count - 1
                    If Date > CDate(.ColHeaderText(5 + i - 12)) Then
                        .CellBackColor(.RowCount, 5 + i - 12) = RGB(226, 226, 226)
                        .ColTag(5 + i - 12) = "PAST"
                    ElseIf Month(CDate(.ColHeaderText(5 + i - 12))) = 12 And i < 16 Then
                        dLastYear = DateAdd("yyyy", -1, CDate(.ColHeaderText(5 + i - 12)))
                        If Date > dLastYear Then
                            .CellBackColor(.RowCount, 5 + i - 12) = RGB(226, 226, 226)
                            .ColTag(5 + i - 12) = "PAST"
                        Else
                            .CellBackColor(.RowCount, 5 + i - 12) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                        End If
                    End If

                    If (r.Fields(i).Value <> "") Then
                        .CellValue(.RowCount, 5 + i - 12) = Format(r.Fields(i).Value, "#0.00")
                    Else
                        .CellValue(.RowCount, 5 + i - 12) = ""
                    End If
                Next i


                ' There are now 7 extra columns at the very right of the table
                .CellBackColor(.RowCount, 5 + i - 12) = RGB(220, 250, 250)
                .CellBackColor(.RowCount, 5 + i - 11) = RGB(220, 250, 250)
                .CellBackColor(.RowCount, 5 + i - 10) = RGB(220, 250, 250)
                .CellBackColor(.RowCount, 5 + i - 9) = RGB(220, 250, 250)
                .CellBackColor(.RowCount, 5 + i - 8) = RGB(220, 250, 250)
                .CellBackColor(.RowCount, 5 + i - 7) = RGB(220, 250, 250)
                .CellBackColor(.RowCount, 5 + i - 6) = RGB(220, 250, 250)

                If (r.Fields("CC").Value) <> "" Then .CellValue(.RowCount, .ColIndex("CC")) = Trim(r.Fields("CC").Value)
                If (r.Fields("Billable").Value) <> "" Then .CellValue(.RowCount, .ColIndex("Billable")) = Trim(r.Fields("Billable").Value)
                If (r.Fields("Portfolio").Value) <> "" Then .CellValue(.RowCount, .ColIndex("Portfolio")) = Trim(r.Fields("Portfolio").Value)
                If (r.Fields("Resource ID").Value) <> "" Then .CellValue(.RowCount, .ColIndex("Resource ID")) = Trim(r.Fields("Resource ID").Value)
                If (r.Fields("F/T or Cont").Value) <> "" Then .CellValue(.RowCount, .ColIndex("F/T or Contract")) = Trim(r.Fields("F/T or Cont").Value)
                If (r.Fields("Lead?").Value) <> "" Then .CellValue(.RowCount, .ColIndex("Lead")) = Trim(r.Fields("Lead?").Value)
                If (r.Fields("System").Value) <> "" Then .CellValue(.RowCount, .ColIndex("System")) = Trim(r.Fields("System").Value)


                r.MoveNext
            Else
                r.MoveNext
            End If
        Loop

        .CollapseAllRows

        If SumType <> bynone Then
            For lParentRow = 1 To .RowCount
                If .RowLevel(lParentRow) = 0 Then
                    For i = 12 To r.Fields.Count - 1
                        dSum = 0
                        For j = 1 To .RowChildCount(lParentRow)
                            If IsNumeric(.CellValue(lParentRow + j, 5 + i - 12)) Then
                                dSum = dSum + .CellValue(lParentRow + j, 5 + i - 12)
                            End If
                        Next

                        .CellValue(lParentRow, 5 + i - 12) = Format(dSum, "#0.00")
                        .CellFont(lParentRow, 5 + i - 12).Bold = True

                        If Date > CDate(.ColHeaderText(5 + i - 12)) Then
                            .CellBackColor(lParentRow, 5 + i - 12) = RGB(180, 180, 180)
                        ElseIf Month(CDate(.ColHeaderText(5 + i - 12))) = 12 And i < 16 Then
                            dLastYear = DateAdd("yyyy", -1, CDate(.ColHeaderText(5 + i - 12)))
                            If Date > dLastYear Then
                                .CellBackColor(lParentRow, 5 + i - 12) = RGB(180, 180, 180)
                            Else
                                .CellBackColor(lParentRow, 5 + i - 12) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                            End If
                        Else
                            .CellBackColor(lParentRow, 5 + i - 12) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                        End If
                    Next
                    ' There are now 7 extra columns at the very right of the table
                    .CellBackColor(lParentRow, 5 + i - 12) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                    .CellBackColor(lParentRow, 5 + i - 11) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                    .CellBackColor(lParentRow, 5 + i - 10) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                    .CellBackColor(lParentRow, 5 + i - 9) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                    .CellBackColor(lParentRow, 5 + i - 8) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                    .CellBackColor(lParentRow, 5 + i - 7) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                    .CellBackColor(lParentRow, 5 + i - 6) = RGB(98, 235, 232)        'RGB(192, 192, 192)
                End If
            Next
        End If


        For i = 1 To .ColCount
            If .ColTag(i) = "PAST" Then
                If frm.chkHidePast.Value Then
                    .ColVisible(i) = False
                Else
                    .ColVisible(i) = True
                End If
            End If
        Next i

        .EndUpdate

        glTotRows = .RowCount

        .CellSelected(1, 1) = True


        frm.lblRowNumber.Caption = Str(.CellValue(1, .ColIndex("vRow"))) + "/" + Str(glTotRows)
        frm.lblSelectedCount.Caption = 1
        frm.lblSelectedFTE.Caption = 0
        frm.lblSelectedCost.Caption = Format(0, "$#,##0.0")

    End With


    r.Close
    scn.Close

    Set r = Nothing
    Set cmdCommand = Nothing
    Set scn = Nothing

End Sub

Sub populate_combos(frm As frmResMan)
    Dim scn As Object
    Dim cmdCommand As Object
    Dim r As Object 'Recordset


    Dim sDB As String


    sDB = MyDocuments() + "\ResDB.accdb"

    Set scn = CreateObject("ADODB.Connection")
    With scn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
        .Open sDB
    End With


    ' Set the command text.
    Set cmdCommand = CreateObject("ADODB.Command")
    Set cmdCommand.activeconnection = scn

    With cmdCommand
        .CommandText = "SELECT DISTINCT [F/T or Cont] FROM tbl_Resources"
        .CommandType = adCmdText
        .Execute
    End With


    ' Open the recordset.
    Set r = CreateObject("ADODB.Recordset")
    Set r.activeconnection = scn
    r.Open cmdCommand


    With frm
        .cmbResStatus.AddItem "ALL"

        Do While Not r.EOF
            If Trim(r.Fields("F/T or Cont") <> "") And Trim(r.Fields("F/T or Cont") <> "-") Then .cmbResStatus.AddItem Trim(r.Fields("F/T or Cont"))

            r.MoveNext
        Loop

        .cmbResStatus.ListIndex = 0
    End With

    '----------------
    With cmdCommand
        .CommandText = "SELECT DISTINCT [CC] FROM tbl_Resources"
        .CommandType = adCmdText
        .Execute
    End With


    ' Open the recordset.
    Set r = CreateObject("ADODB.Recordset")
    Set r.activeconnection = scn
    r.Open cmdCommand


    With frm
        .cmbCC.AddItem "ALL"
        Do While Not r.EOF
            If Trim(r.Fields("CC") <> "") And Trim(r.Fields("CC") <> "-") Then .cmbCC.AddItem Trim(r.Fields("CC"))
            r.MoveNext
        Loop

        .cmbCC.ListIndex = 0
    End With

    '----------------
    With cmdCommand
        .CommandText = "SELECT DISTINCT [Portfolio] FROM tbl_Resources"
        .CommandType = adCmdText
        .Execute
    End With


    ' Open the recordset.
    Set r = CreateObject("ADODB.Recordset")
    Set r.activeconnection = scn
    r.Open cmdCommand


    With frm
        .cmbPortfolio.AddItem "ALL"
        Do While Not r.EOF
            If Trim(r.Fields("Portfolio") <> "") And Trim(r.Fields("Portfolio") <> "-") Then .cmbPortfolio.AddItem Trim(r.Fields("Portfolio"))
            r.MoveNext
        Loop

        .cmbPortfolio.ListIndex = 0
    End With




    Set r = Nothing
    Set cmdCommand = Nothing
    Set scn = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------

Sub Access_Resources()
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
    Dim scn As Object
    Dim c As Object
    Dim tblnew As Object

    Dim sFile As String


    '''Dim cmdCommand As Object
    ''''Dim rstRecordset As Object

    stDB = MyDocuments() + "\ResDB.accdb"        'gwsConfig.Range(gsDB_NAME).Value

    Set oCatalog = CreateObject("ADOX.Catalog")

    On Error Resume Next
    Kill stDB

    On Error GoTo errHandler

    oCatalog.Create "provider='Microsoft.ACE.OLEDB.12.0';" & "Data Source=" & stDB
    sFile = "C:\users\cklahr\Documents\Data\Resource Tracking 2016.xlsm"


    xlLoc = "'" & sFile & "'[Excel 12.0;HDR=YES;IMEX=1;]"        '    '''' IMPORT EVERYTHING AS STRING!!!!

    Set scn = oCatalog.activeconnection
    scn.cursorlocation = 3        'aduseclient

    Load frmResMan
    frmResMan.Show False


    gbCompletedFirstLoad = False

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

    Call populate_Grid(frmResMan, bynone)

    Call populate_combos(frmResMan)


    'sSql = "DROP TABLE tbl_Resources;"
    'scn.Execute sSql




    frmResMan.optSumNone.Value = True



    gbCompletedFirstLoad = True



    frmResMan.iGrid1.SetFocus




finished:
    ''Set cmdCommand = Nothing
    '''Set r = Nothing

    ''Set scn = Nothing
    ''Set oCatalog = Nothing
    ''Set c = Nothing
    ''Set tblnew = Nothing


    Exit Sub

errHandler:
    ErrorMsg = "An error has occurred." & Chr(10) & Chr(10) & Err.Number & " - -" & Err.Description
    MsgBox ErrorMsg, vbCritical, "Error Message - " & ThisWorkbook.Name
    Resume finished

End Sub





