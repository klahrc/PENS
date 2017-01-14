Attribute VB_Name = "MGenPLD"
'---------------------------------------------------------------------------------------
' File   : MGenPLD
' Author : cklahr
' Date   : 2/11/2016
' Purpose: Generates NE Summary
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------
Private Const msMODULE As String = "MGenPLD"

Const adCmdText = 1        ' Required for early binding

Private Function FetchValue(cn As Object, sRetField As String, sCondField1 As String, sCondition1 As String, sCondField2 As String, sCondition2 As String) As String
    Dim cmdCommand As Object
    Set cmdCommand = CreateObject("ADODB.Command")

    Set cmdCommand.activeconnection = cn
    With cmdCommand
        .CommandText = "SELECT [" + sRetField + "] " + "FROM tbl_PortfolioPlan WHERE [" + sCondField1 + "] = " + "'" + sCondition1 + "'" + " AND [" + sCondField2 + "] = " + "'" + sCondition2 + "'"
        .CommandType = adCmdText
        .Execute
    End With


    ' Open the recordset.
    Dim rstRecordset As Object
    Set rstRecordset = CreateObject("ADODB.Recordset")
    Set rstRecordset.activeconnection = cn
    rstRecordset.Open cmdCommand

    Dim a As String
    a = rstRecordset.Fields(sRetField)

    FetchValue = a
End Function

'---------------------------------------------------------------------------------------
' Method : FirstLoad
' Author : cklahr
' Date   : 2/15/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub FirstLoad()
    Dim wbNew As Workbook
    Dim sPPDB As String
    Dim cnConn As Object
    Dim cmdCommand As Object
    Dim rstRecordset As Object
    Dim CellName As String
    Dim nm As Name
    Dim i As Long
    Dim kk As String
    Dim sPLDFileName As String


    sPLDFileName = "PLD_" + Format(Now(), "mmmddyyyy_hhmmss") + ".xlsm"


    Set wbNew = Workbooks.Add(ThisWorkbook.Path & "\" & "PLD.xltm")
    wbNew.Activate


    sPPDB = gsLocal_Folder + "\" + wbNew.Sheets("DATA").Range("O1").Value        ''''' HARDCODED!!!!!!!!!!

    Set cnConn = CreateObject("ADODB.Connection")
    With cnConn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
        .Open sPPDB
    End With


    'Loop through each name and delete
    For Each nm In wbNew.Names
        wbNew.Sheets("DATA").Range(nm).Clear
        nm.Delete
    Next

    ' Set the command text.
    Set cmdCommand = CreateObject("ADODB.Command")
    Set cmdCommand.activeconnection = cnConn
    With cmdCommand
        .CommandText = "SELECT DISTINCT [LOB]FROM tbl_PortfolioPlan"
        .CommandType = adCmdText
        .Execute
    End With

    ' Open the recordset.
    Set rstRecordset = CreateObject("ADODB.Recordset")
    Set rstRecordset.activeconnection = cnConn
    rstRecordset.Open cmdCommand

    i = 15
    With wbNew.Sheets("DATA")
        While Not rstRecordset.EOF
            If rstRecordset.Fields("LOB") <> "" And rstRecordset.Fields("LOB") <> "0" Then
                .Cells(i, 2) = rstRecordset.Fields("LOB")
                i = i + 1
            Else
                ''''' Loggear en debug!!!!
            End If
            rstRecordset.MoveNext
        Wend
        i = i - 1

        rstRecordset.Close

        'Range of Cells Reference (Workbook Scope)

        CellName = "B15:B" & CStr(i)

        Dim cell As Range

        Set cell = wbNew.Sheets("DATA").Range(CellName)
        wbNew.Names.Add Name:="LOBList", RefersTo:=cell


        kk = .Range("B15").Value


        With cmdCommand
            .CommandText = "SELECT DISTINCT[Project Code], [Project Name] FROM tbl_PortfolioPlan WHERE [LOB] = " & "'" & "1117  MI - Shareholder’s Services" & "'"
            '''''.CommandText = "SELECT DISTINCT[Project Code], [Project Name] FROM tbl_PortfolioPlan WHERE [LOB] = " & "'" & kk & "'"


            ''''.CommandText = "SELECT DISTINCT[Project Code], [Project Name] FROM tbl_PortfolioPlan WHERE [LOB] =" & "'" & "1112  MI - Corporate" & "'"
            .CommandType = adCmdText
            .Execute
        End With

        rstRecordset.Open cmdCommand


        i = 15

        While Not rstRecordset.EOF
            .Cells(i, 6) = rstRecordset.Fields("Project Code")
            .Cells(i, 8) = rstRecordset.Fields("Project Code") + " - " + rstRecordset.Fields("Project Name")
            i = i + 1

            rstRecordset.MoveNext
        Wend
        i = i - 1

        rstRecordset.Close


        'Range of Cells Reference (Workbook Scope)

        CellName = "F15:F" & CStr(i)
        Set cell = wbNew.Sheets("DATA").Range(CellName)
        wbNew.Names.Add Name:="PRJCodesList", RefersTo:=cell


        CellName = "H15:H" & CStr(i)
        Set cell = wbNew.Sheets("DATA").Range(CellName)
        wbNew.Names.Add Name:="PRJList", RefersTo:=cell


        kk = .Range("F15").Value

        With cmdCommand
            .CommandText = "SELECT [JAN], [FEB], [MAR], [APR], [MAY], [Jun], [Jul], [Aug], [Sep], [Oct], [Nov], [Dec] FROM tbl_PortfolioPlan WHERE [Project Code] = " & "'" & kk & "'" & _
            " And [Roles] =" & "'" & "NE (Actual + Rem. Plan)" & "'"
            .CommandType = adCmdText
            .Execute
        End With

        rstRecordset.Open cmdCommand

        '.Cells(7, 3).Value = rstRecordset.fields("JAN")
        '.Cells(7, 4).Value = rstRecordset.fields("FEB")
        '.Cells(7, 5).Value = rstRecordset.fields("MAR")
        '.Cells(7, 6).Value = rstRecordset.fields("APR")
        '.Cells(7, 7).Value = rstRecordset.fields("APR")
        '.Cells(7, 8).Value = rstRecordset.fields("MAY")
        '.Cells(7, 9).Value = rstRecordset.fields("JUN")
        '.Cells(7, 10).Value = rstRecordset.fields("JUL")
        '.Cells(7, 11).Value = rstRecordset.fields("AUG")
        '.Cells(7, 12).Value = rstRecordset.fields("SEP")
        '.Cells(7, 13).Value = rstRecordset.fields("OCT")
        '.Cells(7, 14).Value = rstRecordset.fields("NOV")
        '.Cells(7, 15).Value = rstRecordset.fields("DEC")

        rstRecordset.Close

        cnConn.Close
    End With

    wbNew.Sheets("PLD").Activate


    Set cmdCommand = Nothing
    Set rstRecordset = Nothing
    Set cnConn = Nothing


    wbNew.SaveAs FileName:=gwsConfig.Range(gsLocal_Folder).Value & "\" & sPLDFileName, FileFormat:=52

End Sub
