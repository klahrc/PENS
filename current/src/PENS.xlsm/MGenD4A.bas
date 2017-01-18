Attribute VB_Name = "MGenD4A"
'---------------------------------------------------------------------------------------
' File   : MGenD4A
' Author : cklahr
' Date   : 1/17/2016
' Purpose: Module responsible of generating the Pivot Tables
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Module Constant Declarations Follow
'---------------------------------------------------------------------------------------

Private Const adCmdText As Integer = 1                     ' Required for early binding
Private Const pos As Integer = 1
Private Const NEG As Integer = 2

Private Const MINPOSVAR As Double = 25000
Private Const MINNEGVAR As Double = -25000

Private Const msMODULE As String = "MGenD4A"

Sub CondFormat(r As Range)
    Dim cs As ColorScale

    With r
        Set cs = .FormatConditions.AddColorScale(colorscaletype:=3)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority ' Take priority over any other formats

        ' Set the color of the lowest value, with a range up to the next scale criteria. The color should be white.
        With cs.ColorScaleCriteria(1)
            .Type = xlConditionValueLowestValue
            With .FormatColor
                .Color = RGB(248, 105, 107)
                .TintAndShade = 0
            End With
        End With

        ' At the 50th percentile, the color should be orange/green.
        ' Note that you can't set the Value property for all values of Type.
        With cs.ColorScaleCriteria(2)
            '''''''.Type = xlConditionValuePercentile
            .Type = xlConditionValueNumber
            ''''''''.Value = 50
            .Value = 0
            With .FormatColor
                ''''.Color = RGB(255, 235, 132)
                .Color = RGB(255, 255, 255)
                .TintAndShade = 0
            End With
        End With

        ' At the highest value, the color should be orange.
        With cs.ColorScaleCriteria(3)
            .Type = xlConditionValueHighestValue
            With .FormatColor
                .Color = RGB(99, 190, 123)
                .TintAndShade = 0
            End With
        End With
    End With
End Sub

'---------------------------------------------------------------------------------------
' Method : GenPT_PS
' Author : cklahr
' Date   :
' Purpose: Generates Portfolio Seasonality
'---------------------------------------------------------------------------------------
Function GenPT_PS(ByVal wbNew As Workbook, ByVal PC As PivotCache) As Double
    Dim wsPT As Worksheet
    Dim PT As PivotTable
    Dim r As Range
    Dim cs As ColorScale
    Dim dGrandTotal As Double
    Dim i As Integer

    Const sSOURCE As String = "GenPT_PS"

    'wbNew is the D4A workbook...
    Set wsPT = wbNew.Sheets.Add

    wbNew.Windows(1).DisplayGridlines = False

    Set PT = PC.CreatePivotTable(TableDestination:=wbNew.ActiveSheet.Name & "!R1c1", TableName:=wbNew.ActiveSheet.Name, DefaultVersion:=xlPivotTableVersion12)


    PT.PivotFields("Cat").Orientation = xlRowField
    PT.PivotFields("Cat").Subtotals(1) = False             ' 1 = Automatic


    ' Remove Cat = "0"
    ' Dim pi As PivotItem
    ' For Each pi In PT.PivotFields("Cat").PivotItems
    '     If pi.Name = "0" Then pi.Visible = False
    ' Next


    PT.PivotFields("Project Name").Orientation = xlRowField
    PT.PivotFields("Project Name").Subtotals(1) = False
    PT.PivotFields("Project Name").AutoSort Order:=xlDescending, Field:="Program"

    PT.AddDataField PT.PivotFields("JAN"), " JAN", xlSum
    PT.AddDataField PT.PivotFields("FEB"), " FEB", xlSum
    PT.AddDataField PT.PivotFields("MAR"), " MAR", xlSum
    PT.AddDataField PT.PivotFields("APR"), " APR", xlSum
    PT.AddDataField PT.PivotFields("MAY"), " MAY", xlSum
    PT.AddDataField PT.PivotFields("JUN"), " JUN", xlSum
    PT.AddDataField PT.PivotFields("JUL"), " JUL", xlSum
    PT.AddDataField PT.PivotFields("AUG"), " AUG", xlSum
    PT.AddDataField PT.PivotFields("SEP"), " SEP", xlSum
    PT.AddDataField PT.PivotFields("OCT"), " OCT", xlSum
    PT.AddDataField PT.PivotFields("NOV"), " NOV", xlSum
    PT.AddDataField PT.PivotFields("DEC"), " DEC", xlSum

    PT.PivotFields("Roles").Orientation = xlPageField
    PT.PivotFields("Roles").ClearAllFilters
    PT.PivotFields("Roles").CurrentPage = "Total Plan"

    PT.PivotFields("Project Priority").Orientation = xlPageField
    PT.PivotFields("Project Priority").ClearAllFilters
    PT.PivotFields("Project Priority").CurrentPage = "1 - High"

    PT.PivotFields("Do not remove1").Orientation = xlPageField
    PT.PivotFields("Do not remove1").ClearAllFilters
    PT.PivotFields("Do not remove1").PivotItems("0").Visible = False
    PT.PivotFields("Do not remove1").PivotItems("(blank)").Visible = False

    PT.DataBodyRange.NumberFormat = "#,###"                '
    PT.DataLabelRange.HorizontalAlignment = xlRight

    PT.DataPivotField.Value = ""                           'Remove the word  from value column
    If gbFY17 Then
        PT.CompactLayoutRowHeader = "Total Plan 2017"      ' Changes the default 
    Else
        PT.CompactLayoutRowHeader = "Total Plan 2016"      ' Changes the default 
    End If

    wsPT.Range("M1:M2").Merge
    With wsPT.Range("M1")
        If gbFY17 Then
            .Value = 2017
        Else
            .Value = 2016
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(79, 98, 40)
        .Font.Color = vbYellow
    End With


    Set r = wsPT.Range(wsPT.Cells(5, 1), wsPT.Cells(PT.RowRange.Count + 2, 13))

    With r
        Set cs = .FormatConditions.AddColorScale(colorscaletype:=3)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority ' Take priority over any other formats

        ' Set the color of the lowest value, with a range up to the next scale criteria. The color should be white.
        With cs.ColorScaleCriteria(1)
            '''''''''''.Type = xlConditionValueLowestValue
            .Type = xlConditionValueNumber
            .Value = 0
            With .FormatColor
                .Color = RGB(255, 255, 255)
                .TintAndShade = 0
            End With
        End With

        ' At the 50th percentile, the color should be orange/green.
        ' Note that you can't set the Value property for all values of Type.
        With cs.ColorScaleCriteria(2)
            .Type = xlConditionValuePercentile
            .Value = 50
            With .FormatColor
                .Color = RGB(255, 243, 183)
                .TintAndShade = 0
            End With
        End With

        ' At the highest value, the color should be orange.
        With cs.ColorScaleCriteria(3)
            .Type = xlConditionValueHighestValue
            With .FormatColor
                .Color = RGB(255, 192, 0)
                .TintAndShade = 0
            End With
        End With
    End With

    wsPT.Range("B:M").ColumnWidth = 11
    wsPT.Range("$A6").HorizontalAlignment = xlCenter

    PT.DataBodyRange.Font.Color = RGB(255, 220, 100)

    dGrandTotal = 0
    For i = 2 To 13
        dGrandTotal = dGrandTotal + wsPT.Cells(PT.RowRange.Count + 3, i)
    Next

    wsPT.Name = "Plan Seasonality"
    GenPT_PS = dGrandTotal

ErrorExit:
    ' Clean up
    Exit Function

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

'---------------------------------------------------------------------------------------
' Method : GenPT_FTE_NE
' Author : cklahr
' Date   : 1/17/2016
' Purpose: Generates FTE_NE PT
'---------------------------------------------------------------------------------------
Function GenPT_FTE_NE(ByVal wbNew As Workbook, ByVal PC As PivotCache) As Double
    Dim wsPT As Worksheet
    Dim PT As PivotTable
    Dim r As Range
    Dim cs As ColorScale
    Dim dGrandTotal As Double
    Dim i As Integer

    Const sSOURCE As String = "GenPT_FTE_NE"

    'wbNew is the D4A workbook...
    Set wsPT = wbNew.Sheets.Add

    wbNew.Windows(1).DisplayGridlines = False

    Set PT = PC.CreatePivotTable(TableDestination:=wbNew.ActiveSheet.Name & "!R1c1", TableName:=wbNew.ActiveSheet.Name, DefaultVersion:=xlPivotTableVersion12)


    PT.PivotFields("Cat").Orientation = xlRowField
    PT.PivotFields("Cat").Subtotals(1) = False             ' 1 = Automatic


    ' Remove Cat = "0"
    ' Dim pi As PivotItem
    ' For Each pi In PT.PivotFields("Cat").PivotItems
    '     If pi.Name = "0" Then pi.Visible = False
    ' Next


    PT.PivotFields("Program").Orientation = xlRowField
    PT.PivotFields("Program").Subtotals(1) = False
    PT.PivotFields("Program").AutoSort Order:=xlDescending, Field:="Program"

    PT.AddDataField PT.PivotFields("JAN"), " JAN", xlSum
    PT.AddDataField PT.PivotFields("FEB"), " FEB", xlSum
    PT.AddDataField PT.PivotFields("MAR"), " MAR", xlSum
    PT.AddDataField PT.PivotFields("APR"), " APR", xlSum
    PT.AddDataField PT.PivotFields("MAY"), " MAY", xlSum
    PT.AddDataField PT.PivotFields("JUN"), " JUN", xlSum
    PT.AddDataField PT.PivotFields("JUL"), " JUL", xlSum
    PT.AddDataField PT.PivotFields("AUG"), " AUG", xlSum
    PT.AddDataField PT.PivotFields("SEP"), " SEP", xlSum
    PT.AddDataField PT.PivotFields("OCT"), " OCT", xlSum
    PT.AddDataField PT.PivotFields("NOV"), " NOV", xlSum
    PT.AddDataField PT.PivotFields("DEC"), " DEC", xlSum

    PT.PivotFields("Roles").Orientation = xlPageField
    PT.PivotFields("Roles").ClearAllFilters
    PT.PivotFields("Roles").CurrentPage = "FTE NE - MI"

    PT.DataBodyRange.NumberFormat = "###.#"                '
    PT.DataLabelRange.HorizontalAlignment = xlRight

    wsPT.Range("M1:M2").Merge
    With wsPT.Range("M1")
        If gbFY17 Then
            .Value = 2017
        Else
            .Value = 2016
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(79, 98, 40)
        .Font.Color = vbYellow
    End With


    Set r = wsPT.Range(wsPT.Cells(5, 1), wsPT.Cells(PT.RowRange.Count + 2, 13))

    With r
        Set cs = .FormatConditions.AddColorScale(colorscaletype:=3)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority ' Take priority over any other formats

        ' Set the color of the lowest value, with a range up to the next scale criteria. The color should be white.
        With cs.ColorScaleCriteria(1)
            '''''''''''.Type = xlConditionValueLowestValue
            .Type = xlConditionValueNumber
            .Value = 0
            With .FormatColor
                .Color = RGB(255, 255, 255)
                .TintAndShade = 0
            End With
        End With

        ' At the 50th percentile, the color should be orange/green.
        ' Note that you can't set the Value property for all values of Type.
        With cs.ColorScaleCriteria(2)
            .Type = xlConditionValuePercentile
            .Value = 50
            With .FormatColor
                .Color = RGB(234, 241, 221)
                .TintAndShade = 0
            End With
        End With

        ' At the highest value, the color should be orange.
        With cs.ColorScaleCriteria(3)
            .Type = xlConditionValueHighestValue
            With .FormatColor
                .Color = RGB(255, 192, 0)
                .TintAndShade = 0
            End With
        End With
    End With

    wsPT.Range("B:M").ColumnWidth = 5

    dGrandTotal = 0
    For i = 2 To 13
        dGrandTotal = dGrandTotal + wsPT.Cells(PT.RowRange.Count + 3, i)
    Next

    wsPT.Name = "FTE NE"
    GenPT_FTE_NE = dGrandTotal

ErrorExit:
    ' Clean up
    Exit Function

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

'---------------------------------------------------------------------------------------
' Method : GenPT_PROJ_SUMMARY
' Author : cklahr
' Date   : 1/17/2016
' Purpose: Generates PROJ_SUMMARY PT
'---------------------------------------------------------------------------------------
Function GenPT_PROJ_SUMMARY(ByVal wbNew As Workbook, ByVal PC As PivotCache) As Double

    Dim wsPT As Worksheet
    Dim PT As PivotTable
    Dim r As Range
    Dim cs As ColorScale
    Dim dGrandTotal As Double
    Dim i As Integer

    Const sSOURCE As String = "GenPT_PROJ_SUMMARY"

    'wbNew is the D4A workbook...
    Set wsPT = wbNew.Sheets.Add
    wbNew.Windows(1).DisplayGridlines = False

    Set PT = PC.CreatePivotTable(TableDestination:=ActiveSheet.Name & "!R1c1", TableName:=ActiveSheet.Name, DefaultVersion:=xlPivotTableVersion12)

    PT.PivotFields("Cat").Orientation = xlRowField
    PT.PivotFields("Cat").Subtotals(1) = False             ' 1 = Automatic


    ' Remove Cat = "0"
    ' Dim pi As PivotItem
    ' For Each pi In PT.PivotFields("Cat").PivotItems
    '     If pi.Name = "0" Then pi.Visible = False
    ' Next


    PT.PivotFields("Project Name").Orientation = xlRowField
    PT.PivotFields("Project Name").Subtotals(1) = False
    PT.PivotFields("Project Name").AutoSort Order:=xlDescending, Field:="Program"

    PT.AddDataField PT.PivotFields("JAN"), " JAN", xlSum
    PT.AddDataField PT.PivotFields("FEB"), " FEB", xlSum
    PT.AddDataField PT.PivotFields("MAR"), " MAR", xlSum
    PT.AddDataField PT.PivotFields("APR"), " APR", xlSum
    PT.AddDataField PT.PivotFields("MAY"), " MAY", xlSum
    PT.AddDataField PT.PivotFields("JUN"), " JUN", xlSum
    PT.AddDataField PT.PivotFields("JUL"), " JUL", xlSum
    PT.AddDataField PT.PivotFields("AUG"), " AUG", xlSum
    PT.AddDataField PT.PivotFields("SEP"), " SEP", xlSum
    PT.AddDataField PT.PivotFields("OCT"), " OCT", xlSum
    PT.AddDataField PT.PivotFields("NOV"), " NOV", xlSum
    PT.AddDataField PT.PivotFields("DEC"), " DEC", xlSum

    PT.PivotFields("Roles").Orientation = xlPageField
    PT.PivotFields("Roles").ClearAllFilters
    PT.PivotFields("Roles").CurrentPage = "FTE NE - MI"

    PT.DataBodyRange.NumberFormat = "###.#"                '
    PT.DataLabelRange.HorizontalAlignment = xlRight

    wsPT.Range("M1:M2").Merge
    With wsPT.Range("M1")
        If gbFY17 Then
            .Value = 2017
        Else
            .Value = 2016
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(79, 98, 40)
        .Font.Color = vbYellow
    End With


    ''''Set r = wsPT.Range(wsPT.Cells(5, 1), wsPT.Cells(46, 13))

    Set r = wsPT.Range(wsPT.Cells(5, 1), wsPT.Cells(PT.RowRange.Count + 2, 13))


    With r
        Set cs = .FormatConditions.AddColorScale(colorscaletype:=3)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority ' Take priority over any other formats

        ' Set the color of the lowest value, with a range up to the next scale criteria. The color should be white.
        With cs.ColorScaleCriteria(1)
            '''''''''''''''''.Type = xlConditionValueLowestValue

            .Type = xlConditionValueNumber
            .Value = 0

            With .FormatColor
                .Color = RGB(255, 255, 255)
                .TintAndShade = 0
            End With
        End With

        ' At the 50th percentile, the color should be orange/green.
        ' Note that you can't set the Value property for all values of Type.
        With cs.ColorScaleCriteria(2)
            .Type = xlConditionValuePercentile
            .Value = 50
            With .FormatColor
                .Color = RGB(234, 241, 221)
                .TintAndShade = 0
            End With
        End With

        ' At the highest value, the color should be orange.
        With cs.ColorScaleCriteria(3)
            .Type = xlConditionValueHighestValue
            With .FormatColor
                .Color = RGB(255, 192, 0)
                .TintAndShade = 0
            End With
        End With
    End With

    wsPT.Range("B:M").ColumnWidth = 5

    dGrandTotal = 0
    For i = 2 To 13
        dGrandTotal = dGrandTotal + wsPT.Cells(PT.RowRange.Count + 3, i)
    Next

    wsPT.Name = "Proj Summary"


    GenPT_PROJ_SUMMARY = dGrandTotal

ErrorExit:
    ' Clean up
    Exit Function

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

'---------------------------------------------------------------------------------------
' Method : getNumRolesAndPurge
' Author : cklahr
' Date   : 11/16/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function getNumRolesAndPurge(ByVal ws As Worksheet, ByVal iMaxRoles As Integer) As Integer
    Dim i As Integer
    Dim j As Integer

    With ws
        For i = 1 To iMaxRoles
            If UCase(Left(.Cells(i + 2, 2), 3)) = "FTE" Then
                Exit For
            End If
        Next i
    End With

    getNumRolesAndPurge = i - 2

    ' Delete unused rows
    ws.Rows(CStr(i + 2) + ":" & ws.UsedRange.Count).Delete
End Function
'---------------------------------------------------------------------------------------
' Method : GenPT_CC_CAPACITY
' Author : cklahr
' Date   : 1/17/2016
' Purpose: Generates CC_CAPACITY PT
' Assumptions:
'           1. There are no unused roles in the Capacity table (roles with no demand at all)
'           2. Role names in Dashboard and Resource spreadsheet match perfectly
'           3. Tables T1, T2, T3, T4 include Headers and Total row at the bottom
'---------------------------------------------------------------------------------------
Function GenPT_CC_CAPACITY(sRemFolderRS As String, sRemFileNameRS As String) As Double
    Dim wbNew As Workbook
    Dim wsPT As Worksheet
    Dim PT As PivotTable
    Dim r As Range
    Dim dGrandTotal As Double
    Dim i As Integer

    Dim iRolesCount As Integer

    Dim lT1x1 As Long
    Dim lT1x2 As Long
    Dim lT1y1 As Long
    Dim lT1y2 As Long

    Dim lT2x1 As Long
    Dim lT2x2 As Long
    Dim lT2y1 As Long
    Dim lT2y2 As Long

    Dim lT3x1 As Long
    Dim lT3x2 As Long
    Dim lT3y1 As Long
    Dim lT3y2 As Long

    Dim lT4x1 As Long
    Dim lT4x2 As Long
    Dim lT4y1 As Long
    Dim lT4y2 As Long

    Dim sRemoteCCTable As String
    Dim sRemData As String
    Dim sRepFilename As String

    Dim pi As PivotItem
    Dim rngFindValue As Range

    Dim PC As PivotCache


    Const lMAX_ROLES As Integer = 100                      'Max number of roles

    Const sSOURCE As String = "GenPT_CC_CAPACITY"






    Dim sPPDB As String
    sPPDB = gsLocal_Folder + "\" + gwsConfig.Range(gsDB_NAME).Value

    Dim cnConn As Object
    Set cnConn = CreateObject("ADODB.Connection")
    With cnConn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
        .Open sPPDB
    End With


    ' Set the command text.
    Dim cmdCommand As Object
    Set cmdCommand = CreateObject("ADODB.Command")
    Set cmdCommand.activeconnection = cnConn
    With cmdCommand
        .CommandText = "Select * From tbl_PortfolioPlan"
        .CommandType = adCmdText
        .Execute
    End With


    ' Open the recordset.
    Dim rstRecordset As Object
    Set rstRecordset = CreateObject("ADODB.Recordset")
    Set rstRecordset.activeconnection = cnConn
    rstRecordset.Open cmdCommand




    sRemoteCCTable = "B3:N" + CStr(3 + lMAX_ROLES)         ' Find out size of the remote CC table

    sRepFilename = "PENS_CAP_" + Format(Now(), "mmmddyyyy_hhmmss") + ".xlsm"

    ''''''''''''''''''' PATH??????
    Set wbNew = Workbooks.Add(gsLocal_Folder & "\" & "CCC.xltm")

    Set wsPT = wbNew.Sheets("CC Capacity")                 'wbNew is derived from a template...the code is already in the CC Capacity sheet

    ' Create a PivotTable cache and report.
    Set PC = wbNew.PivotCaches.Create(SourceType:=xlExternal)
    Set PC.Recordset = rstRecordset




    wbNew.Windows(1).DisplayGridlines = False

    'Data location and range to copy
    sRemData = "='" + sRemFolderRS + "\[" + sRemFileNameRS + "]" + "CC Capacity'!" + sRemoteCCTable '<< change as required

    'Get Data From Closed Book (link to worksheet)
    With wsPT.Range(wsPT.Cells(3, 2), wsPT.Cells(3 + lMAX_ROLES, 14)) '<< change as required
        .Formula = sRemData
        'convert formula to text
        .Value = .Value
    End With

    iRolesCount = getNumRolesAndPurge(wsPT, lMAX_ROLES)

    ' Initialization
    '*************************************************************************************************************************************
    lT1x1 = 3
    lT1x2 = lT1x1 + iRolesCount + 1                        ' Including Headers and Total Row
    lT1y1 = 2
    lT1y2 = 14

    lT2x1 = 3
    lT2x2 = lT2x1 + iRolesCount + 1                        ' Including Headers and Total Row
    lT2y1 = 16
    lT2y2 = 28

    lT3x1 = 3 + iRolesCount + 8                            ' Including Headers and Total Row
    lT3x2 = lT3x1 + iRolesCount + 1                        ' Including Headers and Total Row
    lT3y1 = 2
    lT3y2 = 14

    lT4x1 = 3 + iRolesCount + 8                            ' Including Headers and Total Row
    lT4x2 = lT4x1 + iRolesCount + 1                        ' Including Headers and Total Row
    lT4y1 = 16
    lT4y2 = 28

    Debug.Print lT1x1
    Debug.Print lT1x2
    Debug.Print lT1y1
    Debug.Print lT1y2

    '    Debug.Print
    '
    '    Debug.Print lT2x1
    '    Debug.Print lT2x2
    '    Debug.Print lT2y1
    '    Debug.Print lT2y2
    '
    '    Debug.Print
    '
    '    Debug.Print lT3x1
    '    Debug.Print lT3x2
    '    Debug.Print lT3y1
    '    Debug.Print lT3y2
    '
    '    Debug.Print
    '
    '    Debug.Print lT4x1
    '    Debug.Print lT4x2
    '    Debug.Print lT4y1
    '    Debug.Print lT4y2

    With wsPT
        .Range(.Cells(lT1x1, lT1y1), .Cells(lT1x2 - 1, lT1y2)).Sort Key1:=.Cells(lT1x1, lT1y1), Header:=xlYes 'Sort range by role!
        .Range(.Cells(lT2x1, lT2y1), .Cells(lT2x2, lT2y2)) = .Range(.Cells(lT1x1, lT1y1), .Cells(lT1x2, lT1y2)).Value 'Copy T1 into T2
        .Range(.Cells(lT1x1 + 1, lT1y1 + 1), .Cells(lT1x2, lT1y2)).Clear 'Clear T1 body only
        .Range(.Cells(lT2x1, lT2y1), .Cells(lT2x2, lT2y2)) = .Range(.Cells(lT2x1, lT2y1), .Cells(lT2x2, lT2y2)).Value 'Copy / paste values from T2 into T2
        .Range(.Cells(lT4x1, lT4y1), .Cells(lT4x2, lT4y2)) = .Range(.Cells(lT2x1, lT2y1), .Cells(lT2x2, lT2y2)).Value 'Copy T4 from T2

        .Range(.Cells(lT2x1 + 1, lT2y1 + 1), .Cells(lT2x2, lT2y2)).NumberFormat = "0%" ' Format T2 body only
        .Range(.Cells(lT4x1 + 1, lT4y1 + 1), .Cells(lT4x2, lT4y2)).NumberFormat = "0.0" 'Format T4 body only

        'Adding formulas for T1 and T2 tables
        .Range(.Cells(lT1x1 + 1, lT1y1 + 1), .Cells(lT1x2, lT1y2)) = "=Q" + CStr(lT4x1 + 1) + "-C" + CStr(lT4x1 + 1)
        .Range(.Cells(lT2x1 + 1, lT2y1 + 1), .Cells(lT2x2, lT2y2)) = "=IF(Q" + CStr(lT4x1 + 1) + "=0,C4,C4/Q" + CStr(lT4x1 + 1) + ")"


    End With

    ' Adding Pivot Table
    '====================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '''Set PT = PC.CreatePivotTable(TableDestination:=ActiveSheet.Cells(lT3x1 - 1, 2), TableName:=ActiveSheet.Name, DefaultVersion:=xlPivotTableVersion12)

    Set PT = PC.CreatePivotTable(TableDestination:=wsPT.Cells(lT3x1 - 1, 2), TableName:="CapVsDemPT", DefaultVersion:=xlPivotTableVersion12)

    PT.PivotFields("Roles").Orientation = xlRowField
    PT.PivotFields("Roles").Subtotals(1) = False           ' 1 = Automatic

    PT.PivotFields("Project Name").Orientation = xlRowField
    PT.PivotFields("Project Name").Subtotals(1) = False
    PT.PivotFields("Project Name").AutoSort Order:=xlDescending, Field:="Program"

    PT.PivotFields("Names").Orientation = xlRowField
    PT.PivotFields("Names").Subtotals(1) = False

    PT.AddDataField PT.PivotFields("JAN"), " Jan", xlSum
    PT.AddDataField PT.PivotFields("FEB"), " Feb", xlSum
    PT.AddDataField PT.PivotFields("MAR"), " Mar", xlSum
    PT.AddDataField PT.PivotFields("APR"), " Apr", xlSum
    PT.AddDataField PT.PivotFields("MAY"), " May", xlSum
    PT.AddDataField PT.PivotFields("JUN"), " Jun", xlSum
    PT.AddDataField PT.PivotFields("JUL"), " Jul", xlSum
    PT.AddDataField PT.PivotFields("AUG"), " Aug", xlSum
    PT.AddDataField PT.PivotFields("SEP"), " Sep", xlSum
    PT.AddDataField PT.PivotFields("OCT"), " Oct", xlSum
    PT.AddDataField PT.PivotFields("NOV"), " Nov", xlSum
    PT.AddDataField PT.PivotFields("DEC"), " Dec", xlSum


    ''''PT.Name = "CapVsDemPT"
    PT.DataPivotField.Value = ""                           'Remove the word  from value column
    PT.CompactLayoutRowHeader = "Role"                     ' Changes the default 
    PT.TableRange2.Font.Color = vbBlack


    If gbFY17 Then
        PT.PivotFields("Project Priority").Orientation = xlPageField
        PT.PivotFields("Project Priority").ClearAllFilters
        PT.PivotFields("Budget Category").Orientation = xlPageField
        PT.PivotFields("Budget Category").ClearAllFilters
    Else
        PT.PivotFields("Project Type").Orientation = xlPageField
        PT.PivotFields("Project Type").ClearAllFilters
    End If

    PT.PivotFields("Account Manager").Orientation = xlPageField
    PT.PivotFields("Account Manager").ClearAllFilters

    PT.DataBodyRange.NumberFormat = "0.0"
    PT.DataLabelRange.HorizontalAlignment = xlRight

    PT.PivotFields("Roles").ShowDetail = False
    PT.PivotFields("Project Name").ShowDetail = False

    'Clear Out Any Previous Filtering
    PT.PivotFields("Roles").ClearAllFilters

    'Enable filtering on multiple items
    PT.PivotFields("Roles").EnableMultiplePageItems = True
    PT.PivotFields("Roles").Subtotals(1) = True

    'Turning off items we do not want showing
    For Each pi In PT.PivotFields("Roles").PivotItems
        Set rngFindValue = wsPT.Range(wsPT.Cells(lT4x1, lT4y1), wsPT.Cells(lT4x2 - 1, lT4y1)).Find(What:=pi.Caption, after:=wsPT.Cells(lT4x1, lT4y1), LookIn:=xlValues, lookAt:=xlWhole, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)

        If rngFindValue Is Nothing Then
            pi.Visible = False
        End If
        '''''If pi.Caption <> "(blank)" Then pi.Visible = False
    Next


    Set r = wsPT.Range(wsPT.Cells(lT1x1 + 1, lT1y1 + 1), wsPT.Cells(lT1x2 - 1, lT1y2))
    Call CondFormat(r)

    Set r = wsPT.Range(wsPT.Cells(lT2x1 + 1, lT2y1 + 1), wsPT.Cells(lT2x2 - 1, lT2y2))
    Call CondFormat(r)

    Set r = wsPT.Range(wsPT.Cells(lT4x1 + 1, lT4y1 + 1), wsPT.Cells(lT4x2 - 1, lT4y2))
    Call CondFormat(r)

    dGrandTotal = 0
    '''For i = 2 To 13
    ''' dGrandTotal2 = dGrandTotal2 + wsPT.Cells(PT.RowRange.Count + 3, i)
    '''Next

    ''''wsPT.Name = "CC Capacity"

    GenPT_CC_CAPACITY = dGrandTotal


    With wsPT
        .Range(.Cells(lT1x1, lT1y1), .Cells(lT1x1, lT1y2)).Font.Bold = True
        .Range(.Cells(lT2x1, lT2y1), .Cells(lT2x1, lT2y2)).Font.Bold = True
        .Range(.Cells(lT4x1, lT4y1), .Cells(lT4x1, lT4y2)).Font.Bold = True

        .Cells(lT4x2, lT4y1) = "Total"
        .Range(.Cells(lT4x2, lT4y1 + 1), .Cells(lT4x2, lT4y2)).FormulaR1C1 = "=SUM(R[-" + CStr(iRolesCount) + "]C:R[-1]C)"
        .Range(.Cells(lT4x2, lT4y1), .Cells(lT4x2, lT4y2)).Interior.Color = RGB(219, 229, 241)
        .Range(.Cells(lT4x2, lT4y1), .Cells(lT4x2, lT4y2)).Font.Bold = True

        .Cells(lT1x2, lT1y1) = "Total"
        .Range(.Cells(lT1x2, lT1y1 + 1), .Cells(lT1x2, lT1y2)).FormulaR1C1 = "=SUM(R[-" + CStr(iRolesCount) + "]C:R[-1]C)"
        .Range(.Cells(lT1x2, lT1y1), .Cells(lT1x2, lT1y2)).Interior.Color = RGB(219, 229, 241)
        .Range(.Cells(lT1x2, lT1y1), .Cells(lT1x2, lT1y2)).Font.Bold = True

        .Cells(lT2x2, lT2y1) = "Total"
        .Range(.Cells(lT2x2, lT2y1 + 1), .Cells(lT2x2, lT2y2)).FormulaR1C1 = "=SUM(R[-" + CStr(iRolesCount) + "]C:R[-1]C)"
        .Range(.Cells(lT2x2, lT2y1), .Cells(lT2x2, lT2y2)).Interior.Color = RGB(219, 229, 241)
        .Range(.Cells(lT2x2, lT2y1), .Cells(lT2x2, lT2y2)).Font.Bold = True

        ' Adjusting row height and col width
        .Columns("A").ColumnWidth = 0.5
        .Columns("O").ColumnWidth = 0.5
        .Range("C:N").ColumnWidth = 7
        .Range("Q:AB").ColumnWidth = 7
        .Columns("P").ColumnWidth = .Columns("B").ColumnWidth

        .Rows(CStr(lT1x1) + ":" + CStr(lT1x2)).RowHeight = 14
        .Rows(CStr(lT4x1) + ":" + CStr(lT4x2)).RowHeight = 14
    End With


    ' Adding labels for the tables
    '==============================
    wsPT.Range(wsPT.Cells(1, lT1y1), wsPT.Cells(1, lT1y2)).Merge

    With wsPT.Cells(1, lT1y1)
        .HorizontalAlignment = xlCenter
        .Value = "FTE Career Centre Capacity vs Demand"
        .Interior.Color = RGB(55, 96, 145)
        .Font.Color = vbWhite
    End With

    wsPT.Range(wsPT.Cells(1, lT2y1), wsPT.Cells(1, lT2y2)).Merge
    With wsPT.Cells(1, lT2y1)
        .HorizontalAlignment = xlCenter
        .Value = "% Over/Under Utilization of FTE"
        .Interior.Color = RGB(55, 96, 145)
        .Font.Color = vbWhite
    End With

    wsPT.Range(wsPT.Cells(lT3x1 - 2, lT3y1), wsPT.Cells(lT3x1 - 2, lT3y2)).Merge
    With wsPT.Cells(lT3x1 - 2, lT3y1)
        .HorizontalAlignment = xlCenter
        .Value = "Project Demand"
        .Interior.Color = RGB(55, 96, 145)
        .Font.Color = vbWhite
    End With

    wsPT.Range(wsPT.Cells(lT4x1 - 2, lT4y1), wsPT.Cells(lT4x1 - 2, lT4y2)).Merge
    With wsPT.Cells(lT4x1 - 2, lT4y1)
        .HorizontalAlignment = xlCenter
        .Value = "Career Centre Resource Capacity"
        .Interior.Color = RGB(55, 96, 145)
        .Font.Color = vbWhite
    End With

    ' Adding labels indicating the Year
    '==================================
    With wsPT.Cells(lT1x1 - 1, 7)                          'G2
        If gbFY17 Then
            .Value = 2017
        Else
            .Value = 2016
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(79, 98, 40)
        .Font.Color = vbYellow
    End With

    With wsPT.Cells(lT2x1 - 1, 21)                         'U2
        If gbFY17 Then
            .Value = 2017
        Else
            .Value = 2016
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(79, 98, 40)
        .Font.Color = vbYellow
    End With

    With wsPT.Cells(lT3x1 - 3, 7)        'G28              ''' TODOS ESTOS NUMBEROS CAMBIAN DEPENDIENDO DE LA CANTIDAD DE FILAS!!!!
        If gbFY17 Then
            .Value = 2017
        Else
            .Value = 2016
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(79, 98, 40)
        .Font.Color = vbYellow
    End With

    With wsPT.Cells(lT4x1 - 3, 21)                         'U30
        If gbFY17 Then
            .Value = 2017
        Else
            .Value = 2016
        End If
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(79, 98, 40)
        .Font.Color = vbYellow
    End With

    wsPT.Copy after:=wsPT
    ActiveSheet.Name = "CCBKP"
    ActiveSheet.Visible = xlSheetHidden                    ' or xlSheetVeryHidden

    wsPT.Activate

    '   Converting to Tables
    '    With wsPT
    '        .ListObjects.Add(xlSrcRange, .Range(wsPT.Cells(lT1x1, lT1y1), .Cells(lT1x2, lT1y2)), , xlYes).Name = "Table1"
    '        .ListObjects("Table1").TableStyle = ""
    '
    '        .ListObjects.Add(xlSrcRange, .Range(wsPT.Cells(lT2x1, lT2y1), .Cells(lT2x2, lT2y2)), , xlYes).Name = "Table2"
    '        .ListObjects("Table2").TableStyle = ""
    '
    '        .ListObjects.Add(xlSrcRange, .Range(wsPT.Cells(lT4x1, lT4y1), .Cells(lT4x2, lT4y2)), , xlYes).Name = "Table4"
    '        .ListObjects("Table4").TableStyle = ""
    '    End With


    ' Adding Name ranges for the Tables
    With wbNew.Sheets("CC Capacity")
        wbNew.Names.Add Name:="Table1", RefersTo:=.Range(.Cells(lT1x1, lT1y1), .Cells(lT1x2, lT1y2)) '''' 
        wbNew.Names.Add Name:="Table2", RefersTo:=.Range(.Cells(lT2x1, lT2y1), .Cells(lT2x2, lT2y2))
        wbNew.Names.Add Name:="Table4", RefersTo:=.Range(.Cells(lT4x1, lT4y1), .Cells(lT4x2, lT4y2))
    End With

    With wbNew.Sheets("CCBKP")                             'Same tables in CCBKP
        wbNew.Names.Add Name:="Table5", RefersTo:=.Range(.Cells(lT1x1, lT1y1), .Cells(lT1x2, lT1y2)) '''' 
        wbNew.Names.Add Name:="Table6", RefersTo:=.Range(.Cells(lT2x1, lT2y1), .Cells(lT2x2, lT2y2))
        wbNew.Names.Add Name:="Table8", RefersTo:=.Range(.Cells(lT4x1, lT4y1), .Cells(lT4x2, lT4y2))
    End With


    ' Doing this to trigger the macro in the template!!!!
    PT.PivotFields("Account Manager").CurrentPage = "ALL"



    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    wbNew.SaveAs FileName:=gsLocal_Folder & "\" & sRepFilename, FileFormat:=52

    gsLastReport = wbNew.Name

    cnConn.Close

    Set cnConn = Nothing
    Set cmdCommand = Nothing
    Set rstRecordset = Nothing

    Set wbNew = Nothing

ErrorExit:
    ' Clean up
    Exit Function

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

'---------------------------------------------------------------------------------------
' Method : GenPT_NES
' Author : cklahr
' Date   : 11/16/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function GenPT_NES(ByVal wbNew As Workbook, ByVal ws As Worksheet) As Long

    Dim wsPT2 As Worksheet
    Dim PC As PivotCache
    Dim PT As PivotTable

    Dim CON As WorkbookConnection
    Dim stCon As String
    Dim spath As String
    Dim sCommandText As String

    Dim r As Range
    Dim cs As ColorScale
    Dim dGrandTotal As Double
    Dim i As Integer

    Dim sData As String

    Const sSOURCE As String = "GenPT_NES"


    'Determine the data range you want to pivot
    sData = ws.Name & "!" & Range("A1:AX10000").Address(ReferenceStyle:=xlR1C1)

    'Create Pivot Cache from Source Data
    Set PC = ActiveWorkbook.PivotCaches.Create( _
    SourceType:=xlDatabase, _
    SourceData:=sData)



    'In theory I'm working on the new D4A workbook...!!!???
    Set wsPT2 = Sheets.Add
    wsPT2.Name = "Portfolio_Projects_PivotTable"


    Set PT = PC.CreatePivotTable(TableDestination:=ActiveSheet.Name & "!R1c1", TableName:=ActiveSheet.Name, DefaultVersion:=xlPivotTableVersion12)

    ''''PT.InGridDropZones = True        ' Classic View
    PT.RowAxisLayout xlTabularRow


    PT.TableStyle2 = "PivotStyleMedium9"



    PT.PivotFields("LOB").Orientation = xlRowField
    PT.PivotFields("LOB").Subtotals(1) = True              ' 1 = Automatic

    If gbFY17 Then
        PT.PivotFields("Budget Category").Orientation = xlRowField
        PT.PivotFields("Budget Category").Subtotals(1) = False ' 1 = Automatic
    Else
        PT.PivotFields("Project Type").Orientation = xlRowField
        PT.PivotFields("Project Type").Subtotals(1) = False ' 1 = Automatic
    End If

    PT.PivotFields("Project Name").Orientation = xlRowField
    PT.PivotFields("Project Name").Subtotals(1) = False    ' 1 = Automatic


    PT.AddDataField PT.PivotFields("Budget"), " Budget", xlSum
    PT.PivotFields(" Budget").NumberFormat = "$#,##0"

    PT.AddDataField PT.PivotFields("Dashboard NE"), " DB NE", xlSum
    PT.PivotFields(" DB NE").NumberFormat = "$#,##0"

    '''PT.AddDataField PT.PivotFields("Dashboard NE - Budget" + vbLf + "($)"), " DB NE - Budget" + vbLf + "($)", xlSum
    PT.AddDataField PT.PivotFields(8), " DB NE - Budget" + vbLf + "($)", xlSum '8 is index for  + vbLf +  field (for some reason VBA doesn't like vbLf in the field name :-(


    ''' Dashboard NE - Budget($)

    PT.PivotFields(" DB NE - Budget" + vbLf + "($)").NumberFormat = "$#,##0"


    PT.CalculatedFields.Add Name:="DNEvsBudget", Formula:="=IF(Budget = 0,0,(Dashboard NE - Budget) / Budget)"
    PT.AddDataField PT.PivotFields("DNEvsBudget"), " DB NE - Budget" + vbLf + "(%)", xlSum
    PT.PivotFields(" DB NE - Budget" + vbLf + "(%)").NumberFormat = "0.0%"



    'PT.AddDataField PT.PivotFields("Dashboard NE - Budget_(%)"), " Dashboard NE - Budget (%)", xlSum
    'PT.PivotFields(" Dashboard NE - Budget (%)").NumberFormat = "0.0%"




    PT.AddDataField PT.PivotFields("YTD Actuals"), " YTD" + vbLf + "Actuals", xlSum
    PT.PivotFields(" YTD" + vbLf + "Actuals").NumberFormat = "$#,##0"

    PT.AddDataField PT.PivotFields("Reporting NE"), " Rep NE", xlSum
    PT.PivotFields(" Rep NE").NumberFormat = "$#,##0"


    '''''''PT.AddDataField PT.PivotFields("Reporting NE - Budget" + vbLf + "($)"), " Rep NE - Budget" + vbLf + "($)", xlSum
    PT.AddDataField PT.PivotFields(9), " Rep NE - Budget" + vbLf + "($)", xlSum '9 is index for  + vbLf +  field(for some reason VBA doesn't like vbLf in the field name :-(
    PT.PivotFields(" Rep NE - Budget" + vbLf + "($)").NumberFormat = "$#,##0"

    PT.CalculatedFields.Add Name:="RNEvsBudget", Formula:="=IF(Budget = 0,0,(Reporting NE - Budget) / Budget)"
    PT.AddDataField PT.PivotFields("RNEvsBudget"), " Rep NE - Budget" + vbLf + "(%)", xlSum
    PT.PivotFields(" Rep NE - Budget" + vbLf + "(%)").NumberFormat = "0.0%"

    ''''PT.CalculatedFields.Add Name:="DNEvsBudget", Formula:="=IF(Budget = 0,0,(Dashboard NE - Budget) / Budget)"

    'With PT.PivotFields("DNEvsBudget")
    '    .Orientation = xlDataField
    '    .Function = xlSum
    '    .Position = 4
    '    .NumberFormat = "0.0%"
    '    .Caption = "BOSTA-%"
    'End With


    ''''''PT.AddDataField PT.PivotFields("Budget / CR Notes"), " Budget / CR Notes", xlSum
    ''''''PT.AddDataField PT.PivotFields("NE Explanation"), " NE Explanation", xlSum


    '''PT.DataBodyRange.NumberFormat = "$#,###"
    '''PT.DataLabelRange.HorizontalAlignment = xlRight



    PT.CompactLayoutColumnHeader = "LOB"
    ''''''PT.CompactLayoutRowHeader = "Cost Details"  '''''''' NO ANDUVO!

    'wsPT2.Columns("A:K").AutoFit

    PT.MergeLabels = True

    wsPT2.Columns("A").ColumnWidth = 22
    wsPT2.Columns("B").ColumnWidth = 21
    wsPT2.Columns("C").ColumnWidth = 48

    wsPT2.Range("D:K").ColumnWidth = 11
    'wsPT2.Columns("G").ColumnWidth = 8

    'wsPT2.Columns("H:I").ColumnWidth = 12
    'wsPT2.Columns("J:K").ColumnWidth = 60



    '''wsPT2.Columns("D:I").ColumnWidth = 21
    wsPT2.Columns("A:C").HorizontalAlignment = xlLeft


    ''''PT.DataLabelRange.HorizontalAlignment = xlCenter


    With wsPT2.Range("A2:I2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        ''.Interior.Color = RGB(79, 98, 40)
        '''.Font.Color = vbYellow
    End With


    Dim lLastRow As Long

    With PT.TableRange1
        lLastRow = .Cells(.Cells.Count).Row
    End With


    With wsPT2
        .Rows(1).Hidden = True
        .Range("A3").Select

        .PageSetup.Orientation = xlLandscape
        .PageSetup.PrintArea = "$A$1:$K$" + CStr(lLastRow)
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False

    End With

    ActiveWindow.FreezePanes = True

    'wbNew is the D4A workbook...
    '''''Set wsPT = wbNew.Sheets.Add
    '''''wbNew.Windows(1).DisplayGridlines = False

    '''''Set PT = PC.CreatePivotTable(TableDestination:=ActiveSheet.Name & "!R1c1", TableName:=ActiveSheet.Name, DefaultVersion:=xlPivotTableVersion12)

ErrorExit:
    ' Clean up
    Exit Function

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

'---------------------------------------------------------------------------------------
' Method : Gen_NES
' Author : cklahr
' Date   : 2/11/2016
' Purpose: Produce NE Summary
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function Gen_NES(ByVal wbNew As Workbook, ByVal sFilePP As String, ByVal bPivot As Boolean) As Double
    Const sSOURCE As String = "ProduceNES"

    Dim PC As Object
    Dim sdConPP As WorkbookConnection
    Dim sdConRS As WorkbookConnection
    Dim sConnPP As String
    Dim sConnRS As String
    Dim sCommandTextPP As String
    Dim sCommandTextRS As String

    Dim dGrandTotalFTE_NE As Double
    Dim dGrandTotalPROJ_SUMMARY As Double
    Dim dGrandTotalCC_CAPACITY As Double

    Dim dRoundUp As Double
    Dim dRoundDown As Double

    Dim wsPT As Worksheet

    Dim cmdCommand2 As Object
    Dim rstRecordset2 As Object


    'Dim sD4AFileName As String
    '   Dim wbNew As Workbook
    '
    '  sD4AFileName = "D4A_" + Format(Now(), "mmmddyyyy_hhmmss") + ".xlsx"
    ' Set wbNew = Workbooks.Add

    'wbNew.Activate


    'wbNew is the D4A workbook...
    Set wsPT = wbNew.Sheets.Add
    wbNew.Windows(1).DisplayGridlines = False

    Application.ErrorCheckingOptions.InconsistentFormula = False


    Dim sPPDB As String
    '''sPPDB = gwsConfig.Range(gsLOCAL_FOLDER).Value + "\" + gwsConfig.Range(gsDB_NAME).Value        'PDASH.accdb

    sPPDB = gsLocal_Folder + "\" + gwsConfig.Range(gsDB_NAME).Value 'PDASH.accdb


    Dim cnConn As Object
    Set cnConn = CreateObject("ADODB.Connection")
    With cnConn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
        .Open sPPDB
    End With


    ' Set the command text.
    Dim cmdCommand As Object
    Set cmdCommand = CreateObject("ADODB.Command")
    Set cmdCommand.activeconnection = cnConn
    With cmdCommand
        '''''@@@@@@@@ .CommandText = "SELECT [Names] FROM tbl_PortfolioPlan WHERE [Roles] = 'Actual - HW Amort'" 'Budget

        If gbFY17 Then
            .CommandText = "SELECT DISTINCT [Project Code], [Project Name], [LOB], [Account Manager], [Activation Status], [Budget Category], [Publish NE] FROM tbl_PortfolioPlan"
        Else
            .CommandText = "SELECT DISTINCT [Project Code], [Project Name], [LOB], [Account Manager], [Activation Status], [Project Type], [Publish NE] FROM tbl_PortfolioPlan"
        End If
        .CommandType = adCmdText
        .Execute
    End With

    ' Open the recordset.
    Dim rstRecordset As Object
    Set rstRecordset = CreateObject("ADODB.Recordset")
    Set rstRecordset.activeconnection = cnConn
    rstRecordset.Open cmdCommand


    Set cmdCommand2 = CreateObject("ADODB.Command")

    Set cmdCommand2.activeconnection = cnConn


    ' Create a PivotTable cache and report.
    ''''''''''''''' Set PC = wbNew.PivotCaches.Create(SourceType:=xlExternal)
    '''''''''''''''Set PC.Recordset = rstRecordset


    Dim i As Long
    i = 2
    With wsPT
        .Cells(1, 1) = "LOB"

        If gbFY17 Then
            .Cells(1, 2) = "Budget Category"
        Else
            .Cells(1, 2) = "Project Type"
        End If
        .Cells(1, 3) = "Project Code"
        .Cells(1, 4) = "Project Name"
        .Cells(1, 5) = "Budget"
        .Cells(1, 6) = "Last Revised BL"
        .Cells(1, 7) = "Dashboard NE"
        .Cells(1, 7).AddComment ("- Yellow: > 5% compared to Revised BL" + vbCrLf + "- Red: > 10% compared to Revised BL")

        .Cells(1, 8) = "Dashboard NE - Budget" + vbLf + "($)"
        .Cells(1, 9) = "Dashboard NE - Budget" + vbLf + "(%)"
        '''.Range(.Cells(1, 8), .Cells(1, 9)).Merge
        '''.Cells(2, 8) = "($)"
        '''.Cells(2, 9) = "(%)"
        .Cells(1, 9).AddComment ("- Yellow: > 0% variance compared to Budget" + vbCrLf + "- Red: > 5% variance compared to Budget")

        .Cells(1, 10) = "Dashboard NE - Revised BL" + vbLf + "($)"
        .Cells(1, 11) = "Dashboard NE - Revised BL" + vbLf + "(%)"
        ''.Range(.Cells(1, 10), .Cells(1, 11)).Merge
        ''.Cells(2, 10) = "($)"
        ''.Cells(2, 11) = "(%)"
        .Cells(1, 11).AddComment ("- Yellow: > 0% variance compared to Revised BL" + vbCrLf + "- Red: > 5% variance compared to Revised BL")

        .Cells(1, 12) = "Reporting NE"
        ''.Cells(2, 12) = "($)"

        .Cells(1, 13) = "Reporting NE - Budget" + vbLf + "($)"
        .Cells(1, 14) = "Reporting NE - Budget" + vbLf + "(%)"
        ''.Range(.Cells(1, 13), .Cells(1, 14)).Merge
        ''.Cells(2, 13) = "($)"
        ''.Cells(2, 14) = "(%)"
        .Cells(1, 14).AddComment ("- Yellow: > 0% variance compared to Budget" + vbCrLf + "- Red: > 5% variance compared to Budget")

        .Cells(1, 15) = "Reporting NE - Revised BL" + vbLf + "($)"
        .Cells(1, 16) = "Reporting NE - Revised BL" + vbLf + "(%)"
        ''.Range(.Cells(1, 15), .Cells(1, 16)).Merge
        ''.Cells(2, 15) = "($)"
        ''.Cells(2, 16) = "(%)"
        .Cells(1, 16).AddComment ("- Yellow: > 0% variance compared to Revised BL" + vbCrLf + "- Red: > 5% variance compared to Revised BL")


        .Cells(1, 17) = "Reporting NE - Dashboard NE" + vbLf + "($)"
        .Cells(1, 18) = "Reporting NE - Dashboard NE" + vbLf + "(%)"
        ''.Range(.Cells(1, 17), .Cells(1, 18)).Merge
        ''.Cells(2, 17) = "($)"
        ''.Cells(2, 18) = "(%)"
        .Cells(1, 18).AddComment ("- Yellow: > 0% variance compared to Dashboard NE" + vbCrLf + "- Red: > 5% variance compared to Dashboard NE")


        .Cells(1, 19) = "YTD Actuals"
        ''.Cells(2, 19) = "($)"
        .Cells(1, 20) = "YTD Actuals % of Reporting NE"
        ''.Cells(2, 20) = "(%)"


        .Cells(1, 21) = "Project Status" + vbLf + "(Calculated)"
        .Cells(1, 22) = "Project Status" + vbLf + "(Dashboard)"
        ''.Range(.Cells(1, 21), .Cells(1, 22)).Merge
        ''.Cells(2, 21) = "(Calculated)"
        ''.Cells(2, 22) = "(Dashboard)"

        .Cells(1, 23) = "Account Manager"
        .Cells(1, 24) = "Activation Status"

        .Cells(1, 25) = "MI"
        .Cells(1, 26) = "IG"
        .Cells(1, 27) = "IPC"
        .Cells(1, 28) = "LifeCo"
        .Cells(1, 29) = "MI"
        .Cells(1, 30) = "IG"
        .Cells(1, 31) = "IPC"
        .Cells(1, 32) = "LifeCo"
        .Cells(1, 33) = "Publish Yes/No"
        ''.Cells(1, 34) = "FTE NE - MI"
        '''.Range(.Cells(1, 34), .Cells(1, 45)).Merge
        .Cells(1, 34) = "FTE NE - MI" + vbLf + "(JAN)"
        .Cells(1, 35) = "FTE NE - MI" + vbLf + "(FEB)"
        .Cells(1, 36) = "FTE NE - MI" + vbLf + "(MAR)"
        .Cells(1, 37) = "FTE NE - MI" + vbLf + "(APR)"
        .Cells(1, 38) = "FTE NE - MI" + vbLf + "(MAY)"
        .Cells(1, 39) = "FTE NE - MI" + vbLf + "(JUN)"
        .Cells(1, 40) = "FTE NE - MI" + vbLf + "(JUL)"
        .Cells(1, 41) = "FTE NE - MI" + vbLf + "(AUG)"
        .Cells(1, 42) = "FTE NE - MI" + vbLf + "(SEP)"
        .Cells(1, 43) = "FTE NE - MI" + vbLf + "(OCT)"
        .Cells(1, 44) = "FTE NE - MI" + vbLf + "(NOV)"
        .Cells(1, 45) = "FTE NE - MI" + vbLf + "(DEC)"
        .Cells(1, 46) = "MI Labor Cost Dashboard NE"
        ''.Cells(2, 46) = "($)"
        .Cells(1, 47) = "MI Labor Cost Budget"
        ''.Cells(2, 47) = "($)"
        .Cells(1, 48) = "MI Labor Cost Revised BL"
        ''.Cells(2, 48) = "($)"
        .Cells(1, 49) = "Budget / CR Notes"
        .Cells(1, 50) = "NE Explanation"

        .Range(.Cells(1, 1), .Cells(1, 50)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(1, 50)).WrapText = True
        .Range(.Cells(1, 1), .Cells(1, 50)).HorizontalAlignment = xlCenter
        .Range(.Cells(1, 1), .Cells(1, 50)).VerticalAlignment = xlTop
        .Range(.Cells(1, 1), .Cells(1, 50)).Interior.Color = RGB(80, 130, 190) 'RGB(238, 236, 225)
        .Range(.Cells(1, 1), .Cells(1, 50)).Font.Color = vbWhite 'RGB(31, 73, 125)

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        While Not rstRecordset.EOF

            ''''.Sheets("Sheet1").Cells(i, 1).Value = CDbl(Trim(rstRecordset.fields("Names")))
            If Len(rstRecordset.Fields("Project Code")) > 1 And rstRecordset.Fields("Project Code") <> "WHATIF" Then
                .Cells(i, 1).Value = Trim(rstRecordset.Fields("LOB"))

                If gbFY17 Then
                    .Cells(i, 2).Value = Trim(rstRecordset.Fields("Budget Category"))
                Else
                    .Cells(i, 2).Value = Trim(rstRecordset.Fields("Project Type"))
                End If

                .Cells(i, 3).Value = Trim(rstRecordset.Fields("Project Code"))
                .Cells(i, 4).Value = Trim(rstRecordset.Fields("Project Name"))
                If gbFY17 Then
                    .Cells(i, 5).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - SW SAAS", "Project Code", rstRecordset.Fields("Project Code"))), "$#,##0") ' Budget
                Else
                    .Cells(i, 5).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW Amort", "Project Code", rstRecordset.Fields("Project Code"))), "$#,##0") ' Budget
                End If
                .Cells(i, 6).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW < $5K", "Project Code", rstRecordset.Fields("Project Code"))), "$#,##0") ' Revised BL
                .Cells(i, 7).Value = Format(FetchValue(cnConn, "Do not remove1", "Roles", "NE (Actual + Rem. Plan)", "Project Code", rstRecordset.Fields("Project Code")), "$#,##0") ' Dashboard NE

                ''.Cells(i, 8).Value = "=G" & i & "-E" & i

                .Cells(i, 8) = "=IF((G" & i & "-E" & i & ")<>0,G" & i & "-E" & i & ","""")"
                .Cells(i, 8).NumberFormat = "$#,##0"

                ' .Cells(i, 9) = "=IF(E" & i & "=0,-100%,H" & i & "/E" & i & ")"
                .Cells(i, 9) = "=IF(E" & i & "=0,-100%,IF(H" & i & "<>"""",H" & i & "/E" & i & ",""""))"



                '=IF(E4=0,-100%,IF(H4<>"",H4/E4,""))

                .Cells(i, 9).NumberFormat = "0%"



                .Cells(i, 10) = "=IF((G" & i & "-F" & i & ")<>0,G" & i & "-F" & i & ","""")"
                .Cells(i, 10).NumberFormat = "$#,##0"
                .Cells(i, 11) = "=IF(F" & i & "=0,-100%,IF(J" & i & "<>"""",J" & i & "/F" & i & ",""""))"
                .Cells(i, 11).NumberFormat = "0%"


                .Cells(i, 12).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW Maint", "Project Code", rstRecordset.Fields("Project Code"))), "$#,##0") ' Reporting NE

                .Cells(i, 13) = "=IF((L" & i & "-E" & i & ")<>0,L" & i & "-E" & i & ","""")"
                .Cells(i, 13).NumberFormat = "$#,##0"
                .Cells(i, 14) = "=IF(E" & i & "=0,-100%,IF(M" & i & "<>"""",M" & i & "/E" & i & ",""""))"
                .Cells(i, 14).NumberFormat = "0%"

                .Cells(i, 15) = "=IF((L" & i & "-F" & i & ")<>0,L" & i & "-F" & i & ","""")"
                .Cells(i, 15).NumberFormat = "$#,##0"
                .Cells(i, 16) = "=IF(F" & i & "=0,-100%,IF(O" & i & "<>"""",O" & i & "/F" & i & ",""""))"
                .Cells(i, 16).NumberFormat = "0%"


                .Cells(i, 17) = "=IF((L" & i & "-G" & i & ")<>0,L" & i & "-G" & i & ","""")"
                .Cells(i, 17).NumberFormat = "$#,##0"
                .Cells(i, 18) = "=IF(G" & i & "=0,-100%,IF(Q" & i & "<>"""",Q" & i & "/G" & i & ",""""))"
                .Cells(i, 18).NumberFormat = "0%"


                '.Cells(i, 12) = "=IF((D" & i & "-G" & i & ")>0,D" & i & "-G" & i & ","""")"
                '.Cells(i, 12).NumberFormat = "$#,##0"
                '.Cells(i, 13) = "=IF((G" & i & "-D" & i & ")>0,G" & i & "-D" & i & ","""")"
                '.Cells(i, 13).NumberFormat = "$#,##0"


                .Cells(i, 19).Value = Format(FetchValue(cnConn, "Names", "Roles", "Actual - Travel", "Project Code", rstRecordset.Fields("Project Code")), "$#,##0") 'YTD Actuals


                ' If Actuals greater than Reporting NE, then overwrite Reporting NE with Actuals and highlith in redish
                If .Cells(i, 19).Value > .Cells(i, 12).Value Then
                    .Cells(i, 12).Value = Format(.Cells(i, 19).Value, "$#,##0")
                    .Cells(i, 12).Font.Color = RGB(156, 0, 6)

                End If


                .Cells(i, 20).Value = "=IFERROR(S" & i & "/L" & i & ",0)"
                .Cells(i, 20).NumberFormat = "0%"


                ''''.Cells(i, 21) = "=IF(ABS(K" & i & ")<0.05,""Green"",IF(ABS(K" & i & ")<0.1,""Yellow"",""Red""))"
                .Cells(i, 21) = "=IF(K" & i & "<>"""", IF(ABS(K" & i & ")<0.05,""Green"",IF(ABS(K" & i & ")<0.1,""Yellow"",""Red"")),""Green"")"


                .Cells(i, 22).Value = Trim((FetchValue(cnConn, "Names", "Roles", "Actual - MI", "Project Code", rstRecordset.Fields("Project Code")))) ' Project Status

                .Cells(i, 23).Value = Trim(rstRecordset.Fields("Account Manager"))
                .Cells(i, 24).Value = Trim(rstRecordset.Fields("Activation Status"))

                .Cells(i, 25).Value = "=L" & i & "*AC" & i
                .Cells(i, 25).NumberFormat = "$#,##0"
                .Cells(i, 26).Value = "=L" & i & "*AD" & i
                .Cells(i, 26).NumberFormat = "$#,##0"
                .Cells(i, 27).Value = "=L" & i & "*AE" & i
                .Cells(i, 27).NumberFormat = "$#,##0"
                .Cells(i, 28).Value = "=L" & i & "*AF" & i
                .Cells(i, 28).NumberFormat = "$#,##0"
                .Cells(i, 29).Value = Trim((FetchValue(cnConn, "Names", "Roles", "Plan - SES", "Project Code", rstRecordset.Fields("Project Code")))) ' MI %
                .Cells(i, 29).NumberFormat = "0%"
                .Cells(i, 30).Value = Trim((FetchValue(cnConn, "Names", "Roles", "Plan - SW Amort", "Project Code", rstRecordset.Fields("Project Code")))) ' IG %
                .Cells(i, 30).NumberFormat = "0%"
                .Cells(i, 31).Value = Trim((FetchValue(cnConn, "Names", "Roles", "Plan - HW < $5K", "Project Code", rstRecordset.Fields("Project Code")))) ' IPC %
                .Cells(i, 31).NumberFormat = "0%"
                .Cells(i, 32).Value = Trim((FetchValue(cnConn, "Names", "Roles", "Plan - Other", "Project Code", rstRecordset.Fields("Project Code")))) ' LC %
                .Cells(i, 32).NumberFormat = "0%"
                .Cells(i, 33).Value = Trim(rstRecordset.Fields("Publish NE"))

                ''''' OPTIMIZE!!!!
                ''''???????.Cells(i, 31).Value = FetchValue(cnConn, "JAN], [FEB], [MAR", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Jan)






                With cmdCommand2

                    .CommandText = "SELECT [JAN], [FEB], [MAR], [APR], [MAY], [Jun], [Jul], [Aug], [Sep], [Oct], [Nov], [Dec] FROM tbl_PortfolioPlan WHERE [Roles] = " + _
                    "'" + "FTE NE - MI" + "'" + " AND [Project Code] = " + "'" + rstRecordset.Fields("Project Code") + "'"


                    .CommandType = adCmdText
                    .Execute
                End With



                ' Open the recordset.

                Set rstRecordset2 = CreateObject("ADODB.Recordset")
                Set rstRecordset2.activeconnection = cnConn
                rstRecordset2.Open cmdCommand2





                ' .Cells(i, 31).Value = FetchValue(cnConn, "JAN", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Jan)
                ' .Cells(i, 32).Value = FetchValue(cnConn, "FEB", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Feb)
                ' .Cells(i, 33).Value = FetchValue(cnConn, "MAR", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Mar)
                ' .Cells(i, 34).Value = FetchValue(cnConn, "APR", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Apr)
                ' .Cells(i, 35).Value = FetchValue(cnConn, "MAY", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (May)
                ' .Cells(i, 36).Value = FetchValue(cnConn, "JUN", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Jun)

                '.Cells(i, 37).Value = FetchValue(cnConn, "JUL", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Jul)
                '.Cells(i, 38).Value = FetchValue(cnConn, "AUG", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Aug)
                '.Cells(i, 39).Value = FetchValue(cnConn, "SEP", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Sep)
                '.Cells(i, 40).Value = FetchValue(cnConn, "OCT", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Oct)
                '.Cells(i, 41).Value = FetchValue(cnConn, "NOV", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Nov)
                '.Cells(i, 42).Value = FetchValue(cnConn, "DEC", "Roles", "FTE NE - MI", "Project Code", rstRecordset.fields("Project Code"))        ' FTE NE - MI (Dec)




                If (Trim(rstRecordset2.Fields("JAN")) <> "") Then
                    .Cells(i, 34) = rstRecordset2.Fields("JAN")
                Else
                    .Cells(i, 34) = ""
                End If


                If Trim((rstRecordset2.Fields("FEB")) <> "") Then
                    .Cells(i, 35) = rstRecordset2.Fields("FEB")
                Else
                    .Cells(i, 35) = ""
                End If


                If (Trim(rstRecordset2.Fields("MAR")) <> "") Then
                    .Cells(i, 36) = rstRecordset2.Fields("MAR")
                Else
                    .Cells(i, 36) = ""
                End If

                If (Trim(rstRecordset2.Fields("APR")) <> "") Then
                    .Cells(i, 37) = rstRecordset2.Fields("APR")
                Else
                    .Cells(i, 37) = ""
                End If

                If (Trim(rstRecordset2.Fields("MAY")) <> "") Then
                    .Cells(i, 38) = rstRecordset2.Fields("MAY")
                Else
                    .Cells(i, 38) = ""
                End If

                If (Trim(rstRecordset2.Fields("JUN")) <> "") Then
                    .Cells(i, 39) = rstRecordset2.Fields("JUN")
                Else
                    .Cells(i, 39) = ""
                End If

                If (Trim(rstRecordset2.Fields("JUL")) <> "") Then
                    .Cells(i, 40) = rstRecordset2.Fields("JUL")
                Else
                    .Cells(i, 40) = ""
                End If

                If (Trim(rstRecordset2.Fields("AUG")) <> "") Then
                    .Cells(i, 41) = rstRecordset2.Fields("AUG")
                Else
                    .Cells(i, 41) = ""
                End If

                If (Trim(rstRecordset2.Fields("SEP")) <> "") Then
                    .Cells(i, 42) = rstRecordset2.Fields("SEP")
                Else
                    .Cells(i, 42) = ""
                End If

                If (Trim(rstRecordset2.Fields("OCT")) <> "") Then
                    .Cells(i, 43) = rstRecordset2.Fields("OCT")
                Else
                    .Cells(i, 43) = ""
                End If

                If (Trim(rstRecordset2.Fields("NOV")) <> "") Then
                    .Cells(i, 44) = rstRecordset2.Fields("NOV")
                Else
                    .Cells(i, 44) = ""
                End If


                If (Trim(rstRecordset2.Fields("DEC")) <> "") Then
                    .Cells(i, 45) = rstRecordset2.Fields("DEC")
                Else
                    .Cells(i, 45) = ""
                End If

                .Cells(i, 46).Value = Format(FetchValue(cnConn, "Do not remove1", "Roles", "FTE NE - MI", "Project Code", rstRecordset.Fields("Project Code")), "$#,##0") ' FTE NE - MI Total ($)
                .Cells(i, 47).Value = Format(FetchValue(cnConn, "Do not remove2", "Roles", "Plan - MI", "Project Code", rstRecordset.Fields("Project Code")), "$#,##0") ' Budget - MI Total ($)

                If Not gbFY17 Then
                    .Cells(i, 48).Value = Format(FetchValue(cnConn, "MI Labor Cost - Revised BL", "Roles", "Actual - HW < $5K", "Project Code", rstRecordset.Fields("Project Code")), "$#,##0") ' MI Labor Cost - Revised BL($)
                End If

                If .Cells(i, 48).Value = "" Then           ' If found no value in MI Labor Cost - Revised BL, then bring Budget - MI Total ($)
                    .Cells(i, 48).Value = Format(.Cells(i, 47).Value, "$#,##0")
                Else
                    .Cells(i, 48).Font.Color = RGB(156, 0, 6)
                End If


                If Not gbFY17 Then
                    .Cells(i, 49).Value = FetchValue(cnConn, "Report", "Roles", "'" + "Actual - MI" + "'", "Project Code", rstRecordset.Fields("Project Code")) ' Budget / CR Notes
                    .Cells(i, 50).Value = FetchValue(cnConn, "NE Explanation", "Roles", "Actual - HW < $5K", "Project Code", rstRecordset.Fields("Project Code")) 'NE Explanation
                    .Cells(i, 50).Font.Color = vbBlue
                End If
                i = i + 1
            End If


            '''''Debug.Print rstRecordset.fields("Names")
            rstRecordset.MoveNext


        Wend

        .Range(.Cells(2, 34), .Cells(i - 1, 45)).NumberFormat = "#0.00"


        .Columns("A:AV").AutoFit

        .Range("A:C").ColumnWidth = 10

        .Range("E:H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 13
        .Columns("J").ColumnWidth = 15
        .Columns("K").ColumnWidth = 13
        .Columns("L").ColumnWidth = 15

        .Columns("M").ColumnWidth = 15
        .Columns("N").ColumnWidth = 13

        .Columns("O").ColumnWidth = 15
        .Columns("P").ColumnWidth = 13

        .Columns("Q").ColumnWidth = 15
        .Columns("R").ColumnWidth = 13

        .Columns("S").ColumnWidth = 15
        .Columns("T").ColumnWidth = 13

        ''.Columns("O").ColumnWidth = 8



        .Columns("U").ColumnWidth = 12
        .Columns("V").ColumnWidth = 12
        .Columns("X").ColumnWidth = 10
        .Range("Y:AB").ColumnWidth = 12
        .Range("AC:AF").ColumnWidth = 8
        .Range("AH:AS").ColumnWidth = 8
        .Columns("AT:AV").ColumnWidth = 13
        .Columns("AW:AX").ColumnWidth = 60

        '.Range(.Cells(1, 1), .Cells(1, 50)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        '.Range(.Cells(2, 2), .Cells(2, 50)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Dim j As Long
        For j = 1 To (i - 1)
            .Range(.Cells(j, 1), .Cells(j, 50)).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(j, 1), .Cells(j, 50)).Borders(xlEdgeBottom).Color = RGB(80, 130, 190)
        Next

        Dim k As Long
        For k = 1 To 51
            .Range(.Cells(1, k), .Cells(i - 1, k)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(1, k), .Cells(i - 1, k)).Borders(xlEdgeLeft).Color = RGB(80, 130, 190)

        Next

        '''''''.Columns("A").ColumnWidth = 1


        ''.Range(.Cells(2, 16), .Cells(i - 1, 16)).HorizontalAlignment = xlCenter        ' Calculated
        ''.Range(.Cells(2, 17), .Cells(i - 1, 17)).HorizontalAlignment = xlCenter        ' Dashboard
        .Range(.Cells(2, 34), .Cells(i - 1, 34)).HorizontalAlignment = xlCenter 'Publish (Y/N)




        ' Removing WrapText for Budget / CR Notes
        .Range(.Cells(2, 49), .Cells(i - 1, 49)).WrapText = False
        .Range(.Cells(2, 50), .Cells(i - 1, 50)).WrapText = False


        ' Dashboard NE column formatting
        With .Range(.Cells(2, 7), .Cells(i - 1, 7))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=($G2/$F2) > 1.1"
            .FormatConditions.Add Type:=xlExpression, Formula1:="=($G2/$F2) > 1.05"

            With .FormatConditions(1)
                .SetFirstPriority
                .Interior.Color = RGB(255, 199, 206)       ' Redish
                .Font.Color = RGB(156, 0, 6)
            End With

            With .FormatConditions(2)
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
            End With
        End With

        ' Dashboard NE - Budget % column formatting
        With .Range(.Cells(2, 9), .Cells(i - 1, 9))

            .FormatConditions.Add Type:=xlExpression, Formula1:="=$I2 = """""
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$I2 > 0.05"
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$I2 > 0"


            With .FormatConditions(1)
                .SetFirstPriority
            End With

            With .FormatConditions(2)
                .Interior.Color = RGB(255, 199, 206)       ' Redish
                .Font.Color = RGB(156, 0, 6)
            End With

            With .FormatConditions(3)
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
            End With

        End With

        ' Dashboard NE - Revised BL % column formatting
        With .Range(.Cells(2, 11), .Cells(i - 1, 11))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$K2 = """""
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$K2 > 0.05"
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$K2 > 0"


            With .FormatConditions(1)
                .SetFirstPriority
            End With


            With .FormatConditions(2)
                .Interior.Color = RGB(255, 199, 206)       ' Redish
                .Font.Color = RGB(156, 0, 6)
            End With

            With .FormatConditions(3)
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
            End With
        End With


        ' Reporting NE - Budget % column formatting
        With .Range(.Cells(2, 14), .Cells(i - 1, 14))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$N2 = """""
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$N2 > 0.05"
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$N2 > 0"


            With .FormatConditions(1)
                .SetFirstPriority
            End With


            With .FormatConditions(2)
                .Interior.Color = RGB(255, 199, 206)       ' Redish
                .Font.Color = RGB(156, 0, 6)
            End With

            With .FormatConditions(3)
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
            End With
        End With


        ' Reporting NE - Revised BL % column formatting
        With .Range(.Cells(2, 16), .Cells(i - 1, 16))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2 = """""
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2 > 0.05"
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$P2 > 0"


            With .FormatConditions(1)
                .SetFirstPriority
            End With


            With .FormatConditions(2)
                .Interior.Color = RGB(255, 199, 206)       ' Redish
                .Font.Color = RGB(156, 0, 6)
            End With

            With .FormatConditions(3)
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
            End With
        End With



        ' Reporting NE - Dashboard NE % column formatting
        With .Range(.Cells(2, 18), .Cells(i - 1, 18))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$R2 = """""
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$R2 > 0.05"
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$R2 > 0"


            With .FormatConditions(1)
                .SetFirstPriority
            End With


            With .FormatConditions(2)
                .Interior.Color = RGB(255, 199, 206)       ' Redish
                .Font.Color = RGB(156, 0, 6)
            End With

            With .FormatConditions(3)
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)
            End With
        End With



        ' Project Status (calculated) column formatting
        With .Range(.Cells(2, 21), .Cells(i - 1, 21))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$U2 =""RED"""
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$U2 =""YELLOW"""
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$U2 =""GREEN"""

            With .FormatConditions(1)
                .SetFirstPriority
                .Interior.Color = RGB(255, 199, 206)       ' Redish
                .Font.Color = RGB(156, 0, 6)
            End With

            With .FormatConditions(2)
                .Interior.Color = RGB(255, 235, 156)       ' Yellow-ish
                .Font.Color = RGB(156, 101, 0)
            End With

            With .FormatConditions(3)
                .Interior.Color = RGB(198, 239, 206)       ' Green-ish
                .Font.Color = RGB(0, 97, 0)
            End With

        End With


        ' Project Status (Dashboard) column formatting
        With .Range(.Cells(2, 22), .Cells(i - 1, 22))
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$V2 =""RED"""
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$V2 =""YELLOW"""
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$V2 =""GREEN"""

            With .FormatConditions(1)
                .SetFirstPriority
                .Interior.Color = RGB(255, 199, 206)       ' Redish
                .Font.Color = RGB(156, 0, 6)
            End With

            With .FormatConditions(2)
                .Interior.Color = RGB(255, 235, 156)       ' Yellow-ish
                .Font.Color = RGB(156, 101, 0)
            End With

            With .FormatConditions(3)
                .Interior.Color = RGB(198, 239, 206)       ' Green-ish
                .Font.Color = RGB(0, 97, 0)
            End With

        End With

        .Range("A1:AG1").AutoFilter



        Range("E3").Select
        ActiveWindow.FreezePanes = True



        ' Retrieve a reference to the used range:
        Dim rng As Range
        Set rng = .Range(.Cells(2, 1), .Cells(i - 1, 50))

        ' Create sort:
        Dim srt As Sort
        ' Include these two lines to make sure you get
        ' IntelliSense help as you work with the Sort object:
        Dim sht As Worksheet
        Set sht = ActiveSheet

        Set srt = sht.Sort

        ' Sort first by state ascending, and then by age descending.
        srt.SortFields.Clear
        srt.SortFields.Add Key:=Columns("A"), _
        SortOn:=xlSortOnValues, Order:=xlAscending
        srt.SortFields.Add Key:=Columns("B"), _
        SortOn:=xlSortOnValues, Order:=xlAscending
        ' Set the sort range:
        srt.SetRange rng
        srt.Header = xlNo
        srt.MatchCase = True
        ' Apply the sort:
        srt.Apply
    End With



    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


    '''''wbNew.SaveAs Filename:=gwsConfig.Range(gsLOCAL_FOLDER).Value & "\" & sD4AFileName

    '''''Set wbNew = Nothing

    cnConn.Close

    Set cnConn = Nothing
    Set cmdCommand = Nothing
    Set rstRecordset = Nothing


    Set cmdCommand2 = Nothing
    Set rstRecordset2 = Nothing


    wsPT.Name = "Portfolio Project Data"
    Gen_NES = 0                                            '???????????????????


    ''''    Set rstRecordset = Nothing



    'Invoke Pivot Table process
    ' Aca pongo un if...se genero Portfolio DAta, etc, etc.

    Dim algo As Long

    ''''''''???????Call copy2tmp(wsPT)

    If bPivot Then
        algo = GenPT_NES(wbNew, wsPT)
    End If



ErrorExit:
    ' Clean up
    Exit Function

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function

'---------------------------------------------------------------------------------------
' Method : LOBsCheck
' Author : cklahr
' Date   : 9/11/2016
' Purpose: xxxxxxxxxxx
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function LOBsCheck(ByVal sLobField As String, ByVal sLobs2check As String, ByVal lChecktype As Long) As Boolean

    Dim i As Long
    Dim sLOBsArray() As String

    sLOBsArray = Split(sLobs2check, ";")

    If (lChecktype = pos) Then
        For i = 0 To UBound(sLOBsArray)
            'If InStr(1, UCase(Trim(sLobField)), UCase(sLOBsArray(i))) > 0 Then
            If UCase(Trim(sLobField)) = UCase(sLOBsArray(i)) Then
                LOBsCheck = True
                Exit Function
            End If
        Next i
        LOBsCheck = False
        Exit Function
    Else
        For i = 0 To UBound(sLOBsArray)
            'If InStr(1, UCase(Trim(sLobField)), UCase(sLOBsArray(i))) > 0 Then
            If UCase(Trim(sLobField)) = UCase(sLOBsArray(i)) Then
                LOBsCheck = False
                Exit Function
            End If
        Next i
        LOBsCheck = True
        Exit Function

    End If

End Function

'---------------------------------------------------------------------------------------
' Method : VarianceGradient
' Author : cklahr
' Date   : 10/1/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub VarianceGradient(ByRef ws As Worksheet, ByVal lCol As Long, ByVal lLastRow As Long, ByVal dMaxNeg As Double, ByVal dMaxPos As Double)
    Dim i As Long
    Dim dVariance As Double
    Dim dPosQuarter As Double

    If dMaxPos > MINPOSVAR Then
        dPosQuarter = (dMaxPos - MINPOSVAR) / 4

        With ws
            For i = 3 To lLastRow
                If .Cells(i, lCol) <> "" And .Cells(i, lCol + 1) <> "" Then 'lCol + 1 should be a Constant! (col 7)
                    dVariance = .Cells(i, lCol)
                    If UCase(Left(.Cells(i, 1), 5)) <> "TOTAL" Then
                        If dVariance > 0 Then
                            Select Case True
                                Case (dVariance > MINPOSVAR) And (dVariance <= dPosQuarter)
                                    .Cells(i, lCol).Interior.Color = RGB(255, 193, 193)

                                Case (dVariance > dPosQuarter) And (dVariance <= 2 * dPosQuarter)
                                    .Cells(i, lCol).Interior.Color = RGB(255, 129, 129)

                                Case (dVariance > 2 * dPosQuarter) And (dVariance <= 3 * dPosQuarter)
                                    .Cells(i, lCol).Interior.Color = RGB(255, 97, 97)

                                Case (dVariance > 3 * dPosQuarter)
                                    .Cells(i, lCol).Interior.Color = RGB(255, 0, 0)
                                    .Cells(i, lCol).Font.Color = vbWhite
                            End Select
                        Else
                            Select Case True
                                Case (dVariance < MINNEGVAR) And (dVariance >= -1 * dPosQuarter)
                                    .Cells(i, lCol).Interior.Color = RGB(255, 255, 183) 'Yellow gradient
                                    '.Cells(i, lCol).Interior.Color = RGB(196, 252, 188)        'Green gradient

                                Case (dVariance < -1 * dPosQuarter) And (dVariance >= -2 * dPosQuarter)
                                    .Cells(i, lCol).Interior.Color = RGB(255, 255, 117) 'Yellow gradient
                                    '.Cells(i, lCol).Interior.Color = RGB(127, 249, 111)        'Green gradient

                                Case (dVariance < -2 * dPosQuarter) And (dVariance >= -3 * dPosQuarter)
                                    .Cells(i, lCol).Interior.Color = RGB(255, 255, 71) 'Yellow gradient
                                    '.Cells(i, lCol).Interior.Color = RGB(42, 245, 15)        'Green gradient

                                Case (dVariance < -3 * dPosQuarter)
                                    .Cells(i, lCol).Interior.Color = RGB(255, 255, 0) 'Yellow gradient
                                    '.Cells(i, lCol).Interior.Color = RGB(26, 165, 7)        'Green gradient

                            End Select
                        End If
                    End If
                End If

            Next i
        End With
    End If


End Sub
'---------------------------------------------------------------------------------------
' Method : ListWithFilter
' Author : cklahr
' Date   : 9/11/2016
' Purpose: xxxxxxxxxxx
' Arguments:
' Pending:
' Comments: PASARLO A FUNCTION???
'---------------------------------------------------------------------------------------
Sub ListWithFilter(ByRef ws As Worksheet, ByVal cnConn As Object, ByVal rst As Object, ByRef i As Long, ByVal sLOBs As String, ByVal Checktype As Long, ByRef dMaxNegVar As Double, ByRef dMaxPosVar As Double)
    With ws
        While Not rst.EOF

            ''''.Sheets("Sheet1").Cells(i, 1).Value = CDbl(Trim(rst.fields("Names")))
            If Len(rst.Fields("Project Code")) > 1 And UCase(rst.Fields("Project Code")) <> "WHATIF" Then

                If (LOBsCheck(rst.Fields("LOB"), sLOBs, Checktype)) Then
                    .Cells(i, 1).Value = Trim(rst.Fields("LOB"))

                    If gbFY17 Then
                        .Cells(i, 2).Value = Trim(rst.Fields("Budget Category"))
                    Else
                        .Cells(i, 2).Value = Trim(rst.Fields("Project Type"))
                    End If
                    ' .Cells(i, 3).Value = Trim(rst.fields("Project Code"))

                    .Cells(i, 3).Value = Trim(rst.Fields("Project Name"))

                    If gbFY17 Then
                        .Cells(i, 4).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - SW SAAS", "Project Code", rst.Fields("Project Code"))), "$#,##0") ' Budget
                    Else
                        .Cells(i, 4).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW Amort", "Project Code", rst.Fields("Project Code"))), "$#,##0") ' Budget
                    End If

                    .Cells(i, 5).Value = Format(FetchValue(cnConn, "Do not remove1", "Roles", "NE (Actual + Rem. Plan)", "Project Code", rst.Fields("Project Code")), "$#,##0") ' Dashboard NE

                    ''.Cells(i, 8).Value = "=G" & i & "-E" & i


                    .Cells(i, 6) = "=IF((E" & i & "-D" & i & ")<>0,E" & i & "-D" & i & ","""")" 'Dashboard NE - Budget
                    .Cells(i, 6).NumberFormat = "$#,##0"


                    If .Cells(i, 6) <> "" Then
                        If .Cells(i, 6) > 0 And .Cells(i, 6) > dMaxPosVar Then dMaxPosVar = .Cells(i, 6)
                        If .Cells(i, 6) < 0 And .Cells(i, 6) < dMaxNegVar Then dMaxNegVar = .Cells(i, 6)
                    End If

                    ' .Cells(i, 9) = "=IF(E" & i & "=0,-100%,H" & i & "/E" & i & ")"
                    .Cells(i, 7) = "=IF(D" & i & "=0,-100%,IF(F" & i & "<>"""",F" & i & "/D" & i & ",""""))"



                    '=IF(E4=0,-100%,IF(H4<>"",H4/E4,""))

                    .Cells(i, 7).NumberFormat = "0%"


                    .Cells(i, 8).Value = Format(FetchValue(cnConn, "Names", "Roles", "Actual - Travel", "Project Code", rst.Fields("Project Code")), "$#,##0") 'YTD Actuals

                    .Cells(i, 9).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW Maint", "Project Code", rst.Fields("Project Code"))), "$#,##0") ' Reporting NE



                    ' If Actuals greater than Reporting NE, then overwrite Reporting NE with Actuals and highlith in redish
                    If .Cells(i, 8).Value > .Cells(i, 9).Value Then
                        .Cells(i, 9).AddComment ("Overwritten with Actuals (Original Reporting NE: " + CStr(Format(.Cells(i, 9).Value, "$#,##0")) + ")")
                        .Cells(i, 9).Value = Format(.Cells(i, 8).Value, "$#,##0")
                        .Cells(i, 9).Font.Color = RGB(156, 0, 6)
                    End If

                    If gbFY17 Then
                        ''''' NEED TO DEVELOP THIS FOR 2017 DASHBOARD!!!!!!!!!
                    Else
                        .Cells(i, 10).Value = FetchValue(cnConn, "Report", "Roles", "'" + "Actual - MI" + "'", "Project Code", rst.Fields("Project Code")) ' Budget / CR Notes
                        .Cells(i, 11).Value = FetchValue(cnConn, "NE Explanation", "Roles", "Actual - HW < $5K", "Project Code", rst.Fields("Project Code")) 'NE Explanation
                        .Cells(i, 11).Font.Color = vbBlue
                    End If

                    i = i + 1
                End If
            End If

            '''''Debug.Print rst.fields("Names")
            rst.MoveNext


        Wend

    End With

End Sub


'---------------------------------------------------------------------------------------
' Method : ListExternals
' Author : cklahr
' Date   : 10/1/2016
' Purpose:
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Sub ListExternals(ByRef ws As Worksheet, ByVal cnConn As Object, ByVal rst As Object, ByRef i As Long, ByRef dMaxNegVar As Double, ByRef dMaxPosVar As Double)
    With ws
        While Not rst.EOF
            If Len(rst.Fields("Project Code")) > 1 And UCase(rst.Fields("Project Code")) <> "WHATIF" Then
                If (Left(rst.Fields("CAT"), 7) = "9 - EXT") Then
                    .Cells(i, 1).Value = Trim(rst.Fields("LOB"))

                    If gbFY17 Then
                        .Cells(i, 2).Value = Trim(rst.Fields("Budget Category"))
                    Else
                        .Cells(i, 2).Value = Trim(rst.Fields("Project Type"))
                    End If
                    .Cells(i, 3).Value = Trim(rst.Fields("Project Name"))

                    If gbFY17 Then
                        .Cells(i, 4).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - SW SAAS", "Project Code", rst.Fields("Project Code"))), "$#,##0") ' Budget
                    Else
                        .Cells(i, 4).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW Amort", "Project Code", rst.Fields("Project Code"))), "$#,##0") ' Budget
                    End If

                    .Cells(i, 5).Value = Format(FetchValue(cnConn, "Do not remove1", "Roles", "NE (Actual + Rem. Plan)", "Project Code", rst.Fields("Project Code")), "$#,##0") ' Dashboard NE
                    .Cells(i, 6) = "=IF((E" & i & "-D" & i & ")<>0,E" & i & "-D" & i & ","""")" 'Dashboard NE - Budget
                    .Cells(i, 6).NumberFormat = "$#,##0"

                    If .Cells(i, 6) <> "" Then
                        If .Cells(i, 6) > 0 And .Cells(i, 6) > dMaxPosVar Then dMaxPosVar = .Cells(i, 6)
                        If .Cells(i, 6) < 0 And .Cells(i, 6) < dMaxNegVar Then dMaxNegVar = .Cells(i, 6)
                    End If

                    ' No % for externals!
                    '.Cells(i, 7) = "=IF(D" & i & "=0,-100%,IF(F" & i & "<>"""",F" & i & "/D" & i & ",""""))"
                    '.Cells(i, 7).NumberFormat = "0%"

                    .Cells(i, 8).Value = Format(FetchValue(cnConn, "Names", "Roles", "Actual - Travel", "Project Code", rst.Fields("Project Code")), "$#,##0") 'YTD Actuals
                    .Cells(i, 9).Value = Format((FetchValue(cnConn, "Names", "Roles", "Actual - HW Maint", "Project Code", rst.Fields("Project Code"))), "$#,##0") ' Reporting NE

                    ' If Actuals greater than Reporting NE, then overwrite Reporting NE with Actuals and highlith in redish
                    If .Cells(i, 8).Value > .Cells(i, 9).Value Then
                        .Cells(i, 9).AddComment ("Overwritten with Actuals (Original Reporting NE: " + CStr(Format(.Cells(i, 9).Value, "$#,##0")) + ")")
                        .Cells(i, 9).Value = Format(.Cells(i, 8).Value, "$#,##0")
                        .Cells(i, 9).Font.Color = RGB(156, 0, 6)
                    End If

                    If gbFY17 Then
                        ''''' NEED TO DEVELOP THIS FOR 2017 DASHBOARD!!!!!!!!!
                    Else
                        .Cells(i, 10).Value = FetchValue(cnConn, "Report", "Roles", "'" + "Actual - MI" + "'", "Project Code", rst.Fields("Project Code")) ' Budget / CR Notes
                        .Cells(i, 11).Value = FetchValue(cnConn, "NE Explanation", "Roles", "Actual - HW < $5K", "Project Code", rst.Fields("Project Code")) 'NE Explanation
                        .Cells(i, 11).Font.Color = vbBlue
                    End If

                    i = i + 1
                End If
            End If

            rst.MoveNext
        Wend
    End With
End Sub

'---------------------------------------------------------------------------------------
' Method : Gen_NES_Short
' Author : cklahr
' Date   : 2/11/2016
' Purpose: Produce NE Summary
' Arguments:
' Pending:
' Comments:
'---------------------------------------------------------------------------------------
Function Gen_NES_Short_GC(ByVal wbNew As Workbook, ByVal sFilePP As String) As Double
    Const sSOURCE As String = "ProduceNES"

    Dim PC As Object
    Dim sdConPP As WorkbookConnection
    Dim sdConRS As WorkbookConnection
    Dim sConnPP As String
    Dim sConnRS As String
    Dim sCommandTextPP As String
    Dim sCommandTextRS As String

    Dim dGrandTotalFTE_NE As Double
    Dim dGrandTotalPROJ_SUMMARY As Double
    Dim dGrandTotalCC_CAPACITY As Double

    Dim dMaxNegVar As Double
    Dim dMaxPosVar As Double

    Dim dRoundUp As Double
    Dim dRoundDown As Double

    Dim wsPT As Worksheet


    Dim sPPDB As String                                    'Portfolio Plan DB
    Dim cnConn As Object
    Dim cmdCommand As Object

    Dim cmdCommand2 As Object
    Dim rstRecordset2 As Object

    Dim lStart As Long
    Dim i As Long

    Dim srt As Sort


    'Dim sD4AFileName As String
    '   Dim wbNew As Workbook
    '
    '  sD4AFileName = "D4A_" + Format(Now(), "mmmddyyyy_hhmmss") + ".xlsx"
    ' Set wbNew = Workbooks.Add

    'wbNew.Activate


    'wbNew is the D4A workbook...
    Set wsPT = wbNew.Sheets.Add
    wbNew.Windows(1).DisplayGridlines = False

    Application.ErrorCheckingOptions.InconsistentFormula = False



    '''sPPDB = gwsConfig.Range(gsLOCAL_FOLDER).Value + "\" + gwsConfig.Range(gsDB_NAME).Value        'PDASH.accdb

    sPPDB = gsLocal_Folder + "\" + gwsConfig.Range(gsDB_NAME).Value 'PDASH.accdb

    Set cnConn = CreateObject("ADODB.Connection")
    With cnConn
        .ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0"
        .Open sPPDB
    End With

    ' Set the command text.

    Set cmdCommand = CreateObject("ADODB.Command")
    Set cmdCommand.activeconnection = cnConn
    With cmdCommand
        '''''@@@@@@@@ .CommandText = "SELECT [Names] FROM tbl_PortfolioPlan WHERE [Roles] = 'Actual - HW Amort'" 'Budget

        If gbFY17 Then
            .CommandText = "SELECT DISTINCT [Project Code], [Project Name], [LOB], [Account Manager], [Activation Status], [Budget Category], [Publish NE], [CAT] FROM tbl_PortfolioPlan"
        Else
            .CommandText = "SELECT DISTINCT [Project Code], [Project Name], [LOB], [Account Manager], [Activation Status], [Project Type], [Publish NE], [CAT] FROM tbl_PortfolioPlan"
        End If
        .CommandType = adCmdText
        .Execute
    End With

    ' Open the recordset.
    Dim rstRecordset As Object
    Set rstRecordset = CreateObject("ADODB.Recordset")
    Set rstRecordset.activeconnection = cnConn
    rstRecordset.Open cmdCommand


    Set cmdCommand2 = CreateObject("ADODB.Command")

    Set cmdCommand2.activeconnection = cnConn


    ' Create a PivotTable cache and report.
    ''''''''''''''' Set PC = wbNew.PivotCaches.Create(SourceType:=xlExternal)
    '''''''''''''''Set PC.Recordset = rstRecordset

    dMaxNegVar = 0
    dMaxPosVar = 0

    i = 3
    With wsPT
        .Cells(1, 1) = "LOB"
        If gbFY17 Then
            Cells(1, 2) = "Budget Category"
        Else
            .Cells(1, 2) = "Project Type"
        End If
        '.Cells(1, 3) = "Project Code"
        .Cells(1, 3) = "Project Name"
        .Cells(1, 4) = "Budget"
        .Cells(2, 4) = "($)"

        .Cells(1, 5) = "Dashboard NE"
        ' .Cells(1, 5).AddComment ("- Yellow: > 5% compared to Revised BL" + vbCrLf + "- Red: > 10% compared to Revised BL")
        .Cells(2, 5) = "($)"

        .Cells(1, 6) = "Dashboard NE - Budget"
        .Range(.Cells(1, 6), .Cells(1, 7)).Merge
        .Cells(2, 6) = "($)"
        .Cells(2, 7) = "(%)"
        ' .Cells(2, 7).AddComment ("- Yellow: > 0% variance compared to Budget" + vbCrLf + "- Red: > 5% variance compared to Budget")

        .Cells(1, 8) = "YTD Actuals"
        .Cells(2, 8) = "($)"


        .Cells(1, 9) = "Reporting NE"
        .Cells(2, 9) = "($)"

        .Cells(1, 10) = "Budget / CR Notes"
        .Cells(1, 11) = "NE Explanation"

        .Range(.Cells(1, 1), .Cells(2, 11)).Font.Bold = True
        .Range(.Cells(1, 1), .Cells(2, 11)).WrapText = True
        .Range(.Cells(1, 1), .Cells(2, 11)).HorizontalAlignment = xlCenter
        .Range(.Cells(1, 1), .Cells(2, 11)).VerticalAlignment = xlTop
        .Range(.Cells(1, 1), .Cells(2, 11)).Interior.Color = RGB(80, 130, 190) ' RGB(238, 236, 225)
        .Range(.Cells(1, 1), .Cells(2, 11)).Font.Color = vbWhite ' RGB(31, 73, 125)

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        Set srt = .Sort

        lStart = i
        Call ListWithFilter(wsPT, cnConn, rstRecordset, i, "1117  MI - Shareholder’s Services", pos, dMaxNegVar, dMaxPosVar)
        If i > lStart Then
            .Cells(i, 1) = "Total Shareholder's Services >>>"
            .Cells(i, 1).Font.Bold = True
            .Cells(i, 4) = "=SUM(D" & CStr(lStart) + ":D" + CStr(i - 1) + ")"
            .Cells(i, 4).Font.Bold = True

            .Cells(i, 5) = "=SUM(E" & CStr(lStart) + ":E" + CStr(i - 1) + ")"
            .Cells(i, 5).Font.Bold = True

            .Cells(i, 6) = "=SUM(F" & CStr(lStart) + ":F" + CStr(i - 1) + ")"
            .Cells(i, 6).Font.Bold = True

            .Cells(i, 8) = "=SUM(H" & CStr(lStart) + ":H" + CStr(i - 1) + ")"
            .Cells(i, 8).Font.Bold = True

            ' Sort first by state ascending, and then by age descending.
            srt.SortFields.Clear
            srt.SortFields.Add Key:=Columns("A"), SortOn:=xlSortOnValues, Order:=xlAscending
            srt.SortFields.Add Key:=Columns("B"), SortOn:=xlSortOnValues, Order:=xlAscending

            ' Set the sort range:
            srt.SetRange .Range(.Cells(lStart, 1), .Cells(i, 11))

            srt.Header = xlNo
            srt.MatchCase = True
            ' Apply the sort:
            srt.Apply

            i = i + 1
            lStart = i
        End If


        rstRecordset.moveFirst
        Call ListWithFilter(wsPT, cnConn, rstRecordset, i, "1115  MI - Products", pos, dMaxNegVar, dMaxPosVar)

        If i > lStart Then
            .Cells(i, 1) = "Total Products >>>"
            .Cells(i, 1).Font.Bold = True
            .Cells(i, 4) = "=SUM(D" & CStr(lStart) + ":D" + CStr(i - 1) + ")"
            .Cells(i, 4).Font.Bold = True

            .Cells(i, 5) = "=SUM(E" & CStr(lStart) + ":E" + CStr(i - 1) + ")"
            .Cells(i, 5).Font.Bold = True

            .Cells(i, 6) = "=SUM(F" & CStr(lStart) + ":F" + CStr(i - 1) + ")"
            .Cells(i, 6).Font.Bold = True

            .Cells(i, 8) = "=SUM(H" & CStr(lStart) + ":H" + CStr(i - 1) + ")"
            .Cells(i, 8).Font.Bold = True


            ' Sort first by state ascending, and then by age descending.
            srt.SortFields.Clear
            srt.SortFields.Add Key:=Columns("A"), SortOn:=xlSortOnValues, Order:=xlAscending
            srt.SortFields.Add Key:=Columns("B"), SortOn:=xlSortOnValues, Order:=xlAscending

            ' Set the sort range:
            srt.SetRange .Range(.Cells(lStart, 1), .Cells(i, 11))

            srt.Header = xlNo
            srt.MatchCase = True
            ' Apply the sort:
            srt.Apply

            i = i + 1
            lStart = i
        End If


        rstRecordset.moveFirst
        Call ListExternals(wsPT, cnConn, rstRecordset, i, dMaxNegVar, dMaxPosVar) ' Make it function!!!!
        '''Call ListWithFilter(wsPT, cnConn, rstRecordset, i, "LifeCo;IG", POS)

        If i > lStart Then
            .Cells(i, 1) = "Total Externals >>>"
            .Cells(i, 1).Font.Bold = True
            .Cells(i, 4) = "=SUM(D" & CStr(lStart) + ":D" + CStr(i - 1) + ")"
            .Cells(i, 4).Font.Bold = True

            .Cells(i, 5) = "=SUM(E" & CStr(lStart) + ":E" + CStr(i - 1) + ")"
            .Cells(i, 5).Font.Bold = True

            .Cells(i, 6) = "=SUM(F" & CStr(lStart) + ":F" + CStr(i - 1) + ")"
            .Cells(i, 6).Font.Bold = True

            .Cells(i, 8) = "=SUM(H" & CStr(lStart) + ":H" + CStr(i - 1) + ")"
            .Cells(i, 8).Font.Bold = True

            ' Sort first by state ascending, and then by age descending.
            srt.SortFields.Clear
            srt.SortFields.Add Key:=Columns("A"), SortOn:=xlSortOnValues, Order:=xlAscending
            srt.SortFields.Add Key:=Columns("B"), SortOn:=xlSortOnValues, Order:=xlAscending

            ' Set the sort range:
            srt.SetRange .Range(.Cells(lStart, 1), .Cells(i, 11))

            srt.Header = xlNo
            srt.MatchCase = True
            ' Apply the sort:
            srt.Apply

            i = i + 1
            lStart = i
        End If

        rstRecordset.moveFirst
        Call ListWithFilter(wsPT, cnConn, rstRecordset, i, "1117  MI - Shareholder’s Services;1115  MI - Products;LifeCo;IG", NEG, dMaxNegVar, dMaxPosVar)

        If i > lStart Then
            .Cells(i, 1) = "Total Remaining >>>"
            .Cells(i, 1).Font.Bold = True
            .Cells(i, 4) = "=SUM(D" & CStr(lStart) + ":D" + CStr(i - 1) + ")"
            .Cells(i, 4).Font.Bold = True

            .Cells(i, 5) = "=SUM(E" & CStr(lStart) + ":E" + CStr(i - 1) + ")"
            .Cells(i, 5).Font.Bold = True

            .Cells(i, 6) = "=SUM(F" & CStr(lStart) + ":F" + CStr(i - 1) + ")"
            .Cells(i, 6).Font.Bold = True

            .Cells(i, 8) = "=SUM(H" & CStr(lStart) + ":H" + CStr(i - 1) + ")"
            .Cells(i, 8).Font.Bold = True

            ' Sort first by state ascending, and then by age descending.
            srt.SortFields.Clear
            srt.SortFields.Add Key:=Columns("A"), SortOn:=xlSortOnValues, Order:=xlAscending
            srt.SortFields.Add Key:=Columns("B"), SortOn:=xlSortOnValues, Order:=xlAscending

            ' Set the sort range:
            srt.SetRange .Range(.Cells(lStart, 1), .Cells(i, 11))

            srt.Header = xlNo
            srt.MatchCase = True
            ' Apply the sort:
            srt.Apply


            i = i + 1
            lStart = i
        End If


        ''''''''.Range(.Cells(3, 34), .Cells(i - 1, 45)).NumberFormat = "#0.00"


        .Columns("A:K").AutoFit

        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 50

        .Range("D:F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 8

        .Columns("H:I").ColumnWidth = 15
        .Columns("J:K").ColumnWidth = 60


        .Range(.Cells(1, 1), .Cells(1, 11)).Borders(xlEdgeTop).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(1, 11)).Borders(xlEdgeTop).Color = RGB(150, 180, 215)
        .Range(.Cells(1, 1), .Cells(1, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(1, 11)).Borders(xlEdgeBottom).Color = RGB(150, 180, 215)
        .Range(.Cells(2, 1), .Cells(2, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range(.Cells(2, 1), .Cells(2, 11)).Borders(xlEdgeBottom).Color = RGB(150, 180, 215)

        Dim j As Long
        For j = 3 To (i - 1)
            If InStr(1, UCase(.Cells(j, 1)), "TOTAL") > 0 Then
                .Range(.Cells(j - 1, 1), .Cells(j - 1, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range(.Cells(j, 1), .Cells(j, 11)).Borders(xlEdgeBottom).LineStyle = xlDouble
                .Range(.Cells(j, 1), .Cells(j, 11)).Borders(xlEdgeBottom).Color = RGB(150, 180, 215)

                .Range(.Cells(j, 1), .Cells(j, 11)).Interior.Color = RGB(184, 204, 228) ' RGB(238, 236, 225)
                .Range(.Cells(j, 1), .Cells(j, 11)).Font.Color = vbBlack ' RGB(31, 73, 125)
            Else
                .Range(.Cells(j, 1), .Cells(j, 11)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range(.Cells(j, 1), .Cells(j, 11)).Borders(xlEdgeBottom).Color = RGB(150, 180, 215)
            End If
        Next

        Dim k As Long
        For k = 2 To 12
            .Range(.Cells(1, k), .Cells(i - 1, k)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Range(.Cells(1, k), .Cells(i - 1, k)).Borders(xlEdgeLeft).Color = RGB(150, 180, 215)
        Next

        '''''''.Columns("A").ColumnWidth = 1

        ' Removing WrapText for Budget / CR Notes
        .Range(.Cells(3, 10), .Cells(i - 1, 10)).WrapText = False
        .Range(.Cells(3, 11), .Cells(i - 1, 11)).WrapText = False



        Call VarianceGradient(wsPT, 6, i - 1, dMaxNegVar, dMaxPosVar) 'Make it function!


        ' Dashboard NE column formatting
        'With .Range(.Cells(3, 7), .Cells(i - 1, 7))
        '    .FormatConditions.Add Type:=xlExpression, Formula1:="=($G1/$F1) > 1.1"
        '    .FormatConditions.Add Type:=xlExpression, Formula1:="=($G1/$F1) > 1.05"'''

        'With .FormatConditions(1)
        '    .SetFirstPriority
        '    .Interior.Color = RGB(255, 199, 206)        ' Redish
        '    .Font.Color = RGB(156, 0, 6)
        'End With

        'With .FormatConditions(2)
        '    .Interior.Color = RGB(255, 235, 156)
        '    .Font.Color = RGB(156, 101, 0)
        'End With
        'End With


        .Range("A1:I1").AutoFilter

        ''''''''''''''.Range("D3").Select
        '''''''''''''''ActiveWindow.FreezePanes = True


        .PageSetup.Orientation = xlLandscape
        .PageSetup.PrintArea = "$A$1:$I$" + CStr(i - 1)
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False



        ' Retrieve a reference to the used range:
        Dim rng As Range
        Set rng = .Range(.Cells(3, 1), .Cells(i - 1, 11))

        ' Create sort:
        'Dim srt As Sort
        ' Include these two lines to make sure you get
        ' IntelliSense help as you work with the Sort object:
        Dim sht As Worksheet
        Set sht = ActiveSheet

        '''Set srt = sht.Sort

        ' Sort first by state ascending, and then by age descending.
        ''''srt.SortFields.Clear
        '''srt.SortFields.Add Key:=Columns("A"), _
        SortOn:=xlSortOnValues, Order:=xlAscending
        ''''srt.SortFields.Add Key:=Columns("B"), _
        SortOn:=xlSortOnValues, Order:=xlAscending
        ' Set the sort range:
        ''''srt.SetRange rng
        '''srt.Header = xlNo
        '''srt.MatchCase = True
        ' Apply the sort:
        '''srt.Apply


    End With



    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True


    '''''wbNew.SaveAs Filename:=gwsConfig.Range(gsLOCAL_FOLDER).Value & "\" & sD4AFileName

    '''''Set wbNew = Nothing

    cnConn.Close

    Set cnConn = Nothing
    Set cmdCommand = Nothing
    Set rstRecordset = Nothing


    Set cmdCommand2 = Nothing
    Set rstRecordset2 = Nothing



    wsPT.Name = "Portfolio Project Data (Short)"

    wsPT.Range("D3").Select
    ActiveWindow.FreezePanes = True

    Gen_NES_Short_GC = 0                                   '???????????????????

    ''''    Set rstRecordset = Nothing

ErrorExit:
    ' Clean up
    Exit Function

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSOURCE) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
End Function



