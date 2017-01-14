Option Explicit

Public Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Dim sNameTest As String
    Dim wsCC As Worksheet
    Dim wsCCBKP As Worksheet

    Dim sTable4Coord As String
    Dim vTable4Coord() As String

    Dim lCurrentRowTable4 As Long
    Dim lPreviousStartRowTable4 As Long
    Dim lNewStartRowTable4 As Long

    On Error Resume Next


    sNameTest = ThisWorkbook.Names("Table1").RefersTo

    If Err = 0 Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual

        Set wsCC = ThisWorkbook.Sheets("CC Capacity")
        Set wsCCBKP = ThisWorkbook.Sheets("CCBKP")        'Veryhidden


        ' Saving current location to go back after Pivot Change
        lCurrentRowTable4 = ActiveCell.Row

        ' Working with Table4 as there is no Named range for Table 3 (the Pivot Table)!!!
        sTable4Coord = wsCC.Range("Table4").Address
        sTable4Coord = Replace(sTable4Coord, ":", "")
        vTable4Coord = Split(sTable4Coord, "$")

        lPreviousStartRowTable4 = Val(vTable4Coord(2))

        '*** Starting actual updates!
        Call wipeTables(wsCC)        'Wipes T1, T2, T4
        ''Call removeFormatting  'Conditional formatting and Total row formatting for T1, T2, T4
        Call PopulateT4(wsCC, wsCCBKP)
        Call PopulateT1T2(wsCC, wsCCBKP)
        ''Call AddTotals
        Call ReapplyFormatting(wsCC, wsCCBKP)        'Conditional formatting, Total row, columns width, etc

        ' Pointing back to the place we were before the change
        ' Working with Table4 as there is no Named range for Table 3!!!
        sTable4Coord = wsCC.Range("Table4").Address
        sTable4Coord = Replace(sTable4Coord, ":", "")
        vTable4Coord = Split(sTable4Coord, "$")

        lNewStartRowTable4 = Val(vTable4Coord(2))

        Application.Goto Sheets("CC Capacity").Cells(lNewStartRowTable4 + lCurrentRowTable4 - lPreviousStartRowTable4, 2), True

        ' Always showing 10 rows above the place we were before (it works fine even with empty tables)
        ActiveWindow.ScrollRow = Val(vTable4Coord(2) - 10)


        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic

        Application.ErrorCheckingOptions.InconsistentFormula = False
    End If
End Sub
Sub CondFormat(ByVal R As Range)
    Dim cs As ColorScale

    With R
        Set cs = .FormatConditions.AddColorScale(colorscaletype:=3)
        .FormatConditions(.FormatConditions.Count).SetFirstPriority        ' Take priority over any other formats

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
Public Function lColNum(ByVal ws As Worksheet, ByVal sColLetter As String) As Long
    lColNum = ws.Range(sColLetter & 1).Column
End Function

Public Function sColLetter(ByVal ws As Worksheet, ByVal lColNum As Long) As String
    sColLetter = Split(ws.Cells(1, lColNum).Address, "$")(1)
End Function
Function wipeTables(ByVal wsCC As Worksheet)
    Dim vRng() As String
    Dim sWipeRange As String
    Dim vFrom() As String

    With wsCC

        'Table1
        vRng = Split(.Range("Table1").Address, "$")
        vFrom = Split(vRng(2), ":")
        sWipeRange = vRng(1) + CStr(Val(vFrom(0)) + 1) + ":" + vRng(3) + vRng(4)
        .Range(sWipeRange).ClearContents

        ' Deleting formatting for Total column T1
        .Range(vRng(1) + vRng(4) + ":" + vRng(3) + vRng(4)).Interior.Color = vbWhite

        'Table2
        vRng = Split(.Range("Table2").Address, "$")
        vFrom = Split(vRng(2), ":")
        sWipeRange = vRng(1) + CStr(Val(vFrom(0)) + 1) + ":" + vRng(3) + vRng(4)
        .Range(sWipeRange).ClearContents

        ' Deleting formatting for Total column T1
        .Range(vRng(1) + vRng(4) + ":" + vRng(3) + vRng(4)).Interior.Color = vbWhite


        'Table4
        vRng = Split(.Range("Table4").Address, "$")
        vFrom = Split(vRng(2), ":")
        sWipeRange = vRng(1) + CStr(Val(vFrom(0)) + 1) + ":" + vRng(3) + vRng(4)
        .Range(sWipeRange).ClearContents

        ' Deleting formatting for Total column T4
        .Range(vRng(1) + vRng(4) + ":" + vRng(3) + vRng(4)).Interior.Color = vbWhite


        ' Delete all Conditional Formatting in this worksheet
        .Cells.FormatConditions.Delete

    End With
End Function

Function PopulateT4(ByVal wsCC As Worksheet, ByVal wsCCBKP As Worksheet)
    Dim PT As PivotTable

    Dim lNumRoles As Long
    Dim lRolesFrom As Long
    Dim lRolesTo As Long

    Dim vRng() As String

    Dim vRng1() As String
    Dim vRng4() As String

    Dim vFrom4() As String

    Dim lFindFrom As Long
    Dim lFindTo As Long

    Dim i As Long

    Dim rngFindValue As Range


    Dim sFindRange As String
    Dim sAfterRange As String

    Set PT = ThisWorkbook.Sheets("CC Capacity").PivotTables("CapVsDemPT")        'Capacity vs Demand PT

    vRng = Split(PT.TableRange1.Address, "$")        'Row From: rng(2) Row To: rng(4)
    lRolesFrom = Val(vRng(2)) + 2        'Remove the first 2 rows (headers)
    lRolesTo = Val(vRng(4)) - 1        'Remove last row (footer)

    Dim vRng8() As String

    vRng8 = Split(wsCCBKP.Range("Table8").Address, "$")        'Using Table4 range from CCBKP!!!!
    lFindFrom = Val(vRng8(2))        'Not removing the first row (header) for the search!
    lFindTo = Val(vRng8(4)) - 1        'Remove last row (footer)

    sFindRange = vRng8(1) + vRng8(2) + vRng8(1) + CStr(lFindTo)        ' Using lFindTo as it has removed the last row already

    sAfterRange = vRng8(1) + vRng8(2)
    sAfterRange = Left(sAfterRange, Len(sAfterRange) - 1)

    For i = lRolesFrom To lRolesTo
        Set rngFindValue = wsCCBKP.Range(sFindRange).Find(What:=wsCC.Cells(i, 2).Value, after:=wsCCBKP.Range(sAfterRange), LookIn:=xlValues, lookAt:=xlWhole, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
        If Not (rngFindValue Is Nothing) Then

            ' Copy over this row in CC:T4
            wsCC.Range(vRng8(1) + CStr(i) + ":" + vRng8(3) + CStr(i)).Value = wsCCBKP.Range(vRng8(1) + CStr(rngFindValue.Row) + ":" + vRng8(3) + CStr(rngFindValue.Row)).Value
        End If

    Next

    lNumRoles = lRolesTo - lRolesFrom + 1

    With wsCC
        ''' Redefine T4 named range
        vRng4 = Split(.Range("Table4").Address, "$")
        vFrom4 = Split(vRng4(2), ":")

        ThisWorkbook.Names("Table4").RefersTo = .Range(vRng4(1) + vRng4(2) + vRng4(3) + CStr(Val(vFrom4(0)) + lNumRoles + 1))


        .Range(vRng4(1) + CStr(i) + ":" + vRng4(3) + CStr(i)).Interior.Color = RGB(219, 229, 241)        ' i already contains the next row after the last filtered row!
        .Range(vRng4(1) + CStr(i) + ":" + vRng4(3) + CStr(i)).Font.Bold = True
        .Range(vRng4(1) + CStr(i)).Value = "Total"
        If (lRolesTo - lRolesFrom) = 0 Then        ' Just 1 role in the table
            .Range(sColLetter(wsCC, lColNum(wsCC, vRng4(1)) + 1) + CStr(i) + ":" + vRng4(3) + CStr(i)).FormulaR1C1 = "=SUM(R[-1]C:R[-1]C)"
        ElseIf (lRolesTo - lRolesFrom) > 1 Then        'More than 1 role in the table
            .Range(sColLetter(wsCC, lColNum(wsCC, vRng4(1)) + 1) + CStr(i) + ":" + vRng4(3) + CStr(i)).FormulaR1C1 = "=SUM(R[-" + CStr(lRolesTo - lRolesFrom + 1) + "]C:R[-1]C)"
        End If
    End With
End Function

Function PopulateT1T2(ByVal wsCC As Worksheet, ByVal wsCCBKP As Worksheet)
    'Table1
    ' TODO: Validate num rows not greater than original (I till can't move the Pivot Table!!!)

    Dim sTable1Coord As String        'Table 1 coordinates in a string
    Dim sTable2Coord As String        'Table 2 coordinates in a string
    Dim sTable4Coord As String        'Table 4 coordinates in a string
    Dim sTable5Coord As String        'Table 5 coordinates in a string (counterpart of Table 1 in CCBKP hidden tab)

    Dim vTable1Coord() As String        'Table 1 coordinates in a an array
    Dim vTable2Coord() As String        'Table 2 coordinates in a an array
    Dim vTable4Coord() As String        'Table 4 coordinates in a an array
    Dim vTable5Coord() As String        'Table 5 coordinates in a an array

    Dim lLastRowInT1 As Long
    Dim lFirstRowinT4 As Long

    '    Dim lLastRowInT5 As Long                               'Last row in T5 (T1 BKP)

    Dim lNumRolesInT4 As Long

    With wsCC
        sTable1Coord = .Range("Table1").Address
        sTable1Coord = Replace(sTable1Coord, ":", "")
        vTable1Coord = Split(sTable1Coord, "$")

        sTable2Coord = .Range("Table2").Address
        sTable2Coord = Replace(sTable2Coord, ":", "")
        vTable2Coord = Split(sTable2Coord, "$")

        sTable4Coord = .Range("Table4").Address
        sTable4Coord = Replace(sTable4Coord, ":", "")
        vTable4Coord = Split(sTable4Coord, "$")

        lNumRolesInT4 = Val(vTable4Coord(4)) - Val(vTable4Coord(2)) - 1

        lLastRowInT1 = Val(vTable1Coord(2)) + lNumRolesInT4 + 1        'Calculating last row in Table 1 according to the number of roles in T4

        lFirstRowinT4 = Val(vTable4Coord(2))

        '        vTable5Coord = Split(wsCCBKP.Range("Table5").Address, "$")
        '        lLastRowInT5 = Val(vTable5Coord(4))
        '
        '        If lLastRowInT1 > lLastRowInT5 Then lLastRowInT1 = lLastRowInT5

        'Adjust gaps between upper and lower tables
        If (lLastRowInT1 + 8 > lFirstRowinT4) Then
            .Rows(CStr(Val(vTable1Coord(4)) + 1) + ":" + CStr(lLastRowInT1)).Insert
        Else
            .Rows(CStr(lLastRowInT1 + 1) + ":" + CStr(Val(vTable4Coord(2)) - 7)).EntireRow.Delete
        End If


        ' Redefine Table1 and Table2 Coordinates
        ' This snippet MUST go after the addition or deletion of rows as it might have affected the ranges!!!
        ThisWorkbook.Names("Table1").RefersTo = .Range(vTable1Coord(1) + vTable1Coord(2) + ":" + vTable1Coord(3) + CStr(lLastRowInT1))
        ThisWorkbook.Names("Table2").RefersTo = .Range(vTable2Coord(1) + vTable2Coord(2) + ":" + vTable2Coord(3) + CStr(lLastRowInT1))

        'Todo: Ideally I should be able also to move the Pivot to the right place

        ' Extract new coordinates!
        sTable1Coord = .Range("Table1").Address
        sTable1Coord = Replace(sTable1Coord, ":", "")
        vTable1Coord = Split(sTable1Coord, "$")

        sTable2Coord = .Range("Table2").Address
        sTable2Coord = Replace(sTable2Coord, ":", "")
        vTable2Coord = Split(sTable2Coord, "$")


        sTable4Coord = .Range("Table4").Address
        sTable4Coord = Replace(sTable4Coord, ":", "")
        vTable4Coord = Split(sTable4Coord, "$")


        .Range(vTable1Coord(1) + CStr(lLastRowInT1) + ":" + vTable1Coord(3) + CStr(lLastRowInT1)).Interior.Color = RGB(219, 229, 241)
        .Range(vTable1Coord(1) + CStr(lLastRowInT1) + ":" + vTable1Coord(3) + CStr(lLastRowInT1)).Font.Bold = True
        .Range(vTable1Coord(1) + CStr(lLastRowInT1)).Value = "Total"


        .Range(vTable2Coord(1) + CStr(lLastRowInT1) + ":" + vTable2Coord(3) + CStr(lLastRowInT1)).Interior.Color = RGB(219, 229, 241)
        .Range(vTable2Coord(1) + CStr(lLastRowInT1) + ":" + vTable2Coord(3) + CStr(lLastRowInT1)).Font.Bold = True
        .Range(vTable2Coord(1) + CStr(lLastRowInT1)).Value = "Total"

        If (vTable1Coord(4) - vTable1Coord(2)) > 1 Then        'There is at least 1 role!
            'Adding formulas for T1 and T2 tables
            .Range(sColLetter(wsCC, lColNum(wsCC, vTable1Coord(1)) + 1) + CStr(Val(vTable1Coord(2)) + 1) + ":" + vTable1Coord(3) + CStr(Val(vTable1Coord(4)) - 1)) = _
            "=Q" + CStr(Val(vTable4Coord(2)) + 1) + "-C" + CStr(Val(vTable4Coord(2)) + 1)

            .Range(sColLetter(wsCC, lColNum(wsCC, vTable2Coord(1)) + 1) + CStr(Val(vTable2Coord(2)) + 1) + ":" + vTable2Coord(3) + CStr(Val(vTable2Coord(4)) - 1)) = _
            "=IF(Q" + CStr(Val(vTable4Coord(2)) + 1) + "=0,C4,C4/Q" + CStr(Val(vTable4Coord(2)) + 1) + ")"

            ' Copying the roles from T4 into T1 and T2
            .Range(vTable1Coord(1) + CStr(Val(vTable1Coord(2)) + 1) + ":" + vTable1Coord(1) + CStr(Val(vTable1Coord(4) - 1))) = "=" + vTable4Coord(1) + CStr(Val(vTable4Coord(2)) + 1)
            .Range(vTable2Coord(1) + CStr(Val(vTable2Coord(2)) + 1) + ":" + vTable2Coord(1) + CStr(Val(vTable2Coord(4) - 1))) = "=" + vTable4Coord(1) + CStr(Val(vTable4Coord(2)) + 1)

            ' Adding Total Row for T1 and T2
            .Range(sColLetter(wsCC, lColNum(wsCC, vTable1Coord(1)) + 1) + CStr(lLastRowInT1) + ":" + vTable1Coord(3) + CStr(lLastRowInT1)).FormulaR1C1 = "=SUM(R[-" + CStr(lNumRolesInT4) + "]C:R[-1]C)"
            .Range(sColLetter(wsCC, lColNum(wsCC, vTable2Coord(1)) + 1) + CStr(lLastRowInT1) + ":" + vTable2Coord(3) + CStr(lLastRowInT1)).FormulaR1C1 = "=SUM(R[-" + CStr(lNumRolesInT4) + "]C:R[-1]C)"
        End If
    End With

End Function


Function ReapplyFormatting(ByVal wsCC As Worksheet, ByVal wsCCBKP As Worksheet)
    Dim R As Range

    Dim sTable1Coord As String
    Dim vTable1Coord() As String

    Dim sTable2Coord As String
    Dim vTable2Coord() As String

    Dim sTable4Coord As String
    Dim vTable4Coord() As String

    With wsCC

        ' Adjusting col width
        .Columns("A").ColumnWidth = 0.5
        .Columns("O").ColumnWidth = 0.5
        .Range("C:N").ColumnWidth = 7
        .Range("Q:AB").ColumnWidth = 7
        .Columns("P").ColumnWidth = wsCCBKP.Columns("P").ColumnWidth        'Reset to T8 size in col P

        ' Todo:Reapply formatting to tables!!!
        sTable1Coord = .Range("Table1").Address
        sTable1Coord = Replace(sTable1Coord, ":", "")
        vTable1Coord = Split(sTable1Coord, "$")

        sTable2Coord = .Range("Table2").Address
        sTable2Coord = Replace(sTable2Coord, ":", "")
        vTable2Coord = Split(sTable2Coord, "$")

        sTable4Coord = .Range("Table4").Address
        sTable4Coord = Replace(sTable4Coord, ":", "")
        vTable4Coord = Split(sTable4Coord, "$")

        Set R = .Range(sColLetter(wsCC, lColNum(wsCC, vTable1Coord(1)) + 1) + CStr(Val(vTable1Coord(2)) + 1) + ":" + vTable1Coord(3) + CStr(Val(vTable1Coord(4)) - 1))
        Call CondFormat(R)

        Set R = .Range(sColLetter(wsCC, lColNum(wsCC, vTable2Coord(1)) + 1) + CStr(Val(vTable2Coord(2)) + 1) + ":" + vTable2Coord(3) + CStr(Val(vTable2Coord(4)) - 1))
        Call CondFormat(R)

        Set R = .Range(sColLetter(wsCC, lColNum(wsCC, vTable4Coord(1)) + 1) + CStr(Val(vTable4Coord(2)) + 1) + ":" + vTable4Coord(3) + CStr(Val(vTable4Coord(4)) - 1))
        Call CondFormat(R)
    End With



End Function




























