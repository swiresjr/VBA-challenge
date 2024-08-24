I received the following **source code** from a **tutor**:

Line 123 - ws.Cells(2, 17).FormulaR1C1 = "=Max(C[-6])"

Line 125 - ws.Cells(3, 17).FormulaR1C1 = "=Min(C[-6])"

Line 128 - ws.Cells(2, 16).FormulaR1C1 = "=XLOOKUP(RC[1],C[-5],C[-7],,0)"

Line 129 - ws.Cells(3, 16).FormulaR1C1 = "=XLOOKUP(RC[1],C[-5],C[-7],,0)"

Lines 137 - 138 - Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _:=False, Transpose:=False

Lines 191 - 192 - "A2:A" & lLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _xlSortNormal 

Lines 193 - 195 - ActiveSheet.Sort.SortFields.Add2 Key:=Range( _ "B2:B" & lLastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _xlSortNormal

Lines 221 - 222 - Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _Formula1:="=0"

Line 223 - Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority

Line 211 - 230 - Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
        
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
    
        .PatternColorIndex = xlAutomatic
        .Color = 7667457
        .TintAndShade = 0
        
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
        
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
    
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        
    End With
    
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("A1").Select
