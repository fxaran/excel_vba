Sub Afiliaciones()

    Range("B2").Select
    Selection.End(xlToRight).Select
    ActiveCell.Offset(0, -2).Select
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.EntireColumn.Hidden = True
    ActiveCell.Offset(1, 1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.Insert Shift:=xlToRight
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.ColumnWidth = 15
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=""Afiliaciones al ""&TEXT(TODAY(),""dd/mm/yyyy"")"
    ActiveCell.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    ActiveCell.Offset(0, -1).Select
    Selection.End(xlDown).Select
    ActiveCell.Select
    Selection.Copy
    ActiveCell.Offset(0, 1).Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=(RC[-1]-RC[-2])/RC[-2]"
    ActiveCell.Select
    Selection.Copy
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.ColumnWidth = 12
    Range("B2").Select

End Sub
