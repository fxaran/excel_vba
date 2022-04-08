Sub Cierre_Afiliaciones()

    uf = Range("B" & Rows.Count).End(xlUp).Row

    For i = 2 To uf
        If Cells(i, 2) <> Empty Then
            NumAfi = NumAfi + 1
        ElseIf Cells(i, 1) = Empty Then
            Exit For
        End If
    Next i
    NumAfi = NumAfi - 2
    
    Range("B2").Select
    Selection.End(xlToRight).Select
    ActiveCell.Offset(1, -1).Select
    
    For i = 1 To NumAfi
        If ActiveCell.Value = "" Then
            ActiveCell.Offset(0, -1).Select
            Selection.Copy
            ActiveCell.Offset(0, 1).Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Application.CutCopyMode = False
        End If
        ActiveCell.Offset(1, 0).Select
    Next i

End Sub
