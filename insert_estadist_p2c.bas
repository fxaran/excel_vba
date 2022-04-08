Sub Insert_Estadist()

'   ***************************************************************
'   Declarar Variables
'   ***************************************************************
    Dim Mes As String
    Dim Anio As String
    Dim SheetNameREG As String
    Dim SheetNameID As String
    Dim SheetNameGCD As String
    Dim SheetNameGOD As String
    Dim SheetNameInf As String
    Dim P2X As String

    Dim MessagDateCompare As String

    Dim LastRow As Integer
    Dim LastRowBef As Integer

    Dim LastRowValue As Date
    Dim LastRowBefValue As Date

    Dim iCol As Long
    Dim Totales(1 To 4) As Long

    Dim FecUltiCierre As Date

    Dim Graf1 As String
    Dim Graf2 As String
    Dim Graf3 As String
    Dim Graf4 As String

    Dim LI As Integer
    Dim uf As Integer
    Dim trango As Integer

    Dim ProductID1 As Range
    Dim ProductID2 As Range

    Dim NumIndErr As Integer

    Dim rest As Integer


'   ***************************************************************
'           Parametros
'   ***************************************************************
    SheetNameID = "Ingreso_Datos"
    Sheets(SheetNameID).Select

    '   Inicializar variables
    Range("K2").Select
    Mes = ActiveCell.Value

    Range("L2").Select
    Anio = ActiveCell.Value

    Range("A1").Select
    P2X = ActiveCell.Value
'   ***************************************************************
'           Fin Parametros
'   ***************************************************************

'   Pestanas
    SheetNameGCD = "Graficas Cierres Diarios"
    SheetNameGOD = "DATA_OPERACIONES_DIARIAS_2022"
    SheetNameInf = "Informe_Graficos"
    SheetNameREG = Mes & "_" & Anio

'   Nombre de las Graficas
    Graf1 = "1 Gráfico"
    Graf2 = "2 Gráfico"
    Graf3 = "3 Gráfico"
    Graf4 = "4 Gráfico"
    Graf5 = "5 Gráfico"

    If Not P2X = "P2C" Then
        P2X = "P2P"
    End If

'   ***************************************************************
'           Carga Indicadores
'   ***************************************************************
'   Actualizar Grafico Causales de Rechazos P2P y P2C
    Sheets(SheetNameID).Select
    rest = 1
    LI = 1

'   Limpiar Celdas Indicadores de errores
    If IsEmpty(Range("N" & LI).Value) = False Then
        Range("N" & LI).Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlToRight)).Select
        Selection.ClearContents
    End If

''   Copiar Celdas Total de Errores por Bancos
    uf = Range("H" & Rows.Count).End(xlUp).Row
    For i = 4 To uf
        If Cells(i, 8) <> Empty Then
            NumIndErr = NumIndErr + 1
        ElseIf Cells(i, 1) = Empty Then
            Exit For
        End If
    Next i

'   Copiar Columna Total Errores
    Range("H4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

'   Pegar seleccion previa
    Range("P" & LI).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

'   Copiar Celdas Codigos por Bancos
    Range("I4").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

'   Pegar seleccion previa
    Range("N" & LI).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

'   Corregir valor
    Selection.TextToColumns Destination:=Range("N" & LI), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

'   Cargar formula de busqueda
    ActiveCell.Offset(0, 1).Select
    For i = 1 To NumIndErr
        ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],homologacion!R26C1:R49C2,2,FALSE)"
            ActiveCell.Offset(1, 0).Select
    Next i

    trango = LI + NumIndErr - 1

    Set ProductID1 = Range("N" & LI & ":N" & trango).Find(14, LookIn:=xlValues, lookat:=xlWhole)
    Set ProductID2 = Range("N" & LI & ":N" & trango).Find(56, LookIn:=xlValues, lookat:=xlWhole)


    If ProductID2 Is Nothing Then
        Debug.Print "Variable Null"
    Else
        err1 = ProductID1.Offset(, 2).Value
        err2 = ProductID2.Offset(, 2).Value
        res = err1 + err2
        Range("P" & ProductID1.Row).Value = res
        Range("N" & ProductID2.Row & ":P" & ProductID2.Row).Select
        Selection.Delete Shift:=xlUp
        rest = 2
    End If

    LineIndErrFinal = NumIndErr + LI - rest

'   Ordenamos
    Range("DataRange").Sort Key1:=Range("P1"), Order1:=xlDescending, Header:=xlYes
'   ***************************************************************
'           Fin Carga Indicadores
'   ***************************************************************


'   ***************************************************************
'           Carga Estadistica
'   ***************************************************************
    Sheets(SheetNameREG).Select
    Range("B2").Select
    Selection.End(xlToRight).Select
    ActiveCell.Offset(-1, 1).Range("A1").Select

    ActiveCell.Range("A1:E1").Select
    Selection.NumberFormat = "m/d/yyyy"
    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Negrita"
        .Size = 16
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
'
    ActiveCell.FormulaR1C1 = "=WORKDAY(RC[-5],1)"
    ActiveSheet.Range(ActiveCell.Offset(0, 4), ActiveCell.Offset(0, 0)).Select
'
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.ColumnWidth = 23.71
'
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
'
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.ColumnWidth = 14
'
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
'
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.ColumnWidth = 14
'
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
'
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.ColumnWidth = 14
'
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
'
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.ColumnWidth = 14
'
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
'
    ActiveSheet.Range(ActiveCell.Offset(24, 4), ActiveCell.Offset(0, 0)).Select
        Selection.Copy
        Selection.End(xlToRight).Select
        ActiveCell.Offset(0, 1).Select
        ActiveSheet.Paste
'
        Application.CutCopyMode = False
'
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(0, 1).Select
'
        For i = 1 To 4
            Let ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-" & i & "]," & SheetNameID & "!R2C2:R25C6," & i + 1 & ",FALSE)"
            If i <> 4 Then
                ActiveCell.Offset(0, 1).Select
            End If
        Next i
'
    ActiveCell.Offset(0, -3).Range("A1").Select
    ActiveSheet.Range(ActiveCell.Offset(0, 3), ActiveCell.Offset(0, 0)).Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Range(ActiveCell.Offset(0, 3), ActiveCell.Offset(22, 0)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Paste
'
    ActiveCell.Offset(-1, 0).Range("A1").Select
    ActiveSheet.Range(ActiveCell.Offset(0, 3), ActiveCell.Offset(23, 0)).Select
    Selection.Copy
'
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    ActiveSheet.Paste
    Application.CutCopyMode = False
'   ***************************************************************
'           Fin Carga Estadistica
'   ***************************************************************

'   ***************************************************************
'           Obtiene la fecha del dia
'   ***************************************************************
'    Sheets(SheetNameREG).Select
'    Range("B2").Select
'    Selection.End(xlToRight).Select
    ActiveCell.Offset(-2, -1).Select
'    Selection.End(xlToLeft).Select
    FecUltiCierre = ActiveCell.Value

'   ***************************************************************
'           Obtiene Totales
'   ***************************************************************
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Select
    For i = 1 To 4
        iCol = ActiveCell.Column
        Totales(i) = WorksheetFunction.Sum(Columns(iCol))
        ActiveCell.Offset(0, 1).Select
    Next i


'   ***************************************************************
'           Ordena Transacciones Aprobadas y Rechazadas (Mayor a Menor)
'   ***************************************************************
    Sheets(SheetNameGCD).Select

'   Ordenamos
    Range("DataRange2").Sort Key1:=Range("I3"), Order1:=xlDescending, Header:=xlYes


'   ***************************************************************
'           Cargar Total Operaciones del dia
'   ***************************************************************
    Sheets(SheetNameGOD).Select
    Range("A1").Select
    Selection.End(xlDown).Select

    ActiveCell.Offset(1, 0).Select

    ActiveCell.Value = FecUltiCierre

    ActiveCell.Offset(0, 1).Select

    For i = LBound(Totales) To UBound(Totales)
        ActiveCell.Value = Totales(i)
        ActiveCell.Offset(0, 1).Select
    Next i


'   ***************************************************************
'           Obtiene intervalo de una semana (7 dias de diferencia)
'   ***************************************************************
    LastRow = ActiveCell.Row
    LastRowBef = (LastRow - 6)
    LastRowValue = Range("A" & LastRow)
    LastRowBefValue = Range("A" & LastRowBef)

    '   Actualizar mensaje
    MessagDateCompare = LastRowBefValue & " al " & LastRowValue


'   ***************************************************************
'           Actualiza Graficos
'   ***************************************************************
    Sheets(SheetNameInf).Select

'    Actualizar Grafico Montos Totales Apro y Recha
    ActiveSheet.ChartObjects(Graf2).Activate
    ActiveChart.SeriesCollection(2).Select
    ActiveChart.SeriesCollection(2).Formula = _
        "=SERIES('" & SheetNameGOD & "'!$E$1,'" & SheetNameGOD & "'!$A$" & (LastRow - 6) & ":$A$" & LastRow & ",'" & SheetNameGOD & "'!$E$" & (LastRow - 6) & ":$E$" & LastRow & ",2)"

    ActiveSheet.ChartObjects(Graf2).Activate
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).Formula = _
        "=SERIES('" & SheetNameGOD & "'!$C$1,'" & SheetNameGOD & "'!$A$" & (LastRow - 6) & ":$A$" & LastRow & ",'" & SheetNameGOD & "'!$C$" & (LastRow - 6) & ":$C$" & LastRow & ",1)"

    ActiveSheet.ChartObjects(Graf2).Activate
        ActiveChart.ChartTitle.Text = _
            "Montos Totales P2C Aprobados y Rechazados " & Chr(10) & "al Cierre del Período desde el " & MessagDateCompare


'   Actualizar Grafico Cantidades Totales de Transacciones
    ActiveSheet.ChartObjects(Graf1).Activate
    ActiveChart.SeriesCollection(2).Select
    ActiveChart.SeriesCollection(2).Formula = _
        "=SERIES('" & SheetNameGOD & "'!$D$1,'" & SheetNameGOD & "'!$A$" & (LastRow - 6) & ":$A$" & LastRow & ",'" & SheetNameGOD & "'!$D$" & (LastRow - 6) & ":$D$" & LastRow & ",2)"

    ActiveSheet.ChartObjects(Graf1).Activate
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).Formula = _
        "=SERIES('" & SheetNameGOD & "'!$B$1,'" & SheetNameGOD & "'!$A$" & (LastRow - 6) & ":$A$" & LastRow & ",'" & SheetNameGOD & "'!$B$" & (LastRow - 6) & ":$B$" & LastRow & ",1)"
'
    ActiveSheet.ChartObjects(Graf1).Activate
    ActiveChart.ChartTitle.Text = _
        "Cantidades Totales de Transacciones P2C Aprobadas y Rechazadas " & Chr(10) & "al Cierre del Período desde " & MessagDateCompare

'   Comparativo una semana anterior
    If P2X = "P2P" Then
        Range("L39").Select
        ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'DATA_OPERACIONES_DIARIAS_2022'!R2C1:R" & LastRow & "C5,2,FALSE)"
        Range("L40").Select
        ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],'DATA_OPERACIONES_DIARIAS_2022'!R7C1:R" & LastRow & "C5,2,FALSE)"
    End If

'   Actualizar Grafico Rechazo de Operaciones
    ActiveSheet.ChartObjects(Graf3).Activate
    ActiveChart.SeriesCollection(1).Select
    ActiveChart.SeriesCollection(1).Formula = _
        "=SERIES(,'" & SheetNameID & "'!$O$2:$O$" & LineIndErrFinal & ",'" & SheetNameID & "'!$P$2:$P$" & LineIndErrFinal & ",1)"

    ActiveSheet.ChartObjects(Graf3).Activate
    ActiveChart.ChartTitle.Text = _
        "Causales de Rechazos de Operaciones " & P2X & Chr(10) & "al cierre del día " & FecUltiCierre

'   Actualizar Grafico Rechazo de Operaciones
    ActiveSheet.ChartObjects(Graf5).Activate
    ActiveChart.ChartTitle.Text = _
        "Transacciones " & P2X & " al Cierre del " & FecUltiCierre

    Range("Z1").Value = "Operaciones Procesadas del día " & FecUltiCierre

End Sub
