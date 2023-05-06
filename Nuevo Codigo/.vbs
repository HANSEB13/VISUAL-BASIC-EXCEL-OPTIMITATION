Sub ProcesoCompleto()
    EliminarColumns "C:C, B:B, D:D, L:L" 'Eliminar varias columnas en una línea
    Traducir "MATE, LAMMAT", "BRILLANTE, LAMBTE", 11 'Traducir columnas 11
    Traducir "Color, CL", "", 7 'Traducir columnas 7
    Traducir "HOLMEN55, BNOBR080, BOND70, BNOBR090, ESM115, BNILM115", "BNAHU080, COOBR090", 8 'Traducir columna 8
    Traducir "ESM. 300, TAILU270", "", 9 'Traducir columna 9
    Traducir "RUSTICO, ENCBIN, CABALLETE, ENCACA", "", 10 'Traducir columna 10
    AjustarAnchoColumns "A2:A2, B2:B2, C2:C2, D2:D2, E2:E2, F2:F2, G2:G2, H2:H2, I2:I2, J2:J2" 'Ajustar el ancho de varias columnas en una línea
    RellenarCeldas "A1:A1000", 30 'Rellenar varias celdas en una línea
    BordesLinea "A1:N1000", 1, 2 'Agregar bordes a todas las celdas
End Sub

Sub EliminarColumns(col As String)
    Range(col).Columns.Delete
End Sub

Sub Traducir(findText As String, replaceText As String, col As Integer)
    Dim i As Integer
    Dim j As Integer

    j = Cells(Rows.Count, 1).End(xlUp).Row 'Encontrar la última fila de la columna 1

    For i = 1 To j
        Cells(i, col).Value = Replace(Cells(i, col).Value, findText, replaceText)
        If replaceText <> "" Then Cells(i, col).Interior.Color = RGB(133, 193, 233)
    Next i
End Sub

Sub AjustarAnchoColumns(col As String)
    Dim arr() As String
    arr = Split(col, ", ")
    For i = 0 To UBound(arr)
        Range(arr(i)).ColumnWidth = Choose(i + 1, 20, 15, 20, 10, 10, 10, 10, 10, 15, 15)
    Next i
    Range("A1:N1").Interior.Color = RGB(100, 140, 280)
End Sub

Sub RellenarCeldas(cellRange As String, height As Double)
    Range(cellRange).RowHeight = height
End Sub

Sub BordesLinea(cellRange As String, colorIndex As Integer, weight As Integer)
    Range(cellRange).BorderAround LineStyle:=xlContinuous, ColorIndex:=colorIndex, Weight:=weight
End Sub
