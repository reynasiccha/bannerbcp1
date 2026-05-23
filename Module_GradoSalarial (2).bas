Attribute VB_Name = "Module_GradoSalarial"
Option Explicit

'===============================================================================
'  ANÁLISIS DE MOVIMIENTOS DE GRADO SALARIAL 2026
'  Autor: Plantilla generada automáticamente
'  Descripción: Calcula transiciones B3→B2 y B2→B1 por mes y segmento
'===============================================================================

' ── Constantes de columnas (basadas en la hoja DATA) ─────────────────────────
Private Const COL_ANO       As Integer = 1   ' Columna A: Año
Private Const COL_MES       As Integer = 2   ' Columna B: Mes
Private Const COL_MATRICULA As Integer = 5   ' Columna E: Matricula
Private Const COL_GS        As Integer = 13  ' Columna M: GS (Grado Salarial)
Private Const COL_GCIA      As Integer = 23  ' Columna W: Gcia

' ── Nombres de hojas ──────────────────────────────────────────────────────────
Private Const SH_DATA    As String = "DATA"
Private Const SH_RESUMEN As String = "RESUMEN"
Private Const SH_CONFIG  As String = "CONFIG"

' ── Colores ───────────────────────────────────────────────────────────────────
Private Const CLR_DARK_BLUE    As Long = 6591744    ' RGB(31,56,100)
Private Const CLR_MID_BLUE     As Long = 11953198   ' RGB(46,117,182)
Private Const CLR_LIGHT_ORANGE As Long = 13694460   ' RGB(252,228,214)
Private Const CLR_LIGHT_GREEN  As Long = 9823970    ' RGB(226,239,218)
Private Const CLR_LIGHT_BLUE   As Long = 12506606   ' RGB(189,215,238)
Private Const CLR_GOLD         As Long = 39423      ' RGB(255,215,0)
Private Const CLR_WHITE        As Long = 16777215   ' RGB(255,255,255)
Private Const CLR_YELLOW_NOTE  As Long = 16775372   ' RGB(255,242,204)

'===============================================================================
'  MACRO PRINCIPAL: ACTUALIZAR RESUMEN
'===============================================================================
Public Sub ActualizarResumen()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Procesando datos..."

    Dim wsData    As Worksheet
    Dim wsResumen As Worksheet
    Dim wsCfg     As Worksheet

    On Error GoTo ErrHandler

    Set wsData    = ThisWorkbook.Sheets(SH_DATA)
    Set wsResumen = ThisWorkbook.Sheets(SH_RESUMEN)
    Set wsCfg     = ThisWorkbook.Sheets(SH_CONFIG)

    '── 1. Leer gerencias desde CONFIG ───────────────────────────────────────
    Dim chatsG(10)    As String
    Dim llamadasG(10) As String
    Dim nC As Integer, nL As Integer
    nC = 0: nL = 0

    Dim r As Long
    For r = 4 To 10
        Dim tipo As String
        Dim gcia As String
        tipo = UCase(Trim(wsCfg.Cells(r, 1).Value))
        gcia = UCase(Trim(wsCfg.Cells(r, 2).Value))
        If gcia = "" Then Exit For
        If tipo = "CHATS" Then
            chatsG(nC) = gcia: nC = nC + 1
        ElseIf tipo = "LLAMADAS" Then
            llamadasG(nL) = gcia: nL = nL + 1
        End If
    Next r

    '── 2. Leer orden de meses desde CONFIG ──────────────────────────────────
    Dim mesOrden(1 To 12) As String
    Dim i As Integer
    For i = 1 To 12
        mesOrden(i) = Trim(wsCfg.Cells(8 + i, 2).Value)
    Next i

    '── 3. Leer DATA → diccionario (mes|mat) → GS|Gcia ──────────────────────
    Dim dictData     As Object
    Dim mesesPresent As Object
    Set dictData     = CreateObject("Scripting.Dictionary")
    Set mesesPresent = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, COL_MATRICULA).End(xlUp).Row

    Application.StatusBar = "Leyendo " & lastRow & " registros..."

    For r = 2 To lastRow
        Dim ano As String
        ano = Trim(CStr(wsData.Cells(r, COL_ANO).Value))
        If ano <> "2026" Then GoTo NextRow

        Dim mes  As String, mat As String
        Dim gs   As String, gciaCel As String
        mes     = Trim(wsData.Cells(r, COL_MES).Value)
        mat     = Trim(wsData.Cells(r, COL_MATRICULA).Value)
        gs      = Trim(wsData.Cells(r, COL_GS).Value)
        gciaCel = Trim(wsData.Cells(r, COL_GCIA).Value)

        If mes = "" Or mat = "" Then GoTo NextRow

        Dim key As String
        key = mes & "|" & mat

        If Not dictData.Exists(key) Then
            dictData.Add key, gs & "|" & gciaCel
        End If
        If Not mesesPresent.Exists(mes) Then
            mesesPresent.Add mes, True
        End If
NextRow:
    Next r

    '── 4. Construir lista ordenada de meses con datos ───────────────────────
    Dim mesesOrd()  As String
    Dim nMeses      As Integer
    nMeses = 0

    For i = 1 To 12
        If mesesPresent.Exists(mesOrden(i)) Then
            ReDim Preserve mesesOrd(nMeses)
            mesesOrd(nMeses) = mesOrden(i)
            nMeses = nMeses + 1
        End If
    Next i

    If nMeses < 2 Then
        MsgBox "Se necesitan al menos 2 meses de datos." & Chr(10) & _
               "Meses encontrados: " & nMeses, vbInformation, "Sin datos suficientes"
        GoTo Cleanup
    End If

    '── 5. Limpiar RESUMEN (filas desde la 5 hacia abajo) ───────────────────
    Dim lastResRow As Long
    lastResRow = wsResumen.Cells(wsResumen.Rows.Count, 1).End(xlUp).Row
    If lastResRow >= 5 Then
        wsResumen.Rows("5:" & lastResRow + 5).Delete
    End If

    '── 6. Calcular y escribir transiciones ──────────────────────────────────
    Dim writeRow As Long
    writeRow = 5

    Dim totB3B2C As Long, totB3B2L As Long
    Dim totB2B1C As Long, totB2B1L As Long
    totB3B2C = 0: totB3B2L = 0: totB2B1C = 0: totB2B1L = 0

    Dim t As Integer
    For t = 0 To nMeses - 2

        Application.StatusBar = "Calculando: " & mesesOrd(t) & " → " & mesesOrd(t + 1)

        Dim mesA As String, mesB As String
        mesA = mesesOrd(t): mesB = mesesOrd(t + 1)

        Dim b3b2C As Long, b3b2L As Long
        Dim b2b1C As Long, b2b1L As Long
        b3b2C = 0: b3b2L = 0: b2b1C = 0: b2b1L = 0

        '-- Iterar colaboradores del mes A --
        Dim kv As Variant
        For Each kv In dictData.Keys
            Dim kParts() As String
            kParts = Split(kv, "|")
            If kParts(0) <> mesA Then GoTo NextKey

            Dim matV As String
            matV = kParts(1)

            ' ¿Sigue activo en mes B?
            Dim keyB As String
            keyB = mesB & "|" & matV
            If Not dictData.Exists(keyB) Then GoTo NextKey

            ' GS en A y B
            Dim vA() As String, vB() As String
            vA = Split(dictData(kv), "|")
            vB = Split(dictData(keyB), "|")

            Dim gsA As String, gsB As String, gciaV As String
            gsA   = vA(0): gsB = vB(0): gciaV = UCase(vA(1))

            ' Determinar servicio
            Dim servicio As String
            servicio = ""
            Dim c As Integer
            For c = 0 To nC - 1
                If InStr(gciaV, chatsG(c)) > 0 Or InStr(chatsG(c), gciaV) > 0 Then
                    servicio = "Chats": Exit For
                End If
            Next c
            If servicio = "" Then
                For c = 0 To nL - 1
                    If InStr(gciaV, llamadasG(c)) > 0 Or InStr(llamadasG(c), gciaV) > 0 Then
                        servicio = "Llamadas": Exit For
                    End If
                Next c
            End If
            If servicio = "" Then GoTo NextKey

            ' Contar transición
            If gsA = "B3" And gsB = "B2" Then
                If servicio = "Chats" Then b3b2C = b3b2C + 1 Else b3b2L = b3b2L + 1
            ElseIf gsA = "B2" And gsB = "B1" Then
                If servicio = "Chats" Then b2b1C = b2b1C + 1 Else b2b1L = b2b1L + 1
            End If
NextKey:
        Next kv

        ' Acumulados
        totB3B2C = totB3B2C + b3b2C: totB3B2L = totB3B2L + b3b2L
        totB2B1C = totB2B1C + b2b1C: totB2B1L = totB2B1L + b2b1L

        ' Escribir bloque
        Dim periodoLbl As String
        periodoLbl = mesA & " → " & mesB

        WriteHeaderRow wsResumen, writeRow, "  ▸  " & periodoLbl
        writeRow = writeRow + 1

        WriteDataRow wsResumen, writeRow, periodoLbl, "B3 → B2", "💬 Chats",    b3b2C, b3b2C + b3b2L, True
        writeRow = writeRow + 1
        WriteDataRow wsResumen, writeRow, periodoLbl, "B3 → B2", "📞 Llamadas", b3b2L, b3b2C + b3b2L, False
        writeRow = writeRow + 1
        WriteDataRow wsResumen, writeRow, periodoLbl, "B2 → B1", "💬 Chats",    b2b1C, b2b1C + b2b1L, True
        writeRow = writeRow + 1
        WriteDataRow wsResumen, writeRow, periodoLbl, "B2 → B1", "📞 Llamadas", b2b1L, b2b1C + b2b1L, False
        writeRow = writeRow + 1

    Next t

    '── 7. Totales acumulados ─────────────────────────────────────────────────
    WriteHeaderRow wsResumen, writeRow, "  ▸  TOTALES ACUMULADOS 2026"
    writeRow = writeRow + 1
    WriteDataRow wsResumen, writeRow, "TOTAL 2026", "B3 → B2", "💬 Chats",    totB3B2C, totB3B2C + totB3B2L, True
    writeRow = writeRow + 1
    WriteDataRow wsResumen, writeRow, "TOTAL 2026", "B3 → B2", "📞 Llamadas", totB3B2L, totB3B2C + totB3B2L, False
    writeRow = writeRow + 1
    WriteDataRow wsResumen, writeRow, "TOTAL 2026", "B2 → B1", "💬 Chats",    totB2B1C, totB2B1C + totB2B1L, True
    writeRow = writeRow + 1
    WriteDataRow wsResumen, writeRow, "TOTAL 2026", "B2 → B1", "📞 Llamadas", totB2B1L, totB2B1C + totB2B1L, False
    writeRow = writeRow + 1

    '── 8. Timestamp ─────────────────────────────────────────────────────────
    Dim tsRow As Long
    tsRow = writeRow + 1
    wsResumen.Merge_Cells_Safe tsRow, 1, tsRow, 8
    Dim tsCell As Range
    Set tsCell = wsResumen.Cells(tsRow, 1)
    tsCell.Value = "Última actualización: " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & _
                   "  |  Meses procesados: " & nMeses & _
                   "  |  Transiciones calculadas: " & (nMeses - 1)
    tsCell.Font.Italic = True
    tsCell.Font.Size = 8
    tsCell.Font.Color = RGB(150, 150, 150)
    tsCell.Font.Name = "Arial"

    wsResumen.Activate
    wsResumen.Cells(1, 1).Select

    MsgBox "✅ Resumen actualizado correctamente." & Chr(10) & Chr(10) & _
           "  Meses procesados: " & nMeses & Chr(10) & _
           "  Transiciones calculadas: " & (nMeses - 1) & Chr(10) & Chr(10) & _
           "  B3→B2 Chats: " & totB3B2C & "   Llamadas: " & totB3B2L & Chr(10) & _
           "  B2→B1 Chats: " & totB2B1C & "   Llamadas: " & totB2B1L, _
           vbInformation, "Análisis Grado Salarial 2026"

Cleanup:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & Chr(10) & _
           "En línea: " & Erl, vbCritical, "Error en macro"
End Sub

'===============================================================================
'  HELPER: ESCRIBIR FILA ENCABEZADO DE PERIODO
'===============================================================================
Private Sub WriteHeaderRow(ws As Worksheet, rowNum As Long, label As String)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 8))
    rng.Merge
    rng.Cells(1, 1).Value = label
    rng.Interior.Color = CLR_DARK_BLUE
    rng.Font.Bold = True
    rng.Font.Color = CLR_GOLD
    rng.Font.Size = 11
    rng.Font.Name = "Arial"
    rng.HorizontalAlignment = xlLeft
    rng.VerticalAlignment = xlCenter
    rng.IndentLevel = 1
    rng.Borders.LineStyle = xlContinuous
    rng.Borders.Color = CLR_MID_BLUE
    rng.Borders.Weight = xlMedium
    ws.Rows(rowNum).RowHeight = 26
End Sub

'===============================================================================
'  HELPER: ESCRIBIR FILA DE DATO
'===============================================================================
Private Sub WriteDataRow(ws As Worksheet, rowNum As Long, _
                          periodo As String, transicion As String, _
                          segmento As String, valor As Long, _
                          total As Long, isChats As Boolean)
    Dim bg As Long
    bg = IIf(isChats, CLR_LIGHT_ORANGE, CLR_LIGHT_GREEN)

    Dim pct As Double
    pct = IIf(total > 0, valor / total, 0)

    ' Barra visual proporcional
    Dim barLen As Integer
    barLen = Int(pct * 15)
    Dim bar As String
    bar = String(barLen, Chr(9608)) & String(15 - barLen, Chr(9617))

    Dim values(1 To 7) As Variant
    values(1) = periodo
    values(2) = transicion
    values(3) = segmento
    values(4) = valor
    values(5) = total
    values(6) = pct
    values(7) = bar

    Dim col As Integer
    For col = 1 To 7
        Dim cell As Range
        Set cell = ws.Cells(rowNum, col)
        cell.Value = values(col)
        cell.Interior.Color = bg
        cell.Font.Name = "Arial"
        cell.Font.Size = 10
        cell.Font.Bold = False
        cell.VerticalAlignment = xlCenter
        cell.HorizontalAlignment = IIf(col <= 3, xlLeft, xlCenter)

        Dim brd As Border
        For Each brd In cell.Borders
            brd.LineStyle = xlContinuous
            brd.Color = RGB(170, 170, 170)
            brd.Weight = xlThin
        Next brd

        If col = 4 Or col = 5 Then cell.NumberFormat = "#,##0"
        If col = 6 Then cell.NumberFormat = "0.0%"
        If col = 7 Then
            cell.Font.Size = 8
            cell.Font.Color = RGB(46, 117, 182)
            cell.HorizontalAlignment = xlLeft
        End If
    Next col

    ws.Rows(rowNum).RowHeight = 20
End Sub

'===============================================================================
'  HELPER: Merge seguro (evita error si ya está mergeado)
'===============================================================================
Private Sub Merge_Cells_Safe(ws As Worksheet, r1 As Long, c1 As Long, _
                               r2 As Long, c2 As Long)
    On Error Resume Next
    ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2)).Merge
    On Error GoTo 0
End Sub
