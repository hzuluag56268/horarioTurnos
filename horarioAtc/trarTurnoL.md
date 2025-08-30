Sub RegistrarTurnoL(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna DAST en HORE_MES
    Dim columnaDAST As Long
    columnaDAST = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "DAST" Then
            columnaDAST = j
            Exit For
        End If
    Next j

    If columnaDAST = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de DAST
    For j = columnaDAST To columnaDAST + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 4
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoL(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna DAST en HORE_MES
    Dim columnaDAST As Long
    columnaDAST = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "DAST" Then
            columnaDAST = j
            Exit For
        End If
    Next j

    If columnaDAST = 0 Then Exit Sub

    ' Buscar la última celda con 4 a la derecha de DAST y borrarla
    Dim ultimaCol As Long
    For j = columnaDAST + 20 To columnaDAST Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 4 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoX(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de diurnas
    For j = columnaDIURNAS To columnaDIURNAS + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoX(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de diurnas y borrarla
    Dim ultimaCol As Long
    For j = columnaDIURNAS + 20 To columnaDIURNAS Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

' Nuevos turnos basados en la tabla
Sub RegistrarTurnoTASR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de diurnas
    For j = columnaDIURNAS To columnaDIURNAS + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoTASR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de diurnas y borrarla
    Dim ultimaCol As Long
    For j = columnaDIURNAS + 20 To columnaDIURNAS Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoNANR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna MAST/NANR en HORE_MES
    Dim columnaMASTNANR As Long
    columnaMASTNANR = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "MAST/NANR" Then
            columnaMASTNANR = j
            Exit For
        End If
    Next j

    If columnaMASTNANR = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de MAST/NANR
    For j = columnaMASTNANR To columnaMASTNANR + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoNANR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna MAST/NANR en HORE_MES
    Dim columnaMASTNANR As Long
    columnaMASTNANR = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "MAST/NANR" Then
            columnaMASTNANR = j
            Exit For
        End If
    Next j

    If columnaMASTNANR = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de MAST/NANR y borrarla
    Dim ultimaCol As Long
    For j = columnaMASTNANR + 20 To columnaMASTNANR Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoNANT(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna TANT/NANT en HORE_MES
    Dim columnaTANTNANT As Long
    columnaTANTNANT = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "TANT/NANT" Then
            columnaTANTNANT = j
            Exit For
        End If
    Next j

    If columnaTANTNANT = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de TANT/NANT
    For j = columnaTANTNANT To columnaTANTNANT + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoNANT(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna TANT/NANT en HORE_MES
    Dim columnaTANTNANT As Long
    columnaTANTNANT = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "TANT/NANT" Then
            columnaTANTNANT = j
            Exit For
        End If
    Next j

    If columnaTANTNANT = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de TANT/NANT y borrarla
    Dim ultimaCol As Long
    For j = columnaTANTNANT + 20 To columnaTANTNANT Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoTANR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de diurnas
    For j = columnaDIURNAS To columnaDIURNAS + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoTANR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de diurnas y borrarla
    Dim ultimaCol As Long
    For j = columnaDIURNAS + 20 To columnaDIURNAS Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoTASA(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de diurnas
    For j = columnaDIURNAS To columnaDIURNAS + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoTASA(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de diurnas y borrarla
    Dim ultimaCol As Long
    For j = columnaDIURNAS + 20 To columnaDIURNAS Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoTANA(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de diurnas
    For j = columnaDIURNAS To columnaDIURNAS + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoTANA(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de diurnas y borrarla
    Dim ultimaCol As Long
    For j = columnaDIURNAS + 20 To columnaDIURNAS Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoSLN4(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna DAST en HORE_MES
    Dim columnaDAST As Long
    columnaDAST = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "DAST" Then
            columnaDAST = j
            Exit For
        End If
    Next j

    If columnaDAST = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de DAST
    For j = columnaDAST To columnaDAST + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 4
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoSLN4(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna DAST en HORE_MES
    Dim columnaDAST As Long
    columnaDAST = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "DAST" Then
            columnaDAST = j
            Exit For
        End If
    Next j

    If columnaDAST = 0 Then Exit Sub

    ' Buscar la última celda con 4 a la derecha de DAST y borrarla
    Dim ultimaCol As Long
    For j = columnaDAST + 20 To columnaDAST Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 4 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoSLN3(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna DAST en HORE_MES
    Dim columnaDAST As Long
    columnaDAST = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "DAST" Then
            columnaDAST = j
            Exit For
        End If
    Next j

    If columnaDAST = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de DAST
    For j = columnaDAST To columnaDAST + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 3
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoSLN3(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna DAST en HORE_MES
    Dim columnaDAST As Long
    columnaDAST = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "DAST" Then
            columnaDAST = j
            Exit For
        End If
    Next j

    If columnaDAST = 0 Then Exit Sub

    ' Buscar la última celda con 3 a la derecha de DAST y borrarla
    Dim ultimaCol As Long
    For j = columnaDAST + 20 To columnaDAST Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 3 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

' Nuevos turnos adicionales
Sub RegistrarTurnoNLPR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna MAST/NANR en HORE_MES
    Dim columnaMASTNANR As Long
    columnaMASTNANR = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "MAST/NANR" Then
            columnaMASTNANR = j
            Exit For
        End If
    Next j

    If columnaMASTNANR = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de MAST/NANR
    For j = columnaMASTNANR To columnaMASTNANR + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoNLPR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna MAST/NANR en HORE_MES
    Dim columnaMASTNANR As Long
    columnaMASTNANR = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "MAST/NANR" Then
            columnaMASTNANR = j
            Exit For
        End If
    Next j

    If columnaMASTNANR = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de MAST/NANR y borrarla
    Dim ultimaCol As Long
    For j = columnaMASTNANR + 20 To columnaMASTNANR Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoNLPT(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna TANT/NANT en HORE_MES
    Dim columnaTANTNANT As Long
    columnaTANTNANT = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "TANT/NANT" Then
            columnaTANTNANT = j
            Exit For
        End If
    Next j

    If columnaTANTNANT = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de TANT/NANT
    For j = columnaTANTNANT To columnaTANTNANT + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoNLPT(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna TANT/NANT en HORE_MES
    Dim columnaTANTNANT As Long
    columnaTANTNANT = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "TANT/NANT" Then
            columnaTANTNANT = j
            Exit For
        End If
    Next j

    If columnaTANTNANT = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de TANT/NANT y borrarla
    Dim ultimaCol As Long
    For j = columnaTANTNANT + 20 To columnaTANTNANT Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoTLPR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de diurnas
    For j = columnaDIURNAS To columnaDIURNAS + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 6
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoTLPR(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna diurnas en HORE_MES
    Dim columnaDIURNAS As Long
    columnaDIURNAS = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "diurnas" Then
            columnaDIURNAS = j
            Exit For
        End If
    Next j

    If columnaDIURNAS = 0 Then Exit Sub

    ' Buscar la última celda con 6 a la derecha de diurnas y borrarla
    Dim ultimaCol As Long
    For j = columnaDIURNAS + 20 To columnaDIURNAS Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 6 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoBANT(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna 5am en HORE_MES
    Dim columna5AM As Long
    columna5AM = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "5am" Then
            columna5AM = j
            Exit For
        End If
    Next j

    If columna5AM = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de 5am
    For j = columna5AM To columna5AM + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 1
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoBANT(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna 5am en HORE_MES
    Dim columna5AM As Long
    columna5AM = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "5am" Then
            columna5AM = j
            Exit For
        End If
    Next j

    If columna5AM = 0 Then Exit Sub

    ' Buscar la última celda con 1 a la derecha de 5am y borrarla
    Dim ultimaCol As Long
    For j = columna5AM + 20 To columna5AM Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 1 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub

Sub RegistrarTurnoBLPT(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna 5am en HORE_MES
    Dim columna5AM As Long
    columna5AM = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "5am" Then
            columna5AM = j
            Exit For
        End If
    Next j

    If columna5AM = 0 Then Exit Sub

    ' Buscar la primera celda vacía a la derecha de 5am
    For j = columna5AM To columna5AM + 20
        If wsHoreMes.Cells(filaTrabajador, j).Value = "" Then
            wsHoreMes.Cells(filaTrabajador, j).Value = 1
            Exit For
        End If
    Next j
End Sub

Sub BorrarTurnoBLPT(trabajador As String)
    Dim wsHoreMes As Worksheet
    Set wsHoreMes = ThisWorkbook.Worksheets("HORE_MES")

    ' Buscar la fila del trabajador en HORE_MES
    Dim filaTrabajador As Long
    filaTrabajador = 0
    Dim i As Long
    For i = 2 To wsHoreMes.Cells(wsHoreMes.Rows.Count, 1).End(xlUp).Row
        If wsHoreMes.Cells(i, 1).Value = trabajador Then
            filaTrabajador = i
            Exit For
        End If
    Next i

    If filaTrabajador = 0 Then Exit Sub

    ' Buscar la columna 5am en HORE_MES
    Dim columna5AM As Long
    columna5AM = 0
    Dim j As Long
    For j = 1 To 100
        If wsHoreMes.Cells(1, j).Value = "5am" Then
            columna5AM = j
            Exit For
        End If
    Next j

    If columna5AM = 0 Then Exit Sub

    ' Buscar la última celda con 1 a la derecha de 5am y borrarla
    Dim ultimaCol As Long
    For j = columna5AM + 20 To columna5AM Step -1
        If wsHoreMes.Cells(filaTrabajador, j).Value = 1 Then
            wsHoreMes.Cells(filaTrabajador, j).Value = ""
            Exit For
        End If
    Next j
End Sub
