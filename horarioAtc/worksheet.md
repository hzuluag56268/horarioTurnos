' Variable global para controlar el estado de la macro
Public MacroActiva As Boolean

' Lista de hojas válidas donde funciona el sistema
Private Function EsHojaValida() As Boolean
    Dim nombreHoja As String
    nombreHoja = ActiveSheet.Name
    
    ' Agregar aquí todas las hojas donde debe funcionar el sistema
    Select Case nombreHoja
        Case "1_QUINC", "2_QUINC" ' Agregar más nombres de hojas según sea necesario
            EsHojaValida = True
        Case Else
            EsHojaValida = False
    End Select
End Function

' Subrutina para activar la macro
Sub ActivarMacro()
    MacroActiva = True
    MsgBox "Macro de registro de turnos ACTIVADA", vbInformation, "Estado de Macro"
End Sub

' Subrutina para desactivar la macro
Sub DesactivarMacro()
    MacroActiva = False
    MsgBox "Macro de registro de turnos DESACTIVADA", vbInformation, "Estado de Macro"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Verificar si la macro está activa antes de procesar
    If MacroActiva = False Then Exit Sub
    
    ' Verificar si estamos en una hoja válida
    If Not EsHojaValida() Then Exit Sub
    
    If Not Intersect(Target, Range("C:AK")) Is Nothing Then
        Application.EnableEvents = False

        ' Manejar múltiples celdas si se copió y pegó
        Dim celda As Range
        For Each celda In Target
            ' Verificar que la celda esté en el rango válido
            If celda.Column >= 3 And celda.Column <= 37 And celda.Row <= 37 Then
                Dim trabajador As String
                trabajador = Cells(celda.Row, 2).Value

                Dim nuevoValor As String
                nuevoValor = celda.Value

                ' Verificar si la columna es par (segunda columna del día)
                Dim esSegundaColumna As Boolean
                esSegundaColumna = (celda.Column Mod 2 = 0)

                ' Procesar turnos según el valor ingresado
                Select Case nuevoValor
                    Case "L"
                        If esSegundaColumna Then
                            Call RegistrarTurnoL(trabajador)
                            celda.Interior.Color = RGB(144, 238, 144) ' Verde claro
                            MsgBox "Turno L registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "X"
                        If esSegundaColumna Then
                            Call RegistrarTurnoX(trabajador)
                            celda.Interior.Color = RGB(255, 165, 0) ' Naranja
                            MsgBox "Turno X registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "TASR"
                        If esSegundaColumna Then
                            Call RegistrarTurnoTASR(trabajador)
                            celda.Interior.Color = RGB(255, 0, 255) ' Magenta
                            MsgBox "Turno TASR registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "NANR"
                        If esSegundaColumna Then
                            Call RegistrarTurnoNANR(trabajador)
                            celda.Interior.Color = RGB(0, 255, 255) ' Cian
                            MsgBox "Turno NANR registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "NANT"
                        If esSegundaColumna Then
                            Call RegistrarTurnoNANT(trabajador)
                            celda.Interior.Color = RGB(128, 0, 128) ' Púrpura
                            MsgBox "Turno NANT registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "TANR"
                        If esSegundaColumna Then
                            Call RegistrarTurnoTANR(trabajador)
                            celda.Interior.Color = RGB(186, 85, 211) ' Magenta oscuro
                            MsgBox "Turno TANR registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "TASA"
                        If esSegundaColumna Then
                            Call RegistrarTurnoTASA(trabajador)
                            celda.Interior.Color = RGB(0, 128, 0) ' Verde oscuro
                            MsgBox "Turno TASA registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "TANA"
                        If esSegundaColumna Then
                            Call RegistrarTurnoTANA(trabajador)
                            celda.Interior.Color = RGB(128, 128, 0) ' Verde oliva
                            MsgBox "Turno TANA registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "SLN4"
                        If esSegundaColumna Then
                            Call RegistrarTurnoSLN4(trabajador)
                            celda.Interior.Color = RGB(255, 0, 0) ' Rojo
                            MsgBox "Turno SLN4 registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "SLN3"
                        If esSegundaColumna Then
                            Call RegistrarTurnoSLN3(trabajador)
                            celda.Interior.Color = RGB(0, 0, 255) ' Azul
                            MsgBox "Turno SLN3 registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "NLPR"
                        If esSegundaColumna Then
                            Call RegistrarTurnoNLPR(trabajador)
                            celda.Interior.Color = RGB(128, 128, 128) ' Gris
                            MsgBox "Turno NLPR registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "NLPT"
                        If esSegundaColumna Then
                            Call RegistrarTurnoNLPT(trabajador)
                            celda.Interior.Color = RGB(64, 64, 64) ' Gris oscuro
                            MsgBox "Turno NLPT registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "TLPR"
                        If esSegundaColumna Then
                            Call RegistrarTurnoTLPR(trabajador)
                            celda.Interior.Color = RGB(192, 192, 192) ' Gris claro
                            MsgBox "Turno TLPR registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "BANT"
                        If Not esSegundaColumna Then
                            Call RegistrarTurnoBANT(trabajador)
                            celda.Interior.Color = RGB(255, 182, 193) ' Rosa claro
                            MsgBox "Turno BANT registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case "BLPT"
                        If Not esSegundaColumna Then
                            Call RegistrarTurnoBLPT(trabajador)
                            celda.Interior.Color = RGB(255, 20, 147) ' Rosa intenso
                            MsgBox "Turno BLPT registrado para " & trabajador, vbInformation, "Confirmación"
                        End If
                    Case Else
                        ' Si se cambia o borra cualquier turno en la segunda columna (o BANT/BLPT en la primera), elimina horas extras y quita el color
                        If esSegundaColumna Or (Not esSegundaColumna And (celda.Interior.Color = RGB(255, 182, 193) Or celda.Interior.Color = RGB(255, 20, 147))) Then
                            ' Verificar el color de la celda para determinar qué turno había antes
                            Dim turnoAnterior As Boolean
                            turnoAnterior = False
                            
                            Select Case celda.Interior.Color
                                Case RGB(144, 238, 144) ' Verde claro - Turno L
                                    Call BorrarTurnoL(trabajador)
                                    turnoAnterior = True
                                Case RGB(255, 165, 0) ' Naranja - Turno X
                                    Call BorrarTurnoX(trabajador)
                                    turnoAnterior = True
                                Case RGB(255, 0, 255) ' Magenta - Turno TASR
                                    Call BorrarTurnoTASR(trabajador)
                                    turnoAnterior = True
                                Case RGB(0, 255, 255) ' Cian - Turno NANR
                                    Call BorrarTurnoNANR(trabajador)
                                    turnoAnterior = True
                                Case RGB(128, 0, 128) ' Púrpura - Turno NANT
                                    Call BorrarTurnoNANT(trabajador)
                                    turnoAnterior = True
                                Case RGB(186, 85, 211) ' Magenta oscuro - Turno TANR
                                    Call BorrarTurnoTANR(trabajador)
                                    turnoAnterior = True
                                Case RGB(0, 128, 0) ' Verde oscuro - Turno TASA
                                    Call BorrarTurnoTASA(trabajador)
                                    turnoAnterior = True
                                Case RGB(128, 128, 0) ' Verde oliva - Turno TANA
                                    Call BorrarTurnoTANA(trabajador)
                                    turnoAnterior = True
                                Case RGB(255, 0, 0) ' Rojo - Turno SLN4
                                    Call BorrarTurnoSLN4(trabajador)
                                    turnoAnterior = True
                                Case RGB(0, 0, 255) ' Azul - Turno SLN3
                                    Call BorrarTurnoSLN3(trabajador)
                                    turnoAnterior = True
                                Case RGB(128, 128, 128) ' Gris - Turno NLPR
                                    Call BorrarTurnoNLPR(trabajador)
                                    turnoAnterior = True
                                Case RGB(64, 64, 64) ' Gris oscuro - Turno NLPT
                                    Call BorrarTurnoNLPT(trabajador)
                                    turnoAnterior = True
                                Case RGB(192, 192, 192) ' Gris claro - Turno TLPR
                                    Call BorrarTurnoTLPR(trabajador)
                                    turnoAnterior = True
                                Case RGB(255, 182, 193) ' Rosa claro - Turno BANT
                                    Call BorrarTurnoBANT(trabajador)
                                    turnoAnterior = True
                                Case RGB(255, 20, 147) ' Rosa intenso - Turno BLPT
                                    Call BorrarTurnoBLPT(trabajador)
                                    turnoAnterior = True
                            End Select
                            
                            celda.Interior.ColorIndex = xlNone
                            
                            ' Solo mostrar mensaje si había un turno válido antes
                            If turnoAnterior Then
                                If nuevoValor = "" Then
                                    MsgBox "Turno borrado para " & trabajador, vbInformation, "Confirmación"
                                Else
                                    MsgBox "Turno cambiado para " & trabajador, vbInformation, "Confirmación"
                                End If
                            End If
                        Else
                            ' Si se escribe otro valor o está en la primera columna, quita el color
                            celda.Interior.ColorIndex = xlNone
                        End If
                End Select
            End If
        Next celda

        Application.EnableEvents = True
    End If
End Sub