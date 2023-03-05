Attribute VB_Name = "Auxiliar"
Dim fila As Long
Dim fila_copiar As Long

Function UltimaFila(hoja As Worksheet, col As Integer) As Long
    UltimaFila = hoja.Cells(Rows.Count, col).End(xlUp).Row
End Function

Sub llena_combo(combo As Object, mensaje As String, col As Integer, hoja As Worksheet)
    Dim fila As Integer
    Dim i As Integer
    fila = Auxiliar.UltimaFila(hoja, col)
    combo.Style = fmStyleDropDownList
    combo.AddItem mensaje
    combo.ListIndex = 0
    For i = 2 To fila
        combo.AddItem hoja.Cells(i, col)
    Next i
End Sub

Sub mensaje_error(mensaje As String, titulo As String)
    MsgBox mensaje, vbCritical, titulo
End Sub

Sub mensaje_exito(mensaje As String, titulo As String)
    MsgBox mensaje, vbInformation, titulo
End Sub
Sub FiltrarPorPais(pais As String)
        fila = Auxiliar.UltimaFila(Hoja1, 1)
        For i = 2 To fila
            If pais = Hoja1.Cells(i, 1) Then
                Hoja1.Range("a" & i & ":l" & i).Copy
                fila_copiar = Hoja2.Cells(Rows.Count, 1).End(xlUp).Row + 1
                Hoja2.Range("a" & fila_copiar).PasteSpecial
            End If
        Next i
End Sub

Sub FiltrarPorCliente(ncliente As String)
        fila = Auxiliar.UltimaFila(Hoja1, 1)
        For i = 2 To fila
            If ncliente = Hoja1.Cells(i, 6) Then
                Hoja1.Range("a" & i & ":l" & i).Copy
                fila_copiar = Hoja2.Cells(Rows.Count, 1).End(xlUp).Row + 1
                Hoja2.Range("a" & fila_copiar).PasteSpecial
            End If
        Next i
End Sub

Sub FiltrarPorAmbos(pais As String, ncliente As String)
        fila = Auxiliar.UltimaFila(Hoja1, 1)
        For i = 2 To fila
            If ncliente = Hoja1.Cells(i, 6) And pais = Hoja1.Cells(i, 1) Then
                Hoja1.Range("a" & i & ":l" & i).Copy
                fila_copiar = Hoja2.Cells(Rows.Count, 1).End(xlUp).Row + 1
                Hoja2.Range("a" & fila_copiar).PasteSpecial
            End If
        Next i
End Sub
Sub generar_reporte()

    Dim i_pais As Integer
    Dim i_cliente As Integer
    Dim ultima As Integer
    
    Dim combopais As String
    Dim combocliente As String
    Dim nombre As String
    
    i_pais = frm_reportes.cbo_pais.ListIndex
    i_cliente = frm_reportes.cbo_ncliente.ListIndex
    
    combopais = frm_reportes.cbo_pais
    combocliente = frm_reportes.cbo_ncliente
     
    If i_pais = 0 And i_cliente = 0 Then
        Call Auxiliar.mensaje_error("Por favor seleccione alguna de las opciones", "Error de selección")
        Exit Sub
    End If
    
    'Limpia la hoja
    ultima = Auxiliar.UltimaFila(Hoja2, 1)
    
    For i = 2 To ultima
        Hoja2.Range("a" & i & ":l" & i).Clear
    Next i
    
    'Filtrar por pais
    If i_pais <> 0 And i_cliente = 0 Then
        Call FiltrarPorPais(combopais)
    End If
     
    'Filtrar por cliente
    If i_pais = 0 And i_cliente <> 0 Then
        Call FiltrarPorCliente(combocliente)
    End If
    
    'Filtrar por ambos
    If i_pais <> 0 And i_cliente <> 0 Then
        Call FiltrarPorAmbos(combopais, combocliente)
    End If


    Hoja2.Activate
    nombre = InputBox("Elige el nombre")
    Hoja2.Copy
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & nombre & ".xlsx"
    ActiveWorkbook.Close
    Hoja1.Activate
    
End Sub

Sub AbrirFormulario()
    frm_reportes.Show
End Sub




