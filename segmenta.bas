Function segmenta(ByRef rango As Range) As Variant()
    
    Dim toma As String, linea As String
    Dim largo As Integer, s As Byte
    Dim caracter As String * 1
    Dim fila As Long, columna As Integer
    Dim A() As Integer
    
'--------------------------------------------'
    ' Vectores
    Dim G() As Variant  ' recoge 1 y 0 para los espacios
    Dim B() As Variant  ' actual segmentado
    Dim L() As Variant  ' destino con espacios
    
    ' variables de control
    Dim col_crit As Integer 'donde inicia las preguntas
    Dim num_lin As Integer  'comentarios (separados por lineas)
    Dim num_crit As Integer '
    'Dim num_filas As Integer
    
    'indices y temporales
    Dim i As Integer    'Indice principal
    Dim j As Integer    'Indice secundario
    Dim aux As Integer  'contador
    Dim marca As String 'tache/paloma
    
'--------------------------------------------'
    ' Recolección de parametros
    
    num_crit = rango.Columns.Count - 1  ' <---- num_crit
       
    aux = 0
    For Each celda In rango
        If aux = 0 Then
            col_crit = celda.Column     ' <---- col_crit
            fila = celda.row            ' <---- fila
            aux = aux + 1
            Exit For
        End If
    Next celda

'--------------------------------------------'
    
    'CAJA GRIS
    'Separa comentarios por columnas
    
    'Define matriz(1, num_crit + 1 )
    L = rango.Value
  
    ' Blanquea L
    For i = 1 To num_crit + 1
        L(1, i) = ""
    Next i

    toma = Cells(fila, col_crit + num_crit).Value
    largo = Len(toma)
       
    'cuenta el número de "alt+enter" de los comentarios
    s = 0
    For i = 1 To largo
        caracter = Mid(toma, i, 1)
        If caracter = Chr(10) Then
            s = s + 1
        End If
    Next i
    
    'Sin comentarios o en blanco
    If s = 0 Then
        L(1, 1) = toma
        segmenta = L
        
    'Separa las columnas
    Else
        ReDim A(largo)
        
        s = 0
        For i = 1 To largo
            caracter = Mid(toma, i, 1)
            If caracter = Chr(10) Then
                s = s + 1
                A(s) = i
            End If
        Next i
        
        'If s = 0 Then Cells(fila, columna + i + 1) = linea: End
        'If s = 0 Then segmenta = toma: End
        A(s + 1) = largo + 1
        
       
        'Rellena L
        For i = 0 To s
            linea = Mid(toma, A(i) + 1, A(i + 1) - A(i) - 1)
            L(1, i + 1) = linea
        Next i
        
        segmenta = L
    
    End If
End Function

