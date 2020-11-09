Attribute VB_Name = "BBDDEXCEL"
Function traer_macro_mail()
    
    Dim conn As ADODB.Connection 'creamos objeto bd
    Dim res As ADODB.Recordset 'creamos objeto bd
    Dim consulta As String

    Set conn = New ADODB.Connection 'bd será un objeto que acojera los datos de conexión con la base de datos
    Set res = New ADODB.Recordset 'Objeto que contendrá el resultado de la consulta

    'conectamos con el fichero "shuttle como base de datos
    'conexion con una base de datos excel
    mi_ruta = "DATA SOURCE=" & ruta_personal & "macro_mail.xlsm"
    With conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties") = "Excel 12.0; HDR=YES"
        .Open mi_ruta
    End With
    'fin de conexion con una base de datos de excel
  
    
    consulta = "SELECT [Enviado por],[Fecha recepción],[NU_TELEFONO],[Asunto],[CC] FROM [Hoja1$]"
    
    Set res = New ADODB.Recordset
    res.CursorLocation = adUseServer
    res.Open Source:=consulta, _
    ActiveConnection:=conn
    
    numero_de_filas = Worksheets("fuente").UsedRange.Rows.Count
    
    'importamos los datos
    If numero_de_filas > 0 Then
        Worksheets("fuente").Range("A1:U" & numero_de_filas).ClearContents
    End If
    Worksheets("fuente").Cells(2, 1).CopyFromRecordset res 'se copia toda la consulta res en una hoja de excel
    
    'rellenamos los campos
    num_campos = res.Fields.Count
    For i = 1 To num_campos
        Worksheets("fuente").Cells(1, i) = res.Fields(i - 1).Name
    
    Next
    
End Function



Function traer_shuttle()

Dim conn As ADODB.Connection 'creamos objeto bd
Dim res As ADODB.Recordset 'creamos objeto bd
Dim consulta As String

Set conn = New ADODB.Connection 'bd será un objeto que acojera los datos de conexión con la base de datos
Set res = New ADODB.Recordset 'Objeto que contendrá el resultado de la consulta

'conectamos con el fichero "shuttle como base de datos
mi_ruta = "DATA SOURCE=" & ruta_personal & "shuttle.xlsx"
With conn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Properties("Extended Properties") = "Excel 12.0; HDR=YES"
    .Open mi_ruta
End With

Dim sh_hojas As Variant
Dim master_hojas As Variant
sh_hojas = Array("OK", "Rechazados")
master_hojas = Array("oks", "noks")
For j = 0 To 1
    
    'realizamos la consulta a una bbdd de Excel
    consulta = "SELECT [" & sh_hojas(j) & "$].[NU_TELEFONO],[" & sh_hojas(j) & "$].[NU_DOCU],[" & sh_hojas(j) & "$].[CLIENTE] FROM [" & sh_hojas(j) & "$]"

    Set res = New ADODB.Recordset
    res.CursorLocation = adUseServer
    res.Open Source:=consulta, _
    ActiveConnection:=conn
    
    numero_de_filas = Worksheets(master_hojas(j)).UsedRange.Rows.Count
    
    If numero_de_filas > 0 Then
        Worksheets(master_hojas(j)).Range("A1:U" & numero_de_filas).ClearContents
    End If
    Worksheets(master_hojas(j)).Cells(2, 1).CopyFromRecordset res 'se copia toda la consulta res en una hoja de excel
    Worksheets(master_hojas(j)).Cells(1, 1) = "NU_TELEFONO"
    Worksheets(master_hojas(j)).Cells(1, 2) = "NU_DOCU"
    Worksheets(master_hojas(j)).Cells(1, 3) = "CLIENTE"
    Worksheets(master_hojas(j)).Cells(1, 4) = "resp"

    new_numero_de_filas = Worksheets(master_hojas(j)).UsedRange.Rows.Count
    'ahora ya podemos recorrer la consulta

    For i = 2 To new_numero_de_filas
        trimado = Trim(Worksheets(master_hojas(j)).Cells(i, 1))
        Worksheets(master_hojas(j)).Cells(i, 1) = trimado
        Worksheets(master_hojas(j)).Cells(i, 4) = 1
        cif_manip = Right(Worksheets(master_hojas(j)).Cells(i, 2), 9)
        Worksheets(master_hojas(j)).Cells(i, 2) = cif_manip
    Next
    Set res = Nothing
Next


End Function

Function cruce_shuttle()

    Dim conn As ADODB.Connection 'creamos objeto bd
    Dim res As ADODB.Recordset 'creamos objeto bd
    Dim consulta As String

    Set conn = New ADODB.Connection 'bd será un objeto que acojera los datos de conexión con la base de datos
    Set res = New ADODB.Recordset 'Objeto que contendrá el resultado de la consulta

    'conectamos con el fichero "shuttle como base de datos
    mi_ruta = "DATA SOURCE=" & ruta_personal & "envio_autorizados.xlsm"
    With conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties") = "Excel 12.0; HDR=YES"
        .Open mi_ruta
    End With
    
    
    Dim master_hojas As Variant
    Dim master_col As Variant
    Dim master_campos As Variant
    master_hojas = Array("oks", "noks")
    master_campos = Array("NU_DOCU", "CLIENTE")
    
    'importamos autorizados y no autorizados
    c = 6 'desde donde empezamos a importar (columna 6)
    For i = 0 To 1
    
    consulta = "SELECT [" & master_hojas(i) & "$].[resp] FROM [fuente$] LEFT JOIN [" & master_hojas(i) & "$] ON [fuente$].[NU_TELEFONO]=[" & master_hojas(i) & "$].[NU_TELEFONO] "
   
    Set res = New ADODB.Recordset
    res.CursorLocation = adUseServer
    res.Open Source:=consulta, _
    ActiveConnection:=conn
    
    numero_de_filas = Worksheets("fuente").Cells(1, 1).CurrentRegion.Rows.Count
    If numero_de_filas > 0 And i = 0 Then
        Worksheets("fuente").Range("F2:J" & numero_de_filas).ClearContents
    End If
    
    Worksheets("fuente").Cells(1, 6) = "auto"
    Worksheets("fuente").Cells(1, 7) = "no_auto"
    Worksheets("fuente").Cells(2, c).CopyFromRecordset res
    c = c + 1
    
    Next
    
    'fin de importacion
    
    'importación del CIF y CLIENTE
    Set res = Nothing
   
    c = 8 'variables que indican en que colimnas empezamos a importar
    d = 9

    For j = 0 To 1
    For i = 0 To 1
    
    If i = 0 Then
        consulta = "SELECT [" & master_hojas(i) & "$].[" & master_campos(j) & "] FROM [fuente$] LEFT JOIN [" & master_hojas(i) & "$] ON [fuente$].[NU_TELEFONO]=[" & master_hojas(i) & "$].[NU_TELEFONO]"
    ElseIf i = 1 Then
        consulta = "SELECT [" & master_hojas(i) & "$].[" & master_campos(j) & "] FROM [fuente$] LEFT JOIN [" & master_hojas(i) & "$] ON [fuente$].[NU_TELEFONO]=[" & master_hojas(i) & "$].[NU_TELEFONO]"
    End If
    
    
    Set res = New ADODB.Recordset
    res.CursorLocation = adUseServer
    res.Open Source:=consulta, _
    ActiveConnection:=conn
    
    numero_de_filas = Worksheets(master_hojas(i)).UsedRange.Rows.Count
    If numero_de_filas > 0 Then
        Worksheets(master_hojas(i)).Range("H2:H" & numero_de_filas).ClearContents
    End If
    
    If j = 0 Then
        Worksheets("fuente").Cells(1, 8) = "CIF"
        Worksheets("fuente").Cells(2, c).CopyFromRecordset res
        c = c + 2
    ElseIf j = 1 Then
        Worksheets("fuente").Cells(1, 9) = "CLIENTE"
        Worksheets("fuente").Cells(2, d).CopyFromRecordset res
        d = d + 2
    
    End If
        
    Next
    
    'fin de importacion

Next

Set res = Nothing

End Function
    
'segundo rastreo para unir las filas de CIFs y CIFS no autorizados y
'CLIENTE y CLIENTE no autorizado

Function rastreo_2()

    numero_de_filas = Worksheets("fuente").UsedRange.Rows.Count
    'variables para recorrer columnas
    c = 8
    e = 10
    For j = 0 To 1
        For i = 2 To numero_de_filas

        If IsEmpty(Worksheets("fuente").Cells(i, c)) = True Then
            Worksheets("fuente").Cells(i, c) = Worksheets("fuente").Cells(i, e)
        End If
        Next

    c = c + 1
    e = e + 1
    Next
    'Range("K:L").Columns.Delete

End Function
