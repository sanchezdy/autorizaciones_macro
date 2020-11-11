Attribute VB_Name = "ACCES"
'esta función guarda las respuestas (autorizados y no autorizados en bbdd)
Function guardado_resp_autori(tabla As String)

Dim conn As ADODB.Connection 'creamos objeto bd
Dim res As ADODB.Recordset 'creamos objeto bd
Dim consulta As String
Dim consulta_2 As String
Dim OA As Object
Dim MAIL As Object


Set OA = CreateObject("Outlook.Application")

fecha_registro = Date

Set conn = New ADODB.Connection 'bd será un objeto que acojera los datos de conexión con la base de datos
Set res = New ADODB.Recordset 'Objeto que contendrá el resultado de la consulta

'''''''''''''''''''''''''''''''''''''''
'Conexión con la base de datos
'conn.Open ("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\alika\OneDrive - Telefonica\master_macro\autorizaciones.accdb")
'mi_ruta = "C:\Users\alika\OneDrive - Telefonica\master_macro\autorizaciones.accdb"
mi_ruta = ruta_personal & tabla
With conn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open mi_ruta
End With

'''''''''''''''''''''''''''''''''''''''


filas_excel = Worksheets("fuente").Cells(1, 1).CurrentRegion.Rows.Count 'obtenemos el total de filas usadas en la hoja fuente


For i = 2 To filas_excel
    
    'comprobar si existe un valor en la base de datos de acces
    'consulta = "SELECT NU_TELEFONO FROM resp_autori WHERE NU_TELEFONO = " & Worksheets("fuente").Cells(i, 3).Value & ""
    'Set res = New ADODB.Recordset
    'res.CursorLocation = adUseServer
    'res.Open Source:=consulta, _
    ActiveConnection:=conn
    
    'If res.EOF And res.BOF Then 'solo se rellenan los campos si existe el valor, si no existe no entra
    
        campo_1 = Worksheets("fuente").Cells(i, 1).Value 'enviado_por
        campo_2 = Worksheets("fuente").Cells(i, 2).Value 'fecha_recepcion
        campo_3 = Worksheets("fuente").Cells(i, 3).Value 'nu_telefono
        campo_4 = Worksheets("fuente").Cells(i, 4).Value 'asunto
        campo_5 = Worksheets("fuente").Cells(i, 7).Value 'no_auto
        campo_6 = Worksheets("fuente").Cells(i, 6).Value 'auto
        campo_7 = Worksheets("fuente").Cells(i, 5).Value 'cc
        campo_8 = fecha_registro 'fecha de registro en bbdd
        campo_9 = Worksheets("fuente").Cells(i, 8).Value 'CIF
        campo_10 = Worksheets("fuente").Cells(i, 9).Value 'CLIENTE
        
        consulta_2 = "INSERT INTO resp_autori(enviado_por,fecha_recepcion,NU_TELEFONO,asunto,no_auto,auto,cc,fecha_registro,CIF,cliente) VALUES ('" & campo_1 & "','" & campo_2 & "'," & campo_3 & ",'" & campo_4 & "','" & campo_5 & "','" & campo_6 & "','" & campo_7 & "','" & campo_8 & "','" & campo_9 & "','" & campo_10 & "')"
        conn.Execute consulta_2
        On Error GoTo Control_error_1

     'End If
    
Next

'desconectamos y liberamos memoria
conn.Close
Set res = Nothing
Set conn = Nothing


Exit Function
Control_error_1:
            'Set Mail = OA.CreateItem(olMailItem)
            MsgBox "Se ha producido ERROR al guardar en BBDD el siguiente registro " & campo_3
                'Mail.SentOnBehalfOfName = "autorizacionespymes@telefonica.com"
                'Mail.To = "dey.sanchezgarcia@telefonica.com"
                'Mail.Subject = "HA EXISTIDO UN ERROR DE CARGA EN BBDD"
                'Mail.Body = "Se ha detectado el siguiente error " & Err.Description & " en el registro con NU_TELEFONO " & campo_3
                'Mail.Send
                Resume Next

End Function
