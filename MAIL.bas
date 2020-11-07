Attribute VB_Name = "MAIL"
Function envio_mail()
    
    'declaracion de variables
    Dim OA As Object
    Dim MI As Object 'creamos un objeto mailitem
    Dim MITWO As Object
    Dim Destinatario As String
    Dim Asunto As String
    Dim Cuerpo As String
    Dim u_fila As Integer
    Dim Contador As Integer
    Dim fecha_actual As Date
    Dim contador_resp As Integer
    Dim contador_noresp As Integer
    Dim contador_autori As Integer
    Dim contador_noautori As Integer
    
    Set OA = CreateObject("Outlook.Application") 'Damos valor al objeto como objeto Outlook
    u_fila = Worksheets("fuente").UsedRange.SpecialCells(xlCellTypeLastCell).Row 'obtenemos el total de filas usadas
    tot_fila = u_fila - 1
    
    fecha_actual = Date
    contador_resp = 0
    contador_autori = 0
    contador_noautori = 0
    contador_noresp = 0
       
    
    For Contador = 1 To tot_fila
        
        Set MI = OA.CreateItem(olMailItem)
        
        If Worksheets("fuente").Cells(Contador + 1, 6).Value = "1" Or Worksheets("fuente").Cells(Contador + 1, 7).Value = "1" Then
            Destinatario = Worksheets("fuente").Cells(Contador + 1, 1).Value
            Asunto = Worksheets("fuente").Cells(Contador + 1, 3).Value & "--" & Worksheets("fuente").Cells(Contador + 1, 8).Value & "--" & Worksheets("fuente").Cells(Contador + 1, 9).Value
            
            
            If Worksheets("fuente").Cells(Contador + 1, 6).Value = "1" Then
                Cuerpo = Worksheets("txt").Cells(1, 2).Value
                contador_autori = contador_autori + 1
                
            End If
            
            If Worksheets("fuente").Cells(Contador + 1, 7).Value = "1" Then
                Cuerpo = Worksheets("txt").Cells(2, 2).Value
                contador_noautori = contador_noautori + 1
               
            End If

            With MI 'with nos permite ejecutar un conjunto de instrucciones sin tener que volver hacer referencia al objeto
                .SentOnBehalfOfName = "autorizacionespymes@telefonica.com" 'se manda el mail en nombre de la cuenta especificada pero no desde la cuenta especificada
                .To = Destinatario
                .CC = Worksheets("fuente").Cells(Contador + 1, 5).Value
                .BCC = "autorizacionespymes@telefonica.com"
                .Subject = Asunto
                .HTMLBody = Cuerpo
                .Send
            End With
            contador_resp = contador_resp + 1
        Else
            Cuerpo = Worksheets("fuente").Cells(Contador + 1, 3).Value
            With MI 'with nos permite ejecutar un conjunto de instrucciones sin tener que volver hacer referencia al objeto
                .SentOnBehalfOfName = "autorizacionespymes@telefonica.com" 'se manda el mail en nombre de la cuenta especificada pero no desde la cuenta especificada
                .To = "autorizacionespymes@telefonica.com"
                .CC = Worksheets("fuente").Cells(Contador + 1, 5).Value
                '.BCC = "autorizacionespymes@telefonica.com"
                .Subject = "PETICIÓN AUTORIZACIÓN NO RESPONDIDA " & Worksheets("fuente").Cells(Contador + 1, 4).Value
                .Body = "La petición de autorización de " & Cuerpo & " no ha sido respondida"
                '.Send
            End With
            contador_noresp = contador_noresp + 1
        End If
     
        
        
    Next
    Set MITWO = OA.CreateItem(olMailItem)
        
    With MITWO 'with nos permite ejecutar un conjunto de instrucciones sin tener que volver hacer referencia al objeto
        .SentOnBehalfOfName = "autorizacionespymes@telefonica.com" 'se manda el mail en nombre de la cuenta especificada pero no desde la cuenta especificada
        .To = "autorizacionespymes@telefonica.com"
        '.BCC = "autorizacionespymes@telefonica.com"
        .Subject = "INFORME DE ENVIOS A: " & fecha_actual
        .Body = "Hoy se han respondido " & contador_resp & " peticiones de autorización de las cuales " & contador_autori & " han sido autorizadas. Hay " & contador_noresp & " peticiones de autorización no respondidas."
        .Send
    End With
    
    Sheets(1).Select
    
   
End Function

