Attribute VB_Name = "PRINCIPAL"
Sub envio_autorizados()
    'traer_macro_mail 'preparacion de la fuente de datos para las respuestas
    'MsgBox "macro mail importada"

    'traer_shuttle
    MsgBox "shuttle importado"

    cruce_shuttle
    'rastreo_2
    MsgBox "masajeo finalizado, vamos a prooceder a enviar las respuestas"

    envio_mail 'envio de las respuestas
    MsgBox "Mails enviados, ATENCIÓN el proceso no ha terminado"
    guardado_resp_autori ("autorizaciones.accdb") 'guardado de las respuestas en bbdd
    'MsgBox "!!!Base de datos actualizada, vamos a crear un BACKUP!!!"
    MsgBox "base de datos actualizada, PROCESO TERMINADO"
    'guardado_resp_autori ("autorizaciones_bkp.accdb")
    'MsgBox "!!!Backup creado, proceso terminado!!!"
End Sub

