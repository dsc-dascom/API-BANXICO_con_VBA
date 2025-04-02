Attribute VB_Name = "API_banxico"

'declaramos el token como constante
Public Const banxico_token As String = "[Token]"

Function API_banxico(serie)

    'esta funcion ejecuta la peticion de los datos mediante la API

    'declaramos la fecha inicial de la consulta
    fecha_inicio = Format(Range("D6").Value, "YYYY-MM-DD") ' "2020-01-01"

    'declaramos la fecha final de la consulta
    fecha_fin = Format(Range("D7").Value, "YYYY-MM-DD") ' "2025-01-01"      'Date
    
    'url de consulta, revisar los parametros en la pagina de SIE-API
    Url = "https://www.banxico.org.mx/SieAPIRest/service/v1/series/" & serie & "/datos/" & fecha_inicio & "/" & fecha_fin & "?mediaType=xml"

    'declaramos el objeto para hacer la conexion
    Set solicitud = CreateObject("MSXML2.ServerXMLHTTP")
    
    'establecemos la conexion
    solicitud.Open "GET", Url, False
    
    'establecemos el encabezados de la solicitud, el cual contiene el token
    solicitud.setRequestHeader "Bmx-Token", banxico_token
     
    'enviamos la solicitud
    solicitud.Send

    'guardamos la respuesta
    Set respuesta = CreateObject("MSXML2.DOMDocument")
    respuesta.LoadXML solicitud.responseText

    'verificamos el contenido de la respuesta
    'MsgBox solicitud.responseText

    'filtramos la respuesta y la guardamos como el resultado de la funcion
    Set API_banxico = respuesta.getElementsByTagName("Obs")
    
    'borramos los datos guardados en la solicitud y repuesta
    Set solicitud = Nothing
    Set respuesta = Nothing

End Function


Sub peso_dolar()

'esta macro usa la serie "SF63528" de Banxico para obtener los datos del tipo de cambio peso-dolar
serie = "SF63528"

'llamamos a la funcion
Set observaciones = API_banxico(serie)

'escribimos los datos
i = 12  'comenzamos en la fila 12
For Each obs In observaciones
    fecha = CDate(obs.SelectSingleNode("fecha").Text) 'convertimos los datos a fechas
    dato = obs.SelectSingleNode("dato").Text
    
    Range("C" & i).Value = fecha
    Range("D" & i).Value = dato
    i = i + 1

Next obs

Range("A2").Select
MsgBox ("Consulta de datos del Tipo de Cambio exitosa")
    
End Sub

