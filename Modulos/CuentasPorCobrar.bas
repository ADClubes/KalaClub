Attribute VB_Name = "CuentasPorCobrar"
Option Explicit

Private Const sSOAPInsert = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
      "<soap:Body>" & _
        "<Insert xmlns=""http://www.sportium.com.mx/"">" & _
          "<pIdSucursal>int</pIdSucursal>" & _
          "<pOrigen>string</pOrigen>" & _
          "<pIdCliente>int</pIdCliente>" & _
          "<pDescripcion>string</pDescripcion>" & _
          "<pIdConcepto>int</pIdConcepto>" & _
          "<pMonto>decimal</pMonto>" & _
          "<pImpuesto>decimal</pImpuesto>" & _
          "<pTotal>decimal</pTotal>" & _
          "<pReferencia>string</pReferencia>" & _
          "<pFechaGeneracion>dateTime</pFechaGeneracion>" & _
          "<pFechaVencimiento>dateTime</pFechaVencimiento>" & _
          "<pEstado>string</pEstado>" & _
          "<pActivo>boolean</pActivo>" & _
          "<pSincronizado>string</pSincronizado>" & _
          "<pIdGrupoUsuario>int</pIdGrupoUsuario>" & _
          "<pIdAsignacionArea>int</pIdAsignacionArea>" & _
        "</Insert>" & _
      "</soap:Body>" & _
    "</soap:Envelope>"

Public Function Insert(ByVal pIdSucursal As Integer, ByVal pOrigen As String, ByVal pIdCliente As Integer, _
            ByVal pDescripcion As String, ByVal pIdConcepto As Integer, ByVal pMonto As Double, ByVal pImpuesto As Double, _
            ByVal pTotal As Double, ByVal pReferencia As String, ByVal pFechaGeneracion As Date, ByVal pFechaVencimiento As Date, _
            ByVal pEstado As String, ByVal pActivo As Boolean, ByVal pSincronizado As String, ByVal pIdGrupoUsuario As Integer, _
            ByVal pIdAsignacionArea As Integer) As String
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    Set parser = New DOMDocument
    
    parser.loadXML sSOAPInsert
    
    'Asigna los valores para el servicio
    'IdSucursal
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pIdSucursal").Text = pIdSucursal
    'Origen
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pOrigen").Text = pOrigen
    'IdCliente
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pIdCliente").Text = pIdCliente
    'Descripcion
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pDescripcion").Text = pDescripcion
    'IdConcepto
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pIdConcepto").Text = pIdConcepto
    'Monto
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pMonto").Text = pMonto
    'Impuesto
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pImpuesto").Text = pImpuesto
    'Total
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pTotal").Text = pTotal
    'Referencia
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pReferencia").Text = pReferencia
    'FechaGeneracion
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pFechaGeneracion").Text = Format(pFechaGeneracion, "yyyy-mm-ddTh:mm:ss")
    'FechaVencimiento
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pFechaVencimiento").Text = Format(pFechaVencimiento, "yyyy-mm-ddTh:mm:ss")
    'Estado
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pEstado").Text = pEstado
    'Activo
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pActivo").Text = Abs(CInt(pActivo))
    'Sincronizado
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pSincronizado").Text = pSincronizado
    'IdGrupoUsuario
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pIdGrupoUsuario").Text = pIdGrupoUsuario
    'IdAsignacionArea
    parser.selectSingleNode("/soap:Envelope/soap:Body/Insert/pIdAsignacionArea").Text = pIdAsignacionArea
    
    DoEvents
    
    enviarComando parser.xml, "http://www.sportium.com.mx/Insert", sRespuesta, "INSERTAR"
    
    Insert = sRespuesta
    
End Function

Private Sub enviarComando(ByVal sXml As String, ByVal sSoapAction As String, ByRef sResp As String, sModo As String)
    ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    Dim oHttReq As XMLHTTP60
    Dim sValor As String
    Dim sUrlWS As String
    
    Set oHttReq = New XMLHTTP60
    
    sUrlWS = ObtieneParametro("URL_WS_CXC")
    
    '
    ' Enviar el comando de forma síncrona (se espera a que se reciba la respuesta)
    oHttReq.Open "POST", sUrlWS, False
    ' las cabeceras a enviar al servicio Web
    ' (no incluir los dos puntos en el nombre de la cabecera)
    oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    oHttReq.setRequestHeader "SOAPAction", sSoapAction
    ' enviar el comando
    oHttReq.send sXml
    DoEvents
    
    Do While True
        DoEvents
        If oHttReq.readyState = 4 Then Exit Do
    Loop

    DoEvents
    '
    ' este será el texto recibido del servicio Web
    sValor = procesarRespuesta(oHttReq.responseText, sModo)
    '
    
    sResp = sValor
    
End Sub

Private Function procesarRespuesta(ByVal s As String, sModo As String) As String
    ' procesar la respuesta recibida del servicio Web
    '
    ' Poner los datos en el analizador de XML
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    parser.loadXML s
    '
    On Error Resume Next
    '
    Select Case sModo
        Case "INSERTAR"
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/InsertResponse/InsertResult").Text
    End Select
    '
    If Err.Number > 0 Then
        MsgBox "Error"
    End If
End Function
