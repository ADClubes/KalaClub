Attribute VB_Name = "CFD"
Option Explicit

Private Const sSOAPGeneraFactura = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
    "<soap:Body>" & _
    "<GeneraFactura xmlns=""http://tempuri.org/"">" & _
        "<IdFacturaKala>int</IdFacturaKala>" & _
        "<serie>string</serie>" & _
        "<EfectoDeComprobante>string</EfectoDeComprobante>" & _
        "</GeneraFactura>" & _
    "</soap:Body>" & _
    "</soap:Envelope>"
    
Private Const sSOAPCancelaFactura = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
    "<soap:Body>" & _
    "<CancelaFactura xmlns=""http://tempuri.org/"">" & _
        "<IdFolio>string</IdFolio>" & _
        "<serie>string</serie>" & _
    "</CancelaFactura>" & _
    "</soap:Body>" & _
    "</soap:Envelope>"
    
Private Const sSOAPGeneraNotaCredito = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
    "<soap:Body>" & _
    "<GeneraNotaCredito xmlns=""http://tempuri.org/"">" & _
        "<NumeroNota>int</NumeroNota>" & _
        "<serie>string</serie>" & _
        "<EfectoDeComprobante>string</EfectoDeComprobante>" & _
        "</GeneraNotaCredito>" & _
    "</soap:Body>" & _
    "</soap:Envelope>"
    
    
Private Const sSOAPPruebaServicio = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
    "<soap:Body>" & _
    "<PruebaServicio xmlns=""http://tempuri.org/"" />" & _
    "</soap:Body>" & _
    "</soap:Envelope>"

    
'---------------------------------------------------------------------------
Public Function CancelaCFD(sNoFolioFac As String, sSerieCFD As String) As String
    
    
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    
    
    Set parser = New DOMDocument
    
    
    
    
    parser.loadXML sSOAPCancelaFactura
    
    'Asigna los valores para el servicio
    'Folio
    parser.selectSingleNode("/soap:Envelope/soap:Body/CancelaFactura/IdFolio").Text = sNoFolioFac
    
    'Serie
    parser.selectSingleNode("/soap:Envelope/soap:Body/CancelaFactura/serie").Text = sSerieCFD
    
       
    DoEvents
    
       
    
    enviarComando parser.xml, "http://tempuri.org/CancelaFactura", sRespuesta, "CANCELAR"
    
    CancelaCFD = sRespuesta
    
End Function
'---------------------------------------------------------------------------
Public Function GeneraCFD(NoFolioFac As Long, sSerieCFD As String, sEfecto As String) As String
    
    
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    
    
    Set parser = New DOMDocument
    
    
    
    
    parser.loadXML sSOAPGeneraFactura
    
    'Asigna los valores para el servicio
    'Folio
    parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraFactura/IdFacturaKala").Text = Str(NoFolioFac)
    
    'Serie
    parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraFactura/serie").Text = sSerieCFD
    
    'Efecto
    parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraFactura/EfectoDeComprobante").Text = sEfecto
    
       
    DoEvents
    
       
    
    enviarComando parser.xml, "http://tempuri.org/GeneraFactura", sRespuesta, "GENERAR"
    
    GeneraCFD = sRespuesta
    
End Function
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Public Function GeneraNotaCreditoCFD(NoFolioNota As Long, sSerieCFD As String, sEfecto As String) As String
    
    
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    
    
    Set parser = New DOMDocument
    
    
    
    
    parser.loadXML sSOAPGeneraNotaCredito
    
    'Asigna los valores para el servicio
    'Folio
    parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraNotaCredito/NumeroNota").Text = Str(NoFolioNota)
    
    'Serie
    parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraNotaCredito/serie").Text = sSerieCFD
    
    'Efecto
    parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraNotaCredito/EfectoDeComprobante").Text = sEfecto
    
       
    DoEvents
    
       
    
    enviarComando parser.xml, "http://tempuri.org/GeneraNotaCredito", sRespuesta, "GENERARNOTA"
    
    GeneraNotaCreditoCFD = sRespuesta
    
End Function

'---------------------------------------------------------------------------
Public Function PruebaCFD() As String
    
    
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    
    
    Set parser = New DOMDocument
    
    
    
    
    parser.loadXML sSOAPPruebaServicio
    
        
       
    DoEvents
    
       
    
    enviarComando parser.xml, "http://tempuri.org/PruebaServicio", sRespuesta, "PROBAR"
    
    PruebaCFD = sRespuesta
    
End Function

Private Sub enviarComando(ByVal sXml As String, ByVal sSoapAction As String, ByRef sResp As String, sModo As String)
    ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    Dim oHttReq As XMLHTTP60
    Dim sValor As String
    Dim sUrlWS As String
    Dim timeOut As Double ' timeout in seconds
    Dim timeOutTime As Double
     
    timeOut = 60
    timeOutTime = Timer + timeOut
    
    
    Set oHttReq = New XMLHTTP60
    
    sUrlWS = ObtieneParametro("URL_WS_CFD")
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
    
    Set oHttReq = Nothing
    
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
    Err.Clear
    On Error Resume Next
    '
    Select Case sModo
        Case "GENERAR"
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraFacturaResponse/GeneraFacturaResult").Text
        Case "GENERARNOTA"
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraNotaCreditoResponse/GeneraNotaCreditoResult").Text
        Case "CANCELAR"
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/CancelaFacturaResponse/CancelaFacturaResult").Text
        Case "PROBAR"
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/PruebaServicioResponse/PruebaServicioResult").Text
    End Select
    '
    If Err.Number > 0 Then
        MsgBox "Error al procesar respuesta " & Err.Number & vbCrLf & Err.Description & vbCrLf & s & vbCrLf & procesarRespuesta
    End If
End Function

