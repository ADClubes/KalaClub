Attribute VB_Name = "CatalogoBancos"
Option Explicit

Private Const sSOAPSelectDataSetBancosCatalogoBancos = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
      "<soap:Body>" & _
        "<SelectDataSetBancosCatalogoBancos xmlns=""http://www.sportium.com.mx/"" />" & _
      "</soap:Body>" & _
    "</soap:Envelope>"

Public Function SelectDataSetBancosCatalogoBancos() As Variant
    Dim parser As DOMDocument
    Dim arrRespuesta As Variant
    
    Set parser = New DOMDocument
    
    parser.loadXML sSOAPSelectDataSetBancosCatalogoBancos
    
    DoEvents
    
    enviarComando parser.xml, "http://www.sportium.com.mx/SelectDataSetBancosCatalogoBancos", arrRespuesta, "CONSULTAR"
    
    SelectDataSetBancosCatalogoBancos = arrRespuesta
    
End Function

Private Sub enviarComando(ByVal sXml As String, ByVal sSoapAction As String, ByRef vResp As Variant, sModo As String)
    ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    Dim oHttReq As XMLHTTP60
    Dim vValor As Variant
    Dim sUrlWS As String
    
    Set oHttReq = New XMLHTTP60
    
    sUrlWS = ObtieneParametro("URL_WS_BANCOS")
    
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
    vValor = procesarRespuesta(oHttReq.responseText, sModo)
    '
    
    vResp = vValor
    
End Sub

Private Function procesarRespuesta(ByVal sXml As String, sModo As String) As Variant
    Dim parser As DOMDocument
    Dim xmlNode As IXMLDOMNode
    Set parser = New DOMDocument
    parser.loadXML sXml
    Dim index As Integer
    Dim intLength As Integer
    Dim arrBancos() As String
    '
    On Error Resume Next
    '
    Select Case sModo
        Case "CONSULTAR"
            intLength = parser.getElementsByTagName("Table").Length - 1
            If intLength < 0 Then Exit Function
            
            ReDim arrBancos(intLength, 1)
            
            For Each xmlNode In parser.getElementsByTagName("Table")
                
                arrBancos(index, 0) = xmlNode.childNodes(0).Text
                arrBancos(index, 1) = xmlNode.childNodes(1).Text
                
                index = index + 1
            Next xmlNode
            
    End Select
    '
    procesarRespuesta = arrBancos
    '
    If Err.Number > 0 Then
        MsgBox "Error"
    End If
End Function




