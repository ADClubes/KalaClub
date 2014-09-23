Attribute VB_Name = "ConvenioDireccionado"
Option Explicit

Private Const sSOAPInsertConvenioKala = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
      "<soap:Body>" & _
        "<InsertConvenioKala xmlns=""http://www.sportium.com.mx/"">" & _
          "<pIdCliente>int</pIdCliente>" & _
          "<pCodigo>string</pCodigo>" & _
          "<pTelefono>string</pTelefono>" & _
          "<pExtension>string</pExtension>" & _
          "<pNombreTitular>string</pNombreTitular>" & _
          "<pNumeroTarjeta>string</pNumeroTarjeta>" & _
          "<pFechaInicioTarjeta>dateTime</pFechaInicioTarjeta>" & _
          "<pFechaVigencia>dateTime</pFechaVigencia>" & _
          "<TipoTarjeta>string</TipoTarjeta>" & _
          "<pIdEmisorTarjeta>int</pIdEmisorTarjeta>" & _
          "<pCodigoSeguridad>string</pCodigoSeguridad>" & _
          "<pCodigoContrato>string</pCodigoContrato>" & _
          "<pIdSucursal>int</pIdSucursal>" & _
          "<pFechaInicioContrato>dateTime</pFechaInicioContrato>" & _
          "<pFechaTerminacionContrato>dateTime</pFechaTerminacionContrato>" & _
          "<pMontoContrato>decimal</pMontoContrato>" & _
          "<pIdBanco>int</pIdBanco>" & _
        "</InsertConvenioKala>" & _
      "</soap:Body>" & _
    "</soap:Envelope>"

Public Function InsertConvenioKala(ByVal pIdCliente As Long, ByVal pCodigo As String, ByVal pTelefono As String, ByVal pExtension As String, _
        ByVal pNombreTitular As String, ByVal pNumeroTarjeta As String, ByVal pFechaInicioTarjeta As Date, _
        ByVal pFechaVigencia As Date, ByVal TipoTarjeta As String, ByVal pIdEmisorTarjeta As Integer, _
        ByVal pCodigoSeguridad As String, ByVal pCodigoContrato As String, ByVal pIdSucursal As Integer, _
        ByVal pFechaInicioContrato As Date, ByVal pFechaTerminacionContrato As Date, ByVal pMontoContrato As Double, _
        ByVal pIdBanco As Integer) As String
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    Set parser = New DOMDocument
    
    parser.loadXML sSOAPInsertConvenioKala
    
    'Asigna los valores para el servicio
    'IdCliente
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pIdCliente").Text = pIdCliente
    'Codigo
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pCodigo").Text = pCodigo
    'Telefono
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pTelefono").Text = pTelefono
    'Extension
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pExtension").Text = pExtension
    'NombreTitular
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pNombreTitular").Text = pNombreTitular
    'NumeroTarjeta
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pNumeroTarjeta").Text = pNumeroTarjeta
    'FechaInicioTarjeta
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pFechaInicioTarjeta").Text = Format(pFechaInicioTarjeta, "yyyy-mm-ddTh:mm:ss")
    'FechaVigencia
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pFechaVigencia").Text = Format(pFechaVigencia, "yyyy-mm-ddTh:mm:ss")
    'TipoTarjeta
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/TipoTarjeta").Text = TipoTarjeta
    'IdEmisorTarjeta
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pIdEmisorTarjeta").Text = pIdEmisorTarjeta
    'CodigoSeguridad
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pCodigoSeguridad").Text = pCodigoSeguridad
    'CodigoContrato
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pCodigoContrato").Text = pCodigoContrato
    'IdSucursal
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pIdSucursal").Text = pIdSucursal
    'FechaInicioContrato
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pFechaInicioContrato").Text = Format(pFechaInicioContrato, "yyyy-mm-ddTh:mm:ss")
    'FechaTerminacionContrato
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pFechaTerminacionContrato").Text = Format(pFechaTerminacionContrato, "yyyy-mm-ddTh:mm:ss")
    'MontoContrato
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pMontoContrato").Text = pMontoContrato
    'IdBanco
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKala/pIdBanco").Text = pIdBanco
    
    DoEvents
    
    enviarComando parser.xml, "http://www.sportium.com.mx/InsertConvenioKala", sRespuesta, "INSERTAR"
    
    InsertConvenioKala = sRespuesta
    
End Function
    
Public Sub UpdateIdConvenioCxC(ByVal pIdConvenioKala As Long, ByVal pIdConvenioCxC As Long)
    Dim iniTrans As Long
    Dim AdoCmdInserta As ADODB.Command
    On Err GoTo err_EliminaAccionista
    
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    'Actualiza el id de cliente para el sistema de CxC
    strSQL = "UPDATE DireccionadosDatos SET IdConvenioCxC = " & CStr(pIdConvenioCxC) & " WHERE IdConvenio = " & CStr(pIdConvenioKala)
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
                    
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    
    Exit Sub
    
err_EliminaAccionista:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub enviarComando(ByVal sXml As String, ByVal sSoapAction As String, ByRef sResp As String, sModo As String)
    ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    Dim oHttReq As XMLHTTP60
    Dim sValor As String
    Dim sUrlWS As String
    
    Set oHttReq = New XMLHTTP60
    
    sUrlWS = ObtieneParametro("URL_WS_DIRECCIONADOS")
    
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
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/InsertConvenioKalaResponse/InsertConvenioKalaResult").Text
    End Select
    '
    If Err.Number > 0 Then
        MsgBox "Error"
    End If
End Function
