Attribute VB_Name = "DireccionesKala"
Option Explicit

Private Const sSOAPInsertDireccionKala = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
      "<soap:Body>" & _
        "<InsertDireccionKala xmlns=""http://www.sportium.com.mx/"">" & _
          "<pIdCliente>int</pIdCliente>" & _
          "<pCalle>string</pCalle>" & _
          "<pNumExt>string</pNumExt>" & _
          "<pNumInt>string</pNumInt>" & _
          "<pIdCodigoPostal>string</pIdCodigoPostal>" & _
          "<pReferencia>string</pReferencia>" & _
          "<pRFC>string</pRFC>" & _
        "</InsertDireccionKala>" & _
      "</soap:Body>" & _
    "</soap:Envelope>"
    
Private Const sSOAPUpDateDireccionKala = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
      "<soap:Body>" & _
        "<UpDateDireccionKala xmlns=""http://www.sportium.com.mx/"">" & _
          "<IdCliente>int</IdCliente>" & _
          "<pIdDireccion>string</pIdDireccion>" & _
          "<pIdCliente>int</pIdCliente>" & _
          "<pCalle>string</pCalle>" & _
          "<pNumExt>string</pNumExt>" & _
          "<pNumInt>string</pNumInt>" & _
          "<pIdCodigoPostal>string</pIdCodigoPostal>" & _
          "<pReferencia>string</pReferencia>" & _
          "<pRFC>string</pRFC>" & _
        "</UpDateDireccionKala>" & _
      "</soap:Body>" & _
    "</soap:Envelope>"

Public Function InsertDireccionKala(ByVal pIdCliente As Long, ByVal pCalle As String, ByVal pNumExt As String, _
            ByVal pNumInt As String, ByVal pIdCodigoPostal As String, ByVal pReferencia As String, ByVal pRFC As String _
            ) As Integer
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    Set parser = New DOMDocument
    
    parser.loadXML sSOAPInsertDireccionKala
    
    'Asigna los valores para el servicio
    'IdCliente
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertDireccionKala/pIdCliente").Text = pIdCliente
    'Calle
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertDireccionKala/pCalle").Text = pCalle
    'NumExt
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertDireccionKala/pNumExt").Text = pNumExt
    'NumInt
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertDireccionKala/pNumInt").Text = pNumInt
    'IdCodigoPostal
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertDireccionKala/pIdCodigoPostal").Text = pIdCodigoPostal
    'Referencia
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertDireccionKala/pReferencia").Text = pReferencia
    'RFC
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertDireccionKala/pRFC").Text = pRFC
    
    DoEvents
    
    enviarComando parser.xml, "http://www.sportium.com.mx/InsertDireccionKala", sRespuesta, "INSERTAR"
    
    InsertDireccionKala = IIf(IsNumeric(sRespuesta), CInt(sRespuesta), 0)
    
End Function

'''
''' Actualiza los datos de la direccion en la base de datos del sistema de CxC
'''
Public Function UpDateDireccionKala(ByVal pIdDireccionCxC As String, ByVal pIdCliente As Long, ByVal pCalle As String, _
            ByVal pNumExt As String, ByVal pNumInt As String, ByVal pIdCodigoPostal As String, ByVal pReferencia As String, _
            ByVal pRFC As String) As String
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    Set parser = New DOMDocument
    
    parser.loadXML sSOAPUpDateDireccionKala
    
    'Asigna los valores para el servicio
    'IdDireccionCxC
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateDireccionKala/pIdDireccion").Text = pIdDireccionCxC
    'IdCliente
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateDireccionKala/pIdCliente").Text = pIdCliente
    'Calle
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateDireccionKala/pCalle").Text = pCalle
    'NumExt
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateDireccionKala/pNumExt").Text = pNumExt
    'NumInt
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateDireccionKala/pNumInt").Text = pNumInt
    'IdCodigoPostal
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateDireccionKala/pIdCodigoPostal").Text = pIdCodigoPostal
    'Referencia
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateDireccionKala/pReferencia").Text = pReferencia
    'RFC
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateDireccionKala/pRFC").Text = pRFC
    
    DoEvents
    
    enviarComando parser.xml, "http://www.sportium.com.mx/UpDateDireccionKala", sRespuesta, "ACTUALIZAR"
    
    UpDateDireccionKala = sRespuesta
    
End Function

'''
''' Actualiza, dentro de la base de Kala, la columna IdDireccionCxC.
'''
Public Sub UpdateIdDireccionCxC(ByVal pIdDireccionKala As Long, ByVal pIdDireccionCxC As Integer)
    Dim iniTrans As Long
    'Dim AdoRcsExAcc As ADODB.Recordset
    Dim AdoCmdInserta As ADODB.Command
    On Err GoTo err_EliminaAccionista
    
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    'Actualiza el id de cliente para el sistema de CxC
    strSQL = "UPDATE DIRECCIONES SET IdDireccionCxC = " & CStr(pIdDireccionCxC) & " WHERE IdDireccion = " & CStr(pIdDireccionKala)
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

'''
''' Obtiene el id de dirección IdDireccionCxC en el sistema Kala.
'''
Public Function SelectIdDireccionCxC(ByVal pIdDireccionKala As Long) As Integer
    Dim AdoRcsExAcc As ADODB.Recordset
    
    strSQL = "SELECT IdDireccionCxC FROM DIRECCIONES WHERE IdDireccion = " & pIdDireccionKala
    Set AdoRcsExAcc = New ADODB.Recordset
    AdoRcsExAcc.ActiveConnection = Conn
    AdoRcsExAcc.CursorLocation = adUseClient
    AdoRcsExAcc.CursorType = adOpenDynamic
    AdoRcsExAcc.LockType = adLockReadOnly
    AdoRcsExAcc.Open strSQL
    
    If Not AdoRcsExAcc.EOF Then
        If Not IsNull(AdoRcsExAcc!IdDireccionCxC) Then SelectIdDireccionCxC = CInt(AdoRcsExAcc!IdDireccionCxC)
    Else
        SelectIdDireccionCxC = 0
    End If
    
    AdoRcsExAcc.Close
    Set AdoRcsExAcc = Nothing
End Function

Private Sub enviarComando(ByVal sXml As String, ByVal sSoapAction As String, ByRef sResp As String, sModo As String)
    ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    Dim oHttReq As XMLHTTP60
    Dim sValor As String
    Dim sUrlWS As String
    
    Set oHttReq = New XMLHTTP60
    
    sUrlWS = ObtieneParametro("URL_WS_DIRECCIONES")
    
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
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/InsertDireccionKalaResponse/InsertDireccionKalaResult").Text
        Case "ACTUALIZAR"
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateDireccionKalaResponse/UpDateDireccionKalaResult").Text
    End Select
    '
    If Err.Number > 0 Then
        MsgBox "Error"
    End If
End Function


