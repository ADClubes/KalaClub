Attribute VB_Name = "CatalogoCliente"
Option Explicit

Private Const sSOAPInsertClienteKala = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
          "<InsertClienteKala xmlns=""http://www.sportium.com.mx/"">" & _
            "<pISOPais>string</pISOPais>" & _
            "<pISOEstado>string</pISOEstado>" & _
            "<pPersonalidad>string</pPersonalidad>" & _
            "<pGenero>string</pGenero>" & _
            "<pNombreRazonSocial>string</pNombreRazonSocial>" & _
            "<pApellidoPaterno>string</pApellidoPaterno>" & _
            "<pApellidoMaterno>string</pApellidoMaterno>" & _
            "<pCURP>string</pCURP>" & _
            "<pIdSucursal>int</pIdSucursal>" & _
            "<pRFC>string</pRFC>" & _
          "</InsertClienteKala>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"

Private Const sSOAPUpdateClienteKala = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
      "<soap:Body>" & _
        "<UpDateClienteKala xmlns=""http://www.sportium.com.mx/"">" & _
          "<IdCliente>int</IdCliente>" & _
          "<pISOPais>string</pISOPais>" & _
          "<pISOEstado>string</pISOEstado>" & _
          "<pPersonalidad>string</pPersonalidad>" & _
          "<pGenero>string</pGenero>" & _
          "<pNombreRazonSocial>string</pNombreRazonSocial>" & _
          "<pApellidoPaterno>string</pApellidoPaterno>" & _
          "<pApellidoMaterno>string</pApellidoMaterno>" & _
          "<pCURP>string</pCURP>" & _
          "<pIdSucursal>int</pIdSucursal>" & _
          "<pRFC>string</pRFC>" & _
        "</UpDateClienteKala>" & _
      "</soap:Body>" & _
    "</soap:Envelope>"
    
Public Sub UpdateIdClienteCxC(ByVal pIdClienteKala As Long, ByVal pIdClienteCxC As Long)
    Dim iniTrans As Long
    'Dim AdoRcsExAcc As ADODB.Recordset
    Dim AdoCmdInserta As ADODB.Command
    On Err GoTo err_EliminaAccionista
    
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    'Actualiza el id de cliente para el sistema de CxC
    strSQL = "UPDATE USUARIOS_CLUB SET IdClienteCxC = " & CStr(pIdClienteCxC) & " WHERE IdMember = " & CStr(pIdClienteKala)
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

Public Function SelectIdClienteCxC(ByVal pIdClienteKala As Long) As Long
    Dim AdoRcsExAcc As ADODB.Recordset
    
    strSQL = "SELECT IdClienteCxC FROM USUARIOS_CLUB WHERE IdMember = " & pIdClienteKala
    Set AdoRcsExAcc = New ADODB.Recordset
    AdoRcsExAcc.ActiveConnection = Conn
    AdoRcsExAcc.CursorLocation = adUseClient
    AdoRcsExAcc.CursorType = adOpenDynamic
    AdoRcsExAcc.LockType = adLockReadOnly
    AdoRcsExAcc.Open strSQL
    
    If Not AdoRcsExAcc.EOF Then
        AdoRcsExAcc.MoveFirst
        If Not IsNull(AdoRcsExAcc!IdClienteCxC) Then
            SelectIdClienteCxC = CLng(AdoRcsExAcc!IdClienteCxC)
        Else
            SelectIdClienteCxC = 0
        End If
    Else
        SelectIdClienteCxC = 0
    End If
    
    AdoRcsExAcc.Close
    Set AdoRcsExAcc = Nothing
End Function

Public Function InsertClienteKala(ByVal pISOPais As String, ByVal pISOEstado As String, _
        ByVal pPersonalidad As String, ByVal pGenero As String, ByVal pNombreRazonSocial As String, _
        ByVal pApellidoPaterno As String, ByVal pApellidoMaterno As String, ByVal pCURP As String, _
        ByVal pIdSucursal As Integer, ByVal pRFC As String) As Long
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    Set parser = New DOMDocument
    
    parser.loadXML sSOAPInsertClienteKala
    
    'Asigna los valores para el servicio
    'Pais
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pISOPais").Text = pISOPais
    'Estado
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pISOEstado").Text = pISOEstado
    'Personalidad
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pPersonalidad").Text = pPersonalidad
    'Genero
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pGenero").Text = pGenero
    'NombreRazonSocial
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pNombreRazonSocial").Text = pNombreRazonSocial
    'ApellidoPaterno
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pApellidoPaterno").Text = pApellidoPaterno
    'ApellidoMaterno
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pApellidoMaterno").Text = pApellidoMaterno
    'CURP
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pCURP").Text = pCURP
    'IdSucursal
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pIdSucursal").Text = pIdSucursal
    'RFC
    parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKala/pRFC").Text = pRFC
    
    DoEvents
    
    enviarComando parser.xml, "http://www.sportium.com.mx/InsertClienteKala", sRespuesta, "INSERTAR"
    
    InsertClienteKala = IIf(IsNumeric(sRespuesta), CLng(sRespuesta), 0)
    
End Function

Public Function UpdateClienteKala(ByVal pIdClienteCxC As Long, ByVal pISOPais As String, ByVal pISOEstado As String, _
        ByVal pPersonalidad As String, ByVal pGenero As String, ByVal pNombreRazonSocial As String, _
        ByVal pApellidoPaterno As String, ByVal pApellidoMaterno As String, ByVal pCURP As String, _
        ByVal pIdSucursal As Integer, ByVal pRFC As String) As Boolean
    Dim parser As DOMDocument
    Dim sRespuesta As String
    
    Set parser = New DOMDocument
    
    parser.loadXML sSOAPUpdateClienteKala
    
    'Asigna los valores para el servicio
    'IdClienteCxC
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/IdCliente").Text = CStr(pIdClienteCxC)
    'Pais
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pISOPais").Text = pISOPais
    'Estado
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pISOEstado").Text = pISOEstado
    'Personalidad
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pPersonalidad").Text = pPersonalidad
    'Genero
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pGenero").Text = pGenero
    'NombreRazonSocial
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pNombreRazonSocial").Text = pNombreRazonSocial
    'ApellidoPaterno
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pApellidoPaterno").Text = pApellidoPaterno
    'ApellidoMaterno
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pApellidoMaterno").Text = pApellidoMaterno
    'CURP
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pCURP").Text = pCURP
    'IdSucursal
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pIdSucursal").Text = pIdSucursal
    'RFC
    parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKala/pRFC").Text = pRFC
    
    DoEvents
    
    enviarComando parser.xml, "http://www.sportium.com.mx/UpDateClienteKala", sRespuesta, "ACTUALIZAR"
    
    UpdateClienteKala = (sRespuesta <> "")
    
End Function

Private Sub enviarComando(ByVal sXml As String, ByVal sSoapAction As String, ByRef sResp As String, sModo As String)
    ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    Dim oHttReq As XMLHTTP60
    Dim sValor As String
    Dim sUrlWS As String
    
    Set oHttReq = New XMLHTTP60
    
    sUrlWS = ObtieneParametro("URL_WS_CLIENTES")
    
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
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/InsertClienteKalaResponse/InsertClienteKalaResult").Text
        Case "ACTUALIZAR"
            procesarRespuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/UpDateClienteKalaResponse/UpDateClienteKalaResult").Text
    End Select
    '
    If Err.Number > 0 Then
        MsgBox "Error"
    End If
End Function
