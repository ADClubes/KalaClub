Attribute VB_Name = "Funciones"
Public Function BuscaInsc(lNumInsc As Long) As Boolean
    Dim adorcs As ADODB.Recordset
    Dim iUnidad As Integer
    
    BuscaInsc = False
    
    iUnidad = Val(ObtieneParametro("NUMERO_CLUB"))
    
    strSQL = "SELECT Unidad, NoInscripcion, FechaAlta"
    strSQL = strSQL & " FROM CuponesPorEntregar"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "(NoInscripcion)=" & lNumInsc
    strSQL = strSQL & " And (Unidad)=" & iUnidad
    strSQL = strSQL & ")"
    
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If adorcs.EOF Then
        
        adorcs.Close
        Set adorcs = Nothing
        
        Exit Function
    End If
    
    adorcs.Close
    Set adorcs = Nothing

    
    BuscaInsc = True
    
End Function

Public Function DatosInscripcion(lNumInsc As Long)
    Dim adorcs As ADODB.Recordset
    
    DatosInscripcion = vbNullString
    
    strSQL = "SELECT USUARIOS_CLUB.A_Paterno & ' ' &  USUARIOS_CLUB.A_Materno & ' ' &  USUARIOS_CLUB.Nombre AS Nombre"
    strSQL = strSQL & " From usuarios_club"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((USUARIOS_CLUB.NoFamilia)=" & lNumInsc & ")"
    strSQL = strSQL & " AND ((USUARIOS_CLUB.IdMember)=[usuarios_club].[idtitular])"
    strSQL = strSQL & ")"
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        DatosInscripcion = adorcs!Nombre
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
End Function
Public Function DatosPago(lNumInsc As Long)
    Dim adorcs As ADODB.Recordset
    
    DatosPago = vbNullString
    
    
    strSQL = "SELECT FECHAS_USUARIO.FechaUltimoPago"
    strSQL = strSQL & " FROM USUARIOS_CLUB INNER JOIN FECHAS_USUARIO ON USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((USUARIOS_CLUB.NoFamilia)=" & lNumInsc & ")"
    strSQL = strSQL & " AND ((USUARIOS_CLUB.IdMember)=[usuarios_club].[idtitular])"
    strSQL = strSQL & ")"
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        DatosPago = adorcs!FechaUltimoPago
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
End Function

Public Function StatusCupones(lNumInsc As Long) As Integer
    
    Dim adorcs As ADODB.Recordset
    Dim iUnidad As Integer
    
    StatusCupones = 0
    
    iUnidad = Val(ObtieneParametro("NUMERO_CLUB"))
    
    strSQL = "SELECT NoInscripcion, FechaCreacion"
    strSQL = strSQL & " FROM CuponesRegalo"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "(NoInscripcion)=" & lNumInsc
    strSQL = strSQL & " And (Unidad)=" & iUnidad
    strSQL = strSQL & ")"
    
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        StatusCupones = 1
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
    
End Function

Public Function GeneraCupones(lNumInsc As Long, lNumCupones As Integer, lDiasVigencia As Long, sConcepto As String) As Boolean
    Dim adocmd As ADODB.Command
    
    Dim lFolio As Long
    Dim iContador As Integer
    Dim iUnidad As Integer
    
    GeneraCupones = True
    
    lFolio = UltimoValorLong("CuponesRegalo", "Folio") + 1
    iUnidad = Val(ObtieneParametro("NUMERO_CLUB"))
    
    strSQL = "INSERT INTO CuponesRegalo ("
    strSQL = strSQL & " Unidad" & ","
    strSQL = strSQL & " NoInscripcion" & ","
    strSQL = strSQL & " FechaCreacion" & ","
    strSQL = strSQL & " HoraCreacion" & ","
    strSQL = strSQL & " Folio" & ","
    strSQL = strSQL & " TotalCupones"
    strSQL = strSQL & ")"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & iUnidad & ","
    strSQL = strSQL & lNumInsc & ","
    strSQL = strSQL & "'" & Format(Date, "dd/mm/yyyy") & "'" & ","
    strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "'" & ","
    strSQL = strSQL & lFolio & ","
    strSQL = strSQL & lNumCupones
    strSQL = strSQL & ")"
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    For iContador = 1 To lNumCupones
    
        strSQL = "INSERT INTO CuponesRegaloDetalle ("
        strSQL = strSQL & " Folio" & ","
        strSQL = strSQL & " Consecutivo" & ","
        strSQL = strSQL & " Concepto" & ","
        strSQL = strSQL & " Vigencia"
        strSQL = strSQL & ")"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lFolio & ","
        strSQL = strSQL & iContador & ","
        strSQL = strSQL & "'" & sConcepto & "'" & ","
        strSQL = strSQL & "'" & Format(Date + lDiasVigencia, "dd/mm/yyyy") & "'"
        strSQL = strSQL & ")"
    
        adocmd.CommandText = strSQL
        adocmd.Execute
        
    Next
    
    
    
End Function

Public Function UltimoValorLong(sNombreTabla, sNombreCampo) As Long
    Dim adorcs As ADODB.Recordset
    
    UltimoValorLong = 0
    
    strSQL = "SELECT Max(" & sNombreCampo & ") As Ultimo"
    strSQL = strSQL & " FROM " & sNombreTabla
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        If IsNull(adorcs!Ultimo) Then
            UltimoValorLong = 0
        Else
            UltimoValorLong = adorcs!Ultimo
        End If
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
    
End Function
Public Sub DisplayCupones(lNumInsc As Long)
    
    strSQL = "SELECT CuponesRegalo.Folio, CuponesRegaloDetalle.Consecutivo, CuponesRegalo.FechaCreacion, CuponesRegalo.HoraCreacion, CuponesRegaloDetalle.Concepto, CuponesRegaloDetalle.Vigencia, CuponesRegalo.Impresiones"
    strSQL = strSQL & " FROM CuponesRegalo INNER JOIN CuponesRegaloDetalle ON CuponesRegalo.Folio = CuponesRegaloDetalle.Folio"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((CuponesRegalo.NoInscripcion)=" & lNumInsc & ")"
    strSQL = strSQL & ")"
    
    LlenaSsDbGrid frmCtrlCUpon.ssdbgCupones, Conn, strSQL, 7
    
End Sub


Public Function SelectPrinter(sNombreImpresora As String) As String
    Dim CurrentPrinter As Printer
    Dim P As Printer
    
    SelectPrinter = Printer.DeviceName
    
    For Each P In Printers
        If P.DeviceName = sNombreImpresora Then
            Set Printer = P
            Exit For
        End If
    Next
    
End Function

Public Function ActualContImp(lFolio As Long) As Boolean
    
    Dim adocmd As ADODB.Command
    
    ActualContImp = True
    
    strSQL = "UPDATE CuponesRegalo SET "
    strSQL = strSQL & "Impresiones = " & "Impresiones + 1"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "(Folio) = " & lFolio
    strSQL = strSQL & ")"
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    Set adocmd = Nothing
    
End Function
