Attribute VB_Name = "FuncionesRentables"
Public Function VigenciaRentable(sNumeroRentable As String, ByRef dFechaPago As Date) As Boolean
    Dim adorcs As ADODB.Recordset
    
    
    VigenciaRentable = False
    
    strSQL = "SELECT IdUsuario, FechaPago"
    strSQL = strSQL & " FROM RENTABLES"
    strSQL = strSQL & " WHERE Numero='" & Format(Trim(sNumeroRentable), "@@@@@@") & "'"
    
    Set adorcs = New ADODB.Recordset
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        If IsNull(adorcs!idusuario) Or adorcs!idusuario = 0 Then
            dFechaPago = CDate("31/12/1980")
        Else
            dFechaPago = adorcs!Fechapago
        End If
        VigenciaRentable = True
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    
    
End Function
