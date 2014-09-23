Attribute VB_Name = "accesoModulo"
Public Function AccesoValido(lidMember As Long) As Boolean

    Dim adorcs As ADODB.Recordset
    Dim dFechaUPago As Date
    Dim lPer As Long
    Dim lIdTitular As Long
    
    Const nPerMax = 1
    Const nDiasMax = 15
    
    AccesoValido = False
    
    
    strSQL = "SELECT IdMember From Ausencias"
    strSQL = strSQL & " Where ("
    strSQL = strSQL & "((IdMember)=" & lidMember & ")"
    strSQL = strSQL & ")"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        adorcs.Close
        Set adorcs = Nothing
        Exit Function
    End If
    
    
    adorcs.Close

    strSQL = "SELECT FECHAS_USUARIO.FechaUltimoPago, USUARIOS_CLUB.IdTitular"
    strSQL = strSQL & " FROM FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember = USUARIOS_CLUB.IdMember"
    strSQL = strSQL & " Where (((FECHAS_USUARIO.IdMember) =  " & lidMember& & "))"
    strSQL = strSQL & " ORDER BY  FECHAS_USUARIO.FechaUltimoPago"
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        dFechaUPago = adorcs!Fechaultimopago
        lIdTitular = adorcs!IdTitular
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    
    lPer = CalculaPeriodos(dFechaUPago, Date, 1)
    
    
    'Si el número de periodos es mayor que el maximo permitido
    If lPer > nPerMax Then
        Exit Function
    End If
    
    If ChecaDireccionado(lIdTitular) <> vbNullString Then 'Si es direccionado
        If lPer < nPerMax Then
            AccesoValido = True
        ElseIf lPer = nPerMax Then
            If Day(Date) <= nDiasMax Then
                AccesoValido = True
            End If
        End If
    Else 'Si es convencional
        If lPer < nPerMax Then
            AccesoValido = True
        ElseIf lPer = nPerMax Then
            If Day(Date) <= nDiasMax Then
                AccesoValido = True
            End If
        End If
    End If
    
    
    

End Function

Public Sub ActAccesoXUsu(lidMember As Long, bActiva As Boolean)

    Dim adorcs As ADODB.Recordset
    Dim sSqlQry As String
    Dim nErrCode As Long
    sSqlQry = "SELECT SECUENCIAL.Secuencial"
    sSqlQry = sSqlQry & " FROM SECUENCIAL INNER JOIN USUARIOS_CLUB ON SECUENCIAL.IdMember = USUARIOS_CLUB.IdMember"
    sSqlQry = sSqlQry & " Where ("
    sSqlQry = sSqlQry & "((USUARIOS_CLUB.IdMember)=" & lidMember & ")"
    sSqlQry = sSqlQry & ")"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open sSqlQry, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        #If SqlServer_ Then
        'Modulo de Ausencias
            ActivaCredSQL 1, adorcs!Secuencial, 1, lidMember, bActiva, False
        #Else
            ActivaCred 1, adorcs!Secuencial, 1, lidMember, bActiva, False
        #End If
         'nErrCode = BloqueaAcceso(lidMember)
    End If
    
    adorcs.Close
    
    Set adorcrs = Nothing
    

End Sub
Public Sub Esperar(lSecs)
    Dim lInicio As Long
    
    lInicio = Timer
    
    Do Until lInicio + lSecs < Timer
        DoEvents
    Loop
    
End Sub

