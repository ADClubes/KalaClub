Attribute VB_Name = "modTurnos"
Public Function OpenShiftF() As Long
    Dim adorcsTurno As ADODB.Recordset
    
    
    OpenShiftF = 0
    
    
    #If SqlServer_ Then
        strSQL = "SELECT TurnoNo"
        strSQL = strSQL & " FROM Turnos"
        strSQL = strSQL & " WHERE "
        '29/06/2007
        strSQL = strSQL & " NumeroCaja=" & iNumeroCaja
        strSQL = strSQL & " AND Cerrado=0"
        strSQL = strSQL & " AND FechaApertura= '" & Format(Date, "yyyymmdd") & "'"
    #Else
        strSQL = "SELECT TurnoNo"
        strSQL = strSQL & " FROM Turnos"
        strSQL = strSQL & " WHERE "
        '29/06/2007
        strSQL = strSQL & " NumeroCaja=" & iNumeroCaja
        strSQL = strSQL & " AND Cerrado=0"
        strSQL = strSQL & " AND FechaApertura=" & "#" & Format(Date, "mm/dd/yyyy") & "#"
    #End If

    Set adorcsTurno = New ADODB.Recordset
    adorcsTurno.CursorLocation = adUseServer
    adorcsTurno.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsTurno.EOF Then
        OpenShiftF = adorcsTurno!TurnoNo
    End If
    
    adorcsTurno.Close
    Set adorcsTurno = Nothing
    
    
End Function


'Regresa el número de corte asignado a un turno y caja
Function BuscaIdCorte(dFecha As Date, nCaja As Long, nTurno As Long) As Long

    Dim adorcs As ADODB.Recordset
    
    
    BuscaIdCorte = 0
    
    #If SqlServer_ Then
        strSQL = "SELECT IdCorteCaja"
        strSQL = strSQL & " FROM CORTE_CAJA"
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " FechaCorte= '" & Format(dFecha, "yyyymmdd") & "'"
        strSQL = strSQL & " AND Caja = " & nCaja
        strSQL = strSQL & " AND Turno = " & nTurno
    #Else
        strSQL = "SELECT IdCorteCaja"
        strSQL = strSQL & " FROM CORTE_CAJA"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FechaCorte)= #" & Format(dFecha, "mm/dd/yyyy") & "#" & ")"
        strSQL = strSQL & " AND ((Caja) = " & nCaja & ")"
        strSQL = strSQL & " AND ((Turno) = " & nTurno & ")"
        strSQL = strSQL & ")"
    #End If
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        BuscaIdCorte = adorcs!IdCorteCaja
    End If
    
    adorcs.Close
    Set adorcs = Nothing

End Function

Public Function OpenShift() As Long
    Dim adorcsTurno As ADODB.Recordset
    
    
    OpenShift = 0
    
    strSQL = "SELECT TurnoNo"
    strSQL = strSQL & " FROM Turnos"
    strSQL = strSQL & " WHERE "
    '29/06/2007
    strSQL = strSQL & " NumeroCaja=" & iNumeroCaja
    strSQL = strSQL & " AND Cerrado=0"

    Set adorcsTurno = New ADODB.Recordset
    adorcsTurno.CursorLocation = adUseServer
    adorcsTurno.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsTurno.EOF Then
        OpenShift = adorcsTurno!TurnoNo
    End If
    
    adorcsTurno.Close
    Set adorcsTurno = Nothing
    
    
End Function

Public Function NextShift() As Long

    Dim adorcsTurno As ADODB.Recordset
    
    
    NextShift = 0
    
    #If SqlServer_ Then
        strSQL = "SELECT Max(TurnoNo) AS Ultimo"
        strSQL = strSQL & " FROM Turnos"
        strSQL = strSQL & " WHERE "
        '29/6/2007
        strSQL = strSQL & " NumeroCaja =" & iNumeroCaja
        strSQL = strSQL & " AND FechaApertura = " & "'" & Format(Date, "yyyymmdd") & "'"
        strSQL = strSQL & " AND FechaCierre > " & "'01/01/1980'"
    #Else
        strSQL = "SELECT Max(TurnoNo) AS Ultimo"
        strSQL = strSQL & " FROM Turnos"
        strSQL = strSQL & " WHERE "
        '29/6/2007
        strSQL = strSQL & " NumeroCaja =" & iNumeroCaja
        strSQL = strSQL & " AND FechaApertura = " & "#" & Format(Date, "mm/dd/yyyy") & "#"
        strSQL = strSQL & " AND FechaCierre > " & "#01/01/1980#"
    #End If
    Set adorcsTurno = New ADODB.Recordset
    adorcsTurno.CursorLocation = adUseServer
    adorcsTurno.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If adorcsTurno.EOF Then
        NextShift = 1
    Else
        If IsNull(adorcsTurno!Ultimo) Then
            NextShift = 1
        Else
            NextShift = adorcsTurno!Ultimo + 1
        End If
    End If
    
    adorcsTurno.Close
    Set adorcsTurno = Nothing

End Function

Public Function ClosedShift(lTurno As Long) As Boolean
    Dim adorcsTurno As ADODB.Recordset
    
    
    ClosedShift = False
    
    strSQL = "SELECT Cerrado"
    strSQL = strSQL & " FROM Turnos"
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " IdTurno=" & lTurno

    Set adorcsTurno = New ADODB.Recordset
    adorcsTurno.CursorLocation = adUseServer
    adorcsTurno.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsTurno.EOF Then
        ClosedShift = adorcsTurno!Cerrado
    End If
    
    adorcsTurno.Close
    Set adorcsTurno = Nothing
    
    
End Function
Public Function RecibosPendientes(dFecha As Date, lCajaNum As Long, lTurnoNum As Long) As Long
    Dim adorcsRec As ADODB.Recordset
    
    RecibosPendientes = 0
    
    #If SqlServer_ Then
        strSQL = "SELECT Count(*) AS Cuenta"
        strSQL = strSQL & " FROM RECIBOS"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " (FechaFactura)=" & "'" & Format(dFecha, "yyyymmdd") & "'"
        strSQL = strSQL & " AND (Caja)=" & lCajaNum
        strSQL = strSQL & " AND (Turno)=" & lTurnoNum
        strSQL = strSQL & " AND (Cancelada) = 0"
        strSQL = strSQL & " AND (Factura) = 0"
        strSQL = strSQL & ")"
    #Else
        strSQL = "SELECT Count(*) AS Cuenta"
        strSQL = strSQL & " FROM RECIBOS"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " (FechaFactura)=" & "#" & Format(dFecha, "mm/dd/yyyy") & "#"
        strSQL = strSQL & " AND (Caja)=" & lCajaNum
        strSQL = strSQL & " AND (Turno)=" & lTurnoNum
        strSQL = strSQL & " AND (Cancelada) = 0"
        strSQL = strSQL & " AND (Factura) = 0"
        strSQL = strSQL & ")"
    #End If
        
    Set adorcsRec = New ADODB.Recordset
    adorcsRec.CursorLocation = adUseServer
    adorcsRec.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsRec.EOF Then
        If IsNull(adorcsRec!Cuenta) Then
            RecibosPendientes = 0
        Else
            RecibosPendientes = adorcsRec!Cuenta
        End If
    End If
    
    adorcsRec.Close
    Set adorcsRec = Nothing
    
    
    
    
End Function

Public Function ValidaCierre(dFechaTurno As Date, lCaja As Long, lTurno As Long, dFechaCierre As Date, dHoraCierre As Date) As Boolean

    Dim adorcs As ADODB.Recordset
    
    ValidaCierre = True
    
    #If SqlServer_ Then
        strSQL = "SELECT TURNOS.FechaApertura, TURNOS.HoraApertura, TURNOS.NumeroCaja, TURNOS.TurnoNo, TURNOS.Cerrado"
        strSQL = strSQL & " From Turnos"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((TURNOS.FechaApertura)='" & Format(dFechaTurno, "yyyymmdd") & "')"
        strSQL = strSQL & " AND ((TURNOS.NumeroCaja)=" & lCaja & ")"
        strSQL = strSQL & " AND ((TURNOS.TurnoNo)=" & lTurno & ")"
        strSQL = strSQL & ")"
    #Else
        strSQL = "SELECT TURNOS.FechaApertura, TURNOS.HoraApertura, TURNOS.NumeroCaja, TURNOS.TurnoNo, TURNOS.Cerrado"
        strSQL = strSQL & " From Turnos"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((TURNOS.FechaApertura)=#" & Format(dFechaTurno, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((TURNOS.NumeroCaja)=" & lCaja & ")"
        strSQL = strSQL & " AND ((TURNOS.TurnoNo)=" & lTurno & ")"
        strSQL = strSQL & ")"
    #End If

    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
    
        If dFechaCierre < adorcs!FechaApertura Then
            ValidaCierre = False
            Exit Function
        End If
    
        If dFechaCierre < adorcs!FechaApertura And dHoraCierre < adorcs!HoraApertura Then
            ValidaCierre = False
            Exit Function
        End If
    
    
    
        
    End If
    
    
    adorcs.Close
    Set adorcs = Nothing
    
End Function

Public Function GetIdCorte(dFecha As Date, iCaja As Integer, lTurno As Long) As Long
    Dim adorcsTurno As ADODB.Recordset
    
    
    GetIdCorte = 0
    
    #If SqlServer_ Then
        strSQL = "SELECT IdTurno"
        strSQL = strSQL & " FROM Turnos"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((FechaApertura)='" & Format(dFecha, "yyyymmdd") & "')"
        strSQL = strSQL & " AND ((NumeroCaja)=" & iCaja & ")"
        strSQL = strSQL & " AND ((TurnoNo)=" & lTurno & ")"
        strSQL = strSQL & ")"
    #Else
        strSQL = "SELECT IdTurno"
        strSQL = strSQL & " FROM Turnos"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((FechaApertura)=#" & Format(dFecha, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((NumeroCaja)=" & iCaja & ")"
        strSQL = strSQL & " AND ((TurnoNo)=" & lTurno & ")"
        strSQL = strSQL & ")"
    #End If

    Set adorcsTurno = New ADODB.Recordset
    adorcsTurno.CursorLocation = adUseServer
    adorcsTurno.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsTurno.EOF Then
        GetIdCorte = adorcsTurno!IdTurno
    End If
    
    adorcsTurno.Close
    Set adorcsTurno = Nothing
End Function

Public Function CloseShift(lIdTurno As Long) As Boolean
    
    Dim adocmdTurnos As ADODB.Command
    
    CloseShift = True
    
    #If SqlServer_ Then
        strSQL = "UPDATE TURNOS SET"
        strSQL = strSQL & " FechaCierre=" & "'" & Format(Now, "yyyymmdd") & "',"
        strSQL = strSQL & " HoraCierre=" & "'" & Format(Now, "Hh:Nn:Ss") & "',"
        strSQL = strSQL & " UsuarioCierre=" & "'" & sDB_User & "',"
        strSQL = strSQL & " Cerrado=-1"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdTurno=" & lIdTurno
    #Else
        strSQL = "UPDATE TURNOS SET"
        strSQL = strSQL & " FechaCierre=" & "#" & Format(Now, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & " HoraCierre=" & "'" & Format(Now, "Hh:Nn:Ss") & "',"
        strSQL = strSQL & " UsuarioCierre=" & "'" & sDB_User & "',"
        strSQL = strSQL & " Cerrado=-1"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdTurno=" & lIdTurno
    #End If
    
    Set adocmdTurnos = New ADODB.Command
    adocmdTurnos.ActiveConnection = Conn
    adocmdTurnos.CommandType = adCmdText
    adocmdTurnos.CommandText = strSQL
    adocmdTurnos.Execute
    
    Set adocmdTurnos = Nothing
    
    
    
End Function

Public Function DatosTurno(dFecha As Date, lCaja As Long, lTurno As Long, ByRef sFolioIni As String, ByRef sSerieIni As String, ByRef dFondoIni As Double) As Boolean
    
    Dim adorcs As ADODB.Recordset
    
    DatosTurno = True
    
    #If SqlServer_ Then
        strSQL = "SELECT FolioApertura, SerieApertura, FondoApertura"
        strSQL = strSQL & " FROM TURNOS"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((FechaApertura)='" & Format(dFecha, "yyyymmdd") & "')"
        strSQL = strSQL & " AND ((NumeroCaja)=" & lCaja & ")"
        strSQL = strSQL & " AND ((TurnoNo)=" & lTurno & ")"
        strSQL = strSQL & ")"
    #Else
        strSQL = "SELECT FolioApertura, SerieApertura, FondoApertura"
        strSQL = strSQL & " FROM TURNOS"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((FechaApertura)=#" & Format(dFecha, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((NumeroCaja)=" & lCaja & ")"
        strSQL = strSQL & " AND ((TurnoNo)=" & lTurno & ")"
        strSQL = strSQL & ")"
    #End If
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        sFolioIni = adorcs!FolioApertura
        sSerieIni = adorcs!SerieApertura
        dFondoIni = adorcs!FondoApertura
    Else
        DatosTurno = False
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    
End Function
