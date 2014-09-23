Attribute VB_Name = "Cupones"
Option Explicit

Public Sub GeneraCupones(sTipoDoc As String, lNumeroDocIni As Long, lNumeroDocFin, ByRef lNumCupones As Long)
    
    Dim adoRcsCupon As ADODB.Recordset
    Dim lDiasVigencia As Integer
    Dim lNumDoc As Long
    
    
    If sTipoDoc = "R" Then
        strSQL = "SELECT RECIBOS.NumeroRecibo, RECIBOS_DETALLE.IdMember, RECIBOS_DETALLE.IdConcepto, RECIBOS_DETALLE.Periodo, RECIBOS_DETALLE.Total,RECIBOS_DETALLE.IdInstructor, RECIBOS_DETALLE.Auxiliar, RECIBOS_DETALLE.Cantidad, CONCEPTO_INGRESOS.NumeroCupones, CONCEPTO_INGRESOS.DiasVigenciaCupones"
        strSQL = strSQL & " FROM (RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo = RECIBOS.NumeroRecibo) INNER JOIN CONCEPTO_INGRESOS ON RECIBOS_DETALLE.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " ((RECIBOS.NumeroRecibo) BETWEEN " & lNumeroDocIni & " AND " & lNumeroDocFin
        strSQL = strSQL & " AND (CONCEPTO_INGRESOS.NumeroCupones) > 0)"
    Else
        strSQL = "SELECT Facturas.NumeroFactura, FACTURAS_DETALLE.IdMember, FACTURAS_DETALLE.IdConcepto, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.Total, FACTURAS_DETALLE.IdInstructor, FACTURAS_DETALLE.Auxiliar, FACTURAS_DETALLE.Cantidad, CONCEPTO_INGRESOS.NumeroCupones, CONCEPTO_INGRESOS.DiasVigenciaCupones"
        strSQL = strSQL & " FROM (FACTURAS_DETALLE INNER JOIN FACTURAS ON FACTURAS_DETALLE.NumeroFactura = FACTURAS.NumeroFactura) INNER JOIN CONCEPTO_INGRESOS ON FACTURAS_DETALLE.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " ((FACTURAS.NumeroFactura) BETWEEN " & lNumeroDocIni & " AND " & lNumeroDocFin
        strSQL = strSQL & " AND (CONCEPTO_INGRESOS.NumeroCupones) > 0)"
    End If
    
    Set adoRcsCupon = New ADODB.Recordset
    adoRcsCupon.ActiveConnection = Conn
    adoRcsCupon.CursorLocation = adUseServer
    adoRcsCupon.CursorType = adOpenForwardOnly
    adoRcsCupon.LockType = adLockReadOnly
    
    adoRcsCupon.Open strSQL
    
    Do Until adoRcsCupon.EOF
        
        lDiasVigencia = adoRcsCupon!DiasVigenciaCupones
            
        If lDiasVigencia = -1 Then
            lDiasVigencia = DateDiff("d", Date, adoRcsCupon!Periodo)
            If lDiasVigencia < 0 Then
                lDiasVigencia = 1
            End If
        End If
        
        If sTipoDoc = "R" Then
            lNumDoc = adoRcsCupon!NumeroRecibo
        Else
            lNumDoc = adoRcsCupon!NumeroFactura
        End If
        
        
        
        CreaCupones adoRcsCupon!Idmember, adoRcsCupon!IdConcepto, adoRcsCupon!IdInstructor, adoRcsCupon!Cantidad * adoRcsCupon!NumeroCupones, lDiasVigencia, sTipoDoc, lNumDoc, IIf(IsNull(adoRcsCupon!Auxiliar), vbNullString, adoRcsCupon!Auxiliar), adoRcsCupon!Total, lNumCupones
        
        adoRcsCupon.MoveNext
    Loop
    
    adoRcsCupon.Close
    Set adoRcsCupon = Nothing
    
    
End Sub
Public Sub CreaCupones(lidMember As Long, lIdConcepto As Long, lIdInstructor, lNumeroCupones As Integer, lDiasVigencia As Integer, sTipoDoc As String, lNumeroDoc As Long, sDatosAd As String, dImporteCupon As Double, ByRef lNumCupones As Long)
    Dim adoCmdCupon As ADODB.Command
    Dim adoRcsCupon As ADODB.Recordset
    
    Dim lNumeroCupon As Long
    
    Set adoCmdCupon = New ADODB.Command
    adoCmdCupon.ActiveConnection = Conn
    adoCmdCupon.CommandType = adCmdText
    
    
    For lNumeroCupon = 1 To lNumeroCupones
        strSQL = "INSERT INTO CUPONES ("
        strSQL = strSQL & " IdMember,"
        strSQL = strSQL & " IdConcepto,"
        '19/07/07
        strSQL = strSQL & " ImporteCupon,"
        strSQL = strSQL & " IdInstructor,"
        strSQL = strSQL & " NumeroCupon,"
        strSQL = strSQL & " TotalCupones,"
        strSQL = strSQL & " TipoDocumento,"
        strSQL = strSQL & " NumeroDocumento,"
        strSQL = strSQL & " FechaAlta,"
        strSQL = strSQL & " FechaVigencia,"
        strSQL = strSQL & " DatosAdicionales,"
        strSQL = strSQL & " Usuario)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lidMember & ","
        strSQL = strSQL & lIdConcepto & ","
        '19/07/07
        strSQL = strSQL & dImporteCupon & ","
        strSQL = strSQL & lIdInstructor & ","
        strSQL = strSQL & lNumeroCupon & ","
        strSQL = strSQL & lNumeroCupones & ","
        strSQL = strSQL & "'" & sTipoDoc & "'" & ","
        strSQL = strSQL & lNumeroDoc & ","
        #If SqlServer_ Then
            strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "'" & ","
            strSQL = strSQL & "'" & Format(Date + lDiasVigencia, "yyyymmdd") & "'" & ","
        #Else
            strSQL = strSQL & "#" & Format(Date, "mm/dd/yyyy") & "#" & ","
            strSQL = strSQL & "#" & Format(Date + lDiasVigencia, "mm/dd/yyyy") & "#" & ","
        #End If
        strSQL = strSQL & "'" & sDatosAd & "'" & ","
        strSQL = strSQL & "'" & sDB_User & "'" & ")"
        
        adoCmdCupon.CommandText = strSQL
        adoCmdCupon.Execute
        
        lNumCupones = lNumCupones + 1
        
    Next
    
    Set adoCmdCupon = Nothing
    
End Sub

