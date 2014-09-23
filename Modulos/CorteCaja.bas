Attribute VB_Name = "CorteCaja"
Public Sub ReportesCorte(iModo As Integer, lIdCorte As Long)
    Dim adocmdCorte As ADODB.Command
    Dim adorcsCorte As ADODB.Recordset
    Dim adorcsCompara As ADODB.Recordset
    
    
    Dim sIdent As String
       
    
    Dim frmRep As frmReportViewer
    
    Dim sCuenta As String
    
    sCuenta = sCuenta & ObtieneParametro("CUENTA_DEPOSITO")
    
    sIdent = Format(Now, "ddmmHhNnSs")
    
    
    Set adocmdCorte = New ADODB.Command
    adocmdCorte.ActiveConnection = Conn
    adocmdCorte.CommandType = adCmdText
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO TMP_CORTE_CAJA ("
        strSQL = strSQL & " IdCorteCaja,"
        strSQL = strSQL & " IdFormaPago,"
        strSQL = strSQL & " IdTPV,"
        strSQL = strSQL & " LoteNumero,"
        strSQL = strSQL & " ImporteOperado,"
        strSQL = strSQL & " ImporteCorte,"
        strSQL = strSQL & " Identificador)"
        strSQL = strSQL & " SELECT FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, ISNULL(CT_AFILIACIONES.IdTerminal, 0) AS IdTPV, PAGOS_FACTURA.LoteNumero, SUM(PAGOS_FACTURA.Importe) AS ImporteOperado, 0 AS ImporteCorte, '" & sIdent & "' AS Ident"
        strSQL = strSQL & " FROM PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura LEFT JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " FACTURAS.IdCorteCaja=" & lIdCorte
        strSQL = strSQL & " AND FACTURAS.Cancelada=0"
        strSQL = strSQL & " GROUP BY FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, CT_AFILIACIONES.IdTerminal, PAGOS_FACTURA.LoteNumero"
    #Else
        strSQL = "INSERT INTO TMP_CORTE_CAJA ("
        strSQL = strSQL & " IdCorteCaja,"
        strSQL = strSQL & " IdFormaPago,"
        strSQL = strSQL & " IdTPV,"
        strSQL = strSQL & " LoteNumero,"
        strSQL = strSQL & " ImporteOperado,"
        strSQL = strSQL & " ImporteCorte,"
        strSQL = strSQL & " Identificador)"
        strSQL = strSQL & " SELECT FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, iif(isnull(CT_AFILIACIONES.IdTerminal),0,CT_AFILIACIONES.IdTerminal) AS IdTPV, PAGOS_FACTURA.LoteNumero, Sum(PAGOS_FACTURA.Importe) AS ImporteOperado, 0 AS ImporteCorte, '" & sIdent & "' AS Ident"
        strSQL = strSQL & " FROM (PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura) LEFT JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FACTURAS.IdCorteCaja)=" & lIdCorte & ")"
        strSQL = strSQL & " AND ((FACTURAS.Cancelada)=0" & ")"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, CT_AFILIACIONES.IdTerminal, PAGOS_FACTURA.LoteNumero"
    #End If
    
    adocmdCorte.CommandText = strSQL
    adocmdCorte.Execute
    
    
    
    
    strSQL = "SELECT CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, CORTE_CAJA_DETALLE.LoteNumero, Sum(CORTE_CAJA_DETALLE.Importe) AS Importe"
    strSQL = strSQL & " FROM CORTE_CAJA_DETALLE INNER JOIN CORTE_CAJA ON CORTE_CAJA_DETALLE.IdCorteCaja = CORTE_CAJA.IdCorteCaja"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((CORTE_CAJA.IdCorteCaja)=" & lIdCorte & "))"
    strSQL = strSQL & " GROUP BY CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, CORTE_CAJA_DETALLE.LoteNumero"
    
    
    Set adorcsCorte = New ADODB.Recordset
    adorcsCorte.CursorLocation = adUseServer
    
    adorcsCorte.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Set adorcsCompara = New ADODB.Recordset
    adorcsCompara.CursorLocation = adUseServer
    
    Dim cmdUpdate As Object
    
    Do While Not adorcsCorte.EOF
        
        strSQL = "SELECT ImporteCorte"
        strSQL = strSQL & " FROM TMP_CORTE_CAJA"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((IdFormaPago)=" & adorcsCorte!IdFormaPago & ")"
        strSQL = strSQL & " AND ((IdTPV)=" & adorcsCorte!IdTPV & ")"
        strSQL = strSQL & " AND ((LoteNumero)='" & adorcsCorte!LoteNumero & "')"
        strSQL = strSQL & " AND ((Identificador)='" & sIdent & "')"
        strSQL = strSQL & ")"
        
        adorcsCompara.Open strSQL, Conn, adOpenForwardOnly, adLockOptimistic
        
        
        If Not adorcsCompara.EOF Then
            #If SqlServer_ Then
                Set cmdUpdate = New ADODB.Command
                cmdUpdate.ActiveConnection = Conn
                cmdUpdate.CommandType = adCmdText
                
                strSQL = "UPDATE TMP_CORTE_CAJA SET ImporteCorte = " & adorcsCorte!Importe
                strSQL = strSQL & " WHERE "
                strSQL = strSQL & " IdFormaPago=" & adorcsCorte!IdFormaPago
                strSQL = strSQL & " AND IdTPV=" & adorcsCorte!IdTPV
                strSQL = strSQL & " AND LoteNumero='" & adorcsCorte!LoteNumero & "'"
                strSQL = strSQL & " AND Identificador='" & sIdent & "'"
                
                cmdUpdate.CommandText = strSQL
                cmdUpdate.Execute
            #Else
            
                adorcsCompara!ImporteCorte = adorcsCorte!Importe
                adorcsCompara.Update
            #End If
        Else
            strSQL = "INSERT INTO TMP_CORTE_CAJA ("
            strSQL = strSQL & " IdCorteCaja,"
            strSQL = strSQL & " IdFormaPago,"
            strSQL = strSQL & " IdTPV,"
            strSQL = strSQL & " LoteNumero,"
            strSQL = strSQL & " ImporteCorte,"
            strSQL = strSQL & " Identificador)"
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & lIdCorte & ","
            strSQL = strSQL & adorcsCorte!IdFormaPago & ","
            strSQL = strSQL & adorcsCorte!IdTPV & ","
            strSQL = strSQL & "'" & adorcsCorte!LoteNumero & "',"
            strSQL = strSQL & adorcsCorte!Importe & ","
            strSQL = strSQL & "'" & sIdent & "')"
            
            
            adocmdCorte.CommandText = strSQL
            adocmdCorte.Execute
            
        End If
        
        adorcsCompara.Close
        
        adorcsCorte.MoveNext
    Loop
    
    
    Set adorcsCompara = Nothing
    
    adorcsCorte.Close
    Set adorcsCorte = Nothing
    
    
    
   
    
    
    strSQL = "SELECT TMP_CORTE_CAJA.IdCorteCaja, CORTE_CAJA.FechaCorte, CORTE_CAJA.HoraCorte, CORTE_CAJA.UsuarioCorte, CORTE_CAJA.Caja, CORTE_CAJA.Turno, CORTE_CAJA.FolioInicial, CORTE_CAJA.SerieInicial, CORTE_CAJA.FolioFinal, CORTE_CAJA.SerieFinal,CORTE_CAJA.FondoInicial, CORTE_CAJA.FondoDejado,CORTE_CAJA.NoFoliosUsados, TMP_CORTE_CAJA.IdFormaPago, FORMA_PAGO.Descripcion, TMP_CORTE_CAJA.IdTPV, CT_TPVS.DescripcionTPV, TMP_CORTE_CAJA.LoteNumero, TMP_CORTE_CAJA.ImporteOperado, TMP_CORTE_CAJA.ImporteCorte," & "CORTE_CAJA.Referencia," & " '" & sCuenta & "' AS Cuenta"
    strSQL = strSQL & " FROM ((TMP_CORTE_CAJA INNER JOIN CORTE_CAJA ON TMP_CORTE_CAJA.IdCorteCaja = CORTE_CAJA.IdCorteCaja) INNER JOIN FORMA_PAGO ON TMP_CORTE_CAJA.IdFormaPago = FORMA_PAGO.IdFormaPago) LEFT JOIN CT_TPVS ON TMP_CORTE_CAJA.IdTPV = CT_TPVS.IdTPV"
    strSQL = strSQL & " WHERE (((TMP_CORTE_CAJA.Identificador) = '" & sIdent & "'))"
    strSQL = strSQL & " ORDER BY TMP_CORTE_CAJA.IdFormaPago, TMP_CORTE_CAJA.IdTPV, TMP_CORTE_CAJA.LoteNumero;"

    
    Set frmRep = New frmReportViewer
    
    
    frmRep.sQuery = strSQL
    
    If iModo = 0 Then
        frmRep.sNombreReporte = sDB_ReportSource & "/" & "cc_cajero.rpt"
    Else
        frmRep.sNombreReporte = sDB_ReportSource & "/" & "cc_supervisor.rpt"
    End If
    frmRep.Show vbModal
    
    #If SqlServer_ Then
        strSQL = "DELETE FROM TMP_CORTE_CAJA"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " TMP_CORTE_CAJA.Identificador='" & sIdent & "'"
    #Else
        strSQL = "DELETE * FROM TMP_CORTE_CAJA"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((TMP_CORTE_CAJA.Identificador)='" & sIdent & "')"
        strSQL = strSQL & ")"
    #End If
    
    adocmdCorte.CommandText = strSQL
    adocmdCorte.Execute
    
    Set adocmdCorte = Nothing
End Sub
