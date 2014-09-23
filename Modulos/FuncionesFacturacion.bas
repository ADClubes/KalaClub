Attribute VB_Name = "FuncionesFacturacion"
Option Explicit
Dim dMonto As Double
Dim dImporte As Double
Dim dDescuento As Double
Dim dIntereses As Double
Dim dIvaPor As Double
Dim dIva As Double
Dim dIvaDescuento As Double
Dim dIvaIntereses As Double

Public Sub RecalculaRenglon()

    dMonto = frmFacturacion.ssdbgFactura.Columns("Cantidad").Value * frmFacturacion.ssdbgFactura.Columns("Importe").Value
    dImporte = dMonto + frmFacturacion.ssdbgFactura.Columns("Intereses").Value
    dDescuento = dImporte * frmFacturacion.ssdbgFactura.Columns("Descuento").Value / 100
    dIntereses = frmFacturacion.ssdbgFactura.Columns("Intereses").Value
    dIvaPor = frmFacturacion.ssdbgFactura.Columns("IvaPor").Value
    dIva = dMonto - Round(dMonto / (1 + dIvaPor), 2)
    dIvaDescuento = dDescuento - Round(dDescuento / (1 + dIvaPor), 2)
    dIvaIntereses = dIntereses - Round(dIntereses / (1 + dIvaPor), 2)
    
    frmFacturacion.ssdbgFactura.Columns("Total").Value = dImporte - dDescuento
    frmFacturacion.ssdbgFactura.Columns("Iva").Value = dIva
    frmFacturacion.ssdbgFactura.Columns("IvaDescuento").Value = dIvaDescuento
    frmFacturacion.ssdbgFactura.Columns("IvaIntereses").Value = dIvaIntereses
    frmFacturacion.ssdbgFactura.Columns("DescMonto").Value = dDescuento
    frmFacturacion.ssdbgFactura.Columns("Total").Value = dImporte - dDescuento
    
    frmFacturacion.ssdbgFactura.Update
    
End Sub

Public Function GetFolio(lNumerodeFolios As Long, iModo As Integer) As Long
    Dim adorcsFolio As ADODB.Recordset
    Dim lNumFac As Long
    Dim i As Long
    Dim lInicia As Long
    Dim lDelay As Long
    Dim lVeces As Long
    Dim lContador As Long
    
    lDelay = 2
    lVeces = 4
    
    On Error Resume Next
    Conn.Errors.Clear
    Err.Clear
    
    If lNumerodeFolios = 0 Then
        GetFolio = 0
        Exit Function
    End If
    
    #If SqlServer_ Then
        Dim adocmdFolio As ADODB.Command
        Set adocmdFolio = New ADODB.Command
        adocmdFolio.ActiveConnection = Conn
        adocmdFolio.CommandType = adCmdStoredProc
        
        If iModo = 0 Then
            adocmdFolio.CommandText = "sp_GetFolioFactura"
        Else
            adocmdFolio.CommandText = "sp_GetFolioRecibo"
        End If
        
        With adocmdFolio
            .Parameters.Append .CreateParameter("@NumerodeFolios", adInteger, adParamInput, 8, lNumerodeFolios)
            .Parameters.Append .CreateParameter("@Modo", adInteger, adParamInput, 8, 0)
            .Parameters.Append .CreateParameter("@Folio", adInteger, adParamOutput, 8)
            .Execute
        End With
        
        GetFolio = adocmdFolio.Parameters("@Folio").Value
        
        Set adocmdFolio = Nothing
        
    #Else
        If iModo = 0 Then
            strSQL = "SELECT NumeroFactura"
        Else
            strSQL = "SELECT NumeroRecibo"
        End If
        
        strSQL = strSQL & " FROM FOLIO_FACTURA"
        
        Set adorcsFolio = New ADODB.Recordset
        adorcsFolio.CursorLocation = adUseServer
        adorcsFolio.CursorType = adOpenDynamic
        adorcsFolio.LockType = adLockPessimistic
        
        adorcsFolio.Open strSQL, Conn
        
        'ocurrio un error
        If Conn.Errors.Count Then
            GetFolio = -1
            Exit Function
        End If
        
        lContador = 0
        
        Do While True
        
            Conn.Errors.Clear
            Err.Clear
            
            If lContador >= lVeces Then
                GetFolio = -1
                Exit Function
            End If
            
            
            If iModo = 0 Then
                lNumFac = adorcsFolio!NumeroFactura + 1
                adorcsFolio!NumeroFactura = lNumFac + lNumerodeFolios - 1
            Else
                lNumFac = adorcsFolio!NumeroRecibo + 1
                adorcsFolio!NumeroRecibo = lNumFac + lNumerodeFolios - 1
            End If
            
            
            If Conn.Errors.Count Then
                If Conn.Errors.Item(0).SQLState = 3218 Then
                    lInicia = Timer
                    Do While Timer < lInicia + lDelay
                        DoEvents
                    Loop
                    lContador = lContador + 1
                End If
            Else
                Exit Do
            End If
        Loop
        
        
        adorcsFolio.Update
        
        adorcsFolio.Close
        Set adorcsFolio = Nothing
        
        On Error GoTo 0
        
        GetFolio = lNumFac
    #End If
    
End Function
Public Function GetFolioSerie(lNumerodeFolios As Long, sSerie As String) As Long
    Dim adorcsFolio As ADODB.Recordset
    Dim lNumFac As Long
    Dim i As Long
    Dim lInicia As Long
    Dim lDelay As Long
    Dim lVeces As Long
    Dim lContador As Long
    
    lDelay = 2
    lVeces = 4
    
    On Error Resume Next
    Conn.Errors.Clear
    Err.Clear
    
    If lNumerodeFolios = 0 Then
        GetFolioSerie = 0
        Exit Function
    End If
    
    #If SqlServer_ Then
        Dim adocmdFolio As ADODB.Command
        Set adocmdFolio = New ADODB.Command
        adocmdFolio.ActiveConnection = Conn
        adocmdFolio.CommandType = adCmdStoredProc
        adocmdFolio.CommandText = "sp_GetFolioFacturaSerie"
        
        With adocmdFolio
            .Parameters.Append .CreateParameter("@NumerodeFolios", adInteger, adParamInput, 8, lNumerodeFolios)
            .Parameters.Append .CreateParameter("@Serie", adVarChar, adParamInput, 8, sSerie)
            .Parameters.Append .CreateParameter("@Folio", adInteger, adParamOutput, 8)
            .Execute
        End With
        
        GetFolioSerie = adocmdFolio.Parameters("@Folio").Value
        
        Set adocmdFolio = Nothing
    #Else
        strSQL = "SELECT NumeroFactura"
        strSQL = strSQL & " FROM FOLIO_FACTURA_SERIE"
        
        If sSerie = vbNullString Then
            strSQL = strSQL & " WHERE SerieFactura Is Null"
        Else
            strSQL = strSQL & " WHERE SerieFactura=" & "'" & sSerie & "'"
        End If
        
        Set adorcsFolio = New ADODB.Recordset
        adorcsFolio.CursorLocation = adUseServer
        adorcsFolio.CursorType = adOpenDynamic
        adorcsFolio.LockType = adLockPessimistic
        
        adorcsFolio.Open strSQL, Conn
        
        'ocurrio un error
        If Conn.Errors.Count Then
            GetFolioSerie = -1
            Exit Function
        End If
        
        lContador = 0
        
        Do While True
        
            Conn.Errors.Clear
            Err.Clear
            
            If lContador >= lVeces Then
                GetFolioSerie = -1
                Exit Function
            End If
            
            lNumFac = adorcsFolio!NumeroFactura + 1
            adorcsFolio!NumeroFactura = lNumFac + lNumerodeFolios - 1
            
            
            
            If Conn.Errors.Count Then
                If Conn.Errors.Item(0).SQLState = 3218 Then
                    lInicia = Timer
                    Do While Timer < lInicia + lDelay
                        DoEvents
                    Loop
                    lContador = lContador + 1
                End If
            Else
                Exit Do
            End If
        Loop
        
        
        adorcsFolio.Update
        
        adorcsFolio.Close
        Set adorcsFolio = Nothing
        
        On Error GoTo 0
        
        GetFolioSerie = lNumFac
    #End If
    
End Function


Public Function UltimoDiaDelMes(dFecha As Date) As Date
    Dim iMes As Integer
    Dim lAno As Long
    Dim dFechaRet As Date
    
    iMes = Month(dFecha) + 1
    lAno = Year(dFecha)
    If iMes = 13 Then
        iMes = 1
        lAno = lAno + 1
    End If
    
    dFechaRet = CDate("01" & "/" & Trim(Str(iMes)) & "/" & Trim(Str(lAno)))
    
    UltimoDiaDelMes = dFechaRet - 1
    
End Function
Function UltimoDiaDelPeriodo(dFecha As Date, lPeriodo As Long, boPeriodoAbierto) As Date
    Dim dFechaRet As Date
    
    Dim lPeriodoAct As Integer
    
    
    If boPeriodoAbierto Then
        dFechaRet = DateAdd("m", lPeriodo, dFecha)
    Else
        lPeriodoAct = PeriodoActual(dFecha, lPeriodo)
        dFechaRet = UltimoDiaDelMes(CDate("01" & "/" & lPeriodoAct * lPeriodo & "/" & Year(dFecha)))
    End If
    
    UltimoDiaDelPeriodo = dFechaRet
    
End Function

Function PeriodoActual(dFecha As Date, lPeriodo As Long) As Integer
    
    Dim iPeriodo As Integer
    
    iPeriodo = Int(Month(dFecha) / lPeriodo)
    
    If (Month(dFecha) Mod lPeriodo) > 0 Then
        iPeriodo = iPeriodo + 1
    End If
    
    PeriodoActual = iPeriodo
    
End Function
'Actualiza las fechas de pago
Public Sub ActualizaFechas(lNumeroInicial As Long, lNumeroFinal As Long, iModo As Integer)
    Dim adorcsDatos As ADODB.Recordset
    Dim adorcsActualiza As ADODB.Recordset
    Dim adocmdActFac As ADODB.Command
    Dim sCancelaFac As String
    Dim lNumero As Long
    Dim cmdUpdate As Object
    
    Err.Clear
    Conn.Errors.Clear
    On Error GoTo Error_Catch
    
    Set adocmdActFac = New ADODB.Command
    adocmdActFac.ActiveConnection = Conn
    adocmdActFac.CommandType = adCmdText

    Set adorcsActualiza = New ADODB.Recordset
    adorcsActualiza.CursorLocation = adUseServer
    
    Set adorcsDatos = New ADODB.Recordset
    adorcsDatos.CursorLocation = adUseServer
    
    For lNumero = lNumeroInicial To lNumeroFinal
    
        sCancelaFac = ""
    
        'Para cuotas de mantenimiento
        strSQL = "SELECT DET.IdMember, DET.IdConcepto, Max(DET.Periodo) AS Fecha"
        If iModo = 0 Then
            strSQL = strSQL & " FROM FACTURAS_DETALLE DET"
            strSQL = strSQL & " WHERE DET.NumeroFactura=" & lNumero
        Else
            strSQL = strSQL & " FROM RECIBOS_DETALLE DET"
            strSQL = strSQL & " WHERE DET.NumeroRecibo=" & lNumero
        End If
        
        strSQL = strSQL & " AND DET.TipoCargo=0"
        strSQL = strSQL & " GROUP BY DET.IdMember, DET.IdConcepto"
        
    
        adorcsDatos.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
        Do Until adorcsDatos.EOF
            
            #If SqlServer_ Then
                strSQL = "SELECT CONVERT(varchar,FechaUltimoPago, 103) AS FechaUltimoPago"
                strSQL = strSQL & " FROM FECHAS_USUARIO"
                strSQL = strSQL & " WHERE IdMember=" & adorcsDatos!Idmember
                strSQL = strSQL & " AND IdConcepto=" & adorcsDatos!IdConcepto
            #Else
                strSQL = "SELECT FechaUltimoPago"
                strSQL = strSQL & " FROM FECHAS_USUARIO"
                strSQL = strSQL & " WHERE IdMember=" & adorcsDatos!Idmember
                strSQL = strSQL & " AND IdConcepto=" & adorcsDatos!IdConcepto
            #End If
                    
            adorcsActualiza.Open strSQL, Conn, adOpenDynamic, adLockOptimistic
        
            '                            1           10                                              6                                10
            sCancelaFac = sCancelaFac & "0" & Format(adorcsDatos!Idmember, "0000000000") & Format(adorcsDatos!IdConcepto, "000000") & adorcsActualiza!Fechaultimopago
            'Mide 27 caracteres
            
            #If SqlServer_ Then
                Set cmdUpdate = New ADODB.Command
                cmdUpdate.ActiveConnection = Conn
                cmdUpdate.CommandType = adCmdText
                
                strSQL = "UPDATE FECHAS_USUARIO SET Fechaultimopago = '" & Format(adorcsDatos!Fecha, "yyyymmdd") & "'"
                strSQL = strSQL & " WHERE IdMember=" & adorcsDatos!Idmember
                strSQL = strSQL & " AND IdConcepto=" & adorcsDatos!IdConcepto
                
                cmdUpdate.CommandText = strSQL
                cmdUpdate.Execute
            #Else
                adorcsActualiza!Fechaultimopago = adorcsDatos!Fecha
                adorcsActualiza.Update
            #End If
            
            adorcsActualiza.Close
        
            adorcsDatos.MoveNext
        Loop
    
        adorcsDatos.Close
    
        'Para Casilleros
        strSQL = "SELECT DET.IdMember, DET.Auxiliar, DET.Total, FAC.FechaFactura, Max(DET.Periodo) AS Fecha"
        If iModo = 0 Then
            strSQL = strSQL & " FROM FACTURAS_DETALLE DET INNER JOIN FACTURAS FAC ON DET.NumeroFactura=FAC.NumeroFactura"
            strSQL = strSQL & " WHERE DET.NumeroFactura=" & lNumero
        Else
            strSQL = strSQL & " FROM RECIBOS_DETALLE DET INNER JOIN RECIBOS FAC ON DET.NumeroRecibo=FAC.NumeroRecibo"
            strSQL = strSQL & " WHERE DET.NumeroRecibo=" & lNumero
        End If
        strSQL = strSQL & " AND DET.TipoCargo=1"
        strSQL = strSQL & " GROUP BY DET.IdMember, DET.Auxiliar, DET.Total, FAC.FechaFactura"
    
        adorcsDatos.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
        Do Until adorcsDatos.EOF
    
            strSQL = "SELECT FechaPago, FechaInicio, ImportePagado, Documento"
            strSQL = strSQL & " FROM RENTABLES"
            strSQL = strSQL & " WHERE IdUsuario=" & adorcsDatos!Idmember
            #If SqlServer_ Then
                strSQL = strSQL & " AND LTRIM(RTrim(Numero))='" & adorcsDatos!Auxiliar & "'"
            #Else
                strSQL = strSQL & " AND Trim(Numero)='" & adorcsDatos!Auxiliar & "'"
            #End If
        
            adorcsActualiza.Open strSQL, Conn, adOpenDynamic, adLockOptimistic
        
            '                            1           10                                           6                                 10
            sCancelaFac = sCancelaFac & "1" & Format(adorcsDatos!Idmember, "0000000000") & Format(adorcsDatos!Auxiliar, "@@@@@@") & Format(adorcsActualiza!Fechapago, "yyyymmdd")
            'Mide 27 caracteres
            
            #If SqlServer_ Then
                Set cmdUpdate = New ADODB.Command
                cmdUpdate.ActiveConnection = Conn
                cmdUpdate.CommandType = adCmdText
                
                strSQL = "UPDATE RENTABLES SET"
                strSQL = strSQL & " Fechapago = '" & Format(adorcsDatos!Fecha, "yyyymmdd") & "',"
                strSQL = strSQL & " FechaInicio='" & Format(adorcsDatos!FechaFactura, "yyyymmdd") & "',"
                strSQL = strSQL & " ImportePagado=" & adorcsDatos!Total & ","
                strSQL = strSQL & " Documento='" & IIf(iModo = 0, "F", "R") & lNumero & "'"
                strSQL = strSQL & " WHERE IdUsuario=" & adorcsDatos!Idmember
                strSQL = strSQL & " AND LTRIM(RTRIM(Numero))='" & adorcsDatos!Auxiliar & "'"
                
                cmdUpdate.CommandText = strSQL
                cmdUpdate.Execute
            #Else
                adorcsActualiza!Fechapago = adorcsDatos!Fecha
                adorcsActualiza!FechaInicio = adorcsDatos!FechaFactura
                adorcsActualiza!ImportePagado = adorcsDatos!Total
                adorcsActualiza!Documento = IIf(iModo = 0, "F", "R") & lNumero
                adorcsActualiza.Update
            #End If
            
            adorcsActualiza.Close
    
            adorcsDatos.MoveNext
        Loop
    
        adorcsDatos.Close
        
        'Para cargos varios
        strSQL = "SELECT DET.IdMember, DET.Auxiliar"
        If iModo = 0 Then
            strSQL = strSQL & " FROM FACTURAS_DETALLE DET"
            strSQL = strSQL & " WHERE DET.NumeroFactura=" & lNumero
        Else
            strSQL = strSQL & " FROM RECIBOS_DETALLE DET"
            strSQL = strSQL & " WHERE DET.NumeroRecibo=" & lNumero
        End If
        strSQL = strSQL & " AND DET.TipoCargo=2"
        strSQL = strSQL & " ORDER BY DET.Auxiliar"
    
        adorcsDatos.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        
        Do Until adorcsDatos.EOF
    
            strSQL = "SELECT Pagado"
            strSQL = strSQL & " FROM CARGOS_VARIOS"
            strSQL = strSQL & " WHERE IdMember=" & adorcsDatos!Idmember
            strSQL = strSQL & " AND IdCargoVario=" & Trim(adorcsDatos!Auxiliar)
        
        
            adorcsActualiza.Open strSQL, Conn, adOpenDynamic, adLockOptimistic
        
            '                            1           10                                           10
            sCancelaFac = sCancelaFac & "2" & Format(adorcsDatos!Idmember, "0000000000") & Format(Val(adorcsDatos!Auxiliar), "0000000000")
            'Mide 21 caracteres
            
            
            #If SqlServer_ Then
                Set cmdUpdate = New ADODB.Command
                cmdUpdate.ActiveConnection = Conn
                cmdUpdate.CommandType = adCmdText
                
                strSQL = "UPDATE CARGOS_VARIOS SET"
                strSQL = strSQL & " Pagado = 1"
                strSQL = strSQL & " WHERE IdMember=" & adorcsDatos!Idmember
                strSQL = strSQL & " AND IdCargoVario=" & Trim(adorcsDatos!Auxiliar)
                
                cmdUpdate.CommandText = strSQL
                cmdUpdate.Execute
            #Else
                adorcsActualiza!Pagado = -1
                adorcsActualiza.Update
            #End If
            
            adorcsActualiza.Close
            
            adorcsDatos.MoveNext
        Loop
    
        adorcsDatos.Close
        
        '----------------
        'Para Cargos por membresia
        strSQL = "SELECT DET.IdMember, DET.Auxiliar"
        If iModo = 0 Then
            strSQL = strSQL & " FROM FACTURAS_DETALLE DET"
            strSQL = strSQL & " WHERE DET.NumeroFactura=" & lNumero
        Else
            strSQL = strSQL & " FROM RECIBOS_DETALLE DET"
            strSQL = strSQL & " WHERE DET.NumeroRecibo=" & lNumero
        End If
        strSQL = strSQL & " AND DET.TipoCargo=3"
        strSQL = strSQL & " ORDER BY DET.Auxiliar"
    
        adorcsDatos.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        
        Do Until adorcsDatos.EOF
    
            strSQL = "SELECT DETALLE_MEM.FechaPago, DETALLE_MEM.Observaciones"
            strSQL = strSQL & " FROM DETALLE_MEM INNER JOIN MEMBRESIAS ON DETALLE_MEM.IdMembresia=MEMBRESIAS.IdMembresia"
            strSQL = strSQL & " WHERE Membresias.IdMember=" & adorcsDatos!Idmember
            strSQL = strSQL & " AND DETALLE_MEM.IdReg=" & Trim(adorcsDatos!Auxiliar)
        
        
            adorcsActualiza.Open strSQL, Conn, adOpenDynamic, adLockOptimistic
        
            '                            1           10                                           10
            sCancelaFac = sCancelaFac & "3" & Format(adorcsDatos!Idmember, "0000000000") & Format(Val(adorcsDatos!Auxiliar), "0000000000")
            'Mide 21 caracteres
            
            #If SqlServer_ Then
                Set cmdUpdate = New ADODB.Command
                cmdUpdate.ActiveConnection = Conn
                cmdUpdate.CommandType = adCmdText
                
                strSQL = "UPDATE DETALLE_MEM SET"
                strSQL = strSQL & " Fechapago = '" & Format(Date, "yyyymmdd") & "',"
                strSQL = strSQL & " Observaciones = '" & IIf(iModo, "R", "F") & lNumero & "'"
                strSQL = strSQL & " FROM DETALLE_MEM INNER JOIN MEMBRESIAS ON DETALLE_MEM.IdMembresia=MEMBRESIAS.IdMembresia"
                strSQL = strSQL & " WHERE Membresias.IdMember=" & adorcsDatos!Idmember
                strSQL = strSQL & " AND DETALLE_MEM.IdReg=" & Trim(adorcsDatos!Auxiliar)
                
                cmdUpdate.CommandText = strSQL
                cmdUpdate.Execute
            #Else
                adorcsActualiza!Fechapago = Format(Date, "dd/mm/yyyy")
                adorcsActualiza!Observaciones = IIf(iModo, "R", "F") & lNumero
                adorcsActualiza.Update
            #End If
            
            adorcsActualiza.Close
            
            adorcsDatos.MoveNext
        Loop
    
        adorcsDatos.Close
        
        '----------------
        If iModo = 0 Then
            strSQL = "INSERT INTO FACTURAS_CANCELA ("
            strSQL = strSQL & " NumeroFactura,"
        Else
            strSQL = "INSERT INTO RECIBOS_CANCELA ("
            strSQL = strSQL & " NumeroRecibo,"
        End If
        strSQL = strSQL & " CadenaCancela1,"
        strSQL = strSQL & " CadenaCancela2)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lNumero & ","
        strSQL = strSQL & "'" & Trim(Mid$(sCancelaFac, 1, 255)) & "',"
        strSQL = strSQL & "'" & Trim(Mid$(sCancelaFac, 256, 255)) & "')"
        
        
        adocmdActFac.CommandText = strSQL
        adocmdActFac.Execute
    
        
    Next
    
    Set adocmdActFac = Nothing
    
    Set adorcsDatos = Nothing
    Set adorcsActualiza = Nothing
    
    Exit Sub
    
Error_Catch:

    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "Error"
    End If
    
End Sub

Public Sub ActualizaDireccion(lidMember As Long)

    Dim adocmdDir As ADODB.Command
    Dim lIdDirNew As Long
    Dim lTipoDir As Long


    
    
    Conn.Errors.Clear
    Err.Clear
    On Error GoTo Error_Catch
    
    
    If Len(frmFacturacion.lblTipoDir.Caption) = 0 Then
        lTipoDir = 0
    Else
        lTipoDir = Val(frmFacturacion.lblTipoDir.Caption)
    End If
    
    
    Select Case lTipoDir
        Case 0 To 2
        
            lIdDirNew = LeeUltReg("DIRECCIONES", "IdDireccion") + 1
            
            strSQL = "INSERT INTO DIRECCIONES ("
            strSQL = strSQL & " IdDireccion,"
            strSQL = strSQL & " IdMember,"
            strSQL = strSQL & " Calle,"
            strSQL = strSQL & " Colonia,"
            strSQL = strSQL & " CodPos,"
            strSQL = strSQL & " IdTipoDireccion,"
            strSQL = strSQL & " RazonSocial,"
            strSQL = strSQL & " RFC,"
            strSQL = strSQL & " Estado,"
            strSQL = strSQL & " Ciudad,"
            strSQL = strSQL & " DelOMuni" & ")"
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & lIdDirNew & ","
            strSQL = strSQL & lidMember & ","
            strSQL = strSQL & "'" & Trim(frmFacturacion.txtFacDireccion.Text) & "',"
            strSQL = strSQL & "'" & Trim(frmFacturacion.txtFacColonia.Text) & "',"
            strSQL = strSQL & Val(frmFacturacion.txtFacCP.Text) & ","
            strSQL = strSQL & "3" & ","
            strSQL = strSQL & "'" & Trim(frmFacturacion.txtFacNombre.Text) & "',"
            strSQL = strSQL & "'" & Trim(frmFacturacion.txtFacRFC.Text) & "',"
            strSQL = strSQL & "'" & Trim(frmFacturacion.ssCmbFacEstado.Text) & "',"
            strSQL = strSQL & "'" & Trim(frmFacturacion.txtFacCiudad.Text) & "',"
            strSQL = strSQL & "'" & Trim(frmFacturacion.txtFacDelOMuni.Text) & "')"
        Case 3
            strSQL = "UPDATE DIRECCIONES SET "
            strSQL = strSQL & " Calle=" & "'" & Trim(frmFacturacion.txtFacDireccion.Text) & "',"
            strSQL = strSQL & " Colonia=" & "'" & Trim(frmFacturacion.txtFacColonia.Text) & "',"
            strSQL = strSQL & " CodPos=" & Val(frmFacturacion.txtFacCP.Text) & ","
            strSQL = strSQL & " IdTipoDireccion=" & "3" & ","
            strSQL = strSQL & " RazonSocial=" & "'" & Trim(frmFacturacion.txtFacNombre.Text) & "',"
            strSQL = strSQL & " RFC=" & "'" & Trim(frmFacturacion.txtFacRFC.Text) & "',"
            strSQL = strSQL & " Estado=" & "'" & Trim(frmFacturacion.ssCmbFacEstado.Text) & "',"
            strSQL = strSQL & " Ciudad=" & "'" & Trim(frmFacturacion.txtFacCiudad.Text) & "',"
            strSQL = strSQL & " DelOMuni=" & "'" & Trim(frmFacturacion.txtFacDelOMuni.Text) & "'"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " IdDireccion=" & frmFacturacion.lblIdDireccion
    End Select
    
    
    Set adocmdDir = New ADODB.Command
    
    adocmdDir.ActiveConnection = Conn
    adocmdDir.CommandType = adCmdText
    adocmdDir.CommandText = strSQL
    adocmdDir.Execute
    
    
    Set adocmdDir = Nothing
    Exit Sub
    
Error_Catch:

    If Error.Number <> 0 Then
    End If
    
End Sub


Public Sub ActualizaAcceso(lidMember As Long)
    Dim adocmdActAcc As ADODB.Command
    
    
    
    Set adocmdActAcc = New ADODB.Command
    adocmdActAcc.ActiveConnection = Conn
    adocmdActAcc.CommandType = adCmdText
    
    
    
    strSQL = "DELETE FROM ACCESO_DERECHOS"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " ACCESO_DERECHOS.IdMember IN"
    strSQL = strSQL & " (SELECT USUARIOS_CLUB.IdMember"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " USUARIOS_CLUB.IdTitular=" & lidMember & ")"
    
    adocmdActAcc.CommandText = strSQL
    adocmdActAcc.Execute
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO ACCESO_DERECHOS (IdMember, FechaAccesoPermitido)"
        strSQL = strSQL & " SELECT USUARIOS_CLUB.IdMember, dateadd(d,10,dateadd(m,1,Max(FECHAS_USUARIO.FechaUltimoPago))) AS FechaAcceso"
        strSQL = strSQL & " FROM USUARIOS_CLUB LEFT JOIN FECHAS_USUARIO ON USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember"
        strSQL = strSQL & " WHERE USUARIOS_CLUB.IdTitular=" & lidMember
        strSQL = strSQL & " GROUP BY USUARIOS_CLUB.IdMember"
    #Else
        strSQL = "INSERT INTO ACCESO_DERECHOS (IdMember, FechaAccesoPermitido)"
        strSQL = strSQL & " SELECT USUARIOS_CLUB.IdMember, dateadd('d',10,dateadd('m',1,Max(FECHAS_USUARIO.FechaUltimoPago))) AS FechaAcceso"
        strSQL = strSQL & " FROM USUARIOS_CLUB LEFT JOIN FECHAS_USUARIO ON USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember"
        strSQL = strSQL & " WHERE USUARIOS_CLUB.IdTitular=" & lidMember
        strSQL = strSQL & " GROUP BY USUARIOS_CLUB.IdMember"
    #End If
    
    adocmdActAcc.CommandText = strSQL
    adocmdActAcc.Execute
    
    #If SqlServer_ Then
        strSQL = "UPDATE ACCESO_DERECHOS"
        strSQL = strSQL & " SET FechaAccesoPermitido=AUSENCIAS.FechaInicial"
        strSQL = strSQL & " FROM"
        strSQL = strSQL & " (ACCESO_DERECHOS INNER JOIN AUSENCIAS ON ACCESO_DERECHOS.IdMember=AUSENCIAS.IdMember) INNER JOIN USUARIOS_CLUB ON ACCESO_DERECHOS.IdMember=USUARIOS_CLUB.Idmember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " (USUARIOS_CLUB.IdTitular=" & lidMember & ")"
        strSQL = strSQL & " AND (AUSENCIAS.FechaInicial <= '" & Format(Date, "yyyymmdd") & "')"
        strSQL = strSQL & " AND (AUSENCIAS.FechaFinal >= '" & Format(Date, "yyyymmdd") & "')"
    #Else
        strSQL = "UPDATE (ACCESO_DERECHOS INNER JOIN AUSENCIAS ON ACCESO_DERECHOS.IdMember=AUSENCIAS.IdMember) INNER JOIN USUARIOS_CLUB ON ACCESO_DERECHOS.IdMember=USUARIOS_CLUB.Idmember SET"
        strSQL = strSQL & " FechaAccesoPermitido=AUSENCIAS.FechaInicial"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " (USUARIOS_CLUB.IdTitular=" & lidMember & ")"
        strSQL = strSQL & " AND (AUSENCIAS.FechaInicial <= #" & Format(Date, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND (AUSENCIAS.FechaFinal >= #" & Format(Date, "mm/dd/yyyy") & "#)"
    #End If
    adocmdActAcc.CommandText = strSQL
    adocmdActAcc.Execute
    
    
    Set adocmdActAcc = Nothing
    
    
End Sub

Public Sub ActivaAcceso(lidMember As Long)
    
    Dim adorsActAcc As ADODB.Recordset
  Dim nErrCode As Long
    
    MDIPrincipal.StatusBar1.Panels(1).Text = "Activando Acceso..."
    
    #If SqlServer_ Then
        strSQL = "SELECT ACCESO_DERECHOS.IdMember, SECUENCIAL.Secuencial"
        strSQL = strSQL & " FROM (ACCESO_DERECHOS INNER JOIN SECUENCIAL"
        strSQL = strSQL & " ON ACCESO_DERECHOS.IdMember=SECUENCIAL.IdMember)"
        strSQL = strSQL & " INNER JOIN USUARIOS_CLUB"
        strSQL = strSQL & " ON ACCESO_DERECHOS.IdMember=USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " (USUARIOS_CLUB.IdTitular=" & lidMember & ")"
        strSQL = strSQL & " AND (ACCESO_DERECHOS.FechaAccesoPermitido >= '" & Format(Date, "yyyymmdd") & "')"
    #Else
        strSQL = "SELECT ACCESO_DERECHOS.IdMember, SECUENCIAL.Secuencial"
        strSQL = strSQL & " FROM (ACCESO_DERECHOS INNER JOIN SECUENCIAL"
        strSQL = strSQL & " ON ACCESO_DERECHOS.IdMember=SECUENCIAL.IdMember)"
        strSQL = strSQL & " INNER JOIN USUARIOS_CLUB"
        strSQL = strSQL & " ON ACCESO_DERECHOS.IdMember=USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " (USUARIOS_CLUB.IdTitular=" & lidMember & ")"
        strSQL = strSQL & " AND (ACCESO_DERECHOS.FechaAccesoPermitido >= #" & Format(Date, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ACCESO_DERECHOS.IdMember NOT IN (Select IdMember from AUSENCIAS)"
    #End If
    
    Set adorsActAcc = New ADODB.Recordset
    
    #If SqlServer_ Then
        adorsActAcc.CursorLocation = adUseClient
    #Else
        adorsActAcc.CursorLocation = adUseServer
    #End If
    
    adorsActAcc.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    'Set Conn = Nothing
    Do Until adorsActAcc.EOF
        #If SqlServer_ Then
            ActivaCredSQL 1, adorsActAcc!Secuencial, 1, adorsActAcc!Idmember, True, False
        #Else
            ActivaCred 1, adorsActAcc!Secuencial, 1, adorsActAcc!Idmember, True, False
        #End If
         'Registro en Torniquetes NITGEN
       
    
    'nErrCode = HabilitaAcceso(adorsActAcc!IdMember)
  
             'If nErrCode <> 0 Then
             '   MsgBox "No se pudo habilitar el usuario en torniquetes,Favor de hacerlo manual. "
             'End If
         
        
        adorsActAcc.MoveNext
    Loop
    'Connection_DB
    adorsActAcc.Close
    
    Set adorsActAcc = Nothing
    
    MDIPrincipal.StatusBar1.Panels(1).Text = ""
    
End Sub
Public Function CalculaPeriodos(dUltimoPago As Date, dFechacalculo, lPeriodo) As Long
    Dim dFechaCompara As Date
    Dim Periodos As Long
    
    
    dFechaCompara = dUltimoPago
    Periodos = 0
    
    
    If dUltimoPago >= dFechacalculo Then
        CalculaPeriodos = 0
        Exit Function
    End If
    
    'If lPeriodo < 12 Then
    Periodos = 1
    'End If
    
    Do While (dFechaCompara <= dFechacalculo)
        dFechaCompara = DateAdd("m", lPeriodo, dFechaCompara)
        
        If dFechaCompara < UltimoDiaDelMes(dFechaCompara) Then
            dFechaCompara = UltimoDiaDelMes(dFechaCompara)
        End If
        
        If dFechaCompara < dFechacalculo Then
            Periodos = Periodos + 1
        End If
    Loop
    
    CalculaPeriodos = Periodos

End Function
Public Function CalculaInteresesPasados(dPeriodo As Date, dPeriodoCalculo As Date, dMonto As Double, dPorcenInter) As Double

    Dim lMeses As Long
    
    lMeses = DateDiff("m", dPeriodo, dPeriodoCalculo)
    
    If lMeses <= 0 Then
        CalculaInteresesPasados = 0
        Exit Function
    End If
    
    
    
    CalculaInteresesPasados = Round(dMonto * dPorcenInter / 100, 2) * lMeses
    

End Function
Public Function CalculaInteresesActuales(dPeriodo As Date, dPeriodoCalculo As Date, dMonto As Double, dPorcenInter) As Double

    Dim lMeses As Long
    Dim lDiasGracia As Long
    
    
    
    
    lMeses = DateDiff("m", dPeriodo, dPeriodoCalculo)
    
    If lMeses < 0 Then
       CalculaInteresesActuales = 0
       Exit Function
    End If
    
    lDiasGracia = Val(ObtieneParametro("DIAS DE GRACIA"))
    
    If lDiasGracia = 0 Then
        lDiasGracia = 10
    End If
    
    Select Case lMeses
        Case 0 'es el mismo mes
            If Day(dPeriodoCalculo) <= lDiasGracia Then
                CalculaInteresesActuales = 0
                Exit Function
            End If
            CalculaInteresesActuales = Round(dMonto * dPorcenInter / 100, 2)
        Case Else
            CalculaInteresesActuales = Round(dMonto * dPorcenInter / 100, 2)
    End Select
End Function
Public Function Date2Days(ByRef dFecha) As Long
    Dim lRet
    
    
    lRet = Year(dFecha) * 365 + Month(dFecha) * 30 + Day(dFecha)
    
     Date2Days = lRet
    
End Function

Public Function ChecaDireccionado(ByRef lidMember As Long) As String

    Dim adorcsChkDir As ADODB.Recordset
    
    ChecaDireccionado = ""
    
    #If SqlServer_ Then
        strSQL = "SELECT TIPODIRECCIONADO"
        strSQL = strSQL & " FROM DIRECCIONADOS"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdMember=" & lidMember
        strSQL = strSQL & " AND Activo=1"
    #Else
        strSQL = "SELECT TIPODIRECCIONADO"
        strSQL = strSQL & " FROM DIRECCIONADOS"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdMember=" & lidMember
        strSQL = strSQL & " AND Activo=-1"
    #End If
    
    Set adorcsChkDir = New ADODB.Recordset
    adorcsChkDir.ActiveConnection = Conn
    adorcsChkDir.CursorLocation = adUseServer
    adorcsChkDir.CursorType = adOpenForwardOnly
    adorcsChkDir.LockType = adLockReadOnly
    adorcsChkDir.Open strSQL
    
    If Not adorcsChkDir.EOF Then
        ChecaDireccionado = adorcsChkDir!TipoDireccionado
    End If
    
    adorcsChkDir.Close
    
    Set adorcsChkDir = Nothing

End Function

Public Sub RestablecePagos()
    
    Dim lI As Long

    For lI = 0 To frmFacturacion.ssdbgPagos.Rows - 1
        frmFacturacion.ssdbgPagos.Bookmark = frmFacturacion.ssdbgPagos.AddItemBookmark(lI)
        frmFacturacion.ssdbgPagos.Columns("Resta").Value = frmFacturacion.ssdbgPagos.Columns("Importe").Value
    Next



End Sub

Public Function InsertaCargoVario(lidMember As Long, lIdConcepto As Long, sDescripcion As String, dImporte As Double, dFechaAlta As Date, dFechaVence As Date) As Boolean
    
    Dim adorcsInsCargo As ADODB.Recordset
    
    Dim adocmdInsCargo As ADODB.Command
    
    InsertaCargoVario = False
    
    Err.Clear
    Conn.Errors.Clear
    
    On Error GoTo Error_Catch
    
    
    If dImporte = -1 Then
    
        strSQL = "SELECT Monto"
        strSQL = strSQL & " FROM CONCEPTO_INGRESOS"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdConcepto=" & lIdConcepto
    
    
        Set adorcsInsCargo = New ADODB.Recordset
        adorcsInsCargo.CursorLocation = adUseServer
    
        adorcsInsCargo.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        
        If Not adorcsInsCargo.EOF Then
            dImporte = adorcsInsCargo!Monto
        End If
    
        adorcsInsCargo.Close
        Set adorcsInsCargo = Nothing
        
    End If
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO CARGOS_VARIOS ("
        strSQL = strSQL & " IdMember" & ","
        strSQL = strSQL & " IdConcepto" & ","
        strSQL = strSQL & " DescripcionCargo" & ","
        strSQL = strSQL & " Ordinal" & ","
        strSQL = strSQL & " NumeroDeCargos" & ","
        strSQL = strSQL & " FechaVencimiento" & ","
        strSQL = strSQL & " FechaAlta" & ","
        strSQL = strSQL & " Importe" & ")"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lidMember & ","
        strSQL = strSQL & lIdConcepto & ","
        strSQL = strSQL & "'" & Trim(sDescripcion) & "',"
        strSQL = strSQL & 1 & ","
        strSQL = strSQL & 1 & ","
        strSQL = strSQL & "'" & Format(dFechaAlta, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Format(dFechaVence, "yyyymmdd") & "',"
        strSQL = strSQL & dImporte & ")"
    #Else
        strSQL = "INSERT INTO CARGOS_VARIOS ("
        strSQL = strSQL & " IdMember" & ","
        strSQL = strSQL & " IdConcepto" & ","
        strSQL = strSQL & " DescripcionCargo" & ","
        strSQL = strSQL & " Ordinal" & ","
        strSQL = strSQL & " NumeroDeCargos" & ","
        strSQL = strSQL & " FechaVencimiento" & ","
        strSQL = strSQL & " FechaAlta" & ","
        strSQL = strSQL & " Importe" & ")"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lidMember & ","
        strSQL = strSQL & lIdConcepto & ","
        strSQL = strSQL & "'" & Trim(sDescripcion) & "',"
        strSQL = strSQL & 1 & ","
        strSQL = strSQL & 1 & ","
        strSQL = strSQL & "#" & Format(dFechaAlta, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "#" & Format(dFechaVence, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & dImporte & ")"
    #End If
    
    Set adocmdInsCargo = New ADODB.Command
    adocmdInsCargo.ActiveConnection = Conn
    adocmdInsCargo.CommandType = adCmdText
    adocmdInsCargo.CommandText = strSQL
    adocmdInsCargo.Execute
    
    Set adocmdInsCargo = Nothing
    
    InsertaCargoVario = True
    
    Exit Function
    
Error_Catch:

    
    
End Function


Function CalculaMantenimientoMes(lidMember As Long, boDesc As Boolean, iFormaP As Integer) As Double
    Dim adorsMantMes As ADODB.Recordset
    Dim dImporte As Double
    Dim dImporteInt As Double
    
    CalculaMantenimientoMes = 0
    
    strSQL = "SELECT USUARIOS_CLUB.IdMember, HISTORICO_CUOTAS.IdTipoUsuario, HISTORICO_CUOTAS.Monto, HISTORICO_CUOTAS.MontoDescuento, HISTORICO_CUOTAS.VigenteDesde, HISTORICO_CUOTAS.VigenteHasta, "
    strSQL = strSQL & "  AUSENCIAS.FechaInicial, AUSENCIAS.Porcentaje"
    strSQL = strSQL & " FROM (USUARIOS_CLUB INNER JOIN HISTORICO_CUOTAS ON USUARIOS_CLUB.IdTipoUsuario = HISTORICO_CUOTAS.IdTipoUsuario)"
    strSQL = strSQL & " LEFT JOIN AUSENCIAS ON USUARIOS_CLUB.IdMember = AUSENCIAS.IdMember"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "(((USUARIOS_CLUB.IdTitular)=" & lidMember & ")"
    strSQL = strSQL & " AND ((HISTORICO_CUOTAS.Periodo) =" & iFormaP & ")"
    
    #If SqlServer_ Then
        strSQL = strSQL & " AND ((HISTORICO_CUOTAS.VigenteDesde) <=" & "'" & Format(Date, "yyyymmdd") & "'" & ")"
        strSQL = strSQL & " AND ((HISTORICO_CUOTAS.VigenteHasta) >=" & "'" & Format(Date, "yyyymmdd") & "'" & "))"
    #Else
        strSQL = strSQL & " AND ((HISTORICO_CUOTAS.VigenteDesde) <=" & "#" & Format(Date, "mm/dd/yyyy") & "#" & ")"
        strSQL = strSQL & " AND ((HISTORICO_CUOTAS.VigenteHasta) >=" & "#" & Format(Date, "mm/dd/yyyy") & "#" & "))"
    #End If
    
    Set adorsMantMes = New ADODB.Recordset
    
    adorsMantMes.CursorLocation = adUseServer
    
    adorsMantMes.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do Until adorsMantMes.EOF
        dImporteInt = IIf(boDesc, adorsMantMes!MontoDescuento, adorsMantMes!Monto)
        
        If Not IsNull(adorsMantMes!FechaInicial) Then
            If CDate(adorsMantMes!FechaInicial) <= Date And Not IsNull(adorsMantMes!Porcentaje) Then
                dImporteInt = Round(dImporteInt * ((100 - adorsMantMes!Porcentaje) / 100), 2)
            End If
        End If
        dImporte = dImporte + dImporteInt
        adorsMantMes.MoveNext
    Loop
    
    adorsMantMes.Close
    Set adorsMantMes = Nothing
    
    CalculaMantenimientoMes = dImporte
    
End Function


Public Function ObtieneDatosConceptoIngresos(ByRef lIdConcepto As Long, ByRef sDescripcion As String, ByRef dMonto As Double, ByRef dIvaPor As Double, ByRef sFacORec As String, ByRef sUnidad As String) As Boolean
    
    Dim adorcs As ADODB.Recordset
    Dim sQuery As String
    
    ObtieneDatosConceptoIngresos = False
    
    sQuery = ""
    sQuery = sQuery + "SELECT Descripcion, Monto, Impuesto1, FacORec, Unidad"
    sQuery = sQuery + " FROM CONCEPTO_INGRESOS"
    sQuery = sQuery + " WHERE ("
    sQuery = sQuery + " (IdConcepto=" & lIdConcepto & ")"
    sQuery = sQuery + ")"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open sQuery, Conn, adOpenForwardOnly, adLockReadOnly
    
    If (Not adorcs.EOF) Then
        sDescripcion = Trim(adorcs!Descripcion)
        dMonto = adorcs!Monto
        dIvaPor = adorcs!Impuesto1
        sFacORec = adorcs!FacORec
        sUnidad = adorcs!Unidad
        ObtieneDatosConceptoIngresos = True
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    
    
End Function

Public Function GetTipoFactura(lNumeroFactura As Long) As Integer

    Dim adorcs As ADODB.Recordset
    
    GetTipoFactura = 0 'Es una factura normal
    
    
    strSQL = vbNullString
    strSQL = strSQL & "SELECT FACTURAS.Marca"
    strSQL = strSQL & " From FACTURAS"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((FACTURAS.NumeroFactura)=" & lNumeroFactura & ")"
    strSQL = strSQL & ")"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        GetTipoFactura = adorcs!Marca
    Else
        GetTipoFactura = -1
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing

End Function



Public Sub MuestraCFD(lNumeroFactura As Long, Optional sTipoDoc As String = "F")
   
    Dim sNombreArcCfd As String
    Dim sCFD As String
    
   
    
    sCFD = NombreArchivoCFD(lNumeroFactura, sTipoDoc, 0)
    
    sNombreArcCfd = ObtieneParametro("RUTA_CFD")
    
    sNombreArcCfd = sNombreArcCfd & "\" & sCFD & ".pdf"
    
    If AbreDoc(sNombreArcCfd) <> 0 Then
    End If

End Sub

Public Function NombreArchivoCFD(lNumeroFactura As Long, sTipoDoc As String, Optional nModo As Integer = 0) As String
    Dim adorcs As ADODB.Recordset
    
    NombreArchivoCFD = vbNullString
    
    If sTipoDoc = "F" Then
        strSQL = "SELECT FechaFactura, NombreArchivo"
        strSQL = strSQL & " FROM FACTURAS"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((NumeroFactura)=" & lNumeroFactura & ")"
        strSQL = strSQL & ")"
    Else
        strSQL = "SELECT FechaNota As FechaFactura, NombreArchivo"
        strSQL = strSQL & " FROM NOTAS_CRED"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((NumeroNota)=" & lNumeroFactura & ")"
        strSQL = strSQL & ")"
    End If
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If nModo = 0 Then
        If Not adorcs.EOF Then
            NombreArchivoCFD = Year(adorcs!FechaFactura)
            NombreArchivoCFD = NombreArchivoCFD & "\" & Format(adorcs!FechaFactura, "m")
            NombreArchivoCFD = NombreArchivoCFD & "\" & adorcs!NombreArchivo
        End If
    Else
        If Not adorcs.EOF Then
            NombreArchivoCFD = adorcs!NombreArchivo
        End If
    End If
    
    adorcs.Close
    Set adorcs = Nothing

    
End Function
Public Function FolioCFD(lNumeroFactura As Long, ByRef sFolio As String, ByRef sSerie As String) As Integer
    
    Dim adorcs As ADODB.Recordset
    
    FolioCFD = 0
    
    strSQL = "SELECT FolioCFD, SerieCFD"
    strSQL = strSQL & " FROM FACTURAS"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((NumeroFactura)=" & lNumeroFactura & ")"
    strSQL = strSQL & ")"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        sFolio = adorcs!FolioCFD
        sSerie = adorcs!SerieCFD
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
End Function

'--------------
Public Sub ObtieneDatosFactura(lidMember As Long, ByRef frmTarget As Form)

    Dim adoRcsDireccion As ADODB.Recordset
    'Busca direccion
    'Primero busca direccion fiscal
    
    strSQL = "SELECT RazonSocial, RFC, CALLE, COLONIA, DELOMUNI, Ciudad, CODPOS, TEL1, TEL2, Estado, IdDireccion, IdTipoDireccion, TipoPersona"
    strSQL = strSQL & " FROM DIRECCIONES DIR"
    strSQL = strSQL & " WHERE IDMEMBER=" & lidMember
    strSQL = strSQL & " AND IDTIPODIRECCION=3"
    
    Set adoRcsDireccion = New ADODB.Recordset
    adoRcsDireccion.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adoRcsDireccion.EOF Then
        
        frmTarget.txtFacNombre.Text = IIf(IsNull(adoRcsDireccion!RazonSocial), "", adoRcsDireccion!RazonSocial)
        frmTarget.txtFacRFC.Text = IIf(IsNull(adoRcsDireccion!rfc), vbNullString, adoRcsDireccion!rfc)
        frmTarget.txtFacDireccion = IIf(IsNull(adoRcsDireccion!calle), vbNullString, adoRcsDireccion!calle)
        frmTarget.txtFacColonia = IIf(IsNull(adoRcsDireccion!colonia), vbNullString, adoRcsDireccion!colonia)
        frmTarget.txtFacDelOMuni.Text = IIf(IsNull(adoRcsDireccion!DeloMuni), vbNullString, adoRcsDireccion!DeloMuni)
        frmTarget.txtFacCiudad.Text = IIf(IsNull(adoRcsDireccion!Ciudad), vbNullString, adoRcsDireccion!Ciudad)
        frmTarget.txtFacCP = IIf(IsNull(adoRcsDireccion!Codpos), vbNullString, Format(adoRcsDireccion!Codpos, "00000"))
        frmTarget.ssCmbFacEstado.Text = IIf(IsNull(adoRcsDireccion!Estado), vbNullString, adoRcsDireccion!Estado)
        frmTarget.txtFacTelefono.Text = IIf(IsNull(adoRcsDireccion!Tel1), vbNullString, adoRcsDireccion!Tel1)
        'lblIdDireccion = IIf(IsNull(adoRcsDireccion!IdDireccion), vbNullString, adoRcsDireccion!IdDireccion)
        'lblTipoDir = IIf(IsNull(adoRcsDireccion!IdTipoDireccion), vbNullString, adoRcsDireccion!IdTipoDireccion)
        If adoRcsDireccion!TipoPersona = "F" Then
            frmTarget.optTipoPer(0).Value = True
        Else
            frmTarget.optTipoPer(1).Value = True
        End If
    Else
        adoRcsDireccion.Close
        strSQL = "SELECT RazonSocial, RFC, CALLE, COLONIA, DELOMUNI, Ciudad, CODPOS, TEL1, TEL2, Estado, IdDireccion, IdTipoDireccion, TipoPersona"
        strSQL = strSQL & " FROM DIRECCIONES DIR"
        strSQL = strSQL & " WHERE IDMEMBER=" & lidMember
        strSQL = strSQL & " AND IDTIPODIRECCION=1"
        
        
        adoRcsDireccion.Open strSQL
        
        If Not adoRcsDireccion.EOF Then
            frmTarget.txtFacNombre.Text = IIf(IsNull(adoRcsDireccion!RazonSocial), "", adoRcsDireccion!RazonSocial)
            frmTarget.txtFacRFC.Text = IIf(IsNull(adoRcsDireccion!rfc), vbNullString, adoRcsDireccion!rfc)
            frmTarget.txtFacDireccion = IIf(IsNull(adoRcsDireccion!calle), vbNullString, adoRcsDireccion!calle)
            frmTarget.txtFacColonia = IIf(IsNull(adoRcsDireccion!colonia), vbNullString, adoRcsDireccion!colonia)
            frmTarget.txtFacDelOMuni.Text = IIf(IsNull(adoRcsDireccion!DeloMuni), vbNullString, adoRcsDireccion!DeloMuni)
            frmTarget.txtFacCiudad.Text = IIf(IsNull(adoRcsDireccion!Ciudad), vbNullString, adoRcsDireccion!Ciudad)
            frmTarget.txtFacCP = IIf(IsNull(adoRcsDireccion!Codpos), vbNullString, Format(adoRcsDireccion!Codpos, "00000"))
            frmTarget.ssCmbFacEstado.Text = IIf(IsNull(adoRcsDireccion!Estado), vbNullString, adoRcsDireccion!Estado)
            frmTarget.txtFacTelefono.Text = IIf(IsNull(adoRcsDireccion!Tel1), vbNullString, adoRcsDireccion!Tel1)
            'Me.lblIdDireccion.Caption = IIf(IsNull(adoRcsDireccion!IdDireccion), vbNullString, adoRcsDireccion!IdDireccion)
            'Me.lblTipoDir.Caption = IIf(IsNull(adoRcsDireccion!IdTipoDireccion), vbNullString, adoRcsDireccion!IdTipoDireccion)
            If adoRcsDireccion!TipoPersona = "F" Then
                frmTarget.optTipoPer(0).Value = True
            Else
                frmTarget.optTipoPer(1).Value = True
            End If
        Else
            frmTarget.txtFacNombre.Text = "" 'Me.lblNombreUsu
            frmTarget.txtFacDireccion = vbNullString
            frmTarget.txtFacColonia = vbNullString
            frmTarget.txtFacDelOMuni.Text = vbNullString
            frmTarget.txtFacCP = vbNullString
            frmTarget.ssCmbFacEstado.Text = vbNullString
            frmTarget.txtFacTelefono.Text = vbNullString
            'lblIdDireccion.Caption = vbNullString
            'lblTipoDir.Caption = vbNullString
            frmTarget.optTipoPer(0).Value = True
        End If
        
    End If
    
    adoRcsDireccion.Close
    Set adoRcsDireccion = Nothing
    
    
    If frmTarget.txtFacNombre.Text = vbNullString Then
        'frmTarget.txtFacNombre.Text = frmTarget.txtNombreInsc.Text
    End If
    
End Sub
'-------------


Public Function ConceptoModificable(lIdConcepto As Long) As Boolean
    Dim adorcs As ADODB.Recordset
    
    
    ConceptoModificable = False
    
    strSQL = "Select EsModificable as M"
    strSQL = strSQL & " From CONCEPTO_INGRESOS"
    strSQL = strSQL & " Where IdConcepto=" & lIdConcepto
    
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        If adorcs!M = 1 Then
            ConceptoModificable = True
        End If
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
    
End Function
Public Function ErrorFolioCFD(lNumeroFactura As Long) As String
Dim adorcs As ADODB.Recordset
strSQL = "Select Mensaje as M"
    strSQL = strSQL & " from [dbo].[CFDI_RESPUESTA]"
    strSQL = strSQL & " Where NombreCFDI='" & lNumeroFactura & "'"
    
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        ErrorFolioCFD = adorcs!M
        
    Else
    ErrorFolioCFD = ""
    
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
    
End Function

Public Function ActualizaFolioCFD(lNumeroFactura As Long, sFolioCFD As String, sSerieCFD As String, sModo As String) As Integer
    Dim adocmd As ADODB.Command
    Dim sNombreArcCfd As String
    Dim lRec As Long
    
    
    
    sNombreArcCfd = sNombreArcCfd & ObtieneParametro("NOMBRE_CFD")
    sNombreArcCfd = sNombreArcCfd & sSerieCFD & Trim(sFolioCFD)
    
    
    If sModo = "F" Then
        strSQL = "UPDATE FACTURAS SET "
    Else
        strSQL = "UPDATE NOTAS_CRED SET "
    End If
    
    strSQL = strSQL & " FolioCFD = " & "'" & sFolioCFD & "',"
    strSQL = strSQL & " SerieCFD = " & "'" & sSerieCFD & "',"
    strSQL = strSQL & " NombreArchivo = " & "'" & sNombreArcCfd & "'"
    strSQL = strSQL & " WHERE ("
    
    If sModo = "F" Then
        strSQL = strSQL & "(NumeroFactura)=" & lNumeroFactura
    Else
        strSQL = strSQL & "(NumeroNota)=" & lNumeroFactura
    End If
    
    strSQL = strSQL & ")"
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    adocmd.Execute lRec
    
    ActualizaFolioCFD = lRec
    
End Function
