VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmGeneraAdeudos 
   Caption         =   "Generacion de adeudos"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pbAvance 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Proceder"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpFechaCalc 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   50528257
      CurrentDate     =   38679
   End
   Begin VB.Label lblTermino 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblInicio 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblCount 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
End
Attribute VB_Name = "FrmGeneraAdeudos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sCargos(1000) As String 'Arreglo para facturar
Dim lIndexsCargos As Long   'Indice del arreglo de facturación

Private Sub cmdOk_Click()
    Dim adocmdAdeudos As ADODB.Command
    Dim adorcsAdeudos As ADODB.Recordset
    
    Dim lCount As Long
    
    #If SqlServer_ Then
        strSQL = "DELETE FROM ADEUDOS"
    #Else
        strSQL = "DELETE * FROM ADEUDOS"
    #End If
    
    Set adocmdAdeudos = New ADODB.Command
    adocmdAdeudos.ActiveConnection = Conn
    adocmdAdeudos.CommandType = adCmdText
    adocmdAdeudos.CommandText = strSQL
    adocmdAdeudos.Execute
    
    Set adocmdAdeudos = Nothing
    
    
    strSQL = "SELECT IdMember"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " WHERE IdTitular=IdMember"
    strSQL = strSQL & " And Status Is Null"
    strSQL = strSQL & " ORDER BY IdMember"
    
    
    Set adorcsAdeudos = New ADODB.Recordset
    adorcsAdeudos.CursorLocation = adUseServer
    
    adorcsAdeudos.Open strSQL, Conn, adOpenKeyset, adLockReadOnly
    
    pbAvance.Max = adorcsAdeudos.RecordCount
    
    
    Me.lblInicio.Caption = Now()
    Me.lblCount.Caption = lCount & " de " & Me.pbAvance.Max & " - " & Format(lCount / Me.pbAvance.Max, "#0.00%")
    DoEvents
    Do Until adorcsAdeudos.EOF
        CalculaMantenimientoAdeudos adorcsAdeudos!Idmember, 1
        adorcsAdeudos.MoveNext
        pbAvance.Value = pbAvance.Value + 1
        lCount = lCount + 1
        Me.lblCount.Caption = lCount & " de " & Me.pbAvance.Max & " - " & Format(lCount / Me.pbAvance.Max, "#0.00%")
        DoEvents
    Loop
    
    
    adorcsAdeudos.Close
    Set adorcsAdeudos = Nothing
    
    Me.lblTermino.Caption = Now()
    
End Sub


Private Sub CalculaMantenimientoAdeudos(ByRef lidMember As Long, ByVal iPeriodoCalculo As Integer)
    
    Dim AdoRcsUsuarios As ADODB.Recordset
    Dim adoRcsFac As ADODB.Recordset
    Dim adorcsAus As ADODB.Recordset
    
    Dim i As Long   'Variable para ciclos for
    
    
    
    Dim iUltimoDiadelMes As Integer
    Dim siProporcion As Single
    Dim iDiasProporcion As Integer
    
    
    Dim dFechaUltimoPago As Date
    Dim dPeriodo As Date
    Dim dPeriodoIni As Date
    
    Dim lPeriodos As Long
    
    Dim dCuota As Double
    Dim dCantidad As Double
    Dim dDescuentoPor As Double
    Dim dIvaPor As Double
    Dim lFormaPago As Long
    Dim dMonto As Double
    
    
    Dim dIntereses As Double
    Dim dDescuento As Double
                
    Dim dIva As Double
    Dim dIvaDescuento As Double
    Dim dIvaIntereses As Double
                
    Dim dImporte As Double
                
                
    Dim dTotal As Double
    
    
    
    Dim nPorAus As Integer
    
    
    Dim boDireccionar As Boolean
        
        
    Dim adocmdCalcMant As ADODB.Command
        
    boDireccionar = False
    
    
    
    Set adocmdCalcMant = New ADODB.Command
    adocmdCalcMant.ActiveConnection = Conn
    adocmdCalcMant.CommandType = adCmdText
        
    strSQL = "SELECT CONCEPTO_TIPO.IdConcepto, CONCEPTO_TIPO.Periodo, CONCEPTO_INGRESOS.Descripcion, CONCEPTO_INGRESOS.Impuesto1, CONCEPTO_INGRESOS.Impuesto2, CONCEPTO_INGRESOS.FacORec, FECHAS_USUARIO.FechaUltimoPago, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.IdMember, USUARIOS_CLUB.NumeroFamiliar, USUARIOS_CLUB.IdTipoUsuario, TIPO_USUARIO.Descripcion "
    strSQL = strSQL & " FROM (((USUARIOS_CLUB INNER JOIN CONCEPTO_TIPO ON USUARIOS_CLUB.IdTipoUsuario = CONCEPTO_TIPO.IdTipoUsuario) INNER JOIN CONCEPTO_INGRESOS ON CONCEPTO_TIPO.IdConcepto = CONCEPTO_INGRESOS.IdConcepto) INNER JOIN FECHAS_USUARIO ON (USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember) AND (CONCEPTO_TIPO.IdConcepto = FECHAS_USUARIO.IdConcepto)) LEFT JOIN TIPO_USUARIO ON USUARIOS_CLUB.IdTipoUsuario=TIPO_USUARIO.IdTipoUsuario"
    strSQL = strSQL & " WHERE USUARIOS_CLUB.IdTitular=" & lidMember
    Set AdoRcsUsuarios = New ADODB.Recordset
        
    AdoRcsUsuarios.ActiveConnection = Conn
    AdoRcsUsuarios.CursorLocation = adUseServer
    AdoRcsUsuarios.CursorType = adOpenForwardOnly
    AdoRcsUsuarios.LockType = adLockReadOnly
    AdoRcsUsuarios.Open strSQL

    
    If ChecaDireccionado(lidMember) <> "" Then
        boDireccionar = True
    End If


        
    'Crea e inicializa adoRcsFac
    Set adoRcsFac = New ADODB.Recordset
    adoRcsFac.ActiveConnection = Conn
    adoRcsFac.CursorLocation = adUseServer
    adoRcsFac.CursorType = adOpenForwardOnly
    adoRcsFac.LockType = adLockReadOnly
    
    'Crea e inicializa adorcsAuse
    Set adorcsAus = New ADODB.Recordset
    adorcsAus.ActiveConnection = Conn
    adorcsAus.CursorLocation = adUseServer
    adorcsAus.CursorType = adOpenForwardOnly
    adorcsAus.LockType = adLockReadOnly

    Do While Not AdoRcsUsuarios.EOF
    
        siProporcion = 1
        
        dFechaUltimoPago = AdoRcsUsuarios!Fechaultimopago
        
        iUltimoDiadelMes = Day(UltimoDiaDelMes(dFechaUltimoPago))
        iDiasProporcion = iUltimoDiadelMes - Day(dFechaUltimoPago) + 1
        
        'Para el pago proporcional de los dias
        If iDiasProporcion > 1 Then
            siProporcion = (1 / iUltimoDiadelMes) * iDiasProporcion
        End If
        
        dPeriodo = dFechaUltimoPago
        dPeriodoIni = dPeriodo
        
        'Si la ultima fecha de pago no coincide con el el
        'ultimo dia del mes (siProporcion < 1 ) ajusta la fecha
        'para calcular el periodo completo y despues multiplica la cuota
        'por el factor de proporcion.
        
        If siProporcion < 1 Then
            dPeriodo = UltimoDiaDelMes(DateAdd("m", -1, dPeriodo))
            dPeriodoIni = dPeriodo
            
            dFechaUltimoPago = dPeriodo
            
            
        End If
        
        
        
        
        'lPeriodos = CalculaPeriodos(dFechaUltimopago, Me.dtpFechaCalc.Value, AdoRcsUsuarios!Periodo)
        lPeriodos = CalculaPeriodos(dFechaUltimoPago, Me.dtpFechaCalc.Value, iPeriodoCalculo)
        
            
        For i = 1 To lPeriodos
        
            If i > 1 Then
                siProporcion = 1
            End If
            
            dPeriodoIni = CDate(dPeriodo + 1)
            'dPeriodo = DateAdd("m", i * AdoRcsUsuarios!Periodo, dFechaUltimopago)
            dPeriodo = DateAdd("m", i * iPeriodoCalculo, dFechaUltimoPago)
            dPeriodo = UltimoDiaDelMes(dPeriodo)
            
            #If SqlServer_ Then
                strSQL = "SELECT MONTO, MONTODESCUENTO"
                strSQL = strSQL & " FROM HISTORICO_CUOTAS"
                strSQL = strSQL & " WHERE IdTipoUsuario=" & AdoRcsUsuarios!idtipousuario
                'strSQL = strSQL & " AND Periodo=" & AdoRcsUsuarios!Periodo
                strSQL = strSQL & " AND Periodo=" & iPeriodoCalculo
                strSQL = strSQL & " AND VigenteDesde <=" & "'" & Format(dPeriodo, "yyyymmdd") & "'"
                strSQL = strSQL & " AND VigenteHasta >=" & "'" & Format(dPeriodo, "yyyymmdd") & "'"
            #Else
                strSQL = "SELECT MONTO, MONTODESCUENTO"
                strSQL = strSQL & " FROM HISTORICO_CUOTAS"
                strSQL = strSQL & " WHERE IdTipoUsuario=" & AdoRcsUsuarios!idtipousuario
                'strSQL = strSQL & " AND Periodo=" & AdoRcsUsuarios!Periodo
                strSQL = strSQL & " AND Periodo=" & iPeriodoCalculo
                strSQL = strSQL & " AND VigenteDesde <=" & "#" & dPeriodo & "#"
                strSQL = strSQL & " AND VigenteHasta >=" & "#" & dPeriodo & "#"
            #End If
            
            adoRcsFac.Open strSQL
                
            If adoRcsFac.EOF Then
                dCuota = 0
                dCantidad = 1
                dDescuentoPor = 0
                dIvaPor = 0
                lFormaPago = 0
            Else
                If boDireccionar Then
                    dCuota = adoRcsFac!MontoDescuento
                Else
                    dCuota = adoRcsFac!Monto
                End If
                dCantidad = 1
                dDescuentoPor = 0
                dIvaPor = AdoRcsUsuarios!Impuesto1 / 100
                lFormaPago = AdoRcsUsuarios!Periodo
            End If
                
            adoRcsFac.Close
            
            'Para ausencias
            #If SqlServer_ Then
                strSQL = "SELECT FechaInicial, FechaFinal, Porcentaje"
                strSQL = strSQL & " FROM AUSENCIAS"
                strSQL = strSQL & " WHERE IdMember=" & AdoRcsUsuarios!Idmember
                strSQL = strSQL & " AND FechaInicial <='" & Format(dPeriodoIni, "yyyymmdd") & "'"
            #Else
                strSQL = "SELECT FechaInicial, FechaFinal, Porcentaje"
                strSQL = strSQL & " FROM AUSENCIAS"
                strSQL = strSQL & " WHERE IdMember=" & AdoRcsUsuarios!Idmember
                strSQL = strSQL & " AND FechaInicial <=#" & Format(dPeriodoIni, "mm/dd/yyyy") & "#"
                'strSQL = strSQL & " AND FechaFinal >=#" & Format(dPeriodoIni, "mm/dd/yyyy") & "#"
            #End If
            
            adorcsAus.Open strSQL
            
            If adorcsAus.EOF Then
                nPorAus = 0
            Else
                nPorAus = adorcsAus!Porcentaje / 100
            End If
            
            adorcsAus.Close
                
            dMonto = dCuota * dCantidad * (1 - nPorAus)
            
            'Toma en cuenta el factor de proporcion
            
            dMonto = Round(dMonto * siProporcion, 2)
            
            dIntereses = CalculaInteresesPasados(dPeriodo, Me.dtpFechaCalc.Value, dMonto, 3)
            dIntereses = dIntereses + CalculaInteresesActuales(dPeriodo, Me.dtpFechaCalc.Value, dMonto, 3)
            dDescuento = 0
                
            dIva = dMonto - Round(dMonto / (1 + dIvaPor), 2)
            dIvaDescuento = dDescuento - Round(dDescuento / (1 + dIvaPor), 2)
            dIvaIntereses = dIntereses - Round(dIntereses / (1 + dIvaPor), 2)
                
            dImporte = dMonto + dIntereses
                
                
            dTotal = dImporte - dDescuento
                
                
                
                
            'Columnas del grid
            '0  Concepto
            '1  Nombre
            '2  Periodo
            '3  Cantidad
            '4  Importe
            '5  Intereses
            '6  Descuento
            '7  Total
            '8  Clave
            '9  IvaPor          No Visible
            '10  Iva            No Visible
            '11 IvaDescuento    No Visible
            '12 IvaIntereses    No Visible
            '13 DescMonto       No Visible
            '14 IdMember        No Visible
            '15 NoFamiliar      No Visible
            '16 Periodo         No Visible
            '17 IdTipoUsuario   No Visible
            '18 TipoCargo       No Visible
            '19 Auxiliar        No Visible
                
            'sCargos(lIndexsCargos) = Format(Date2Days(dPeriodo) + Val(AdoRcsUsuarios!NumeroFamiliar), "0000000000") & vbTab & AdoRcsUsuarios.Fields("Tipo_Usuario.Descripcion") & IIf(nPorAus > 0, "/AUS", "") & IIf(siProporcion < 1, "/PROP. " & iDiasProporcion & " DIAS", "") & IIf(Me.chkDireccionar.Value = 1, " DIRECCIONADO", "") & vbTab & Trim(AdoRcsUsuarios!a_paterno) & " " & Trim(AdoRcsUsuarios!a_materno) & " " & Trim(AdoRcsUsuarios!Nombre) & vbTab & dPeriodo & vbTab & dCantidad & vbTab & dMonto & vbTab & dIntereses & vbTab & dDescuento & vbTab & dImporte & vbTab & AdoRcsUsuarios!Idconcepto & vbTab & dIvaPor & vbTab & dIva & vbTab & dIvaDescuento & vbTab & dIvaIntereses & vbTab & 0 & vbTab & AdoRcsUsuarios!IdMember & vbTab & AdoRcsUsuarios!NumeroFamiliar & vbTab & lFormaPago & vbTab & AdoRcsUsuarios!idtipousuario & vbTab & 0 & vbTab & vbNullString & vbTab & AdoRcsUsuarios!FacORec
            'sCargos(lIndexsCargos) = Format(Date2Days(dPeriodo) + Val(AdoRcsUsuarios!NumeroFamiliar), "0000000000") & vbTab & AdoRcsUsuarios.Fields("Tipo_Usuario.Descripcion") & IIf(nPorAus > 0, "/AUS", "") & IIf(siProporcion < 1, "/PROP. " & iDiasProporcion & " DIAS", "") & IIf(Me.chkDireccionar.Value = 1, " DIRECCIONADO", "") & vbTab & Trim(AdoRcsUsuarios!a_paterno) & " " & Trim(AdoRcsUsuarios!a_materno) & " " & Trim(AdoRcsUsuarios!Nombre) & vbTab & dPeriodo & vbTab & dCantidad & vbTab & dMonto & vbTab & dIntereses & vbTab & dDescuento & vbTab & dImporte & vbTab & AdoRcsUsuarios!Idconcepto & vbTab & dIvaPor & vbTab & dIva & vbTab & dIvaDescuento & vbTab & dIvaIntereses & vbTab & 0 & vbTab & AdoRcsUsuarios!IdMember & vbTab & AdoRcsUsuarios!NumeroFamiliar & vbTab & iPeriodoCalculo & vbTab & AdoRcsUsuarios!idtipousuario & vbTab & 0 & vbTab & vbNullString & vbTab & AdoRcsUsuarios!FacORec
            'aFacturacion(lIndexFacturacion) = Format(Date2Days(dPeriodo) + Val(AdoRcsUsuarios!NumeroFamiliar), "0000000000") & vbTab & AdoRcsUsuarios.Fields("Tipo_Usuario.Descripcion") & IIf(nPorAus > 0, "/AUS", "") & IIf(siProporcion < 1, "/PROP. " & iDiasProporcion & " DIAS", "") & IIf(Me.chkDireccionar.Value = 1, " DIRECCIONADO", "") & vbTab & Trim(AdoRcsUsuarios!a_paterno) & " " & Trim(AdoRcsUsuarios!a_materno) & " " & Trim(AdoRcsUsuarios!Nombre) & vbTab & dPeriodo & vbTab & dCantidad & vbTab & dMonto & vbTab & dIntereses & vbTab & dDescuento & vbTab & dImporte & vbTab & AdoRcsUsuarios!Idconcepto & vbTab & dIvaPor & vbTab & dIva & vbTab & dIvaDescuento & vbTab & dIvaIntereses & vbTab & 0 & vbTab & AdoRcsUsuarios!IdMember & vbTab & AdoRcsUsuarios!NumeroFamiliar & vbTab & iPeriodoCalculo & vbTab & AdoRcsUsuarios!idtipousuario & vbTab & 0 & vbTab & vbNullString & vbTab & AdoRcsUsuarios!FacORec
            
            #If SqlServer_ Then
                strSQL = "INSERT INTO ADEUDOS ("
                strSQL = strSQL & " IdMember,"
                strSQL = strSQL & " IdConcepto,"
                strSQL = strSQL & " Periodo,"
                strSQL = strSQL & " Cantidad,"
                strSQL = strSQL & " Monto,"
                strSQL = strSQL & " Intereses,"
                strSQL = strSQL & " Descuento)"
                strSQL = strSQL & " VALUES ("
                strSQL = strSQL & AdoRcsUsuarios!Idmember & ","
                strSQL = strSQL & AdoRcsUsuarios!IdConcepto & ","
                strSQL = strSQL & "'" & Format(dPeriodo, "yyyymmdd") & "',"
                strSQL = strSQL & 1 & ","
                strSQL = strSQL & dMonto & ","
                strSQL = strSQL & dIntereses & ","
                strSQL = strSQL & dDescuento & ")"
            #Else
                strSQL = "INSERT INTO ADEUDOS ("
                strSQL = strSQL & " IdMember,"
                strSQL = strSQL & " IdConcepto,"
                strSQL = strSQL & " Periodo,"
                strSQL = strSQL & " Cantidad,"
                strSQL = strSQL & " Monto,"
                strSQL = strSQL & " Intereses,"
                strSQL = strSQL & " Descuento)"
                strSQL = strSQL & " VALUES ("
                strSQL = strSQL & AdoRcsUsuarios!Idmember & ","
                strSQL = strSQL & AdoRcsUsuarios!IdConcepto & ","
                strSQL = strSQL & "#" & Format(dPeriodo, "mm/dd/yyyy") & "#,"
                strSQL = strSQL & 1 & ","
                strSQL = strSQL & dMonto & ","
                strSQL = strSQL & dIntereses & ","
                strSQL = strSQL & dDescuento & ")"
            #End If
            
            adocmdCalcMant.CommandText = strSQL
            adocmdCalcMant.Execute
            
        Next
        AdoRcsUsuarios.MoveNext
    Loop
        
        
    Set adocmdCalcMant = Nothing
        
    Set adoRcsFac = Nothing
    
    Set adorcsAus = Nothing
        
    AdoRcsUsuarios.Close
    Set AdoRcsUsuarios = Nothing
End Sub

Private Sub Form_Load()
    Me.dtpFechaCalc.Value = Date
End Sub
