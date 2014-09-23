VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmFacturaRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación de recibos"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   39
      Top             =   6090
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Key             =   "Proceso"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
            Key             =   "Tarea"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmMonto 
      Caption         =   "Monto"
      Height          =   615
      Left            =   7320
      TabIndex        =   35
      Top             =   2280
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         MaxLength       =   13
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmInsc 
      Caption         =   "Datos Inscripción"
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   1080
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtIdMember 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         MaxLength       =   13
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtNombreInsc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   34
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox txtInscripcion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   13
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdBusca 
      Caption         =   "&Buscar"
      Default         =   -1  'True
      Height          =   495
      Left            =   8280
      TabIndex        =   31
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame frmFolio 
      Caption         =   "Folio Recibo"
      Height          =   735
      Left            =   3600
      TabIndex        =   29
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtFolio 
         Height          =   285
         Left            =   120
         MaxLength       =   13
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame frmDatosFactura 
      Caption         =   "Datos de Facturación"
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   6975
      Begin VB.OptionButton optTipoPer 
         Caption         =   "Persona Moral"
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   41
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optTipoPer 
         Caption         =   "Persona Física"
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   40
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtFacNombre 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   16
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txtFacDireccion 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1560
         Width           =   6495
      End
      Begin VB.TextBox txtFacColonia 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtFacCiudad 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox txtFacRFC 
         Height          =   285
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtFacCP 
         Height          =   285
         Left            =   3720
         MaxLength       =   5
         TabIndex        =   11
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtFacDelOMuni 
         Height          =   285
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtFacTelefono 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   9
         Top             =   3240
         Width           =   1695
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbFacEstado 
         Height          =   255
         Left            =   4680
         TabIndex        =   37
         Top             =   2760
         Width           =   2055
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         _Version        =   196616
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   3200
         Columns(0).Caption=   "Estado"
         Columns(0).Name =   "Estado"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "CveEstado"
         Columns(1).Name =   "CveEstado"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Direccion"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Colonia"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Delegación/Municipio"
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "RFC"
         Height          =   255
         Left            =   4680
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "C.P."
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Estado"
         Height          =   255
         Left            =   4680
         TabIndex        =   20
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblIdDireccion 
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblTipoDir 
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3240
         Width           =   1095
      End
   End
   Begin VB.Frame frmTurno 
      Caption         =   "Turno"
      Height          =   735
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      Begin VB.TextBox txtTurno 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   13
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame frmModo 
      Caption         =   "Buscar"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton ctlOpt 
         Caption         =   "Por turno"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton ctlOpt 
         Caption         =   "Por recibo"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame frmFecha 
      Caption         =   "Fecha"
      Height          =   735
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   59179011
         CurrentDate     =   38973
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdFacturar 
      Caption         =   "&Facturar"
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmFacturaRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lidMember As Long
Dim lNoFamilia As Long


Private Sub cmdBusca_Click()
        
    Dim dMonto As Double
    Dim lInscripcionVarios As Long
    
    
    If Me.ctlOpt(0).Value Then
        If Me.txtFolio.Text = vbNullString Then
            MsgBox "Indicar el folio del recibo", vbExclamation, "Verifique"
            Exit Sub
        End If
    Else
    
    End If
    
    
    Screen.MousePointer = vbHourglass

    lInscripcionVarios = Val(ObtieneParametro("INSCRIPCION PARA VARIOS"))
    
    If lInscripcionVarios = 0 Then
        lInscripcionVarios = 80000
    End If
    
    

    If Me.ctlOpt(0).Value Then
        If Not ValidaRecibo(lidMember) Then
            Me.txtFolio.SetFocus
            Exit Sub
        End If
        
        
        ObtieneDatosUsuario lidMember
        ObtieneDatosFactura lidMember, Me
        
        
    Else
        Me.txtFacNombre.Text = "VENTAS PUBLICO EN GENERAL"
        Me.txtFacDireccion.Text = ObtieneParametro("CALLE FISCAL")
        Me.txtFacColonia.Text = ObtieneParametro("COLONIA FISCAL")
        Me.txtFacDelOMuni.Text = ObtieneParametro("DELEGACION FISCAL")
        Me.txtFacCiudad.Text = ObtieneParametro("CIUDAD FISCAL")
        Me.txtFacCP.Text = ObtieneParametro("CP FISCAL")
        Me.ssCmbFacEstado.Text = ObtieneParametro("ESTADO FISCAL")
        Me.txtFacRFC.Text = "XAXX010101000"
        Me.optTipoPer(0).Value = True
        
        ObtieneDatosUsuarioXInsc lInscripcionVarios
        
        
    End If
    
    
    
    dMonto = BuscaImporte()
    
    Screen.MousePointer = vbDefault
    
    If dMonto = 0 Then
        MsgBox "No hay datos para facturar!", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    Me.frmInsc.Visible = True
    Me.frmDatosFactura.Visible = True
    Me.cmdCancelar.Visible = True
    Me.cmdFacturar.Visible = True
    
    
    
    
    Me.txtMonto.Text = Format(dMonto, "#,0.00")
    
    
    
    Me.frmMonto.Visible = True
    
    Me.cmdFacturar.Default = True
    
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdFacturar_Click()
    
    Dim lNumeroFactura As Long
    Dim lNumeroFolioFactura As Long
    
    Dim lTurno As Long
    
    Dim iInitTrans As Integer
    Dim lIdTitular As Long
    Dim lIdFamilia As Long
    
    Dim sNombreFactura As String
    Dim sDireccion As String
    Dim sColonia As String
    Dim sDelegacion As String
    Dim sCiudad As String
    Dim sEstado As String
    Dim sCP As String
    Dim sRfc As String
    Dim sTelefono As String
    Dim sObserva As String
    
    Dim lRowCount As Long
    Dim lRowPagos As Long
    
    
    Dim adocmdFactura As ADODB.Command
    Dim adoRcsTotal As ADODB.Recordset
    Dim AdoRcsPagos As ADODB.Recordset
    
    Dim iResp As Integer
    
    '01/12/09
    Dim dIvaPor As Double
    
    
    Dim dTotalFactura As Double
    
    Dim sTipoPersona  As String
    
    Dim sSerieCFD As String
    Dim sFolioCFD As String
    
    lIdTitular = lidMember
    lIdFamilia = Val(Me.txtInscripcion.Text)
    
    
    
    sTipoPersona = "F"
    
    
    If Me.optTipoPer(0).Value Then 'Persona física
        If (Len(Me.txtFacRFC) <> 13) Then
            iResp = MsgBox("El RFC debe ser de 13 caracteres para personas físicas" & vbCrLf & "¿Desea emitir la factura SIN el IVA desglosado?", vbYesNo + vbQuestion, "Confirme")
            If iResp = vbNo Then
                Exit Sub
            End If
        End If
    Else 'Persona Moral
        If (Len(Me.txtFacRFC) <> 12) Then
            iResp = MsgBox("El RFC debe ser de 12 caracteres para personas morales" & vbCrLf & "¿Desea emitir la factura SIN el IVA desglosado?", vbYesNo + vbQuestion, "Confirme")
            If iResp = vbNo Then
                Exit Sub
            End If
            Exit Sub
        End If
    End If
    
    
        
    If Me.optTipoPer(0).Value Then
        sTipoPersona = "F"
    Else
        sTipoPersona = "M"
    End If
        
    
    
    
    'Cuando es por turno se forza a persona moral
    If Me.ctlOpt(1).Value Then
        sTipoPersona = "M"
    End If
    
    
    
    If MsgBox("¿Desea Generar la factura?", vbQuestion + vbOKCancel, "Confirme") = vbCancel Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    dIvaPor = Val(ObtieneParametro("IVA_GENERAL")) / 100
    
    
    'Por turno
    If Me.ctlOpt(1).Value Then
        sNombreFactura = Trim(Me.txtFacNombre.Text)
        sDireccion = Trim(Me.txtFacDireccion.Text)
        sColonia = Trim(Me.txtFacColonia.Text)
        sDelegacion = Trim(Me.txtFacDelOMuni.Text)
        sCiudad = Trim(Me.txtFacCiudad.Text)
        sEstado = Trim(Me.ssCmbFacEstado.Text)
        sCP = Trim(Me.txtFacCP.Text)
        sRfc = "XAXX010101000"
        sTelefono = vbNullString
        sObserva = "GENERADA AUTOMATICAMENTE"
    Else 'Por folio
        sNombreFactura = Trim(Me.txtFacNombre.Text)
        sDireccion = Trim(Me.txtFacDireccion.Text)
        sColonia = Trim(Me.txtFacColonia.Text)
        sDelegacion = Trim(Me.txtFacDelOMuni.Text)
        sCiudad = Trim(Me.txtFacCiudad.Text)
        sEstado = Trim(Me.ssCmbFacEstado.Text)
        sCP = Trim(Me.txtFacCP.Text)
        sRfc = Trim(Me.txtFacRFC.Text)
        sTelefono = Trim(Me.txtFacTelefono.Text)
        sObserva = "GENERADA AUTOMATICAMENTE"
    End If
    
    
    Me.StatusBar1.Panels("Proceso").Text = "Obteniendo turno"
    
    
    'Turno abierto
    lTurno = OpenShiftF()
    
    If lTurno = 0 Then
        MsgBox "No hay turno abierto!", vbCritical, "Verifique"
        Exit Sub
    End If
    
    
    'Obtiene folios para las facturas
    lNumeroFactura = GetFolio(1, 0)
    lNumeroFolioFactura = GetFolioSerie(1, sSerieFactura)
    
    If lNumeroFactura = -1 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error al obtener folio, reintente", vbCritical
        Exit Sub
    End If
    
    
    
    
    
    
    MDIPrincipal.StatusBar1.Panels(1).Text = "Guardando Factura(s)"
    
    Err.Clear
    Conn.Errors.Clear
    On Error GoTo Error_Catch
    iInitTrans = Conn.BeginTrans
    
    
    'Inserta el encabezado de la factura
    Set adocmdFactura = New ADODB.Command
    adocmdFactura.ActiveConnection = Conn
    adocmdFactura.CommandType = adCmdText
    
    Me.StatusBar1.Panels("Proceso").Text = "Insertando factura encabezado"
    Me.StatusBar1.Panels("Tarea").Text = ""
    
    strSQL = "INSERT INTO FACTURAS"
    strSQL = strSQL & " ( NumeroFactura,"
    strSQL = strSQL & " Folio,"
    strSQL = strSQL & " Serie,"
    strSQL = strSQL & " IdTitular,"
    strSQL = strSQL & " NoFamilia,"
    strSQL = strSQL & " FechaFactura,"
    strSQL = strSQL & " HoraFactura,"
    strSQL = strSQL & " NombreFactura,"
    strSQL = strSQL & " CalleFactura,"
    strSQL = strSQL & " ColoniaFactura,"
    strSQL = strSQL & " DelFactura,"
    strSQL = strSQL & " CiudadFactura,"
    strSQL = strSQL & " EstadoFactura,"
    strSQL = strSQL & " CodPos,"
    strSQL = strSQL & " RFC,"
    strSQL = strSQL & " Tel1,"
    strSQL = strSQL & " Observaciones,"
    strSQL = strSQL & " ImporteConLetra,"
    strSQL = strSQL & " Usuario,"
    strSQL = strSQL & " Turno,"
    strSQL = strSQL & " Caja,"
    strSQL = strSQL & " Direccionado,"
    strSQL = strSQL & " Marca,"
    strSQL = strSQL & " TipoPersona)"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & lNumeroFactura & ","
    strSQL = strSQL & lNumeroFolioFactura & ","
    strSQL = strSQL & "'" & sSerieFactura & "', "
    strSQL = strSQL & lIdTitular & ", "
    strSQL = strSQL & lIdFamilia & ", "
    #If SqlServer_ Then
        strSQL = strSQL & "'" & Format(Now, "yyyymmdd") & "', "
    #Else
        strSQL = strSQL & "#" & Format(Now, "mm/dd/yyyy") & "#, "
    #End If
    strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "', "
    strSQL = strSQL & "'" & Trim(sNombreFactura) & "', "
    strSQL = strSQL & "'" & Trim(sDireccion) & "', "
    strSQL = strSQL & "'" & Trim(sColonia) & "', "
    strSQL = strSQL & "'" & Trim(sDelegacion) & "', "
    strSQL = strSQL & "'" & Trim(sCiudad) & "', "
    strSQL = strSQL & "'" & Trim(sEstado) & "', "
    strSQL = strSQL & "'" & Trim(sCP) & "', "
    strSQL = strSQL & "'" & Trim(sRfc) & "', "
    strSQL = strSQL & "'" & Trim(sTelefono) & "', "
    strSQL = strSQL & "'" & Trim(sObserva) & "', "
    strSQL = strSQL & "'" & Trim(vbNullString) & "',"
    strSQL = strSQL & "'" & Trim(sDB_User) & "',"
    strSQL = strSQL & lTurno & ","
    strSQL = strSQL & iNumeroCaja & ","
    strSQL = strSQL & "'" & vbNullString & "',"
    strSQL = strSQL & 1 & ","
    strSQL = strSQL & "'" & sTipoPersona & "')"
    
    adocmdFactura.CommandText = strSQL
    adocmdFactura.Execute
    
    
    'Inserta el detalle de la factura
    'Por turno
    If Me.ctlOpt(1).Value Then
        #If SqlServer_ Then
            strSQL = "INSERT INTO FACTURAS_DETALLE ( NumeroFactura, Renglon, IdConcepto, IdMember, NumeroFamiliar, IdTipoUsuario, Periodo, FormaPago, Concepto, Cantidad, Importe, Intereses, DescuentoPorciento, Descuento, Total, IvaPorciento, Iva, IvaIntereses, IvaDescuento, TipoCargo, Auxiliar, IdInstructor)"
            strSQL = strSQL & " SELECT " & lNumeroFactura & " AS NumeroFactura, RECIBOS_DETALLE.Renglon, RECIBOS_DETALLE.IdConcepto, RECIBOS_DETALLE.IdMember, RECIBOS_DETALLE.NumeroFamiliar, RECIBOS_DETALLE.IdTipoUsuario, RECIBOS_DETALLE.Periodo, RECIBOS_DETALLE.FormaPago, RECIBOS_DETALLE.Concepto, RECIBOS_DETALLE.Cantidad, RECIBOS_DETALLE.Importe, RECIBOS_DETALLE.Intereses, RECIBOS_DETALLE.DescuentoPorciento, RECIBOS_DETALLE.Descuento, RECIBOS_DETALLE.Total, RECIBOS_DETALLE.IvaPorciento, RECIBOS_DETALLE.Iva, RECIBOS_DETALLE.IvaIntereses, RECIBOS_DETALLE.IvaDescuento, RECIBOS_DETALLE.TipoCargo, RECIBOS_DETALLE.Auxiliar, RECIBOS_DETALLE.IdInstructor"
            strSQL = strSQL & " From RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo=RECIBOS.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) = " & 0 & ")"
            'strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        #Else
            strSQL = "INSERT INTO FACTURAS_DETALLE ( NumeroFactura, Renglon, IdConcepto, IdMember, NumeroFamiliar, IdTipoUsuario, Periodo, FormaPago, Concepto, Cantidad, Importe, Intereses, DescuentoPorciento, Descuento, Total, IvaPorciento, Iva, IvaIntereses, IvaDescuento, TipoCargo, Auxiliar, IdInstructor)"
            strSQL = strSQL & " SELECT " & lNumeroFactura & " AS NumeroFactura, RECIBOS_DETALLE.Renglon, RECIBOS_DETALLE.IdConcepto, RECIBOS_DETALLE.IdMember, RECIBOS_DETALLE.NumeroFamiliar, RECIBOS_DETALLE.IdTipoUsuario, RECIBOS_DETALLE.Periodo, RECIBOS_DETALLE.FormaPago, RECIBOS_DETALLE.Concepto, RECIBOS_DETALLE.Cantidad, RECIBOS_DETALLE.Importe, RECIBOS_DETALLE.Intereses, RECIBOS_DETALLE.DescuentoPorciento, RECIBOS_DETALLE.Descuento, RECIBOS_DETALLE.Total, RECIBOS_DETALLE.IvaPorciento, RECIBOS_DETALLE.Iva, RECIBOS_DETALLE.IvaIntereses, RECIBOS_DETALLE.IvaDescuento, RECIBOS_DETALLE.TipoCargo, RECIBOS_DETALLE.Auxiliar, RECIBOS_DETALLE.IdInstructor"
            strSQL = strSQL & " From RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo=RECIBOS.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) = " & 0 & ")"
            'strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        #End If
    Else 'Por Folio
        #If SqlServer_ Then
            strSQL = "INSERT INTO FACTURAS_DETALLE ( NumeroFactura, Renglon, IdConcepto, IdMember, NumeroFamiliar, IdTipoUsuario, Periodo, FormaPago, Concepto, Cantidad, Importe, Intereses, DescuentoPorciento, Descuento, Total, IvaPorciento, Iva, IvaIntereses, IvaDescuento, TipoCargo, Auxiliar, IdInstructor)"
            strSQL = strSQL & " SELECT " & lNumeroFactura & " AS NumeroFactura, RECIBOS_DETALLE.Renglon, RECIBOS_DETALLE.IdConcepto, RECIBOS_DETALLE.IdMember, RECIBOS_DETALLE.NumeroFamiliar, RECIBOS_DETALLE.IdTipoUsuario, RECIBOS_DETALLE.Periodo, RECIBOS_DETALLE.FormaPago, RECIBOS_DETALLE.Concepto, RECIBOS_DETALLE.Cantidad, RECIBOS_DETALLE.Importe, RECIBOS_DETALLE.Intereses, RECIBOS_DETALLE.DescuentoPorciento, RECIBOS_DETALLE.Descuento, RECIBOS_DETALLE.Total, RECIBOS_DETALLE.IvaPorciento, RECIBOS_DETALLE.Iva, RECIBOS_DETALLE.IvaIntereses, RECIBOS_DETALLE.IvaDescuento, RECIBOS_DETALLE.TipoCargo, RECIBOS_DETALLE.Auxiliar, RECIBOS_DETALLE.IdInstructor"
            strSQL = strSQL & " From RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo=RECIBOS.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.NumeroRecibo) =" & Trim(Me.txtFolio.Text) & ")"
            'strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        #Else
            strSQL = "INSERT INTO FACTURAS_DETALLE ( NumeroFactura, Renglon, IdConcepto, IdMember, NumeroFamiliar, IdTipoUsuario, Periodo, FormaPago, Concepto, Cantidad, Importe, Intereses, DescuentoPorciento, Descuento, Total, IvaPorciento, Iva, IvaIntereses, IvaDescuento, TipoCargo, Auxiliar, IdInstructor)"
            strSQL = strSQL & " SELECT " & lNumeroFactura & " AS NumeroFactura, RECIBOS_DETALLE.Renglon, RECIBOS_DETALLE.IdConcepto, RECIBOS_DETALLE.IdMember, RECIBOS_DETALLE.NumeroFamiliar, RECIBOS_DETALLE.IdTipoUsuario, RECIBOS_DETALLE.Periodo, RECIBOS_DETALLE.FormaPago, RECIBOS_DETALLE.Concepto, RECIBOS_DETALLE.Cantidad, RECIBOS_DETALLE.Importe, RECIBOS_DETALLE.Intereses, RECIBOS_DETALLE.DescuentoPorciento, RECIBOS_DETALLE.Descuento, RECIBOS_DETALLE.Total, RECIBOS_DETALLE.IvaPorciento, RECIBOS_DETALLE.Iva, RECIBOS_DETALLE.IvaIntereses, RECIBOS_DETALLE.IvaDescuento, RECIBOS_DETALLE.TipoCargo, RECIBOS_DETALLE.Auxiliar, RECIBOS_DETALLE.IdInstructor"
            strSQL = strSQL & " From RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo=RECIBOS.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.NumeroRecibo) =" & Trim(Me.txtFolio.Text) & ")"
            'strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        #End If
    End If
    
    adocmdFactura.CommandText = strSQL
    adocmdFactura.Execute
    
    
    'Actualiza el total de la factura
    'Por turno
    If Me.ctlOpt(1).Value Then
        #If SqlServer_ Then
            strSQL = "SELECT Sum(RECIBOS_DETALLE.Total) AS Total"
            strSQL = strSQL & " FROM RECIBOS INNER JOIN RECIBOS_DETALLE ON RECIBOS.NumeroRecibo = RECIBOS_DETALLE.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) = " & 0 & ")"
            'strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        #Else
            strSQL = "SELECT Sum(RECIBOS_DETALLE.Total) AS Total"
            strSQL = strSQL & " FROM RECIBOS INNER JOIN RECIBOS_DETALLE ON RECIBOS.NumeroRecibo = RECIBOS_DETALLE.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) = " & 0 & ")"
            'strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        #End If
    Else
        #If SqlServer_ Then
            strSQL = "SELECT Sum(RECIBOS_DETALLE.Total) AS Total"
            strSQL = strSQL & " FROM RECIBOS INNER JOIN RECIBOS_DETALLE ON RECIBOS.NumeroRecibo = RECIBOS_DETALLE.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.NumeroRecibo) =" & Trim(Me.txtFolio.Text) & ")"
            'strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        #Else
            strSQL = "SELECT Sum(RECIBOS_DETALLE.Total) AS Total"
            strSQL = strSQL & " FROM RECIBOS INNER JOIN RECIBOS_DETALLE ON RECIBOS.NumeroRecibo = RECIBOS_DETALLE.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.NumeroRecibo) =" & Trim(Me.txtFolio.Text) & ")"
            'strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        #End If
    End If
    
    
    Set adoRcsTotal = New ADODB.Recordset
    adoRcsTotal.CursorLocation = adUseServer
    
    adoRcsTotal.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adoRcsTotal.EOF Then
        dTotalFactura = adoRcsTotal!Total
    End If
    
    adoRcsTotal.Close
    Set adoRcsTotal = Nothing
    
    'Actualiza el total de la factura
    strSQL = "UPDATE FACTURAS SET"
    strSQL = strSQL & " Total=" & dTotalFactura
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((NumeroFactura) = " & lNumeroFactura & ")"
    strSQL = strSQL & ")"
    
    adocmdFactura.CommandText = strSQL
    adocmdFactura.Execute
    
    
    
    
    
    'Copia la forma de pago
    
    ' Por Turno
    If Me.ctlOpt(1).Value Then
        #If SqlServer_ Then
            strSQL = "SELECT FORMA_PAGO.IdFormaPago, PAGOS_RECIBO.OpcionPago, PAGOS_RECIBO.Referencia, PAGOS_RECIBO.IdAfiliacion, PAGOS_RECIBO.LoteNumero, PAGOS_RECIBO.OperacionNumero, PAGOS_RECIBO.ImporteRecibido, PAGOS_RECIBO.FechaOperacion, PAGOS_RECIBO.Importe"
            strSQL = strSQL & " FROM (PAGOS_RECIBO INNER JOIN FORMA_PAGO ON PAGOS_RECIBO.IdFormaPago = FORMA_PAGO.IdFormaPago) INNER JOIN RECIBOS ON PAGOS_RECIBO.NumeroRecibo = RECIBOS.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & "((RECIBOS.FechaFactura) = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) =" & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) =" & 0 & ")"
            strSQL = strSQL & ")"
            strSQL = strSQL & " ORDER BY PAGOS_RECIBO.NumeroRecibo, PAGOS_RECIBO.Renglon"
        #Else
            strSQL = "SELECT FORMA_PAGO.IdFormaPago, PAGOS_RECIBO.OpcionPago, PAGOS_RECIBO.Referencia, PAGOS_RECIBO.IdAfiliacion, PAGOS_RECIBO.LoteNumero, PAGOS_RECIBO.OperacionNumero, PAGOS_RECIBO.ImporteRecibido, PAGOS_RECIBO.FechaOperacion, PAGOS_RECIBO.Importe"
            strSQL = strSQL & " FROM (PAGOS_RECIBO INNER JOIN FORMA_PAGO ON PAGOS_RECIBO.IdFormaPago = FORMA_PAGO.IdFormaPago) INNER JOIN RECIBOS ON PAGOS_RECIBO.NumeroRecibo = RECIBOS.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & "((RECIBOS.FechaFactura) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) =" & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) =" & 0 & ")"
            strSQL = strSQL & ")"
            strSQL = strSQL & " ORDER BY PAGOS_RECIBO.NumeroRecibo, PAGOS_RECIBO.Renglon"
        #End If
    Else 'Por folio
        #If SqlServer_ Then
            strSQL = "SELECT FORMA_PAGO.IdFormaPago, PAGOS_RECIBO.OpcionPago, PAGOS_RECIBO.Referencia, PAGOS_RECIBO.IdAfiliacion, PAGOS_RECIBO.LoteNumero, PAGOS_RECIBO.OperacionNumero, PAGOS_RECIBO.ImporteRecibido, PAGOS_RECIBO.FechaOperacion, PAGOS_RECIBO.Importe"
            strSQL = strSQL & " FROM (PAGOS_RECIBO INNER JOIN FORMA_PAGO ON PAGOS_RECIBO.IdFormaPago = FORMA_PAGO.IdFormaPago) INNER JOIN RECIBOS ON PAGOS_RECIBO.NumeroRecibo = RECIBOS.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & "((RECIBOS.FechaFactura) = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) =" & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) =" & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.NumeroRecibo) =" & Trim(Me.txtFolio.Text) & ")"
            strSQL = strSQL & ")"
            strSQL = strSQL & " ORDER BY PAGOS_RECIBO.NumeroRecibo, PAGOS_RECIBO.Renglon"
        #Else
            strSQL = "SELECT FORMA_PAGO.IdFormaPago, PAGOS_RECIBO.OpcionPago, PAGOS_RECIBO.Referencia, PAGOS_RECIBO.IdAfiliacion, PAGOS_RECIBO.LoteNumero, PAGOS_RECIBO.OperacionNumero, PAGOS_RECIBO.ImporteRecibido, PAGOS_RECIBO.FechaOperacion, PAGOS_RECIBO.Importe"
            strSQL = strSQL & " FROM (PAGOS_RECIBO INNER JOIN FORMA_PAGO ON PAGOS_RECIBO.IdFormaPago = FORMA_PAGO.IdFormaPago) INNER JOIN RECIBOS ON PAGOS_RECIBO.NumeroRecibo = RECIBOS.NumeroRecibo"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & "((RECIBOS.FechaFactura) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) =" & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) =" & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.NumeroRecibo) =" & Trim(Me.txtFolio.Text) & ")"
            strSQL = strSQL & ")"
            strSQL = strSQL & " ORDER BY PAGOS_RECIBO.NumeroRecibo, PAGOS_RECIBO.Renglon"
        #End If
    End If
    
    Set AdoRcsPagos = New ADODB.Recordset
    AdoRcsPagos.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Do Until AdoRcsPagos.EOF
    
        Me.StatusBar1.Panels("Proceso").Text = "Insertando pagos"
        Me.StatusBar1.Panels("Tarea").Text = "Recibo # " & lRowPagos
    
        lRowPagos = lRowPagos + 1
    
        strSQL = "INSERT INTO PAGOS_FACTURA ("
        strSQL = strSQL & " NumeroFactura, "
        strSQL = strSQL & " Renglon, "
        strSQL = strSQL & " IdFormaPago, "
        strSQL = strSQL & " OpcionPago, "
        strSQL = strSQL & " Importe, "
        strSQL = strSQL & " Referencia, "
        strSQL = strSQL & " IdAfiliacion, "
        strSQL = strSQL & " LoteNumero, "
        strSQL = strSQL & " OperacionNumero, "
        strSQL = strSQL & " ImporteRecibido, "
        strSQL = strSQL & " FechaOperacion) "
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lNumeroFactura & ", "
        strSQL = strSQL & lRowPagos & ", "
        strSQL = strSQL & AdoRcsPagos!IdFormaPago & ","
        strSQL = strSQL & "'" & AdoRcsPagos!OpcionPago & "'" & ","
        strSQL = strSQL & AdoRcsPagos!Importe & ","
        strSQL = strSQL & "'" & AdoRcsPagos!Referencia & "',"
        strSQL = strSQL & AdoRcsPagos!IdAfiliacion & ","
        strSQL = strSQL & "'" & AdoRcsPagos!LoteNumero & "',"
        strSQL = strSQL & "'" & AdoRcsPagos!OperacionNumero & "',"
        strSQL = strSQL & AdoRcsPagos!ImporteRecibido & ","
        #If SqlServer_ Then
            strSQL = strSQL & "'" & Format(AdoRcsPagos!FechaOperacion, "yyyymmdd") & "')"
        #Else
            strSQL = strSQL & "#" & Format(AdoRcsPagos!FechaOperacion, "mm/dd/yyyy") & "#)"
        #End If
                
        adocmdFactura.CommandText = strSQL
        adocmdFactura.Execute
        
        AdoRcsPagos.MoveNext
    Loop
    
    AdoRcsPagos.Close
    
    Set AdoRcsPagos = Nothing
    
    
    Me.StatusBar1.Panels("Proceso").Text = "Marcando recibos como facturados"
    Me.StatusBar1.Panels("Tarea").Text = ""
    
    'Actualiza el campo factura en recibos.
    strSQL = "UPDATE RECIBOS"
    strSQL = strSQL & " SET FACTURA = " & lNumeroFactura
    strSQL = strSQL & " WHERE ("
    
    'Por folio
    If Me.ctlOpt(0).Value Then
        strSQL = strSQL & "((NumeroRecibo) = " & Trim(Me.txtFolio.Text) & ")"
    Else
        'Si es por turno
        'Actualiza aquellas facturas de la fecha, turno, caja
        'que no esten canceladas y que no hayan sido facturadas previamente.
        #If SqlServer_ Then
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) = " & 0 & ")"
        #Else
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & lTurno & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) = " & 0 & ")"
        #End If
    End If
    
    strSQL = strSQL & ")"
    
    adocmdFactura.CommandText = strSQL
    adocmdFactura.Execute
    
    
    'Crea el registro en Facturas_Cancela
    
    Me.StatusBar1.Panels("Proceso").Text = "Creando registro de cancelación"
    Me.StatusBar1.Panels("Tarea").Text = ""
    
    strSQL = "INSERT INTO FACTURAS_CANCELA ("
    strSQL = strSQL & " NumeroFactura,"
    strSQL = strSQL & " CadenaCancela1)"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & lNumeroFactura & ","
    strSQL = strSQL & "'" & "900001" & "')"
    
    adocmdFactura.CommandText = strSQL
    adocmdFactura.Execute
    
    
    Conn.CommitTrans
    
    
    Me.StatusBar1.Panels("Proceso").Text = "Generando CFD en " & ObtieneParametro("URL_WS_CFD")
    
    sSerieCFD = ObtieneParametro("SERIE_CFD_FACTURA_CAJA")
    
    sFolioCFD = GeneraCFD(lNumeroFactura, sSerieCFD, "ingreso")
    
    If Len(sFolioCFD) > 12 Then
        MsgBox "Ocurrio un error generando el CFD" & vbCrLf & sFolioCFD, vbCritical, "Error"
    Else
        If sFolioCFD <> vbNullString Then
            Me.StatusBar1.Panels("Proceso").Text = "Actualizando FolioCFD"
            DoEvents
            If ActualizaFolioCFD(lNumeroFactura, sFolioCFD, sSerieCFD, "F") = 0 Then
            End If
        End If
    End If
    
    Me.StatusBar1.Panels("Proceso").Text = "Terminado"
    Me.StatusBar1.Panels("Tarea").Text = "Se creo la factura " & lNumeroFactura
    
    iInitTrans = 0
    
    
    Set adocmdFactura = Nothing
    
    Me.frmInsc.Visible = False
    Me.frmDatosFactura.Visible = False
    Me.frmMonto.Visible = False
    
    Me.cmdCancelar.Visible = False
    Me.cmdFacturar.Visible = False
    
    Me.cmdBusca.Default = True
    
    
    Me.txtFolio.Text = vbNullString
    Me.txtFacNombre.Text = vbNullString
    Me.txtFacDireccion.Text = vbNullString
    Me.txtFacColonia.Text = vbNullString
    Me.txtFacDelOMuni.Text = vbNullString
    Me.txtFacCiudad.Text = vbNullString
    Me.txtFacRFC.Text = vbNullString
    Me.txtFacCP.Text = vbNullString
    Me.txtFacTelefono.Text = vbNullString
    
    Me.txtInscripcion.Text = vbNullString
    Me.txtNombreInsc.Text = vbNullString
    Me.txtMonto.Text = vbNullString
    
    
    Screen.MousePointer = vbDefault
    
    'Para imprimir la factura
    
    
    lNumFacIniImp = lNumeroFactura
    lNumFacFinImp = lNumeroFactura
   
    lNumFolioFacIniImp = sSerieFactura & lNumeroFolioFactura
    lNumFolioFacFinImp = sSerieFactura & lNumeroFolioFactura
    
    
    Dim frmImp As New frmImpFac
    
    frmImp.cModo = "F"
    frmImp.Tag = "F"
    
    frmImp.lNumeroInicial = lNumeroFactura
    frmImp.lNumeroFinal = lNumeroFactura
    
    frmImp.Show 1
    
    Exit Sub

Error_Catch:
    
    If iInitTrans Then
        Conn.RollbackTrans
    End If
    
    Screen.MousePointer = vbDefault
    
    MsgError
    
End Sub



Private Sub ctlOpt_Click(index As Integer)
    
    If Me.ctlOpt(0).Value Then
        frmFolio.Visible = True
        frmFecha.Visible = False
        frmTurno.Visible = False
    Else
        frmFolio.Visible = False
        frmFecha.Visible = True
        frmTurno.Visible = True
    End If
    
End Sub

Private Sub Form_Activate()
    'Me.txtFolio.SetFocus
End Sub

Private Sub Form_Load()
    Me.Height = 6825
    Me.Width = 10110
    
    
  
    
    Me.dtpFecha.Value = Date
    
    Me.ctlOpt(0).Value = True
    
    Me.txtTurno.Text = OpenShiftF()
    
    strSQL = "SELECT nomEntFederativa, cveEntFederativa "
    strSQL = strSQL & " FROM ENTFEDERATIVA"
    strSQL = strSQL & " ORDER BY nomEntFederativa"
    
    LlenaSsCombo Me.ssCmbFacEstado, Conn, strSQL, 2
        
    
    
    
    CentraForma MDIPrincipal, Me
End Sub




Private Function BuscaImporte() As Double
    
    Dim adorcsBusca As ADODB.Recordset
    
    BuscaImporte = 0
    
    #If SqlServer_ Then
        If Me.ctlOpt(1).Value Then
            strSQL = "SELECT Sum(RECIBOS_DETALLE.Total) AS Total"
            strSQL = strSQL & " FROM ((RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo = RECIBOS.NumeroRecibo))"
            strSQL = strSQL & " INNER JOIN CONCEPTO_INGRESOS ON RECIBOS_DETALLE.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & Me.txtTurno.Text & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) = " & 0 & ")"
            strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        Else
            strSQL = "SELECT Sum(RECIBOS_DETALLE.Total) AS Total"
            strSQL = strSQL & " FROM (RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo = RECIBOS.NumeroRecibo)"
            strSQL = strSQL & " INNER JOIN CONCEPTO_INGRESOS ON RECIBOS_DETALLE.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & Me.txtTurno.Text & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.NumeroRecibo) =" & Me.txtFolio.Text & ")"
            strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        End If
    #Else
        If Me.ctlOpt(1).Value Then
            strSQL = "SELECT Sum(RECIBOS_DETALLE.Total) AS Total"
            strSQL = strSQL & " FROM ((RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo = RECIBOS.NumeroRecibo))"
            strSQL = strSQL & " INNER JOIN CONCEPTO_INGRESOS ON RECIBOS_DETALLE.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & Me.txtTurno.Text & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.Factura) = " & 0 & ")"
            strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        Else
            strSQL = "SELECT Sum(RECIBOS_DETALLE.Total) AS Total"
            strSQL = strSQL & " FROM (RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo = RECIBOS.NumeroRecibo)"
            strSQL = strSQL & " INNER JOIN CONCEPTO_INGRESOS ON RECIBOS_DETALLE.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((RECIBOS.FechaFactura) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((RECIBOS.Turno) = " & Me.txtTurno.Text & ")"
            strSQL = strSQL & " AND ((RECIBOS.Caja) = " & iNumeroCaja & ")"
            strSQL = strSQL & " AND ((RECIBOS.Cancelada) = " & 0 & ")"
            strSQL = strSQL & " AND ((RECIBOS.NumeroRecibo) =" & Me.txtFolio.Text & ")"
            strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.NoFacturable) = " & 0 & ")"
            strSQL = strSQL & " )"
        End If
    #End If
    
    Set adorcsBusca = New ADODB.Recordset
    adorcsBusca.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsBusca.EOF Then
        If Not IsNull(adorcsBusca!Total) Then
            BuscaImporte = adorcsBusca!Total
        End If
    End If
    
    
End Function

Private Function ValidaRecibo(ByRef lidMember As Long) As Boolean
    
    Dim adorcsRecibo As ADODB.Recordset
    
    
    ValidaRecibo = False

    strSQL = "SELECT RECIBOS.NumeroRecibo, RECIBOS.IdTitular, RECIBOS.NoFamilia, RECIBOS.FechaFactura, RECIBOS.Caja, RECIBOS.Turno, RECIBOS.Cancelada, RECIBOS.Factura"
    strSQL = strSQL & " FROM RECIBOS"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & " ((RECIBOS.NumeroRecibo) = " & Trim(Me.txtFolio.Text) & ")"
    strSQL = strSQL & ")"
    
    Set adorcsRecibo = New ADODB.Recordset
    
    adorcsRecibo.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsRecibo.EOF Then
        If adorcsRecibo!FechaFactura <> Date Then
            MsgBox "Este recibo NO es del día de hoy", vbExclamation, "Verifique"
            Exit Function
        End If
        
        If adorcsRecibo!Caja <> iNumeroCaja Then
            MsgBox "Este recibo Es de otra caja", vbExclamation, "Verifique"
            Exit Function
        End If
        
        If adorcsRecibo!Turno <> Val(Me.txtTurno.Text) Then
            MsgBox "Este recibo NO es del turno activo", vbExclamation, "Verifique"
            Exit Function
        End If
        
        If adorcsRecibo!Cancelada = True Then
            MsgBox "Este recibo está Cancelado", vbExclamation, "Verifique"
            Exit Function
        End If
            
        If adorcsRecibo!Factura <> 0 Then
            MsgBox "Este recibo YA fue facturado", vbExclamation, "Verifique"
            Exit Function
        End If
        
        lidMember = adorcsRecibo!IdTitular
        
    End If
    
    
    adorcsRecibo.Close
    
    Set adorcsRecibo = Nothing
    
    ValidaRecibo = True
    
End Function



Private Sub ObtieneDatosUsuario(lidMember As Long)
    Dim adorcsUsuario As ADODB.Recordset
    
    strSQL = "SELECT A_MATERNO & ' ' & A_PATERNO & ' ' &  NOMBRE AS NOMBRE, NoFamilia"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & " ((IdMember) = " & lidMember & ")"
    strSQL = strSQL & ")"
    
    
    Set adorcsUsuario = New ADODB.Recordset
    adorcsUsuario.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsUsuario.EOF Then
        Me.txtInscripcion.Text = adorcsUsuario!NoFamilia
        Me.txtNombreInsc.Text = adorcsUsuario!Nombre
    End If
    
    adorcsUsuario.Close
    Set adorcsUsuario = Nothing
    
End Sub

Private Sub ObtieneDatosUsuarioXInsc(lIdFamilia As Long)
    Dim adorcsUsuario As ADODB.Recordset
    
    #If SqlServer_ Then
        strSQL = "SELECT A_MATERNO + ' ' + A_PATERNO + ' ' +  NOMBRE AS NOMBRE, NoFamilia,IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " NoFamilia = " & lIdFamilia
        strSQL = strSQL & " AND IdMember = IdTitular"
    #Else
        strSQL = "SELECT A_MATERNO & ' ' & A_PATERNO & ' ' &  NOMBRE AS NOMBRE, NoFamilia,IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((NoFamilia) = " & lIdFamilia & ")"
        strSQL = strSQL & " AND ((IdMember) = IdTitular)"
        strSQL = strSQL & ")"
    #End If
    
    Set adorcsUsuario = New ADODB.Recordset
    adorcsUsuario.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsUsuario.EOF Then
        Me.txtInscripcion.Text = adorcsUsuario!NoFamilia
        Me.txtNombreInsc.Text = adorcsUsuario!Nombre
        lidMember = adorcsUsuario!Idmember
    End If
    
    adorcsUsuario.Close
    Set adorcsUsuario = Nothing
    
End Sub

