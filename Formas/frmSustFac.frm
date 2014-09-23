VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmSustFac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sustitución de facturas"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   47
      Top             =   7350
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   "Proceso"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   "Tarea"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   7440
      TabIndex        =   46
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7440
      TabIndex        =   45
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdGeneraNotaCredito 
      Caption         =   "Genera N. de Crédito"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   44
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkNoNC 
      Caption         =   "No generar Nota de crédito"
      Height          =   375
      Left            =   7320
      TabIndex        =   38
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox chkGuarda 
      Caption         =   "Guarda datos como fiscales"
      Height          =   375
      Left            =   7320
      TabIndex        =   37
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Información de factura a sustituir"
      Height          =   1215
      Left            =   120
      TabIndex        =   31
      Top             =   1200
      Visible         =   0   'False
      Width           =   6975
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Importe"
         Height          =   255
         Left            =   5280
         TabIndex        =   43
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Turno"
         Height          =   255
         Left            =   4200
         TabIndex        =   42
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Caja"
         Height          =   255
         Left            =   3240
         TabIndex        =   41
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1440
         TabIndex        =   40
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Inscripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5280
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblCaja 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3240
         TabIndex        =   35
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblTurno 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4200
         TabIndex        =   34
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblFecha 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   33
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblInsc 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Factura a sustituir"
      Height          =   855
      Left            =   120
      TabIndex        =   28
      Top             =   240
      Width           =   6975
      Begin VB.TextBox txtSerie 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Busca"
         Default         =   -1  'True
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtNoFactura 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Serie"
         Height          =   255
         Left            =   2520
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "No. de Factura"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame frmDatosFactura 
      Caption         =   "Datos de Facturación (Nueva Factura)"
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtFacTelefono 
         Height          =   285
         Left            =   1680
         MaxLength       =   20
         TabIndex        =   14
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtFacDelOMuni 
         Height          =   285
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   13
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtFacCP 
         Height          =   285
         Left            =   3720
         MaxLength       =   5
         TabIndex        =   12
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtFacRFC 
         Height          =   285
         Left            =   4800
         MaxLength       =   20
         TabIndex        =   11
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtFacCiudad 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox txtFacColonia 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txtFacDireccion 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1560
         Width           =   6495
      End
      Begin VB.TextBox txtFacNombre 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   7
         Top             =   960
         Width           =   4455
      End
      Begin VB.OptionButton optTipoPer 
         Caption         =   "Persona Física"
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optTipoPer 
         Caption         =   "Persona Moral"
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbFacEstado 
         Height          =   255
         Left            =   4680
         TabIndex        =   15
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
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblTipoDir 
         Enabled         =   0   'False
         Height          =   255
         Left            =   5760
         TabIndex        =   26
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblIdDireccion 
         Enabled         =   0   'False
         Height          =   255
         Left            =   4800
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Estado"
         Height          =   255
         Left            =   4680
         TabIndex        =   24
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "C.P."
         Height          =   255
         Left            =   3720
         TabIndex        =   23
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "RFC"
         Height          =   255
         Left            =   4680
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Ciudad"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Delegación/Municipio"
         Height          =   255
         Left            =   4680
         TabIndex        =   19
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Colonia"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Direccion"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5400
      TabIndex        =   20
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdProceder 
      Caption         =   "Sustituir Factura"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "frmSustFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iModo As Integer

Private lNumFactura As Long
Private Sub cmdBusca_Click()
    
    Dim adorcs As ADODB.Recordset
    Dim sCadFolio As String
    
    Dim lidMember As Long
    Dim lNumFac As Long
    
    Me.txtNoFactura.Text = UCase(Me.txtNoFactura.Text)
    Me.txtSerie.Text = UCase(Me.txtSerie.Text)
    
    
    sCadFolio = Trim(Me.txtNoFactura.Text)
    sCadFolio = "('" & sCadFolio & "'," & "'0" & sCadFolio & "'," & "'00" & sCadFolio & "'," & "'000" & sCadFolio & "')"
    
    strSQL = ""
    
    strSQL = "SELECT F.Numerofactura, F.FolioCFD, F.SerieCFD, F.NoFamilia, F.IdTitular, F.FechaFactura, F.NombreFactura, F.Cancelada, F.Caja, F.Turno, F.Total, "
    strSQL = strSQL & " F.NombreFactura, F.CalleFactura, F.ColoniaFactura, F.DelFactura, F.CiudadFactura, F.EstadoFactura, F.CodPos, F.RFC, F.Tel1"
    strSQL = strSQL & " FROM FACTURAS F"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & " ((F.FolioCFD) In " & sCadFolio & ")"
    strSQL = strSQL & " And ((F.SerieCFD)='" & Trim(Me.txtSerie.Text) & "')"
    strSQL = strSQL & ")"
    
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    If adorcs.EOF Then
        adorcs.Close
        Set adorcs = Nothing
        MsgBox "¡Factura no localizada!", vbExclamation, "Verifique"
        'Me.Frame1.Visible = False
        If Me.txtNoFactura.Visible Then Me.txtNoFactura.SetFocus
        Exit Sub
    End If
    
    If adorcs!Cancelada = True Then
        adorcs.Close
        Set adorcs = Nothing
        MsgBox "¡Esta factura está cancelada!", vbExclamation, "Error"
        'Me.Frame1.Visible = False
        Me.txtNoFactura.SetFocus
        Exit Sub
    End If
    
    
    Me.lblInsc.Caption = adorcs!NoFamilia
    Me.lblFecha.Caption = Format(adorcs!FechaFactura, "dd/mm/yyyy")
    Me.lblCaja.Caption = adorcs!Caja
    Me.lblTurno.Caption = adorcs!Turno
    Me.lblTotal.Caption = Format(adorcs!Total, "$#,#0.00")
    
    lNumFactura = adorcs!NumeroFactura
    lidMember = adorcs!IdTitular
    
    ObtieneDatosFactura lidMember, Me
    
    Me.Frame2.Visible = True
    Me.frmDatosFactura.Visible = True
    
    adorcs.Close
    Set adorcs = Nothing
    
    
    Me.Frame1.Visible = True
    
    If iModo = 0 Then
        Me.cmdProceder.Enabled = True
    Else
        Me.cmdGeneraNotaCredito.Enabled = True
    End If
    
    
    
End Sub

Private Sub cmdGeneraNotaCredito_Click()
    If Not GeneraNCdesdeFactura(lNumFactura) Then
        MsgBox "Ocurrio un error al generar la nota de crédito", vbExclamation, "Error"
    End If
    
    Unload Me
    
End Sub

Private Sub cmdProceder_Click()
    
    Dim iResp As VbMsgBoxResult
    
    
    Dim lTurno As Long
    Dim lNumFac As Long
    Dim lNotaCred As Long
    
    Dim sSerieCFD As String
    Dim sFolioCFD As String

    'iResp = MsgBox("Se generará una nueva factura, asi como una nota de crédito" & vbCrLf & "¿Desea continuar?", vbOKCancel + vbQuestion, "Confirme")
    
    'If iResp = vbCancel Then
    '    Exit Sub
    'End If
    
    lTurno = OpenShiftF()
    
    If lTurno = 0 Then
        MsgBox "No hay turno abierto!", vbCritical, "Verifique"
        Exit Sub
    End If
    
    
    
    lNumFac = GeneraFactura(lNumFactura)
    
    Me.StatusBar1.Panels("Proceso").Text = "Generando CFD en " & ObtieneParametro("URL_WS_CFD")
    
    If iNumeroCaja <> 2 Then
        sSerieCFD = ObtieneParametro("SERIE_CFD_FACTURA_CAJA")
    Else
        sSerieCFD = ObtieneParametro("SERIE_CFD_FACTURA_DIRE")
    End If
    
    sFolioCFD = GeneraCFD(lNumFac, sSerieCFD, "ingreso")
    
    If Len(sFolioCFD) > 12 Then
        MsgBox "Ocurrio un error generando el CFD" & vbCrLf & sFolioCFD, vbCritical, "Error"
    Else
        If sFolioCFD <> vbNullString Then
            Me.StatusBar1.Panels("Proceso").Text = "Actualizando FolioCFD"
            DoEvents
            If ActualizaFolioCFD(lNumFac, sFolioCFD, sSerieCFD, "F") = 0 Then
            End If
        End If
    End If
    
    Me.StatusBar1.Panels("Proceso").Text = "Terminado"
    Me.StatusBar1.Panels("Tarea").Text = "Se creo la factura " & lNumFac

    
    
    
    
    
    
    If Not GeneraNCdesdeFactura(lNumFactura) Then
        MsgBox "Ocurrio un error al generar la nota de crédito", vbExclamation, "Error"
    End If
    
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    GeneraNCdesdeFactura CLng(Me.Text1.Text)
    MsgBox "Nota Generada", vbInformation, "Ok"
    Me.Text1.Text = vbNullString
End Sub

Private Sub Form_Activate()
    'Si es para sustituir factura iModo = 0
    If iModo = 0 Then
        Me.cmdProceder.Visible = True
        Me.cmdGeneraNotaCredito.Visible = False
    Else
        Me.cmdProceder.Visible = False
        Me.cmdGeneraNotaCredito.Visible = True
        Me.Caption = "Generar nota de Crédito"
    End If
End Sub

Private Sub Form_Load()
    CentraForma MDIPrincipal, Me
End Sub



Private Sub txtNoFactura_GotFocus()
    
    Me.txtNoFactura.SelStart = 0
    Me.txtNoFactura.SelLength = Len(Me.txtNoFactura.Text)
    
End Sub


Private Function GeneraFactura(lFacActual As Long) As Long
    Dim lNumeroFactura As Long
    Dim lNumeroFolioFactura As Long
    
    Dim lTurno As Long
    
    Dim iInitTrans As Integer
    
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
    
    Dim dIvaPor As Double
    
    
    Dim dTotalFactura As Double
    
    Dim sTipoPersona  As String
    
    Dim lIdTitular As Long
    Dim lNoFamilia As Long
    Dim sFolioFac As String
    Dim dFechaFac As Date
    Dim doTotalFac As Double
    
    
    sTipoPersona = "F"
    
    
    GeneraFactura = 0
    
    
    If Me.optTipoPer(0).Value Then 'Persona física
        If (Len(Me.txtFacRFC) <> 13) Then
            iResp = MsgBox("El RFC debe ser de 13 caracteres para personas físicas" & vbCrLf & "¿Desea emitir la factura SIN el IVA desglosado?", vbYesNo + vbQuestion, "Confirme")
            If iResp = vbNo Then
                Exit Function
            End If
        End If
    Else 'Persona Moral
        If (Len(Me.txtFacRFC) <> 12) Then
            iResp = MsgBox("El RFC debe ser de 12 caracteres para personas morales" & vbCrLf & "¿Desea emitir la factura SIN el IVA desglosado?", vbYesNo + vbQuestion, "Confirme")
            If iResp = vbNo Then
                Exit Function
            End If
            Exit Function
        End If
    End If
    
    
        
    If Me.optTipoPer(0).Value Then
        sTipoPersona = "F"
    Else
        sTipoPersona = "M"
    End If
        
    
    
    
    
    
    
    If MsgBox("¿Desea Generar la factura?", vbQuestion + vbOKCancel, "Confirme") = vbCancel Then
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    
    'Obtiene los valores de la factura actual
    
    strSQL = "SELECT FACTURAS.Folio, FACTURAS.Serie, FACTURAS.FolioCFD, FACTURAS.SerieCFD, FACTURAS.IdTitular, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.Total"
    strSQL = strSQL & " FROM FACTURAS"
    strSQL = strSQL & " Where ("
    strSQL = strSQL & "((FACTURAS.NumeroFactura)=" & lFacActual & ")"
    strSQL = strSQL & ")"

    Set adoRcsTotal = New ADODB.Recordset
    adoRcsTotal.CursorLocation = adUseServer
    
    adoRcsTotal.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adoRcsTotal.EOF Then
        lIdTitular = adoRcsTotal!IdTitular
        lNoFamilia = adoRcsTotal!NoFamilia
        dFechaFac = adoRcsTotal!FechaFactura
        doTotalFac = adoRcsTotal!Total
        
        
        If IsNull(adoRcsTotal!FolioCFD) Then
            sFolioFac = adoRcsTotal!Folio & adoRcsTotal!Serie
        Else
            sFolioFac = adoRcsTotal!FolioCFD & adoRcsTotal!SerieCFD
        End If
        
        
    End If
    
    adoRcsTotal.Close
    Set adoRcsTotal = Nothing
    
    
    
    dIvaPor = Val(ObtieneParametro("IVA_GENERAL")) / 100
    
    
    sNombreFactura = Trim(Me.txtFacNombre.Text)
    sDireccion = Trim(Me.txtFacDireccion.Text) & "."
    sColonia = Trim(Me.txtFacColonia.Text)
    sDelegacion = Trim(Me.txtFacDelOMuni.Text)
    sCiudad = Trim(Me.txtFacCiudad.Text)
    sEstado = Trim(Me.ssCmbFacEstado.Text)
    sCP = Trim(Me.txtFacCP.Text)
    sRfc = Trim(Me.txtFacRFC.Text)
    sTelefono = Trim(Me.txtFacTelefono.Text)
    sObserva = "SUSTITUYE A LA FACTURA " & sFolioFac & " DEL " & dFechaFac
    
    
    'Me.StatusBar1.Panels("Proceso").Text = "Obteniendo turno"
    
    
    'Turno abierto
    lTurno = OpenShiftF()
    
    If lTurno = 0 Then
        MsgBox "No hay turno abierto!", vbCritical, "Verifique"
        Exit Function
    End If
    
    
    'Obtiene folios para las factura
    lNumeroFactura = GetFolio(1, 0)
    lNumeroFolioFactura = GetFolioSerie(1, sSerieFactura)
    
    If lNumeroFactura = -1 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error al obtener folio, reintente", vbCritical
        Exit Function
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
    
    'Me.StatusBar1.Panels("Proceso").Text = "Insertando factura encabezado"
    'Me.StatusBar1.Panels("Tarea").Text = ""
    
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
    strSQL = strSQL & " Total,"
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
    strSQL = strSQL & lNoFamilia & ", "
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
    strSQL = strSQL & doTotalFac & ","
    strSQL = strSQL & "'" & Trim(sDB_User) & "',"
    strSQL = strSQL & lTurno & ","
    strSQL = strSQL & iNumeroCaja & ","
    strSQL = strSQL & "'" & vbNullString & "',"
    strSQL = strSQL & 2 & ","
    strSQL = strSQL & "'" & sTipoPersona & "')"
    
    adocmdFactura.CommandText = strSQL
    adocmdFactura.Execute
    
    
    'Inserta el detalle de la factura
    strSQL = "INSERT INTO FACTURAS_DETALLE ( NumeroFactura, Renglon, IdConcepto, IdMember, NumeroFamiliar, IdTipoUsuario, Periodo, FormaPago, Concepto, Cantidad, Importe, Intereses, DescuentoPorciento, Descuento, Total, IvaPorciento, Iva, IvaIntereses, IvaDescuento, TipoCargo, Auxiliar, IdInstructor, Unidad)"
    strSQL = strSQL & " SELECT " & lNumeroFactura & " AS NumeroFactura, FACTURAS_DETALLE.Renglon, FACTURAS_DETALLE.IdConcepto, FACTURAS_DETALLE.IdMember, FACTURAS_DETALLE.NumeroFamiliar, FACTURAS_DETALLE.IdTipoUsuario, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.FormaPago, FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Cantidad, FACTURAS_DETALLE.Importe, FACTURAS_DETALLE.Intereses, FACTURAS_DETALLE.DescuentoPorciento, FACTURAS_DETALLE.Descuento, FACTURAS_DETALLE.Total, FACTURAS_DETALLE.IvaPorciento, FACTURAS_DETALLE.Iva, FACTURAS_DETALLE.IvaIntereses, FACTURAS_DETALLE.IvaDescuento, FACTURAS_DETALLE.TipoCargo, FACTURAS_DETALLE.Auxiliar, FACTURAS_DETALLE.IdInstructor, FACTURAS_DETALLE.Unidad"
    strSQL = strSQL & " From FACTURAS_DETALLE INNER JOIN FACTURAS ON FACTURAS_DETALLE.NumeroFactura=FACTURAS.NumeroFactura"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((FACTURAS.NumeroFactura) =" & lFacActual & ")"
    strSQL = strSQL & " )"
    
    adocmdFactura.CommandText = strSQL
    adocmdFactura.Execute
    
    
    
    
    'Me.StatusBar1.Panels("Proceso").Text = "Insertando pagos"
    'Me.StatusBar1.Panels("Tarea").Text = "Recibo # " & lRowPagos
    
    
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
    strSQL = strSQL & 1 & ", "
    strSQL = strSQL & 11 & ","
    strSQL = strSQL & "'" & "'" & ","
    strSQL = strSQL & doTotalFac & ","
    strSQL = strSQL & "'" & "',"
    strSQL = strSQL & 0 & ","
    strSQL = strSQL & "'" & "',"
    strSQL = strSQL & "'" & "',"
    strSQL = strSQL & doTotalFac & ","
    #If SqlServer_ Then
        strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "')"
    #Else
        strSQL = strSQL & "#" & Format(Date, "mm/dd/yyyy") & "#)"
    #End If
                
    adocmdFactura.CommandText = strSQL
    adocmdFactura.Execute
        
    
    'Me.StatusBar1.Panels("Proceso").Text = "Marcando recibos como facturados"
    'Me.StatusBar1.Panels("Tarea").Text = ""
    
    
    'Crea el registro en Facturas_Cancela
    
    'Me.StatusBar1.Panels("Proceso").Text = "Creando registro de cancelación"
    'Me.StatusBar1.Panels("Tarea").Text = ""
    
    strSQL = "INSERT INTO FACTURAS_CANCELA ("
    strSQL = strSQL & " NumeroFactura,"
    strSQL = strSQL & " CadenaCancela1,"
    strSQL = strSQL & " CadenaCancela2)"
    strSQL = strSQL & " SELECT " & lNumeroFactura & " AS NumeroFactura, CadenaCancela1, CadenaCancela2"
    strSQL = strSQL & " FROM FACTURAS_CANCELA"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((NumeroFactura)=" & lFacActual & ")"
    strSQL = strSQL & ")"
    
    adocmdFactura.CommandText = strSQL
    adocmdFactura.Execute
    
    
    Conn.CommitTrans
    
    'Me.StatusBar1.Panels("Proceso").Text = "Terminado"
    'Me.StatusBar1.Panels("Tarea").Text = "Se creo la factura " & lNumeroFactura
    
    iInitTrans = 0
    
    
    Set adocmdFactura = Nothing
    
    
    
    
       
    
    Screen.MousePointer = vbDefault
    
    'Para imprimir la factura
    
    
    'lNumFacIniImp = lNumeroFactura
    'lNumFacFinImp = lNumeroFactura
   
    'lNumFolioFacIniImp = sSerieFactura & lNumeroFolioFactura
    'lNumFolioFacFinImp = sSerieFactura & lNumeroFolioFactura
    
    
    'Dim frmImp As New frmImpFac
    
    'frmImp.cModo = "F"
    'frmImp.Tag = "F"
    
    'frmImp.lNumeroInicial = lNumeroFactura
    'frmImp.lNumeroFinal = lNumeroFactura
    
    'frmImp.Show 1
    
    GeneraFactura = lNumeroFactura
    
    Exit Function

Error_Catch:
    
    If iInitTrans Then
        Conn.RollbackTrans
    End If
    
    Screen.MousePointer = vbDefault
    
    MsgError
    
End Function

Private Function GeneraNCdesdeFactura(lNumFactura As Long) As Boolean
    
    Dim lNumNotaCred As Long
    
    Dim adorcs As ADODB.Recordset
    Dim adocmd As ADODB.Command
    
    Dim sSerieCFD As String
    Dim sFolioCFD As String
    
    GeneraNCdesdeFactura = False
    
    
    
    strSQL = "SELECT Max(NumeroNota) As Ultimo"
    strSQL = strSQL & " FROM NOTAS_CRED"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly

    If Not adorcs.EOF Then
        lNumNotaCred = adorcs!Ultimo + 1
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    
    
    'Inserta el encabezado de la nota de crédito
    strSQL = "INSERT INTO NOTAS_CRED ( NumeroNota, NumeroFactura, IdTitular, NoFamilia, FechaNota, HoraNota, NombreNota, CalleNota, ColoniaNota, DelNota, CiudadNota, EstadoNota, CodPos, RFC, Tel1, Total, USUARIO, Direccionado )"
    strSQL = strSQL & "SELECT  "
    strSQL = strSQL & lNumNotaCred & " AS NumeroNota" & ","
    strSQL = strSQL & "FACTURAS.NumeroFactura, FACTURAS.IdTitular, FACTURAS.NoFamilia,"
    #If SqlServer_ Then
        strSQL = strSQL & "'" & Format(Now, "yyyymmdd") & "',"
    #Else
        strSQL = strSQL & "#" & Format(Now, "mm/dd/yyyy") & "#,"
    #End If
    strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "',"
    strSQL = strSQL & "FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Total,"
    strSQL = strSQL & "'" & Trim(sDB_User) & "' AS USUARIO,"
    strSQL = strSQL & "FACTURAS.Direccionado"
    strSQL = strSQL & " From FACTURAS"
    strSQL = strSQL & " WHERE FACTURAS.NumeroFactura=" & lNumFactura
    
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    
    
    'Inserta el detalle de la nota de crédito
    strSQL = "INSERT INTO NOTAS_CRED_DETALLE ( NumeroNota, Renglon, IdConcepto, IdMember, NumeroFamiliar, IdTipoUsuario, Periodo, FormaPago, Concepto, Cantidad, Importe, Intereses, DescuentoPorciento, Descuento, Total, IvaPorciento, Iva, IvaIntereses, IvaDescuento, TipoCargo, Auxiliar, IdInstructor, Unidad )"
    strSQL = strSQL & "SELECT "
    strSQL = strSQL & lNumNotaCred & " AS NumeroNota" & ","
    strSQL = strSQL & "FACTURAS_DETALLE.Renglon, FACTURAS_DETALLE.IdConcepto, FACTURAS_DETALLE.IdMember, FACTURAS_DETALLE.NumeroFamiliar, FACTURAS_DETALLE.IdTipoUsuario, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.FormaPago, FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Cantidad, FACTURAS_DETALLE.Importe, FACTURAS_DETALLE.Intereses, FACTURAS_DETALLE.DescuentoPorciento, FACTURAS_DETALLE.Descuento, FACTURAS_DETALLE.Total, FACTURAS_DETALLE.IvaPorciento, FACTURAS_DETALLE.Iva, FACTURAS_DETALLE.IvaIntereses, FACTURAS_DETALLE.IvaDescuento, FACTURAS_DETALLE.TipoCargo, FACTURAS_DETALLE.Auxiliar, FACTURAS_DETALLE.IdInstructor, FACTURAS_DETALLE.Unidad"
    strSQL = strSQL & " From FACTURAS_DETALLE"
    strSQL = strSQL & " WHERE FACTURAS_DETALLE.NumeroFactura=" & lNumFactura

    adocmd.CommandText = strSQL
    adocmd.Execute
    
    
    Set adocmd = Nothing
    
    
    Me.StatusBar1.Panels("Proceso").Text = "Generando CFD en " & ObtieneParametro("URL_WS_CFD")
    
    sSerieCFD = ObtieneParametro("SERIE_CFD_NOTA_CREDITO")
    
    sFolioCFD = GeneraNotaCreditoCFD(lNumNotaCred, sSerieCFD, "egreso")
    
    If Len(sFolioCFD) > 12 Then
        MsgBox "Ocurrio un error generando el CFD" & vbCrLf & sFolioCFD, vbCritical, "Error"
    Else
        If sFolioCFD <> vbNullString Then
            Me.StatusBar1.Panels("Proceso").Text = "Actualizando FolioCFD"
            DoEvents
            If ActualizaFolioCFD(lNumNotaCred, sFolioCFD, sSerieCFD, "N") = 0 Then
            End If
        End If
    End If
    
    Me.StatusBar1.Panels("Proceso").Text = "Terminado"
    Me.StatusBar1.Panels("Tarea").Text = "Se creo la N. de crédito " & lNumNotaCred
    
    
    
    
    MsgBox "Nota de crédito Generada", vbInformation, "Ok"
    
    GeneraNCdesdeFactura = True

End Function

Private Sub txtNoFactura_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
            SendKeys vbTab
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
