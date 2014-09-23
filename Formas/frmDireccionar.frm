VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmDireccionar 
   Caption         =   "Direccionar pagos"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "frmDireccionar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Height          =   735
      Left            =   8520
      Picture         =   "frmDireccionar.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   735
      Left            =   9240
      Picture         =   "frmDireccionar.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame frmTipoDireccionado 
      Caption         =   "  Tipo de direccionado  "
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   3240
      TabIndex        =   7
      Top             =   4320
      Width           =   3615
      Begin VB.OptionButton optTD 
         Caption         =   "Tarjeta de débito"
         Height          =   255
         Left            =   1920
         TabIndex        =   9
         Top             =   320
         Width           =   1575
      End
      Begin VB.OptionButton optTC 
         Caption         =   "Tarjeta de crédito"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   320
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   5520
      MaxLength       =   8
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpFechaAlta 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   102105089
      CurrentDate     =   38663
   End
   Begin VB.CommandButton cmdCancelar 
      Enabled         =   0   'False
      Height          =   735
      Left            =   8520
      Picture         =   "frmDireccionar.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdBorrar 
      Height          =   735
      Left            =   7800
      Picture         =   "frmDireccionar.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdAgregar 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7800
      Picture         =   "frmDireccionar.frx":12EA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdModificar 
      Height          =   735
      Left            =   7080
      Picture         =   "frmDireccionar.frx":15F4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtNombre 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   60
      TabIndex        =   3
      Top             =   3960
      Width           =   5175
   End
   Begin VB.CheckBox chkActivo 
      Caption         =   "Activo"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtTitular 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   3360
      Width           =   5535
   End
   Begin VB.TextBox txtNoTit 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   975
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbDireccionados 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _Version        =   196616
      DataMode        =   2
      Cols            =   8
      Col.Count       =   8
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   17383
      _ExtentY        =   5106
      _StockProps     =   79
      Caption         =   "Datos para direccionar pagos"
   End
   Begin VB.CommandButton cmdNuevo 
      Height          =   735
      Left            =   7080
      Picture         =   "frmDireccionar.frx":1A36
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdNuevoConvenio 
      Height          =   735
      Left            =   7080
      Picture         =   "frmDireccionar.frx":2300
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblImporte 
      Caption         =   "Importe"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   20
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblFechaAlta 
      Caption         =   "Fecha de alta"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label lblTitular 
      Caption         =   "Titular"
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblNoTit 
      Caption         =   "# Titular"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   975
   End
End
Attribute VB_Name = "frmDireccionar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de pagos direccionados          *
'*  Daniel Hdez                                                 *
'*  07 / Noviembre / 2004                                       *
'*  Ultima actualización: 14 / Noviembre / 2005                 *
'****************************************************************

Const DATOSDIRS = 8

Public nTitular As Long
Public sTitular As String

Dim mAncDir(DATOSDIRS) As Integer
Dim mEncDir(DATOSDIRS) As String
Dim sNombre As String
Dim nImporte As Double
Dim nActivo As Byte
Dim dFechaAlta As Date
Dim bTarjCredito As Boolean
Dim bNuevo As Boolean
Dim nPosIni As Variant
Dim sTextMain As String
Dim frmHDir As frmayuda

Private Sub LlenaDirs()
    Dim rsDirs As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String


    Me.ssdbDireccionados.RemoveAll
    
    strSQL = "SELECT Direccionados.Nombre, Direccionados.Activo, Direccionados.FechaAlta, Direccionados.Importe, "
    strSQL = strSQL & " Direccionados.TipoDireccionado,Usuarios_Club.Nombre + ' ' + Usuarios_Club.A_Paterno + ' ' + Usuarios_Club.A_Materno  AS NOMBRETITULAR, Direccionados.idMember, "
    strSQL = strSQL & " Direccionados.idReg FROM Usuarios_Club inner join Direccionados on Direccionados.idmember=Usuarios_Club.idmember"
    strSQL = strSQL & " WHERE Direccionados.idMember='" & nTitular & "'"
    strSQL = strSQL & " ORDER BY Direccionados.Nombre"

    
    Set rsDirs = New ADODB.Recordset
    
'    rsDirs.ActiveConnection = Conn
'    rsDirs.LockType = adLockReadOnly
'    rsDirs.CursorType = adOpenForwardOnly
'    rsDirs.CursorLocation = adUseServer
'    rsDirs.Open strSQL
With rsDirs
        .Source = strSQL
        .ActiveConnection = Conn
        .CursorType = adOpenStatic
        .CursorLocation = adUseServer
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With
'    sCampos = "Direccionados.Nombre, Direccionados.Activo, "
'    sCampos = sCampos & "Direccionados.FechaAlta, Direccionados.Importe, "
'    sCampos = sCampos & "Direccionados.TipoDireccionado, "
'
'    #If SqlServer_ Then
'        sCampos = sCampos & "(SELECT (Usuarios_Club.Nombre + ' ' + Usuarios_Club.A_Paterno + ' ' + Usuarios_Club.A_Materno) FROM Usuarios_Club "
'    #Else
'        sCampos = sCampos & "(SELECT (Usuarios_Club.Nombre & ' ' & Usuarios_Club.A_Paterno & ' ' & Usuarios_Club.A_Materno) FROM Usuarios_Club "
'    #End If
'
'    sCampos = sCampos & "WHERE Direccionados.idMember=Usuarios_Club.idMember) AS NOMBRETITULAR, "
'
'    sCampos = sCampos & "Direccionados.idMember, Direccionados.idReg "
'
'    sTablas = "Direccionados"
'    sTablas = sTablas & ""
'
 '   InitRecordSet rsDirs, sCampos, sTablas, "Direccionados.idMember=" & nTitular, "Direccionados.Nombre", Conn
    With rsDirs
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                Me.ssdbDireccionados.AddItem .Fields("Nombre") & vbTab & _
                IIf(.Fields("Activo"), 1, 0) & vbTab & _
                Format(.Fields("FechaAlta"), "dd / mmm / yyyy") & vbTab & _
                Format(Round(.Fields("Importe"), 2), "###,###,##0.#0") & vbTab & _
                .Fields("TipoDireccionado") & vbTab & _
                .Fields(5) & vbTab & _
                .Fields("idMember") & vbTab & _
                .Fields("idReg")
            
                .MoveNext
            Loop
            
'            MuestraTit
        End If
    
        .Close
    End With
    Set rsDirs = Nothing
    
    If (Me.ssdbDireccionados.Rows <= 0) Then
        Me.ssdbDireccionados.Enabled = False
    Else
        Me.ssdbDireccionados.Bookmark = Me.ssdbDireccionados.AddItemBookmark(Me.ssdbDireccionados.AddItemRowIndex(nPosIni))
    End If
    
    If (Me.ssdbDireccionados.Enabled) Then
        Me.ssdbDireccionados.SetFocus
    End If
End Sub


Private Sub InitssdbDirs()
    'Asigna valores a la matriz de encabezados
    mEncDir(0) = "Nombre"
    mEncDir(1) = "Activo"
    mEncDir(2) = "Fecha alta"
    mEncDir(3) = "Importe"
    mEncDir(4) = "Direccionado a"
    mEncDir(5) = "Titular"
    mEncDir(6) = "# Titular"
    mEncDir(7) = "# Registro"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid Me.ssdbDireccionados, mEncDir
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncDir(0) = 4270
    mAncDir(1) = 700
    mAncDir(2) = 1400
    mAncDir(3) = 1200
    mAncDir(4) = 1700
    mAncDir(5) = 4500
    mAncDir(6) = 1000
    mAncDir(7) = 1000

    'Asigna el ancho de cada columna
    DefAnchossGrid Me.ssdbDireccionados, mAncDir
    
    Me.ssdbDireccionados.Columns(2).Alignment = ssCaptionAlignmentCenter
    Me.ssdbDireccionados.Columns(3).Alignment = ssCaptionAlignmentRight
    Me.ssdbDireccionados.Columns(4).Alignment = ssCaptionAlignmentCenter
    
    Me.ssdbDireccionados.Columns(1).Style = ssStyleCheckBox
    
    Me.ssdbDireccionados.Columns(5).Visible = False
    Me.ssdbDireccionados.Columns(6).Visible = False
    Me.ssdbDireccionados.Columns(7).Visible = False
End Sub

Private Sub cmdAgregar_Click()
    If (Cambios) Then
        If (GuardaDatos) Then
            bNuevo = True
            ClrCtrls
            InitVars
            
            LlenaDirs
        End If
    End If
    
    Me.txtNombre.SetFocus
End Sub

Private Function Cambios() As Boolean
    Cambios = True
    
    If (sNombre <> Trim$(Me.txtNombre.Text)) Then
        Exit Function
    End If
    
    If (nImporte <> Val(Me.txtImporte.Text)) Then
        Exit Function
    End If
    
    If (nActivo <> Me.chkActivo.Value) Then
        Exit Function
    End If
    
    If (dFechaAlta <> Me.dtpFechaAlta.Value) Then
        Exit Function
    End If
    
    If (bTarjCredito <> Me.optTC.Value) Then
        Exit Function
    End If
    
    Cambios = False
End Function

Private Function GuardaDatos() As Boolean
    Const DATOSDIR = 7
    Dim mFields(DATOSDIR) As String
    Dim mValues(DATOSDIR) As Variant
    Dim sCondicion As String

    GuardaDatos = False

    If (DatosCorrectos) Then
        mFields(0) = "idReg"
        mFields(1) = "idMember"
        mFields(2) = "FechaAlta"
        mFields(3) = "TipoDireccionado"
        mFields(4) = "Nombre"
        mFields(5) = "Importe"
        mFields(6) = "Activo"
        
        #If SqlServer_ Then
            mValues(1) = Val(Me.txtNoTit.Text)
            mValues(2) = Format(Me.dtpFechaAlta.Value, "yyyymmdd")
            mValues(3) = IIf(Me.optTC.Value, "TC", "TD")
            mValues(4) = Trim$(UCase$(Me.txtNombre.Text))
            mValues(5) = IIf(Val(Me.txtImporte.Text) > 0, CLng(Me.txtImporte.Text), 0)
            mValues(6) = Me.chkActivo.Value
        #Else
            mValues(1) = Val(Me.txtNoTit.Text)
            mValues(2) = Format(Me.dtpFechaAlta.Value, "dd/mm/yyyy")
            mValues(3) = IIf(Me.optTC.Value, "TC", "TD")
            mValues(4) = Trim$(UCase$(Me.txtNombre.Text))
            mValues(5) = IIf(Val(Me.txtImporte.Text) > 0, CLng(Me.txtImporte.Text), 0)
            mValues(6) = Me.chkActivo.Value
        #End If
        
        If (bNuevo) Then
            mValues(0) = LeeUltReg("Direccionados", "idReg") + 1
            
            If (AgregaRegistro("Direccionados", mFields, DATOSDIR, mValues, Conn)) Then
                GuardaDatos = True
                bNuevo = False
            Else
                MsgBox "No se agregó la información.", vbCritical, "KalaSystems"
            End If
        Else
            mValues(0) = Val(Me.ssdbDireccionados.Columns(7).Value)
            sCondicion = "idReg=" & Val(Me.ssdbDireccionados.Columns(7).Value)
        
            If (CambiaReg("Direccionados", mFields, DATOSDIR, mValues, sCondicion, Conn)) Then
                GuardaDatos = True
            Else
                MsgBox "No se realizaron los cambios.", vbCritical, "KalaSystems"
            End If
        End If
    End If
End Function

Private Function DatosCorrectos() As Boolean
    DatosCorrectos = False

    If (Trim$(Me.txtNombre.Text) = "") Then
        MsgBox "Se debe escribir un nombre.", vbExclamation, "KalaSystems"
        Me.txtNombre.SetFocus
        Exit Function
    End If
    
    If (Val(Me.txtImporte.Text) <= 0) Then
        MsgBox "El importe debe ser mayor a cero.", vbExclamation, "KalaSystems"
        Me.txtImporte.SetFocus
        Exit Function
    End If
    
'    If (Trim$(Me.txtTipoDireccionado.Text) = "") Then
'        MsgBox "Se debe escribir el tipo de direccionado.", vbExclamation, "KalaSystems"
'        Me.txtTipoDireccionado.SetFocus
'        Exit Function
'    End If

    DatosCorrectos = True
End Function

Private Sub cmdBorrar_Click()
    Dim nOk As Long

    If Not ChecaSeguridad(Me.Name, "cmdBorrar") Then
        Exit Sub
    End If

    If (Me.ssdbDireccionados.Rows > 0) Then
        nOk = MsgBox("¿Desea borrar el Direccionado seleccionado?", vbYesNo, "KalaSystems")

        If (nOk = vbYes) Then
            If (EliminaReg("Direccionados", "idReg=" & Me.ssdbDireccionados.Columns(7).Value, "", Conn)) Then
                LlenaDirs
            Else
                MsgBox "No se eliminó el registro.", vbInformation, "KalaSystems"
            End If
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    ClrCtrls
    ActivaCtrls False
    CtrlsEdicion True
    
    If (Me.ssdbDireccionados.Rows > 0) Then
        Me.ssdbDireccionados.SetFocus
    End If
End Sub

Private Sub cmdImprime_Click()
    If Not ChecaSeguridad(Me.Name, "cmdImprime") Then
        Exit Sub
    End If
    
    Dim frmImpfto As frmReportViewer
    Dim sQry As String
    Dim sRazonSocial As String
    
    If Me.ssdbDireccionados.Rows = 0 Then
        Exit Sub
    End If
    
    If CDate(Me.ssdbDireccionados.Columns(2).Value) <> Date Then
        MsgBox "Ya no es posible imprimir este formato", vbExclamation + vbOKOnly, "Verifique"
        
        Exit Sub
    End If
    
    sRazonSocial = ObtieneParametro("EMPRESA CONTRATO")
    
    sQry = "SELECT DireccionadosDatos.IdConvenio, DireccionadosDatos.FechaAlta, DireccionadosDatos.NoInscripcion, DireccionadosDatos.IdMember, DireccionadosDatos.TipoDireccionado, DireccionadosDatos.NumeroCuenta, DireccionadosDatos.NumeroCuentaVisible, DireccionadosDatos.BancoEmisor, DireccionadosDatos.Nombre, DireccionadosDatos.A_Paterno, DireccionadosDatos.A_Materno, DireccionadosDatos.FechaExpedicion, DireccionadosDatos.FechaVencimiento, DireccionadosDatos.CodigoSeguridad, DireccionadosDatos.OperadorTarjeta, DireccionadosDatos.Importe, DireccionadosDatos.Activo, DireccionadosDatos.Email, DireccionadosDatos.Telefono, DireccionadosDatos.AplicaAPartir, DireccionadosDatos.UsuarioAlta, USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno" & ",'" & sRazonSocial & "' AS RazonSocial"
    sQry = sQry & " FROM DireccionadosDatos INNER JOIN USUARIOS_CLUB ON DireccionadosDatos.IdMember=USUARIOS_CLUB.IdMember"
    sQry = sQry & " Where ("
    sQry = sQry & "((DireccionadosDatos.IdConvenio)=" & Me.ssdbDireccionados.Columns(7).Value & ")"
    sQry = sQry & ")"
    
    Set frmImpfto = New frmReportViewer
    
    frmImpfto.sNombreReporte = sDB_ReportSource & "\fto_alta_direccionado.rpt"
    frmImpfto.sQuery = sQry
    
    frmImpfto.Show vbModal
    
End Sub

Private Sub cmdModificar_Click()
    If Not ChecaSeguridad(Me.Name, "cmdModificar") Then
        Exit Sub
    End If
    
    If (Me.ssdbDireccionados.Rows > 0) Then
        bNuevo = False
        nPosIni = Me.ssdbDireccionados.Bookmark
    
        LeeDatos
        
        CtrlsEdicion False
        ActivaCtrls True
        InitVars
                
        Me.txtNombre.SetFocus
    End If
End Sub

Private Sub cmdNuevo_Click()
    If Not ChecaSeguridad(Me.Name, "cmdNuevo") Then
        Exit Sub
    End If
    
    bNuevo = True

    CtrlsEdicion False
    ActivaCtrls True
    ClrCtrls
    
    InitVars

    Me.txtNombre.SetFocus
End Sub

Private Sub cmdNuevoConvenio_Click()
    If Not ChecaSeguridad(Me.Name, "cmdNuevoConvenio") Then
        Exit Sub
    End If
    
    Dim frmConvenio As frmDireccionarNuevo
    
    Set frmConvenio = New frmDireccionarNuevo
    
    frmConvenio.lTitular = nTitular
    
    frmConvenio.Show vbModal
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.dtpFechaAlta.Value = Date
    InitssdbDirs
    LlenaDirs
End Sub

Private Sub Form_Load()
    nPosIni = 0
    sTextMain = MDIPrincipal.StatusBar1.Panels.Item(1).Text
    
    Me.txtNoTit.Text = nTitular
    Me.txtTitular.Text = sTitular
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Direccionar pagos"
    
    #If SqlServer_ Then
        Dim strDireccionar As String
        strDireccionar = ObtieneParametro("DIRECCIONAR_NUEVO")
        
        If strDireccionar <> "" Then
            bDireccionado = CBool(strDireccionar)
        Else
            bDireccionado = False
        End If
    #End If
    
    If bDireccionado = True Then
        cmdNuevo.Visible = False
        cmdNuevoConvenio.Visible = True
        cmdImprime.Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDireccionar = Nothing
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextMain
End Sub

Private Sub ActivaCtrls(bValor As Boolean)
    With Me
        .txtNombre.Enabled = bValor
        .txtImporte.Enabled = bValor
        .chkActivo.Enabled = bValor
        .dtpFechaAlta.Enabled = bValor
        .frmTipoDireccionado.Enabled = bValor
    End With
End Sub

Private Sub CtrlsEdicion(bValor As Boolean)
    With Me
        .ssdbDireccionados.Enabled = bValor
        .cmdNuevo.Enabled = bValor
        .cmdModificar.Enabled = bValor
        .cmdBorrar.Enabled = bValor
        .cmdAgregar.Enabled = Not bValor
        .cmdCancelar.Enabled = Not bValor
    End With
End Sub

Private Sub ClrCtrls()
    With Me
        .txtNombre.Text = Me.txtTitular.Text
        .txtImporte.Text = ""
        .chkActivo.Value = 0
        .dtpFechaAlta.Value = Date
        .optTC.Value = True
    End With
End Sub

Private Sub InitVars()
    sNombre = Trim$(Me.txtNombre.Text)
    
    nImporte = 0
    If (Val(Me.txtImporte.Text) > 0) Then
        nImporte = CDbl(Me.txtImporte.Text)
    End If
    
    nActivo = Me.chkActivo.Value
    dFechaAlta = Me.dtpFechaAlta.Value
    bTarjCredito = Me.optTC.Value
End Sub

Private Sub LeeDatos()
    With Me.ssdbDireccionados
        Me.txtNombre.Text = .Columns(0).Text
        Me.txtImporte.Text = .Columns(3).Value
        Me.chkActivo.Value = Val(.Columns(1).Value)
        Me.dtpFechaAlta.Value = .Columns(2).Value
        Me.optTC.Value = IIf(.Columns(4).Text = "TC", True, False)
        Me.optTD.Value = Not Me.optTC.Value
    End With
End Sub
