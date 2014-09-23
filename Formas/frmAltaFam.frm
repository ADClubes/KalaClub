VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmAltaFam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de familiares"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "frmAltaFam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10365
   Begin VB.Frame frmFamiliar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10005
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbParentesco 
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Width           =   2175
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
         Columns(0).Caption=   "Parentesco"
         Columns(0).Name =   "Parentesco"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1244
         Columns(1).Caption=   "Clave"
         Columns(1).Name =   "Clave"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   4800
         TabIndex        =   50
         Top             =   5160
         Width           =   255
      End
      Begin VB.TextBox txtUltPago 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   47
         Top             =   5160
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpFechaInicio 
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Top             =   5160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   105316353
         CurrentDate     =   39175
      End
      Begin VB.TextBox txtNumFam 
         Height          =   285
         Left            =   3960
         TabIndex        =   45
         Top             =   6000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtImagen 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   14
         Top             =   6000
         Width           =   855
      End
      Begin VB.Frame frmImagen 
         Height          =   3615
         Left            =   6960
         TabIndex        =   24
         Top             =   960
         Width           =   2775
         Begin VB.Image imgFoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   3255
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdHTipo 
         Height          =   305
         Left            =   930
         Picture         =   "frmAltaFam.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   " Tipos de usuario disponibles "
         Top             =   4320
         Width           =   425
      End
      Begin VB.CommandButton cmdHPais 
         Height          =   305
         Left            =   930
         Picture         =   "frmAltaFam.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   " Lista de países "
         Top             =   3600
         Width           =   425
      End
      Begin VB.TextBox txtCveTipoFam 
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   21
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox txtPaisFam 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   17
         Top             =   3600
         Width           =   4335
      End
      Begin VB.TextBox txtTipoFam 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   4320
         Width           =   4215
      End
      Begin VB.TextBox txtCvePaisFam 
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   15
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox txtFamilia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   6000
         Width           =   855
      End
      Begin VB.TextBox txtSec 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   6000
         Width           =   855
      End
      Begin VB.TextBox txtTitFam 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   6000
         Width           =   855
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   6720
         MaxLength       =   60
         TabIndex        =   3
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtPaterno 
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtEdad 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtMaterno 
         Height          =   285
         Left            =   3360
         MaxLength       =   60
         TabIndex        =   2
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtCve 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         MaxLength       =   5
         TabIndex        =   4
         Top             =   6000
         Width           =   615
      End
      Begin VB.TextBox txtProf 
         Height          =   285
         Left            =   2400
         MaxLength       =   60
         TabIndex        =   8
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   120
         MaxLength       =   60
         TabIndex        =   10
         Top             =   2880
         Width           =   4455
      End
      Begin VB.TextBox txtCel 
         Height          =   285
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Frame frmSexo 
         Caption         =   " Sexo "
         Height          =   855
         Left            =   4920
         TabIndex        =   18
         Top             =   2520
         Width           =   1575
         Begin VB.OptionButton optMasculino 
            Caption         =   "Masculino"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   520
            Width           =   1095
         End
         Begin VB.OptionButton optFemenino 
            Caption         =   "Femenino"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdGuardar 
         Height          =   795
         Left            =   8880
         Picture         =   "frmAltaFam.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   " Guardar registro "
         Top             =   4680
         Width           =   795
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   675
         Left            =   8880
         Picture         =   "frmAltaFam.frx":09E0
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   " Salir "
         Top             =   5640
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtpRegistro 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   105316353
         CurrentDate     =   37995
      End
      Begin MSComCtl2.DTPicker dtpNacio 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   105316353
         CurrentDate     =   37995
      End
      Begin VB.Label lblParen 
         Caption         =   "Parentesco"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Mantenimiento Pagado Hasta"
         Height          =   255
         Left            =   2400
         TabIndex        =   49
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Inicio. de uso de Inst."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label lblImagen 
         Caption         =   "# imagen"
         Height          =   255
         Left            =   3000
         TabIndex        =   44
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label lblTipoFam 
         Caption         =   "Cve. tipo"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lblPaisFam 
         Caption         =   "Cve. país"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label lblFamilia 
         Caption         =   "# Familia"
         Height          =   255
         Left            =   2040
         TabIndex        =   41
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label lblSec 
         Caption         =   "# Sec."
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Top             =   5760
         Width           =   735
      End
      Begin VB.Label lblCveTit 
         Caption         =   "Cve. titular"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre(s)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6720
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblMaterno 
         Caption         =   "Apellido materno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3360
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblPaterno 
         Caption         =   "Apellido paterno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblEdad 
         Caption         =   "Edad"
         Height          =   255
         Left            =   1680
         TabIndex        =   35
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblNacio 
         Caption         =   "Nació"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblRegistro 
         Caption         =   "Inscripción"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   33
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblCve 
         Caption         =   "Cve"
         Height          =   255
         Left            =   5040
         TabIndex        =   32
         Top             =   5760
         Width           =   495
      End
      Begin VB.Label lblProf 
         Caption         =   "Profesión"
         Height          =   255
         Left            =   2400
         TabIndex        =   31
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label lblCel 
         Caption         =   "Celular"
         Height          =   255
         Left            =   4920
         TabIndex        =   29
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblTipoUsuario 
         Caption         =   "Tipo de usuario"
         Height          =   255
         Left            =   1440
         TabIndex        =   28
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblPais 
         Caption         =   "País de origen"
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   3360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAltaFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de familiares                   *
'*  Daniel Hdez                                                 *
'*  02 / Septiembre / 2004                                      *
'*  Ult Act: 03 / Septiembre / 2005                             *
'****************************************************************


Public bNvoFam As Boolean
Public nCveFam As Integer
Public nAyuda As Byte

Dim sFormaAnt As String
Dim sTextToolBar As String
Dim sPaterno As String
Dim sMaterno As String
Dim sNombre As String
Dim sProf As String
Dim dNacio As Date
Dim dRegistro As Date
Dim sCel As String
Dim sEmail As String
Dim bFemenino As Boolean
Dim sTipoFam As String
Dim sPaisFam As String
Dim nTipoAnt As Integer
Dim nTipoNvo As Integer


Private Sub cmdSalir_Click()
Dim Respuesta As Integer

    If (Cambios) Then
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Registro de familiares")
        
        If (Respuesta = vbYes) Then
            If (GuardaDatos) Then
                Unload Me
            Else
                Exit Sub
            End If
        End If
    End If

    Unload Me
End Sub


Private Sub dtpNacio_Change()
    Me.txtEdad.Text = Format(Edad(Me.dtpNacio.Value), "#0.00")
    Me.txtCveTipoFam.Text = ""
    Me.txtTipoFam.Text = ""
End Sub

Private Sub Form_Load()
    CentraForma MDIPrincipal, frmAltaFam
    sTextToolBar = Trim(MDIPrincipal.StatusBar1.Panels.Item(1).Text)
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Registro de familiares"
    
    'Guarda el nombre de la forma inmediata anterior
    sFormaAnt = Forms(Forms.Count - 2).Name
    
    ClearCtrls
    
    If (Not bNvoFam) Then
        LeeDatos
    Else
        'Clave para el nuevo registro
        nCveFam = 0
    End If
    
    InitVar
    
End Sub


Private Sub dtpNacio_LostFocus()
'    Me.txtEdad.Text = Format(Edad(Me.dtpNacio.Value), "#0.00")
'    Me.txtCveTipoFam.Text = ""
'    Me.txtTipoFam.Text = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If (sFormaAnt = "frmAltaSocios") Then
        frmAltaFam.LlenaFam
    Else
        frmDatosSocios.Refresca
        frmDatosSocios.WindowState = 2
    End If
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub


Private Sub ClearCtrls()
    With frmAltaFam
        If (sFormaAnt = "frmAltaSocios") Then
            .txtTitFam.Text = Val(frmAltaSocios.txtTitCve.Text)
            .txtFamilia.Text = Val(frmAltaSocios.txtFamilia.Text)
        Else
            .txtTitFam.Text = ""
            .txtFamilia.Text = ""
        End If
        
        .txtTitFam.Enabled = False
        
        If (bNvoFam) Then
            .txtCve.Text = BuscaCve("Usuarios_Club", "IdMember")
        Else
            .txtCve.Text = ""
        End If
        
        .txtPaterno.Text = ""
        .txtMaterno.Text = ""
        .txtNombre.Text = ""
        .txtProf.Text = ""
        .dtpNacio.Value = Date
        .txtEdad.Text = ""
        .dtpRegistro.Value = Date
        .txtCel.Text = ""
        .txtEmail.Text = ""
        .txtCvePaisFam.Text = ""
        .txtPaisFam.Text = ""
        .txtCveTipoFam.Text = ""
        .txtTipoFam.Text = ""
        .optFemenino.Value = True
        
        '03/04/2007
        .dtpFechaInicio.Value = Date
        .txtUltPago.Text = ""
        
        .txtSec.Text = ""
        .txtImagen.Text = ""
        
        strSQL = "SELECT P.Parentesco, P.Clave"
        strSQL = strSQL & " From PARENTESCO P"
        strSQL = strSQL & " ORDER BY P.Parentesco"
        
        LlenaSsCombo Me.sscmbParentesco, Conn, strSQL, 2
        
        Me.sscmbParentesco.Text = ""

        
    End With
End Sub


Private Sub InitVar()
    sPaterno = Trim(Me.txtPaterno.Text)
    sMaterno = Trim(Me.txtMaterno.Text)
    sNombre = Trim(Me.txtNombre.Text)
    sProf = Trim(Me.txtProf.Text)
    dNacio = Me.dtpNacio.Value
    dRegistro = Me.dtpRegistro.Value
    sCel = Trim(Me.txtCel.Text)
    sEmail = Trim(Me.txtEmail.Text)
    
    If (Me.optFemenino.Value) Then
        bFemenino = True
    Else
        bFemenino = False
    End If
    
    sPaisFam = Trim(Me.txtPaisFam.Text)
    sTipoFam = Trim(Me.txtTipoFam.Text)
    nTipoAnt = Val(Me.txtCveTipoFam.Text)
    nTipoNvo = Val(Me.txtCveTipoFam.Text)
End Sub


Private Sub cmdGuardar_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdGuardar.Name) Then
        Exit Sub
    End If
    
    If (Cambios) Then
        If (GuardaDatos) Then
        
            If (sFormaAnt = "frmAltaSocios") Then
                bNvoFam = True
                ClearCtrls
            End If
        
            'Inicializa las variables
            InitVar
        Else
            MsgBox "No se registraron los datos, verifique la información.", vbCritical, "KalaSystems"
        End If
    End If
    
    Me.txtPaterno.SetFocus
End Sub


Private Function ChecaDatos()
    Dim sCond As String
    Dim sCamp As String

    ChecaDatos = False
    
    If (bNvoFam) Then
        If (IsNumeric(Me.txtCve.Text)) Then
            If (Val(Me.txtCve.Text) <= 0) Then
                MsgBox "La clave del familiar debe ser mayor que cero.", vbExclamation, "KalaSystems"
                Me.txtCve.SetFocus
                Exit Function
            End If
            
            If (Me.dtpRegistro.Value > Date) Then
                MsgBox "La Fecha de ingreso no puede ser mayor al dia de hoy!", vbExclamation, "Verifique"
                Me.dtpRegistro.SetFocus
                Exit Function
            End If
            
'            If (ExisteXValor("IdMember", "Usuarios_Club", "IdMember=" & Val(Me.txtCve.Text), Conn, "")) Then
'                MsgBox "La clave seleccionada ya se encuentra en uso.", vbExclamation, "KalaSystems"
'
'                If (Me.txtCve.Enabled) Then
'                    Me.txtCve.SetFocus
'                End If
'                Exit Function
'            End If
        Else
            MsgBox "La Clave seleccionada para el familiar es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtCve.SetFocus
            Exit Function
        End If
    End If
    
    If (Trim(Me.txtPaterno.Text) = "") Then
        MsgBox "El apellido paterno no puede quedar en blanco.", vbExclamation, "KalaSystems"
        Me.txtPaterno.SetFocus
        Exit Function
    End If
    
    If (Trim(Me.txtMaterno.Text) = "") Then
        MsgBox "El apellido materno no puede quedar en blanco.", vbExclamation, "KalaSystems"
        Me.txtMaterno.SetFocus
        Exit Function
    End If
    
    If (Trim(Me.txtNombre.Text) = "") Then
        MsgBox "El nombre no puede quedar en blanco.", vbExclamation, "KalaSystems"
        Me.txtNombre.SetFocus
        Exit Function
    End If
    
    '11/09/2007
    If Me.dtpNacio.Value >= Date Then
        MsgBox "La Fecha de nacimiento debe ser menor al dia de hoy!", vbExclamation, "Verifique"
        Me.dtpNacio.SetFocus
        Exit Function
    End If
    
    '18/04/08
    If Me.sscmbParentesco.Text = vbNullString Then
        MsgBox "Seleccionar el parentesco!", vbExclamation, "Verifique"
        Me.sscmbParentesco.SetFocus
        Exit Function
    End If
    
    If (Trim(Me.txtTipoFam.Text) = "") Then
        MsgBox "Se debe seleccionar un tipo de familiar.", vbExclamation, "KalaSystems"
        Me.txtCveTipoFam.SetFocus
        Exit Function
    End If
    
    If (Trim(Me.txtPaisFam.Text) = "") Then
        MsgBox "Se debe seleccionar un país de origen.", vbExclamation, "KalaSystems"
        Me.txtCvePaisFam.SetFocus
        Exit Function
    End If
    
    
'    If Me.dtpFechaInicio.Value < Date Then
'        MsgBox "La Fecha de inicio de uso de instalaciones debe ser mayor o igual a la fecha actual!", vbExclamation, "Verifique"
'        Me.dtpFechaInicio.SetFocus
'        Exit Function
'    End If
    
    If Not ChecaIntegrantesMembresia() Then
        Exit Function
    End If
    
    
    
    ChecaDatos = True
End Function


Private Function GuardaDatos() As Boolean
    Const DATOSFAMILIAR = 18
    Const DATOSALTA = 17
    Const DATOSTZONE = 4
    Dim mFieldsFam(DATOSFAMILIAR) As String
    Dim mValuesFam(DATOSFAMILIAR) As Variant
    Dim mFieldsAlta(DATOSALTA) As String
    Dim mValuesAlta(DATOSALTA) As Variant
    Dim mFieldsZone(DATOSTZONE) As String
    Dim mValuesZone(DATOSTZONE) As Variant
    Dim InitTrans As Long
    Dim nSecFam As Long
    Dim nSecFamNuevo As Long
    Dim nErrCode As Long
    Dim nTor As Long

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Usuarios_Club
    mFieldsFam(0) = "IdMember"
    mFieldsFam(1) = "Nombre"
    mFieldsFam(2) = "A_Paterno"
    mFieldsFam(3) = "A_Materno"
    mFieldsFam(4) = "FechaNacio"
    mFieldsFam(5) = "Sexo"
    mFieldsFam(6) = "IdPais"
    mFieldsFam(7) = "IdTipoUsuario"
    mFieldsFam(8) = "IdTitular"
    mFieldsFam(9) = "Email"
    mFieldsFam(10) = "Celular"
    mFieldsFam(11) = "Profesion"
    mFieldsFam(12) = "FechaIngreso"
    mFieldsFam(13) = "NoFamilia"
    mFieldsFam(14) = "FotoFile"
    mFieldsFam(15) = "NumeroFamiliar"
    'gpo 18/08/08
    mFieldsFam(16) = "Parentesco"
    mFieldsFam(17) = "UFechaPago"
    
    'Campos de la tabla Altas
    mFieldsAlta(0) = "IdMember"
    mFieldsAlta(1) = "Nombre"
    mFieldsAlta(2) = "A_Paterno"
    mFieldsAlta(3) = "A_Materno"
    mFieldsAlta(4) = "FechaNacio"
    mFieldsAlta(5) = "Sexo"
    mFieldsAlta(6) = "IdPais"
    mFieldsAlta(7) = "IdTipoUsuario"
    mFieldsAlta(8) = "IdTitular"
    mFieldsAlta(9) = "Email"
    mFieldsAlta(10) = "Celular"
    mFieldsAlta(11) = "Profesion"
    mFieldsAlta(12) = "FechaIngreso"
    mFieldsAlta(13) = "IdUsuario"
    mFieldsAlta(14) = "Fecha"
    mFieldsAlta(15) = "NoFamilia"
    mFieldsAlta(16) = "FotoFile"

    'Valores para la tabla Usuarios_Club
    #If SqlServer_ Then
        mValuesFam(0) = Val(Me.txtCve.Text)
        mValuesFam(1) = Trim(UCase(Me.txtNombre.Text))
        mValuesFam(2) = Trim(UCase(Me.txtPaterno.Text))
        mValuesFam(3) = Trim(UCase(Me.txtMaterno.Text))
        mValuesFam(4) = Format(Me.dtpNacio.Value, "yyyymmdd")
        mValuesFam(5) = IIf(Me.optFemenino.Value, "F", "M")
        mValuesFam(6) = Val(Me.txtCvePaisFam.Text)
        mValuesFam(7) = Val(Me.txtCveTipoFam.Text)
        mValuesFam(8) = Val(Me.txtTitFam.Text)
        mValuesFam(9) = Trim(Me.txtEmail.Text)
        mValuesFam(10) = Trim(UCase(Me.txtCel.Text))
        mValuesFam(11) = Trim(UCase(Me.txtProf.Text))
        mValuesFam(12) = Format(Me.dtpRegistro, "yyyymmdd")
        mValuesFam(13) = Val(Me.txtFamilia.Text)
        
        If (bNvoFam) Then
            mValuesFam(0) = LeeUltReg("Usuarios_Club", "IdMember") + 1
            mValuesFam(14) = LeeUltReg("FolioFoto", "idFoto") + 1
            mValuesFam(15) = LeeUltimoFamiliar(Val(Me.txtTitFam.Text)) + 1
        Else
            mValuesFam(14) = Val(Me.txtImagen.Text)
            mValuesFam(15) = Val(Me.txtNumFam.Text)
        End If
        
        mValuesFam(16) = Me.sscmbParentesco.Columns("Clave").Value
        
        mValuesFam(17) = Format(Me.dtpFechaInicio.Value, "yyyymmdd")
    #Else
        mValuesFam(0) = Val(Me.txtCve.Text)
        mValuesFam(1) = Trim(UCase(Me.txtNombre.Text))
        mValuesFam(2) = Trim(UCase(Me.txtPaterno.Text))
        mValuesFam(3) = Trim(UCase(Me.txtMaterno.Text))
        mValuesFam(4) = Format(Me.dtpNacio.Value, "dd/mm/yyyy")
        mValuesFam(5) = IIf(Me.optFemenino.Value, "F", "M")
        mValuesFam(6) = Val(Me.txtCvePaisFam.Text)
        mValuesFam(7) = Val(Me.txtCveTipoFam.Text)
        mValuesFam(8) = Val(Me.txtTitFam.Text)
        mValuesFam(9) = Trim(Me.txtEmail.Text)
        mValuesFam(10) = Trim(UCase(Me.txtCel.Text))
        mValuesFam(11) = Trim(UCase(Me.txtProf.Text))
        mValuesFam(12) = Format(Me.dtpRegistro, "dd/mm/yyyy")
        mValuesFam(13) = Val(Me.txtFamilia.Text)
    
        If (bNvoFam) Then
            mValuesFam(0) = LeeUltReg("Usuarios_Club", "IdMember") + 1
            mValuesFam(14) = LeeUltReg("FolioFoto", "idFoto") + 1
            mValuesFam(15) = LeeUltimoFamiliar(Val(Me.txtTitFam.Text)) + 1
        Else
            mValuesFam(14) = Val(Me.txtImagen.Text)
            mValuesFam(15) = Val(Me.txtNumFam.Text)
        End If
        
        mValuesFam(16) = Me.sscmbParentesco.Columns("Clave").Value
        
        mValuesFam(17) = Format(Me.dtpFechaInicio.Value, "dd/mm/yyyy")
    #End If
    
    'Valores para la tabla Altas
    #If SqlServer_ Then
        mValuesAlta(0) = mValuesFam(0)
        mValuesAlta(1) = mValuesFam(1)
        mValuesAlta(2) = mValuesFam(2)
        mValuesAlta(3) = mValuesFam(3)
        mValuesAlta(4) = mValuesFam(4)
        mValuesAlta(5) = mValuesFam(5)
        mValuesAlta(6) = mValuesFam(6)
        mValuesAlta(7) = mValuesFam(7)
        mValuesAlta(8) = mValuesFam(8)
        mValuesAlta(9) = mValuesFam(9)
        mValuesAlta(10) = mValuesFam(10)
        mValuesAlta(11) = mValuesFam(11)
        mValuesAlta(12) = mValuesFam(12)
        mValuesAlta(13) = LeeXValor("IdUsuario", "Usuarios_Sistema", "Login_Name='" & sDB_User & "'", "IdUsuario", "n", Conn)
        mValuesAlta(14) = Format(Date, "yyyymmdd")
        mValuesAlta(15) = mValuesFam(13)
        mValuesAlta(16) = mValuesFam(14)
    #Else
        mValuesAlta(0) = mValuesFam(0)
        mValuesAlta(1) = mValuesFam(1)
        mValuesAlta(2) = mValuesFam(2)
        mValuesAlta(3) = mValuesFam(3)
        mValuesAlta(4) = mValuesFam(4)
        mValuesAlta(5) = mValuesFam(5)
        mValuesAlta(6) = mValuesFam(6)
        mValuesAlta(7) = mValuesFam(7)
        mValuesAlta(8) = mValuesFam(8)
        mValuesAlta(9) = mValuesFam(9)
        mValuesAlta(10) = mValuesFam(10)
        mValuesAlta(11) = mValuesFam(11)
        mValuesAlta(12) = mValuesFam(12)
        mValuesAlta(13) = LeeXValor("IdUsuario", "Usuarios_Sistema", "Login_Name='" & sDB_User & "'", "IdUsuario", "n", Conn)
        mValuesAlta(14) = Format(Date, "dd/mm/yyyy")
        mValuesAlta(15) = mValuesFam(13)
        mValuesAlta(16) = mValuesFam(14)
    #End If

    If (bNvoFam) Then
        'Inicia el registro de los datos en las tablas
        InitTrans = Conn.BeginTrans
    
        'Registra los datos de la nueva direccion
        If (AgregaRegistro("Usuarios_Club", mFieldsFam, DATOSFAMILIAR, mValuesFam, Conn)) Then
        
            'Registra las fechas de inicio para cuotas de Mtto
            If (RegistraFechasMtto(CLng(mValuesFam(0)), Val(Me.txtCveTipoFam.Text), Me.dtpRegistro.Value)) Then
        
                'Registra los datos en la tabla de altas
                If (AgregaRegistro("Altas", mFieldsAlta, DATOSALTA, mValuesAlta, Conn)) Then
                
                    'Campos de la tabla Time_Zone_User
                    mFieldsZone(0) = "IdReg"
                    mFieldsZone(1) = "NoFamilia"
                    mFieldsZone(2) = "IdMember"
                    mFieldsZone(3) = "IdTimeZone"
                
                    'Valores para la tabla Time_Zone_Users
                    mValuesZone(0) = LeeUltReg("Time_Zone_Users", "IdReg") + 1
                    mValuesZone(1) = mValuesFam(13)
                    mValuesZone(2) = mValuesFam(0)
                    mValuesZone(3) = 0
                
                    'Registra los datos en las zonas
                    If (AgregaRegistro("Time_Zone_Users", mFieldsZone, DATOSTZONE, mValuesZone, Conn)) Then
                
                        'Incrementa el contador en la tabla de ALTAS
                        If (Altas_Bajas(True)) Then
                        
                            'Registra la clave del familiar en la tabla de secuenciales
                            nSecFam = AsignaSec(CLng(mValuesFam(0)), False)
                            nSecFamNuevo = AsignaSecNuevo(CLng(mValuesFam(0)), False)
                            
                            'Si se registro correctamente
                            If (nSecFam > 0 And nSecFamNuevo > 0) Then
                            
                                'Incrementa el ultimo numero de folio de las fotos
                                If (IncFolio("FolioFoto", "idFoto", 1)) Then
                                    'Baja a disco los nuevos datos
                                    Conn.CommitTrans
                                    
                                    
                                    
                                    'Muestra el numero del registro de la direccion
                                    Me.txtCve.Text = mValuesFam(0)
                                    Me.txtCve.REFRESH
                                    
                                    'Actualiza el numero de secuencial
                                    Me.txtSec.Text = nSecFam
                                    Me.txtSec.REFRESH
                                    
                                    'Actualiza el numero del fotofile
                                    Me.txtImagen.Text = mValuesFam(14)
                                    Me.txtImagen.REFRESH
                                    
                                    
                                                                        
                        
                                    bNvoFam = False
                                    GuardaDatos = True
                                    
                                    MsgBox "Los datos se dieron de alta correctamente. " & vbLf & "Favor de registrar la huella y/o Tarjeta del Usuario. ", vbInformation, "Ok"
                                    
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            'En caso de algun error no baja a disco los nuevos datos
            If InitTrans > 0 Then
                Conn.RollbackTrans
            End If
            
            MsgBox "El registro no fue completado.", vbCritical, "Error"
        End If
    Else

        If (Val(Me.txtCve.Text) > 0) Then
            InitTrans = Conn.BeginTrans
    
            'Actualiza los datos del familiar
            If (CambiaReg("Usuarios_Club", mFieldsFam, DATOSFAMILIAR, mValuesFam, "IdMember=" & Val(Me.txtCve.Text), Conn)) Then
                
                If (nTipoAnt <> nTipoNvo) Then
                    If (ActConcFact(CInt(mValuesFam(0)), nTipoAnt, nTipoNvo)) Then
                        
                        Conn.CommitTrans
                        MsgBox "Los datos se actualizaron correctamente.", vbInformation, "KalaSystems"
                        GuardaDatos = True
                
                        
                    Else
                        If InitTrans > 0 Then
                            Conn.RollbackTrans
                        End If
                        MsgBox "No se realizaron los cambios.", vbCritical, "KalaSystems"
                    End If
                Else
                    
                    Conn.CommitTrans
                
                    MsgBox "Los datos se actualizaron correctamente.", vbInformation, "KalaSystems"
                    GuardaDatos = True
                
                    
                End If
            Else
                If InitTrans > 0 Then
                    Conn.RollbackTrans
                End If
                MsgBox "No se realizaron los cambios.", vbCritical, "KalaSystems"
            End If
        End If

    End If
    'Registro en Torniquetes NITEGEN
            
            'nTor = (68 * 16777216) + CLng(Me.txtSec.Text)
            'Do While nErrCode <> 805372429
             'nErrCode = AgregaAcceso(CLng(mValuesFam(0)), CStr(mValuesFam(1) & " " & mValuesFam(2) & " " & mValuesFam(3)), nTor)
             
             
                'MsgBox "No se pudo registrar el usuario en torniquetes,Se va a volver a intentar..."
             'Loop
             
    If (sFormaAnt = "frmDatosSocios") Then
            
        'Pasa a mayusculas el contenido de las cajas de texto
        CambiaAMayusculas
    End If
End Function


'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    Cambios = True

    If (sPaterno <> Trim(Me.txtPaterno.Text)) Then
        Exit Function
    End If
    
    If (sMaterno <> Trim(Me.txtMaterno.Text)) Then
        Exit Function
    End If
    
    If (sNombre <> Trim(Me.txtNombre.Text)) Then
        Exit Function
    End If

    If (sProf <> Trim(Me.txtProf.Text)) Then
        Exit Function
    End If

    If (dNacio <> Me.dtpNacio.Value) Then
        Exit Function
    End If

    If (dRegistro <> Me.dtpRegistro.Value) Then
        Exit Function
    End If

    If (sCel <> Trim(Me.txtCel.Text)) Then
        Exit Function
    End If

    If (sEmail <> Trim(Me.txtEmail.Text)) Then
        Exit Function
    End If
    
    If (bFemenino <> Me.optFemenino.Value) Then
        Exit Function
    End If

    If (sTipoFam <> Trim(Me.txtTipoFam.Text)) Then
        nTipoNvo = Val(Me.txtCveTipoFam.Text)
        Exit Function
    End If

    If (sPaisFam <> Trim(Me.txtPaisFam.Text)) Then
        Exit Function
    End If

    Cambios = False
End Function


Private Sub CambiaAMayusculas()
    With Me
        .txtPaterno.Text = UCase(.txtPaterno.Text)
        .txtPaterno.REFRESH
        
        .txtMaterno.Text = UCase(.txtMaterno.Text)
        .txtMaterno.REFRESH
        
        .txtNombre.Text = UCase(.txtNombre.Text)
        .txtNombre.REFRESH

        .txtProf.Text = UCase(.txtProf.Text)
        .txtProf.REFRESH
    End With
End Sub


Public Sub LlenaFam()
Const DATOSFAM = 15
Dim rsFams As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String
Dim sCond As String
Dim mAncFam(DATOSFAM) As Integer
Dim mEncFam(DATOSFAM) As String


    ClrTxtFam

    frmAltaSocios.ssdbFamiliares.RemoveAll

    sCampos = "Usuarios_Club.A_Paterno, Usuarios_Club.A_Materno, "
    sCampos = sCampos & "Usuarios_Club.Nombre, Usuarios_Club.IdMember, "
    sCampos = sCampos & "Secuencial.Secuencial, Tipo_Usuario.Descripcion, "
    sCampos = sCampos & " Usuarios_Club.FechaNacio, Usuarios_Club.FechaIngreso, "
    sCampos = sCampos & " Usuarios_Club.Profesion, Usuarios_Club.Celular, "
    sCampos = sCampos & "Usuarios_Club.Email, Usuarios_Club.Sexo, Paises.Pais, "
    sCampos = sCampos & "Usuarios_Club.FotoFile, Usuarios_Club.NumeroFamiliar "
    
    sTablas = "((Usuarios_Club LEFT JOIN Tipo_Usuario ON Usuarios_Club.IdTipoUsuario=Tipo_Usuario.IdTipoUsuario) "
    sTablas = sTablas & "LEFT JOIN Paises ON Usuarios_Club.IdPais=Paises.IdPais) "
    sTablas = sTablas & "LEFT JOIN Secuencial ON Usuarios_Club.IdMember=Secuencial.IdMember "
    
    sCond = "Usuarios_Club.NoFamilia=" & Val(frmAltaSocios.txtFamilia.Text) & " AND "
    sCond = sCond & "Usuarios_Club.IdMember <> " & Val(frmAltaSocios.txtTitCve.Text)
    
    InitRecordSet rsFams, sCampos, sTablas, sCond, "", Conn
    With rsFams
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                frmAltaSocios.ssdbFamiliares.AddItem .Fields("A_Paterno") & vbTab & _
                .Fields("A_Materno") & vbTab & _
                .Fields("Nombre") & vbTab & _
                .Fields("idMember") & vbTab & _
                .Fields("Secuencial") & vbTab & _
                .Fields("Descripcion") & vbTab & _
                .Fields("FechaNacio") & vbTab & _
                .Fields("FechaIngreso") & vbTab & _
                .Fields("Profesion") & vbTab & _
                .Fields("Celular") & vbTab & _
                .Fields("Email") & vbTab & _
                .Fields("Sexo") & vbTab & _
                .Fields("Pais") & vbTab & _
                .Fields("FotoFile") & vbTab & _
                .Fields("NumeroFamiliar")
            
                .MoveNext
            Loop
        End If
    
        .Close
    End With
    Set rsFams = Nothing
    
    'Asigna valores a la matriz de encabezados
    mEncFam(0) = "A. Paterno"
    mEncFam(1) = "A. Materno"
    mEncFam(2) = "Nombre"
    mEncFam(3) = "Cve"
    mEncFam(4) = "# Sec."
    mEncFam(5) = "Tipo de usuario"
    mEncFam(6) = "Nació"
    mEncFam(7) = "Ingreso"
    mEncFam(8) = "Profesion"
    mEncFam(9) = "Celular"
    mEncFam(10) = "Email"
    mEncFam(11) = "Sexo"
    mEncFam(12) = "País"
    mEncFam(13) = "Foto"
    mEncFam(14) = "# Fam."

    'Asigna los encabezados de las columnas
    DefHeaderssGrid frmAltaSocios.ssdbFamiliares, mEncFam
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncFam(0) = 2500
    mAncFam(1) = 2500
    mAncFam(2) = 2500
    mAncFam(3) = 800
    mAncFam(4) = 800
    mAncFam(5) = 1800
    mAncFam(6) = 1100
    mAncFam(7) = 1100
    mAncFam(8) = 2500
    mAncFam(9) = 1800
    mAncFam(10) = 1800
    mAncFam(11) = 800
    mAncFam(12) = 2200
    mAncFam(13) = 700
    mAncFam(14) = 700
    
    'Asigna el ancho de cada columna
    DefAnchossGrid frmAltaSocios.ssdbFamiliares, mAncFam
End Sub


Private Sub ClrTxtFam()
    With frmAltaSocios
        .txtProf.Text = ""
        .txtTipoUser.Text = ""
        .txtFamFoto.Text = ""
        .txtEmail.Text = ""
        .txtCel.Text = ""
        .txtSec.Text = ""
        .txtNacio.Text = ""
        .txtIngreso.Text = ""
        .txtPais.Text = ""
        .txtSexo.Text = ""
        .txtCveFam.Text = ""
    End With
End Sub


Private Sub LeeDatos()
    If (sFormaAnt = "frmAltaSocios") Then
        LeeDatosDbGrid
    Else
        LeeDatosRecordset
    End If
End Sub


Private Sub LeeDatosDbGrid()
Dim rsFams As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String
Dim sCond As String


    sCampos = "Usuarios_Club.A_Paterno, Usuarios_Club.A_Materno, "
    sCampos = sCampos & "Usuarios_Club.Nombre, Usuarios_Club.IdMember, "
    sCampos = sCampos & "Secuencial.Secuencial, Tipo_Usuario.Descripcion, "
    sCampos = sCampos & " Usuarios_Club.FechaNacio, Usuarios_Club.FechaIngreso, "
    sCampos = sCampos & " Usuarios_Club.Profesion, Usuarios_Club.Celular, "
    sCampos = sCampos & "Usuarios_Club.Email, Usuarios_Club.Sexo, Paises.Pais, "
    sCampos = sCampos & "Usuarios_Club.FotoFile, Usuarios_Club.NumeroFamiliar, "
    sCampos = sCampos & "Usuarios_Club.UFechaPago, "
    'gpo 18/04/08
    sCampos = sCampos & "Usuarios_Club.Parentesco "
    '22 julio 2011
    'HABILITAR PARA DIRECCIONADOS
'    sCampos = sCampos & "Usuarios_Club.ISOPais, "
'    sCampos = sCampos & "Usuarios_Club.ISOEstado "
    
    sTablas = "((Usuarios_Club LEFT JOIN Tipo_Usuario ON Usuarios_Club.IdTipoUsuario=Tipo_Usuario.IdTipoUsuario) "
    sTablas = sTablas & "LEFT JOIN Paises ON Usuarios_Club.IdPais=Paises.IdPais) "
    '03/04/2007
    sTablas = sTablas & "LEFT JOIN Secuencial ON Usuarios_Club.IdMember=Secuencial.IdMember "
    
'    sCond = " Usuarios_Club.NoFamilia=" & Val(frmAltaSocios.txtFamilia.Text) & " AND "
    sCond = sCond & "Usuarios_Club.IdMember=" & Val(frmAltaSocios.txtCveFam.Text)
    
    InitRecordSet rsFams, sCampos, sTablas, sCond, "", Conn
    With rsFams
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                nCveFam = .Fields("IdMember")
                frmAltaFam.txtCve.Text = nCveFam
                frmAltaFam.txtCve.Enabled = False
            
                If (.Fields("A_Paterno") <> "") Then
                    frmAltaFam.txtPaterno.Text = .Fields("A_Paterno")
                End If
                
                If (.Fields("A_Materno") <> "") Then
                    frmAltaFam.txtMaterno.Text = .Fields("A_Materno")
                End If
                
                If (.Fields("Nombre") <> "") Then
                    frmAltaFam.txtNombre.Text = .Fields("Nombre")
                End If
                
                If (.Fields("Profesion") <> "") Then
                    frmAltaFam.txtProf.Text = .Fields("Profesion")
                End If
                
                frmAltaFam.dtpNacio.Value = .Fields("FechaNacio")
                frmAltaFam.txtEdad.Text = Format(Edad(.Fields("FechaNacio")), "#0.00")
                
                frmAltaFam.dtpRegistro.Value = .Fields("FechaIngreso")
                
                If (.Fields("Celular") <> "") Then
                    frmAltaFam.txtCel.Text = .Fields("Celular")
                End If
                
                If (.Fields("Email") <> "") Then
                    frmAltaFam.txtEmail.Text = .Fields("Email")
                End If
                
                If (.Fields("Secuencial") <> "") Then
                    frmAltaFam.txtSec.Text = .Fields("Secuencial")
                End If
                
                If (.Fields("Sexo") = "F") Then
                    frmAltaFam.optFemenino.Value = True
                Else
                    frmAltaFam.optMasculino.Value = True
                End If
                
                If (.Fields("Pais") <> "") Then
                    frmAltaFam.txtPaisFam.Text = .Fields("Pais")
                    frmAltaFam.txtCvePaisFam.Text = LeeXValor("IdPais", "Paises", "Pais='" & Trim(Me.txtPaisFam.Text) & "'", "IdPais", "n", Conn)
                End If
                
                If (.Fields("Descripcion") <> "") Then
                    frmAltaFam.txtTipoFam.Text = .Fields("Descripcion")
                    frmAltaFam.txtCveTipoFam.Text = LeeXValor("IdTipoUsuario", "Tipo_Usuario", "Descripcion='" & Trim(Me.txtTipoFam.Text) & "'", "IdTipoUsuario", "n", Conn)
                End If
                
                If (Not IsNull(.Fields("FotoFile"))) Then
                    frmAltaFam.txtImagen.Text = .Fields("FotoFile")
                End If
                
                If (Dir(sG_RutaFoto & "\" & Trim(.Fields("FotoFile")) & ".jpg") <> "") Then
                    frmAltaFam.imgFoto.Picture = LoadPicture(sG_RutaFoto & "\" & Trim(.Fields("FotoFile")) & ".jpg")
                Else
                    frmAltaFam.imgFoto.Picture = LoadPicture("")
                End If
                
                If (.Fields("NumeroFamiliar") <> "") Then
                    frmAltaFam.txtNumFam.Text = .Fields("NumeroFamiliar")
                End If
                
                frmAltaFam.dtpFechaInicio.Value = IIf(IsNull(.Fields("UFechaPago")), 0, .Fields("UFechaPago"))
                
                If bNvoFam Then
                    Me.dtpFechaInicio.Enabled = True
                Else
                    Me.dtpFechaInicio.Enabled = False
                End If
                
                'gpo 18/04/08
                If Not IsNull(.Fields("Parentesco")) Then
                    BuscaSSCombo Me.sscmbParentesco, .Fields("Parentesco"), 1
                    If Me.sscmbParentesco.Columns("Clave").Value = .Fields("Parentesco") Then
                        Me.sscmbParentesco.Text = Me.sscmbParentesco.Columns("Parentesco").Value
                    End If
                    
                End If
                
                '22 julio 2011
                'HABILITAR PARA DIRECCIONADOS
'                If Not IsNull(.Fields("ISOPais")) Then
'                    Call LlenaComboPaises
'                    cboPaises.Text = GetNombrePais(.Fields("ISOPais"))
'
'                    Call LlenaComboEstados(.Fields("ISOPais"))
'                    If Not IsNull(.Fields("ISOEstado")) Then
'                        cboEstados.Text = GetNombreEstado(.Fields("ISOEstado"))
'                    End If
'                End If
                
                .MoveNext
            Loop
        End If
        
        .Close
    End With
    
    
    
    strSQL = "SELECT  FECHAS_USUARIO.FechaUltimoPago"
    strSQL = strSQL & " FROM (USUARIOS_CLUB INNER JOIN CONCEPTO_TIPO ON USUARIOS_CLUB.IdTipoUsuario = CONCEPTO_TIPO.IdTipoUsuario) INNER JOIN FECHAS_USUARIO ON (USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember) AND (CONCEPTO_TIPO.IdConcepto = FECHAS_USUARIO.IdConcepto)"
    strSQL = strSQL & " Where (((USUARIOS_CLUB.IdMember) =" & nCveFam & "))"
    
    rsFams.Open strSQL
    
    If Not rsFams.EOF Then
        Me.txtUltPago.Text = rsFams!Fechaultimopago
        Select Case DateDiff("d", rsFams!Fechaultimopago, Date)
            Case Is <= 10  'Estan al corriente
                Me.txtStatus.BackColor = &HFF00&
            Case 11 To 39
                Me.txtStatus.BackColor = &HFFFF&
            Case Is >= 40
                Me.txtStatus.BackColor = &HFF&
        End Select
        
        
        
        
        
    End If
    
    rsFams.Close
    
    Set rsFams = Nothing
    
    
    
    
    
End Sub


Private Sub LeeDatosRecordset()
Dim sCampos As String
Dim sTablas As String
Dim sCondicion As String
Dim rsDatosFam As ADODB.Recordset

    sCampos = "Usuarios_Club.IdMember, Usuarios_Club.A_Paterno, Usuarios_Club.A_Materno, "
    sCampos = sCampos & "Usuarios_Club.Nombre, Usuarios_Club.Profesion, "
    sCampos = sCampos & "Usuarios_Club.FechaNacio, Usuarios_Club.FechaIngreso, "
    sCampos = sCampos & "Usuarios_Club.Celular, Usuarios_Club.Email, Usuarios_Club.NoFamilia, "
    sCampos = sCampos & "Usuarios_Club.Sexo, Tipo_Usuario.Descripcion, "
    sCampos = sCampos & "Usuarios_Club.IdPais, Usuarios_Club.IdTipoUsuario, "
    sCampos = sCampos & "Paises.Pais, Usuarios_Club.IdTitular, Secuencial.Secuencial, "
    sCampos = sCampos & "Usuarios_Club.FotoFile, Usuarios_Club.NumeroFamiliar"
    
    sTablas = "((Usuarios_Club LEFT JOIN Tipo_Usuario ON Usuarios_Club.IdTipoUsuario=Tipo_Usuario.IdTipoUsuario) "
    sTablas = sTablas & "LEFT JOIN Paises ON Usuarios_Club.IdPais=Paises.IdPais) "
    sTablas = sTablas & "LEFT JOIN Secuencial ON Usuarios_Club.IdMember=Secuencial.IdMember "
    
    sCondicion = "Usuarios_Club.IdMember=" & nCveFam

    InitRecordSet rsDatosFam, sCampos, sTablas, sCondicion, "", Conn
    
    With rsDatosFam
        If (.RecordCount > 0) Then
'            nCveFam = .Fields("Usuarios_Club!IdMember")
            frmAltaFam.txtCve.Text = nCveFam
            frmAltaFam.txtCve.Enabled = False
            
            frmAltaFam.txtTitFam.Text = .Fields("Usuarios_Club!IdTitular")
            frmAltaFam.txtFamilia.Text = .Fields("Usuarios_Club!NoFamilia")
        
            If (.Fields("Usuarios_Club!A_Paterno") <> "") Then
                frmAltaFam.txtPaterno.Text = .Fields("Usuarios_Club!A_Paterno")
            End If
            
            If (.Fields("Usuarios_Club!A_Materno") <> "") Then
                frmAltaFam.txtMaterno.Text = .Fields("Usuarios_Club!A_Materno")
            End If
            
            If (.Fields("Usuarios_Club!Nombre") <> "") Then
                frmAltaFam.txtNombre.Text = .Fields("Usuarios_Club!Nombre")
            End If
            
            If (.Fields("Usuarios_Club!Profesion") <> "") Then
                frmAltaFam.txtProf.Text = .Fields("Usuarios_Club!Profesion")
            End If
            
            frmAltaFam.dtpNacio.Value = .Fields("Usuarios_Club!FechaNacio")
            frmAltaFam.txtEdad.Text = Format(Edad(.Fields("Usuarios_Club!FechaNacio")), "#0.00")
            
            frmAltaFam.dtpRegistro.Value = .Fields("Usuarios_Club!FechaIngreso")
            
            If (.Fields("Usuarios_Club!Celular") <> "") Then
                frmAltaFam.txtCel.Text = .Fields("Usuarios_Club!Celular")
            End If
            
            If (.Fields("Usuarios_Club!Email") <> "") Then
                frmAltaFam.txtEmail.Text = .Fields("Usuarios_Club!Email")
            End If
            
            If (.Fields("Secuencial!Secuencial") <> "") Then
                frmAltaFam.txtSec.Text = .Fields("Secuencial!Secuencial")
            End If
            
            If (.Fields("Usuarios_Club!Sexo") = "F") Then
                frmAltaFam.optFemenino.Value = True
            Else
                frmAltaFam.optMasculino.Value = True
            End If
            
            If (.Fields("Tipo_Usuario!Descripcion") <> "") Then
                frmAltaFam.txtCveTipoFam.Text = .Fields("Usuarios_Club!IdTipoUsuario")
                frmAltaFam.txtTipoFam.Text = .Fields("Tipo_Usuario!Descripcion")
            End If
            
            If (.Fields("Paises!Pais") <> "") Then
                frmAltaFam.txtCvePaisFam.Text = .Fields("Usuarios_Club!IdPais")
                frmAltaFam.txtPaisFam.Text = .Fields("Paises!Pais")
            End If
            
            If (Not IsNull(.Fields("Usuarios_Club!FotoFile"))) Then
                frmAltaFam.txtImagen.Text = .Fields("Usuarios_Club!FotoFile")
            End If
            
            If (Dir(sG_RutaFoto & "\" & Trim(.Fields("Usuarios_Club!FotoFile")) & ".jpg") <> "") Then
                frmAltaFam.imgFoto.Picture = LoadPicture(sG_RutaFoto & "\" & Trim(.Fields("Usuarios_Club!FotoFile")) & ".jpg")
            Else
                frmAltaFam.imgFoto.Picture = LoadPicture("")
            End If
            
            If (.Fields("Usuarios_Club!NumeroFamiliar") <> "") Then
                frmAltaFam.txtNumFam.Text = .Fields("Usuarios_Club!NumeroFamiliar")
            End If
            
            
            
        End If
        
        .Close
    End With
    
    Set rsDatosFam = Nothing
End Sub






Private Sub txtCvePaisFam_LostFocus()
    If (Trim(Me.txtCvePaisFam.Text) <> "") Then
        If (IsNumeric(Me.txtCvePaisFam.Text)) Then
            Me.txtPaisFam.Text = LeeXValor("Pais", "Paises", "IdPais=" & Val(Me.txtCvePaisFam.Text), "Pais", "s", Conn)
            
            If (Trim(Me.txtPaisFam.Text) = "VACIO") Then
                MsgBox "El país seleccionado no existe en la base de datos.", vbExclamation, "KalaSystems"
                Me.txtPaisFam.Text = ""
                Me.txtCvePaisFam.Text = ""
                Me.txtCvePaisFam.SetFocus
            End If
        Else
            MsgBox "La clave del país es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtPaisFam.Text = ""
            Me.txtCvePaisFam.Text = ""
            Me.txtCvePaisFam.SetFocus
        End If
    End If
    
    Me.txtCvePaisFam.REFRESH
End Sub


Private Sub txtCveTipoFam_LostFocus()
Dim sCondicion As String

    If (Trim(Me.txtEdad.Text) = "") Then
        MsgBox "Se debe especificar la fecha de nacimiento para" & Chr(13) & "mostrar los tipos de usuario correspondiente.", vbInformation, "KalaSystems"
        Me.txtCveTipoFam.Text = ""
        Me.txtTipoFam.Text = ""
        Me.dtpNacio.SetFocus
        Exit Sub
    End If

    If (Trim(Me.txtCveTipoFam.Text) <> "") Then
        If (IsNumeric(Me.txtCveTipoFam.Text)) Then
        
            sCondicion = "IdTipoUsuario=" & Val(Me.txtCveTipoFam.Text) & " AND "
            sCondicion = sCondicion & "(" & Int(Val(Me.txtEdad.Text)) & " BETWEEN EdadMinima AND EdadMaxima )"
        
            Me.txtTipoFam.Text = LeeXValor("Descripcion", "Tipo_Usuario", sCondicion, "Descripcion", "s", Conn)
            
            If (Trim(Me.txtTipoFam.Text) = "VACIO") Then
                MsgBox "El tipo de usuario seleccionado es incorrecto.", vbExclamation, "KalaSystems"
                Me.txtTipoFam.Text = ""
                Me.txtCveTipoFam.Text = ""
                Me.txtCveTipoFam.SetFocus
            End If
        Else
            MsgBox "La clave del tipo de usuario debe ser numérica.", vbExclamation, "KalaSystems"
            Me.txtTipoFam.Text = ""
            Me.txtCveTipoFam.Text = ""
            Me.txtCveTipoFam.SetFocus
        End If
    End If
    
    Me.txtCveTipoFam.REFRESH
End Sub




'************************************************************
'*                          Ayudas                          *
'************************************************************

Private Sub cmdHPais_Click()
Const DATOSPAIS = 2
Dim sCadena As String
Dim mFAyuda(DATOSPAIS) As String
Dim mAAyuda(DATOSPAIS) As Integer
Dim mCAyuda(DATOSPAIS) As String
Dim mEAyuda(DATOSPAIS) As String

    nAyuda = 1

    Set frmHPais = New frmayuda
    
    mFAyuda(0) = "Países ordenados por clave"
    mFAyuda(1) = "Países ordenados por nombre"
    
    mAAyuda(0) = 800
    mAAyuda(1) = 2500
    
    mCAyuda(0) = "IdPais"
    mCAyuda(1) = "Pais"
    
    mEAyuda(0) = "Clave"
    mEAyuda(1) = "País"
    
    With frmHPais
        .nColActiva = 1
        .nColsAyuda = DATOSPAIS
        .sTabla = "Paises"
        
        .sCondicion = ""
        .sTitAyuda = "Lista de países"
        .lAgregar = True
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
    If (Trim(Me.txtCvePaisFam.Text) <> "") Then
        Me.txtPaisFam.Text = LeeXValor("Pais", "Paises", "IdPais=" & Val(Me.txtCvePaisFam.Text), "Pais", "s", Conn)
    End If
    
    Me.cmdHPais.SetFocus
End Sub


Private Sub cmdHTipo_Click()
Const DATOSTIPO = 4
Dim sCadena As String
Dim mFAyuda(DATOSTIPO) As String
Dim mAAyuda(DATOSTIPO) As Integer
Dim mCAyuda(DATOSTIPO) As String
Dim mEAyuda(DATOSTIPO) As String


    If (Trim(Me.txtEdad.Text) = vbNullString) Then
        MsgBox "Se debe especificar la fecha de nacimiento para" & Chr(13) & "mostrar los tipos de usuario correspondiente.", vbInformation, "KalaSystems"
        Me.dtpNacio.SetFocus
        Exit Sub
    End If
    
    nAyuda = 2

    Set frmHTipo = New frmayuda
    
    mFAyuda(0) = "Tipos de familiar ordenados por clave"
    mFAyuda(1) = "Tipos de familiar ordenados por descripción"
    mFAyuda(2) = "Tipos de familiar ordenados por edad mínima"
    mFAyuda(3) = "Tipos de familiar ordenados por edad máxima"
    
    mAAyuda(0) = 800
    mAAyuda(1) = 3800
    mAAyuda(2) = 900
    mAAyuda(3) = 900
    
    mCAyuda(0) = "IdTipoUsuario"
    mCAyuda(1) = "Descripcion"
    mCAyuda(2) = "EdadMinima"
    mCAyuda(3) = "EdadMaxima"
    
    mEAyuda(0) = "Clave"
    mEAyuda(1) = "Descripción"
    mEAyuda(2) = "Mínima"
    mEAyuda(3) = "Máxima"
    
    With frmHTipo
        .nColActiva = 1
        .nColsAyuda = DATOSTIPO
        .sTabla = "Tipo_Usuario"
        
        .sCondicion = "Familiar<>0 AND (" & Int(Val(Me.txtEdad.Text)) & " BETWEEN EdadMinima AND EdadMaxima )"
        .sTitAyuda = "Tipos de usuario titular"
        .lAgregar = True
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
    If (Trim(Me.txtCveTipoFam.Text) <> "") Then
        Me.txtTipoFam.Text = LeeXValor("Descripcion", "Tipo_Usuario", "IdTipoUsuario=" & Val(Me.txtCveTipoFam.Text), "Descripcion", "s", Conn)
    End If
    
    Me.cmdHTipo.SetFocus
End Sub

Public Function ChecaIntegrantesMembresia() As Boolean
    Dim adorcschkInt As ADODB.Recordset
    
    Dim lIdTipoMembresia As Long
    
    
    
    lIdTipoMembresia = 0
    
    ChecaIntegrantesMembresia = False
    
    strSQL = "SELECT IdTipoMembresia"
    strSQL = strSQL & " FROM Membresias"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember=" & Trim(Me.txtTitFam.Text)


    Set adorcschkInt = New ADODB.Recordset
    adorcschkInt.CursorLocation = adUseServer
    
    adorcschkInt.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly

    If Not adorcschkInt.EOF Then
        lIdTipoMembresia = adorcschkInt!IdTipoMembresia
    End If
    
    adorcschkInt.Close
    Set adorcschkInt = Nothing
    
    If lIdTipoMembresia = 0 Then
        'MsgBox "No se ha definido el tipo de inscripción!", vbExclamation, "Verifique"
        ChecaIntegrantesMembresia = True
        Exit Function
    End If
    
    If lIdTipoMembresia = 1 Then
        MsgBox "Este tipo de inscripción no admite familiares!", vbExclamation, "Verifique"
        Exit Function
    End If
    
    If lIdTipoMembresia = 2 Then
    
        If Val(Me.txtCveTipoFam.Text) <> 5 And Val(Me.txtCveTipoFam.Text) <> 6 And Val(Me.txtCveTipoFam.Text) <> 7 And Val(Me.txtCveTipoFam.Text) <> 9 Then
            MsgBox "Este tipo de inscripción sólo admite" & vbCr & "familiares de los tipos 7, 8 o 9", vbExclamation, "Verifique"
            Exit Function
        End If
        
        strSQL = "SELECT IdTipoUsuario"
        strSQL = strSQL & " FROM USUARIOS_CLUB"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdTitular =" & Trim(Me.txtTitFam.Text)
        strSQL = strSQL & " AND IdTipoUsuario IN ( 5, 7, 9 )"
        
        Set adorcschkInt = New ADODB.Recordset
        adorcschkInt.CursorLocation = adUseServer
        
        adorcschkInt.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        
        If Not adorcschkInt.EOF Then
            adorcschkInt.Close
            Set adorcschkInt = Nothing
            MsgBox "Esta inscripción ya tiene un agregado!", vbExclamation, "Verifique"
            Exit Function
        End If
        
        adorcschkInt.Close
        Set adorcschkInt = Nothing

        
    End If
    
    ChecaIntegrantesMembresia = True
    

End Function
