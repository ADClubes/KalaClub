VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSocios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información de Usuarios"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   14700
   Icon            =   "frmSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   14700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCartaEntrega 
      Height          =   615
      Left            =   12240
      Picture         =   "frmSocios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdFormatos 
      Height          =   615
      Left            =   12240
      Picture         =   "frmSocios.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdFechasPago 
      Height          =   615
      Left            =   12240
      Picture         =   "frmSocios.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Actualiza fechas de pago"
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdStatus 
      Caption         =   "Command1"
      Height          =   615
      Left            =   12240
      TabIndex        =   64
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdDoctos 
      Height          =   615
      Left            =   13080
      Picture         =   "frmSocios.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Control de documentos"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdContrato 
      Height          =   615
      Left            =   13920
      Picture         =   "frmSocios.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Contrato de prestación"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdEntregaFacturas 
      Height          =   615
      Left            =   13080
      Picture         =   "frmSocios.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Entrega Facturas"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdMensajes 
      Height          =   615
      Left            =   12240
      Picture         =   "frmSocios.frx":1C96
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Mensajes"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdPrintHoja 
      Height          =   615
      Left            =   13920
      Picture         =   "frmSocios.frx":20D8
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Hoja de condiciones"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton cmdPases 
      Height          =   615
      Left            =   13080
      Picture         =   "frmSocios.frx":251A
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Pases y Credenciales"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdDesactiva 
      Height          =   615
      Left            =   13080
      Picture         =   "frmSocios.frx":295C
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "  Desactivar familiar  "
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdModifica 
      Height          =   615
      Left            =   12240
      Picture         =   "frmSocios.frx":2C66
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "  Modificar datos  "
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   735
      Left            =   5520
      Picture         =   "frmSocios.frx":30A8
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "  Salir  "
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdDireccionar 
      Height          =   615
      Left            =   13920
      Picture         =   "frmSocios.frx":33B2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "  Direccionar pagos  "
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdDesactivaTodas 
      Height          =   615
      Left            =   13920
      Picture         =   "frmSocios.frx":36BC
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "  Desactivar todos  "
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton cmdActTodas 
      Height          =   615
      Left            =   13920
      Picture         =   "frmSocios.frx":39C6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "  Activar todos  "
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdActiva 
      Height          =   615
      Left            =   13080
      Picture         =   "frmSocios.frx":4498
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "  Activar familiar  "
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.OptionButton optClave 
      Caption         =   "No. familia"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox txtNoRegs 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbFamiliares 
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   6135
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Cols            =   12
      Col.Count       =   12
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
      SelectTypeRow   =   1
      BackColorOdd    =   12648384
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   10821
      _ExtentY        =   6800
      _StockProps     =   79
      Caption         =   "Integrantes de la familia"
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbBusca 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6135
      ScrollBars      =   2
      _Version        =   196616
      DataMode        =   2
      Cols            =   13
      Col.Count       =   13
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
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns(0).Width=   3200
      Columns(0).DataType=   8
      Columns(0).FieldLen=   4096
      _ExtentX        =   10821
      _ExtentY        =   5318
      _StockProps     =   79
      Caption         =   "Resultados de la busqueda"
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdBuscar 
      Default         =   -1  'True
      Height          =   735
      Left            =   3120
      Picture         =   "frmSocios.frx":47A2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "  Buscar coincidencias  "
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton optNombre 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame frmDatosTitular 
      Caption         =   "  Datos generales "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   6480
      TabIndex        =   15
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtInscripcion 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   56
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtCveTit 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   54
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txtMontoMem 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   52
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtNoFamilia 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   28
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtTel2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   25
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtTel1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox txtTitular 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox txtMembresia 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   18
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtPropMem 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtCveMembresia 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblInscripcion 
         Caption         =   "Inscripción"
         Height          =   255
         Left            =   2760
         TabIndex        =   57
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblCveTitular 
         Caption         =   "# Titular"
         Height          =   255
         Left            =   1560
         TabIndex        =   55
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblMontoMem 
         Caption         =   "Monto Memb."
         Height          =   255
         Left            =   4080
         TabIndex        =   53
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblNoFamilia 
         Caption         =   "# Familia"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label lblTel2 
         Caption         =   "Teléfono 2"
         Height          =   255
         Left            =   2760
         TabIndex        =   27
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblTel1 
         Caption         =   "Teléfono 1"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblTitular 
         Caption         =   "Nombre del titular"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblMembresia 
         Caption         =   "Tipo de inscripción"
         Height          =   255
         Left            =   1320
         TabIndex        =   21
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblCveMembresia 
         Caption         =   "# Inscripción"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblPropMem 
         Caption         =   "Propietario de la inscripción"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   3015
      End
   End
   Begin VB.Frame frmDatosFam 
      Caption         =   "  Información del titular o familiares  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   6480
      TabIndex        =   30
      Top             =   4320
      Width           =   8055
      Begin VB.CommandButton cmdCopiaFoto 
         Height          =   615
         Left            =   4440
         Picture         =   "frmSocios.frx":506C
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H0000FFFF&
         Height          =   285
         Left            =   4680
         TabIndex        =   65
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtFamNombre 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   50
         Top             =   600
         Width           =   4935
      End
      Begin VB.TextBox txtFoto 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   48
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox txtSecuencial 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   46
         Top             =   3480
         Width           =   1095
      End
      Begin VB.Frame frmFamiliar 
         Height          =   3375
         Left            =   5280
         TabIndex        =   43
         Top             =   360
         Width           =   2535
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   10
            Visible         =   0   'False
            X1              =   360
            X2              =   840
            Y1              =   840
            Y2              =   480
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H000000FF&
            BorderWidth     =   10
            FillColor       =   &H000000FF&
            Height          =   735
            Left            =   200
            Shape           =   3  'Circle
            Top             =   310
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Image imgFamiliar 
            Height          =   3015
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.TextBox txtFamFecIngreso 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtFamFecUPago 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   35
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtFamTipoUser 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox txtFamStatus 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   33
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtFamParentesco 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   32
         Top             =   3480
         Width           =   1815
      End
      Begin VB.TextBox txtFamFecNacio 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   31
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label lblFamNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblFoto 
         Caption         =   "# Foto"
         Height          =   255
         Left            =   3480
         TabIndex        =   49
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label lblSecuencial 
         Caption         =   "# secuencial"
         Height          =   255
         Left            =   2280
         TabIndex        =   47
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label lblFamParentesco 
         Alignment       =   2  'Center
         Caption         =   "Parentesco"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label lblFamStatus 
         Alignment       =   2  'Center
         Caption         =   "Status"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblFamTipo 
         Caption         =   "Tipo de familiar"
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblFamUPago 
         Alignment       =   2  'Center
         Caption         =   "Ultima fecha de pago"
         Height          =   255
         Left            =   2520
         TabIndex        =   39
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblFamFecIngreso 
         Alignment       =   2  'Center
         Caption         =   "Fecha de Ingreso"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblFechaNacio 
         Alignment       =   2  'Center
         Caption         =   "Fecha de nacimiento"
         Height          =   255
         Left            =   2520
         TabIndex        =   37
         Top             =   2520
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdPrintMem 
      Height          =   615
      Left            =   12240
      Picture         =   "frmSocios.frx":54AE
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   " Imprimir contrato  "
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblNoRegs 
      Caption         =   "# Coincidencias"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblBuscar 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'*  Formulario para alta y modificacion de socios   *
'*  Daniel Hdez                                     *
'*  20 / Octubre / 2004                             *
'*  Ultima actualización: 14 / Noviembre / 2005     *
'****************************************************


'Public bLoaded As Boolean

'Const DATOSBUSQ = 12
'gpo 25/11/2005
Const DATOSBUSQ = 13
Const DATOSFAM = 12
Const COLFOTO = 4

Dim mEncBusca(DATOSBUSQ) As String
Dim mAncBusca(DATOSBUSQ) As Integer

Dim mEncFam(DATOSFAM) As String
Dim mAncFam(DATOSFAM) As Integer

Dim sTextMain As String

Dim bOculta As Boolean



Private Sub BuscaTitulares()

    

    Dim rsTits As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String
    Dim sCondicion As String
    Dim sOrder As String
    
    Dim sQryIni As String



    'Checa la concordancia entre el tipo de dato a buscar
    'y la opcion de por nombre o por clave
    If Me.optNombre.Value And IsNumeric(Trim(Me.txtBuscar.Text)) Then
        Me.optClave.Value = True
    End If
    
    If Me.optClave.Value And Not IsNumeric(Trim(Me.txtBuscar.Text)) Then
        Me.optNombre.Value = True
    End If
    
    
    
    'Genera la porcion del query dependiendo si se busca por clave o por nombre
    If (Me.optNombre.Value) Then
        #If SqlServer_ Then
            sQryIni = "((Nombre + ' ' + A_Paterno + ' ' + A_Materno) LIKE '%" & Trim$(UCase$(Me.txtBuscar.Text)) & "%') "
        #Else
            sQryIni = "((Nombre & ' ' & A_Paterno & ' ' & A_Materno) LIKE '%" & Trim$(UCase$(Me.txtBuscar.Text)) & "%') "
        #End If
    Else
        sQryIni = "NoFamilia=" & Int(CDbl(Me.txtBuscar.Text))
    End If
    
    
    'Armado del query
    #If SqlServer_ Then
        strSQL = "SELECT U.NoFamilia, U.Name, T.Nombretitular, U.IdTitular, M.IdMembresia, M.Nombrepropietario, M.IdTipoMembresia, M.Descripcion, M.Monto, D.TEL1, D.TEL2, U.Inscripcion, U.IdMember, U.NumeroFamiliar, F.FechaUltimopago"
        strSQL = strSQL & " FROM ((((SELECT USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.NumeroFamiliar, USUARIOS_CLUB.Nombre +  ' ' + USUARIOS_CLUB.A_Paterno + ' ' +  USUARIOS_CLUB.A_Materno AS Name, USUARIOS_CLUB.IdMember,  USUARIOS_CLUB.IdTitular, USUARIOS_CLUB.Inscripcion  FROM USUARIOS_CLUB WHERE " & sQryIni & ") AS U"
        strSQL = strSQL & " LEFT JOIN (SELECT USUARIOS_CLUB.IdMember, USUARIOS_CLUB.Nombre + ' ' + USUARIOS_CLUB.A_Paterno +  ' ' +  USUARIOS_CLUB.A_Materno AS NombreTitular FROM USUARIOS_CLUB) AS T ON U.IdTitular=T.IdMember) LEFT JOIN (SELECT MEMBRESIAS.IdMember, MEMBRESIAS.IdMembresia , MEMBRESIAS.NombrePropietario, MEMBRESIAS.IdtipoMembresia, TIPO_MEMBRESIA.Descripcion, MEMBRESIAS.Monto FROM MEMBRESIAS INNER JOIN TIPO_MEMBRESIA ON MEMBRESIAS.IdTipoMembresia=TIPO_MEMBRESIA.IdTipoMembresia) AS M ON M.IdMember=U.IdTitular)  LEFT JOIN (SELECT DIRECCIONES.IdMember, DIRECCIONES.Tel1, DIRECCIONES.Tel2 FROM DIRECCIONES WHERE IdTipoDireccion=1) AS D ON U.IdTitular=D.IdMember)  LEFT JOIN"
        strSQL = strSQL & "(SELECT FECHAS_USUARIO.IdMember, FECHAS_USUARIO.FechaUltimopago FROM FECHAS_USUARIO) AS F ON U.IdMember=F.IdMember"
    
        strSQL = strSQL & " ORDER BY "
    
        If (Me.optNombre.Value) Then
            strSQL = strSQL & "(U.Name)"
        Else
            strSQL = strSQL & "U.NumeroFamiliar"
        End If
        
    #Else
        strSQL = "SELECT U.NoFamilia, U.Name, T.Nombretitular, U.IdTitular, M.IdMembresia, M.Nombrepropietario, M.IdTipoMembresia, M.Descripcion, M.Monto, D.TEL1, D.TEL2, U.Inscripcion, U.IdMember, U.NumeroFamiliar, F.FechaUltimopago"
        strSQL = strSQL & " FROM ((((SELECT USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.NumeroFamiliar, USUARIOS_CLUB.Nombre &  ' ' & USUARIOS_CLUB.A_Paterno & ' ' &  USUARIOS_CLUB.A_Materno AS Name, USUARIOS_CLUB.IdMember,  USUARIOS_CLUB.IdTitular, USUARIOS_CLUB.Inscripcion  FROM USUARIOS_CLUB WHERE " & sQryIni & ") AS U"
        strSQL = strSQL & " LEFT JOIN (SELECT USUARIOS_CLUB.IdMember, USUARIOS_CLUB.Nombre & ' ' & USUARIOS_CLUB.A_Paterno &  ' ' &  USUARIOS_CLUB.A_Materno AS NombreTitular FROM USUARIOS_CLUB) AS T ON U.IdTitular=T.IdMember) LEFT JOIN (SELECT MEMBRESIAS.IdMember, MEMBRESIAS.IdMembresia , MEMBRESIAS.NombrePropietario, MEMBRESIAS.IdtipoMembresia, TIPO_MEMBRESIA.Descripcion, MEMBRESIAS.Monto FROM MEMBRESIAS INNER JOIN TIPO_MEMBRESIA ON MEMBRESIAS.IdTipoMembresia=TIPO_MEMBRESIA.IdTipoMembresia) AS M ON M.IdMember=U.IdTitular)  LEFT JOIN (SELECT DIRECCIONES.IdMember, DIRECCIONES.Tel1, DIRECCIONES.Tel2 FROM DIRECCIONES WHERE IdTipoDireccion=1) AS D ON U.IdTitular=D.IdMember)  LEFT JOIN"
        strSQL = strSQL & "(SELECT FECHAS_USUARIO.IdMember, FECHAS_USUARIO.FechaUltimopago FROM FECHAS_USUARIO) AS F ON U.IdMember=F.IdMember"
    
        strSQL = strSQL & " ORDER BY "
    
        If (Me.optNombre.Value) Then
            strSQL = strSQL & "(U.Name)"
        Else
            strSQL = strSQL & "U.NumeroFamiliar"
        End If
    #End If
    

'    sCampos = "TITULARES!NoFamilia, (TITULARES!Nombre & ' ' & TITULARES!A_Paterno & ' ' & TITULARES!A_Materno) AS NAME, "
'
'    sCampos = sCampos & "(SELECT (Usuarios_Club!Nombre & ' ' & Usuarios_Club!A_Paterno & ' ' & Usuarios_Club!A_Materno) FROM Usuarios_Club "
'    sCampos = sCampos & "WHERE TITULARES.idTitular=Usuarios_Club.idMember) AS NOMBRETITULAR, "
'    sCampos = sCampos & "TITULARES!idTitular, "
'
'    sCampos = sCampos & "Membresias!idMembresia, Membresias!NombrePropietario, Membresias!idTipoMembresia, "
'    sCampos = sCampos & "TIPOMEM!Descripcion, Membresias!Monto, "
'
'    sCampos = sCampos & "Direcciones!Tel1, Direcciones!Tel2, "
'
'    sCampos = sCampos & "(SELECT Usuarios_Club!Inscripcion FROM Usuarios_Club "
'    sCampos = sCampos & "WHERE TITULARES.idTitular=Usuarios_Club.idMember) AS INSCRIPCION "
'
'    'gpo 25/11/2005
'    sCampos = sCampos & ", TITULARES!IdMember"
'
'    sTablas = "((Usuarios_Club AS TITULARES LEFT JOIN Membresias ON TITULARES.idTitular=Membresias.idMember) "
'    sTablas = sTablas & "LEFT JOIN Tipo_Membresia AS TIPOMEM ON Membresias.idTipoMembresia=TIPOMEM.idTipoMembresia) "
'    sTablas = sTablas & "LEFT JOIN Direcciones ON (TITULARES.idTitular=Direcciones.idMember AND Direcciones.idTipoDireccion=1) "
'
'    If (Me.optNombre.Value) Then
'        sCondicion = "((TITULARES!Nombre & ' ' & TITULARES!A_Paterno & ' ' & TITULARES!A_Materno) LIKE '%" & Trim$(UCase$(Me.txtBuscar.Text)) & "%') "
'        sOrder = "(TITULARES!Nombre & ' ' & TITULARES!A_Paterno & ' ' & TITULARES!A_Materno)"
'    Else
'        sCondicion = "TITULARES!NoFamilia=" & Int(CDbl(Me.txtBuscar.Text))
'        'sOrder = "TITULARES!NoFamilia"
'        'gpo 09/12/2005
'        sOrder = "TITULARES!NumeroFamiliar"
'    End If
'
'    InitRecordSet rsTits, sCampos, sTablas, sCondicion, sOrder, Conn
    
    
    Set rsTits = New ADODB.Recordset
    rsTits.CursorLocation = adUseServer
    
    rsTits.Open strSQL, Conn, adOpenKeyset, adLockReadOnly
    
    
    
    'Llena el ssdbGrid
    With rsTits
        If (.RecordCount > 0) Then
            Me.txtNoRegs.Text = .RecordCount
        
            .MoveFirst
            Do While (Not .EOF)
                Me.ssdbBusca.AddItem .Fields("NoFamilia") & vbTab & _
                .Fields(1) & vbTab & _
                .Fields(2) & vbTab & _
                .Fields("idTitular") & vbTab & _
                .Fields("idMembresia") & vbTab & _
                .Fields("NombrePropietario") & vbTab & _
                .Fields("idTipoMembresia") & vbTab & _
                .Fields("Descripcion") & vbTab & _
                .Fields("Monto") & vbTab & _
                .Fields("Tel1") & vbTab & _
                .Fields("Tel2") & vbTab & _
                .Fields(11) & vbTab & _
                .Fields(12)

                
                .MoveNext
            Loop
            
            Me.ssdbBusca.SetFocus
            MuestraDatosTit
            
            Me.imgFamiliar.Picture = LoadPicture("")
            
            MuestraFam
            MuestraFoto
            MuestraDatosFam
        Else
            MsgBox "No existen usuarios con éstas características.", vbExclamation, "Verifique"
            Me.txtBuscar.SelStart = 0
            Me.txtBuscar.SelLength = Len(Me.txtBuscar.Text)
            Me.txtBuscar.SetFocus
        End If
    
        .Close
    End With
    Set rsTits = Nothing
End Sub


Private Sub cmdActiva_Click()
    Dim nTor As Long
    
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdActiva.Name) Then
        Exit Sub
    End If
    
    If Not AccesoValido(Me.ssdbFamiliares.Columns(0).Value) Then
        MsgBox "¡No es posible activar a este usuario!", vbCritical, "Verificar"
        Exit Sub
    End If
    
    
    #If SqlServer_ Then
        ActivaCredSQL 1, CLng(Me.txtSecuencial.Text), 1, Me.ssdbFamiliares.Columns(0).Value, True, True
    #Else
        ActivaCred 1, CLng(Me.txtSecuencial.Text), 1, Me.ssdbFamiliares.Columns(0).Value, True, True
    #End If
    
'Registro en Torniquetes NITEGEN
    'Set Conn = Nothing
    'nErrCode = HabilitaAccesoM(CLng(Me.ssdbFamiliares.Columns(0).Value))
       
             'If nErrCode <> 0 Then
             '   MsgBox "No se pudo habilitar el usuario en torniquetes,Favor de hacerlo manual. "
             'Else
             'MsgBox "Se habilitó el usuario en torniquetes. "
             'End If
    'Connection_DB
    Me.ssdbFamiliares.SetFocus
End Sub


Private Sub cmdActTodas_Click()
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdActTodas.Name) Then
        Exit Sub
    End If
    
    EnviaDatosCred True
    
End Sub


Private Sub cmdCartaEntrega_Click()
    Dim frmReport As frmReportViewer
    
    If Not ChecaSeguridad(Me.Name, Me.cmdCartaEntrega.Name) Then
        Exit Sub
    End If
    
    
    #If SqlServer_ Then
        strSQL = "SELECT USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Nombre + ' ' + USUARIOS_CLUB.A_Paterno + ' ' +  USUARIOS_CLUB.A_Materno AS Nombre, USUARIOS_CLUB.FechaIngreso"
        strSQL = strSQL & " From USUARIOS_CLUB"
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " USUARIOS_CLUB.NoFamilia=" & Me.ssdbBusca.Columns(0).Value
        strSQL = strSQL & " AND USUARIOS_CLUB.IdMember = USUARIOS_CLUB.IdTitular"
    #Else
        strSQL = "SELECT USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Nombre & ' ' & USUARIOS_CLUB.A_Paterno & ' ' &  USUARIOS_CLUB.A_Materno AS Nombre, USUARIOS_CLUB.FechaIngreso"
        strSQL = strSQL & " From USUARIOS_CLUB"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((USUARIOS_CLUB.NoFamilia)=" & Me.ssdbBusca.Columns(0).Value & ")"
        strSQL = strSQL & " AND (USUARIOS_CLUB.IdMember = USUARIOS_CLUB.IdTitular)"
        strSQL = strSQL & ")"
    #End If
    
    Set frmReport = New frmReportViewer
    
    frmReport.sNombreReporte = sDB_ReportSource & "\cartabienvenida.rpt"
    frmReport.sQuery = strSQL
   
    
    
    
     frmReport.Show vbModal
     
     
    
End Sub

Private Sub cmdContrato_Click()
    
    Dim frmReport As frmReportViewer
    
    
    
    If Not ChecaSeguridad(Me.Name, Me.cmdContrato.Name) Then
        Exit Sub
    End If
    
    
    Set frmReport = New frmReportViewer
    
       
    
    
    frmReport.sNombreReporte = sDB_ReportSource & "\contratoprestacionservicios.rpt"
    frmReport.strValor1 = Me.txtTitular.Text
    frmReport.strValor2 = ObtieneParametro("EMPRESA CONTRATO")
    frmReport.strValor3 = ObtieneParametro("PERSONA CONTRATO")
    frmReport.strValor4 = Format(Date, "Long Date")
    
    
    frmReport.Show vbModal
    
        

    
    
    
    
    
    
    
End Sub

Private Sub cmdCopiaFoto_Click()


    If Not ChecaSeguridad(Me.Name, Me.cmdCopiaFoto.Name) Then
        Exit Sub
    End If
    
    
    Dim fsObjectCopy As FileSystemObject

    
    Me.CommonDialog1.DialogTitle = "Nombre de archivo"
    Me.CommonDialog1.CancelError = True
    On Error GoTo ErrCommonDialog
    
    
    
    Me.CommonDialog1.Filter = "Archivos gráficos jpg (*.jpg)|*.jpg|"
    
    Me.CommonDialog1.FilterIndex = 1
    
    
    Me.CommonDialog1.FileName = Me.txtCveMembresia.Text & " - " & Trim(Me.txtFamNombre.Text)
    
    'Me.CommonDialog1.InitDir
    
    Me.CommonDialog1.ShowSave
    
    
    
    
    
    Set fsObjectCopy = New FileSystemObject
    
    
    
    fsObjectCopy.CopyFile sG_RutaFoto & "\" & Trim(Me.txtFoto.Text) & ".jpg", Trim(Me.CommonDialog1.FileName), True
    
    Set fsObjectCopy = Nothing
    
    
    'Me.txtNomArchivo.Text = Trim(Me.CommonDialog1.FileName)
    'If Len(Me.txtNomArchivo.Text) > 0 Then
        
    'End If
    
    Exit Sub
ErrCommonDialog:
    Exit Sub
End Sub

'Private Sub cmdContrato_Click()
'    Dim pdfDoc As New clsPDFCreator
'    Dim strNombreArchivo As String
'    Dim strCadena As String
'    Dim ArchivoOrigen As String
'
'    Dim fs As FileSystemObject
'    Dim inputFile As TextStream
'
'
'    Dim strLine As String
'    Dim i As Long
'    Dim lPos As Long
'
'    Dim lRen As Long
'
'    ArchivoOrigen = sDB_ReportSource & "\contrato.txt"
'
'    strNombreArchivo = App.Path & "\contrato.pdf"
'
'    Set fs = New FileSystemObject
'    Set inputFile = fs.OpenTextFile(ArchivoOrigen, ForReading, False)
'
'
'    strCadena = inputFile.ReadAll
'
'    inputFile.Close
'
'    Set fs = Nothing
'
'
'
'
'    With pdfDoc
'        .Title = "Contrato de prestacion de servicios"
'        .ScaleMode = pdfMillimeter
'        .PaperSize = pdf85x11
'        .Margin = 0
'        .Orientation = pdfPortrait
'
'        .EncodeASCII85 = False
'
'        .InitPDFFile strNombreArchivo
'
'
'
'
'
'        '.LoadFont "Fnt1", "Arial", pdfNormal
'        .LoadFont "Fnt1", "Courier New"
'        .LoadFontStandard "Fnt3", "Courier New", pdfBoldItalic
'
'        .BeginPage
'
'        lPos = 1
'        lRen = 1
'        Do While lPos <= Len(strCadena)
'
'            'Obtiene una linea
'            strLine = Mid(strCadena, lPos, 105)
'
'            'Checa si tiene salto de linea
'            i = InStr(1, strLine, vbCrLf, vbBinaryCompare)
'
'            If i > 0 Then
'                strLine = Mid(strCadena, lPos, i + 1)
'                lPos = lPos + i + 1
'            Else
'                lPos = lPos + 105
'            End If
'
'
'
'            .DrawText 20, 269 - lRen * 3, strLine, "Fnt1", 8, pdfAlignLeft
'
'            lRen = lRen + 1
'
'            If lRen >= 80 Then
'                .EndPage
'                .BeginPage
'                lRen = 1
'            End If
'
'        Loop
'
'
'
'        .EndPage
'        .ClosePDFFile
'
'
'    End With
'
'
'    Call Shell("rundll32.exe url.dll,FileProtocolHandler " & (strNombreArchivo), vbMaximizedFocus)
'
'
'
'End Sub






Private Sub cmdDesActiva_Click()
    Dim nErrCode As Long
    Dim sTor As String
    
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdDesActiva.Name) Then
        Exit Sub
    End If
    
    #If SqlServer_ Then
        ActivaCredSQL 1, CLng(Me.txtSecuencial.Text), 1, Me.ssdbFamiliares.Columns(0).Value, False, True
    #Else
        ActivaCred 1, CLng(Me.txtSecuencial.Text), 1, Me.ssdbFamiliares.Columns(0).Value, False, True
    #End If
    'Registro en Torniquetes NITEGEN
    
'    nErrCode = BloqueaAccesoM(Me.ssdbFamiliares.Columns(0).Value)
'
'    If nErrCode <> 0 Then
'                MsgBox "No se pudo bloquear al usuario en torniquetes,Favor de hacerlo manual. "
'                Else
'                MsgBox "Se bloqueó al usuario en torniquetes. "
'    End If
    Me.ssdbFamiliares.SetFocus
End Sub


Private Sub cmdDesactivaTodas_Click()
    
    '27/07/06
    Dim iRespuesta As Integer
    
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdDesactivaTodas.Name) Then
        Exit Sub
    End If
    
    '27/07/06
    iRespuesta = MsgBox("¿Está seguro que desea bloquear" & vbLf & "a todos los integrantes?", vbQuestion + vbOKCancel, "Confirme")
        
    If iRespuesta = vbCancel Then
        Exit Sub
    End If
    
    EnviaDatosCred False
End Sub


Private Sub EnviaDatosCred(bActivar As Boolean)
    Dim i As Integer
    Dim nPosIni As Variant
    Dim sTor As String
    Dim nTor As Long
    
    If (Me.ssdbFamiliares.Rows > 0) Then
    
       
    
        nPosIni = Me.ssdbFamiliares.Bookmark
        Me.ssdbFamiliares.Bookmark = Me.ssdbFamiliares.AddItemBookmark(0)
    
        For i = 0 To (Me.ssdbFamiliares.Rows - 1)
            Me.ssdbFamiliares.Bookmark = Me.ssdbFamiliares.AddItemBookmark(CLng(i))
            If bActivar Then
                If Not AccesoValido(Me.ssdbFamiliares.Columns(0).Value) Then
                    MsgBox "No es posible activar a este usuario", vbCritical, "Error"
                Else
                    #If SqlServer_ Then
                        ActivaCredSQL 1, Me.ssdbFamiliares.Columns(9).Value, 1, Me.ssdbFamiliares.Columns(0).Value, bActivar, False
                    #Else
                        ActivaCred 1, Me.ssdbFamiliares.Columns(9).Value, 1, Me.ssdbFamiliares.Columns(0).Value, bActivar, False
                    #End If
                'Set Conn = Nothing
                'nErrCode = HabilitaAcceso(Me.ssdbFamiliares.Columns(0).Value)
            
                'Connection_DB
                End If
                
            Else
                #If SqlServer_ Then
                    ActivaCredSQL 1, Me.ssdbFamiliares.Columns(9).Value, 1, Me.ssdbFamiliares.Columns(0).Value, bActivar, False
                #Else
                    ActivaCred 1, Me.ssdbFamiliares.Columns(9).Value, 1, Me.ssdbFamiliares.Columns(0).Value, bActivar, False
                #End If
                'nErrCode = BloqueaAcceso(Me.ssdbFamiliares.Columns(0).Value)
            End If
        Next i
        
        MsgBox "Proceso ejecutado"
        
        Me.ssdbFamiliares.Bookmark = Me.ssdbFamiliares.AddItemBookmark(Me.ssdbFamiliares.AddItemRowIndex(nPosIni))
        Me.ssdbBusca.SetFocus
    End If
End Sub


Private Sub cmdDireccionar_Click()
    If (Me.ssdbBusca.Rows = 0) Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdDireccionar.Name) Then
        Exit Sub
    End If

    
    
    frmDireccionar.nTitular = Val(Me.txtCveTit.Text)
    frmDireccionar.sTitular = Trim$(Me.txtTitular.Text)
    Load frmDireccionar
    frmDireccionar.Show (1)
    
End Sub


Private Sub cmdDoctos_Click()
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdDoctos.Name) Then
        Exit Sub
    End If

    frmCtrlDoctos.lIdTitular = CSng(Me.txtCveTit.Text)
    frmCtrlDoctos.Show vbModal
    
End Sub

Private Sub cmdEntregaFacturas_Click()
    
    Dim frmEntrega As frmEntregaFacturas
    
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdEntregaFacturas.Name) Then
        Exit Sub
    End If
    
    Set frmEntrega = New frmEntregaFacturas
    frmEntrega.lIdTitular = CSng(Me.txtCveTit.Text)
    frmEntrega.Show vbModal
    
    
End Sub

Private Sub cmdFechasPago_Click()
    
    Dim frmFechas As frmFechaPago
    
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdFechasPago.Name) Then
        Exit Sub
    End If
    
    
    Set frmFechas = New frmFechaPago
    frmFechas.lNoFamilia = Me.ssdbBusca.Columns(0).Value
    frmFechas.Show vbModal
    
    
End Sub

Private Sub cmdFormatos_Click()
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdFormatos.Name) Then
        Exit Sub
    End If

    frmFormatos.lIdTitular = CSng(Me.txtCveTit.Text)
    frmFormatos.Show vbModal
End Sub

Private Sub cmdMensajes_Click()
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdMensajes.Name) Then
        Exit Sub
    End If

    frmMensajesMovs.lIdTitular = CSng(Me.txtCveTit.Text)
    frmMensajesMovs.Show vbModal
End Sub

Private Sub cmdModifica_Click()
    If (Me.ssdbBusca.Rows > 0) Then
        frmAltaSocios.bSocioNvo = False
        frmAltaSocios.sFormaAnterior = "frmSocios"
        Load frmAltaSocios
        frmAltaSocios.Show (1)
    End If
    
    Me.ssdbBusca.SetFocus
End Sub


Private Sub cmdPases_Click()
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdPases.Name) Then
        Exit Sub
    End If

    frmCredyPases.lIdTitular = CSng(Me.txtCveTit.Text)
    frmCredyPases.lIdMemberSel = CSng(Me.ssdbFamiliares.Columns(0).Value)
    frmCredyPases.Show vbModal
End Sub

Private Sub cmdPrintHoja_Click()

    Dim frmHoja As frmHojaCond
    
    If Me.ssdbBusca.Rows = 0 Then
        MsgBox "Seleccione un usuario!", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdPrintHoja.Name) Then
        Exit Sub
    End If
    
    
    
    If ExistWindow("frmHojaCond") Then
        frmHojaCond.WindowState = 0
        Unload frmHojaCond
    End If
    
    
    Set frmHoja = New frmHojaCond
    frmHoja.nidTitular = CSng(Me.txtCveTit.Text)

    frmHoja.Show vbModal
    
    
    
End Sub

Private Sub cmdPrintMem_Click()
    
'    If Me.ssdbBusca.Rows = 0 Then
'        MsgBox "Seleccione un usuario!", vbExclamation, "Verifique"
'        Exit Sub
'    End If
'
'    If ExistWindow("frmRepMembresia") Then
'        frmRepMembresia.WindowState = 0
'    End If
'
'    Load frmRepMembresia
'
'    frmRepMembresia.nidTitular = CSng(Me.txtCveTit.Text)
'    frmRepMembresia.nTotalMem = CDbl(Me.txtMontoMem.Text)
'
'    frmRepMembresia.Show

    



    
'    Me.ssdbBusca.SetFocus
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub cmdStatus_Click()
    
    Dim frmStat As frmStatus
    
    If Me.ssdbBusca.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdStatus.Name) Then
        Exit Sub
    End If
    
    Set frmStat = New frmStatus
    
    frmStat.lIdTitular = CSng(Me.txtCveTit.Text)
    
    frmStat.Show vbModal
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSocios = Nothing
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextMain
End Sub




Private Sub ssdbBusca_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    With Me.ssdbBusca
        ClrTxtFam
        Me.imgFamiliar.Picture = LoadPicture("")
    
        If (.Rows > 0) Then
            MuestraDatosTit
            MuestraFam
            MuestraFoto
            MuestraDatosFam
'
'            If (.Rows = 1) Then
'                MuestraFoto
'
'                Me.imgFamiliar.Picture = LoadPicture("")
'            End If
        End If
    End With
End Sub


Private Sub ssdbFamiliares_DblClick()
Dim nPosIni As Variant

    Exit Sub

    If (Me.ssdbFamiliares.Rows > 0) Then
        nPosIni = Me.ssdbFamiliares.Bookmark
    
        frmAltaFam.bNvoFam = False
        frmAltaFam.nCveFam = Me.ssdbFamiliares.Columns(0).Value
        Load frmAltaFam
        frmAltaFam.Show (1)
        
        If (Me.ssdbBusca.Rows > 0) Then
            ClrTxtFam
            Me.imgFamiliar.Picture = LoadPicture("")
            
            MuestraFam
            Me.ssdbFamiliares.Bookmark = Me.ssdbFamiliares.AddItemBookmark(Me.ssdbFamiliares.AddItemRowIndex(nPosIni))
            MuestraFoto
            MuestraDatosFam
        End If
    End If
End Sub


Private Sub ssdbFamiliares_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    With Me.ssdbFamiliares
        If (.Rows > 0) Then
            MuestraFoto
            
            MuestraDatosFam
        End If
    End With
End Sub


Private Sub MuestraFoto()
    If (Dir(sG_RutaFoto & "\" & Me.ssdbFamiliares.Columns(COLFOTO).Value & ".jpg") <> "") Then
        Me.imgFamiliar.Picture = LoadPicture(sG_RutaFoto & "\" & Me.ssdbFamiliares.Columns(COLFOTO).Value & ".jpg")
    
        If (Trim$(Me.ssdbFamiliares.Columns(6).Text) = "INACTIVO") Then
            Me.Shape1.Visible = True
            Me.Line1.Visible = True
        Else
            Me.Shape1.Visible = False
            Me.Line1.Visible = False
        End If
    Else
        Me.imgFamiliar.Picture = LoadPicture("")
        Me.Shape1.Visible = False
        Me.Line1.Visible = False
    End If
End Sub


Private Sub MuestraDatosTit()
    With Me
        .txtTitular.Text = .ssdbBusca.Columns(2).Text
        .txtCveMembresia.Text = .ssdbBusca.Columns(4).Value
        .txtPropMem.Text = .ssdbBusca.Columns(5).Value
        .txtMembresia.Text = "(" & .ssdbBusca.Columns(6).Value & ") " & .ssdbBusca.Columns(7).Value
'        Bloqueo Fin de semana
'        If .ssdbBusca.Columns(7).Value Like "*FIN*" Then
'            cmdActiva.Enabled = False
'            cmdActTodas.Enabled = False
'        Else
'            cmdActiva.Enabled = True
'            cmdActTodas.Enabled = True
'        End If
        
        
        If (Val(.ssdbBusca.Columns(8).Value) > 0) Then
            .txtMontoMem.Text = Format(Round(CDbl(.ssdbBusca.Columns(8).Value), 2), "###,###,##0.#0")
        Else
            .txtMontoMem.Text = 0
        End If
        
        .txtTel1.Text = .ssdbBusca.Columns(9).Text
        .txtTel2.Text = .ssdbBusca.Columns(10).Text
        .txtNoFamilia.Text = .ssdbBusca.Columns(0).Value
        .txtCveTit.Text = .ssdbBusca.Columns(3).Value
        .txtInscripcion.Text = .ssdbBusca.Columns(11).Text
        
        
        
    End With
End Sub


Private Sub MuestraFam()
    Dim rsFamiliares As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String
    Dim sCondicion As String
    Dim sOrder As String

    Dim vBookMark As Variant

    Me.ssdbFamiliares.RemoveAll
    
    
    
    
    

'    sCampos = "FAMILIARES!idMember, (FAMILIARES!Nombre & ' ' & FAMILIARES!A_Paterno & ' ' & FAMILIARES!A_Materno) AS NAME, "
'    sCampos = sCampos & "FAMILIARES!FechaNacio, FAMILIARES!FechaIngreso, "
'    sCampos = sCampos & "FAMILIARES!FotoFile, FAMILIARES!UFechaPago, "
'    sCampos = sCampos & "FAMILIARES!Status, FAMILIARES!idTipoUsuario, "
'
'    sCampos = sCampos & "TIPOUSER!Descripcion, Secuencial!Secuencial, "
'
'    sCampos = sCampos & "Parentesco!Parentesco, FAMILIARES!idTitular "
'
'    sTablas = "((Usuarios_Club AS FAMILIARES LEFT JOIN Tipo_Usuario AS TIPOUSER ON FAMILIARES.idTipoUsuario=TIPOUSER.idTipoUsuario) "
'    sTablas = sTablas & "LEFT JOIN Secuencial ON FAMILIARES.idMember=Secuencial.idMember) "
'    sTablas = sTablas & "LEFT JOIN Parentesco ON TIPOUSER.Parentesco=Parentesco.Clave"
'
'    sCondicion = "FAMILIARES.idTitular=" & Me.ssdbBusca.Columns(3).Value
'    'sOrder = "(FAMILIARES!Nombre & ' ' & FAMILIARES!A_Paterno & ' ' & FAMILIARES!A_Materno)"
'    'gpo 15/11/2005
'    sOrder = "FAMILIARES!NumeroFamiliar"
'
'    InitRecordSet rsFamiliares, sCampos, sTablas, sCondicion, sOrder, Conn


    
    
    #If SqlServer_ Then
        strSQL = "SELECT U.IdMember, (U.Nombre + ' ' + U.A_Paterno + ' ' + U.A_Materno) AS NAME, U.FechaNacio, U.FechaIngreso, U.FotoFile, U.UFechaPago, U.Status, U.IdTipoUsuario, T.Descripcion, S.Secuencial, P.Parentesco, U.IdTitular, F.FechaUltimoPago"
        strSQL = strSQL & " FROM (((USUARIOS_CLUB AS U INNER JOIN TIPO_USUARIO AS T ON U.IdTipoUsuario = T.IdTipoUsuario) INNER JOIN SECUENCIAL AS S ON U.IdMember = S.IdMember) INNER JOIN PARENTESCO AS P ON T.Parentesco = P.Clave) INNER JOIN FECHAS_USUARIO F ON U.IdMember = F.IdMember"
        strSQL = strSQL & " WHERE U.IdTitular= " & Me.ssdbBusca.Columns(3).Value
        strSQL = strSQL & " ORDER BY U.NumeroFamiliar"
    #Else
        strSQL = "SELECT U.IdMember, (U.Nombre & ' ' & U.A_Paterno & ' ' & U.A_Materno) AS NAME, U.FechaNacio, U.FechaIngreso, U.FotoFile, U.UFechaPago, U.Status, U.IdTipoUsuario, T.Descripcion, S.Secuencial, P.Parentesco, U.IdTitular, F.FechaUltimoPago"
        strSQL = strSQL & " FROM (((USUARIOS_CLUB AS U INNER JOIN TIPO_USUARIO AS T ON U.IdTipoUsuario = T.IdTipoUsuario) INNER JOIN SECUENCIAL AS S ON U.IdMember = S.IdMember) INNER JOIN PARENTESCO AS P ON T.Parentesco = P.Clave) INNER JOIN FECHAS_USUARIO F ON U.IdMember = F.IdMember"
        strSQL = strSQL & " WHERE U.IdTitular= " & Me.ssdbBusca.Columns(3).Value
        strSQL = strSQL & " ORDER BY U.NumeroFamiliar"
    #End If
    Set rsFamiliares = New ADODB.Recordset
    rsFamiliares.CursorLocation = adUseServer
    
    rsFamiliares.Open strSQL, Conn, adOpenKeyset, adLockReadOnly
    
    With rsFamiliares
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                Me.ssdbFamiliares.AddItem .Fields("idMember") & vbTab & _
                .Fields(1) & vbTab & _
                Format(.Fields("FechaNacio"), "dd/mmm/yyyy") & vbTab & _
                Format(.Fields("FechaIngreso"), "dd/mmm/yyyy") & vbTab & _
                .Fields("FotoFile") & vbTab & _
                Format(.Fields("FechaUltimoPago"), "dd/mmm/yyyy") & vbTab & _
                .Fields("Status") & vbTab & _
                .Fields("idTipoUsuario") & vbTab & _
                .Fields("Descripcion") & vbTab & _
                .Fields("Secuencial") & vbTab & _
                .Fields("Parentesco") & vbTab & _
                .Fields("idTitular")
                
                'gpo 25/11/2005
                If .Fields("IdMember") = Val(Me.ssdbBusca.Columns(12).Value) Then
                    vBookMark = Me.ssdbFamiliares.AddItemBookmark(Me.ssdbFamiliares.Rows - 1)
                End If
                
                .MoveNext
            Loop
            
'            MuestraFoto
'            MuestraDatosFam
        End If

        .Close
    End With
    Set rsFamiliares = Nothing
    
    Me.ssdbFamiliares.Bookmark = vBookMark
    
End Sub


Private Sub MuestraDatosFam()
    With Me
        .txtFamNombre.Text = .ssdbFamiliares.Columns(1).Text
        .txtFamFecNacio.Text = .ssdbFamiliares.Columns(2).Value
        .txtFamFecIngreso.Text = .ssdbFamiliares.Columns(3).Value
        .txtFamFecUPago.Text = .ssdbFamiliares.Columns(5).Value
        .txtFamTipoUser.Text = "(" & .ssdbFamiliares.Columns(7).Value & ") " & .ssdbFamiliares.Columns(8).Value
        .txtFamStatus.Text = .ssdbFamiliares.Columns(6).Value
        .txtFamParentesco.Text = .ssdbFamiliares.Columns(10).Text
        .txtSecuencial.Text = .ssdbFamiliares.Columns(9).Value
        .txtFoto.Text = .ssdbFamiliares.Columns(4).Value
        
        
        Select Case DateDiff("d", CDate(.txtFamFecUPago.Text), Date)
            Case Is <= 10  'Estan al corriente
                Me.txtStatus.BackColor = &HFF00&
            Case 11 To 30
                Me.txtStatus.BackColor = &HFFFF&
            Case Is >= 31
                Me.txtStatus.BackColor = &HFF&
        End Select
        
        
    End With
End Sub


Private Sub cmdBuscar_Click()
    If (Trim$(Me.txtBuscar.Text) <> vbNullString) Then
        
        OcultaCols
        
        ClrTxtTit
    
        Me.ssdbBusca.RemoveAll
        Me.ssdbFamiliares.RemoveAll
        Me.imgFamiliar.Picture = LoadPicture("")
    
        BuscaTitulares
    End If
End Sub


Private Sub Form_Load()
    With Me
        .Top = 0
        .Left = 0
        .Height = 8925
        .Width = 15050 '14175
    End With
    
    InitColsSocios
    
    bLoaded = True
    bOculta = True
    
    sTextMain = MDIPrincipal.StatusBar1.Panels.Item(1).Text
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Consulta y modificación de socios"
End Sub


Private Sub OcultaCols()
Dim i As Integer

    
    If (bOculta) Then
        For i = 2 To (Me.ssdbBusca.Cols - 1)
            Me.ssdbBusca.Columns(i).Visible = False
        Next i
    
        For i = 2 To (Me.ssdbFamiliares.Cols - 1)
            Me.ssdbFamiliares.Columns(i).Visible = False
        Next i
    
        bOculta = False
    End If
End Sub


Private Sub ClrTxtTit()
    With Me
        .txtTitular.Text = ""
        .txtCveMembresia.Text = ""
        .txtMembresia.Text = ""
        .txtPropMem.Text = ""
        .txtMontoMem.Text = ""
        .txtTel1.Text = ""
        .txtTel2.Text = ""
        .txtNoFamilia.Text = ""
        .txtCveTit.Text = ""
        .txtInscripcion.Text = ""
    End With
End Sub


Private Sub ClrTxtFam()
    With Me
        .txtFamNombre.Text = ""
        .txtFamFecIngreso.Text = ""
        .txtFamFecUPago.Text = ""
        .txtFamTipoUser.Text = ""
        .txtFamStatus.Text = ""
        .txtFamFecNacio.Text = ""
        .txtFamParentesco.Text = ""
        .txtSecuencial.Text = ""
        .txtFoto.Text = ""
    End With
End Sub


Private Sub InitColsSocios()
Dim i As Byte

    'Asigna valores a la matriz de encabezados
    mEncBusca(0) = "# Familia"
    mEncBusca(1) = "Nombre"
    mEncBusca(2) = "Titular"
    mEncBusca(3) = "# Titular"
    mEncBusca(4) = "# membresia"
    mEncBusca(5) = "Propietario"
    mEncBusca(6) = "# tipo"
    mEncBusca(7) = "Tipo membresia"
    mEncBusca(8) = "Monto Memb."
    mEncBusca(9) = "Teléfono 1"
    mEncBusca(10) = "Teléfono 2"
    mEncBusca(11) = "Inscripción"
    mEncBusca(12) = "IdMember"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid Me.ssdbBusca, mEncBusca
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncBusca(0) = 1000
    mAncBusca(1) = 4550
    mAncBusca(2) = 4250
    mAncBusca(3) = 1000
    mAncBusca(4) = 1000
    mAncBusca(5) = 4250
    mAncBusca(6) = 1000
    mAncBusca(7) = 2500
    mAncBusca(8) = 1000
    mAncBusca(9) = 1200
    mAncBusca(10) = 1200
    mAncBusca(11) = 1200
    mAncBusca(12) = 1200
    
    'Asigna el ancho de cada columna
    DefAnchossGrid Me.ssdbBusca, mAncBusca
    
    Me.ssdbBusca.Columns(0).Alignment = ssCaptionAlignmentRight
    Me.ssdbBusca.Columns(2).Alignment = ssCaptionAlignmentRight


    'Asigna valores a la matriz de encabezados
    mEncFam(0) = "# usuario"
    mEncFam(1) = "Nombre"
    mEncFam(2) = "Fecha nacio"
    mEncFam(3) = "Fec. ingreso"
    mEncFam(4) = "Foto"
    mEncFam(5) = "Fec. pago"
    mEncFam(6) = "Status"
    mEncFam(7) = "# tipo"
    mEncFam(8) = "Tipo usuario"
    mEncFam(9) = "# secuencial"
    mEncFam(10) = "# foto"
    mEncFam(11) = "# titular"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid Me.ssdbFamiliares, mEncFam

    'Asigna valores a la matriz que define el ancho de cada columna
    mAncFam(0) = 1000
    mAncFam(1) = 4550
    mAncFam(2) = 1745
    mAncFam(3) = 1100
    mAncFam(4) = 1100
    mAncFam(5) = 1100
    mAncFam(6) = 1000
    mAncFam(7) = 1000
    mAncFam(8) = 3500
    mAncFam(9) = 1000
    mAncFam(10) = 1000
    mAncFam(11) = 1000

    'Asigna el ancho de cada columna
    DefAnchossGrid Me.ssdbFamiliares, mAncFam

    Me.ssdbFamiliares.Columns(0).Alignment = ssCaptionAlignmentRight
    Me.ssdbFamiliares.Columns(2).Alignment = ssCaptionAlignmentCenter
End Sub
