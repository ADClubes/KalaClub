VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL3N.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmAltaSocios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de usuarios"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   Icon            =   "frmAltaSocios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin TabDlg.SSTab sstabSocios 
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Generales"
      TabPicture(0)   =   "frmAltaSocios.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblImagen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdModMembresia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmDatosAccion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frmFoto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSalir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdGuardar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtImagen"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdMembresias"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Direcciones"
      TabPicture(1)   =   "frmAltaSocios.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdDelDir"
      Tab(1).Control(1)=   "cmdModDir"
      Tab(1).Control(2)=   "cmdAddDir"
      Tab(1).Control(3)=   "ssdbDireccion"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Familiares"
      TabPicture(2)   =   "frmAltaSocios.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtFamFoto"
      Tab(2).Control(1)=   "txtSec"
      Tab(2).Control(2)=   "txtCveFam"
      Tab(2).Control(3)=   "txtSexo"
      Tab(2).Control(4)=   "txtPais"
      Tab(2).Control(5)=   "txtTipoUser"
      Tab(2).Control(6)=   "txtProf"
      Tab(2).Control(7)=   "txtCel"
      Tab(2).Control(8)=   "txtEmail"
      Tab(2).Control(9)=   "txtIngreso"
      Tab(2).Control(10)=   "txtNacio"
      Tab(2).Control(11)=   "cmdDelFam"
      Tab(2).Control(12)=   "cmdModFam"
      Tab(2).Control(13)=   "cmdAddFam"
      Tab(2).Control(14)=   "frameFamiliar"
      Tab(2).Control(15)=   "ssdbFamiliares"
      Tab(2).Control(16)=   "lblFamFoto"
      Tab(2).Control(17)=   "lblSec"
      Tab(2).Control(18)=   "lblCveFam"
      Tab(2).Control(19)=   "lblTipoUser"
      Tab(2).Control(20)=   "lblProf"
      Tab(2).Control(21)=   "lblCel"
      Tab(2).Control(22)=   "lblEmail"
      Tab(2).Control(23)=   "lblSexo"
      Tab(2).Control(24)=   "lblPaisOrigen"
      Tab(2).Control(25)=   "lblIngreso"
      Tab(2).Control(26)=   "lblFechaNacio"
      Tab(2).ControlCount=   27
      TabCaption(3)   =   "Rentables"
      TabPicture(3)   =   "frmAltaSocios.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ssdbRenta"
      Tab(3).Control(1)=   "cmdPeriodoPago"
      Tab(3).Control(2)=   "cmdContratoRentable"
      Tab(3).Control(3)=   "cmdDelRenta"
      Tab(3).Control(4)=   "cmdAddRenta"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Ausencias"
      TabPicture(4)   =   "frmAltaSocios.frx":093A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdDelFalta"
      Tab(4).Control(1)=   "cmdModFalta"
      Tab(4).Control(2)=   "cmdAddFalta"
      Tab(4).Control(3)=   "ssdbAusencias"
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Facturas"
      TabPicture(5)   =   "frmAltaSocios.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtObs"
      Tab(5).Control(1)=   "txtTotalFact"
      Tab(5).Control(2)=   "ssdbDetalle"
      Tab(5).Control(3)=   "ssdbFacturas"
      Tab(5).Control(4)=   "lblTotalFact"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Acceso"
      TabPicture(6)   =   "frmAltaSocios.frx":0972
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame2"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "MultiClub"
      TabPicture(7)   =   "frmAltaSocios.frx":098E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).ControlCount=   0
      TabCaption(8)   =   "Datos Emergencia"
      TabPicture(8)   =   "frmAltaSocios.frx":09AA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame4"
      Tab(8).Control(1)=   "fraBeneficiarios"
      Tab(8).Control(2)=   "cmdGuardaDatoEmer"
      Tab(8).Control(3)=   "cmdAnexoBene"
      Tab(8).ControlCount=   4
      Begin VB.CommandButton cmdMembresias 
         Enabled         =   0   'False
         Height          =   550
         Left            =   5040
         Picture         =   "frmAltaSocios.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   132
         ToolTipText     =   "Inscripciones"
         Top             =   7140
         Width           =   550
      End
      Begin VB.TextBox txtImagen 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7200
         TabIndex        =   131
         Top             =   6060
         Width           =   855
      End
      Begin VB.CommandButton cmdGuardar 
         Height          =   550
         Left            =   5760
         Picture         =   "frmAltaSocios.frx":0CD0
         Style           =   1  'Graphical
         TabIndex        =   130
         ToolTipText     =   " Guardar datos "
         Top             =   7140
         Width           =   550
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   550
         Left            =   6480
         Picture         =   "frmAltaSocios.frx":1112
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   " Salir "
         Top             =   7140
         Width           =   550
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos del titular "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   120
         TabIndex        =   87
         Top             =   2460
         Width           =   6975
         Begin VB.CommandButton cmdHTipo 
            Height          =   305
            Left            =   960
            Picture         =   "frmAltaSocios.frx":141C
            Style           =   1  'Graphical
            TabIndex        =   109
            ToolTipText     =   " Tipos de usuarios "
            Top             =   3480
            Width           =   425
         End
         Begin VB.CommandButton cmdHPais 
            Height          =   305
            Left            =   960
            Picture         =   "frmAltaSocios.frx":1566
            Style           =   1  'Graphical
            TabIndex        =   108
            ToolTipText     =   "  Lista de países "
            Top             =   2880
            Width           =   425
         End
         Begin VB.TextBox txtCveTipo 
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   107
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox txtCvePais 
            Height          =   285
            Left            =   120
            MaxLength       =   4
            TabIndex        =   106
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtTipoTit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   105
            Top             =   3480
            Width           =   5295
         End
         Begin VB.TextBox txtPaisTit 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   104
            Top             =   2880
            Width           =   3615
         End
         Begin VB.TextBox txtSecuencial 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4440
            TabIndex        =   103
            Top             =   2280
            Width           =   855
         End
         Begin VB.TextBox txtFamilia 
            Enabled         =   0   'False
            Height          =   285
            Left            =   5280
            MaxLength       =   7
            TabIndex        =   102
            Top             =   2880
            Width           =   735
         End
         Begin VB.Frame frmTitSexo 
            Caption         =   " Sexo "
            ForeColor       =   &H000000FF&
            Height          =   975
            Left            =   5400
            TabIndex        =   98
            Top             =   1440
            Width           =   1335
            Begin VB.OptionButton optFemenino 
               Caption         =   "Femenino"
               Height          =   255
               Left            =   120
               TabIndex        =   101
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optMasculino 
               Caption         =   "Masculino"
               Height          =   255
               Left            =   120
               TabIndex        =   100
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.TextBox txtTitCel 
            Height          =   285
            Left            =   3720
            MaxLength       =   20
            TabIndex        =   97
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtTitEmail 
            Height          =   285
            Left            =   120
            MaxLength       =   60
            TabIndex        =   99
            Top             =   2280
            Width           =   4215
         End
         Begin VB.TextBox txtTitProf 
            Height          =   285
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   94
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtTitCve 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6120
            MaxLength       =   6
            TabIndex        =   92
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtTitMaterno 
            Height          =   285
            Left            =   3480
            MaxLength       =   60
            TabIndex        =   91
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox txtTitEdad 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   90
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtTitPaterno 
            Height          =   285
            Left            =   120
            MaxLength       =   60
            TabIndex        =   89
            Top             =   480
            Width           =   3255
         End
         Begin VB.TextBox txtTitNombre 
            Height          =   285
            Left            =   120
            MaxLength       =   60
            TabIndex        =   93
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtNoIns 
            Height          =   285
            Left            =   2280
            TabIndex        =   88
            Top             =   4200
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker dtpTitRegistro 
            Height          =   285
            Left            =   2280
            TabIndex        =   96
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   49872897
            CurrentDate     =   37995
         End
         Begin MSComCtl2.DTPicker dtpTitNacio 
            Height          =   285
            Left            =   120
            TabIndex        =   95
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   49872897
            CurrentDate     =   37995
         End
         Begin MSComCtl2.DTPicker dtpFechaUPago 
            Height          =   285
            Left            =   120
            TabIndex        =   110
            Top             =   4200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   49872897
            CurrentDate     =   38362
         End
         Begin VB.Label lblCveTipo 
            Caption         =   "Cve. Tipo"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   3240
            Width           =   735
         End
         Begin VB.Label lblCvePais 
            Caption         =   "Cve. País"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   127
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label lblTipoTit 
            Caption         =   "Tipo de usuario"
            Height          =   255
            Left            =   1440
            TabIndex        =   126
            Top             =   3240
            Width           =   3015
         End
         Begin VB.Label lblSecuencial 
            Caption         =   "# Sec."
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   4440
            TabIndex        =   125
            Top             =   2040
            Width           =   855
         End
         Begin VB.Label lblFamilia 
            Caption         =   "# Fam."
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   5400
            TabIndex        =   124
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label lblPais 
            Caption         =   "País de origen"
            Height          =   255
            Left            =   1440
            TabIndex        =   123
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lblTitCel 
            Caption         =   "Celular"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3720
            TabIndex        =   122
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lblTitEmail 
            Caption         =   "Email"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   121
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblTitProf 
            Caption         =   "Profesión"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3480
            TabIndex        =   120
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblTitCve 
            Caption         =   "# Reg."
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   6120
            TabIndex        =   119
            Top             =   2640
            Width           =   615
         End
         Begin VB.Label lblTitRegistro 
            Caption         =   "Inscripción"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2280
            TabIndex        =   118
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblNacio 
            Caption         =   "Nació"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblTitEdad 
            Caption         =   "Edad"
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   1560
            TabIndex        =   116
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblTitPaterno 
            Caption         =   "Apellido paterno"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label lblTitMaterno 
            Caption         =   "Apellido materno"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3480
            TabIndex        =   114
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblTitNombre 
            Caption         =   "Nombre(s)"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   113
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblFechaIniMtto 
            Caption         =   "Inicio de uso de Inst."
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   3960
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "No. Ins."
            Height          =   255
            Left            =   2280
            TabIndex        =   111
            Top             =   3960
            Width           =   615
         End
      End
      Begin VB.Frame frmFoto 
         Height          =   3135
         Left            =   7200
         TabIndex        =   86
         Top             =   2460
         Width           =   2415
         Begin VB.Image imgFoto 
            BorderStyle     =   1  'Fixed Single
            Height          =   2775
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame frmDatosAccion 
         Caption         =   " Datos de la acción o inscripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   67
         Top             =   1020
         Width           =   9495
         Begin VB.OptionButton optMembresia 
            Caption         =   "Inscripcion"
            Height          =   255
            Left            =   2280
            TabIndex        =   85
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optRentista 
            Caption         =   "Rentista"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1320
            TabIndex        =   84
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optPropietario 
            Caption         =   "Propietario"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   1095
         End
         Begin VB.Frame frmAccion 
            Caption         =   " Seleccione una acción "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   75
            Top             =   480
            Width           =   3255
            Begin VB.CommandButton cmdAyuda 
               Enabled         =   0   'False
               Height          =   305
               Left            =   2640
               Picture         =   "frmAltaSocios.frx":16B0
               Style           =   1  'Graphical
               TabIndex        =   79
               ToolTipText     =   "  Mostrar acciones disponibles  "
               Top             =   480
               Width           =   425
            End
            Begin VB.TextBox txtNumero 
               Enabled         =   0   'False
               Height          =   285
               Left            =   1800
               MaxLength       =   5
               TabIndex        =   78
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox txtTipo 
               Enabled         =   0   'False
               Height          =   285
               Left            =   960
               MaxLength       =   8
               TabIndex        =   77
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox txtSerie 
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               MaxLength       =   8
               TabIndex        =   76
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lblNumero 
               Caption         =   "Número"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   1800
               TabIndex        =   82
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lblTipo 
               Caption         =   "Tipo"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   960
               TabIndex        =   81
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblSerie 
               Caption         =   "Serie"
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   80
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Frame frmAccionista 
            Caption         =   " Dueño de la acción "
            Height          =   1215
            Left            =   3480
            TabIndex        =   68
            Top             =   120
            Width           =   5895
            Begin VB.TextBox txtTel2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Left            =   4200
               TabIndex        =   72
               Top             =   840
               Width           =   1455
            End
            Begin VB.TextBox txtTel1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Left            =   2520
               TabIndex        =   71
               Top             =   840
               Width           =   1455
            End
            Begin VB.TextBox txtCveAccionista 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Left            =   1080
               TabIndex        =   70
               Top             =   840
               Width           =   855
            End
            Begin VB.TextBox txtNombre 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   69
               Top             =   360
               Width           =   5535
            End
            Begin VB.Label lblTels 
               Alignment       =   1  'Right Justify
               Caption         =   "Tels."
               Height          =   255
               Left            =   2040
               TabIndex        =   74
               Top             =   885
               Width           =   375
            End
            Begin VB.Label lblCveAccionista 
               Caption         =   "# Accionista"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   885
               Width           =   975
            End
         End
      End
      Begin VB.CommandButton cmdDelDir 
         Height          =   615
         Left            =   -66120
         Picture         =   "frmAltaSocios.frx":17FA
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Eliminar direccion"
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton cmdModDir 
         Height          =   615
         Left            =   -66960
         Picture         =   "frmAltaSocios.frx":1B04
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   " Modificar dirección "
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton cmdAddDir 
         Height          =   615
         Left            =   -67800
         Picture         =   "frmAltaSocios.frx":1F46
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   " Agregar dirección "
         Top             =   1260
         Width           =   615
      End
      Begin VB.TextBox txtFamFoto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67800
         TabIndex        =   63
         Top             =   5145
         Width           =   855
      End
      Begin VB.TextBox txtSec 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70800
         TabIndex        =   62
         Top             =   5745
         Width           =   735
      End
      Begin VB.TextBox txtCveFam 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -66840
         TabIndex        =   61
         Top             =   6300
         Width           =   735
      End
      Begin VB.TextBox txtSexo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -67440
         TabIndex        =   60
         Top             =   6300
         Width           =   495
      End
      Begin VB.TextBox txtPais 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -69960
         TabIndex        =   59
         Top             =   6300
         Width           =   2415
      End
      Begin VB.TextBox txtTipoUser 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71760
         TabIndex        =   58
         Top             =   5145
         Width           =   3855
      End
      Begin VB.TextBox txtProf 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74760
         TabIndex        =   57
         Top             =   5145
         Width           =   2895
      End
      Begin VB.TextBox txtCel 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -69960
         TabIndex        =   56
         Top             =   5745
         Width           =   3015
      End
      Begin VB.TextBox txtEmail 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74760
         TabIndex        =   55
         Top             =   5745
         Width           =   3855
      End
      Begin VB.TextBox txtIngreso 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -72360
         TabIndex        =   54
         Top             =   6300
         Width           =   2295
      End
      Begin VB.TextBox txtNacio 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   -74760
         TabIndex        =   53
         Top             =   6300
         Width           =   2295
      End
      Begin VB.CommandButton cmdDelFam 
         Height          =   615
         Left            =   -66000
         Picture         =   "frmAltaSocios.frx":2810
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Eliminar Familiar"
         Top             =   5580
         Width           =   615
      End
      Begin VB.CommandButton cmdModFam 
         Height          =   615
         Left            =   -66000
         Picture         =   "frmAltaSocios.frx":2B1A
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   " Modificar familiar "
         Top             =   4860
         Width           =   615
      End
      Begin VB.CommandButton cmdAddFam 
         Height          =   615
         Left            =   -66840
         Picture         =   "frmAltaSocios.frx":2F5C
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   " Agregar familiar "
         Top             =   4860
         Width           =   615
      End
      Begin VB.Frame frameFamiliar 
         Height          =   3255
         Left            =   -67800
         TabIndex        =   49
         Top             =   1500
         Width           =   2415
         Begin VB.Image imgFamiliar 
            BorderStyle     =   1  'Fixed Single
            Height          =   2895
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdAddRenta 
         Height          =   615
         Left            =   -66840
         Picture         =   "frmAltaSocios.frx":3826
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   " Agregar rentable"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdDelRenta 
         Height          =   615
         Left            =   -66120
         Picture         =   "frmAltaSocios.frx":40F0
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Eliminar rentable"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtObs 
         Enabled         =   0   'False
         Height          =   495
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   6240
         Width           =   7935
      End
      Begin VB.TextBox txtTotalFact 
         Alignment       =   1  'Right Justify
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
         Left            =   -66720
         TabIndex        =   40
         Top             =   6540
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelFalta 
         Height          =   615
         Left            =   -66120
         Picture         =   "frmAltaSocios.frx":43FA
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Eliminar"
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton cmdModFalta 
         Height          =   615
         Left            =   -66960
         Picture         =   "frmAltaSocios.frx":4704
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   " Modificar datos "
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton cmdAddFalta 
         Height          =   615
         Left            =   -67800
         Picture         =   "frmAltaSocios.frx":4B46
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   " Registrar ausencias "
         Top             =   1260
         Width           =   615
      End
      Begin VB.CommandButton cmdModMembresia 
         Caption         =   "Mod.Mem"
         Height          =   550
         Left            =   2880
         TabIndex        =   36
         Top             =   7140
         Visible         =   0   'False
         Width           =   550
      End
      Begin VB.Frame Frame2 
         ClipControls    =   0   'False
         Height          =   5535
         Left            =   -74880
         TabIndex        =   27
         Top             =   1380
         Width           =   9495
         Begin VB.CommandButton cmdConsAcc 
            Caption         =   "Consulta"
            Height          =   495
            Left            =   3120
            TabIndex        =   32
            Top             =   4920
            Width           =   1095
         End
         Begin VB.ListBox lstUsuarios 
            Height          =   3765
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   4095
         End
         Begin VB.CheckBox chkNoExcep 
            Caption         =   "Sin Excepciones"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   5040
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtpFecFin 
            Height          =   375
            Left            =   1680
            TabIndex        =   29
            Top             =   4320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   49872897
            CurrentDate     =   38870
         End
         Begin MSComCtl2.DTPicker dtpFecIni 
            Height          =   375
            Left            =   120
            TabIndex        =   30
            Top             =   4320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   49872897
            CurrentDate     =   38870
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgAcceso 
            Height          =   5055
            Left            =   4320
            TabIndex        =   33
            Top             =   240
            Width           =   5055
            _Version        =   196616
            DataMode        =   2
            Col.Count       =   4
            SelectTypeCol   =   0
            SelectTypeRow   =   0
            RowHeight       =   423
            Columns.Count   =   4
            Columns(0).Width=   2328
            Columns(0).Caption=   "Fecha"
            Columns(0).Name =   "Fecha"
            Columns(0).DataField=   "Column 0"
            Columns(0).DataType=   8
            Columns(0).FieldLen=   256
            Columns(1).Width=   1588
            Columns(1).Caption=   "Hora"
            Columns(1).Name =   "Hora"
            Columns(1).DataField=   "Column 1"
            Columns(1).DataType=   8
            Columns(1).FieldLen=   256
            Columns(2).Width=   2461
            Columns(2).Caption=   "Entrada/Salida"
            Columns(2).Name =   "EoS"
            Columns(2).CaptionAlignment=   2
            Columns(2).DataField=   "Column 2"
            Columns(2).DataType=   8
            Columns(2).FieldLen=   256
            Columns(3).Width=   1561
            Columns(3).Caption=   "Excepcion"
            Columns(3).Name =   "Excepcion"
            Columns(3).DataField=   "Column 3"
            Columns(3).DataType=   8
            Columns(3).FieldLen=   256
            _ExtentX        =   8916
            _ExtentY        =   8916
            _StockProps     =   79
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
         Begin VB.Label Label2 
            Caption         =   "Fec.Inicial"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Fec.Final"
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   4080
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Contacto en caso de emergencia"
         Height          =   2175
         Left            =   -74760
         TabIndex        =   18
         Top             =   960
         Width           =   9255
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   0
            Left            =   3240
            MaxLength       =   80
            TabIndex        =   22
            Top             =   360
            Width           =   5775
         End
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   1
            Left            =   3240
            MaxLength       =   15
            TabIndex        =   21
            Top             =   840
            Width           =   5775
         End
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   2
            Left            =   3240
            MaxLength       =   25
            TabIndex        =   20
            Top             =   1320
            Width           =   5775
         End
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   3
            Left            =   3240
            MaxLength       =   100
            TabIndex        =   19
            Top             =   1800
            Width           =   5775
         End
         Begin VB.Label Label8 
            Caption         =   "En caso de accidente comunicarse con: "
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label9 
            Caption         =   "Parentesco:"
            Height          =   375
            Left            =   2280
            TabIndex        =   25
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "Telefono:"
            Height          =   255
            Left            =   2400
            TabIndex        =   24
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Domicilio:"
            Height          =   255
            Left            =   2400
            TabIndex        =   23
            Top             =   1800
            Width           =   855
         End
      End
      Begin VB.Frame fraBeneficiarios 
         Caption         =   "Beneficiarios"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   6
         Top             =   3600
         Visible         =   0   'False
         Width           =   9255
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   4
            Left            =   1320
            MaxLength       =   80
            TabIndex        =   12
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   5
            Left            =   1320
            MaxLength       =   80
            TabIndex        =   11
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   6
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   10
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   7
            Left            =   4920
            MaxLength       =   15
            TabIndex        =   9
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   8
            Left            =   7800
            MaxLength       =   6
            TabIndex        =   8
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtCtrlEmer 
            Height          =   285
            Index           =   9
            Left            =   7800
            MaxLength       =   6
            TabIndex        =   7
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Beneficiario 1:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Beneficiario 2:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Porcentaje"
            Height          =   255
            Left            =   8040
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Parentesco"
            Height          =   255
            Left            =   5760
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   2520
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdGuardaDatoEmer 
         Caption         =   "Guardar"
         Height          =   495
         Left            =   -66840
         TabIndex        =   5
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton cmdContratoRentable 
         Height          =   615
         Left            =   -74760
         Picture         =   "frmAltaSocios.frx":5410
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Contrato"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton cmdAnexoBene 
         Height          =   615
         Left            =   -74640
         Picture         =   "frmAltaSocios.frx":5852
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Contrato"
         Top             =   6000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdPeriodoPago 
         Height          =   615
         Left            =   -74040
         TabIndex        =   2
         Top             =   1320
         Width           =   615
      End
      Begin SSDataWidgets_B.SSDBGrid ssdbDetalle 
         Height          =   1935
         Left            =   -74760
         TabIndex        =   42
         Top             =   4260
         Width           =   9255
         _Version        =   196616
         DataMode        =   2
         Cols            =   12
         Col.Count       =   12
         AllowUpdate     =   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   16325
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Detalle de la factura"
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
      Begin SSDataWidgets_B.SSDBGrid ssdbFacturas 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   43
         Top             =   1020
         Width           =   9255
         _Version        =   196616
         DataMode        =   2
         Cols            =   11
         Col.Count       =   11
         AllowUpdate     =   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   16325
         _ExtentY        =   5530
         _StockProps     =   79
         Caption         =   "Datos de la factura"
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
      Begin SSDataWidgets_B.SSDBGrid ssdbRenta 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   44
         Top             =   1980
         Width           =   9255
         _Version        =   196616
         DataMode        =   2
         Cols            =   5
         Col.Count       =   5
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   16325
         _ExtentY        =   8493
         _StockProps     =   79
         Caption         =   "Información de Arts. rentables"
      End
      Begin SSDataWidgets_B.SSDBGrid ssdbFamiliares 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   45
         Top             =   1620
         Width           =   6855
         _Version        =   196616
         DataMode        =   2
         Cols            =   15
         Col.Count       =   15
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   12091
         _ExtentY        =   5530
         _StockProps     =   79
         Caption         =   "Información de familiares"
      End
      Begin SSDataWidgets_B.SSDBGrid ssdbDireccion 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   46
         Top             =   1980
         Width           =   9255
         _Version        =   196616
         DataMode        =   2
         Cols            =   14
         Col.Count       =   14
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   16325
         _ExtentY        =   8493
         _StockProps     =   79
         Caption         =   "Información de direcciones"
      End
      Begin SSDataWidgets_B.SSDBGrid ssdbAusencias 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   133
         Top             =   1980
         Width           =   9255
         _Version        =   196616
         DataMode        =   2
         Cols            =   8
         Col.Count       =   8
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeRow   =   1
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   16325
         _ExtentY        =   8493
         _StockProps     =   79
         Caption         =   "Fechas de ausencias"
      End
      Begin VB.Label lblImagen 
         Caption         =   "# imagen"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   7200
         TabIndex        =   146
         Top             =   5820
         Width           =   975
      End
      Begin VB.Label lblFamFoto 
         Caption         =   "# imagen"
         Height          =   255
         Left            =   -67800
         TabIndex        =   145
         Top             =   4905
         Width           =   855
      End
      Begin VB.Label lblSec 
         Caption         =   "# Sec."
         Height          =   255
         Left            =   -70800
         TabIndex        =   144
         Top             =   5505
         Width           =   735
      End
      Begin VB.Label lblCveFam 
         Caption         =   "Clave"
         Height          =   255
         Left            =   -66840
         TabIndex        =   143
         Top             =   6060
         Width           =   615
      End
      Begin VB.Label lblTipoUser 
         Caption         =   "Tipo de usuario"
         Height          =   255
         Left            =   -71760
         TabIndex        =   142
         Top             =   4905
         Width           =   1695
      End
      Begin VB.Label lblProf 
         Caption         =   "Profesión"
         Height          =   255
         Left            =   -74760
         TabIndex        =   141
         Top             =   4905
         Width           =   2055
      End
      Begin VB.Label lblCel 
         Caption         =   "Celular"
         Height          =   255
         Left            =   -69960
         TabIndex        =   140
         Top             =   5505
         Width           =   1695
      End
      Begin VB.Label lblEmail 
         Caption         =   "Email"
         Height          =   255
         Left            =   -74760
         TabIndex        =   139
         Top             =   5505
         Width           =   1815
      End
      Begin VB.Label lblSexo 
         Caption         =   "Sexo"
         Height          =   255
         Left            =   -67440
         TabIndex        =   138
         Top             =   6060
         Width           =   495
      End
      Begin VB.Label lblPaisOrigen 
         Caption         =   "País de origen"
         Height          =   255
         Left            =   -69960
         TabIndex        =   137
         Top             =   6060
         Width           =   2175
      End
      Begin VB.Label lblIngreso 
         Caption         =   "Fecha de ingreso"
         Height          =   255
         Left            =   -72360
         TabIndex        =   136
         Top             =   6060
         Width           =   1815
      End
      Begin VB.Label lblFechaNacio 
         Caption         =   "Fecha de nacimiento"
         Height          =   255
         Left            =   -74760
         TabIndex        =   135
         Top             =   6060
         Width           =   1815
      End
      Begin VB.Label lblTotalFact 
         Alignment       =   2  'Center
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66720
         TabIndex        =   134
         Top             =   6300
         Width           =   1095
      End
   End
   Begin VB.Label lblParameter 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmAltaSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*  Formulario para capturar los datos de la familia                *
'*  Daniel Hdez                                                     *
'*  20/09/2004                                                      *
'*  Ultima actualización: 27/12/2007                                *
'********************************************************************
Option Explicit

Public nAyuda As Byte
Public bSocioNvo As Boolean
Public nTitCve As Integer
'Public nPosFam As Variant
Public sFormaAnterior As String

Dim frmHAcciones As frmayuda
Dim frmHTipo As frmayuda
Dim frmHPais As frmayuda
Dim sTextToolBar As String

'Datos del titular
Public sTitPaterno As String
Public sTitMaterno As String
Public sTitNombre As String
Dim sTitPais As String
Dim sTitTipo As String
Dim sTitCel As String
Dim sTitEmail As String
Dim dTitNacio As Date
Dim dTitRegistro As Date
Dim sProf As String
Dim bFemenino As Boolean
Dim nPais As Integer
Dim nTipo As Integer
Dim nTipoAnt As Integer
Dim nTipoNvo As Integer
'11/10/2005 gpo
Dim sNoIns As String
Dim dFecUPago As Date

'Banderas para el ctrl de cambios
Public bMemb As Boolean
Public bProp As Boolean
Public bRent As Boolean
Dim lCambio As Boolean

'Datos del accionista
Public nAccCve As Integer
Dim sAccTel1 As String
Dim sAccTel2 As String
Dim sSerie As String
Dim sTipo As String
Dim nNumero As Integer

'Cols de los datos de las facturas
Const DATOSFACT = 11
Dim mAncFacts(DATOSFACT) As Integer
Dim mEncFacts(DATOSFACT) As String

'Cols del detalle de las facturas
Const DATOSDET = 12
Dim mAncDet(DATOSDET) As Integer
Dim mEncDet(DATOSDET) As String

'Torniquetes NITGEN
Dim nErrCode As Long




Private Sub cmbUsuarios_DblClick()
    Dim adoRcsAcceso As ADODB.Recordset
    
    #If SqlServer_ Then
        strSQL = "SELECT FECHA, HORA, CASE WHEN ISNULL(ENT_SAL,0) = 1 THEN 'ENTRADA' ELSE 'SALIDA' END AS EoS, EXCEPCION"
        strSQL = strSQL & " FROM (ACCESO1 LEFT JOIN SECUENCIAL ON ACCESO1.Secuencial=SECUENCIAL.Secuencial)"
        strSQL = strSQL & " LEFT JOIN USUARIOS_CLUB ON SECUENCIAL.IdMember=USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & "USUARIOS_CLUB.IdMember=" & Trim(Me.txtTitCve)
    #Else
        strSQL = "SELECT FECHA, HORA, IIF(ENT_SAL = 1, 'ENTRADA','SALIDA') AS EoS, EXCEPCION"
        strSQL = strSQL & " FROM (ACCESO1 LEFT JOIN SECUENCIAL ON ACCESO1.Secuencial=SECUENCIAL.Secuencial)"
        strSQL = strSQL & " LEFT JOIN USUARIOS_CLUB ON SECUENCIAL.IdMember=USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & "USUARIOS_CLUB.IdMember=" & Trim(Me.txtTitCve)
    #End If
    
    Set adoRcsAcceso = New ADODB.Recordset
    adoRcsAcceso.CursorLocation = adUseServer
    adoRcsAcceso.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Do Until adoRcsAcceso.EOF
        Me.ssdbgAcceso.AddItem adoRcsAcceso!Fecha & vbTab & adoRcsAcceso!Hora & vbTab & adoRcsAcceso!EOS & vbTab & adoRcsAcceso!Excepcion
        adoRcsAcceso.MoveNext
    Loop
    
    adoRcsAcceso.Close
    Set adoRcsAcceso = Nothing
    
End Sub

Private Sub cmdAddDir_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdAddDir.Name) Then
        Exit Sub
    End If
    
    
    frmAltaDir.bNvaDir = True
    frmAltaDir.Show (1)
'    frmAltaSocios.dgDireccion.SetFocus
End Sub

Private Sub cmdAddFalta_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdAddFalta.Name) Then
        Exit Sub
    End If
    
    frmAusencias.bNvaAus = True
    frmAusencias.Show (1)
    
    frmAltaSocios.ssdbAusencias.SetFocus
End Sub

Private Sub cmdAddFam_Click()
    Dim nRen As Variant
    
    If Not ChecaSeguridad(Me.Name, Me.cmdAddFam.Name) Then
        Exit Sub
    End If
    
    If (Me.ssdbFamiliares.Rows > 0) Then
        nRen = Me.ssdbFamiliares.Bookmark
    Else
        nRen = 0
    End If

    frmAltaFam.bNvoFam = True
    frmAltaFam.Show (1)
    
    frmAltaSocios.ssdbFamiliares.SetFocus
    
    If (Me.ssdbFamiliares.Rows > 0) Then
        Me.ssdbFamiliares.Bookmark = Me.ssdbFamiliares.AddItemBookmark(Me.ssdbFamiliares.AddItemRowIndex(nRen))
        
        Call ssdbFamiliares_RowColChange(Me.ssdbFamiliares.Rows, Me.ssdbFamiliares.Cols)
        
        Me.ssdbFamiliares.SetFocus
    End If
End Sub

Private Sub cmdAddRenta_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdAddRenta.Name) Then
        Exit Sub
    End If
    
    frmAltaRenta.bNewRent = True
    frmAltaRenta.Show (1)
    
    frmAltaSocios.ssdbRenta.SetFocus
End Sub

Private Sub cmdCambia_Click()
    If Not ChecaSeguridad(Me.Name, "") Then
        Exit Sub
    End If
End Sub

Private Sub cmdAnexoBene_Click()
    Dim frmreporte As frmReportViewer
    
    Set frmreporte = New frmReportViewer
    
    strSQL = "SELECT USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.FechaNacio, EMERGENCIA_DATOS.NombreBeneficiario1, EMERGENCIA_DATOS.ParentescoBeneficiario1, EMERGENCIA_DATOS.PorcentajeBeneficiario1, EMERGENCIA_DATOS.NombreBeneficiario2, EMERGENCIA_DATOS.ParentescoBeneficiario2, EMERGENCIA_DATOS.PorcentajeBeneficiario2, DIRECCIONES.Calle, DIRECCIONES.Colonia, DIRECCIONES.CodPos, DIRECCIONES.Estado, DIRECCIONES.Ciudad, DIRECCIONES.DeloMuni"
    strSQL = strSQL & " FROM (USUARIOS_CLUB LEFT JOIN EMERGENCIA_DATOS ON USUARIOS_CLUB.IdMember = EMERGENCIA_DATOS.IdMember) LEFT JOIN DIRECCIONES ON USUARIOS_CLUB.IdMember = DIRECCIONES.IdMember"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((USUARIOS_CLUB.IdMember)=" & Trim(Me.txtTitCve.Text) & ")"
    strSQL = strSQL & " AND ((DIRECCIONES.IdTipoDireccion)=1));"
    
    frmreporte.sNombreReporte = sDB_ReportSource & "\" & "anexo_beneficiarios_seguro.rpt"
    frmreporte.sQuery = strSQL
    
    frmreporte.Show vbModal
End Sub

Private Sub cmdConsAcc_Click()
    Dim adoRcsAcceso As ADODB.Recordset
    Dim sSite As String
    Dim sSiteNew As String
    Dim sSiteCad As String
    
    If Me.lstUsuarios.ListIndex < 0 Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    sSite = ObtieneParametro("SITECODE")
    sSiteNew = ObtieneParametro("SITECODENEW")
    
    sSite = Right$(sSite, 3)
    
    If sSiteNew <> vbNullString Then
        sSiteNew = Right$(sSiteNew, 3)
    End If
    
    sSiteCad = "('" & sSite & "'"
    
    If sSiteNew <> vbNullString Then
        sSiteCad = sSiteCad & ",'" & sSiteNew & "'"
    End If
    
    sSiteCad = sSiteCad & ")"
    
    #If SqlServer_ Then
        strSQL = "SELECT FECHA, HORA, CASE WHEN ISNULL(ENT_SAL,0) = 1 THEN 'ENTRADA' ELSE 'SALIDA' END AS EoS, EXCEPCION"
        strSQL = strSQL & " FROM (ACCESO1 LEFT JOIN SECUENCIAL ON ACCESO1.Secuencial=SECUENCIAL.Secuencial)"
        strSQL = strSQL & " LEFT JOIN USUARIOS_CLUB ON SECUENCIAL.IdMember=USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " USUARIOS_CLUB.IdMember=" & Me.lstUsuarios.ItemData(Me.lstUsuarios.ListIndex)
        strSQL = strSQL & " AND ACCESO1.Fecha Between " & "'" & Format(Me.dtpFecIni.Value, "yyyymmdd") & "'" & " AND " & "'" & Format(Me.dtpFecFin.Value, "yyyymmdd") & "'"
        strSQL = strSQL & " AND SITE IN" & sSiteCad
        If Me.chkNoExcep.Value Then
            strSQL = strSQL & " AND EXCEPCION=Space(2)"
        End If
        strSQL = strSQL & " ORDER BY ACCESO1.Fecha, ACCESO1.Hora"
    #Else
        strSQL = "SELECT FECHA, HORA, IIF(ENT_SAL = 1, 'ENTRADA','SALIDA') AS EoS, EXCEPCION"
        strSQL = strSQL & " FROM (ACCESO1 LEFT JOIN SECUENCIAL ON ACCESO1.Secuencial=SECUENCIAL.Secuencial)"
        strSQL = strSQL & " LEFT JOIN USUARIOS_CLUB ON SECUENCIAL.IdMember=USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " USUARIOS_CLUB.IdMember=" & Me.lstUsuarios.ItemData(Me.lstUsuarios.ListIndex)
        strSQL = strSQL & " AND ACCESO1.Fecha Between " & "#" & Format(Me.dtpFecIni.Value, "mm/dd/yyyy") & "#" & " AND " & "#" & Format(Me.dtpFecFin.Value, "mm/dd/yyyy") & "#"
        strSQL = strSQL & " AND SITE IN" & sSiteCad
        If Me.chkNoExcep.Value Then
            strSQL = strSQL & " AND EXCEPCION=Space(2)"
        End If
        strSQL = strSQL & " ORDER BY ACCESO1.Fecha, ACCESO1.Hora"
    #End If
        
    Set adoRcsAcceso = New ADODB.Recordset
    adoRcsAcceso.CursorLocation = adUseServer
    adoRcsAcceso.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.ssdbgAcceso.RemoveAll
    
    Do Until adoRcsAcceso.EOF
        Me.ssdbgAcceso.AddItem adoRcsAcceso!Fecha & vbTab & adoRcsAcceso!Hora & vbTab & adoRcsAcceso!EOS & vbTab & adoRcsAcceso!Excepcion
        adoRcsAcceso.MoveNext
    Loop
    
    adoRcsAcceso.Close
    Set adoRcsAcceso = Nothing
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdContratoRentable_Click()
    Dim frmreporte As frmReportViewer
    
    Set frmreporte = New frmReportViewer
    
    If Me.ssdbRenta.Rows = 0 Then
        Exit Sub
    End If
    
    strSQL = "SELECT RENTABLES.Numero, RENTABLES.Area, RENTABLES.FechaInicio, RENTABLES.FechaPago, RENTABLES.ImportePagado, USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.FechaNacio, PARAMETROS.Valor"
    strSQL = strSQL & " FROM PARAMETROS, RENTABLES INNER JOIN USUARIOS_CLUB ON RENTABLES.IdUsuario = USUARIOS_CLUB.IdMember"
    strSQL = strSQL & " WHERE ((LTRIM(RTrim(RENTABLES.Numero))='" & Trim(Me.ssdbRenta.Columns("Número").Value) & "')"
    strSQL = strSQL & " AND ((PARAMETROS.Nombre_Parametro)='EMPRESA CONTRATO'))"
    
    frmreporte.sNombreReporte = sDB_ReportSource & "\" & "contratolockers.rpt"
    frmreporte.sQuery = strSQL
    
    frmreporte.Show vbModal
    
End Sub

Private Sub cmdDelDir_Click()
    Dim nAnswer As Integer

    If Not ChecaSeguridad(Me.Name, Me.cmdDelDir.Name) Then
        Exit Sub
    End If

    If (Me.ssdbDireccion.Rows > 0) Then
        nAnswer = MsgBox("¿Realmente desea borrar el domicilio seleccionado?", vbYesNo, "Registro de direcciones")
        
        If (nAnswer = vbYes) Then
            If (EliminaReg("Direcciones", "idDireccion=" & Me.ssdbDireccion.Columns(10).Value, "", Conn)) Then
                frmAltaDir.LlenaDirs
                frmAltaSocios.ssdbDireccion.SetFocus
            End If
        End If
    End If
End Sub

Private Sub cmdDelFalta_Click()
    Dim nAnswer As Integer

    If Not ChecaSeguridad(Me.Name, Me.cmdDelFalta.Name) Then
        Exit Sub
    End If

    If (Me.ssdbAusencias.Rows > 0) Then
        nAnswer = MsgBox("¿Desea borrar el período de ausencia?", vbYesNo, "Registro de ausencias")
        
        If (nAnswer = vbYes) Then
            If (EliminaReg("Ausencias", "idAusencia=" & Me.ssdbAusencias.Columns(0).Value, "", Conn)) Then
                frmAusencias.ActivaMiembro (Me.ssdbAusencias.Columns(4).Value)
                frmAusencias.LlenaAusencias
            End If
        End If
    End If
    
    frmAltaSocios.ssdbAusencias.SetFocus
End Sub

Private Sub cmdDelFam_Click()
    Dim nAnswer As Integer

    If Not ChecaSeguridad(Me.Name, Me.cmdDelFam.Name) Then
        Exit Sub
    End If

    If (Me.ssdbFamiliares.Rows > 0) Then
        nAnswer = MsgBox("¿Realmente desea borrar al familiar seleccionado?", vbYesNo, "Registro de familiares")
        
        If (nAnswer = vbYes) Then
            If (QuitaFamiliar(Me.ssdbFamiliares.Columns(3).Value, 0)) Then
                frmAltaFam.LlenaFam
                
                frmAltaSocios.ssdbFamiliares.SetFocus
                
                If (Me.ssdbFamiliares.Rows > 0) Then
                    Me.ssdbFamiliares.row = 0
                    Call ssdbFamiliares_RowColChange(Me.ssdbFamiliares.Rows, Me.ssdbFamiliares.Cols)
                End If
            Else
                MsgBox "No se realizó la baja del familiar seleccionado.", vbExclamation, "KalaSystems"
            End If
        End If
    End If
End Sub

Private Sub cmdDelRenta_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdDelRenta.Name) Then
        Exit Sub
    End If
    
    If (Me.ssdbRenta.Rows > 0) Then
        frmAltaRenta.QuitaRenta
        frmAltaRenta.LlenaRenta
        
        frmAltaSocios.ssdbRenta.SetFocus
    End If
End Sub

Private Sub cmdGuardaDatoEmer_Click()
    Dim adocmd As ADODB.Command
    Dim lTrans As Long
    
    If Not ChecaSeguridad(Me.Name, Me.cmdGuardaDatoEmer.Name) Then
        Exit Sub
    End If
    
    If bSeguroVida = True Then
        If Val(Me.txtCtrlEmer(8).Text) + Val(Me.txtCtrlEmer(9).Text) <> 100 Then
            MsgBox "La suma de los porcentajes de los" & vbCrLf & "debe ser igual a 100%", vbCritical, "Verifique"
            Exit Sub
        End If
    End If
    
    Conn.Errors.Clear
    
    On Error GoTo Error_Catch
    
    lTrans = Conn.BeginTrans
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    
    #If SqlServer_ Then
        strSQL = "DELETE"
    #Else
        strSQL = "DELETE *"
    #End If
    
    strSQL = strSQL & " FROM EMERGENCIA_DATOS"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember = " & Trim(Me.txtTitCve.Text)
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    strSQL = "INSERT INTO EMERGENCIA_DATOS ("
    strSQL = strSQL & " IdMember,"
    strSQL = strSQL & " NombreEmergencia,"
    strSQL = strSQL & " ParentescoEmergencia,"
    strSQL = strSQL & " TelefonosEmergencia,"
    strSQL = strSQL & " DomicilioEmergencia,"
    strSQL = strSQL & " NombreBeneficiario1,"
    strSQL = strSQL & " ParentescoBeneficiario1,"
    strSQL = strSQL & " PorcentajeBeneficiario1,"
    strSQL = strSQL & " NombreBeneficiario2,"
    strSQL = strSQL & " ParentescoBeneficiario2,"
    strSQL = strSQL & " PorcentajeBeneficiario2)"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & Trim(Me.txtTitCve.Text) & ","
    strSQL = strSQL & "'" & Trim(Me.txtCtrlEmer(0).Text) & "',"
    strSQL = strSQL & "'" & Trim(Me.txtCtrlEmer(1).Text) & "',"
    strSQL = strSQL & "'" & Trim(Me.txtCtrlEmer(2).Text) & "',"
    strSQL = strSQL & "'" & Trim(Me.txtCtrlEmer(3).Text) & "',"
    strSQL = strSQL & "'" & Trim(Me.txtCtrlEmer(4).Text) & "',"
    strSQL = strSQL & "'" & Trim(Me.txtCtrlEmer(6).Text) & "',"
    strSQL = strSQL & Val(Trim(Me.txtCtrlEmer(8).Text)) / 100 & ","
    strSQL = strSQL & "'" & Trim(Me.txtCtrlEmer(5).Text) & "',"
    strSQL = strSQL & "'" & Trim(Me.txtCtrlEmer(7).Text) & "',"
    strSQL = strSQL & Val(Trim(Me.txtCtrlEmer(9).Text)) / 100 & ")"
    
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    Set adocmd = Nothing
    Conn.CommitTrans
    
    On Error GoTo 0
    
    MsgBox "Datos de emergencia actualizados", vbInformation, "Correcto"
    
    Exit Sub
    
Error_Catch:
    
    If lTrans > 0 Then
        Conn.RollbackTrans
    End If
    
    MsgError
    
End Sub

Private Sub cmdGuardar_Click()
    Dim i As Byte
    
    If Not ChecaSeguridad(Me.Name, Me.cmdGuardar.Name) Then
        Exit Sub
    End If

    If (Cambios) Then
        If (GuardaDatos) Then
            HabilitaTabs True
            
            Me.cmdMembresias.Enabled = True
            
            LlenaRecordsets
            
            'Inicializa las variables
            InitVarAcc
            InitVarTit
            
            lCambio = False
        Else
            MsgBox "No se registraron los datos, verifique la información.", vbCritical, "KalaSystems"
        End If
    End If
    
    If (Me.optMembresia.Value) Then
        Me.txtTitPaterno.SetFocus
    Else
        Me.txtSerie.SetFocus
    End If
End Sub

Private Sub cmdMembresias_Click()
    If (Val(Me.txtTitCve.Text) > 0) Then
        frmMembresia.sFormaAnterior = "frmAltaSocios"
        Load frmMembresia
        frmMembresia.Show (1)
    End If
End Sub

Private Sub cmdModDir_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdModDir.Name) Then
        Exit Sub
    End If

    If (Me.ssdbDireccion.Rows > 0) Then
        frmAltaDir.bNvaDir = False
        frmAltaDir.Show (1)
        
        Me.ssdbDireccion.SetFocus
    End If
End Sub

Private Sub cmdModFalta_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdModFalta.Name) Then
        Exit Sub
    End If
    
    If (Me.ssdbAusencias.Rows > 0) Then
        frmAusencias.bNvaAus = False
        frmAusencias.Show (1)
    End If
    
    frmAltaSocios.ssdbAusencias.SetFocus
End Sub

Private Sub cmdModFam_Click()
    Dim nRen As Variant

    If Not ChecaSeguridad(Me.Name, Me.cmdModFam.Name) Then
        Exit Sub
    End If
    
    If (Me.ssdbFamiliares.Rows > 0) Then
        nRen = Me.ssdbFamiliares.Bookmark
    
        frmAltaFam.bNvoFam = False
        frmAltaFam.Show (1)
        
        If (Me.ssdbFamiliares.Rows > 0) Then
            Me.ssdbFamiliares.Bookmark = Me.ssdbFamiliares.AddItemBookmark(Me.ssdbFamiliares.AddItemRowIndex(nRen))
        
            Call ssdbFamiliares_RowColChange(Me.ssdbFamiliares.Rows, Me.ssdbFamiliares.Cols)
            
            Me.ssdbFamiliares.SetFocus
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Dim Respuesta As Integer

    If (Cambios) Then
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Registro de direcciones")
        
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

Private Sub dtpTitNacio_LostFocus()
    Me.txtTitEdad.Text = Format(Edad(Me.dtpTitNacio.Value), "#0.00")
End Sub

Private Sub Form_Activate()

    ConfigForma True
    
    Me.lblParameter.Caption = "LOADED"
    
'    Me.Visible = True
End Sub

Private Sub Form_Load()
    Me.lblParameter.Caption = ""
    
    HabilitaOpciones False
    ClearCtrlsAcc
    ClearCtrlsTit
    HabilitaOpciones True
    
'    nPosFam = 0
    
    InitColsFact
    Me.sstabSocios.Tab = 0
    
    If (Not bSocioNvo) Then
        If (sFormaAnterior = "frmDatosSocios") Then
            nTitCve = frmDatosSocios.adoSocios.Recordset.Fields("IdMember")
'            frmAltaSocios.sTitPaterno = frmDatosSocios.adoSocios.Recordset.Fields("Usuarios_Club!A_Paterno")
'            frmAltaSocios.sTitMaterno = frmDatosSocios.adoSocios.Recordset.Fields("Usuarios_Club!A_Materno")
'            frmAltaSocios.sTitNombre = frmDatosSocios.adoSocios.Recordset.Fields("Usuarios_Club!Nombre")
        ElseIf (sFormaAnterior = "frmSocios") Then
            nTitCve = frmSocios.ssdbBusca.Columns(3).Value
        Else
            nTitCve = 0
        End If
        
        LlenaTabs
        
        HabilitaTabs True
        
'        Me.dtpFechaUPago.Enabled = False
        
        Me.cmdMembresias.Enabled = True
        
        Me.txtFamilia.Enabled = False
    Else
        HabilitaTabs False
    End If
    
    InitVarAcc
    InitVarTit
    
    
    #If SqlServer_ Then
        Dim strSeguroVida As String
        strSeguroVida = ObtieneParametro("SEGURO_VIDA")
        
        If strSeguroVida <> "" Then
            bSeguroVida = CBool(strSeguroVida)
        Else
            bSeguroVida = False
        End If
    #End If
    
    If bSeguroVida = True Then
        fraBeneficiarios.Visible = True
        cmdAnexoBene.Visible = True
    Else
        fraBeneficiarios.Visible = False
        cmdAnexoBene.Visible = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set frmAltaSocios = Nothing
'
'    ConfigForma False
'
'    If (sFormaAnterior = "frmDatosSocios") Then
'        frmDatosSocios.WindowState = 2
'    End If
'
'    Dim tEmpresaResult As Integer
'
'    tEmpresaResult = TieneEmpresa
'
'    If tEmpresaResult > 0 Then
'        Dim msgResultado As Variant
'        Dim strMsg As String
'
'        Select Case tEmpresaResult
'            Case 1: strMsg = "La inscripción no cuenta con dirección de empresa. ¿Desea salir?"
'            Case 2: strMsg = "Faltan por acompletar los datos de la dirección de trabajo. ¿Desea salir?"
'        End Select
'
'        msgResultado = MsgBox(strMsg, vbYesNo, "KalaClub - Datos de empresa")
'
'        If msgResultado = vbNo Then
'            sstabSocios.Tab = 1
'            Cancel = 1
'        End If
'    End If
End Sub

'08/Dic/2011 UCM
Private Function TieneEmpresa() As Integer
    Dim rsDireccion As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String
    Dim sCondicion As String

    sCampos = "IdDireccion, IdMember, IdTipoDireccion, RazonSocial, Area "
    
    sTablas = "DIRECCIONES"
    
    sCondicion = "IdMember = " & nTitCve & " AND IdTipoDireccion = 2"
    
    InitRecordSet rsDireccion, sCampos, sTablas, sCondicion, "IdDireccion", Conn
    
    If rsDireccion.EOF Then
        TieneEmpresa = 1
    Else
         Do While Not rsDireccion.EOF
            If rsDireccion.Fields("RazonSocial").Value <> "" And rsDireccion.Fields("Area").Value <> "" Then
                TieneEmpresa = 0
                Exit Do
            End If
            
            TieneEmpresa = 2
            
            rsDireccion.MoveNext
        Loop
    End If
    
    rsDireccion.Close
    Set rsDireccion = Nothing
End Function

Private Sub ConfigForma(bEntrada As Boolean)
    If (bEntrada) Then
'        With Me
'            .Top = 0
'            .Left = 0
'        End With
    
'        'Centra el formulario
'        CENTRAFORMA MDIPrincipal, frmAltaSocios

        sTextToolBar = Trim(MDIPrincipal.StatusBar1.Panels.Item(1).Text)
        MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Altas, cambios y bajas de usuarios del club"
    Else
        MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
    End If
End Sub

'Asigna valores a las variables del accionista
Private Sub InitVarAcc()
    'Datos de la accion
    sSerie = Trim(Me.txtSerie.Text)
    sTipo = Trim(Me.txtTipo.Text)
    nNumero = Val(Me.txtNumero.Text)
    nAccCve = Val(Me.txtCveAccionista.Text)
End Sub

'Asigna valores a las variables del titular
Private Sub InitVarTit()
    'Datos del titular
    sTitPaterno = Trim(Me.txtTitPaterno.Text)
    sTitMaterno = Trim(Me.txtTitMaterno.Text)
    sTitNombre = Trim(Me.txtTitNombre.Text)
    sProf = Trim(Me.txtTitProf.Text)
    dTitNacio = Me.dtpTitNacio.Value
    dTitRegistro = Me.dtpTitRegistro.Value
    sTitCel = Trim(Me.txtTitCel.Text)
    sTitEmail = Trim(Me.txtTitEmail.Text)
    sTitPais = Trim(Me.txtPaisTit.Text)
    nPais = Val(Me.txtCvePais.Text)
    sTitTipo = Trim(Me.txtTipoTit.Text)
    nTipo = Val(Me.txtCveTipo.Text)
    bFemenino = Me.optFemenino.Value
    nTipoAnt = Val(Me.txtTipoTit.Text)
    nTipoNvo = Val(Me.txtTipoTit.Text)
    '11/10/2005 gpo
    sNoIns = Trim(Me.txtNoIns.Text)
    dFecUPago = Me.dtpFechaUPago.Value
End Sub

'Limpia las cajas de texto de los datos del accionista
Private Sub ClearCtrlsAcc()
    With Me
        .txtSerie.Text = ""
        .txtTipo.Text = ""
        .txtNumero.Text = ""
        .txtNombre.Text = ""
        .txtCveAccionista.Text = ""
        .txtTel1.Text = ""
        .txtTel2.Text = ""
    End With
End Sub

'Limpia los ctrls de los datos del titular
Private Sub ClearCtrlsTit()
    With Me
        .txtTitCve.Text = BuscaCve("Usuarios_Club", "IdMember")
        .txtFamilia.Text = LeeUltReg("usuarios_Club", "NoFamilia") + 1
        .txtTitPaterno.Text = ""
        .txtTitMaterno.Text = ""
        .txtTitNombre.Text = ""
        .txtTitProf.Text = ""
        .dtpTitNacio.Value = Format(Date, "dd/mm/yyyy")
        .dtpTitRegistro.Value = Format(Date, "dd/mm/yyyy")
        .txtTitCel.Text = ""
        .txtTitEmail.Text = ""
        .txtPaisTit.Text = ""
        .txtCvePais.Text = ""
        .txtTipoTit.Text = ""
        .txtCveTipo.Text = ""
'        .optFemenino.Value = True
        .txtSecuencial.Text = ""
        .dtpFechaUPago.Value = Format(CDate("01/01/1900"), "dd/mm/yyyy")
        'gpo 11/10/2005
        .txtNoIns.Text = ""
    End With
End Sub

Private Sub optMembresia_Click()
    If (Me.optMembresia.Enabled) Then
        If (bSocioNvo) Then
            ClearCtrlsAcc
            ClearCtrlsTit
            
            InitVarAcc
            InitVarTit
        Else
            'Cambia de membresia a propietario o rentista
            If (Not bMemb) Then
                ClearCtrlsAcc
                'InitVarAcc
                
                'Deshabilita los ctrls de la accion
                Me.txtSerie.Enabled = False
                Me.txtTipo.Enabled = False
                Me.txtNumero.Enabled = False
                
                lCambio = True
            End If
        End If
        
        Me.cmdAyuda.Enabled = False
        Me.txtSerie.Enabled = False
        Me.txtTipo.Enabled = False
        Me.txtNumero.Enabled = False
        
        'Cambia las banderas del tipo de uso de la accion
        bMemb = True
        bProp = False
        bRent = False
        
        Me.txtTitPaterno.Enabled = True
        Me.txtTitMaterno.Enabled = True
        Me.txtTitNombre.Enabled = True
        
        If (Me.lblParameter.Caption = "LOADED") Then
            Me.txtTitPaterno.SetFocus
        End If
    Else
        lCambio = False
    End If
End Sub

Private Sub optPropietario_Click()
    If (Me.optPropietario.Enabled) Then
        If (bSocioNvo) Then
            ClearCtrlsAcc
            ClearCtrlsTit
            
            InitVarAcc
            InitVarTit
        Else
            If (Not bProp) Then
                'Cambia de rentista a propietario
                If (bRent) Then
                    ClearCtrlsAcc
                    ClearCtrlsTit
                    
'                    InitVarAcc
'                    InitVarTit
                End If
                
                'En el cambio de membresia o rentista a propietario
                nAccCve = 0
                
                lCambio = True
            End If
        End If
        
        'Cambia las banderas del tipo de uso de la accion
        bProp = True
        bRent = False
        bMemb = False
              
        'Habilita los ctrls de la accion
        Me.txtSerie.Enabled = True
        Me.txtTipo.Enabled = True
        Me.txtNumero.Enabled = True
        Me.cmdAyuda.Enabled = True
    
        'Deshabilita los ctrls del nombre del titular
        Me.txtTitPaterno.Enabled = False
        Me.txtTitMaterno.Enabled = False
        Me.txtTitNombre.Enabled = False
    
        If (Me.lblParameter.Caption = "LOADED") Then
            Me.txtSerie.SetFocus
        End If
    Else
        lCambio = False
    End If
End Sub

Private Sub optRentista_Click()
    If (Me.optRentista.Enabled) Then
        If (bSocioNvo) Then
            ClearCtrlsAcc
            ClearCtrlsTit
            
            InitVarAcc
            InitVarTit
        Else
            'Cambia de membresia a rentista
            If (bMemb) Then
                nAccCve = 0
                lCambio = True
                
            'Cambia de propietario a rentista
            ElseIf (bProp) Then
                ClearCtrlsTit
                InitVarTit
                
                lCambio = True
            End If
        End If
        
        'Cambia las banderas del tipo de uso de la accion
        bRent = True
        bProp = False
        bMemb = False

        'Habilita los ctrls de la accion
        Me.txtSerie.Enabled = True
        Me.txtTipo.Enabled = True
        Me.txtNumero.Enabled = True
        Me.cmdAyuda.Enabled = True
        
        'Habilita los ctrls del nombre del titular
        Me.txtTitPaterno.Enabled = True
        Me.txtTitMaterno.Enabled = True
        Me.txtTitNombre.Enabled = True
        
        If (Me.lblParameter.Caption = "LOADED") Then
            Me.txtSerie.SetFocus
        End If
    Else
        lCambio = False
    End If
End Sub

Private Function ChecaDatos()
    Dim sCond As String
    Dim sCamp As String

    ChecaDatos = False

    If ((Me.optPropietario.Value) Or (Me.optRentista.Value)) Then
        If (Trim(Me.txtSerie.Text) = "") Then
            MsgBox "Se debe seleccionar una serie para la acción.", vbExclamation, "KalaSystems"
            Me.txtSerie.SetFocus
            Exit Function
        End If
        
        If ((Trim(Me.txtTipo.Text) = "")) Then
            MsgBox "Se debe seleccionar un tipo de serie para la acción.", vbExclamation, "KalaSystems"
            Me.txtTipo.SetFocus
            Exit Function
        End If
        
        If ((Trim(Me.txtNumero.Text) = "")) Then
            MsgBox "Se debe seleccionar un número de accion.", vbExclamation, "KalaSystems"
            Me.txtNumero.SetFocus
            Exit Function
        End If
        
        If (IsNumeric(Me.txtNumero.Text)) Then
            If (Val(Me.txtNumero.Text) <= 0) Then
                MsgBox "El número de la acción es incorrecto.", vbExclamation, "KalaSystems"
                Me.txtNumero.SetFocus
                Exit Function
            End If
        Else
            MsgBox "El número de la acción es incorrecto.", vbExclamation, "KalaSystems"
            Me.txtNumero.SetFocus
            Exit Function
        End If
        
        sCamp = "Serie, Tipo, Numero, IdPropietario"
        
        sCond = "Serie='" & Trim(Me.txtSerie.Text) & "' AND "
        sCond = sCond & "Tipo='" & Trim(Me.txtTipo.Text) & "' AND "
        sCond = sCond & "Numero=" & Val(Me.txtNumero.Text) & " AND "
        sCond = sCond & "IdPropietario>0"
        
        'Revisa que exista el numero de la accion en la tabla de acciones asignadas
        If (Not ExisteXValor(sCamp, "Titulos", sCond, Conn, "")) Then
            MsgBox "La acción seleccionada no existe o no está asignada.", vbExclamation, "KalaSystems"
            Me.txtSerie.SetFocus
            Exit Function
        End If
        
        'Revisa que no se encuentre en uso la accion
        sCamp = "Serie, Tipo, Numero"
        
        sCond = "Serie='" & Trim(Me.txtSerie.Text) & "' AND "
        sCond = sCond & "Tipo='" & Trim(Me.txtTipo.Text) & "' AND "
        sCond = sCond & "Numero=" & Val(Me.txtNumero.Text)
        
        'Revisa que no se encuentre en uso la accion
        If (Not ExisteXValor(sCamp, "Titulos", sCond, Conn, "")) Then
            MsgBox "La acción seleccionada no existe o no está asignada.", vbExclamation, "KalaSystems"
            Me.txtSerie.SetFocus
            Exit Function
        End If
    End If
    
    If ((Me.optRentista.Value) Or (Me.optMembresia.Value)) Then
        If (Trim(Me.txtTitPaterno.Text) = "") Then
            MsgBox "Se debe escribir el apellido paterno.", vbExclamation, "KalaSystems"
            Me.txtTitPaterno.SetFocus
            Exit Function
        End If
        
        If (Trim(Me.txtTitMaterno.Text) = "") Then
            MsgBox "Se debe escribir el apellido materno.", vbExclamation, "KalaSystems"
            Me.txtTitMaterno.SetFocus
            Exit Function
        End If
        
        If (Trim(Me.txtTitNombre.Text) = "") Then
            MsgBox "Se debe escribir el nombre del titular.", vbExclamation, "KalaSystems"
            Me.txtTitNombre.SetFocus
            Exit Function
        End If
    End If
    
    If CDate(Me.dtpTitNacio.Value) > Date Then
        MsgBox "La fecha de Nacimiento no puede ser mayor a la fecha actual!", vbExclamation, "KalaSystems"
        Me.dtpTitNacio.SetFocus
        Exit Function
    End If
    
    If DateDiff("yyyy", CDate(Me.dtpTitNacio.Value), Date) > 120 Then
        MsgBox "La edad del titular no puede ser mayor de 120 años!", vbExclamation, "KalaSystems"
        Me.dtpTitNacio.SetFocus
        Exit Function
    End If
    
    If Int(DateDiff("d", CDate(Me.dtpTitNacio.Value), Date) / 365.25) < 16 Then
        MsgBox "La edad del titular no puede ser menor de 16 años!", vbExclamation, "KalaSystems"
        Me.dtpTitNacio.SetFocus
        Exit Function
    End If
    
    If (Trim(Me.txtTipoTit.Text) = "") Then
        MsgBox "Se debe seleccionar un tipo de titular.", vbExclamation, "KalaSystems"
        Me.txtCveTipo.SetFocus
        Exit Function
    End If
    
    If (Trim(Me.txtPaisTit.Text) = "") Then
        MsgBox "Se debe seleccionar un país de origen.", vbExclamation, "KalaSystems"
        Me.txtCvePais.SetFocus
        Exit Function
    End If
    
    If Not ValidaEmailAddress(Me.txtTitEmail.Text) Then
        MsgBox "El correo electrónico no es válido", vbExclamation, "Verifique"
        Me.txtTitEmail.SetFocus
        Exit Function
    End If
    
    If Trim(txtTitProf.Text) = "" Then
        MsgBox "Favor de ingresar la profesión del titular.", vbExclamation, "Verifique"
        Me.txtTitProf.SetFocus
        Exit Function
    End If
    
    If Trim(txtTitCel.Text) = "" Then
        MsgBox "Favor de capturar un número de celular.", vbExclamation, "Verifique"
        Me.txtTitCel.SetFocus
        Exit Function
    End If
    
    If Not (optFemenino.Value Xor optMasculino.Value) Then
        MsgBox "Favor de seleccionar el sexo del titular.", vbExclamation, "Verifique"
        'Me.optFemenino.SetFocus
        Exit Function
    End If
    
    If dtpFechaUPago.Value = CDate("01/01/1900") Then
        MsgBox "Favor de seleccionar la fecha de uso de instalaciones.", vbExclamation, "Verifique"
        Me.dtpFechaUPago.SetFocus
        Exit Function
    End If
    
    ChecaDatos = True
End Function

Private Sub ssdbFacturas_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    With Me.ssdbFacturas
        If (.Rows > 0) Then
            Me.txtObs.Text = .Columns(10).Text
            
            MostrarDetalle
            MostrarTotal
        End If
    End With
End Sub

Private Sub ssdbFamiliares_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    With Me.ssdbFamiliares
        If (.Rows > 0) Then
'            nPosFam = Me.ssdbFamiliares.Bookmark
        
            Me.txtCveFam.Text = .Columns(3).Value
            Me.txtSec.Text = .Columns(4).Value
            Me.txtTipoUser.Text = .Columns(5).Text
            Me.txtNacio.Text = Format(.Columns(6).Value, "dd / mmm / yyyy")
            Me.txtIngreso.Text = Format(.Columns(7).Value, "dd / mmm / yyyy")
            Me.txtProf.Text = .Columns(8).Text
            Me.txtCel.Text = .Columns(9).Text
            Me.txtEmail.Text = .Columns(10).Text
            Me.txtSexo.Text = .Columns(11).Text
            Me.txtPais.Text = .Columns(12).Text
            Me.txtFamFoto.Text = .Columns(13).Value
            
            If (Dir(sG_RutaFoto & "\" & Trim$(.Columns(13).Value) & ".jpg") <> "") Then
                Me.imgFamiliar.Picture = LoadPicture(sG_RutaFoto & "\" & Trim$(.Columns(13).Value) & ".jpg")
            Else
                Me.imgFamiliar.Picture = LoadPicture("")
            End If
        End If
    End With
End Sub

Private Sub sstabSocios_Click(PreviousTab As Integer)
    If Me.sstabSocios.Tab = 6 Then
        Llena_List_Usuarios
        Me.ssdbgAcceso.RemoveAll
        Me.dtpFecIni.Value = Date
        Me.dtpFecFin.Value = Date
    End If
End Sub

Private Sub txtCtrlEmer_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 1, 3, 4, 5, 6, 7
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case 8, 9
            Select Case KeyAscii
                Case 8 ' Tecla backspace
                    KeyAscii = KeyAscii
                Case 46 'punto decimal
                    If InStr(1, Me.txtCtrlEmer(Index).Text, ".", 1) > 0 Then
                        KeyAscii = 0
                    End If
                    
                    KeyAscii = KeyAscii
                Case 48 To 57 ' Numeros del 0 al 9
                    KeyAscii = KeyAscii
                Case Else
                    KeyAscii = 0
                    'MsgBox "Solo admite números", vbInformation, "Ver"
            End Select
        Case 16
            Select Case KeyAscii
                Case 8 ' Tecla backspace
                    KeyAscii = KeyAscii
                Case 48 To 57 ' Numeros del 0 al 9
                    KeyAscii = KeyAscii
                Case Else
                    KeyAscii = 0
                    'MsgBox "Solo admite números", vbInformation, "Productos"
            End Select
    End Select
End Sub

Private Sub txtCvePais_LostFocus()
    Dim sCond As String

    If (Trim(Me.txtCvePais.Text) <> "") Then
        If (IsNumeric(Me.txtCvePais.Text)) Then
            sCond = "IdPais=" & Val(Me.txtCvePais.Text)
            Me.txtPaisTit.Text = LeeXValor("Pais", "Paises", sCond, "Pais", "s", Conn)
            
            If (Trim(Me.txtPaisTit.Text) = "VACIO") Then
                MsgBox "El país seleccionado no existe en la base de datos.", vbExclamation, "KalaSystems"
                Me.txtPaisTit.Text = ""
                Me.txtCvePais.Text = ""
                Me.txtCvePais.SetFocus
            End If
        Else
            MsgBox "La clave del país es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtPaisTit.Text = ""
            Me.txtCvePais.Text = ""
            Me.txtCvePais.SetFocus
        End If
    End If
    
    Me.txtCvePais.REFRESH
End Sub

Private Sub txtCveTipo_LostFocus()
    Dim sCond As String

    If (Trim(Me.txtCveTipo.Text) <> "") Then
        If (IsNumeric(Me.txtCveTipo.Text)) Then
            sCond = "IdTipoUsuario=" & Val(Me.txtCveTipo.Text) & " AND Parentesco='TI'"
            If (Me.optPropietario.Value) Then
                sCond = sCond & " AND Tipo='PROPIETARIO'"
            ElseIf (Me.optRentista.Value) Then
                sCond = sCond & " AND Tipo='RENTISTA'"
            ElseIf (Me.optMembresia.Value) Then
                sCond = sCond & " AND Tipo='MEMBRESIA'"
            End If
            
            
            '02/10/2010 gpo
            
            'sCond = sCond & "AND Vigencia >= '" & Format(Date, "mm/dd/yyyy") & "'"
            
            
            Me.txtTipoTit.Text = LeeXValor("Descripcion", "Tipo_Usuario", sCond, "Descripcion", "s", Conn)
            
            If (Trim(Me.txtTipoTit.Text) = "VACIO") Then
                MsgBox "El tipo de usuario seleccionado no existe o no" & Chr(13) & "corresponde con las características del usuario.", vbExclamation, "KalaSystems"
                Me.txtTipoTit.Text = ""
                Me.txtCveTipo.Text = ""
                Me.txtCveTipo.SetFocus
            End If
        Else
            MsgBox "La clave del tipo de usuario es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtTipoTit.Text = ""
            Me.txtCveTipo.Text = ""
            Me.txtCveTipo.SetFocus
        End If
    End If
    
    Me.txtCveTipo.REFRESH
End Sub

Private Sub txtNumero_LostFocus()
    If ((Trim(Me.txtSerie.Text) <> "") And (Trim(Me.txtTipo.Text) <> "") And (Trim(Me.txtNumero.Text) <> "")) Then
        BuscaDatosAcc
    
        If (Me.optRentista.Value) Then
            Me.txtTitPaterno.SetFocus
        ElseIf (Me.optPropietario.Value) Then
            Me.txtTitPaterno.Text = sTitPaterno
            Me.txtTitMaterno.Text = sTitMaterno
            Me.txtTitNombre.Text = sTitNombre
            
            If (Me.txtTitCve.Enabled) Then
                Me.txtTitCve.SetFocus
            End If
        End If
    End If
End Sub

Private Function GuardaDatos() As Boolean
    Const DATOSTITULAR = 18
    Const DATOSADICIONALES = 5
    Const DATOSALTA = 17
    Const DATOSTZONE = 4
    Dim bCreado As Boolean
    Dim mFieldsTit(DATOSTITULAR) As String
    Dim mValuesTit(DATOSTITULAR) As Variant
    Dim mFieldsAdic(DATOSADICIONALES) As String
    Dim mValuesAdic(DATOSADICIONALES) As Variant
    Dim mFieldsAlta(DATOSALTA) As String
    Dim mValuesAlta(DATOSALTA) As Variant
    Dim mFieldsZone(DATOSTZONE) As String
    Dim mValuesZone(DATOSTZONE) As Variant
    Dim sTipoUsoAccion As String
    Dim InitTrans As Long
    Dim nSecTit As Long
    Dim nSecTitNuevo As Long
    Dim nTor As Long

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Usuarios_Club
    mFieldsTit(0) = "IdMember"
    mFieldsTit(1) = "Nombre"
    mFieldsTit(2) = "A_Paterno"
    mFieldsTit(3) = "A_Materno"
    mFieldsTit(4) = "FechaNacio"
    mFieldsTit(5) = "Sexo"
    mFieldsTit(6) = "IdPais"
    mFieldsTit(7) = "IdTipoUsuario"
    mFieldsTit(8) = "Email"
    mFieldsTit(9) = "Celular"
    mFieldsTit(10) = "Profesion"
    mFieldsTit(11) = "FechaIngreso"
    mFieldsTit(12) = "IdTitular"
    mFieldsTit(13) = "NoFamilia"
    mFieldsTit(14) = "FotoFile"
    mFieldsTit(15) = "UFechaPago"
    mFieldsTit(16) = "NumeroFamiliar"
    '11/10/2005 gpo
    mFieldsTit(17) = "Inscripcion"
    
    'Campos de la tabla de Usuarios_Titulo
    mFieldsAdic(0) = "IdMember"
'    mFieldsAdic(1) = "IdTipoPago"
    mFieldsAdic(1) = "Serie"
    mFieldsAdic(2) = "Tipo"
    mFieldsAdic(3) = "Numero"
    mFieldsAdic(4) = "IdTipoUsoAccion"
    
    'Campos de la tabla Altas
    mFieldsAlta(0) = "IdMember"
    mFieldsAlta(1) = "Nombre"
    mFieldsAlta(2) = "A_Paterno"
    mFieldsAlta(3) = "A_Materno"
    mFieldsAlta(4) = "FechaNacio"
    mFieldsAlta(5) = "Sexo"
    mFieldsAlta(6) = "IdPais"
    mFieldsAlta(7) = "IdTipoUsuario"
    mFieldsAlta(8) = "Email"
    mFieldsAlta(9) = "Celular"
    mFieldsAlta(10) = "Profesion"
    mFieldsAlta(11) = "FechaIngreso"
    mFieldsAlta(12) = "IdTitular"
    mFieldsAlta(13) = "IdUsuario"
    mFieldsAlta(14) = "Fecha"
    mFieldsAlta(15) = "NoFamilia"
    mFieldsAlta(16) = "FotoFile"
    
    'Valores para la tabla de Usuarios_Club
    If (bSocioNvo) Then
        mValuesTit(0) = LeeUltReg("Usuarios_Club", "IdMember") + 1
        mValuesTit(13) = LeeUltReg("Usuarios_Club", "NoFamilia") + 1
        mValuesTit(14) = LeeUltReg("FolioFoto", "idFoto") + 1
        mValuesTit(16) = 1
    Else
        mValuesTit(0) = Val(Me.txtTitCve.Text)
        mValuesTit(13) = Val(Me.txtFamilia.Text)
        mValuesTit(14) = Val(Me.txtImagen.Text)
        
        mValuesTit(16) = 1
    End If
    
    #If SqlServer_ Then
        mValuesTit(1) = Trim(UCase(Me.txtTitNombre.Text))
        mValuesTit(2) = Trim(UCase(Me.txtTitPaterno.Text))
        mValuesTit(3) = Trim(UCase(Me.txtTitMaterno.Text))
        mValuesTit(4) = Format(Me.dtpTitNacio.Value, "yyyymmdd")
        mValuesTit(5) = IIf(Me.optFemenino.Value, "F", "M")
        mValuesTit(6) = Val(Me.txtCvePais.Text)
        mValuesTit(7) = Val(Me.txtCveTipo.Text)
        mValuesTit(8) = Trim(Me.txtTitEmail.Text)
        mValuesTit(9) = Trim(Me.txtTitCel.Text)
        mValuesTit(10) = Trim(UCase(Me.txtTitProf.Text))
        mValuesTit(11) = Format(Me.dtpTitRegistro.Value, "yyyymmdd")
        mValuesTit(12) = mValuesTit(0)
        mValuesTit(15) = Format(Me.dtpFechaUPago.Value, "yyyymmdd")
        mValuesTit(17) = Trim(Me.txtNoIns.Text)
    #Else
        mValuesTit(1) = Trim(UCase(Me.txtTitNombre.Text))
        mValuesTit(2) = Trim(UCase(Me.txtTitPaterno.Text))
        mValuesTit(3) = Trim(UCase(Me.txtTitMaterno.Text))
        mValuesTit(4) = Format(Me.dtpTitNacio.Value, "dd/mm/yyyy")
        mValuesTit(5) = IIf(Me.optFemenino.Value, "F", "M")
        mValuesTit(6) = Val(Me.txtCvePais.Text)
        mValuesTit(7) = Val(Me.txtCveTipo.Text)
        mValuesTit(8) = Trim(Me.txtTitEmail.Text)
        mValuesTit(9) = Trim(Me.txtTitCel.Text)
        mValuesTit(10) = Trim(UCase(Me.txtTitProf.Text))
        mValuesTit(11) = Format(Me.dtpTitRegistro.Value, "dd/mm/yyyy")
        mValuesTit(12) = mValuesTit(0)
        mValuesTit(15) = Format(Me.dtpFechaUPago.Value, "dd/mm/yyyy")
        mValuesTit(17) = Trim(Me.txtNoIns.Text)
    #End If
    
    'Valores para la tabla Usuarios_Titulo
    mValuesAdic(0) = mValuesTit(0)
    
'    If (Me.cbComoPaga.ListIndex < 0) Then
'        mValuesAdic(1) = LeeXValor("IdTipoPago, Descripcion", "Tipo_Pago", "Descripcion='" & Trim(Me.cbComoPaga.Text) & "'", "IdTipoPago", "n", Conn)
'    Else
'        mValuesAdic(1) = Me.cbComoPaga.ItemData(Me.cbComoPaga.ListIndex)
'    End If
    
    If (Not Me.optMembresia.Value) Then
        mValuesAdic(1) = Trim(UCase(Me.txtSerie.Text))
        mValuesAdic(2) = Trim(UCase(Me.txtTipo.Text))
        mValuesAdic(3) = Val(Me.txtNumero.Text)
        
        sTipoUsoAccion = IIf(Me.optPropietario.Value, "PROPIETARIO", "RENTISTA")
    Else
        mValuesAdic(1) = ""
        mValuesAdic(2) = ""
        mValuesAdic(3) = 0
        
        sTipoUsoAccion = "MEMBRESIA"
    End If
    
    mValuesAdic(4) = LeeXValor("IdTipoUsoAccion, Descripcion", "Tipo_Uso_Accion", "Descripcion='" & sTipoUsoAccion & "'", "IdTipoUsoAccion", "n", Conn)
    
    'Valores para la tabla Altas
    mValuesAlta(0) = mValuesTit(0)
    mValuesAlta(1) = mValuesTit(1)
    mValuesAlta(2) = mValuesTit(2)
    mValuesAlta(3) = mValuesTit(3)
    mValuesAlta(4) = mValuesTit(4)
    mValuesAlta(5) = mValuesTit(5)
    mValuesAlta(6) = mValuesTit(6)
    mValuesAlta(7) = mValuesTit(7)
    mValuesAlta(8) = mValuesTit(8)
    mValuesAlta(9) = mValuesTit(9)
    mValuesAlta(10) = mValuesTit(10)
    mValuesAlta(11) = mValuesTit(11)
    mValuesAlta(12) = mValuesTit(12)
    mValuesAlta(13) = LeeXValor("IdUsuario", "Usuarios_Sistema", "Login_Name='" & sDB_User & "'", "IdUsuario", "n", Conn)
    #If SqlServer_ Then
        mValuesAlta(14) = Format(Date, "yyyymmdd")
    #Else
        mValuesAlta(14) = Format(Date, "dd/mm/yyyy")
    #End If
    
    mValuesAlta(15) = mValuesTit(13)
    mValuesAlta(16) = mValuesTit(14)
    
    If (bSocioNvo) Then
        'Inicia el registro de los datos en las tablas
        InitTrans = Conn.BeginTrans
    
        'Registra los datos del titular
        
        If (AgregaRegistro("Usuarios_Club", mFieldsTit, DATOSTITULAR, mValuesTit, Conn)) Then
            
            'MsgBox "Usuarios titulo", vbOKCancel, ""            'Registra los datos adicionales
            If (AgregaRegistro("Usuarios_Titulo", mFieldsAdic, DATOSADICIONALES, mValuesAdic, Conn)) Then
            
                                
                'Registra las fechas de inicio para cuotas de Mtto
                If (RegistraFechasMtto(CLng(mValuesTit(0)), Val(Me.txtCveTipo.Text), Me.dtpFechaUPago.Value)) Then
                
                    
                    'Registra los datos en la tabla de altas
                    If (AgregaRegistro("Altas", mFieldsAlta, DATOSALTA, mValuesAlta, Conn)) Then
                    
                        'Campos de la tabla Time_Zone_User
                        mFieldsZone(0) = "IdReg"
                        mFieldsZone(1) = "NoFamilia"
                        mFieldsZone(2) = "IdMember"
                        mFieldsZone(3) = "IdTimeZone"
                    
                        'Valores para la tabla Time_Zone_Users
                        mValuesZone(0) = LeeUltReg("Time_Zone_Users", "IdReg") + 1
                        mValuesZone(1) = mValuesTit(13)
                        mValuesZone(2) = mValuesTit(0)
                        mValuesZone(3) = 0
                    
                        
                        'Registra los datos en las zonas horarias
                        If (AgregaRegistro("Time_Zone_Users", mFieldsZone, DATOSTZONE, mValuesZone, Conn)) Then
                    
                            'Incrementa el contador en la tabla de ALTAS
                            
                            If (Altas_Bajas(True)) Then
                                
                                
                                'Registra la clave del titular en la tabla de secuenciales
                                nSecTit = AsignaSec(CLng(mValuesTit(0)), False)
                                nSecTitNuevo = AsignaSecNuevo(CLng(mValuesTit(0)), False)
                                
                                'Si se registro correctamente
                                If (nSecTit > 0 And nSecTitNuevo > 0) Then
                                
                                    
                                    'Incrementa el ultimo numero de folio de las fotos
                                    If (IncFolio("FolioFoto", "IdFoto", 1)) Then
                                    
                                        'Baja a disco los nuevos datos
                                        Conn.CommitTrans
                                    
                                        'Deshabilita la caja de texto del numero de familia
                                        Me.txtFamilia.Enabled = False
                                        
                                        'Refresca el numero de usuario que se asigno
                                        Me.txtTitCve.Text = mValuesTit(0)
                                        Me.txtTitCve.REFRESH
                                        
                                        'gpo 16/10/2005
                                        'Me.txtCveFam.Text = mValuesTit(13)
                                        'Me.txtCveFam.REFRESH
                                        
                                        'gpo 02/04/2006
                                        Me.txtFamilia.Text = mValuesTit(13)
                                        Me.txtFamilia.REFRESH
                                        
                                        'Actualiza el numero de secuencial
                                        Me.txtSecuencial.Text = nSecTit
                                        Me.txtSecuencial.REFRESH
                                        
                                        'Actualiza el numero de fotografia
                                        Me.txtImagen.Text = mValuesTit(14)
                                        Me.txtImagen.REFRESH
                                        
                                        MsgBox "Los datos se dieron de alta con la clave: " & mValuesTit(0) & vbLf & "y # Familia: " & mValuesTit(13) & "." & vbLf & "Favor de registrar la huella y/o Tarjeta del Usuario. ", vbInformation, "KalaSystems"
                            
                                        bSocioNvo = False
                                        GuardaDatos = True
                                        
                                        
                                    End If
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
            
            MsgBox "El registro no fue completado.", vbCritical, "KalaSystems"
        End If
    Else

        If (Val(Me.txtTitCve.Text) > 0) Then
            InitTrans = Conn.BeginTrans

            'Actualiza los datos del titular
            If (CambiaReg("Usuarios_Club", mFieldsTit, DATOSTITULAR, mValuesTit, "IdMember=" & Val(Me.txtTitCve.Text), Conn)) Then
            
                'Actualiza los datos adicionales
                If (CambiaReg("Usuarios_Titulo", mFieldsAdic, DATOSADICIONALES, mValuesAdic, "IdMember=" & Val(Me.txtTitCve.Text), Conn)) Then
                
                    If (nTipoAnt <> nTipoNvo) Then
                        If (ActConcFact(CInt(mValuesTit(0)), nTipoAnt, nTipoNvo)) Then
                            
                            Conn.CommitTrans
                            MsgBox "Los datos de la clave: " & mValuesTit(0) & " se actualizaron correctamente", vbInformation, "Modificación de datos"
                            GuardaDatos = True
                            
                        Else
                            Conn.RollbackTrans
                            MsgBox "No se realizaron los cambios.", vbCritical, "KalaSystems"
                        End If
                    Else
                        Conn.CommitTrans
                        
                        MsgBox "Los datos se actualizaron correctamente.", vbInformation, "KalaSystems"
                        GuardaDatos = True
                
                        
                    End If
                Else
                    MsgBox "No se realizaron los cambios del usuario en los títulos.", vbCritical, "KalaSystems"
                    If InitTrans > 0 Then
                        Conn.RollbackTrans
                    End If
                End If
                
            Else
                MsgBox "No se realizaron los cambios del titular.", vbCritical, "KalaSystems"
                If InitTrans > 0 Then
                    Conn.RollbackTrans
                End If
            End If

        'Else
        '    MsgBox "Para dar de alta nuevos datos cierre la ventana y presione el botón: Nuevos datos", vbExclamation, "Agregar nuevos datos"
        End If

    End If
            'Registro en Torniquetes NITEGEN
           'nTor = (68 * 16777216) + CLng(nSecTit)
           
           'Do While nErrCode <> 805372429
             'nErrCode = AgregaAcceso(CLng(mValuesTit(0)), mValuesTit(1) & " " & mValuesTit(2) & " " & mValuesTit(3), nTor)

               ' MsgBox "No se pudo registrar el usuario en torniquetes,Se va a volver a intentar..."
            ' Loop
    'Pasa a mayusculas el contenido de las cajas de texto
    CambiaAMayusculas
End Function

'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    If (Not Me.optMembresia.Value) Then
        If (sSerie <> Trim(Me.txtSerie.Text)) Then
            Cambios = True
            Exit Function
        End If
        
        If (sTipo <> Trim(Me.txtTipo.Text)) Then
            Cambios = True
            Exit Function
        End If
        
        If (nNumero <> Val(Me.txtNumero.Text)) Then
            Cambios = True
            Exit Function
        End If
    End If

    If (sTitPaterno <> Trim(Me.txtTitPaterno.Text)) Then
        Cambios = True
        Exit Function
    End If
    
    If (sTitMaterno <> Trim(Me.txtTitMaterno.Text)) Then
        Cambios = True
        Exit Function
    End If
    
    If (sTitNombre <> Trim(Me.txtTitNombre.Text)) Then
        Cambios = True
        Exit Function
    End If
    
    If (sProf <> Trim(Me.txtTitProf.Text)) Then
        Cambios = True
        Exit Function
    End If
    
    If (dTitNacio <> Me.dtpTitNacio.Value) Then
        Cambios = True
        Exit Function
    End If
    
    If (dTitRegistro <> Me.dtpTitRegistro.Value) Then
        Cambios = True
        Exit Function
    End If
    
    If (sTitCel <> Trim(Me.txtTitCel.Text)) Then
        Cambios = True
        Exit Function
    End If
    
    If (sTitEmail <> Trim(Me.txtTitEmail.Text)) Then
        Cambios = True
        Exit Function
    End If
    
    If (sTitPais <> Trim(Me.txtPaisTit.Text)) Then
        Cambios = True
        Exit Function
    End If
    
    If (sTitTipo <> Trim(Me.txtTipoTit.Text)) Then
        nTipoNvo = Val(Me.txtCveTipo.Text)
        Cambios = True
        Exit Function
    End If
    
    If (bFemenino <> Me.optFemenino.Value) Then
        Cambios = True
        Exit Function
    End If
    
    '11/10/2005 gpo
    If (sNoIns <> Trim(Me.txtNoIns.Text)) Then
        Cambios = True
        Exit Function
    End If
    
    If (dFecUPago <> Me.dtpFechaUPago.Value) Then
        Cambios = True
        Exit Function
    End If
    
    
    'Cambios en la forma de uso de la accion
    If (lCambio) Then
        Cambios = True
        Exit Function
    End If
    
    Cambios = False
End Function

Private Sub CambiaAMayusculas()
    With Me
        .txtTitPaterno.Text = UCase(.txtTitPaterno.Text)
        .txtTitPaterno.REFRESH
        
        .txtTitMaterno.Text = UCase(.txtTitMaterno.Text)
        .txtTitMaterno.REFRESH
        
        .txtTitNombre.Text = UCase(.txtTitNombre.Text)
        .txtTitNombre.REFRESH
        
        .txtTitProf.Text = UCase(.txtTitProf.Text)
        .txtTitProf.REFRESH
    End With
End Sub

'Busca los datos del accionista
Private Sub BuscaDatosAcc()
    Dim rsDatosAcc As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String
    Dim sCondicion As String

    sCampos = "Titulos.Serie, Titulos.Tipo, Titulos.Numero, Titulos.IdPropietario, "
    sCampos = sCampos & "Accionistas.A_Paterno, Accionistas.A_Materno, Accionistas.Nombre, "
    sCampos = sCampos & "Accionistas.Telefono_1, Accionistas.Telefono_2"
    
    sTablas = "Titulos LEFT JOIN Accionistas ON Titulos.IdPropietario=Accionistas.IdPropTitulo"
    
    sCondicion = "Titulos.Serie='" & Me.txtSerie.Text & "' AND "
    sCondicion = sCondicion & "Titulos.Tipo='" & Me.txtTipo.Text & "' AND "
    sCondicion = sCondicion & "Titulos.Numero=" & Val(Me.txtNumero.Text) & " AND "
    sCondicion = sCondicion & "Titulos.IdPropietario>0"
    
    InitRecordSet rsDatosAcc, sCampos, sTablas, sCondicion, "", Conn

    With rsDatosAcc
        If (.RecordCount > 0) Then
            'Datos para el titular o accionista
            sTitPaterno = .Fields("A_Paterno")
            sTitMaterno = .Fields("A_Materno")
            sTitNombre = .Fields("Nombre")
            
            'Datos de la accion
            Me.txtSerie.Text = .Fields("Serie")
            Me.txtTipo.Text = .Fields("Tipo")
            Me.txtNumero.Text = .Fields("Numero")
            
            'Datos del accionista
            Me.nAccCve = .Fields("IdPropietario")
            Me.txtCveAccionista.Text = nAccCve
            Me.txtNombre.Text = Trim(sTitPaterno) & " " & Trim(sTitMaterno) & " " & Trim(sTitNombre)
            Me.txtTel1.Text = .Fields("Telefono_1")
            Me.txtTel2.Text = .Fields("Telefono_2")
        End If
        
        .Close
    End With

    Set rsDatosAcc = Nothing
End Sub

Private Sub LlenaTabs()
    Screen.MousePointer = vbHourglass
    
    LlenaGrales
    LlenaRecordsets
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub LlenaRecordsets()
    frmAltaDir.LlenaDirs
    frmAltaFam.LlenaFam
    frmAltaRenta.LlenaRenta
    frmAusencias.LlenaAusencias
    LlenaFacts
    LlenaDatosEmergencia
End Sub

Private Sub LlenaFacts()
    Dim rsFacts As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String
    Dim sCondicion As String

    Me.ssdbFacturas.RemoveAll

    sCampos = "NumeroFactura, FolioCFD AS Folio, SerieCFD As Serie, FechaFactura, NombreFactura, CalleFactura, "
    sCampos = sCampos & "ColoniaFactura, DelFactura, CiudadFactura, "
    sCampos = sCampos & "EstadoFactura, RFC, Observaciones "
    
    sTablas = "Facturas"
    
    'gpo 22/12/2005
    sCondicion = "(idTitular=" & Val(Me.txtTitCve.Text) & ")"
    sCondicion = sCondicion & " AND Cancelada=0"
    
    InitRecordSet rsFacts, sCampos, sTablas, sCondicion, "NumeroFactura", Conn
    
    With rsFacts
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                Me.ssdbFacturas.AddItem .Fields("NumeroFactura") & vbTab & _
                .Fields("Folio") & .Fields("Serie") & vbTab & _
                Format(.Fields("FechaFactura"), "dd / mmm / yyyy") & vbTab & _
                .Fields("NombreFactura") & vbTab & _
                .Fields("CalleFactura") & vbTab & _
                .Fields("ColoniaFactura") & vbTab & _
                .Fields("DelFactura") & vbTab & _
                .Fields("CiudadFactura") & vbTab & _
                .Fields("EstadoFactura") & vbTab & _
                .Fields("RFC") & vbTab & _
                .Fields("Observaciones")
            
                .MoveNext
            Loop
        End If
    
        .Close
    End With
    Set rsFacts = Nothing
End Sub

Private Sub MostrarDetalle()
    Dim rsDetalle As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String

    Me.ssdbDetalle.RemoveAll

    sCampos = "DETALLE.Periodo, DETALLE.idConcepto, DETALLE.Concepto, "
    sCampos = sCampos & "DETALLE.Cantidad, DETALLE.Importe, DETALLE.Intereses, "
    sCampos = sCampos & "DETALLE.Descuento, DETALLE.Iva, DETALLE.Total, "
    sCampos = sCampos & "DETALLE.idMember, DETALLE.idTipoUsuario, Tipo_Usuario.Descripcion "
    
    sTablas = "FACTURAS_DETALLE AS DETALLE LEFT JOIN Tipo_Usuario ON DETALLE.idTipoUsuario=Tipo_Usuario.idTipoUsuario"
    
    InitRecordSet rsDetalle, sCampos, sTablas, "NumeroFactura=" & Me.ssdbFacturas.Columns(0).Value, "", Conn
    With rsDetalle
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                Me.ssdbDetalle.AddItem Format(.Fields("Periodo"), "dd / mmm / yyyy") & vbTab & _
                .Fields("idConcepto") & vbTab & _
                .Fields("Concepto") & vbTab & _
                .Fields("Cantidad") & vbTab & _
                (.Fields("Total") - .Fields("Intereses") - .Fields("Iva") + .Fields("Descuento")) & vbTab & _
                .Fields("Intereses") & vbTab & _
                .Fields("Descuento") & vbTab & _
                .Fields("Iva") & vbTab & _
                .Fields("Total") & vbTab & _
                .Fields("idMember") & vbTab & _
                .Fields("idTipoUsuario") & vbTab & _
                .Fields("Descripcion")

                .MoveNext
            Loop
        End If
    
        .Close
    End With
    Set rsDetalle = Nothing
End Sub

Private Function MostrarTotal()
    Dim rsTotFact As ADODB.Recordset

    InitRecordSet rsTotFact, "SUM(Total) AS TOTFACT", "FACTURAS_DETALLE", "NumeroFactura=" & Me.ssdbFacturas.Columns(0).Value, "", Conn
    If (Not IsNull(rsTotFact.Fields(0))) Then
        Me.txtTotalFact.Text = Format(rsTotFact.Fields(0), "###,###,##0.00")
    Else
        Me.txtTotalFact.Text = "0.00"
    End If
    
    rsTotFact.Close
    Set rsTotFact = Nothing
End Function

Private Sub InitColsFact()
    '***    Datos de la factura     ***

    'Asigna valores a la matriz de encabezados
    mEncFacts(0) = "# Fact."
    mEncFacts(1) = "Folio"
    mEncFacts(2) = "Fecha"
    mEncFacts(3) = "Nombre o razón social"
    mEncFacts(4) = "Calle"
    mEncFacts(5) = "Colonia"
    mEncFacts(6) = "Delegación o municipio"
    mEncFacts(7) = "Ciudad"
    mEncFacts(8) = "Estado"
    mEncFacts(9) = "RFC"
    mEncFacts(10) = "Observaciones"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid Me.ssdbFacturas, mEncFacts
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncFacts(0) = 900
    mAncFacts(1) = 900
    mAncFacts(2) = 1500
    mAncFacts(3) = 5500
    mAncFacts(4) = 5500
    mAncFacts(5) = 4500
    mAncFacts(6) = 4500
    mAncFacts(7) = 3500
    mAncFacts(8) = 2200
    mAncFacts(9) = 1800
    mAncFacts(10) = 3000
    
    'Asigna el ancho de cada columna
    DefAnchossGrid Me.ssdbFacturas, mAncFacts
    
    Me.ssdbFacturas.Columns(0).Alignment = ssCaptionAlignmentRight
    Me.ssdbFacturas.Columns(1).Alignment = ssCaptionAlignmentCenter
    
    Me.ssdbFacturas.Columns(9).Visible = False

    '***    Detalles    ***

    'Asigna valores a la matriz de encabezados
    mEncDet(0) = "Período"
    mEncDet(1) = "# Conc."
    mEncDet(2) = "Concepto"
    mEncDet(3) = "Cantidad"
    mEncDet(4) = "Importe"
    mEncDet(5) = "Intereses"
    mEncDet(6) = "Descuento"
    mEncDet(7) = "IVA"
    mEncDet(8) = "Total"
    mEncDet(9) = "# Usuario"
    mEncDet(10) = "# Tipo"
    mEncDet(11) = "Tipo usuario"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid Me.ssdbDetalle, mEncDet
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncDet(0) = 1500
    mAncDet(1) = 900
    mAncDet(2) = 4500
    mAncDet(3) = 900
    mAncDet(4) = 1100
    mAncDet(5) = 1100
    mAncDet(6) = 1100
    mAncDet(7) = 1100
    mAncDet(8) = 1800
    mAncDet(9) = 1000
    mAncDet(10) = 900
    mAncDet(11) = 3500
    
    'Asigna el ancho de cada columna
    DefAnchossGrid Me.ssdbDetalle, mAncDet
    
    Me.ssdbDetalle.Columns(0).Alignment = ssCaptionAlignmentCenter
    Me.ssdbDetalle.Columns(1).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(3).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(4).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(5).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(6).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(7).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(8).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(9).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(10).Alignment = ssCaptionAlignmentRight
End Sub

'************************************************************
'*                          Ayudas                          *
'************************************************************

Private Sub cmdAyuda_Click()
    Const DATOSACCION = 9
    Dim sCadena As String
    Dim mFAyuda(DATOSACCION) As String
    Dim mAAyuda(DATOSACCION) As Integer
    Dim mCAyuda(DATOSACCION) As String
    Dim mEAyuda(DATOSACCION) As String

    nAyuda = 1

    Set frmHAcciones = New frmayuda
    
    mFAyuda(0) = "Acciones ordenadas por serie"
    mFAyuda(1) = "Acciones ordenadas por tipo"
    mFAyuda(2) = "Acciones ordenadas por número"
    mFAyuda(3) = "Acciones ordenadas por apellido paterno"
    mFAyuda(4) = "Acciones ordenadas por apellido materno"
    mFAyuda(5) = "Acciones ordenadas por nombre"
    mFAyuda(6) = "Acciones ordenadas por Cve del accionista"
    mFAyuda(7) = "Acciones ordenadas por Tel. 1"
    mFAyuda(8) = "Acciones ordenadas por Tel. 2"
    
    mAAyuda(0) = 800
    mAAyuda(1) = 800
    mAAyuda(2) = 800
    mAAyuda(3) = 2500
    mAAyuda(4) = 2500
    mAAyuda(5) = 2500
    mAAyuda(6) = 1100
    mAAyuda(7) = 1100
    mAAyuda(8) = 1100
    
    mCAyuda(0) = "Serie"
    mCAyuda(1) = "Tipo"
    mCAyuda(2) = "Numero"
    mCAyuda(3) = "A_Paterno"
    mCAyuda(4) = "A_Materno"
    mCAyuda(5) = "Nombre"
    mCAyuda(6) = "IdPropTitulo"
    mCAyuda(7) = "Telefono_1"
    mCAyuda(8) = "Telefono_2"
    
    mEAyuda(0) = "Serie"
    mEAyuda(1) = "Tipo"
    mEAyuda(2) = "Número"
    mEAyuda(3) = "A. Paterno"
    mEAyuda(4) = "A. Materno"
    mEAyuda(5) = "Nombre(s)"
    mEAyuda(6) = "# Accionista"
    mEAyuda(7) = "Tel. 1"
    mEAyuda(8) = "Tel. 2"
    
    With frmHAcciones
        .nColActiva = 0
        .nColsAyuda = DATOSACCION
        .sTabla = "Titulos AS T LEFT JOIN Accionistas ON T.IdPropietario=Accionistas.IdPropTitulo"
        
        .sCondicion = "IdPropietario>0 AND "
        .sCondicion = .sCondicion & " NOT EXISTS( "
        .sCondicion = .sCondicion & "SELECT * FROM Usuarios_Titulo AS UT "
        .sCondicion = .sCondicion & "WHERE T.Serie=UT.Serie AND T.Tipo=UT.Tipo AND T.Numero=UT.Numero )"
        .sTitAyuda = "Acciones aún no asignadas"
        .lAgregar = False
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
    If (Val(Me.txtCveAccionista.Text) <> nAccCve) Then
        nAccCve = Val(Me.txtCveAccionista.Text)
        BuscaDatosAcc
        InitVarAcc
        
        If (Me.optRentista.Value) Then
            Me.txtTitPaterno.Text = ""
            Me.txtTitMaterno.Text = ""
            Me.txtTitNombre.Text = ""
            
            sTitPaterno = ""
            sTitMaterno = ""
            sTitNombre = ""
            
            Me.txtTitPaterno.SetFocus
        ElseIf (Me.optPropietario.Value) Then
            Me.txtTitPaterno.Text = sTitPaterno
            Me.txtTitMaterno.Text = sTitMaterno
            Me.txtTitNombre.Text = sTitNombre
            
            Me.txtTitPaterno.Enabled = False
            Me.txtTitMaterno.Enabled = False
            Me.txtTitNombre.Enabled = False
            
            If (Me.txtTitCve.Enabled) Then
                Me.txtTitCve.SetFocus
            End If
        End If
    Else
        Me.cmdAyuda.SetFocus
    End If
End Sub

Private Sub cmdHPais_Click()
    Const DATOSPAIS = 2
    Dim sCadena As String
    Dim mFAyuda(DATOSPAIS) As String
    Dim mAAyuda(DATOSPAIS) As Integer
    Dim mCAyuda(DATOSPAIS) As String
    Dim mEAyuda(DATOSPAIS) As String

    nAyuda = 2

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
    
    If (Trim(Me.txtCvePais.Text) <> "") Then
        Me.txtPaisTit.Text = LeeXValor("Pais", "Paises", "IdPais=" & Val(Me.txtCvePais.Text), "Pais", "s", Conn)
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

    nAyuda = 3
    
    If (Trim(Me.txtTitEdad.Text) = vbNullString) Then
        MsgBox "Se debe especificar la fecha de nacimiento para" & Chr(13) & "mostrar los tipos de usuario correspondiente.", vbInformation, "KalaSystems"
        Me.dtpTitNacio.SetFocus
        Exit Sub
    End If
    
    Set frmHTipo = New frmayuda
    
    mFAyuda(0) = "Tipos de titular ordenados por clave"
    mFAyuda(1) = "Tipos de titular ordenados por descripción"
    mFAyuda(2) = "Tipos de titular ordenados por edad mínima"
    mFAyuda(3) = "Tipos de titular ordenados por edad máxima"
    
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
        
        .sCondicion = "Parentesco='TI'"
        .sCondicion = .sCondicion & " And Familiar = 0"
        .sCondicion = .sCondicion & " And (" & Int(Val(Me.txtTitEdad.Text)) & " BETWEEN EdadMinima AND EdadMaxima )"
        #If SqlServer_ Then
            .sCondicion = .sCondicion & " And Vigencia >= '" & Format(Date, "yyyymmdd") & "'"
        #Else
            .sCondicion = .sCondicion & " And Vigencia >= #" & Format(Date, "mm/dd/yyyy") & "#"
        #End If
        
        
        If (Me.optPropietario.Value) Then
            .sCondicion = .sCondicion & " AND Tipo='PROPIETARIO'"
        ElseIf (Me.optRentista.Value) Then
            .sCondicion = .sCondicion & " AND Tipo='RENTISTA'"
        ElseIf (Me.optMembresia.Value) Then
            .sCondicion = .sCondicion & " AND Tipo='MEMBRESIA'"
        End If
        
        .sTitAyuda = "Tipos de usuario titular"
        .lAgregar = True
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
    If (Trim(Me.txtCveTipo.Text) <> "") Then
        Me.txtTipoTit.Text = LeeXValor("Descripcion", "Tipo_Usuario", "IdTipoUsuario=" & Val(Me.txtCveTipo.Text), "Descripcion", "s", Conn)
    End If
    
    Me.cmdHTipo.SetFocus
End Sub

Private Sub Llena_List_Usuarios()
    Dim adorcslstUsr As ADODB.Recordset
    
    Me.lstUsuarios.Clear
    
    #If SqlServer_ Then
        Set adorcslstUsr = New ADODB.Recordset
        strSQL = "SELECT ( RTRIM(NOMBRE) + ' ' + RTRIM(A_PATERNO) + ' ' + RTRIM(A_MATERNO) ) As Nombre, IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB"
        strSQL = strSQL & " WHERE IdTitular=" & Me.txtTitCve
        strSQL = strSQL & " ORDER BY NumeroFamiliar"
    #Else
        Set adorcslstUsr = New ADODB.Recordset
        strSQL = "SELECT ( TRIM(NOMBRE) & ' ' & TRIM(A_PATERNO) & ' ' & TRIM(A_MATERNO) ) As Nombre, IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB"
        strSQL = strSQL & " WHERE IdTitular=" & Me.txtTitCve
        strSQL = strSQL & " ORDER BY NumeroFamiliar"
    #End If
    
    Set adorcslstUsr = New ADODB.Recordset
    adorcslstUsr.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do Until adorcslstUsr.EOF
        Me.lstUsuarios.AddItem adorcslstUsr!Nombre
        Me.lstUsuarios.ItemData(Me.lstUsuarios.NewIndex) = adorcslstUsr!Idmember
        adorcslstUsr.MoveNext
    Loop
    
    adorcslstUsr.Close
    Set adorcslstUsr = Nothing
    
End Sub

Private Sub LlenaDatosEmergencia()
    Dim adorcs As ADODB.Recordset
    
    On Error GoTo Error_Catch
    
    strSQL = "SELECT E.NombreEmergencia, E.ParentescoEmergencia, E.TelefonosEmergencia, E.DomicilioEmergencia, E.NombreBeneficiario1, E.ParentescoBeneficiario1, E.PorcentajeBeneficiario1, E.NombreBeneficiario2, E.ParentescoBeneficiario2, E.PorcentajeBeneficiario2"
    strSQL = strSQL & " FROM EMERGENCIA_DATOS E"
    strSQL = strSQL & " WHERE (((E.IdMember)=" & Me.txtTitCve & "))"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        
        Me.txtCtrlEmer(0).Text = Trim(adorcs!NombreEmergencia)
        Me.txtCtrlEmer(1).Text = Trim(adorcs!ParentescoEmergencia)
        Me.txtCtrlEmer(2).Text = Trim(adorcs!TelefonosEmergencia)
        Me.txtCtrlEmer(3).Text = Trim(adorcs!DomicilioEmergencia)
        Me.txtCtrlEmer(4).Text = Trim(adorcs!NombreBeneficiario1)
        Me.txtCtrlEmer(5).Text = Trim(adorcs!NombreBeneficiario2)
        Me.txtCtrlEmer(6).Text = Trim(adorcs!ParentescoBeneficiario1)
        Me.txtCtrlEmer(7).Text = Trim(adorcs!ParentescoBeneficiario2)
        Me.txtCtrlEmer(8).Text = adorcs!PorcentajeBeneficiario1 * 100
        Me.txtCtrlEmer(9).Text = adorcs!PorcentajeBeneficiario2 * 100
        
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
    On Error GoTo 0
    
    Exit Sub
    
Error_Catch:

    MsgError
    
End Sub
