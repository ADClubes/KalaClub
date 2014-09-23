VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSelecReportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelecReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   9915
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar a Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1800
      Width           =   1425
   End
   Begin VB.Frame frmDatos 
      Caption         =   "Datos Complementarios..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   300
      TabIndex        =   22
      Top             =   2100
      Width           =   8000
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   1110
         TabIndex        =   38
         Text            =   "cboDatos"
         Top             =   1485
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   3420
         TabIndex        =   15
         Text            =   "cboDatos"
         Top             =   315
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1860
         TabIndex        =   14
         Text            =   "cboDatos"
         Top             =   1545
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   4320
         TabIndex        =   18
         Text            =   "cboDatos"
         Top             =   225
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         ItemData        =   "frmSelecReportes.frx":030A
         Left            =   1230
         List            =   "frmSelecReportes.frx":031A
         TabIndex        =   17
         Text            =   "cboDatos"
         Top             =   750
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   4155
         TabIndex        =   16
         Text            =   "cboDatos"
         Top             =   810
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         ItemData        =   "frmSelecReportes.frx":0346
         Left            =   885
         List            =   "frmSelecReportes.frx":036E
         TabIndex        =   13
         Text            =   "cboDatos"
         Top             =   1290
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   435
         TabIndex        =   12
         Text            =   "cboDatos"
         Top             =   1920
         Visible         =   0   'False
         Width           =   5370
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Text            =   "cboDatos"
         Top             =   390
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.ComboBox cboDatos 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2385
         TabIndex        =   10
         Text            =   "cboDatos"
         Top             =   1725
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   195
         TabIndex        =   9
         Top             =   975
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2070
         TabIndex        =   8
         Top             =   300
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3285
         TabIndex        =   7
         Top             =   1545
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Index           =   2
         Left            =   4455
         TabIndex        =   6
         Top             =   1500
         Visible         =   0   'False
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   106102784
         CurrentDate     =   37984
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   5
         Top             =   375
         Visible         =   0   'False
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   106102784
         CurrentDate     =   37984
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Index           =   0
         Left            =   4020
         TabIndex        =   4
         Top             =   1185
         Visible         =   0   'False
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   106102784
         CurrentDate     =   37984
      End
      Begin VB.Label lblNumero 
         Alignment       =   1  'Right Justify
         Caption         =   "Caja:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblNumero 
         Alignment       =   1  'Right Justify
         Caption         =   "Turno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "$Socio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   510
         TabIndex        =   39
         Top             =   1620
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "&Serie:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   975
         TabIndex        =   37
         Top             =   1095
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "&Mes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   5880
         TabIndex        =   36
         Top             =   1245
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "&Membresía:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   5955
         TabIndex        =   35
         Top             =   1125
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "&Año:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3705
         TabIndex        =   34
         Top             =   825
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "&Accionista:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   3735
         TabIndex        =   33
         Top             =   825
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblNumero 
         Alignment       =   1  'Right Justify
         Caption         =   "Número &Final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   -30
         TabIndex        =   32
         Top             =   1695
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblNumero 
         Alignment       =   1  'Right Justify
         Caption         =   "Número &Inicial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   5805
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblNumero 
         Alignment       =   1  'Right Justify
         Caption         =   "&Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   2160
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "&Sexo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   5925
         TabIndex        =   29
         Top             =   645
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de &Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   -990
         TabIndex        =   28
         Top             =   1410
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo &Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   3675
         TabIndex        =   27
         Top             =   135
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo &Rentable:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   3465
         TabIndex        =   26
         Top             =   810
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha &Final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   5325
         TabIndex        =   25
         Top             =   210
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha &Inicial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4965
         TabIndex        =   24
         Top             =   105
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         Caption         =   "&Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   3270
         TabIndex        =   23
         Top             =   780
         Visible         =   0   'False
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccione el Reporte a Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   300
      TabIndex        =   21
      Top             =   615
      Width           =   8000
      Begin VB.ComboBox cboReportes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   3
         Top             =   855
         Width           =   6405
      End
      Begin VB.ComboBox cboGrupo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   350
         Width           =   6400
      End
      Begin VB.Label lblReporte 
         Caption         =   "&Reporte:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   500
         TabIndex        =   2
         Top             =   1000
         Width           =   840
      End
      Begin VB.Label lblGrupo 
         Caption         =   "&Grupo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   500
         TabIndex        =   0
         Top             =   500
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   1425
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Vista Preliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   180
      Width           =   1425
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSelecReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' * * * * * * * * SELECCION DE REPORTES * * * * * * * * * * * * * * * * *
' Objetivo: PRESENTA LOS REPORTES DISPONIBLES PARA SU IMPRESIÓN
' Autor:
' Fecha: DICIEMBRE de 2002
' Modificado (Totalmente): DICIEMBRE de 2003
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
    Dim intVisibles As Integer
    Dim ablnFecha(0 To 2) As Boolean
    Dim ablnNumero(0 To 4) As Boolean
    Dim ablnDatos(0 To 8) As Boolean
    Dim strTitulo, strTag, strNombre As String
    Dim frmRepSeleccionado As frmReportesCrystal
    Dim AdoRcsReportes As ADODB.Recordset
    Dim strSQL As String

Private Sub cboDatos_GotFocus(Index As Integer)
    cboDatos(Index).SelStart = 0
    cboDatos(Index).SelLength = Len(cboDatos(Index))
End Sub

Private Sub cboDatos_LostFocus(Index As Integer)
    If (cboDatos(Index).Text <> "") And (cboDatos(Index).ListIndex < 0) Then
        MsgBox "Seleccione el DATO faltante."
        cboDatos(Index).SetFocus
    End If
End Sub

Private Sub cboGrupo_Click()
    Dim strCampo1 As String, strCampo2 As String
    Dim intGrupo As Integer
    Call Limpia
    
    'Llena el combo de la lista de reportes disponibles
    If cboGrupo.ListIndex = -1 Then Exit Sub
    intGrupo = cboGrupo.ItemData(cboGrupo.ListIndex)
    strSQL = "SELECT  descripcion, idreporte FROM reportes WHERE grupo = " & _
                    intGrupo & " ORDER BY nombre"
    strCampo1 = "descripcion"
    strCampo2 = "idreporte"
    Call LlenaCombos(cboReportes, strSQL, strCampo1, strCampo2)
    If cboReportes.ListIndex <> -1 Then
        cboReportes.ListIndex = 0
    End If
End Sub

Private Sub cboReportes_Click()
    Dim i, intClaveReporte As Integer
    
    Call Limpia
    If cboReportes.ListIndex < 0 Then
        Exit Sub
    End If
    'Determina los controles que aparecerán
    intClaveReporte = cboReportes.ItemData(cboReportes.ListIndex)
    strSQL = "SELECT * FROM reportes WHERE idreporte = " & intClaveReporte
    Set AdoRcsReportes = New ADODB.Recordset
    AdoRcsReportes.ActiveConnection = Conn
    AdoRcsReportes.CursorLocation = adUseClient
    AdoRcsReportes.CursorType = adOpenDynamic
    AdoRcsReportes.LockType = adLockReadOnly
    AdoRcsReportes.Open strSQL
    If Not AdoRcsReportes.EOF Then
        ablnFecha(0) = AdoRcsReportes.Fields("fecha")
        ablnFecha(1) = AdoRcsReportes.Fields("fecha_inicial")
        ablnFecha(2) = AdoRcsReportes.Fields("fecha_final")
        ablnNumero(0) = AdoRcsReportes.Fields("num")
        ablnNumero(1) = AdoRcsReportes.Fields("num_inicial")
        ablnNumero(2) = AdoRcsReportes.Fields("num_final")
        ablnNumero(3) = AdoRcsReportes.Fields("Turno")
        ablnNumero(4) = AdoRcsReportes.Fields("NumeroCaja")
        ablnDatos(0) = AdoRcsReportes.Fields("accionista")
        ablnDatos(1) = AdoRcsReportes.Fields("anio")
        ablnDatos(2) = AdoRcsReportes.Fields("membresia")
        ablnDatos(3) = AdoRcsReportes.Fields("mes")
        ablnDatos(4) = AdoRcsReportes.Fields("serie")
        ablnDatos(5) = AdoRcsReportes.Fields("sexo")
        ablnDatos(6) = AdoRcsReportes.Fields("tipo_pago")
        ablnDatos(7) = AdoRcsReportes.Fields("tipo_rentable")
        ablnDatos(8) = AdoRcsReportes.Fields("tipo_usuario")
        strNombre = AdoRcsReportes.Fields("nombre")
        strTitulo = AdoRcsReportes.Fields("titulo")
        strTag = AdoRcsReportes.Fields("tag")
    End If
    For i = 0 To 2      'Presenta los Date Pickers
        dtpFecha(i).Visible = ablnFecha(i)
        lblFecha(i).Visible = ablnFecha(i)
        If dtpFecha(i).Visible = True Then
            intVisibles = intVisibles + 1
            Call AcomodaControl(lblFecha(i), dtpFecha(i), intVisibles)
            If intVisibles > 3 Then Exit Sub
        End If
    Next i
    For i = 0 To 4      'Presenta los Cuadros de Texto
        txtNumero(i).Visible = ablnNumero(i)
        lblNumero(i).Visible = ablnNumero(i)
        If txtNumero(i).Visible = True Then
            intVisibles = intVisibles + 1
            Call AcomodaControl(lblNumero(i), txtNumero(i), intVisibles)
            If intVisibles > 3 Then Exit Sub
        End If
    Next i
    For i = 0 To 8      'Presenta los Combos
        cboDatos(i).Visible = ablnDatos(i)
        lblDatos(i).Visible = ablnDatos(i)
        If cboDatos(i).Visible = True Then
            intVisibles = intVisibles + 1
            Call AcomodaControl(lblDatos(i), cboDatos(i), intVisibles)
            If intVisibles > 3 Then Exit Sub
        End If
    Next i
End Sub

Private Sub cboReportes_LostFocus()
    If (cboReportes.Text <> "") And (cboReportes.ListIndex < 0) Then
        MsgBox "Seleccione un REPORTE de la lista"
        cboReportes.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdExportar_Click()
On Error GoTo ERROR_EXPORTAR

    Dim ExcelApp As Excel.Application
    Dim ExcelWB As Excel.Workbook
    Dim ExcelWS As Excel.Worksheet
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim row As Integer
    Dim intClaveReporte As Integer
    Dim nombreReporte As String
    
    'Exit Sub
    
    
    If cboGrupo.Text = "" Then
        MsgBox "¡ Seleccione un Grupo !", vbInformation, " ¡ Dato Faltante !"
        cboGrupo.SetFocus
        Exit Sub
    End If
    
    If cboReportes.Text = "" Then
        MsgBox "¡ Seleccione un Reporte !", vbInformation, " ¡ Dato Faltante !"
        cboReportes.SetFocus
        Exit Sub
    End If
    
    If (dtpFecha(1).Visible = True) And (dtpFecha(2).Visible = True) And (dtpFecha(1).Value > dtpFecha(2).Value) Then
        MsgBox "¡ La Fecha Final No Puede Ser Menor Que la Fecha Inicial !", vbInformation, " ¡ Dato Erroneo !"
        dtpFecha(2).SetFocus
        Exit Sub
    End If
    
    For i = 0 To 4
        If (txtNumero(i).Visible = True) And (txtNumero(i).Text = "") Then
            MsgBox "¡ Introduzca en Número Faltante !", vbInformation, " ¡ Dato Faltante !"
            txtNumero(i).SetFocus
            Exit Sub
        End If
    Next i

    If (txtNumero(1).Visible = True) And (txtNumero(2).Visible = True) And (Val(txtNumero(1).Text) > Val(txtNumero(2).Text)) Then
        MsgBox "¡ El Límite Superior No Puede Ser Menor Que el Límite Inferior !", vbInformation, " ¡ Dato Erroneo !"
        txtNumero(2).SetFocus
        Exit Sub
    End If
    
    For i = 0 To 8
        If (cboDatos(i).Visible = True) And (cboDatos(i).Text = "") Then
            MsgBox "¡ Seleccione el Dato Faltante !", vbInformation, " ¡ Dato Faltante !"
            cboDatos(i).SetFocus
            Exit Sub
        End If
    Next i
    
    nombreReporte = cboGrupo.Text & "_" & cboReportes.Text & ".xlsx"
    
    Screen.MousePointer = vbHourglass
    
    Me.CommonDialog1.DialogTitle = "Nombre de archivo de salida"
    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.InitDir = "%USERPROFILE%\Mis documentos"
    Me.CommonDialog1.Filter = "Archivos de Microsoft Excel (*.xlsx)|*.xlsx|"
    Me.CommonDialog1.FilterIndex = 1
    Me.CommonDialog1.FileName = Trim(nombreReporte)
    Me.CommonDialog1.ShowSave
    
    If Len(Trim(Me.CommonDialog1.FileName)) > 0 Then
        nombreReporte = Trim(Me.CommonDialog1.FileName)
    End If
    
    Call Obtiene_Query
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = Conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenDynamic
    rs.LockType = adLockReadOnly
    rs.Open strSQL
    
    If rs.EOF Then
        MsgBox "No se encontró información para este reporte."
        Set rs = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelWB = ExcelApp.Workbooks.Add()
    Set ExcelWS = ExcelWB.Worksheets(1)
    
    For i = 1 To rs.Fields.Count
        ExcelWS.Cells(1, i) = rs.Fields(i - 1).Name
        DoEvents
    Next i
    
    row = 2
    If Not rs.EOF Then rs.MoveFirst
    While Not rs.EOF
        For i = 1 To rs.Fields.Count
            ExcelWS.Cells(row, i) = rs.Fields(i - 1).Value
            DoEvents
        Next i
        
        rs.MoveNext
        row = row + 1
    Wend
    
    ExcelWB.Close savechanges:=True, FileName:=nombreReporte
    rs.Close
    Set rs = Nothing
    Screen.MousePointer = vbDefault
    
    MsgBox "El reporte se exportó correctamente."
    Exit Sub
    
ERROR_EXPORTAR:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub cmdImprimir_Click()
    Dim i, intClaveReporte As Integer
    
    If cboGrupo.Text = "" Then
        MsgBox "¡ Seleccione un Grupo !", vbInformation, " ¡ Dato Faltante !"
        cboGrupo.SetFocus
        Exit Sub
    End If
    
    If cboReportes.Text = "" Then
        MsgBox "¡ Seleccione un Reporte !", vbInformation, " ¡ Dato Faltante !"
        cboReportes.SetFocus
        Exit Sub
    End If
    
        

    If (dtpFecha(1).Visible = True) And (dtpFecha(2).Visible = True) And (dtpFecha(1).Value > dtpFecha(2).Value) Then
        MsgBox "¡ La Fecha Final No Puede Ser Menor Que la Fecha Inicial !", vbInformation, " ¡ Dato Erroneo !"
        dtpFecha(2).SetFocus
        Exit Sub
    End If
    
    For i = 0 To 4
        If (txtNumero(i).Visible = True) And (txtNumero(i).Text = "") Then
            MsgBox "¡ Introduzca en Número Faltante !", vbInformation, " ¡ Dato Faltante !"
            txtNumero(i).SetFocus
            Exit Sub
        End If
    Next i

    If (txtNumero(1).Visible = True) And (txtNumero(2).Visible = True) And (Val(txtNumero(1).Text) > Val(txtNumero(2).Text)) Then
        MsgBox "¡ El Límite Superior No Puede Ser Menor Que el Límite Inferior !", vbInformation, " ¡ Dato Erroneo !"
        txtNumero(2).SetFocus
        Exit Sub
    End If
    
    For i = 0 To 8
        If (cboDatos(i).Visible = True) And (cboDatos(i).Text = "") Then
            MsgBox "¡ Seleccione el Dato Faltante !", vbInformation, " ¡ Dato Faltante !"
            cboDatos(i).SetFocus
            Exit Sub
        End If
    Next i
    
' Valida valores y prepara reporte
    sReportes.Nombre = Trim(strNombre)
    sReportes.Titulo = Trim(strTitulo)

'    sReportes.xImpPara(0) = Fecha
'    sReportes.xImpPara(1) = Mes
'    sReportes.xImpPara(2) = Año
'    sReportes.xImpPara(3) = Cliente
'    sReportes.xImpPara(4) = Rep
'
'    If DTPFinal.Visible = True Then
'        sReportes.xImpPara(0) = "#" & Format(Me.DTPFinal.Value, "mm/dd/yyyy") & "#"
'    End If
'    If CmbAnio.Visible = True Then
'        sReportes.xImpPara(0) = Val(Me.CmbAnio.Text)
'    End If
'    If cmbMes.Visible = True Then
'        sReportes.xImpPara(0) = cmbMes.Text
'    End If
'    If cmbCliente.Visible = True Then
'        sReportes.xImpPara(3) = Me.cmbCliente.ItemData(Me.cmbCliente.ListIndex)
'    End If
'    If Me.cmbRep.Visible Then
'        sReportes.xImpPara(1) = Month(Me.DTPFinal.Value)
'        sReportes.xImpPara(2) = Year(Me.DTPFinal.Value)
'        sReportes.xImpPara(4) = Me.cmbRep.ItemData(Me.cmbRep.ListIndex)
'    End If

    Set frmRepSeleccionado = New frmReportesCrystal
    frmRepSeleccionado.Tag = Trim(strTag)
    Load frmRepSeleccionado
    frmRepSeleccionado.Show
End Sub

Private Sub Form_Activate()
    With frmSelecReportes
        .Height = 5145
        .Left = 0
        .Top = 0
        .Width = 10000
    End With
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strCampo1, strCampo2, strCampo3 As String
    
    intVisibles = 0
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Impresión de Reportes"
    
    '07/Dic/2011 UCM
'    If sDB_NivelUser = 0 Then
'        'Llena combo Grupo
'        strSQL = "SELECT * FROM GRUPO_REPORTES ORDER BY Descripcion"
'    Else
'        'Llena combo Grupo
'        strSQL = "SELECT * FROM GRUPO_REPORTES WHERE Privado = 0 ORDER BY Descripcion"
'    End If
    strSQL = "usp_Catalogo_Reportes"
    strCampo3 = sDB_NivelUser
    strCampo1 = "descripcion"
    strCampo2 = "idgporeportes"
    'Call LlenaCombos(cboGrupo, strSQL, strCampo1, strCampo2)
    Call LlenaComboSP(cboGrupo, strSQL, strCampo1, strCampo2, strCampo3)
    'Inicializa Fechas
    dtpFecha(0).Value = Now()
    dtpFecha(1).Value = Now() '"01/" & Month(Now()) & "/" & Year(Now()) 'DateAdd("m", -3, DTPFinal.Value)
    dtpFecha(2).Value = Now() 'DateAdd("m", 1, dtpFecha(1).Value)
    
    'Llena combo Accionista
    'strSQL = "SELECT idproptitulo, (a_paterno & ' ' & a_materno & ' ' & nombre) as Nombres " & _
                    "FROM accionistas ORDER BY a_paterno, a_materno, nombre"
    'strCampo1 = "Nombres"
    'strCampo2 = "idproptitulo"
    'Call LlenaCombos(cboDatos(0), strSQL, strCampo1, strCampo2)
    
    'Llena combo Año
    For i = Year(Now()) - 10 To Year(Now()) + 10
        cboDatos(1).AddItem i
    Next i
    
    'Llena combo Membresia
    'strSQL = "SELECT idmembresia, descripcion FROM membresias ORDER BY idmembresia"
    'strCampo1 = "descripcion"
    'strCampo2 = "idmembresia"
    'Call LlenaCombos(cboDatos(2), strSQL, strCampo1, strCampo2)
    
    'El cboDatos(3) está prellenado con los meses del año
    
    'Llena combo Serie
    'strSQL = "SELECT DISTINCT serie FROM titulos ORDER BY serie"
    'strCampo1 = "serie"
    'strCampo2 = ""
    'Call LlenaCombos(cboDatos(4), strSQL, strCampo1, strCampo2)
    
    'El cboDatos(5) está prellenado con Masculino y Femenino
    
    'Llena combo Tipo de Pago
    'strSQL = "SELECT idformapago, descripcion FROM forma_pago ORDER BY idformapago"
    'strCampo1 = "descripcion"
    'strCampo2 = "idformapago"
    'Call LlenaCombos(cboDatos(6), strSQL, strCampo1, strCampo2)
    
    'Llena combo Tipo de Rentable
    'strSQL = "SELECT idtiporentable, descripcion FROM tipo_rentables ORDER BY idtiporentable"
    'strCampo1 = "descripcion"
    'strCampo2 = "idtiporentable"
    'Call LlenaCombos(cboDatos(7), strSQL, strCampo1, strCampo2)
    
    'Llena combo Tipo de Usuario
    'strSQL = "SELECT idtipousuario, descripcion FROM tipo_usuario ORDER BY idtipousuario"
    'strCampo1 = "descripcion"
    'strCampo2 = "idtipousuario"
    'Call LlenaCombos(cboDatos(8), strSQL, strCampo1, strCampo2)
    
    'Llena combo Socio
    'strSQL = "SELECT idmember, (nombre + ' ' + a_paterno + ' ' + a_materno) as Nombre_Completo FROM usuarios_club ORDER BY nombre"
    'strCampo1 = "Nombre_Completo"
    'strCampo2 = "idmember"
    'Call LlenaCombos(cboDatos(9), strSQL, strCampo1, strCampo2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Pantalla Principal"
End Sub

Sub AcomodaControl(Etiqueta As Control, UnControl As Control, intPosicion As Integer)
    Etiqueta.Left = 200
    UnControl.Left = 2400
    Select Case intPosicion
        Case 1
            Etiqueta.Top = 500
            UnControl.Top = 350
        Case 2
            Etiqueta.Top = 1000
            UnControl.Top = 850
        Case 3
            Etiqueta.Top = 1500
            UnControl.Top = 1350
        Case 4
            Etiqueta.Top = 2000
            UnControl.Top = 1850
        End Select
End Sub

Sub Limpia()
    Dim i As Integer
    Dim ctrLabel As Label
    Dim ctrText As TextBox
    Dim ctrDTP As DTPicker
    Dim ctrCombo As ComboBox
    
    For i = 0 To 3
        txtNumero(i).Text = ""
    Next i
    For i = 0 To 8
        cboDatos(i).ListIndex = -1
    Next i
    For Each ctrLabel In lblFecha
        ctrLabel.Visible = False
    Next
    For Each ctrDTP In dtpFecha
        ctrDTP.Visible = False
    Next
    For Each ctrLabel In lblNumero
        ctrLabel.Visible = False
    Next
    For Each ctrText In txtNumero
        ctrText.Visible = False
    Next
    For Each ctrLabel In lblDatos
        ctrLabel.Visible = False
    Next
    For Each ctrCombo In cboDatos
        ctrCombo.Visible = False
    Next
    intVisibles = 0
End Sub

Private Sub txtNumero_GotFocus(Index As Integer)
    txtNumero(Index).SelStart = 0
    txtNumero(Index).SelLength = Len(txtNumero(Index))
End Sub

Private Sub txtNumero_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 22, 46, 48 To 57    'Backspace, <Ctrl+V>, punto y del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

'''11.Ago.2011
Sub Obtiene_Query()
    Dim AdoRcsReportes As ADODB.Recordset
    Dim strParametro, strProc, strvar As String
    Dim lngPos As Long
    Dim sqlQuery As String
    
    On Error GoTo Err_Obtiene_Query
        
    sqlQuery = "SELECT * FROM reportes WHERE idreporte = " & _
                    frmSelecReportes.cboReportes.ItemData(frmSelecReportes.cboReportes.ListIndex)
    
    Set AdoRcsReportes = New ADODB.Recordset
    AdoRcsReportes.ActiveConnection = Conn
    AdoRcsReportes.CursorLocation = adUseClient
    AdoRcsReportes.CursorType = adOpenDynamic
    AdoRcsReportes.LockType = adLockReadOnly
    AdoRcsReportes.Open sqlQuery
    
    strParametro = sReportes.xImpPara(0)
       
    'Arma el query para el reporte
    strSQL = ""
    If Not AdoRcsReportes.EOF Then
        #If SqlServer_ Then
            strProc = Trim(AdoRcsReportes!sqlQuerySQL)
        #Else
            strProc = Trim(AdoRcsReportes!sqlQuery)
        #End If
        strSQL = ""
        Do While 1
            lngPos = InStr(strProc, "<@")
            If lngPos = 0 Then
                strSQL = strSQL & strProc
                Exit Do
            Else
                strSQL = strSQL & Left$(strProc, lngPos - 1)
            End If
            strProc = Mid$(strProc, lngPos)
            lngPos = InStr(strProc, ">")
            strvar = Left$(strProc, lngPos)
            strProc = Mid$(strProc, lngPos + 1)
            Select Case strvar
                Case "<@Fecha>"
                    #If SqlServer_ Then
                        strSQL = strSQL & " '" & Format(frmSelecReportes.dtpFecha(0).Value, "yyyymmdd") & "' "
                    #Else
                        strSQL = strSQL & " #" & Format(frmSelecReportes.dtpFecha(0).Value, "mm/dd/yyyy") & "# "
                    #End If
                Case "<@FechaInicial>"
                    #If SqlServer_ Then
                        strSQL = strSQL & " '" & Format(frmSelecReportes.dtpFecha(1).Value, "yyyymmdd") & "' "
                    #Else
                        strSQL = strSQL & " #" & Format(frmSelecReportes.dtpFecha(1).Value, "mm/dd/yyyy") & "# "
                    #End If
                Case "<@FechaFinal>"
                    #If SqlServer_ Then
                        strSQL = strSQL & " '" & Format(frmSelecReportes.dtpFecha(2).Value, "yyyymmdd") & "' "
                    #Else
                        strSQL = strSQL & " #" & Format(frmSelecReportes.dtpFecha(2).Value, "mm/dd/yyyy") & "# "
                    #End If
                Case "<@Num>"
                    strSQL = strSQL & Val(frmSelecReportes.txtNumero(0).Text)
                Case "<@NumInicial>"
                    strSQL = strSQL & Val(frmSelecReportes.txtNumero(1).Text)
                Case "<@NumFinal>"
                    strSQL = strSQL & Val(frmSelecReportes.txtNumero(2).Text)
                Case "<@IdAccionista>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(0).ItemData(frmSelecReportes.cboDatos(0).ListIndex)
                Case "<@Anio>"
                    strSQL = strSQL & Val(frmSelecReportes.cboDatos(1).Text)
                Case "<@IdMembresia>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(2).ItemData(frmSelecReportes.cboDatos(2).ListIndex)
                Case "<@Mes>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(3).ListIndex + 1
                Case "<@Serie>"
                    strSQL = strSQL & "'" & Trim(frmSelecReportes.cboDatos(4).Text) & "'"
                Case "<@Sexo>"
                    strSQL = strSQL & "'" & Trim(frmSelecReportes.cboDatos(5).Text) & "'"
                Case "<@TipoPago>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(6).ItemData(frmSelecReportes.cboDatos(6).ListIndex)
                Case "<@TipoRentable>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(7).ItemData(frmSelecReportes.cboDatos(7).ListIndex)
                Case "<@TipoUsuario>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(8).ItemData(frmSelecReportes.cboDatos(8).ListIndex)
                Case "<@Socio>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(9).ItemData(frmSelecReportes.cboDatos(9).ListIndex)
                Case "<@Turno>"
                    If Val(frmSelecReportes.txtNumero(3).Text) > 0 Then
                        strSQL = strSQL & Trim(frmSelecReportes.txtNumero(3).Text)
                    Else
                        #If SqlServer_ Then
                            strSQL = Replace(strSQL, " AND FACTURAS.TURNO=", "")
                        #Else
                            strSQL = strSQL & "True"
                        #End If
                    End If
                Case "<@Caja>"
                    If Val(frmSelecReportes.txtNumero(4).Text) > 0 Then
                        strSQL = strSQL & Trim(frmSelecReportes.txtNumero(4).Text)
                    Else
                        strSQL = strSQL & "True"
                    End If
            End Select
        Loop
    End If
    
    Exit Sub
Err_Obtiene_Query:
        GeneraMensajeError Err.Number
End Sub

