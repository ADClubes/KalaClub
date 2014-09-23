VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConsultas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultas"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10500
   Icon            =   "frmConsultas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmConsultas.frx":030A
   ScaleHeight     =   6195
   ScaleWidth      =   10500
   Begin VB.Frame frmBusca 
      Height          =   6180
      Left            =   12000
      TabIndex        =   30
      Top             =   -30
      Width           =   10455
      Begin VB.Frame frmBuscaPor 
         Height          =   1020
         Left            =   510
         TabIndex        =   36
         Top             =   300
         Width           =   9390
         Begin VB.OptionButton optBusca 
            Caption         =   "C.U.R.P."
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   46
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "R.F.C."
            Height          =   240
            Index           =   1
            Left            =   210
            TabIndex        =   45
            Top             =   600
            Width           =   840
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "Num. Empleado"
            Height          =   225
            Index           =   2
            Left            =   1905
            TabIndex        =   44
            Top             =   315
            Width           =   1440
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "Nombre"
            Height          =   225
            Index           =   3
            Left            =   1905
            TabIndex        =   43
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "1er Apellido"
            Height          =   225
            Index           =   4
            Left            =   4005
            TabIndex        =   42
            Top             =   300
            Width           =   1215
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "Colonia"
            Height          =   225
            Index           =   5
            Left            =   4005
            TabIndex        =   41
            Top             =   600
            Width           =   960
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "Entidad Federativa"
            Height          =   225
            Index           =   6
            Left            =   5535
            TabIndex        =   40
            Top             =   300
            Width           =   1785
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "Departamento"
            Height          =   225
            Index           =   7
            Left            =   5535
            TabIndex        =   39
            Top             =   615
            Width           =   1500
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "Puesto"
            Height          =   225
            Index           =   8
            Left            =   7950
            TabIndex        =   38
            Top             =   300
            Width           =   960
         End
         Begin VB.OptionButton optBusca 
            Caption         =   "Turno"
            Height          =   225
            Index           =   9
            Left            =   7935
            TabIndex        =   37
            Top             =   615
            Width           =   915
         End
      End
      Begin VB.TextBox txtBusca 
         Alignment       =   2  'Center
         Height          =   330
         Left            =   510
         TabIndex        =   35
         Top             =   1410
         Width           =   4410
      End
      Begin VB.ComboBox cboBusca 
         Height          =   315
         Left            =   5250
         TabIndex        =   34
         Top             =   1410
         Visible         =   0   'False
         Width           =   4665
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   450
         Left            =   3285
         TabIndex        =   33
         Top             =   1845
         Width           =   3615
      End
      Begin VB.CommandButton cmdCancela 
         Caption         =   "&Cancelar"
         Height          =   450
         Left            =   5925
         TabIndex        =   32
         Top             =   5610
         Width           =   2715
      End
      Begin VB.CommandButton cmdAceptar 
         Cancel          =   -1  'True
         Caption         =   "&Aceptar"
         Height          =   450
         Left            =   1785
         TabIndex        =   31
         Top             =   5580
         Width           =   2715
      End
      Begin MSAdodcLib.Adodc AdoDcBusca 
         Height          =   345
         Left            =   3840
         Top             =   3810
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   609
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "AdoDcBusca"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid dtGrdBusca 
         Bindings        =   "frmConsultas.frx":0614
         Height          =   2970
         Left            =   510
         TabIndex        =   47
         Top             =   2415
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   5239
         _Version        =   393216
         ForeColor       =   8388608
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "RESULTADOS DE LA BÚSQUEDA"
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "numempleado"
            Caption         =   "Número"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "apellidopat"
            Caption         =   "1er Apellido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "apellidomat"
            Caption         =   "2do Apellido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "nombre"
            Caption         =   "Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "nomdepartamento"
            Caption         =   "Departamento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "cveempleado"
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1860.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   2355.024
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2085.166
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Busqueda de Emplados por:"
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
         Left            =   3810
         TabIndex        =   48
         Top             =   15
         Width           =   3075
      End
   End
   Begin VB.Frame frmLaborales 
      Caption         =   "Datos Laborales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   1500
      TabIndex        =   7
      Top             =   4980
      Width           =   8730
      Begin VB.Label lblTrabajo 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   6840
         TabIndex        =   27
         Top             =   315
         Width           =   1725
      End
      Begin VB.Label lblTrabajo 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1605
         TabIndex        =   26
         Top             =   705
         Width           =   6150
      End
      Begin VB.Label lblTrabajo 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   1590
         TabIndex        =   25
         Top             =   330
         Width           =   4185
      End
      Begin VB.Label lblLaboral 
         Alignment       =   1  'Right Justify
         Caption         =   "Turno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   6060
         TabIndex        =   17
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblLaboral 
         Alignment       =   1  'Right Justify
         Caption         =   "Puesto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   675
         TabIndex        =   16
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblLaboral 
         Alignment       =   1  'Right Justify
         Caption         =   "Departamento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   345
         Width           =   1350
      End
   End
   Begin VB.Frame frmLogo 
      Height          =   5520
      Left            =   120
      TabIndex        =   3
      Top             =   570
      Width           =   1245
      Begin VB.Image imgKala 
         Height          =   5115
         Left            =   135
         Picture         =   "frmConsultas.frx":062D
         Stretch         =   -1  'True
         Top             =   225
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   825
      Width           =   990
   End
   Begin VB.Frame frmGenerales 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   1515
      TabIndex        =   0
      Top             =   1665
      Width           =   8685
      Begin VB.Label lblGral 
         Alignment       =   1  'Right Justify
         Caption         =   "C.U.R.P.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   855
         TabIndex        =   29
         Top             =   2435
         Width           =   855
      End
      Begin VB.Label lblPersonales 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   7
         Left            =   1900
         TabIndex        =   28
         Top             =   2865
         Width           =   4140
      End
      Begin VB.Label lblPersonales 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   1900
         TabIndex        =   24
         Top             =   2435
         Width           =   5160
      End
      Begin VB.Label lblPersonales 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   1900
         TabIndex        =   23
         Top             =   2005
         Width           =   1950
      End
      Begin VB.Label lblPersonales 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   1900
         TabIndex        =   22
         Top             =   1575
         Width           =   5280
      End
      Begin VB.Label lblPersonales 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   6615
         TabIndex        =   21
         Top             =   2010
         Width           =   1035
      End
      Begin VB.Label lblPersonales 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   20
         Top             =   1155
         Width           =   3555
      End
      Begin VB.Label lblPersonales 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   1900
         TabIndex        =   19
         Top             =   715
         Width           =   5235
      End
      Begin VB.Label lblPersonales 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   1900
         TabIndex        =   18
         Top             =   330
         Width           =   5265
      End
      Begin VB.Label lblGral 
         Alignment       =   1  'Right Justify
         Caption         =   "Teléfono(s):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   105
         TabIndex        =   14
         Top             =   2865
         Width           =   1605
      End
      Begin VB.Label lblGral 
         Alignment       =   1  'Right Justify
         Caption         =   "R. F. C.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   915
         TabIndex        =   13
         Top             =   2005
         Width           =   795
      End
      Begin VB.Label lblGral 
         Alignment       =   1  'Right Justify
         Caption         =   "C. P.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   5400
         TabIndex        =   12
         Top             =   2010
         Width           =   1095
      End
      Begin VB.Label lblGral 
         Alignment       =   1  'Right Justify
         Caption         =   "Ent. Federativa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   1575
         Width           =   1575
      End
      Begin VB.Label lblGral 
         Alignment       =   1  'Right Justify
         Caption         =   "Del. o Mun.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   615
         TabIndex        =   10
         Top             =   1154
         Width           =   1095
      End
      Begin VB.Label lblGral 
         Alignment       =   1  'Right Justify
         Caption         =   "Colonia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   715
         Width           =   1605
      End
      Begin VB.Label lblGral 
         Alignment       =   1  'Right Justify
         Caption         =   "Calle y Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   330
         Width           =   1605
      End
   End
   Begin VB.Label lblNombre 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Index           =   2
      Left            =   3060
      TabIndex        =   6
      Top             =   1005
      Width           =   5640
   End
   Begin VB.Label lblNombre 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   1
      Left            =   3060
      TabIndex        =   5
      Top             =   540
      Width           =   5640
   End
   Begin VB.Label lblNombre 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Index           =   0
      Left            =   3060
      TabIndex        =   4
      Top             =   90
      Width           =   5640
   End
   Begin VB.Image imgFoto 
      Height          =   1455
      Left            =   1485
      Stretch         =   -1  'True
      Top             =   105
      Width           =   1305
   End
   Begin VB.Shape Shape1 
      Height          =   1260
      Left            =   1305
      Shape           =   3  'Circle
      Top             =   270
      Width           =   1620
   End
   Begin VB.Line Line1 
      X1              =   1725
      X2              =   2490
      Y1              =   1320
      Y2              =   420
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "SIN FOTOGRAFÍA DISPONIBLE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1530
      TabIndex        =   49
      Top             =   570
      Width           =   1185
   End
End
Attribute VB_Name = "frmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytOpcion As Byte
Dim lngCveEmpleado As Long
Dim AdoRcsCons As ADODB.Recordset, AdoComCons As ADODB.Command

Private Sub cboBusca_Click()
    Presenta_datos
End Sub

Private Sub cmdAceptar_Click()
    Dim strArchivoExistente, strArchivoFoto As String
    txtBusca.Text = ""
    frmBusca.Left = 12000
    strSQL = "SELECT emp.apellidopat, emp.apellidomat, emp.nombre, " & _
                "emp.calle, emp.numero, emp.colonia, entfederativa.nomentfederativa, " & _
                "emp.cp, delomuni.nomdelomuni, emp.rfc, emp.curp, " & _
                "departamentos.nomdepartamento, puestos.nompuesto, " & _
                "turnos.nomturno, emp.telefono, emp.cveempleado FROM " & _
                "((((empleados emp LEFT JOIN departamentos ON emp.departamento " & _
                "= departamentos.cvedepartamento) LEFT JOIN puestos ON " & _
                "emp.puesto = puestos.cvepuesto) LEFT JOIN turnos ON emp.turno " & _
                "= turnos.cveturno) LEFT JOIN entfederativa ON emp.entfederativa = " & _
                "entfederativa.cveentfederativa) LEFT JOIN delomuni ON " & _
                "emp.delomuni = delomuni.cvedelomuni " & _
                "WHERE emp.cveEmpleado = " & lngCveEmpleado
        
    Set AdoRcsCons = New ADODB.Recordset
    AdoRcsCons.ActiveConnection = Conn
    AdoRcsCons.CursorLocation = adUseClient
    AdoRcsCons.CursorType = adOpenDynamic
    AdoRcsCons.LockType = adLockReadOnly
    AdoRcsCons.Open strSQL
    If Not AdoRcsCons.EOF Then
        lblNombre(0) = AdoRcsCons!apellidopat
        lblNombre(1) = AdoRcsCons!apellidomat
        lblNombre(2) = AdoRcsCons!Nombre
        lblPersonales(0) = AdoRcsCons!calle & " No. " & AdoRcsCons!numero
        lblPersonales(1) = AdoRcsCons!colonia
        lblPersonales(2) = AdoRcsCons!nomdelomuni
        lblPersonales(3) = AdoRcsCons!cp
        lblPersonales(4) = AdoRcsCons!nomentfederativa
        lblPersonales(5) = AdoRcsCons!rfc
        lblPersonales(6) = AdoRcsCons!curp
        lblPersonales(7) = AdoRcsCons!telefono
        lblTrabajo(0) = AdoRcsCons!nomdepartamento
        lblTrabajo(1) = AdoRcsCons!nompuesto
        lblTrabajo(2) = AdoRcsCons!nomturno
    End If
    strArchivoFoto = sDB_DataSource & "\saifargotof\" & lngCveEmpleado & ".jpg"
    
    strArchivoExistente = Dir$(strArchivoFoto, vbNormal)
    If strArchivoExistente <> "" Then
        imgFoto.Picture = LoadPicture(strArchivoFoto)
    Else
        imgFoto.Picture = Nothing
    End If
End Sub

Private Sub cmdBuscar_Click()
    frmBusca.Left = 0
End Sub

Private Sub cmdCancela_Click()
    txtBusca.Text = ""
    frmBusca.Left = 12000
End Sub

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub cmdMostrar_Click()
    Call Presenta_datos
End Sub

Private Sub dtGrdBusca_Click()
    If Not IsNull(AdoDcBusca.Recordset.Fields("cveempleado")) Then
        lngCveEmpleado = AdoDcBusca.Recordset.Fields("cveempleado")
    End If
End Sub

Private Sub dtGrdBusca_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not IsNull(AdoDcBusca.Recordset.Fields("cveempleado")) Then
        lngCveEmpleado = AdoDcBusca.Recordset.Fields("cveempleado")
    End If
End Sub

Private Sub Form_Activate()
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "CONSULTA DE REGISTROS DE EMPLEADOS"
   With frmConsultas
      .Height = 6645
      .Left = 0
      .Top = 0
      .Width = 10590
   End With
End Sub

Private Sub optBusca_Click(Index As Integer)
    Dim strCampo1, strCampo2 As String
    cmdMostrar.Enabled = True
    txtBusca.Text = ""
    Select Case Index
        Case 0, 1
            txtBusca.Visible = True
            cboBusca.Clear
            cboBusca.Visible = False
            txtBusca.SetFocus
            If Index = 0 Then
                bytOpcion = 0
            Else
                bytOpcion = 1
            End If
        Case 2
            bytOpcion = 2
            txtBusca.Visible = False
            strSQL = "SELECT DISTINCT numempleado, campocero FROM empleados " & _
                        "ORDER BY numempleado"
            strCampo1 = "numempleado"
            strCampo2 = "campocero"
            Call LlenaCombos(cboBusca, strSQL, strCampo1, strCampo2)
            cboBusca.ListIndex = 0
            cboBusca.Visible = True
            cboBusca.SetFocus
        Case 3
            bytOpcion = 3
            txtBusca.Visible = False
            strSQL = "SELECT DISTINCT nombre, campocero FROM empleados " & _
                        "ORDER BY nombre"
            strCampo1 = "nombre"
            strCampo2 = "campocero"
            Call LlenaCombos(cboBusca, strSQL, strCampo1, strCampo2)
            cboBusca.ListIndex = 0
            cboBusca.Visible = True
            cboBusca.SetFocus
        Case 4
            bytOpcion = 4
            txtBusca.Visible = False
            strSQL = "SELECT DISTINCT apellidopat, campocero FROM empleados " & _
                        "ORDER BY apellidopat"
            strCampo1 = "apellidopat"
            strCampo2 = "campocero"
            Call LlenaCombos(cboBusca, strSQL, strCampo1, strCampo2)
            cboBusca.ListIndex = 0
            cboBusca.Visible = True
            cboBusca.SetFocus
        Case 5
            bytOpcion = 5
            txtBusca.Visible = False
            strSQL = "SELECT DISTINCT colonia, campocero FROM empleados " & _
                        "ORDER BY colonia"
            strCampo1 = "colonia"
            strCampo2 = "campocero"
            Call LlenaCombos(cboBusca, strSQL, strCampo1, strCampo2)
            cboBusca.ListIndex = 0
            cboBusca.Visible = True
            cboBusca.SetFocus
        Case 6
            bytOpcion = 6
            txtBusca.Visible = False
            strSQL = "SELECT * FROM entfederativa ORDER BY nomentfederativa"
            strCampo1 = "nomEntFederativa"
            strCampo2 = "cveEntFederativa"
            Call LlenaCombos(cboBusca, strSQL, strCampo1, strCampo2)
            cboBusca.ListIndex = 0
            cboBusca.Visible = True
            cboBusca.SetFocus
        Case 7
            bytOpcion = 7
            txtBusca.Visible = False
            strSQL = "SELECT * FROM departamentos ORDER BY nomdepartamento"
            strCampo1 = "nomdepartamento"
            strCampo2 = "cvedepartamento"
            Call LlenaCombos(cboBusca, strSQL, strCampo1, strCampo2)
            cboBusca.ListIndex = 0
            cboBusca.Visible = True
            cboBusca.SetFocus
        Case 8
            bytOpcion = 8
            txtBusca.Visible = False
            strSQL = "SELECT * FROM puestos ORDER BY nompuesto"
            strCampo1 = "nompuesto"
            strCampo2 = "cvepuesto"
            Call LlenaCombos(cboBusca, strSQL, strCampo1, strCampo2)
            cboBusca.ListIndex = 0
            cboBusca.Visible = True
            cboBusca.SetFocus
        Case 9
            bytOpcion = 9
            txtBusca.Text = ""
            txtBusca.Visible = False
            strSQL = "SELECT * FROM turnos ORDER BY nomturno"
            strCampo1 = "nomturno"
            strCampo2 = "cveturno"
            Call LlenaCombos(cboBusca, strSQL, strCampo1, strCampo2)
            cboBusca.ListIndex = 0
            cboBusca.Visible = True
            cboBusca.SetFocus
        End Select
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub Actualiza_Grid()
    AdoDcBusca.Refresh
    Me.dtGrdBusca.Refresh
End Sub

Sub Presenta_datos()
    Dim intCampo As Integer
    'Obtiene el itemdata del combo
    If cboBusca.Visible = True Then
        intCampo = cboBusca.ItemData(cboBusca.ListIndex)
    End If
    strSQL = "SELECT numempleado, apellidopat, apellidomat, nombre, " & _
                "departamentos.nomdepartamento, cveempleado FROM empleados " & _
                "INNER JOIN departamentos ON empleados.departamento = " & _
                "departamentos.cvedepartamento WHERE "
    Select Case bytOpcion
        Case 0
            strSQL = strSQL & "curp = '" & Trim(txtBusca.Text) & "' "
        Case 1
            strSQL = strSQL & "rfc = '" & Trim(txtBusca.Text) & "' "
        Case 2
            strSQL = strSQL & "numempleado = " & Trim(cboBusca.Text) & " "
        Case 3
            strSQL = strSQL & "nombre = '" & Trim(cboBusca.Text) & "' "
        Case 4
            strSQL = strSQL & "apellidopat = '" & Trim(cboBusca.Text) & "' "
        Case 5
            strSQL = strSQL & "colonia = '" & Trim(cboBusca.Text) & "' "
        Case 6
            strSQL = strSQL & "entfederativa = " & intCampo & " "
        Case 7
            strSQL = strSQL & "departamento = " & intCampo & " "
        Case 8
            strSQL = strSQL & "puesto = " & intCampo & " "
        Case 9
            strSQL = strSQL & "turno = " & intCampo & " "
    End Select
    strSQL = strSQL & "ORDER BY numempleado"
    AdoDcBusca.ConnectionString = Conn
    Me.AdoDcBusca.CursorLocation = adUseClient
    Me.AdoDcBusca.LockType = adLockOptimistic
    Me.AdoDcBusca.CursorType = adOpenKeyset
    Me.AdoDcBusca.RecordSource = strSQL
    Call Actualiza_Grid
End Sub
