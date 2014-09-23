VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmCompra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compra de Títulos"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10155
   Icon            =   "frmCompra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCompra.frx":030A
   ScaleHeight     =   6075
   ScaleWidth      =   10155
   Begin TabDlg.SSTab SSTab1 
      Height          =   5370
      Left            =   105
      TabIndex        =   0
      Top             =   660
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   9472
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tipo de Comprador"
      TabPicture(0)   =   "frmCompra.frx":0614
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSiguiente(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSalir(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Datos Generales"
      TabPicture(1)   =   "frmCompra.frx":0630
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblAccionista"
      Tab(1).Control(1)=   "lblFecha"
      Tab(1).Control(2)=   "dtpFecha"
      Tab(1).Control(3)=   "txtAccionista"
      Tab(1).Control(4)=   "frmPersonales"
      Tab(1).Control(5)=   "frmEmpresa"
      Tab(1).Control(6)=   "cmdAtras(0)"
      Tab(1).Control(7)=   "cmdSiguiente(1)"
      Tab(1).Control(8)=   "cmdSalir(1)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Asignación de Títulos"
      TabPicture(2)   =   "frmCompra.frx":064C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdSalir(2)"
      Tab(2).Control(1)=   "cmdAtras(1)"
      Tab(2).Control(2)=   "cmdGuardar"
      Tab(2).Control(3)=   "cmdUnoDer"
      Tab(2).Control(4)=   "Frame1"
      Tab(2).Control(5)=   "cmdTodosDer"
      Tab(2).Control(6)=   "cmdUnoIzq"
      Tab(2).Control(7)=   "cmdTodosIzq"
      Tab(2).Control(8)=   "SSGrdDispon"
      Tab(2).Control(9)=   "SSGrdSelec"
      Tab(2).ControlCount=   10
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   2
         Left            =   -70110
         Picture         =   "frmCompra.frx":0668
         TabIndex        =   69
         ToolTipText     =   "Salir"
         Top             =   4700
         Width           =   1500
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   -70095
         Picture         =   "frmCompra.frx":0972
         TabIndex        =   68
         ToolTipText     =   "Salir"
         Top             =   4700
         Width           =   1500
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   6420
         Picture         =   "frmCompra.frx":0C7C
         TabIndex        =   67
         ToolTipText     =   "Salir"
         Top             =   4700
         Width           =   1500
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "&Siguiente >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   -66900
         Picture         =   "frmCompra.frx":0F86
         TabIndex        =   63
         ToolTipText     =   "Siguiente"
         Top             =   4700
         Width           =   1500
      End
      Begin VB.CommandButton cmdSiguiente 
         Caption         =   "&Siguiente >>"
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
         Height          =   500
         Index           =   0
         Left            =   8100
         Picture         =   "frmCompra.frx":13C8
         TabIndex        =   62
         ToolTipText     =   "Siguiente"
         Top             =   4700
         Width           =   1500
      End
      Begin VB.CommandButton cmdAtras 
         Caption         =   "<< &Atrás"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   -68500
         TabIndex        =   61
         ToolTipText     =   "Atrás"
         Top             =   4700
         Width           =   1500
      End
      Begin VB.CommandButton cmdAtras 
         Caption         =   "<< &Atrás"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   0
         Left            =   -68500
         TabIndex        =   60
         ToolTipText     =   "Atrás"
         Top             =   4700
         Width           =   1500
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Terminar y Guardar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   -66900
         Picture         =   "frmCompra.frx":180A
         TabIndex        =   59
         ToolTipText     =   "Guardar"
         Top             =   4700
         Width           =   1500
      End
      Begin VB.Frame Frame2 
         Height          =   2850
         Left            =   1650
         TabIndex        =   52
         Top             =   1140
         Width           =   6675
         Begin VB.ComboBox cboExAccionista 
            Height          =   315
            Left            =   495
            TabIndex        =   66
            Top             =   1830
            Visible         =   0   'False
            Width           =   5730
         End
         Begin VB.OptionButton optTipo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Index           =   2
            Left            =   4470
            Picture         =   "frmCompra.frx":1C4C
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   360
            Width           =   1185
         End
         Begin VB.ComboBox cboAccionista 
            Height          =   315
            Left            =   500
            TabIndex        =   55
            Top             =   1770
            Visible         =   0   'False
            Width           =   5730
         End
         Begin VB.OptionButton optTipo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Index           =   1
            Left            =   2647
            Picture         =   "frmCompra.frx":1F56
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   360
            Width           =   1185
         End
         Begin VB.OptionButton optTipo 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Index           =   0
            Left            =   825
            Picture         =   "frmCompra.frx":2260
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Ex Accionista"
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
            Left            =   2295
            TabIndex        =   65
            Top             =   1170
            Width           =   1875
         End
         Begin VB.Label lblQuienCompra 
            Caption         =   "Seleccione el Nombre de la Persona que realiza la Compra"
            Height          =   195
            Left            =   495
            TabIndex        =   58
            Top             =   1530
            Visible         =   0   'False
            Width           =   4515
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Nuevo Accionista"
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
            Left            =   4095
            TabIndex        =   57
            Top             =   1170
            Width           =   1875
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Accionista"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   705
            TabIndex        =   56
            Top             =   1170
            Width           =   1425
         End
      End
      Begin VB.Frame frmEmpresa 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   -74745
         TabIndex        =   34
         Top             =   3000
         Width           =   9420
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   11
            Left            =   7215
            TabIndex        =   41
            Top             =   1095
            Width           =   2000
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   10
            Left            =   4950
            TabIndex        =   40
            Top             =   1095
            Width           =   2000
         End
         Begin VB.ComboBox cboDelegacion 
            Height          =   315
            Index           =   1
            Left            =   200
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1095
            Width           =   4500
         End
         Begin VB.ComboBox cboEntidad 
            Height          =   315
            Index           =   1
            Left            =   6810
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   465
            Width           =   2430
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   9
            Left            =   4485
            TabIndex        =   37
            Top             =   480
            Width           =   2100
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   8
            Left            =   2340
            TabIndex        =   36
            Top             =   480
            Width           =   1950
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   7
            Left            =   200
            TabIndex        =   35
            Top             =   480
            Width           =   1965
         End
         Begin VB.Label lblDatos 
            Caption         =   "Teléfono &2:"
            Height          =   255
            Index           =   15
            Left            =   7215
            TabIndex        =   48
            Top             =   870
            Width           =   1320
         End
         Begin VB.Label lblDatos 
            Caption         =   "Teléfono &1:"
            Height          =   255
            Index           =   14
            Left            =   4950
            TabIndex        =   47
            Top             =   855
            Width           =   1305
         End
         Begin VB.Label lblDatos 
            Caption         =   "Dele&gación o Municipio:"
            Height          =   255
            Index           =   13
            Left            =   195
            TabIndex        =   46
            Top             =   855
            Width           =   1725
         End
         Begin VB.Label lblDatos 
            Caption         =   "Ent&idad Federativa:"
            Height          =   255
            Index           =   12
            Left            =   6810
            TabIndex        =   45
            Top             =   225
            Width           =   1380
         End
         Begin VB.Label lblDatos 
            Caption         =   "Co&lonia:"
            Height          =   255
            Index           =   11
            Left            =   4485
            TabIndex        =   44
            Top             =   225
            Width           =   765
         End
         Begin VB.Label lblDatos 
            Caption         =   "C&alle y Número:"
            Height          =   255
            Index           =   10
            Left            =   2340
            TabIndex        =   43
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label lblDatos 
            Caption         =   "Nom&bre:"
            Height          =   255
            Index           =   9
            Left            =   195
            TabIndex        =   42
            Top             =   225
            Width           =   1305
         End
      End
      Begin VB.Frame frmPersonales 
         Caption         =   "Personales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2220
         Left            =   -74745
         TabIndex        =   15
         Top             =   780
         Width           =   9420
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   6
            Left            =   7215
            TabIndex        =   24
            Top             =   1725
            Width           =   2000
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   5
            Left            =   4950
            TabIndex        =   23
            Top             =   1725
            Width           =   2000
         End
         Begin VB.ComboBox cboDelegacion 
            Height          =   315
            Index           =   0
            Left            =   200
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1725
            Width           =   4500
         End
         Begin VB.ComboBox cboEntidad 
            Height          =   315
            Index           =   0
            Left            =   6045
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1095
            Width           =   3165
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   4
            Left            =   2940
            TabIndex        =   20
            Top             =   1095
            Width           =   3045
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   3
            Left            =   200
            TabIndex        =   19
            Top             =   1095
            Width           =   2685
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   2
            Left            =   6510
            TabIndex        =   18
            Top             =   480
            Width           =   2685
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   1
            Left            =   3390
            TabIndex        =   17
            Top             =   480
            Width           =   2685
         End
         Begin VB.TextBox txtDatos 
            Height          =   300
            Index           =   0
            Left            =   200
            TabIndex        =   16
            Top             =   495
            Width           =   2685
         End
         Begin VB.Label lblDatos 
            Caption         =   "Teléfono &2:"
            Height          =   255
            Index           =   8
            Left            =   7215
            TabIndex        =   33
            Top             =   1470
            Width           =   930
         End
         Begin VB.Label lblDatos 
            Caption         =   "Teléfono &1:"
            Height          =   255
            Index           =   7
            Left            =   4950
            TabIndex        =   32
            Top             =   1470
            Width           =   1305
         End
         Begin VB.Label lblDatos 
            Caption         =   "&Delegación o Municipio:"
            Height          =   255
            Index           =   6
            Left            =   195
            TabIndex        =   31
            Top             =   1470
            Width           =   1725
         End
         Begin VB.Label lblDatos 
            Caption         =   "&Entidad Federativa:"
            Height          =   255
            Index           =   5
            Left            =   6045
            TabIndex        =   30
            Top             =   855
            Width           =   1380
         End
         Begin VB.Label lblDatos 
            Caption         =   "C&olonia:"
            Height          =   255
            Index           =   4
            Left            =   2940
            TabIndex        =   29
            Top             =   855
            Width           =   1320
         End
         Begin VB.Label lblDatos 
            Caption         =   "&Calle y Número:"
            Height          =   255
            Index           =   3
            Left            =   195
            TabIndex        =   28
            Top             =   855
            Width           =   1305
         End
         Begin VB.Label lblDatos 
            Caption         =   "&Nombre(s):"
            Height          =   255
            Index           =   2
            Left            =   6510
            TabIndex        =   27
            Top             =   225
            Width           =   765
         End
         Begin VB.Label lblDatos 
            Caption         =   "Apellido &Materno:"
            Height          =   255
            Index           =   1
            Left            =   3390
            TabIndex        =   26
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label lblDatos 
            Caption         =   "Apellido &Paterno:"
            Height          =   255
            Index           =   0
            Left            =   195
            TabIndex        =   25
            Top             =   225
            Width           =   1305
         End
      End
      Begin VB.TextBox txtAccionista 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Left            =   -71430
         TabIndex        =   14
         Top             =   435
         Width           =   1860
      End
      Begin VB.CommandButton cmdUnoDer 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70575
         TabIndex        =   11
         ToolTipText     =   "Seleccionar"
         Top             =   1890
         Width           =   1110
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   -74790
         TabIndex        =   5
         Top             =   585
         Width           =   9570
         Begin VB.OptionButton optCompra 
            Caption         =   "Comprar Acciones a un &Accionista"
            Height          =   330
            Index           =   1
            Left            =   915
            TabIndex        =   7
            Top             =   465
            Width           =   2745
         End
         Begin VB.OptionButton optCompra 
            Caption         =   "Comprar Acciones al &Club"
            Height          =   330
            Index           =   0
            Left            =   915
            TabIndex        =   6
            Top             =   150
            Width           =   2310
         End
         Begin SSDataWidgets_B.SSDBCombo SSCboAccionista 
            Height          =   315
            Left            =   4200
            TabIndex        =   8
            Top             =   390
            Width           =   810
            DataFieldList   =   "Column 0"
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
            DefColWidth     =   176
            RowHeight       =   423
            Columns.Count   =   2
            Columns(0).Width=   1244
            Columns(0).Caption=   "Número"
            Columns(0).Name =   "NUMERO"
            Columns(0).Alignment=   1
            Columns(0).DataField=   "Column 0"
            Columns(0).FieldLen=   256
            Columns(1).Width=   5794
            Columns(1).Caption=   "Accionista"
            Columns(1).Name =   "ACCIONISTA"
            Columns(1).DataField=   "Column 1"
            Columns(1).FieldLen=   256
            _ExtentX        =   1429
            _ExtentY        =   556
            _StockProps     =   93
            BackColor       =   -2147483643
         End
         Begin VB.Label lblComboAccionista 
            Caption         =   "A &Quien Le Compra:"
            Height          =   255
            Left            =   4200
            TabIndex        =   10
            Top             =   150
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.Label lblNombre 
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
            Height          =   285
            Left            =   5160
            TabIndex        =   9
            Top             =   405
            Width           =   4155
         End
      End
      Begin VB.CommandButton cmdTodosDer 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70575
         TabIndex        =   4
         ToolTipText     =   "Seleccionar Todo"
         Top             =   2355
         Width           =   1110
      End
      Begin VB.CommandButton cmdUnoIzq 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70575
         TabIndex        =   3
         ToolTipText     =   "Des-seleccionar"
         Top             =   3150
         Width           =   1110
      End
      Begin VB.CommandButton cmdTodosIzq 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70575
         TabIndex        =   2
         ToolTipText     =   "Des-seleccionar Todo"
         Top             =   3630
         Width           =   1110
      End
      Begin SSDataWidgets_B.SSDBGrid SSGrdDispon 
         Height          =   2800
         Left            =   -73860
         TabIndex        =   12
         Top             =   1695
         Width           =   2835
         _Version        =   196616
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   3
         HeadFont3D      =   3
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   1270
         Columns(0).Caption=   "SERIE"
         Columns(0).Name =   "SERIE"
         Columns(0).Alignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   1667
         Columns(1).Caption=   "NUMERO"
         Columns(1).Name =   "NUMERO"
         Columns(1).Alignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   1005
         Columns(2).Caption=   "TIPO"
         Columns(2).Name =   "TIPO"
         Columns(2).Alignment=   2
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         _ExtentX        =   5001
         _ExtentY        =   4939
         _StockProps     =   79
         Caption         =   "DISPONIBLES"
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
      Begin SSDataWidgets_B.SSDBGrid SSGrdSelec 
         Height          =   2800
         Left            =   -69045
         TabIndex        =   13
         Top             =   1680
         Width           =   2835
         _Version        =   196616
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Col.Count       =   3
         HeadFont3D      =   3
         RowHeight       =   423
         Columns.Count   =   3
         Columns(0).Width=   1270
         Columns(0).Caption=   "SERIE"
         Columns(0).Name =   "SERIE"
         Columns(0).Alignment=   2
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(0).Locked=   -1  'True
         Columns(1).Width=   1667
         Columns(1).Caption=   "NUMERO"
         Columns(1).Name =   "NUMERO"
         Columns(1).Alignment=   2
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   2
         Columns(1).FieldLen=   256
         Columns(1).Locked=   -1  'True
         Columns(2).Width=   1005
         Columns(2).Caption=   "TIPO"
         Columns(2).Name =   "TIPO"
         Columns(2).Alignment=   2
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(2).Locked=   -1  'True
         _ExtentX        =   5001
         _ExtentY        =   4939
         _StockProps     =   79
         Caption         =   "SELECCIONADOS"
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
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   -67890
         TabIndex        =   49
         Top             =   450
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   37957
      End
      Begin VB.Label lblFecha 
         Caption         =   "&Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -68685
         TabIndex        =   51
         Top             =   570
         Width           =   690
      End
      Begin VB.Label lblAccionista 
         Caption         =   "P&ropietario No:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -73200
         TabIndex        =   50
         Top             =   540
         Width           =   1665
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "COMPRA DE TÍTULOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   3330
      TabIndex        =   1
      Top             =   90
      Width           =   4170
   End
End
Attribute VB_Name = "frmCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA COMPRA DE ACCIONES
' Objetivo: PERMITE LA COMPRA DE ACCIONES
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim blnExAccionista As Boolean
    Dim AdoRcsAccionista As ADODB.Recordset
    Dim AdoRcsExAccionista As ADODB.Recordset
    
Private Sub cboAccionista_Click()
    Dim i As Integer
    Dim lngAccionista As Long
    If cboAccionista.ListIndex < 0 Then
        Exit Sub
    End If
    lngAccionista = cboAccionista.ItemData(cboAccionista.ListIndex)
    strSQL = "SELECT * FROM accionistas WHERE idproptitulo = " & lngAccionista
    Set AdoRcsAccionista = New ADODB.Recordset
    AdoRcsAccionista.ActiveConnection = Conn
    AdoRcsAccionista.LockType = adLockOptimistic
    AdoRcsAccionista.CursorType = adOpenKeyset
    AdoRcsAccionista.CursorLocation = adUseServer
    AdoRcsAccionista.Open strSQL
    If Not AdoRcsAccionista.EOF Then
        txtAccionista.Text = lngAccionista
        txtAccionista.Enabled = False
        txtDatos(0).Text = AdoRcsAccionista!A_Paterno
        txtDatos(1).Text = AdoRcsAccionista!A_Materno
        txtDatos(2).Text = AdoRcsAccionista!Nombre
        txtDatos(3).Text = AdoRcsAccionista!calle
        txtDatos(4).Text = AdoRcsAccionista!colonia
        txtDatos(5).Text = AdoRcsAccionista!telefono_1
        txtDatos(6).Text = AdoRcsAccionista!telefono_2
        txtDatos(7).Text = AdoRcsAccionista!Empresa
        txtDatos(8).Text = AdoRcsAccionista!e_calle
        txtDatos(9).Text = AdoRcsAccionista!e_colonia
        txtDatos(10).Text = AdoRcsAccionista!e_telefono_1
        txtDatos(11).Text = AdoRcsAccionista!e_telefono_2
        For i = 0 To 1
            If MuestraElementoCombo(cboEntidad(i), AdoRcsAccionista!ent_federativa) Then
            End If
            If MuestraElementoCombo(cboDelegacion(i), AdoRcsAccionista!delegamunici) Then
            End If
        Next i
    End If
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = True
    cmdSiguiente(0).Enabled = True
End Sub

Private Sub cboEntidad_Click(index As Integer)
    Dim lngCveDeloMuni As Long
    Dim strCampo1, strCampo2 As String
    
    'Llena combo Delegación o Municipio
    If cboEntidad(index).ListIndex < 0 Then Exit Sub
    lngCveDeloMuni = cboEntidad(index).ItemData(cboEntidad(index).ListIndex)
    strSQL = "SELECT cvedelomuni, nomdelomuni FROM delgamunici " & _
                    "WHERE entidadfed = " & lngCveDeloMuni & " ORDER BY nomdelomuni"
    strCampo1 = "nomdelomuni"
    strCampo2 = "cvedelomuni"
    Call LlenaCombos(cboDelegacion(index), strSQL, strCampo1, strCampo2)
End Sub

Private Sub cboExAccionista_Click()
    Dim lngAccionista As Long
    Dim intRespuesta As Integer
    If cboExAccionista.ListIndex < 0 Then
        Exit Sub
    End If
    lngAccionista = cboExAccionista.ItemData(cboExAccionista.ListIndex)
    strSQL = "SELECT * FROM exaccionistas WHERE idexaccionista = " & lngAccionista
    Set AdoRcsExAccionista = New ADODB.Recordset
    AdoRcsExAccionista.ActiveConnection = Conn
    AdoRcsExAccionista.LockType = adLockOptimistic
    AdoRcsExAccionista.CursorType = adOpenKeyset
    AdoRcsExAccionista.CursorLocation = adUseServer
    AdoRcsExAccionista.Open strSQL
    If Not AdoRcsExAccionista.EOF Then
        Call LlenaTxtAccionista
        txtDatos(0).Text = AdoRcsExAccionista!A_Paterno
        txtDatos(1).Text = AdoRcsExAccionista!A_Materno
        txtDatos(2).Text = AdoRcsExAccionista!Nombre
        txtDatos(3).Text = AdoRcsExAccionista!calle
        txtDatos(4).Text = AdoRcsExAccionista!colonia
        txtDatos(5).Text = AdoRcsExAccionista!telefono_1
        txtDatos(6).Text = AdoRcsExAccionista!telefono_2
        txtDatos(7).Text = AdoRcsExAccionista!Empresa
        txtDatos(8).Text = AdoRcsExAccionista!e_calle
        txtDatos(9).Text = AdoRcsExAccionista!e_colonia
        txtDatos(10).Text = AdoRcsExAccionista!e_telefono_1
        txtDatos(11).Text = AdoRcsExAccionista!e_telefono_2
        cboEntidad(0).Text = AdoRcsExAccionista!ent_federativa
        cboEntidad(1).Text = AdoRcsExAccionista!e_ent_federativa
        cboDelegacion(0).Text = AdoRcsExAccionista!delegamunici
        cboDelegacion(1).Text = AdoRcsExAccionista!e_delegamunici
    End If
    strSQL = "SELECT * FROM accionistas WHERE a_paterno = '" & txtDatos(0).Text & _
                    "' AND a_materno = '" & Trim(txtDatos(1).Text) & "' AND nombre = '" & _
                    txtDatos(2).Text & "'"
    Set AdoRcsAccionista = New ADODB.Recordset
    AdoRcsAccionista.ActiveConnection = Conn
    AdoRcsAccionista.LockType = adLockOptimistic
    AdoRcsAccionista.CursorType = adOpenKeyset
    AdoRcsAccionista.CursorLocation = adUseServer
    AdoRcsAccionista.Open strSQL
    If Not AdoRcsAccionista.EOF Then
        intRespuesta = MsgBox("Ya existe un Accionista con el Nombre de: " & _
                                                AdoRcsAccionista!Nombre & " " & _
                                                AdoRcsAccionista!A_Paterno & " " & _
                                                AdoRcsAccionista!A_Materno & _
                                            ". ¿Desea crear otro Accionista con el mismo Nombre? ", _
                                            vbOKCancel + vbQuestion, "Accionista Duplicado")
        If intRespuesta <> 1 Then
            SSTab1.TabVisible(1) = False
            SSTab1.TabVisible(2) = False
            cmdSiguiente(0).Enabled = False
            Exit Sub
        End If
    End If
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(2) = True
    cmdSiguiente(0).Enabled = True
End Sub

Private Sub cmdAtras_Click(index As Integer)
    Select Case index
        Case 0
            SSTab1.Tab = 0
        Case 1
            SSTab1.Tab = 1
    End Select
End Sub

Private Sub cmdGuardar_Click()
    If VerificaDatos = True Then
        Call GuardaDatos
    End If
End Sub

Private Sub cmdSalir_Click(index As Integer)
    Unload Me
End Sub

Private Sub cmdSiguiente_Click(index As Integer)
    Select Case index
        Case 0
            SSTab1.Tab = 1
        Case 1
            SSTab1.Tab = 2
    End Select
End Sub

Private Sub cmdTodosDer_Click()
Dim i As Integer
    SSGrdDispon.MoveFirst
    For i = 0 To SSGrdDispon.Rows - 1
        SSGrdDispon.SelBookmarks.Add SSGrdDispon.Bookmark
        SSGrdDispon.MoveNext
    Next i
    Call Mueve_A_La_Derecha
    If optCompra(1).Value = True Then
        blnExAccionista = True
    End If
End Sub

Private Sub cmdTodosIzq_Click()
Dim i As Integer
    SSGrdSelec.MoveFirst
    For i = 0 To SSGrdSelec.Rows - 1
        SSGrdSelec.SelBookmarks.Add SSGrdSelec.Bookmark
        SSGrdSelec.MoveNext
    Next i
    Call Mueve_A_La_Izquierda
    blnExAccionista = False
End Sub

Private Sub cmdUnoDer_Click()
    Call Mueve_A_La_Derecha
    If SSGrdDispon.Rows = 0 And optCompra(1).Value = True Then
        blnExAccionista = True
    End If
End Sub

Private Sub cmdUnoIzq_Click()
    Call Mueve_A_La_Izquierda
    blnExAccionista = False
End Sub

Private Sub Form_Activate()
    Unload frmSelecReportes
    Me.Top = 0
    Me.Left = 0
    Me.Height = 6525
    Me.Width = 10245
End Sub

Private Sub Form_Load()
    Dim strCampo1, strCampo2 As String
    blnExAccionista = False
    optTipo(0).Value = False
    optTipo(1).Value = False
    'Llena Combo Accionistas
    strSQL = "SELECT idproptitulo, a_paterno & ' ' & a_materno & ' ' & nombre as accionistas" & _
                    " FROM accionistas ORDER BY a_paterno, a_materno, nombre"
    strCampo1 = "accionistas"
    strCampo2 = "idproptitulo"
    Call LlenaCombos(cboAccionista, strSQL, strCampo1, strCampo2)
    'Llena Combo Ex Accionistas
    strSQL = "SELECT idexaccionista, a_paterno & ' ' & a_materno & ' ' & nombre as accionistas" & _
                    " FROM exaccionistas ORDER BY a_paterno, a_materno, nombre"
    strCampo1 = "accionistas"
    strCampo2 = "idexaccionista"
    Call LlenaCombos(cboExAccionista, strSQL, strCampo1, strCampo2)
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    dtpFecha.Value = Now
    optCompra(0).Value = True
    'Llena los combos de Entidad Federativa
    strSQL = "SELECT cveentfederativa, nomentfederativa FROM entfederativa"
    strCampo1 = "nomentfederativa"
    strCampo2 = "cveentfederativa"
    Call LlenaCombos(cboEntidad(0), strSQL, strCampo1, strCampo2)
    Call LlenaCombos(cboEntidad(1), strSQL, strCampo1, strCampo2)
    Call LlenaComboAccionistas
End Sub

Private Sub optCompra_Click(index As Integer)
    lblNombre.Caption = ""
    Select Case index
        Case 0
            lblComboAccionista.Visible = False
            SSCboAccionista.Visible = False
        Case 1
            lblComboAccionista.Visible = True
            SSCboAccionista.Visible = True
    End Select
    Call ActualizaGrid
End Sub

Private Sub optTipo_Click(index As Integer)
    Select Case index
        Case 0                           'Accionista ya existente
            cboAccionista.Text = ""
            lblQuienCompra.Visible = True
            cboAccionista.Visible = True
            cboExAccionista.Visible = False
            SSTab1.TabVisible(1) = False
            SSTab1.TabVisible(2) = False
            cmdSiguiente(0).Enabled = False
            optCompra(0).Value = True
            txtAccionista.Enabled = False
        Case 1                          ' Ex Accionista
            cboExAccionista.Text = ""
            lblQuienCompra.Visible = True
            cboExAccionista.Visible = True
            cboAccionista.Visible = False
            SSTab1.TabVisible(1) = False
            SSTab1.TabVisible(2) = False
            cmdSiguiente(0).Enabled = False
            optCompra(0).Value = True
            txtAccionista.Enabled = True
        Case 2                          'Nuevo accionista
            Call Limpia
            optCompra(0).Value = True
            SSTab1.TabVisible(1) = True
            SSTab1.TabVisible(2) = True
            Call LlenaTxtAccionista
            txtAccionista.Enabled = True
            cmdSiguiente(0).Enabled = True
            SSTab1.Tab = 1
    End Select
End Sub

Private Sub SSCboAccionista_Click()
    If (SSCboAccionista.Columns(1).Text = cboAccionista.Text) Or (SSCboAccionista.Columns(1).Text = cboExAccionista.Text) Then
        MsgBox "¡ No puede Seleccionar al accionista " & SSCboAccionista.Columns(1).Text & ", pues es quien realiza la compra !, ¡ Por favor seleccione a otro !", vbExclamation
        SSCboAccionista.Text = ""
        SSCboAccionista.SetFocus
        Exit Sub
    End If
    Call ActualizaGrid
    lblNombre.Caption = SSCboAccionista.Columns(1).Text
End Sub

Private Sub SSGrdDispon_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
    DispPromptMsg = 0
End Sub

Private Sub SSGrdSelec_BeforeDelete(Cancel As Integer, DispPromptMsg As Integer)
    DispPromptMsg = 0
End Sub

Private Sub txtAccionista_GotFocus()
    txtAccionista.SelStart = 0
    txtAccionista.SelLength = Len(txtAccionista)
End Sub

Private Sub txtAccionista_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 22, 48 To 57     'Backspace, <Ctrl+V> y del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtDatos_GotFocus(index As Integer)
    txtDatos(index).SelStart = 0
    txtDatos(index).SelLength = Len(txtDatos(index))
End Sub

Private Sub txtDatos_KeyPress(index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function VerificaDatos()
    If SSGrdSelec.Rows = 0 Then
        MsgBox "¡ Para ser Accionista debe Comprar al menos un Título !", _
                    vbOKOnly + vbExclamation, "Propietarios de Títulos (Captura)"
        VerificaDatos = False
        SSGrdDispon.SetFocus
        SSTab1.Tab = 2
        Exit Function
    End If
    If (txtAccionista.Text = "") Then
        MsgBox "¡ Favor de Digitar el Número del Accionista, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Propietarios de Títulos (Captura)"
        VerificaDatos = False
        txtAccionista.SetFocus
        SSTab1.Tab = 1
        Exit Function
    End If
    If (txtDatos(0).Text = "") Then
        MsgBox "¡ Favor de Digitar el Apellido Paterno, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Propietarios de Títulos (Captura)"
        VerificaDatos = False
        txtDatos(0).SetFocus
        SSTab1.Tab = 1
        Exit Function
    End If
    If (txtDatos(2).Text = "") Then
        MsgBox "¡ Favor de Digitar el Nombre, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Propietarios de Títulos (Captura)"
        VerificaDatos = False
        txtDatos(2).SetFocus
        SSTab1.Tab = 1
        Exit Function
    End If
    VerificaDatos = True
End Function

Sub ActualizaGrid()
    SSGrdDispon.RemoveAll
    SSGrdSelec.RemoveAll
    cmdUnoDer.Enabled = False
    cmdTodosDer.Enabled = False
    cmdUnoIzq.Enabled = False
    cmdTodosIzq.Enabled = False
    strSQL = "SELECT serie, numero, tipo FROM titulos WHERE idpropietario = "
    If optCompra(0).Value = True Then
        strSQL = strSQL & "0"
    Else
        strSQL = strSQL & IIf(SSCboAccionista.Text = "", -1, SSCboAccionista.Text)
    End If
    strSQL = strSQL & " ORDER BY numero"
    Set AdoRcsAccionista = New ADODB.Recordset
    AdoRcsAccionista.ActiveConnection = Conn
    AdoRcsAccionista.LockType = adLockOptimistic
    AdoRcsAccionista.CursorType = adOpenKeyset
    AdoRcsAccionista.CursorLocation = adUseServer
    AdoRcsAccionista.Open strSQL
    If Not AdoRcsAccionista.EOF Then
        Do While Not AdoRcsAccionista.EOF
            SSGrdDispon.AddItem AdoRcsAccionista!Serie + _
            Chr$(9) + Str(AdoRcsAccionista!Numero) + _
            Chr$(9) + AdoRcsAccionista!tipo
            AdoRcsAccionista.MoveNext
        Loop
        cmdUnoDer.Enabled = True
        cmdTodosDer.Enabled = True
    End If
End Sub

Sub GuardaDatos()
    Dim i As Integer
    Dim lngNumero As Long
    Dim strTipo, strSerie As String
    Dim AdoCmdInserta As ADODB.Command
    On Error GoTo err_Guarda
    cmdGuardar.Enabled = False
    For i = 0 To 2
        cmdSalir(i).Enabled = False
    Next i
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    If (optTipo(1).Value = True) Or (optTipo(2).Value = True) Then
        'Guarda los datos del Nuevo Accionista en la tabla ACCIONISTAS
        #If SqlServer_ Then
            strSQL = "INSERT INTO accionistas (idproptitulo, fecha_alta, a_paterno, " & _
                        "a_materno, nombre, calle, colonia, ent_federativa, delegamunici, " & _
                        "telefono_1, telefono_2, empresa, e_calle, e_colonia, " & _
                        "e_ent_federativa, e_delegamunici, e_telefono_1, e_telefono_2) " & _
                        "VALUES (" & Val(txtAccionista.Text) & ", '" & _
                        Format(dtpFecha, "yyyymmdd") & "', '" & txtDatos(0).Text & _
                        "', '" & txtDatos(1).Text & "', '" & txtDatos(2).Text & "', '" & _
                        txtDatos(3).Text & "', '" & txtDatos(4).Text & "', '" & _
                        cboEntidad(0).Text & "', '" & cboDelegacion(0).Text & "', '" & _
                        txtDatos(5).Text & "', '" & txtDatos(6).Text & "', '" & _
                        txtDatos(7).Text & "', '" & txtDatos(8).Text & "', '" & _
                        txtDatos(9).Text & "', '" & cboEntidad(1).Text & "', '" & _
                        cboDelegacion(1).Text & "', '" & txtDatos(10).Text & "', '" & _
                        txtDatos(11).Text & "')"
        #Else
            strSQL = "INSERT INTO accionistas (idproptitulo, fecha_alta, a_paterno, " & _
                        "a_materno, nombre, calle, colonia, ent_federativa, delegamunici, " & _
                        "telefono_1, telefono_2, empresa, e_calle, e_colonia, " & _
                        "e_ent_federativa, e_delegamunici, e_telefono_1, e_telefono_2) " & _
                        "VALUES (" & Val(txtAccionista.Text) & ", #" & _
                        Format(dtpFecha, "mm/dd/yyyy") & "#, '" & txtDatos(0).Text & _
                        "', '" & txtDatos(1).Text & "', '" & txtDatos(2).Text & "', '" & _
                        txtDatos(3).Text & "', '" & txtDatos(4).Text & "', '" & _
                        cboEntidad(0).Text & "', '" & cboDelegacion(0).Text & "', '" & _
                        txtDatos(5).Text & "', '" & txtDatos(6).Text & "', '" & _
                        txtDatos(7).Text & "', '" & txtDatos(8).Text & "', '" & _
                        txtDatos(9).Text & "', '" & cboEntidad(1).Text & "', '" & _
                        cboDelegacion(1).Text & "', '" & txtDatos(10).Text & "', '" & _
                        txtDatos(11).Text & "')"
        #End If
        
        Set AdoCmdInserta = New ADODB.Command
        AdoCmdInserta.ActiveConnection = Conn
        AdoCmdInserta.CommandText = strSQL
        AdoCmdInserta.Execute
    End If
    
    SSGrdSelec.MoveFirst
    For i = 0 To SSGrdSelec.Rows - 1
        strSerie = SSGrdSelec.Columns("serie").CellValue(SSGrdSelec.Bookmark)
        strTipo = SSGrdSelec.Columns("tipo").CellValue(SSGrdSelec.Bookmark)
        lngNumero = Val(SSGrdSelec.Columns("numero").CellValue(SSGrdSelec.Bookmark))
        
        'Actualiza la información de los títulos comprados en la tabla TÍTULOS
        #If SqlServer_ Then
            strSQL = "UPDATE titulos SET idpropietario = " & Val(txtAccionista.Text) & _
                        ",  fecha_asignacion = '" & Format(dtpFecha.Value, "yyyymmdd") & _
                        "' WHERE tipo = '" & strTipo & "' AND numero = " & lngNumero & _
                        " AND serie = '" & strSerie & "'"
        #Else
            strSQL = "UPDATE titulos SET idpropietario = " & Val(txtAccionista.Text) & _
                        ",  fecha_asignacion = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & _
                        "# WHERE tipo = '" & strTipo & "' AND numero = " & lngNumero & _
                        " AND serie = '" & strSerie & "'"
        #End If
        
        Set AdoCmdInserta = New ADODB.Command
        AdoCmdInserta.ActiveConnection = Conn
        AdoCmdInserta.CommandText = strSQL
        AdoCmdInserta.Execute
        
        'Insertamos los movimientos en el Histórico en la tabla HISTOACCIONES
        #If SqlServer_ Then
            strSQL = "INSERT INTO histoacciones (tipo, numero, serie, " & _
                       "tipomovimiento, fechamovimiento, propanterior, propactual) " & _
                       "VALUES ('" & strTipo & "', " & lngNumero & ", '" & strSerie & _
                       "', 'COMPRA', '" & Format(dtpFecha.Value, "yyyymmdd") & "', '"
        #Else
            strSQL = "INSERT INTO histoacciones (tipo, numero, serie, " & _
                       "tipomovimiento, fechamovimiento, propanterior, propactual) " & _
                       "VALUES ('" & strTipo & "', " & lngNumero & ", '" & strSerie & _
                       "', 'COMPRA', #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#, '"
        #End If
        
        If optCompra(0).Value = True Then
            strSQL = strSQL & "CLUB', '"
        Else
            strSQL = strSQL & lblNombre.Caption & "', '"
        End If
        strSQL = strSQL & txtDatos(0).Text & " " & txtDatos(1).Text & " " & _
                        txtDatos(2).Text & "')"

        Set AdoCmdInserta = New ADODB.Command
        AdoCmdInserta.ActiveConnection = Conn
        AdoCmdInserta.CommandText = strSQL
        AdoCmdInserta.Execute
        
        SSGrdSelec.MoveNext
    Next i
    
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    If blnExAccionista = True Then
        Call EliminaAccionista(Val(SSCboAccionista.Text))
    End If
    MsgBox "¡ Transacción Terminada !", vbOKOnly + vbInformation, "Propietarios de Títulos"
    cmdGuardar.Enabled = True
    For i = 0 To 2
        cmdSalir(i).Enabled = True
    Next i
    Call Limpia
    optTipo(1).Value = False
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    cmdGuardar.Enabled = True
    For i = 0 To 2
        cmdSalir(i).Enabled = True
    Next i
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Sub Limpia()
    Dim i As Integer
    Dim strCampo1, strCampo2 As String
    blnExAccionista = False
    Call LlenaTxtAccionista
    dtpFecha.Value = Now
    For i = 0 To 11
        txtDatos(i).Text = ""
    Next
    For i = 0 To 1
        cboEntidad(i).ListIndex = -1
        cboDelegacion(i).ListIndex = -1
    Next i
    SSGrdDispon.RemoveAll
    SSGrdSelec.RemoveAll
    optCompra(0).Value = False
    optCompra(0).Enabled = True
    optCompra(1).Value = False
    optCompra(1).Enabled = True
    SSCboAccionista.Text = ""
    lblNombre.Caption = ""
    lblComboAccionista.Visible = False
    SSCboAccionista.Enabled = True
    SSCboAccionista.Visible = False
    lblComboAccionista.Enabled = True
    'Llena Combo Accionistas
    strSQL = "SELECT idproptitulo, a_paterno & ' ' & a_materno & ' ' & nombre as accionistas" & _
                    " FROM accionistas ORDER BY a_paterno, a_materno, nombre"
    strCampo1 = "accionistas"
    strCampo2 = "idproptitulo"
    Call LlenaCombos(cboAccionista, strSQL, strCampo1, strCampo2)
    'Llena Combo Ex Accionistas
    strSQL = "SELECT idexaccionista, a_paterno & ' ' & a_materno & ' ' & nombre as accionistas" & _
                    " FROM exaccionistas ORDER BY a_paterno, a_materno, nombre"
    strCampo1 = "accionistas"
    strCampo2 = "idexaccionista"
    Call LlenaCombos(cboExAccionista, strSQL, strCampo1, strCampo2)
    Call LlenaComboAccionistas
    cboAccionista.ListIndex = -1
    cboAccionista.Visible = False
    cboExAccionista.Visible = False
    optTipo(0).Value = False
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(2) = False
    lblQuienCompra.Visible = False
    cmdSiguiente(0).Enabled = False
End Sub

Sub LlenaComboAccionistas()
    SSCboAccionista.RemoveAll
        
    strSQL = "SELECT idproptitulo, a_paterno & ' ' & a_materno & ' ' & nombre as accionistas" & _
                    " FROM accionistas ORDER BY a_paterno, a_materno, nombre"
    Set AdoRcsAccionista = New ADODB.Recordset
    AdoRcsAccionista.ActiveConnection = Conn
    AdoRcsAccionista.LockType = adLockOptimistic
    AdoRcsAccionista.CursorType = adOpenKeyset
    AdoRcsAccionista.CursorLocation = adUseServer
    AdoRcsAccionista.Open strSQL
    If Not AdoRcsAccionista.EOF Then
        Do While Not AdoRcsAccionista.EOF
            SSCboAccionista.AddItem Str(AdoRcsAccionista!idproptitulo) + Chr$(9) + _
                                                        AdoRcsAccionista!accionistas
            AdoRcsAccionista.MoveNext
        Loop
    End If
End Sub

Sub LlenaTxtAccionista()
    Dim lngAnterior, lngAccionista As Long
    'Llena txtAccionista
    lngAccionista = 1
    strSQL = "SELECT idproptitulo FROM accionistas ORDER BY idproptitulo"
    Set AdoRcsAccionista = New ADODB.Recordset
    AdoRcsAccionista.ActiveConnection = Conn
    AdoRcsAccionista.LockType = adLockOptimistic
    AdoRcsAccionista.CursorType = adOpenKeyset
    AdoRcsAccionista.CursorLocation = adUseServer
    AdoRcsAccionista.Open strSQL
    If AdoRcsAccionista.EOF Then
        lngAccionista = 1
        txtAccionista.Text = lngAccionista
        Exit Sub
    End If
    AdoRcsAccionista.MoveFirst
    Do While Not AdoRcsAccionista.EOF
        If AdoRcsAccionista.Fields!idproptitulo <> "1" Then
            If Val(AdoRcsAccionista.Fields!idproptitulo) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngAccionista
        AdoRcsAccionista.MoveNext
        If Not AdoRcsAccionista.EOF Then lngAccionista = AdoRcsAccionista.Fields!idproptitulo
    Loop
    txtAccionista.Text = lngAnterior + 1
End Sub

Sub Mueve_A_La_Derecha()
    Dim i As Integer
    'Copia los registros seleccionados del primer grid al segundo
    For i = 0 To SSGrdDispon.SelBookmarks.Count - 1
        SSGrdSelec.AddItem SSGrdDispon.Columns("serie").CellValue(SSGrdDispon.SelBookmarks(i)) + Chr$(9) + _
                                         Str(SSGrdDispon.Columns("numero").CellValue(SSGrdDispon.SelBookmarks(i))) + Chr$(9) + _
                                         SSGrdDispon.Columns("tipo").CellValue(SSGrdDispon.SelBookmarks(i))
        cmdUnoIzq.Enabled = True
        cmdTodosIzq.Enabled = True
    Next i
    
    'Elimina los renglones seleccionados del primer grid
    SSGrdDispon.DeleteSelected
    DoEvents
    'Verifica si se ha quedado vacío el grid de la izquierda
    If SSGrdDispon.Rows = 0 Then
        cmdUnoDer.Enabled = False
        cmdTodosDer.Enabled = False
    End If
    optCompra(0).Enabled = False
    optCompra(1).Enabled = False
    SSCboAccionista.Enabled = False
    lblComboAccionista.Enabled = False
End Sub

Sub Mueve_A_La_Izquierda()
Dim i As Integer
    'Copia los registros seleccionados del segundo grid al primero
    For i = 0 To SSGrdSelec.SelBookmarks.Count - 1
        SSGrdDispon.AddItem SSGrdSelec.Columns("serie").CellValue(SSGrdSelec.SelBookmarks(i)) + Chr$(9) + _
                                          Str(SSGrdSelec.Columns("numero").CellValue(SSGrdSelec.SelBookmarks(i))) + Chr$(9) + _
                                          SSGrdSelec.Columns("tipo").CellValue(SSGrdSelec.SelBookmarks(i))
        cmdUnoDer.Enabled = True
        cmdTodosDer.Enabled = True
    Next i
    
    'Elimina los renglones seleccionados del segundo grid
    SSGrdSelec.DeleteSelected
        
    'Verifica si se ha quedado vacío el grid de la derecha
    If SSGrdSelec.Rows = 0 Then
        cmdUnoIzq.Enabled = False
        cmdTodosIzq.Enabled = False
    End If
End Sub

Private Sub txtDatos_LostFocus(index As Integer)
    Dim j, intRespuesta As Integer
    If (index <> 0) And (index <> 1) And (index <> 2) Then Exit Sub
    strSQL = "SELECT * FROM accionistas WHERE a_paterno = '" & txtDatos(0).Text & _
                    "' AND a_materno = '" & Trim(txtDatos(1).Text) & "' AND nombre = '" & _
                    txtDatos(2).Text & "'"
    Set AdoRcsAccionista = New ADODB.Recordset
    AdoRcsAccionista.ActiveConnection = Conn
    AdoRcsAccionista.LockType = adLockOptimistic
    AdoRcsAccionista.CursorType = adOpenKeyset
    AdoRcsAccionista.CursorLocation = adUseServer
    AdoRcsAccionista.Open strSQL
    If Not AdoRcsAccionista.EOF Then
        intRespuesta = MsgBox("Ya existe un Accionista con el Nombre de: " & _
                                                AdoRcsAccionista!Nombre & " " & _
                                                AdoRcsAccionista!A_Paterno & " " & _
                                                AdoRcsAccionista!A_Materno & _
                                            ". ¿Desea crear otro Accionista con el mismo Nombre? ", _
                                            vbOKCancel + vbQuestion, "Accionista Duplicado")
        If intRespuesta <> 1 Then
            For j = 0 To 2
                txtDatos(j).Text = ""
            Next j
            txtDatos(0).SetFocus
            Exit Sub
        End If
    End If
End Sub
