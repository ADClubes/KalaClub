VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmpleados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos al Catalogo de Empleados"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10186
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Generales"
      TabPicture(0)   =   "frmEmpleados.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCtrl(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCtrl(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCtrl(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCtrl(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCtrl(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblCtrl(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCtrl(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblCtrl(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblCtrl(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblCtrl(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCtrl(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblCtrl(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblCtrl(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblCtrl(13)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "DTPicker1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCtrl(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCtrl(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCtrl(2)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "dtpFechaNaci"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtCtrl(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtCtrl(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCtrl(5)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCtrl(6)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtCtrl(7)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtCtrl(8)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtCtrl(9)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtCtrl(10)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtCtrl(11)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Laborales"
      TabPicture(1)   =   "frmEmpleados.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Acceso"
      TabPicture(2)   =   "frmEmpleados.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   11
         Left            =   6000
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   10
         Left            =   3120
         TabIndex        =   27
         Text            =   "Text1"
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   5280
         Width           =   2655
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   8
         Left            =   3120
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   6
         Left            =   6000
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   5
         Left            =   3120
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   3840
         Width           =   2655
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   3000
         Width           =   8535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sexo"
         Height          =   615
         Left            =   4200
         TabIndex        =   11
         Top             =   1800
         Width           =   2655
         Begin VB.OptionButton optSexo 
            Caption         =   "Masculino"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optSexo 
            Caption         =   "Femenino"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   1935
         End
      End
      Begin MSComCtl2.DTPicker dtpFechaNaci 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62521345
         CurrentDate     =   39196
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   2
         Left            =   6120
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   960
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   62521345
         CurrentDate     =   39196
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Celular"
         Height          =   375
         Index           =   13
         Left            =   6000
         TabIndex        =   31
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Telefono 2"
         Height          =   375
         Index           =   12
         Left            =   3120
         TabIndex        =   30
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Telefono 1"
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   29
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Código Postal"
         Height          =   375
         Index           =   10
         Left            =   3120
         TabIndex        =   25
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Estado"
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Ciudad"
         Height          =   375
         Index           =   8
         Left            =   6000
         TabIndex        =   21
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Delegación/Municipio"
         Height          =   375
         Index           =   7
         Left            =   3120
         TabIndex        =   19
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Colonia"
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Direccion"
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "FechaAlta"
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Fecha Nacimiento"
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Apellido Materno"
         Height          =   375
         Index           =   2
         Left            =   6120
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Apellido Paterno"
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Nombre"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

