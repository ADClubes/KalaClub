VERSION 5.00
Begin VB.Form frmBajas 
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   8805
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
      Height          =   2655
      Left            =   1395
      TabIndex        =   10
      Top             =   1575
      Width           =   7245
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
         Left            =   100
         TabIndex        =   26
         Top             =   285
         Width           =   1605
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
         TabIndex        =   25
         Top             =   615
         Width           =   1605
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
         Left            =   600
         TabIndex        =   24
         Top             =   945
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
         Left            =   105
         TabIndex        =   23
         Top             =   1275
         Width           =   1575
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
         Left            =   4905
         TabIndex        =   22
         Top             =   1275
         Width           =   1095
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
         Left            =   855
         TabIndex        =   21
         Top             =   1605
         Width           =   795
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
         TabIndex        =   20
         Top             =   2265
         Width           =   1605
      End
      Begin VB.Label lblPersonales 
         Caption         =   "CUAUHTEMOC 55 - 5"
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
         TabIndex        =   19
         Top             =   255
         Width           =   5265
      End
      Begin VB.Label lblPersonales 
         Caption         =   "BARRIO SAN PEDRO"
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
         TabIndex        =   18
         Top             =   585
         Width           =   5235
      End
      Begin VB.Label lblPersonales 
         Caption         =   "D.F."
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
         Left            =   1900
         TabIndex        =   17
         Top             =   1245
         Width           =   3555
      End
      Begin VB.Label lblPersonales 
         Caption         =   "16090"
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
         Left            =   6120
         TabIndex        =   16
         Top             =   1245
         Width           =   1035
      End
      Begin VB.Label lblPersonales 
         Caption         =   "XOCHIMILCO"
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
         TabIndex        =   15
         Top             =   915
         Width           =   5280
      End
      Begin VB.Label lblPersonales 
         Caption         =   "HEMD641229HV7"
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
         TabIndex        =   14
         Top             =   1575
         Width           =   1950
      End
      Begin VB.Label lblPersonales 
         Caption         =   "54 89 28 79"
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
         TabIndex        =   13
         Top             =   2235
         Width           =   5160
      End
      Begin VB.Label lblPersonales 
         Caption         =   "HEMD641229HDFRRV03"
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
         TabIndex        =   12
         Top             =   1905
         Width           =   3015
      End
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
         Left            =   825
         TabIndex        =   11
         Top             =   1935
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
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
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   990
      Width           =   990
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   7665
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   435
      Width           =   990
   End
   Begin VB.Frame frmLogo 
      Height          =   5010
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   1275
      Begin VB.Image imgKala 
         Height          =   4605
         Left            =   135
         Picture         =   "frmBajas.frx":0000
         Stretch         =   -1  'True
         Top             =   225
         Width           =   975
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
      Height          =   1140
      Left            =   1380
      TabIndex        =   0
      Top             =   4320
      Width           =   7215
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
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   1605
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
         Left            =   105
         TabIndex        =   5
         Top             =   525
         Width           =   1605
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
         Left            =   4875
         TabIndex        =   4
         Top             =   525
         Width           =   855
      End
      Begin VB.Label lblTrabajo 
         Caption         =   "DESARROLLO"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   255
         Width           =   5160
      End
      Begin VB.Label lblTrabajo 
         Caption         =   "PROGRAMADOR"
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
         Left            =   1800
         TabIndex        =   2
         Top             =   495
         Width           =   3330
      End
      Begin VB.Label lblTrabajo 
         Caption         =   "ÚNICO"
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
         Left            =   5820
         TabIndex        =   1
         Top             =   495
         Width           =   1230
      End
   End
   Begin VB.Image imgFoto 
      Height          =   1455
      Left            =   1365
      Picture         =   "frmBajas.frx":1DC5
      Stretch         =   -1  'True
      Top             =   15
      Width           =   1305
   End
   Begin VB.Label lblNombre 
      Caption         =   "HERNÁNDEZ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Index           =   0
      Left            =   2775
      TabIndex        =   29
      Top             =   0
      Width           =   4755
   End
   Begin VB.Label lblNombre 
      Caption         =   "MARTÍNEZ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Index           =   1
      Left            =   2775
      TabIndex        =   28
      Top             =   457
      Width           =   4755
   End
   Begin VB.Label lblNombre 
      Caption         =   "DAVID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Index           =   2
      Left            =   2775
      TabIndex        =   27
      Top             =   914
      Width           =   4755
   End
End
Attribute VB_Name = "frmBajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
