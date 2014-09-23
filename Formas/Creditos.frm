VERSION 5.00
Begin VB.Form Creditos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Créditos"
   ClientHeight    =   4515
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   3975
   Icon            =   "Creditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4515
   ScaleWidth      =   3975
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3840
      Top             =   120
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   3720
      Width           =   3735
      Begin VB.CommandButton Command1 
         Caption         =   "&OK"
         Height          =   400
         Left            =   2700
         TabIndex        =   14
         Top             =   160
         Width           =   975
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   2
         Left            =   600
         Max             =   11
         TabIndex        =   12
         Top             =   240
         Value           =   1
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   3645
      TabIndex        =   0
      Top             =   105
      Width           =   3705
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   5760
         Left            =   240
         ScaleHeight     =   5760
         ScaleWidth      =   3255
         TabIndex        =   1
         Top             =   -2430
         Width           =   3255
         Begin VB.Frame Frame2 
            Height          =   30
            Left            =   120
            TabIndex        =   7
            Top             =   2520
            Width           =   3000
         End
         Begin VB.Frame Frame1 
            Height          =   30
            Left            =   120
            TabIndex        =   3
            Top             =   1440
            Width           =   3000
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   $"Creditos.frx":030A
            Height          =   1035
            Left            =   120
            TabIndex        =   10
            Top             =   3720
            Width           =   2895
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "KalaClub © 2003 -2007"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   3360
            Width           =   2895
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Copyright"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   2880
            Width           =   2775
         End
         Begin VB.Label Label4 
            Caption         =   "e-mail: kalasys@kalasystems.com.mx"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "http://www.kalasystems.com.mx"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   3015
         End
         Begin VB.Label Label2 
            Caption         =   "Este programa está protegido por las leyes de Derechos de Autor. 2007. "
            Height          =   615
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Créditos."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   375
            Left            =   360
            TabIndex        =   2
            Top             =   240
            Width           =   2415
         End
      End
   End
End
Attribute VB_Name = "Creditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Picture1.Height = 3600
    Picture1.Width = 3700
    Picture2.Width = Picture1.Width - 400
    Picture2.Left = 200
    Picture2.Top = 3300
End Sub

Private Sub HScroll1_Change()
    Label8.Caption = HScroll1.Value
End Sub

Private Sub Timer1_Timer()
    If Picture2.Top < -4640 Then
        Picture2.Top = 3300
        HScroll1.Value = HScroll1.Value + 1
        If HScroll1.Value > 10 Then HScroll1.Value = 1
    End If
    Picture2.Top = Picture2.Top - 10 * HScroll1.Value
End Sub
