VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmActiva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activa por código"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.OptionButton optCtrl 
      Caption         =   "Inactiva"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton optCtrl 
      Caption         =   "Activa"
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   1
      Left            =   2520
      MaxLength       =   5
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   0
      Left            =   600
      MaxLength       =   5
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Codigo"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Prefijo"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmActiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEnviar_Click()
    
    Dim sCred As String
    Dim sTor As String
    Dim nErrCode As Long
    
    
    If Me.txtCtrl(0).Text = vbNullString Then
        Exit Sub
    End If
    
    If Me.txtCtrl(1).Text = vbNullString Then
        Exit Sub
    End If
    
    sCred = Format(Trim(Me.txtCtrl(0).Text), "00000") & Format(Trim(Me.txtCtrl(1).Text), "00000")
    sTor = (Me.txtCtrl(0).Text * 16777216) + Me.txtCtrl(1).Text
    
    If Me.optCtrl(0).Value Then
        #If SqlServer_ Then
            ActivaCred2SQL 1, sCred, 1, 0, True, True
        #Else
            ActivaCred2 1, sCred, 1, 0, True, True
        #End If
        'nErrCode = AgregaAccesoManual(Me.txtCtrl(0).Text & Me.txtCtrl(1).Text, sTor)
    Else
        #If SqlServer_ Then
            ActivaCred2SQL 1, sCred, 1, 0, False, True
        #Else
            ActivaCred2 1, sCred, 1, 0, False, True
        #End If
        'nErrCode = BloqueaAcceso(Me.txtCtrl(0).Text & Me.txtCtrl(1).Text)
    End If
    
'    If nErrCode <> 0 Then
'                MsgBox "No se pudo registrar el usuario en torniquetes,Favor de hacerlo manual"
'    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForma MDIPrincipal, Me
End Sub



Private Sub UpDown1_DownClick()
    Me.txtCtrl(1).Text = Val(Me.txtCtrl(1).Text) - 1
End Sub

Private Sub UpDown1_UpClick()
    Me.txtCtrl(1).Text = Val(Me.txtCtrl(1).Text) + 1
End Sub
