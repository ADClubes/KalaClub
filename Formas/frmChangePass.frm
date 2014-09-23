VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de contraseña"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtControl 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2160
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txtControl 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2160
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Debe contener por lo menos 1 letra mayuscula, 1 numero y tener una longitud de 6 a 12 caracteres"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtControl 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2160
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtControl 
      Height          =   375
      Index           =   0
      Left            =   2160
      MaxLength       =   15
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirme Contraseña Nueva"
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Contraseña Nueva"
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Contraseña actual"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    If Me.txtControl(0).Text = vbNullString Then
        MsgBox "Indicar un usuario", vbCritical, "Verifique"
        Me.txtControl(0).SetFocus
        Exit Sub
    End If
    
    If Me.txtControl(1).Text = vbNullString Then
        MsgBox "Indicar la contraseña actual", vbCritical, "Verifique"
        Me.txtControl(1).SetFocus
        Exit Sub
    End If
    
    If Me.txtControl(2).Text = vbNullString Then
        MsgBox "Indicar la contraseña nueva", vbCritical, "Verifique"
        Me.txtControl(2).SetFocus
        Exit Sub
    End If
    
    If Me.txtControl(3).Text = vbNullString Then
        MsgBox "Indicar la confirmación de la contraseña", vbCritical, "Verifique"
        Me.txtControl(3).SetFocus
        Exit Sub
    End If
    
    If Me.txtControl(2).Text <> Me.txtControl(3).Text Then
        MsgBox "La nueva contraseña y su confirmación no coinciden", vbCritical, "Verifique"
        Me.txtControl(3).SetFocus
        Exit Sub
    End If
    
    If Not ChecaPassword(Me.txtControl(0).Text, Me.txtControl(1).Text) Then
        MsgBox "La contraseña o el usuario actual no son correctos!", vbCritical, "Verifique"
        Me.txtControl(0).SetFocus
        Exit Sub
    End If
    
    If ChecaPassword(Me.txtControl(0).Text, Me.txtControl(2).Text) Then
        MsgBox "La nueva contraseña debe ser distinta a la actual!", vbCritical, "Verifique"
        Me.txtControl(2).SetFocus
        Exit Sub
    End If
    
    If Not ValidaExpReg(Me.txtControl(2).Text, "^(?=.*\d)(?=.*[a-z])(?=.*[A-Z]).{6,12}") Then
        MsgBox "La nueva contraseña no cumple con las características necesarias", vbCritical, "Verifique"
        Me.txtControl(2).SetFocus
        Exit Sub
    End If
    
    If CambiaPassword(Me.txtControl(0).Text, Me.txtControl(2).Text) = 0 Then
        MsgBox "Contraseña cambiada", vbInformation, "Ok"
    Else
        MsgBox "Ocurrio un error al cambiar la contraseña", vbCritical, "Verifique"
        Me.txtControl(0).SetFocus
        Exit Sub
    End If

    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
End Sub

Private Sub txtControl_GotFocus(Index As Integer)
    Me.txtControl(Index).SelStart = 0
    Me.txtControl(Index).SelLength = Len(Me.txtControl(Index).Text)
End Sub
