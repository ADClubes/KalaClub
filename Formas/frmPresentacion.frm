VERSION 5.00
Begin VB.Form frmPresentacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3750
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   5400
   ControlBox      =   0   'False
   Icon            =   "frmPresentacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDatos 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5055
      Begin VB.Label lblDireccion 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label lblNombreClub 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4815
      End
   End
   Begin VB.Timer tmrTimeOut 
      Interval        =   65535
      Left            =   390
      Top             =   3180
   End
   Begin VB.TextBox txtPsw 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1920
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3180
      Width           =   2040
   End
   Begin VB.CommandButton cmdCancelar 
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
      Height          =   465
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3090
      Width           =   1050
   End
   Begin VB.TextBox txtLogin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   1905
      MaxLength       =   15
      TabIndex        =   0
      Top             =   2595
      Width           =   2040
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2490
      Width           =   1050
   End
   Begin VB.Image imgLogo 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "&Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   870
      TabIndex        =   5
      Top             =   2685
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "C&ontraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   390
      TabIndex        =   4
      Top             =   3285
      Width           =   1485
   End
End
Attribute VB_Name = "frmPresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA DE PRESENTACIÓN
' Objetivo: MUESTRA UNA PANTALLA CON EL LOGO DE LA EMPRESA
' Programado por:
' Fecha: OCTUBRE DE 2003
' Modificado: DICIEMBRE DE 2004
' ************************************************************************
Option Explicit
Dim intContador As Integer 'Variable que cuenta el número de intentos
Const SM_CLEANBOOT& = 67

Private Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long

Private Sub cmdAceptar_Click()
    Dim AdoRcsPresenta As ADODB.Recordset
    
    Dim sHora As String
    sDB_User = Trim(txtLogin.Text)
    sDB_PW = Trim(txtPsw.Text)
    LoginOk = False
    
    strSQL = "SELECT IdUsuario, uPassword, Nombre, IdPerfil, FechaVencePass"
    strSQL = strSQL & " FROM USUARIOS_SISTEMA"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " Login_Name = '" & Trim(txtLogin.Text) & "'"
    strSQL = strSQL & " AND Login_Name IS NOT NULL"
    strSQL = strSQL & " AND Status ='A'"
    
    Set AdoRcsPresenta = New ADODB.Recordset
    
    AdoRcsPresenta.ActiveConnection = Conn
    AdoRcsPresenta.LockType = adLockReadOnly
    AdoRcsPresenta.CursorType = adOpenForwardOnly
    AdoRcsPresenta.CursorLocation = adUseServer
    AdoRcsPresenta.Open strSQL
    
    If Not AdoRcsPresenta.EOF Then
        AdoRcsPresenta.MoveFirst
        Do While Not AdoRcsPresenta.EOF
            If AdoRcsPresenta.Fields("upassword") = sDB_PW Then
                
                '20/06/08
                'Valida la fecha de vencimiento de la contraseña
                If AdoRcsPresenta!FechaVencePass < Date Then
                    MsgBox "Contraseña caducada", vbCritical, "Error"
                    LoginOk = False
                    ChangePassword = True
                    
                    Unload Me
                    Exit Sub
                End If
                
                ' Carga variables globales
                sHora = Hour(Time) & Minute(Time) & Second(Time)
                sDB_IdUser = Trim(AdoRcsPresenta.Fields("idusuario") & sHora)
                sDB_NivelUser = AdoRcsPresenta.Fields("IdPerfil")
                iDB_IdUser = AdoRcsPresenta.Fields("IdUsuario")
                LoginOk = True
                Exit Do
            End If
            AdoRcsPresenta.MoveNext
        Loop
        If Not LoginOk Then
            intContador = intContador + 1
            If intContador > 2 Then
                MsgBox "LO SIENTO, ¡NO LO CONOZCO!", vbCritical, "Error"
                LoginOk = False
                Unload Me
                Exit Sub
            End If
            MsgBox "¡Usuario o contraseña inválidos!", vbExclamation, "Error"
            Me.txtPsw.SetFocus
            DoEvents
            Exit Sub
        End If
    Else
        LoginOk = False
        intContador = intContador + 1
        If intContador > 2 Then
            MsgBox "LO SIENTO, ¡NO LO CONOZCO!", vbCritical, "Error"
            Unload Me
            Exit Sub
        End If
        MsgBox "¡Usuario o contraseña inválidos!", vbExclamation, "Error"
        txtLogin.SetFocus
        Exit Sub
    End If
    
    AdoRcsPresenta.Close
    Set AdoRcsPresenta = Nothing
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    LoginOk = False
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim sNombreClub As String
    Dim sDireccion As String
    
    Dim sPathLogo As String
    
    sPathLogo = sDB_DataSource & "\rc\logo.jpg"
    
    If sNombreClub = vbNullString Then
        sNombreClub = ObtieneParametro("NOMBRE DEL CLUB")
    End If
    
    If sDireccion = vbNullString Then
        sDireccion = ObtieneParametro("DIRECCION DEL CLUB")
    End If
    
    On Error Resume Next
    If Dir(sPathLogo) <> "" Then
        On Error Resume Next
        Me.imgLogo.Picture = LoadPicture(sPathLogo)
        On Error GoTo 0
    End If
    
    Me.lblNombreClub.Caption = sNombreClub
    Me.lblDireccion.Caption = sDireccion
    
    Me.txtLogin.Text = sDB_User
    Me.txtPsw.SetFocus
End Sub

Private Sub Form_Load()
    Dim result As Long
    
    result = GetSystemMetrics(SM_CLEANBOOT)
    
    Select Case result
        Case 0
            'OK
'            MsgBox "System started in normal mode."
        Case 1
            MsgBox "Sistema iniciado en modo seguro."
            End
        Case 2
            MsgBox "Sistema iniciado en modo seguro con funciones de red."
            End
        Case Else
            MsgBox "Inicio de sistema con valor desconocido."
            End
    End Select
    
    With frmPresentacion
        .MousePointer = 0
        .Width = 5490
        .Height = 3840
        .Left = (Screen.Width / 2) - (frmPresentacion.Width / 2)
        .Top = (Screen.Height / 2) - (frmPresentacion.Height / 2)
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrTimeOut.Interval = 0
End Sub

Private Sub Image2_DblClick()
'    txtLogin.Text = "ADMIN"
'    txtPsw.Text = "poderoso"
'    cmdAceptar.Value = True
End Sub

Private Sub tmrTimeOut_Timer()
    Me.cmdCancelar.Value = True
End Sub

Private Sub txtLogin_GotFocus()
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(Me.txtLogin.Text)
    DoEvents
End Sub

Private Sub txtPsw_GotFocus()
    Me.txtPsw.SelStart = 0
    Me.txtPsw.SelLength = Len(Me.txtPsw.Text)
    DoEvents
End Sub
