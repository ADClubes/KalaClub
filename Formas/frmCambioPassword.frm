VERSION 5.00
Begin VB.Form frmCambioPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Contraseña"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5310
   Icon            =   "frmCambioPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5310
   Begin VB.TextBox txtChPsswrd 
      Alignment       =   2  'Center
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2715
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1440
      Width           =   1845
   End
   Begin VB.TextBox txtChPsswrd 
      Alignment       =   2  'Center
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2040
      Width           =   1845
   End
   Begin VB.TextBox txtChPsswrd 
      Alignment       =   2  'Center
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2700
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   1845
   End
   Begin VB.ComboBox cboLogin 
      Height          =   315
      Left            =   2760
      TabIndex        =   1
      Top             =   525
      Width           =   1875
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   1080
      Picture         =   "frmCambioPassword.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   2760
      Width           =   795
   End
   Begin VB.CommandButton cmdGuardar 
      Height          =   840
      Left            =   3600
      Picture         =   "frmCambioPassword.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guardar"
      Top             =   2640
      Width           =   795
   End
   Begin VB.Label lblChPsswrd 
      BackStyle       =   0  'Transparent
      Caption         =   "&Confirme Nueva Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   720
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblChPsswrd 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nueva Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   1560
      Width           =   1650
   End
   Begin VB.Label lblChPsswrd 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña &Actual:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1650
   End
   Begin VB.Label lblChPsswrd 
      BackStyle       =   0  'Transparent
      Caption         =   "&Login Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1515
   End
End
Attribute VB_Name = "frmCambioPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA CAMBIO DE CONTRASEÑAS
' Objetivo: PERMITE MODIFICAR LAS CONTRASEÑAS
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim intEntos As Integer
    Dim AdoChPsw As ADODB.Recordset
        
Private Function VerificaDatos()
    Dim i As Integer

    For i = 0 To 2      'Verificamos que no estén vacíos
        If (txtChPsswrd(i).Enabled = True) And (txtChPsswrd(i).Text = vbNullString) Then
            MsgBox "¡ Favor de Llenar Todas las Casillas, Pues No son Opcionales !", _
                        vbOKOnly + vbExclamation, "Cambio de Contraseña"
            VerificaDatos = False
            txtChPsswrd(i).SetFocus
            Exit Function
        End If
    Next i
    
    If txtChPsswrd(0).Enabled = True Then
        'Verificamos la Contraseña Actual
        strSQL = "SELECT upassword FROM usuarios_sistema WHERE idusuario = " & _
                        cboLogin.ItemData(cboLogin.ListIndex)
        Set AdoChPsw = New ADODB.Recordset
        AdoChPsw.ActiveConnection = Conn
        AdoChPsw.LockType = adLockOptimistic
        AdoChPsw.CursorType = adOpenKeyset
        AdoChPsw.CursorLocation = adUseServer
        AdoChPsw.Open strSQL
        If Not AdoChPsw.EOF Then
            If AdoChPsw!upassword <> txtChPsswrd(0).Text Then
                intEntos = intEntos + 1
                If intEntos = 3 Then
                    MsgBox "¡ Lo Siento Amigo ! ¡ NO Lo Conzco !", _
                            vbOKOnly + vbCritical, "Cambio de Contraseña"
                    Unload Me
                    Exit Function
                End If
                MsgBox "¡ La Contraseña Actual Ingresada NO es Correcta !", _
                vbOKOnly + vbExclamation, "Cambio de Contraseña"
                VerificaDatos = False
                txtChPsswrd(0).SetFocus
                Exit Function
            End If
        End If
    End If
    
    'Verificamos la Confirmación de la Nueva Contrseña
    If txtChPsswrd(2).Text <> txtChPsswrd(1).Text Then
        MsgBox "¡ La Confirmación de la Contraseña NO es Correcta !", _
                    vbOKOnly + vbExclamation, "Cambio de Contraseña"
        VerificaDatos = False
        txtChPsswrd(2).SetFocus
        Exit Function
    End If
    
    'Verificamos que la contraseña anterior y la nueva no sean iguales
    
    If txtChPsswrd(1).Text <> txtChPsswrd(2).Text Then
        MsgBox "¡ La nueva contraseña NO puede ser igual a la anterior !", _
                    vbOKOnly + vbExclamation, "Cambio de Contraseña"
        VerificaDatos = False
        txtChPsswrd(1).SetFocus
        Exit Function
    End If
    
    
    VerificaDatos = True
End Function

Private Sub cboLogin_Click()
    Select Case sDB_NivelUser
        Case 0
'            If cboLogin.Text = "ADMIN" Then
'                lblChPsswrd(1).Enabled = True
'                txtChPsswrd(0).Enabled = True
'            Else
                lblChPsswrd(1).Enabled = False
                txtChPsswrd(0).Enabled = False
'            End If
        Case Else
            lblChPsswrd(1).Enabled = True
            txtChPsswrd(0).Enabled = True
    End Select
End Sub

Private Sub cmdGuardar_Click()
    If VerificaDatos = True Then
        Call RemplazaDatos
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strCampo1, strCampo2 As String
    intEntos = 0
    frmCatalogos.Enabled = False
    If (sDB_NivelUser = 0) Then
        lblChPsswrd(1).Enabled = False
        txtChPsswrd(0).Enabled = False
        strSQL = "SELECT idusuario, login_name FROM usuarios_sistema ORDER BY login_name"
    Else
        strSQL = "SELECT idusuario, login_name FROM usuarios_sistema " & _
        "WHERE login_name = '" & MDIPrincipal.StatusBar1.Panels.Item(5).Text & "'"
    End If
    strCampo1 = "login_name"
    strCampo2 = "idusuario"
    Call LlenaCombos(cboLogin, strSQL, strCampo1, strCampo2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
        frmCatalogos.Enabled = True
End Sub

Private Sub RemplazaDatos()
   Dim AdoCmdRemplaza As ADODB.Command
    On Error GoTo err_Guarda
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    'Actualizamos la nueva Contraseña
    strSQL = "UPDATE usuarios_sistema SET upassword = '" & Trim(txtChPsswrd(1).Text) & _
                    "' WHERE idusuario = " & cboLogin.ItemData(cboLogin.ListIndex)
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    MsgBox "¡ Contraseña Actualizada !"
    Unload Me
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub txtchpsswrd_GotFocus(Index As Integer)
    txtChPsswrd(Index).SelStart = 0
    txtChPsswrd(Index).SelLength = Len(txtChPsswrd(Index))
End Sub
