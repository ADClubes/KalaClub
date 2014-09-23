VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUsuariosSistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios del Sistema (Captura)"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6885
   Icon            =   "frmUsuariosSistema.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6885
   Begin VB.TextBox txtUsuariosSistema 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   360
      TabIndex        =   13
      Top             =   3480
      Width           =   1080
   End
   Begin VB.TextBox txtUsuariosSistema 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   345
      TabIndex        =   11
      Top             =   2520
      Width           =   5160
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   330
      Left            =   4455
      TabIndex        =   9
      Top             =   1515
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      Format          =   48758785
      CurrentDate     =   37974
   End
   Begin VB.TextBox txtUsuariosSistema 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   375
      TabIndex        =   1
      Top             =   510
      Width           =   1080
   End
   Begin VB.TextBox txtUsuariosSistema 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   1
      Left            =   1830
      MaxLength       =   15
      TabIndex        =   3
      Top             =   510
      Width           =   1980
   End
   Begin VB.TextBox txtUsuariosSistema 
      Alignment       =   2  'Center
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   345
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1515
      Width           =   1560
   End
   Begin VB.TextBox txtUsuariosSistema 
      Alignment       =   2  'Center
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   2385
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1515
      Width           =   1560
   End
   Begin VB.CommandButton cmdGuardar 
      Height          =   840
      Left            =   5940
      Picture         =   "frmUsuariosSistema.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Guardar"
      Top             =   555
      Width           =   795
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   5940
      Picture         =   "frmUsuariosSistema.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Salir"
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label lblUsuarioSistema 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Nivel:"
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
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblUsuarioSistema 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   270
      Index           =   5
      Left            =   4410
      TabIndex        =   8
      Top             =   1215
      Width           =   660
   End
   Begin VB.Label lblUsuarioSistema 
      BackStyle       =   0  'Transparent
      Caption         =   "Id &Usuario:"
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
      Index           =   0
      Left            =   375
      TabIndex        =   0
      Top             =   210
      Width           =   1080
   End
   Begin VB.Label lblUsuarioSistema 
      BackStyle       =   0  'Transparent
      Caption         =   "&Login: (Máx: 15)"
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
      Index           =   1
      Left            =   1830
      TabIndex        =   2
      Top             =   210
      Width           =   1935
   End
   Begin VB.Label lblUsuarioSistema 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C&ontraseña:(Máx: 10)"
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
      Index           =   2
      Left            =   345
      TabIndex        =   4
      Top             =   1215
      Width           =   1830
   End
   Begin VB.Label lblUsuarioSistema 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Conf&irme Contraseña:"
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
      Index           =   3
      Left            =   2385
      TabIndex        =   6
      Top             =   1215
      Width           =   1860
   End
   Begin VB.Label lblUsuarioSistema 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Nombre:"
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
      Index           =   4
      Left            =   345
      TabIndex        =   10
      Top             =   2220
      Width           =   735
   End
   Begin VB.Label lblClave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Left            =   345
      TabIndex        =   12
      Top             =   510
      Width           =   1110
   End
End
Attribute VB_Name = "frmUsuariosSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA: USUARIOS DEL SISTEMA
' Objetivo: CATÁLOGO DE USUARIOS DEL SISTEMA
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsUsuariosSistema As ADODB.Recordset
    
Private Function VerificaDatos()
    Dim i, intInicio As Integer
    If txtUsuariosSistema(0).Visible = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    For i = intInicio To 3
        Select Case i
            Case 0
                If (txtUsuariosSistema(i).Text = "") Then
                    MsgBox "¡ Favor de Llenar la Casilla CLAVE, Pues No es Opcional !", _
                                 vbOKOnly + vbExclamation, "Usuarios Sistema (Captura)"
                    VerificaDatos = False
                    txtUsuariosSistema(i).SetFocus
                    Exit Function
                End If
            Case 1
                If txtUsuariosSistema(i).Text = "" Then
                    MsgBox "¡ Favor de Llenar la Casilla LOGIN NAME, Pues No es Opcional.", _
                                vbOKOnly + vbExclamation, "Usuarios Sistema (Captura)"
                    VerificaDatos = False
                    txtUsuariosSistema(i).SetFocus
                    Exit Function
                End If
            Case 2
                If (frmCatalogos.lblModo.Caption = "A") And (Trim(txtUsuariosSistema(i)) = "") Then
                    MsgBox "¡ Favor de Llenar la Casilla CONTRASEÑA, Pues No es Opcional.", _
                                vbOKOnly + vbExclamation, "Usuarios Sistema (Captura)"
                    VerificaDatos = False
                    txtUsuariosSistema(i).SetFocus
                    Exit Function
                End If
            Case 3
                If (frmCatalogos.lblModo.Caption = "A") And _
                    Trim(txtUsuariosSistema(i)) <> Trim(txtUsuariosSistema(2)) Then
                    MsgBox "¡ La Confirmación de la CONTRASEÑA no coincide !", _
                                vbOKOnly + vbExclamation, "Usuarios Sistema (Captura)"
                    VerificaDatos = False
                    txtUsuariosSistema(i).SetFocus
                    Exit Function
                End If
        End Select
        VerificaDatos = True
    Next
    
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT idusuario FROM usuarios_sistema WHERE idusuario = " & _
                        Val(Trim(txtUsuariosSistema(0).Text))
        Set AdoRcsUsuariosSistema = New ADODB.Recordset
        AdoRcsUsuariosSistema.ActiveConnection = Conn
        AdoRcsUsuariosSistema.LockType = adLockOptimistic
        AdoRcsUsuariosSistema.CursorType = adOpenKeyset
        AdoRcsUsuariosSistema.CursorLocation = adUseServer
        AdoRcsUsuariosSistema.Open strSQL
        If Not AdoRcsUsuariosSistema.EOF Then
            MsgBox "Ya Existe Un Registro Con La Clave: " & _
                        txtUsuariosSistema(0).Text, vbInformation + vbOKOnly, "Usuarios Sistema"
            AdoRcsUsuariosSistema.Close
            VerificaDatos = False
            txtUsuariosSistema(0).SetFocus
            Exit Function
        Else
            AdoRcsUsuariosSistema.Close
            VerificaDatos = True
        End If
    End If
    
    
    
    'Se verifica la fecha de caducidad
    If DateDiff("d", Date, Me.dtpFecha.Value) < 1 Then
        MsgBox "La fecha de caducidad de la clabe deber ser al menos un dia", vbExclamation, "Verifique"
        VerificaDatos = False
        Me.dtpFecha.SetFocus
        Exit Function
    End If
    
    VerificaDatos = True
    
End Function

Private Sub cmdGuardar_Click()
    If VerificaDatos = True Then
        If frmCatalogos.lblModo.Caption = "A" Then
            Call GuardaDatos
        Else
            Call RemplazaDatos
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmCatalogos.Enabled = False
    dtpFecha.Value = Now
    If frmCatalogos.lblModo.Caption = "A" Then
        txtUsuariosSistema(0).Visible = True
        Call Llena_txtUsuariosSistema
    Else
        txtUsuariosSistema(0).Visible = False
        lblUsuarioSistema(2).Enabled = False
        lblUsuarioSistema(3).Enabled = False
        txtUsuariosSistema(2).Enabled = False
        txtUsuariosSistema(3).Enabled = False
        Call LlenaDatos
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        frmCatalogos.Enabled = True
End Sub

Private Sub GuardaDatos()
    Dim AdoCmdInserta As ADODB.Command
    On Error GoTo err_Guarda
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO usuarios_sistema ("
        strSQL = strSQL & " idusuario,"
        strSQL = strSQL & " login_name,"
        strSQL = strSQL & " upassword,"
        strSQL = strSQL & " nombre,"
        strSQL = strSQL & " fecha_alta,"
        strSQL = strSQL & " idPerfil,"
        strSQL = strSQL & " FechaVencePass)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & Val(Trim(txtUsuariosSistema(0).Text)) & ","
        strSQL = strSQL & "'" & Trim(txtUsuariosSistema(1).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtUsuariosSistema(2).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtUsuariosSistema(4).Text) & "',"
        strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "',"
        strSQL = strSQL & Val(Trim(txtUsuariosSistema(5).Text)) & ","
        strSQL = strSQL & "'" & Format(Me.dtpFecha.Value, "yyyymmdd") & "')"
    #Else
        strSQL = "INSERT INTO usuarios_sistema ("
        strSQL = strSQL & " idusuario,"
        strSQL = strSQL & " login_name,"
        strSQL = strSQL & " upassword,"
        strSQL = strSQL & " nombre,"
        strSQL = strSQL & " fecha_alta,"
        strSQL = strSQL & " idPerfil,"
        strSQL = strSQL & " FechaVencePass)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & Val(Trim(txtUsuariosSistema(0).Text)) & ","
        strSQL = strSQL & "'" & Trim(txtUsuariosSistema(1).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtUsuariosSistema(2).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtUsuariosSistema(4).Text) & "',"
        strSQL = strSQL & "#" & Format(Date, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & Val(Trim(txtUsuariosSistema(5).Text)) & ","
        strSQL = strSQL & "#" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
    #End If
    
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    
    
    'Borra la seguridad del usuario
    #If SqlServer_ Then
        strSQL = "DELETE"
        strSQL = strSQL & " FROM SEGURIDAD_USUARIO"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " SEGURIDAD_USUARIO.IdUsuario=" & Val(Trim(txtUsuariosSistema(0).Text))
    #Else
        strSQL = ""
        strSQL = strSQL & "DELETE SEGURIDAD_USUARIO.IdUsuario"
        strSQL = strSQL & " FROM SEGURIDAD_USUARIO"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " (SEGURIDAD_USUARIO.IdUsuario)=" & Val(Trim(txtUsuariosSistema(0).Text))
        strSQL = strSQL & ")"
    #End If
    
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    
    'Inserta la seguridad del usuario
    strSQL = ""
    strSQL = strSQL & "INSERT INTO SEGURIDAD_USUARIO"
    strSQL = strSQL & " SELECT " & Val(Trim(txtUsuariosSistema(0).Text)) & " AS IdUsuario, SEGURIDAD_CT_PERFILES_DETALLE.IdObjeto AS IdObjeto"
    strSQL = strSQL & " From SEGURIDAD_CT_PERFILES_DETALLE"
    strSQL = strSQL & " WHERE ((IdPerfil)=" & Trim(txtUsuariosSistema(5).Text) & ")"
    
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    MsgBox "¡Usuario Ingresado!"
    Call Limpia
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub RemplazaDatos()
   Dim AdoCmdRemplaza As ADODB.Command
    On Error GoTo err_Guarda
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    'Eliminamos elregistro existente
    #If SqlServer_ Then
        strSQL = "UPDATE usuarios_sistema"
        strSQL = strSQL & " SET login_name = '" & Trim(txtUsuariosSistema(1).Text) & "',"
        strSQL = strSQL & " FechaVencePass = '" & Format(dtpFecha.Value, "yyyymmdd") & "',"
        strSQL = strSQL & " nombre = '" & Trim(txtUsuariosSistema(4).Text) & "',"
        strSQL = strSQL & " idPerfil=" & Val(Trim(Me.txtUsuariosSistema(5).Text))
        strSQL = strSQL & " WHERE idusuario = " & Val(lblClave.Caption)
    #Else
        strSQL = "UPDATE usuarios_sistema"
        strSQL = strSQL & " SET login_name = '" & Trim(txtUsuariosSistema(1).Text) & "',"
        strSQL = strSQL & " FechaVencePass = #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & " nombre = '" & Trim(txtUsuariosSistema(4).Text) & "',"
        strSQL = strSQL & " idPerfil=" & Val(Trim(Me.txtUsuariosSistema(5).Text))
        strSQL = strSQL & " WHERE idusuario = " & Val(lblClave.Caption)
    #End If
    
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Borra la seguridad del usuario
    #If SqlServer_ Then
        strSQL = "DELETE"
        strSQL = strSQL & " FROM SEGURIDAD_USUARIO"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " SEGURIDAD_USUARIO.IdUsuario=" & Val(lblClave.Caption)
    #Else
        strSQL = ""
        strSQL = strSQL & "DELETE SEGURIDAD_USUARIO.IdUsuario"
        strSQL = strSQL & " FROM SEGURIDAD_USUARIO"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " (SEGURIDAD_USUARIO.IdUsuario)=" & Val(lblClave.Caption)
        strSQL = strSQL & ")"
    #End If
    
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Inserta la seguridad del usuario
    strSQL = ""
    strSQL = strSQL & "INSERT INTO SEGURIDAD_USUARIO"
    strSQL = strSQL & " SELECT " & Trim(Me.lblClave.Caption) & " AS IdUsuario, SEGURIDAD_CT_PERFILES_DETALLE.IdObjeto AS IdObjeto"
    strSQL = strSQL & " From SEGURIDAD_CT_PERFILES_DETALLE"
    strSQL = strSQL & " WHERE ((IdPerfil)=" & Trim(txtUsuariosSistema(5).Text) & ")"
    
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    
    
    
    Conn.CommitTrans                            'Termina transacción
    Screen.MousePointer = vbDefault
    
    MsgBox "¡Registro Actualizado!"
    Unload Me
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub txtUsuariosSistema_GotFocus(Index As Integer)
    txtUsuariosSistema(Index).SelStart = 0
    txtUsuariosSistema(Index).SelLength = Len(txtUsuariosSistema(Index))
End Sub

Private Sub txtUsuariosSistema_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            Select Case KeyAscii
                Case 8, 22, 48 To 57            'Backspace, <Ctrl+V> y del 0 al 9
                    KeyAscii = KeyAscii
                Case Else
                    KeyAscii = 0
                End Select
        Case 1, 2, 3
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Sub Limpia()
    Dim i As Integer
    For i = 0 To 5
        txtUsuariosSistema(i).Text = ""
    Next i
    lblClave.Caption = ""
    Call Llena_txtUsuariosSistema
    txtUsuariosSistema(0).SetFocus
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM usuarios_sistema WHERE idusuario = " & _
                    Val(Trim(frmCatalogos.lblModo.Caption))
    Set AdoRcsUsuariosSistema = New ADODB.Recordset
    AdoRcsUsuariosSistema.ActiveConnection = Conn
    AdoRcsUsuariosSistema.LockType = adLockOptimistic
    AdoRcsUsuariosSistema.CursorType = adOpenKeyset
    AdoRcsUsuariosSistema.CursorLocation = adUseServer
    AdoRcsUsuariosSistema.Open strSQL
    If Not AdoRcsUsuariosSistema.EOF Then
        lblClave.Caption = AdoRcsUsuariosSistema!IdUsuario
        txtUsuariosSistema(1).Text = AdoRcsUsuariosSistema!Login_Name
        txtUsuariosSistema(4).Text = AdoRcsUsuariosSistema!Nombre
        txtUsuariosSistema(5).Text = AdoRcsUsuariosSistema!IdPerfil
        dtpFecha.Value = Format(IIf(IsNull(AdoRcsUsuariosSistema!FechaVencePass), Date, AdoRcsUsuariosSistema!FechaVencePass), "dd/mm/yyyy")
    End If
End Sub


Sub Llena_txtUsuariosSistema()
    Dim lngAnterior, lngUsuariosSistema As Long
    'Llena txtUsuariosSistema
    lngUsuariosSistema = 1
    strSQL = "SELECT idusuario FROM usuarios_Sistema ORDER BY idusuario"
    Set AdoRcsUsuariosSistema = New ADODB.Recordset
    AdoRcsUsuariosSistema.ActiveConnection = Conn
    AdoRcsUsuariosSistema.LockType = adLockOptimistic
    AdoRcsUsuariosSistema.CursorType = adOpenKeyset
    AdoRcsUsuariosSistema.CursorLocation = adUseServer
    AdoRcsUsuariosSistema.Open strSQL
    If AdoRcsUsuariosSistema.EOF Then
        lngUsuariosSistema = 1
        txtUsuariosSistema(0).Text = lngUsuariosSistema
        Exit Sub
    End If
    AdoRcsUsuariosSistema.MoveFirst
    Do While Not AdoRcsUsuariosSistema.EOF
        If AdoRcsUsuariosSistema.Fields!IdUsuario <> "1" Then
            If Val(AdoRcsUsuariosSistema.Fields!IdUsuario) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngUsuariosSistema
        AdoRcsUsuariosSistema.MoveNext
        If Not AdoRcsUsuariosSistema.EOF Then lngUsuariosSistema = AdoRcsUsuariosSistema.Fields!IdUsuario
    Loop
    txtUsuariosSistema(0).Text = lngAnterior + 1
End Sub
