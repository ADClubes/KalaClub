VERSION 5.00
Begin VB.Form frmBancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bancos (Captura)"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5205
   Icon            =   "frmBancos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5205
   Begin VB.TextBox txtBancos 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   1515
      TabIndex        =   1
      Top             =   630
      Width           =   1215
   End
   Begin VB.TextBox txtBancos 
      Height          =   330
      Index           =   1
      Left            =   1515
      TabIndex        =   3
      Top             =   1080
      Width           =   3435
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      Height          =   840
      Left            =   2955
      Picture         =   "frmBancos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar"
      Top             =   1515
      Width           =   795
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   4125
      Picture         =   "frmBancos.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   1515
      Width           =   795
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BANCOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   742
      TabIndex        =   7
      Top             =   90
      Width           =   3720
   End
   Begin VB.Label lblClave 
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
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   675
      Width           =   1095
   End
   Begin VB.Label lblBancos 
      BackStyle       =   0  'Transparent
      Caption         =   "&Clave:"
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
      Left            =   150
      TabIndex        =   0
      Top             =   750
      Width           =   1230
   End
   Begin VB.Label lblBancos 
      BackStyle       =   0  'Transparent
      Caption         =   "&Descripción:"
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
      Left            =   150
      TabIndex        =   2
      Top             =   1185
      Width           =   1230
   End
End
Attribute VB_Name = "frmBancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA BANCOS
' Objetivo: CATÁLOGO DE BANCOS
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsBancos As ADODB.Recordset
        
Private Function VerificaDatos()
    Dim i, intInicio As Integer
    If txtBancos(0).Visible = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    For i = intInicio To 1
        If (txtBancos(i).Text = "") Then
            MsgBox "¡ Favor de Llenar Todas las Casillas, Pues No son Opcionales !", _
                        vbOKOnly + vbExclamation, "Bancos (Captura)"
            VerificaDatos = False
            txtBancos(i).SetFocus
            Exit Function
        End If
        VerificaDatos = True
    Next i
    
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT idbanco FROM bancos WHERE idbanco = " & _
                        Val(Trim(txtBancos(0).Text))
        Set AdoRcsBancos = New ADODB.Recordset
        AdoRcsBancos.ActiveConnection = Conn
        AdoRcsBancos.LockType = adLockOptimistic
        AdoRcsBancos.CursorType = adOpenKeyset
        AdoRcsBancos.CursorLocation = adUseServer
        AdoRcsBancos.Open strSQL
        If Not AdoRcsBancos.EOF Then
            MsgBox "Ya Existe Un Registro Con La Clave: " & _
                        txtBancos(0).Text, vbInformation + vbOKOnly, "Bancos"
            AdoRcsBancos.Close
            VerificaDatos = False
            txtBancos(0).SetFocus
            Exit Function
        Else
            AdoRcsBancos.Close
            VerificaDatos = True
        End If
    End If
End Function

Private Sub cmdGuardar_Click()
    Dim blnGuarda As Boolean
    blnGuarda = VerificaDatos
    If blnGuarda = True Then
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
    If frmCatalogos.lblModo.Caption = "A" Then
        txtBancos(0).Visible = True
        Call Llena_txtBancos
    Else
        txtBancos(0).Visible = False
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
    strSQL = "INSERT INTO bancos (idbanco, banco) VALUES (" & _
                 Val(Trim(txtBancos(0))) & ", '" & Trim(txtBancos(1)) & "')"
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    MsgBox "¡Registro Ingresado!"
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
    strSQL = "DELETE FROM bancos WHERE idbanco = " & Val(lblClave.Caption)
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Ahora Insertamos el nuevo registro que sustituye al anterior
    strSQL = "INSERT INTO bancos (idbanco, banco) VALUES (" & _
                    Val(Trim(lblClave.Caption)) & ", '" & Trim(txtBancos(1)) & "')"
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    Conn.CommitTrans      'Termina transacción
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

Private Sub txtBancos_GotFocus(Index As Integer)
    txtBancos(Index).SelStart = 0
    txtBancos(Index).SelLength = Len(txtBancos(Index))
End Sub

Private Sub txtBancos_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0
            Select Case KeyAscii
                Case 8, 22, 48 To 57 'Backspace, <Ctrl+V> y del 0 al 9
                    KeyAscii = KeyAscii
                Case Else
                    KeyAscii = 0
                End Select
        Case 1
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Sub Limpia()
    Dim i As Integer
    For i = 0 To 1
        txtBancos(i).Text = ""
    Next i
    Call Llena_txtBancos
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM bancos WHERE idbanco = " & _
                  Val(Trim(frmCatalogos.lblModo.Caption))
    Set AdoRcsBancos = New ADODB.Recordset
    AdoRcsBancos.ActiveConnection = Conn
    AdoRcsBancos.LockType = adLockOptimistic
    AdoRcsBancos.CursorType = adOpenKeyset
    AdoRcsBancos.CursorLocation = adUseServer
    AdoRcsBancos.Open strSQL
    If Not AdoRcsBancos.EOF Then
        lblClave.Caption = AdoRcsBancos!idbanco
        txtBancos(1).Text = AdoRcsBancos!banco
    End If
End Sub

Sub Llena_txtBancos()
    Dim lngAnterior, lngBancos As Long
    'Llena txtbancos
    lngBancos = 1
    strSQL = "SELECT idbanco FROM bancos ORDER BY idbanco"
    Set AdoRcsBancos = New ADODB.Recordset
    AdoRcsBancos.ActiveConnection = Conn
    AdoRcsBancos.LockType = adLockOptimistic
    AdoRcsBancos.CursorType = adOpenKeyset
    AdoRcsBancos.CursorLocation = adUseServer
    AdoRcsBancos.Open strSQL
    If AdoRcsBancos.EOF Then
        lngBancos = 1
        txtBancos(0).Text = lngBancos
        Exit Sub
    End If
    AdoRcsBancos.MoveFirst
    Do While Not AdoRcsBancos.EOF
        If AdoRcsBancos.Fields!idbanco <> "1" Then
            If Val(AdoRcsBancos.Fields!idbanco) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngBancos
        AdoRcsBancos.MoveNext
        If Not AdoRcsBancos.EOF Then lngBancos = AdoRcsBancos.Fields!idbanco
    Loop
    txtBancos(0).Text = lngAnterior + 1
End Sub
