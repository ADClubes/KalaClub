VERSION 5.00
Begin VB.Form frmTipoClase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Clases (Captura)"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4965
   Icon            =   "frmTipoClase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4965
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   3870
      Picture         =   "frmTipoClase.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   1860
      Width           =   795
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      Height          =   840
      Left            =   2700
      Picture         =   "frmTipoClase.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar"
      Top             =   1860
      Width           =   795
   End
   Begin VB.TextBox txtTipoClase 
      Height          =   330
      Index           =   1
      Left            =   1290
      TabIndex        =   1
      Top             =   1350
      Width           =   3435
   End
   Begin VB.TextBox txtTipoClase 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   930
      Width           =   1215
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TIPOS DE CLASES"
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
      Left            =   495
      TabIndex        =   7
      Top             =   210
      Width           =   3975
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
      Left            =   90
      TabIndex        =   6
      Top             =   1455
      Width           =   1230
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
      Left            =   90
      TabIndex        =   5
      Top             =   1020
      Width           =   1230
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
      Left            =   1500
      TabIndex        =   4
      Top             =   945
      Width           =   1095
   End
End
Attribute VB_Name = "frmTipoClase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA TIPOS DE CLASES
' Objetivo: CATÁLOGO DE TIPOS DE CLASES
' Programado por:
' Fecha: FEBRERO DE 2004
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsTipoClase As ADODB.Recordset
        
Private Function VerificaDatos()
    Dim i, intInicio As Integer
    If txtTipoClase(0).Visible = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    For i = intInicio To 1
        If (txtTipoClase(i).Text = "") Then
            MsgBox "¡ Favor de Llenar Todas las Casillas, Pues No son Opcionales !", _
                        vbOKOnly + vbExclamation, "TipoClase (Captura)"
            VerificaDatos = False
            txtTipoClase(i).SetFocus
            Exit Function
        End If
        VerificaDatos = True
    Next i
    
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT idtipoclase FROM tipo_clase WHERE idtipoclase = " & _
                        Val(Trim(txtTipoClase(0).Text))
        Set AdoRcsTipoClase = New ADODB.Recordset
        AdoRcsTipoClase.ActiveConnection = Conn
        AdoRcsTipoClase.LockType = adLockOptimistic
        AdoRcsTipoClase.CursorType = adOpenKeyset
        AdoRcsTipoClase.CursorLocation = adUseServer
        AdoRcsTipoClase.Open strSQL
        If Not AdoRcsTipoClase.EOF Then
            MsgBox "Ya Existe Un Registro Con La Clave: " & _
                        txtTipoClase(0).Text, vbInformation + vbOKOnly, "TipoClase"
            AdoRcsTipoClase.Close
            VerificaDatos = False
            txtTipoClase(0).SetFocus
            Exit Function
        Else
            AdoRcsTipoClase.Close
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
        txtTipoClase(0).Visible = True
        Call Llena_txtTipoClase
    Else
        txtTipoClase(0).Visible = False
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
    strSQL = "INSERT INTO Tipo_Clase (idtipoclase, descripcion) VALUES (" & _
                    Val(Trim(txtTipoClase(0))) & ", '" & Trim(txtTipoClase(1)) & "')"
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
    strSQL = "DELETE FROM Tipo_Clase WHERE idtipoclase = " & Val(lblClave.Caption)
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Ahora Insertamos el nuevo registro que sustituye al anterior
    strSQL = "INSERT INTO Tipo_Clase (idtipoclase, descripcion) VALUES (" & _
                    Val(lblClave) & ", '" & Trim(txtTipoClase(1)) & "')"
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

Private Sub txtTipoClase_GotFocus(Index As Integer)
    txtTipoClase(Index).SelStart = 0
    txtTipoClase(Index).SelLength = Len(txtTipoClase(Index))
End Sub

Private Sub txtTipoClase_KeyPress(Index As Integer, KeyAscii As Integer)
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
        txtTipoClase(i).Text = ""
    Next i
    Call Llena_txtTipoClase
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM tipo_clase WHERE idtipoclase = " & Val(frmCatalogos.lblModo.Caption)
    Set AdoRcsTipoClase = New ADODB.Recordset
    AdoRcsTipoClase.ActiveConnection = Conn
    AdoRcsTipoClase.LockType = adLockOptimistic
    AdoRcsTipoClase.CursorType = adOpenKeyset
    AdoRcsTipoClase.CursorLocation = adUseServer
    AdoRcsTipoClase.Open strSQL
    If Not AdoRcsTipoClase.EOF Then
        lblClave.Caption = AdoRcsTipoClase!idtipoclase
        txtTipoClase(1).Text = AdoRcsTipoClase!Descripcion
    End If
End Sub

Sub Llena_txtTipoClase()
    Dim lngAnterior, lngTipoClase As Long
    'Llena txtTipoClase
    lngTipoClase = 1
    strSQL = "SELECT idtipoclase FROM tipo_clase ORDER BY idtipoclase"
    Set AdoRcsTipoClase = New ADODB.Recordset
    AdoRcsTipoClase.ActiveConnection = Conn
    AdoRcsTipoClase.LockType = adLockOptimistic
    AdoRcsTipoClase.CursorType = adOpenKeyset
    AdoRcsTipoClase.CursorLocation = adUseServer
    AdoRcsTipoClase.Open strSQL
    If AdoRcsTipoClase.EOF Then
        lngTipoClase = 1
        txtTipoClase(0).Text = lngTipoClase
        Exit Sub
    End If
    AdoRcsTipoClase.MoveFirst
    Do While Not AdoRcsTipoClase.EOF
        If AdoRcsTipoClase.Fields!idtipoclase <> "1" Then
            If Val(AdoRcsTipoClase.Fields!idtipoclase) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngTipoClase
        AdoRcsTipoClase.MoveNext
        If Not AdoRcsTipoClase.EOF Then lngTipoClase = AdoRcsTipoClase.Fields!idtipoclase
    Loop
    txtTipoClase(0).Text = lngAnterior + 1
End Sub
