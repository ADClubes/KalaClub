VERSION 5.00
Begin VB.Form frmPaises 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paises"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5190
   Icon            =   "frmPaises.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5190
   Begin VB.TextBox txtPais 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   1425
      TabIndex        =   1
      Top             =   555
      Width           =   1230
   End
   Begin VB.TextBox txtPais 
      Height          =   330
      Index           =   1
      Left            =   1425
      TabIndex        =   3
      Top             =   1005
      Width           =   3585
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2865
      Picture         =   "frmPaises.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar"
      Top             =   1470
      Width           =   795
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   4230
      Picture         =   "frmPaises.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   1470
      Width           =   795
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAÍSES"
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
      Left            =   608
      TabIndex        =   7
      Top             =   60
      Width           =   3975
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
      Height          =   315
      Left            =   1455
      TabIndex        =   6
      Top             =   570
      Width           =   1110
   End
   Begin VB.Label lblRentables 
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
      Top             =   675
      Width           =   1245
   End
   Begin VB.Label lblRentables 
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
      Left            =   135
      TabIndex        =   2
      Top             =   1110
      Width           =   1245
   End
End
Attribute VB_Name = "frmPaises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA: PAISES
' Objetivo: CATÁLOGO DE LOS PAISES
' Programado por:
' Fecha: JULIO 2004
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsPaises As ADODB.Recordset
        
Private Function VerificaDatos()
    Dim i, intInicio As Integer
    If txtPais(0).Visible = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    For i = intInicio To 1
        If (txtPais(i).Text = "") Then
            MsgBox "¡ Favor de Llenar Todas las Casillas, Pues NO son Opcionales !", _
                        vbOKOnly + vbExclamation, "Paises (Captura)"
            VerificaDatos = False
            txtPais(i).SetFocus
            Exit Function
        End If
        VerificaDatos = True
    Next i
    
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT idpais FROM paises WHERE idpais = " & _
                        Val(Trim(txtPais(0).Text))
        Set AdoRcsPaises = New ADODB.Recordset
        AdoRcsPaises.ActiveConnection = Conn
        AdoRcsPaises.LockType = adLockOptimistic
        AdoRcsPaises.CursorType = adOpenKeyset
        AdoRcsPaises.CursorLocation = adUseServer
        AdoRcsPaises.Open strSQL
        If Not AdoRcsPaises.EOF Then
            MsgBox "Ya Existe Un Registro Con La Clave: " & _
                        txtPais(0).Text, vbInformation + vbOKOnly, "Paises"
            AdoRcsPaises.Close
            VerificaDatos = False
            txtPais(0).SetFocus
            Exit Function
        Else
            AdoRcsPaises.Close
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
        txtPais(0).Visible = True
        Call Llena_txtpais
    Else
        txtPais(0).Visible = False
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
    strSQL = "INSERT INTO paises (idpais, pais) VALUES (" & Val(Trim(txtPais(0))) & ", '" & Trim(txtPais(1)) & "')"
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
    strSQL = "DELETE FROM paises WHERE idpais = " & Val(lblClave.Caption)
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Ahora Insertamos el nuevo registro que sustituye al anterior
    strSQL = "INSERT INTO paises (idpais, pais) VALUES (" & Val(Trim(lblClave.Caption)) & ", '" & Trim(txtPais(1)) & "')"
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

Private Sub txtpais_GotFocus(Index As Integer)
    txtPais(Index).SelStart = 0
    txtPais(Index).SelLength = Len(txtPais(Index))
End Sub

Private Sub txtpais_KeyPress(Index As Integer, KeyAscii As Integer)
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
        txtPais(i).Text = ""
    Next i
    Call Llena_txtpais
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM paises WHERE idpais = " & Val(Trim(frmCatalogos.lblModo.Caption))
    Set AdoRcsPaises = New ADODB.Recordset
    AdoRcsPaises.ActiveConnection = Conn
    AdoRcsPaises.LockType = adLockOptimistic
    AdoRcsPaises.CursorType = adOpenKeyset
    AdoRcsPaises.CursorLocation = adUseServer
    AdoRcsPaises.Open strSQL
    If Not AdoRcsPaises.EOF Then
        lblClave.Caption = AdoRcsPaises!idpais
        txtPais(1).Text = AdoRcsPaises!pais
    End If
End Sub

Sub Llena_txtpais()
    Dim lngAnterior, lngPaises As Long
    'Llena txtpais
    lngPaises = 1
    strSQL = "SELECT idpais FROM paises ORDER BY idpais"
    Set AdoRcsPaises = New ADODB.Recordset
    AdoRcsPaises.ActiveConnection = Conn
    AdoRcsPaises.LockType = adLockOptimistic
    AdoRcsPaises.CursorType = adOpenKeyset
    AdoRcsPaises.CursorLocation = adUseServer
    AdoRcsPaises.Open strSQL
    If AdoRcsPaises.EOF Then
        lngPaises = 1
        txtPais(0).Text = lngPaises
        Exit Sub
    End If
    AdoRcsPaises.MoveFirst
    Do While Not AdoRcsPaises.EOF
        If AdoRcsPaises.Fields!idpais <> "1" Then
            If Val(AdoRcsPaises.Fields!idpais) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngPaises
        AdoRcsPaises.MoveNext
        If Not AdoRcsPaises.EOF Then lngPaises = AdoRcsPaises.Fields!idpais
    Loop
    txtPais(0).Text = lngAnterior + 1
End Sub
