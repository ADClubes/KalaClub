VERSION 5.00
Begin VB.Form frmPagos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos (Captura)"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5265
   Icon            =   "frmPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5265
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   4230
      Picture         =   "frmPagos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   1620
      Width           =   795
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      Height          =   840
      Left            =   2835
      Picture         =   "frmPagos.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar"
      Top             =   1620
      Width           =   795
   End
   Begin VB.TextBox txtPagos 
      Height          =   330
      Index           =   1
      Left            =   1500
      TabIndex        =   3
      Top             =   1125
      Width           =   3555
   End
   Begin VB.TextBox txtPagos 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   1500
      TabIndex        =   1
      Top             =   675
      Width           =   1395
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FORMAS DE PAGO"
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
      Left            =   645
      TabIndex        =   7
      Top             =   165
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
      Left            =   1515
      TabIndex        =   6
      Top             =   705
      Width           =   1125
   End
   Begin VB.Label lblPagos 
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
      Left            =   285
      TabIndex        =   2
      Top             =   1230
      Width           =   1230
   End
   Begin VB.Label lblPagos 
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
      Left            =   285
      TabIndex        =   0
      Top             =   795
      Width           =   1230
   End
End
Attribute VB_Name = "frmPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA: PAGOS
' Objetivo: CATÁLOGO DE PAGOS
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsPagos As ADODB.Recordset
        
Private Function VerificaDatos()
    Dim i, intInicio As Integer
    If txtPagos(0).Visible = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    For i = intInicio To 1
        If (txtPagos(i).Text = "") Then
            MsgBox "¡ Favor de Llenar Todas las Casillas, Pues No son Opcionales.", _
                        vbOKOnly + vbExclamation, "pagos (Captura)"
            VerificaDatos = False
            txtPagos(i).SetFocus
            Exit Function
        End If
        VerificaDatos = True
    Next i
    
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT idformapago FROM forma_pago WHERE idformapago = " & _
                        Val(Trim(txtPagos(0).Text))
        Set AdoRcsPagos = New ADODB.Recordset
        AdoRcsPagos.ActiveConnection = Conn
        AdoRcsPagos.LockType = adLockOptimistic
        AdoRcsPagos.CursorType = adOpenKeyset
        AdoRcsPagos.CursorLocation = adUseServer
        AdoRcsPagos.Open strSQL
        If Not AdoRcsPagos.EOF Then
            MsgBox "Ya Existe Un Registro Con La Clave: " & _
                        txtPagos(0).Text, vbInformation + vbOKOnly, "Forma de Pagos"
            AdoRcsPagos.Close
            VerificaDatos = False
            txtPagos(0).SetFocus
            Exit Function
        Else
            AdoRcsPagos.Close
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
        txtPagos(0).Visible = True
        Call Llena_txtPagos
    Else
        txtPagos(0).Visible = False
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
    strSQL = "INSERT INTO forma_pago (idformapago, descripcion) VALUES (" & _
                 Val(Trim(txtPagos(0))) & ", '" & Trim(txtPagos(1)) & "')"
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
    strSQL = "DELETE FROM forma_pago WHERE idformapago = " & Val(lblClave.Caption)
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Ahora Insertamos el nuevo registro que sustituye al anterior
    strSQL = "INSERT INTO forma_pago (idformapago, descripcion) VALUES (" & _
                    Val(Trim(lblClave.Caption)) & ", '" & Trim(txtPagos(1)) & "')"
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

Private Sub txtPagos_GotFocus(Index As Integer)
    txtPagos(Index).SelStart = 0
    txtPagos(Index).SelLength = Len(txtPagos(Index))
End Sub

Private Sub txtpagos_KeyPress(Index As Integer, KeyAscii As Integer)
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
        txtPagos(i).Text = ""
    Next i
    Call Llena_txtPagos
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM forma_pago WHERE idformapago = " & _
                  Val(Trim(frmCatalogos.lblModo.Caption))
    Set AdoRcsPagos = New ADODB.Recordset
    AdoRcsPagos.ActiveConnection = Conn
    AdoRcsPagos.LockType = adLockOptimistic
    AdoRcsPagos.CursorType = adOpenKeyset
    AdoRcsPagos.CursorLocation = adUseServer
    AdoRcsPagos.Open strSQL
    If Not AdoRcsPagos.EOF Then
        lblClave.Caption = AdoRcsPagos!idformapago
        txtPagos(1).Text = AdoRcsPagos!Descripcion
    End If
End Sub

Sub Llena_txtPagos()
    Dim lngAnterior, lngPagos As Long
    'Llena txtPAgos
    lngPagos = 1
    strSQL = "SELECT idformapago FROM forma_pago ORDER BY idformapago"
    Set AdoRcsPagos = New ADODB.Recordset
    AdoRcsPagos.ActiveConnection = Conn
    AdoRcsPagos.LockType = adLockOptimistic
    AdoRcsPagos.CursorType = adOpenKeyset
    AdoRcsPagos.CursorLocation = adUseServer
    AdoRcsPagos.Open strSQL
    If AdoRcsPagos.EOF Then
        lngPagos = 1
        txtPagos(0).Text = lngPagos
        Exit Sub
    End If
    AdoRcsPagos.MoveFirst
    Do While Not AdoRcsPagos.EOF
        If AdoRcsPagos.Fields!idformapago <> "1" Then
            If Val(AdoRcsPagos.Fields!idformapago) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngPagos
        AdoRcsPagos.MoveNext
        If Not AdoRcsPagos.EOF Then lngPagos = AdoRcsPagos.Fields!idformapago
    Loop
    txtPagos(0).Text = lngAnterior + 1
End Sub

