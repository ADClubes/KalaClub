VERSION 5.00
Begin VB.Form frmTipoRentable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo Rentables"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5310
   Icon            =   "frmTipoRentables.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmTipoRentables.frx":0442
   ScaleHeight     =   2700
   ScaleWidth      =   5310
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   4245
      Picture         =   "frmTipoRentables.frx":B4BE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   1635
      Width           =   795
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
      Left            =   2895
      Picture         =   "frmTipoRentables.frx":B7C8
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar"
      Top             =   1635
      Width           =   795
   End
   Begin VB.TextBox txtRentables 
      Height          =   330
      Index           =   1
      Left            =   1455
      TabIndex        =   1
      Top             =   1170
      Width           =   3585
   End
   Begin VB.TextBox txtRentables 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   1455
      TabIndex        =   0
      Top             =   720
      Width           =   1230
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TIPOS RENTABLES"
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
      Left            =   810
      TabIndex        =   7
      Top             =   75
      Width           =   3975
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
      Left            =   165
      TabIndex        =   6
      Top             =   1275
      Width           =   1245
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
      Left            =   180
      TabIndex        =   5
      Top             =   840
      Width           =   1245
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
      Left            =   1485
      TabIndex        =   4
      Top             =   735
      Width           =   1110
   End
End
Attribute VB_Name = "frmTipoRentable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA: TIPO RENTABLES
' Objetivo: CATÁLOGO DE LOS ITEMS QUE SON RENTABLES
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsRentables As ADODB.Recordset
        
Private Function VerificaDatos()
    Dim i, intInicio As Integer
    If txtRentables(0).Visible = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    For i = intInicio To 1
        If (txtRentables(i).Text = "") Then
            MsgBox "¡ Favor de Llenar Todas las Casillas, Pues NO son Opcionales !", _
                        vbOKOnly + vbExclamation, "Tipo Rentables (Captura)"
            VerificaDatos = False
            txtRentables(i).SetFocus
            Exit Function
        End If
        VerificaDatos = True
    Next i
    
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT idtiporentable FROM tipo_rentables WHERE idtiporentable = " & _
                        Val(Trim(txtRentables(0).Text))
        Set AdoRcsRentables = New ADODB.Recordset
        AdoRcsRentables.ActiveConnection = Conn
        AdoRcsRentables.LockType = adLockOptimistic
        AdoRcsRentables.CursorType = adOpenKeyset
        AdoRcsRentables.CursorLocation = adUseServer
        AdoRcsRentables.Open strSQL
        If Not AdoRcsRentables.EOF Then
            MsgBox "Ya Existe Un Registro Con La Clave: " & _
                        txtRentables(0).Text, vbInformation + vbOKOnly, "Tipos Rentables"
            AdoRcsRentables.Close
            VerificaDatos = False
            txtRentables(0).SetFocus
            Exit Function
        Else
            AdoRcsRentables.Close
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
        txtRentables(0).Visible = True
        Call Llena_txtRentables
    Else
        txtRentables(0).Visible = False
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
    strSQL = "INSERT INTO tipo_rentables (idtiporentable, descripcion) VALUES (" & _
                 Val(Trim(txtRentables(0))) & ", '" & Trim(txtRentables(1)) & "')"
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
    strSQL = "DELETE FROM tipo_rentables WHERE idtiporentable = " & Val(lblClave.Caption)
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Ahora Insertamos el nuevo registro que sustituye al anterior
    strSQL = "INSERT INTO tipo_rentables (idtiporentable, descripcion) VALUES (" & _
                    Val(Trim(lblClave.Caption)) & ", '" & Trim(txtRentables(1)) & "')"
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

Private Sub txtRentables_GotFocus(Index As Integer)
    txtRentables(Index).SelStart = 0
    txtRentables(Index).SelLength = Len(txtRentables(Index))
End Sub

Private Sub txtRentables_KeyPress(Index As Integer, KeyAscii As Integer)
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
        txtRentables(i).Text = ""
    Next i
    Call Llena_txtRentables
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM tipo_rentables WHERE idtiporentable = " & _
                  Val(Trim(frmCatalogos.lblModo.Caption))
    Set AdoRcsRentables = New ADODB.Recordset
    AdoRcsRentables.ActiveConnection = Conn
    AdoRcsRentables.LockType = adLockOptimistic
    AdoRcsRentables.CursorType = adOpenKeyset
    AdoRcsRentables.CursorLocation = adUseServer
    AdoRcsRentables.Open strSQL
    If Not AdoRcsRentables.EOF Then
        lblClave.Caption = AdoRcsRentables!idtiporentable
        txtRentables(1).Text = AdoRcsRentables!Descripcion
    End If
End Sub

Sub Llena_txtRentables()
    Dim lngAnterior, lngRentables As Long
    'Llena txtRentables
    lngRentables = 1
    strSQL = "SELECT idtiporentable FROM tipo_rentables ORDER BY idtiporentable"
    Set AdoRcsRentables = New ADODB.Recordset
    AdoRcsRentables.ActiveConnection = Conn
    AdoRcsRentables.LockType = adLockOptimistic
    AdoRcsRentables.CursorType = adOpenKeyset
    AdoRcsRentables.CursorLocation = adUseServer
    AdoRcsRentables.Open strSQL
    If AdoRcsRentables.EOF Then
        lngRentables = 1
        txtRentables(0).Text = lngRentables
        Exit Sub
    End If
    AdoRcsRentables.MoveFirst
    Do While Not AdoRcsRentables.EOF
        If AdoRcsRentables.Fields!idtiporentable <> "1" Then
            If Val(AdoRcsRentables.Fields!idtiporentable) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngRentables
        AdoRcsRentables.MoveNext
        If Not AdoRcsRentables.EOF Then lngRentables = AdoRcsRentables.Fields!idtiporentable
    Loop
    txtRentables(0).Text = lngAnterior + 1
End Sub
