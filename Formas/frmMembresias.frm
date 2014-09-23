VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMembresias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Membresías (Captura)"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7080
   Icon            =   "frmMembresias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7080
   Begin VB.VScrollBar vsbAnios 
      Height          =   285
      Left            =   5010
      Max             =   99
      Min             =   1
      TabIndex        =   17
      Top             =   2955
      Value           =   1
      Width           =   210
   End
   Begin VB.TextBox txtMembresia 
      Height          =   330
      Index           =   4
      Left            =   4410
      TabIndex        =   13
      Top             =   2925
      Width           =   840
   End
   Begin MSComCtl2.DTPicker dtpAlta 
      Height          =   285
      Left            =   2070
      TabIndex        =   9
      Top             =   2490
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   57147393
      CurrentDate     =   37971
   End
   Begin VB.TextBox txtMembresia 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   2
      Left            =   2070
      TabIndex        =   5
      Top             =   1920
      Width           =   1000
   End
   Begin VB.TextBox txtMembresia 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   0
      Left            =   2070
      TabIndex        =   1
      Top             =   765
      Width           =   870
   End
   Begin VB.TextBox txtMembresia 
      Height          =   540
      Index           =   1
      Left            =   2070
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1215
      Width           =   3675
   End
   Begin VB.TextBox txtMembresia 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   3
      Left            =   4725
      TabIndex        =   7
      Text            =   "0"
      Top             =   1920
      Width           =   1000
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      Height          =   840
      Left            =   6000
      Picture         =   "frmMembresias.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Guardar"
      Top             =   1050
      Width           =   795
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   6000
      Picture         =   "frmMembresias.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   2190
      Width           =   795
   End
   Begin MSComCtl2.DTPicker dtpVigencia 
      Height          =   285
      Left            =   2070
      TabIndex        =   11
      Top             =   2985
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      Format          =   57147393
      CurrentDate     =   37971
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MEMBRESÍAS"
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
      Left            =   1553
      TabIndex        =   18
      Top             =   225
      Width           =   3975
   End
   Begin VB.Label lblMembresia 
      BackStyle       =   0  'Transparent
      Caption         =   "D&uración:                años."
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
      Left            =   3525
      TabIndex        =   12
      Top             =   3030
      Width           =   2445
   End
   Begin VB.Label lblMembresia 
      Alignment       =   1  'Right Justify
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
      Left            =   1410
      TabIndex        =   0
      Top             =   855
      Width           =   555
   End
   Begin VB.Label lblMembresia 
      Alignment       =   1  'Right Justify
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
      Left            =   900
      TabIndex        =   2
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label lblMembresia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Precio de Contado: $"
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
      Index           =   2
      Left            =   210
      TabIndex        =   4
      Top             =   1995
      Width           =   1815
   End
   Begin VB.Label lblMembresia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "&Inscripción: $"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   1995
      Width           =   1170
   End
   Begin VB.Label lblMembresia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de &Alta:"
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
      Left            =   675
      TabIndex        =   8
      Top             =   2535
      Width           =   1275
   End
   Begin VB.Label lblMembresia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha &Vigencia:"
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
      Left            =   570
      TabIndex        =   10
      Top             =   3030
      Width           =   1410
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
      Left            =   2010
      TabIndex        =   16
      Top             =   780
      Width           =   1110
   End
End
Attribute VB_Name = "frmMembresias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA MEMBRESIAS
' Objetivo: CATÁLOGO DE MEMBRESÍAS
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsMembresias As ADODB.Recordset
    
Private Function VerificaDatos()
    Dim i, intInicio As Integer
    If txtMembresia(0).Visible = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    For i = intInicio To 4
        Select Case i
            Case 0
                If (txtMembresia(i).Text = "") Then
                    MsgBox "¡ Favor de Llenar la Casilla CLAVE, Pues No es Opcional !", _
                                 vbOKOnly + vbExclamation, "Membresias (Captura)"
                    VerificaDatos = False
                    txtMembresia(i).SetFocus
                    Exit Function
                End If
            Case 1
                If txtMembresia(i).Text = "" Then
                    MsgBox "¡ Favor de Llenar la Casilla DESCRIPCIÓN, Pues No es Opcional.", _
                                vbOKOnly + vbExclamation, "Membresias (Captura)"
                    VerificaDatos = False
                    txtMembresia(i).SetFocus
                    Exit Function
                End If
            Case 2
                If (Not IsNumeric(txtMembresia(i))) Or (txtMembresia(i).Text = "") Then
                    MsgBox "¡ Precio de Contado Incorrecto o Vacío !", vbOKOnly + vbExclamation, _
                                "Membresias (Captura)"
                    VerificaDatos = False
                    txtMembresia(i).SetFocus
                    Exit Function
                End If
            Case 3
                If (Not IsNumeric(txtMembresia(i))) Or (txtMembresia(i).Text = "") Then
                    MsgBox "¡ Precio de Inscripción Incorrecto o Vacío !", vbOKOnly + vbExclamation, _
                                "Membresias (Captura)"
                    VerificaDatos = False
                    txtMembresia(i).SetFocus
                    Exit Function
                End If
            Case 4
                If (Val(Trim(txtMembresia(i))) < 1) Or (Val(Trim(txtMembresia(i))) > 99) _
                    Or (txtMembresia(i).Text = "") Then
                    MsgBox "¡ Duración Incorrecta o Vacía !", vbOKOnly + vbExclamation, _
                                "Membresias (Captura)"
                    VerificaDatos = False
                    txtMembresia(i).SetFocus
                    Exit Function
                End If
        End Select
    Next
    If dtpAlta > dtpVigencia Then
        MsgBox "¡ Fechas Incorrectas !", vbOKOnly + vbExclamation, "Membresias (Captura)"
        VerificaDatos = False
        dtpAlta.SetFocus
        Exit Function
    End If
    VerificaDatos = True
    
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT idmembresia FROM membresias WHERE idmembresia = " & _
                    Val(Trim(txtMembresia(0).Text))
        Set AdoRcsMembresias = New ADODB.Recordset
        AdoRcsMembresias.ActiveConnection = Conn
        AdoRcsMembresias.LockType = adLockOptimistic
        AdoRcsMembresias.CursorType = adOpenKeyset
        AdoRcsMembresias.CursorLocation = adUseServer
        AdoRcsMembresias.Open strSQL
        If Not AdoRcsMembresias.EOF Then
            MsgBox "Ya Existe Un Registro Con La Clave: " & _
                        txtMembresia(0).Text, vbInformation + vbOKOnly, "Membresias"
            AdoRcsMembresias.Close
            VerificaDatos = False
            txtMembresia(0).SetFocus
            Exit Function
        Else
            AdoRcsMembresias.Close
            VerificaDatos = True
        End If
    End If
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
    If frmCatalogos.lblModo.Caption = "A" Then
        txtMembresia(0).Visible = True
        txtMembresia(4).Text = vsbAnios.Value
        Call Llena_txtMembresia
    Else
        txtMembresia(0).Visible = False
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
    strSQL = "INSERT INTO membresias (idmembresia, descripcion, " & _
                    "precio_contado, inscripcion, fecha_alta, fecha_vigencia, " & _
                    "duracion) VALUES (" & Val(txtMembresia(0).Text) & ", '" & _
                    Trim(txtMembresia(1).Text) & "', " & Val(txtMembresia(2).Text) & _
                    ", " & Val(txtMembresia(3).Text) & ", '" & _
                    Format(dtpAlta, "dd/mm/yyyy") & "', '" & _
                    Format(dtpVigencia, "dd/mm/yyyy") & "', " & _
                    Val(txtMembresia(4).Text) & ")"
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
    strSQL = "DELETE FROM membresias WHERE idmembresia = " & _
                    Val(lblClave.Caption)
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Ahora Insertamos el nuevo registro que sustituye al anterior
    strSQL = "INSERT INTO membresias (idmembresia, descripcion, " & _
                    "precio_contado, inscripcion, fecha_alta, fecha_vigencia, " & _
                    "duracion) VALUES (" & Val(lblClave.Caption) & ", '" & _
                    Trim(txtMembresia(1).Text) & "', " & Val(txtMembresia(2).Text) & _
                    ", " & Val(txtMembresia(3).Text) & ", '" & _
                    Format(dtpAlta, "dd/mm/yyyy") & "', '" & _
                    Format(dtpVigencia, "dd/mm/yyyy") & "', " & _
                    Val(txtMembresia(4).Text) & ")"
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

Private Sub txtMembresia_GotFocus(Index As Integer)
    txtMembresia(Index).SelStart = 0
    txtMembresia(Index).SelLength = Len(txtMembresia(Index))
End Sub

Private Sub txtMembresia_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 2, 3, 4
            Select Case KeyAscii
                Case 8, 22, 46, 48 To 57  'Backspace, <Ctrl+V>, punto y del 0 al 9
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
    For i = 0 To 4
        txtMembresia(i).Text = ""
    Next i
    lblClave.Caption = ""
    Call Llena_txtMembresia
    txtMembresia(0).SetFocus
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM membresias WHERE idmembresia = " & _
                  Val(frmCatalogos.lblModo.Caption)
    Set AdoRcsMembresias = New ADODB.Recordset
    AdoRcsMembresias.ActiveConnection = Conn
    AdoRcsMembresias.LockType = adLockOptimistic
    AdoRcsMembresias.CursorType = adOpenKeyset
    AdoRcsMembresias.CursorLocation = adUseServer
    AdoRcsMembresias.Open strSQL
    If Not AdoRcsMembresias.EOF Then
        lblClave.Caption = AdoRcsMembresias!idmembresia
        txtMembresia(1).Text = AdoRcsMembresias!Descripcion
        txtMembresia(2).Text = AdoRcsMembresias!precio_contado
        txtMembresia(3).Text = AdoRcsMembresias!inscripcion
        dtpAlta.Value = Format(AdoRcsMembresias!fecha_alta, "dd/mm/yy")
        dtpVigencia.Value = Format(AdoRcsMembresias!fecha_vigencia, "dd/mm/yy")
        txtMembresia(4).Text = AdoRcsMembresias!duracion
    End If
End Sub

Private Sub vsbAnios_Change()
    txtMembresia(4).Text = 100 - vsbAnios.Value
End Sub


Sub Llena_txtMembresia()
    Dim lngAnterior, lngMembresia As Long
    'Llena txtMembresia
    lngMembresia = 1
    strSQL = "SELECT idmembresia FROM membresias ORDER BY idmembresia"
    Set AdoRcsMembresias = New ADODB.Recordset
    AdoRcsMembresias.ActiveConnection = Conn
    AdoRcsMembresias.LockType = adLockOptimistic
    AdoRcsMembresias.CursorType = adOpenKeyset
    AdoRcsMembresias.CursorLocation = adUseServer
    AdoRcsMembresias.Open strSQL
    If AdoRcsMembresias.EOF Then
        lngMembresia = 1
        txtMembresia(0).Text = lngMembresia
        Exit Sub
    End If
    AdoRcsMembresias.MoveFirst
    Do While Not AdoRcsMembresias.EOF
        If AdoRcsMembresias.Fields!idmembresia <> "1" Then
            If Val(AdoRcsMembresias.Fields!idmembresia) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngMembresia
        AdoRcsMembresias.MoveNext
        If Not AdoRcsMembresias.EOF Then lngMembresia = AdoRcsMembresias.Fields!idmembresia
    Loop
    txtMembresia(0).Text = lngAnterior + 1
End Sub
