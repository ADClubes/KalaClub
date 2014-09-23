VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMultiplesRentables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rentables (Captura Múltiple)"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6510
   Icon            =   "frmAgregaMultiples.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6510
   Begin VB.Frame Frame1 
      Caption         =   "Números a Generar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   200
      TabIndex        =   8
      Top             =   2500
      Width           =   3420
      Begin VB.TextBox txtFijo 
         Alignment       =   2  'Center
         Height          =   345
         Index           =   0
         Left            =   255
         TabIndex        =   10
         Top             =   600
         Width           =   440
      End
      Begin VB.TextBox txtDesde 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   1095
         TabIndex        =   12
         Top             =   600
         Width           =   440
      End
      Begin VB.TextBox txtHasta 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   1740
         TabIndex        =   14
         Top             =   600
         Width           =   440
      End
      Begin VB.TextBox txtFijo 
         Alignment       =   2  'Center
         Height          =   345
         Index           =   1
         Left            =   2700
         TabIndex        =   16
         Top             =   600
         Width           =   440
      End
      Begin VB.Label lblRentables 
         Caption         =   "(*) &Prefijo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   330
         Width           =   885
      End
      Begin VB.Label lblRentables 
         Caption         =   "&Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1035
         TabIndex        =   11
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblRentables 
         Caption         =   "&Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1740
         TabIndex        =   13
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lblRentables 
         Caption         =   "(*) &Subfijo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2430
         TabIndex        =   15
         Top             =   330
         Width           =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   750
         X2              =   1020
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   765
         X2              =   1035
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   2295
         X2              =   2565
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         X1              =   2310
         X2              =   2580
         Y1              =   720
         Y2              =   720
      End
   End
   Begin MSComctlLib.ProgressBar pgrAvance 
      Height          =   285
      Left            =   200
      TabIndex        =   21
      Top             =   3930
      Visible         =   0   'False
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame frmMuestra 
      Caption         =   "Muestra a Generar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   3750
      TabIndex        =   19
      Top             =   2500
      Width           =   2565
      Begin VB.Label lblMuestra 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   105
         TabIndex        =   20
         Top             =   315
         Width           =   2355
      End
   End
   Begin VB.TextBox txtRentables 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1965
      TabIndex        =   7
      Top             =   1995
      Width           =   1530
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   5220
      Picture         =   "frmAgregaMultiples.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Salir"
      Top             =   1350
      Width           =   795
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      Height          =   840
      Left            =   5205
      Picture         =   "frmAgregaMultiples.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Guardar"
      Top             =   210
      Width           =   795
   End
   Begin VB.ComboBox cboTipoRentable 
      Height          =   315
      Left            =   1545
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   555
      Width           =   3165
   End
   Begin VB.Frame frmSexo 
      Caption         =   "Sexo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   200
      TabIndex        =   2
      Top             =   1050
      Width           =   4425
      Begin VB.OptionButton optSexo 
         Caption         =   "&Indistinto"
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
         Left            =   3045
         TabIndex        =   5
         Top             =   240
         Width           =   1110
      End
      Begin VB.OptionButton optSexo 
         Caption         =   "&Femenino"
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
         Left            =   1725
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
      Begin VB.OptionButton optSexo 
         Caption         =   "&Masculino"
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
         Left            =   315
         TabIndex        =   3
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MÚLTIPLES RENTABLES"
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
      Left            =   1275
      TabIndex        =   24
      Top             =   105
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(Máximo 10 Caracteres)"
      Height          =   180
      Left            =   225
      TabIndex        =   23
      Top             =   2175
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "(*)  Opcional"
      Height          =   195
      Left            =   390
      TabIndex        =   22
      Top             =   3945
      Width           =   1230
   End
   Begin VB.Label lblRentables 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo &Rentable:"
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
      Left            =   200
      TabIndex        =   0
      Top             =   630
      Width           =   1365
   End
   Begin VB.Label lblRentables 
      BackStyle       =   0  'Transparent
      Caption         =   "(*) &Ubicación:"
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
      Left            =   225
      TabIndex        =   6
      Top             =   1920
      Width           =   1320
   End
End
Attribute VB_Name = "frmMultiplesRentables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA GENERACIÓN DE RENTABLES MULTIPLE
' Objetivo: GENERA ITEMS RENTABLES DE MANERA MASIVA
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsRentables As ADODB.Recordset
    Dim strSexo, strNumero As String
    Dim intTipoRentable As Integer
    
Private Function VerificaDatos()
    Dim intCadena As Integer
    Dim strCadena As String
    If (cboTipoRentable.Text = "") Then
        MsgBox "¡ Favor de seleccionar el TIPO RENTABLE, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de rentables (Captura)"
        VerificaDatos = False
        cboTipoRentable.SetFocus
        Exit Function
    End If
    If (optSexo(0).Value = vbUnchecked) And (optSexo(1).Value = vbUnchecked) And (optSexo(2).Value = vbUnchecked) Then
        MsgBox "¡ Favor de Seleccionar el SEXO, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de rentables (Captura)"
        VerificaDatos = False
        Exit Function
    End If
    If (txtDesde.Text = "") Then
        MsgBox "¡ Favor de Digitar el Valor Inicial DESDE, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de rentables (Captura)"
        VerificaDatos = False
        txtDesde.SetFocus
        Exit Function
    End If
    If (txtHasta.Text = "") Then
        MsgBox "¡ Favor de Digitar el Valor Final HASTA, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de rentables (Captura)"
        VerificaDatos = False
        txtHasta.SetFocus
        Exit Function
    End If
    If (Val(Trim(txtDesde.Text)) > Val(Trim(txtHasta.Text))) Then
        MsgBox "¡ No son Correctos los Datos DESDE y/o HASTA !", _
                    vbOKOnly + vbExclamation, "Tipo de rentables (Captura)"
        VerificaDatos = False
        txtDesde.SetFocus
        Exit Function
    End If
    strCadena = txtFijo(0).Text & txtHasta.Text & txtFijo(1).Text
    If Len(strCadena) > 6 Then
        MsgBox "¡ El Largo de Uno o Más Números a Generar, Excede los 6 Caracteres !", _
                    vbOKOnly + vbExclamation, "Tipo de rentables (Captura)"
        VerificaDatos = False
        txtFijo(0).SetFocus
        Exit Function
    End If
    VerificaDatos = True
End Function

Private Sub cmdGuardar_Click()
    Dim blnGuarda As Boolean
    blnGuarda = VerificaDatos
    If blnGuarda = True Then
        Call GuardaDatos
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strCampo1, strCampo2 As String
    frmCatalogos.Enabled = False
    
    'Llena los combo Id Tipo Rentable
    strSQL = "SELECT idtiporentable, descripcion FROM tipo_rentables"
    strCampo1 = "descripcion"
    strCampo2 = "idtiporentable"
    Call LlenaCombos(cboTipoRentable, strSQL, strCampo1, strCampo2)

'    Set AdoRcsRentables = New ADODB.Recordset
'    AdoRcsRentables.ActiveConnection = Conn
'    AdoRcsRentables.LockType = adLockOptimistic
'    AdoRcsRentables.CursorType = adOpenKeyset
'    AdoRcsRentables.CursorLocation = adUseServer
'    AdoRcsRentables.Open strSQL
'    cboTipoRentable.Clear
'    Do While Not AdoRcsRentables.EOF
'      cboTipoRentable.AddItem AdoRcsRentables.Fields!descripcion
'      'Llena el campo alterno del combo
'      cboTipoRentable.ItemData(cboTipoRentable.NewIndex) = AdoRcsRentables.Fields!idtiporentable
'      AdoRcsRentables.MoveNext
'    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCatalogos.Enabled = True
End Sub

Private Sub GuardaDatos()
    Dim AdoCmdInserta As ADODB.Command
    Dim i, intNoGuardados, intRespuesta, intTotal As Integer
    On Error GoTo err_Guarda
    intRespuesta = MsgBox("Se Generarán los Registros Indicados. " & _
                                        "¿Está Seguro?", vbOKCancel + vbQuestion, "Generación Múltiple")
    If intRespuesta = 2 Then
        Exit Sub
    End If
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    cmdGuardar.Enabled = False
    cmdSalir.Enabled = False
    frmMuestra.Caption = "Guardando:"
    intTotal = (Val(txtHasta.Text) - Val(txtDesde.Text)) + 1
    pgrAvance.Visible = True
    pgrAvance.Max = intTotal
    pgrAvance.Value = 0
    intNoGuardados = 0
    intTipoRentable = cboTipoRentable.ItemData(cboTipoRentable.ListIndex)
    For i = Val(txtDesde.Text) To Val(txtHasta.Text)
        strNumero = txtFijo(0) & i & txtFijo(1)    'Asignamos el Número a Insertar
        Call LlenaEspacios                      'Llenamos los espacios del numero
        lblMuestra.Caption = strNumero
        'Consultamos para ver si no exixte previamente es registro
        strSQL = "SELECT idtiporentable FROM rentables WHERE (idtiporentable = " & _
                        intTipoRentable & ") AND (numero = '" & strNumero & "')"
        Set AdoRcsRentables = New ADODB.Recordset
        AdoRcsRentables.ActiveConnection = Conn
        AdoRcsRentables.LockType = adLockOptimistic
        AdoRcsRentables.CursorType = adOpenKeyset
        AdoRcsRentables.CursorLocation = adUseServer
        AdoRcsRentables.Open strSQL
        If AdoRcsRentables.EOF Then
            strSQL = "INSERT INTO rentables (idtiporentable, numero, sexo, " & _
                            "ubicacion) VALUES (" & intTipoRentable & ", '" & strNumero & _
                            "', '" & strSexo & "', '" & Trim(txtRentables.Text) & "')"
            Set AdoCmdInserta = New ADODB.Command
            AdoCmdInserta.ActiveConnection = Conn
            AdoCmdInserta.CommandText = strSQL
            AdoCmdInserta.Execute
        Else
            intNoGuardados = intNoGuardados + 1
        End If
        pgrAvance.Value = pgrAvance.Value + 1
    Next i
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    MsgBox "¡Proceso Terminado! Registros Insertados: " & intTotal - intNoGuardados & ", Registros NO Insertados: " & intNoGuardados
    cmdGuardar.Enabled = True
    cmdSalir.Enabled = True
    Call Limpia
    frmCatalogos.AdoDcCatal.REFRESH
    frmCatalogos.grdCatalogos.REFRESH
    frmCatalogos.lblTotal.Caption = Format(frmCatalogos.AdoDcCatal.Recordset.RecordCount, "#######")
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    cmdGuardar.Enabled = True
    cmdSalir.Enabled = True
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub optSexo_Click(Index As Integer)
    Select Case Index
        Case 0
            strSexo = "M"
        Case 1
            strSexo = "F"
        Case 2
            strSexo = "X"
    End Select
End Sub

Private Sub txtDesde_Change()
    Call ActualizaMuestra
End Sub

Private Sub txtDesde_GotFocus()
    txtDesde.SelStart = 0
    txtDesde.SelLength = Len(txtDesde)
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 22, 48 To 57     'Backspace, <Ctrl+V> y del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtFijo_Change(Index As Integer)
    Call ActualizaMuestra
End Sub

Private Sub txtFijo_GotFocus(Index As Integer)
    txtFijo(Index).SelStart = 0
    txtFijo(Index).SelLength = Len(txtFijo(Index))
End Sub

Private Sub txtFijo_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtHasta_Change()
    Call ActualizaMuestra
End Sub

Private Sub txtHasta_GotFocus()
    txtHasta.SelStart = 0
    txtHasta.SelLength = Len(txtHasta)
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 22, 48 To 57     'Backspace, <Ctrl+V> y del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtRentables_GotFocus()
    txtRentables.SelStart = 0
    txtRentables.SelLength = Len(txtRentables)
End Sub

Private Sub txtRentables_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub ActualizaMuestra()
    lblMuestra = txtFijo(0).Text & txtDesde.Text & txtFijo(1).Text & " - " & _
                        txtFijo(0).Text & txtHasta.Text & txtFijo(1).Text
End Sub

Sub Limpia()
    Dim i As Integer
    cboTipoRentable.ListIndex = -1
    txtRentables.Text = ""
    For i = 0 To 2
        If i < 2 Then txtFijo(i).Text = ""
        optSexo(i).Value = False
    Next i
    txtDesde.Text = ""
    txtHasta.Text = ""
    frmMuestra.Caption = "Muestra a Generar:"
End Sub

Sub LlenaEspacios()
    Dim intLargo As Integer
    If Left(strNumero, 1) <> " " Then
        intLargo = Len(strNumero)
        If intLargo < 6 Then
            While intLargo < 6
                strNumero = " " & strNumero
                intLargo = Len(strNumero)
            Wend
        End If
    End If
End Sub
