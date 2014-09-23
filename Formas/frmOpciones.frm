VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones de Configuración"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7710
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmOpciones.frx":030A
   ScaleHeight     =   4845
   ScaleWidth      =   7710
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   500
      Left            =   3945
      Picture         =   "frmOpciones.frx":0614
      TabIndex        =   7
      Top             =   4245
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   500
      Left            =   5595
      TabIndex        =   8
      Top             =   4245
      Width           =   1500
   End
   Begin VB.Frame frmUnico 
      Height          =   2175
      Left            =   605
      TabIndex        =   10
      Top             =   825
      Width           =   6500
      Begin VB.CheckBox chkBooleano 
         Caption         =   "Falso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2580
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   1470
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   2595
         TabIndex        =   6
         Top             =   1305
         Visible         =   0   'False
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   56950785
         CurrentDate     =   37998
      End
      Begin VB.TextBox txtTexto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   225
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1245
         Visible         =   0   'False
         Width           =   6000
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   3
         Top             =   1305
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.ComboBox cboParametro 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   255
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   495
         Width           =   6000
      End
      Begin VB.Label lblValor 
         Caption         =   "&Valor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2955
         TabIndex        =   2
         Top             =   975
         Width           =   600
      End
      Begin VB.Label lblParametro 
         Caption         =   "&Parámetro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   255
         TabIndex        =   0
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Label Label2 
      Caption         =   " Descripción: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   12
      Top             =   2985
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "OPCIONES DE CONFIGURACIÓN PARÁMETROS GLOBALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   690
      Left            =   1470
      TabIndex        =   11
      Top             =   120
      Width           =   4770
   End
   Begin VB.Label lblDescripcion 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   600
      TabIndex        =   9
      Top             =   3210
      Width           =   6495
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA: OPCIONES
' Objetivo: PERMITE LA MODIFICACIÓN DE LOS PARÁMETROS DE
'               CONFIGURACIÓN DEL SISTEMA
' Programado por:
' Fecha: ENERO DE 2004
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsOpciones As ADODB.Recordset
    Dim AdoCmdRemplaza As ADODB.Command

Private Sub cboParametro_Click()
    strSQL = "SELECT * FROM parametros WHERE nombre_parametro = '" & cboParametro.Text & "'"
    Set AdoRcsOpciones = New ADODB.Recordset
    AdoRcsOpciones.ActiveConnection = Conn
    AdoRcsOpciones.LockType = adLockOptimistic
    AdoRcsOpciones.CursorType = adOpenKeyset
    AdoRcsOpciones.CursorLocation = adUseServer
    AdoRcsOpciones.Open strSQL
    If Not AdoRcsOpciones.EOF Then
        Select Case AdoRcsOpciones!tipo
            Case "N"
                txtNumero.Visible = True
                txtTexto.Visible = False
                chkBooleano.Visible = False
                dtpFecha.Visible = False
                txtNumero.Text = AdoRcsOpciones!Valor
            Case "T"
                txtNumero.Visible = False
                txtTexto.Visible = True
                chkBooleano.Visible = False
                dtpFecha.Visible = False
                txtTexto.Text = AdoRcsOpciones!Valor
            Case "B"
                txtNumero.Visible = False
                txtTexto.Visible = False
                chkBooleano.Visible = True
                dtpFecha.Visible = False
                If AdoRcsOpciones!Valor = "VERDADERO" Then
                    chkBooleano.Value = 1
                Else
                    chkBooleano.Value = 0
                End If
            Case "F"
                txtNumero.Visible = False
                txtTexto.Visible = False
                chkBooleano.Visible = False
                dtpFecha.Visible = True
                dtpFecha.Value = Format(AdoRcsOpciones!Valor, "dd/mm/yy")
        End Select
    End If
    lblDescripcion.Caption = cboParametro.Text & Chr$(13) & AdoRcsOpciones!Descripcion
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    On Error GoTo err_Guarda
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    strSQL = "UPDATE parametros SET valor = "
    Select Case AdoRcsOpciones!tipo
        Case "N"
            strSQL = strSQL & "'" & Trim(txtNumero.Text) & "'"
        Case "T"
            strSQL = strSQL & "'" & Trim(txtTexto.Text) & "'"
        Case "B"
            If chkBooleano.Value = 1 Then
                strSQL = strSQL & "'VERDADERO'"
            Else
                strSQL = strSQL & "'FALSO'"
            End If
        Case "F"
            strSQL = strSQL & "'" & Format(dtpFecha, "dd/mm/yy") & "'"
    End Select
    strSQL = strSQL & " WHERE nombre_parametro = '" & AdoRcsOpciones!nombre_parametro & "'"
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
     
    cmdGuardar.Enabled = False
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    Exit Sub
err_Guarda:
    Screen.MousePointer = Default
    cmdGuardar.Enabled = False
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub chkBooleano_Click()
    If chkBooleano.Value = 1 Then
        chkBooleano.Caption = "Verdadero"
    Else
        chkBooleano.Caption = "Falso"
    End If
End Sub

Private Sub chkBooleano_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdGuardar.Enabled = True
End Sub

Private Sub dtpFecha_Change()
        cmdGuardar.Enabled = True
End Sub

Private Sub Form_Load()
    Dim strCampo1, strCampo2 As String
    'Llena los combo Id Tipo Rentable
    strSQL = "SELECT nombre_parametro FROM parametros ORDER BY nombre_parametro"
    strCampo1 = "nombre_parametro"
    strCampo2 = ""
    Call LlenaCombos(cboParametro, strSQL, strCampo1, strCampo2)
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 22, 46, 48 To 57 'Backspace, <Ctrl+V>, punto y del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
    cmdGuardar.Enabled = True
End Sub

Private Sub txtTexto_KeyPress(KeyAscii As Integer)
    cmdGuardar.Enabled = True
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
