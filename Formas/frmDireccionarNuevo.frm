VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmDireccionarNuevo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Direccionar pagos"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Tarjetahabiente"
      Height          =   1215
      Left            =   240
      TabIndex        =   32
      Top             =   2880
      Width           =   6015
      Begin VB.TextBox txtCtrl 
         Height          =   405
         Index           =   7
         Left            =   3960
         MaxLength       =   60
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtCtrl 
         Height          =   405
         Index           =   6
         Left            =   2040
         MaxLength       =   60
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtCtrl 
         Height          =   405
         Index           =   0
         Left            =   120
         MaxLength       =   60
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Ap. Materno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4080
         TabIndex        =   35
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Ap. Paterno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSComCtl2.DTPicker dtpFechaAplica 
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   7440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   98631681
      CurrentDate     =   40273
   End
   Begin VB.TextBox txtTelefono 
      Height          =   375
      Left            =   240
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtEmail 
      Height          =   375
      Left            =   240
      MaxLength       =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   6015
   End
   Begin VB.TextBox txtCtrl 
      Height          =   405
      Index           =   5
      Left            =   360
      MaxLength       =   10
      TabIndex        =   13
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox txtCtrl 
      Alignment       =   2  'Center
      Height          =   405
      Index           =   4
      Left            =   360
      MaxLength       =   4
      TabIndex        =   10
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox txtCtrl 
      Alignment       =   2  'Center
      Height          =   405
      Index           =   3
      Left            =   3720
      MaxLength       =   4
      TabIndex        =   12
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtCtrl 
      Alignment       =   2  'Center
      Height          =   405
      Index           =   2
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   11
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtCtrl 
      Height          =   405
      Index           =   1
      Left            =   3720
      MaxLength       =   18
      TabIndex        =   9
      Top             =   5520
      Width           =   2535
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbEmisorTarjeta 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   5520
      Width           =   3015
      DataFieldList   =   "Column 0"
      AllowInput      =   0   'False
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   4868
      Columns(0).Caption=   "Tipo"
      Columns(0).Name =   "Tipo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "IdTipo"
      Columns(1).Name =   "IdTipo"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   5318
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbBancos 
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   4680
      Width           =   2535
      DataFieldList   =   "Column 0"
      AllowInput      =   0   'False
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   5715
      Columns(0).Caption=   "Banco"
      Columns(0).Name =   "Banco"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "IdBanco"
      Columns(1).Name =   "IdBanco"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.ComboBox cmbTipoDireccionado 
      Height          =   315
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4680
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker dtpFechaAlta 
      Height          =   405
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   98631681
      CurrentDate     =   38663
   End
   Begin VB.CommandButton cmdCancelar 
      Enabled         =   0   'False
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   735
      Left            =   4800
      Picture         =   "frmDireccionarNuevo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8160
      Width           =   735
   End
   Begin VB.CommandButton cmdGuarda 
      Enabled         =   0   'False
      Height          =   735
      Left            =   2880
      Picture         =   "frmDireccionarNuevo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8160
      Width           =   735
   End
   Begin VB.TextBox txtTitular 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label lblAplica 
      Caption         =   "Aplica a partir"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   31
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label lblTelefono 
      Caption         =   "Teléfono"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lblEmail 
      Caption         =   "Correo electrónico"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Tipo tarjeta"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Emitida (mmaa)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Banco emisor"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   26
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de cuenta"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label lblCVV 
      Alignment       =   2  'Center
      Caption         =   "Código seguridad"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label lblFecVen 
      Caption         =   "Vence (mmaa)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label lblCta 
      Caption         =   "Numero de tarjeta/CLABE"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   5280
      Width           =   2295
   End
   Begin VB.Label lblImporte 
      Caption         =   "Importe"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label lblFechaAlta 
      Caption         =   "Fecha de alta"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   20
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label lblTitular 
      Caption         =   "Titular"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmDireccionarNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de pagos direccionados          *
'*  Daniel Hdez                                                 *
'*  07 / Noviembre / 2004                                       *
'*  Ultima actualización: 14 / Noviembre / 2005                 *
'****************************************************************
Option Explicit

Const DATOSDIRS = 8

Public lTitular As Long
Public sTitular As String

Public lNoInscripcion As Long

Dim mAncDir(DATOSDIRS) As Integer
Dim mEncDir(DATOSDIRS) As String
Dim sNombre As String
Dim nImporte As Double
Dim nActivo As Byte
Dim dFechaAlta As Date
Dim bTarjCredito As Boolean
Dim bNuevo As Boolean
Dim nPosIni As Variant
Dim sTextMain As String
Dim frmHDir As frmayuda
Dim dImporteMant As Double

Private Function Cambios() As Boolean
    Cambios = True
    
'    If (sNombre <> Trim$(Me.txtNombre.Text)) Then
'        Exit Function
'    End If
    
'    If (nImporte <> Val(Me.txtImporte.Text)) Then
'        Exit Function
'    End If
    
'    If (nActivo <> Me.chkActivo.Value) Then
'        Exit Function
'    End If
    
    If (dFechaAlta <> Me.dtpFechaAlta.Value) Then
        Exit Function
    End If
    
'    If (bTarjCredito <> Me.optTC.Value) Then
'        Exit Function
'    End If
    
    Cambios = False
End Function

Private Function DatosCorrectos() As Boolean
    DatosCorrectos = False
    Dim sError As String

    If Me.txtEmail.Text = vbNullString Then
        MsgBox "Capture un correo electrónico", vbExclamation, "Verifique"
        Me.txtEmail.SetFocus
        Exit Function
    End If
    
    If Me.txtTelefono.Text = vbNullString Then
        MsgBox "Capture un teléfono", vbExclamation, "Verifique"
        Me.txtTelefono.SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(0).Text = vbNullString Then
        MsgBox "Capture el nombre del tarjetahabiente", vbExclamation, "Verifique"
        Me.txtCtrl(0).SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(6).Text = vbNullString Then
        MsgBox "Capture el Apellido Paterno del tarjetahabiente", vbExclamation, "Verifique"
        Me.txtCtrl(6).SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(7).Text = vbNullString Then
        MsgBox "Capture el Apellido Materno del tarjetahabiente" & vbCrLf & "si no tiene, capture una letra X", vbExclamation, "Verifique"
        Me.txtCtrl(7).SetFocus
        Exit Function
    End If
    
    If Me.cmbTipoDireccionado.Text = vbNullString Then
        MsgBox "Indique el tipo de cuenta", vbExclamation, "Verifique"
        Me.cmbTipoDireccionado.SetFocus
        Exit Function
    End If
    
    If Me.ssCmbBancos.Text = vbNullString Then
        MsgBox "Indique el banco", vbExclamation, "Verifique"
        Me.ssCmbBancos.SetFocus
        Exit Function
    End If
    
    If Me.ssCmbEmisorTarjeta.Text = vbNullString Then
        MsgBox "Indique el tipo de tarjeta", vbExclamation, "Verifique"
        Me.ssCmbEmisorTarjeta.SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(1).Text = vbNullString Then
        MsgBox "Capture el numero de tarjeta", vbExclamation, "Verifique"
        Me.txtCtrl(1).SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(4).Text = vbNullString Then
        MsgBox "Capture el código de seguridad", vbExclamation, "Verifique"
        Me.txtCtrl(4).SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(3).Text = vbNullString Then
        MsgBox "Capture la fecha de vencimiento", vbExclamation, "Verifique"
        Me.txtCtrl(3).SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(5).Text = vbNullString Then
        MsgBox "Indique el importe", vbExclamation, "Verifique"
        Me.txtCtrl(5).SetFocus
        Exit Function
    End If
    
    
    If Not ValidaEmailAddress(Trim(Me.txtEmail.Text)) Then
        MsgBox "Verifique el correo electrónico", vbExclamation, "Verifique"
        Me.txtEmail.SetFocus
        Exit Function
    End If
    
    
    If Len(Me.txtCtrl(1).Text) < 16 Then
        If Me.ssCmbBancos.Text = "AMEX" Then
            If Len(Me.txtCtrl(1).Text) < 15 Then
                MsgBox "Verifique la cantidad de digitos del numero de tarjeta", vbExclamation, "Verifique"
                Me.txtCtrl(1).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If cmbTipoDireccionado.Text = "CLABE" And Len(Me.txtCtrl(1).Text) < 18 Then
        MsgBox "Verifique la cantidad de digitos de la cuenta CLABE", vbExclamation, "Verifique"
        Me.txtCtrl(1).SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(2).Text <> vbNullString And Len(Me.txtCtrl(2).Text) < 4 Then
        MsgBox "Verifique la cantidad de digitos de la fecha de Emisión", vbExclamation, "Verifique"
        Me.txtCtrl(2).SetFocus
        Exit Function
    End If
    
    If (Me.txtCtrl(2).Text <> vbNullString) And (Val(Left(Me.txtCtrl(2).Text, 2)) < 1 Or Val(Left(Me.txtCtrl(2).Text, 2)) > 12) Then
        MsgBox "El mes de la fecha de emisión es incorrecto", vbExclamation, "Verifique"
        Me.txtCtrl(2).SetFocus
        Exit Function
    End If
    
    If (Me.txtCtrl(2).Text <> vbNullString) And (Val(Right(Me.txtCtrl(2).Text, 2)) + 2000 > Year(Date)) Then
        MsgBox "El año de la fecha de emision es mayor al año actual", vbExclamation, "Verifique"
        Me.txtCtrl(2).SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(2).Text <> vbNullString And (Val(Right(Me.txtCtrl(2).Text, 2)) + 2000) * 12 + Val(Left(Me.txtCtrl(2).Text, 2)) >= Year(Date) * 12 + Month(Date) Then
        MsgBox "La fecha de emision debe ser igual o menor al mes actual", vbExclamation, "Verifique"
        Me.txtCtrl(2).SetFocus
        Exit Function
    End If
    
    If Len(Me.txtCtrl(3).Text) < 4 Then
        MsgBox "Verifique la cantidad de digitos de la fecha de vencimiento", vbExclamation, "Verifique"
        Me.txtCtrl(3).SetFocus
        Exit Function
    End If
    
    If Val(Left(Me.txtCtrl(3).Text, 2)) < 1 Or Val(Left(Me.txtCtrl(3).Text, 2)) > 12 Then
        MsgBox "El mes de la fecha de vencimiento es incorrecto", vbExclamation, "Verifique"
        Me.txtCtrl(3).SetFocus
        Exit Function
    End If
    
    If Val(Right(Me.txtCtrl(3).Text, 2)) + 2000 < Year(Date) Then
        MsgBox "El año de la fecha de vencimiento es menor al año actual", vbExclamation, "Verifique"
        Me.txtCtrl(3).SetFocus
        Exit Function
    End If
    
    If (Val(Right(Me.txtCtrl(3).Text, 2)) + 2000) * 12 + Val(Left(Me.txtCtrl(3).Text, 2)) <= Year(Date) * 12 + Month(Date) Then
        MsgBox "La fecha de vencimiento debe ser mayor", vbExclamation, "Verifique"
        Me.txtCtrl(3).SetFocus
        Exit Function
    End If
    
    If Val(Me.txtCtrl(5).Text) = 0 Then
        MsgBox "El importe debe ser mayor a cero", vbExclamation, "Verifique"
        Me.txtCtrl(5).SetFocus
        Exit Function
    End If
    
    If Me.dtpFechaAplica.Value < Me.dtpFechaAlta.Value Then
        MsgBox "La fecha de aplicación debe ser mayor que la fecha de alta", vbExclamation, "Verifique"
        Me.dtpFechaAplica.SetFocus
        Exit Function
    End If
    
    If Day(Me.dtpFechaAplica.Value) <> 1 Then
        MsgBox "La fecha de aplicación siempre debe ser el dia 1 de cada mes", vbExclamation, "Verifique"
        Me.dtpFechaAplica.SetFocus
        Exit Function
    End If
    
    If DateSerial(Val(Right(Me.txtCtrl(3).Text, 2)) + 2000, Val(Left(Me.txtCtrl(3).Text, 2)), 1) < Me.dtpFechaAplica.Value Then
        MsgBox "La fecha de vencimiento deber ser mayor que la fecha de aplicación!", vbExclamation, "Verifique"
        Me.dtpFechaAplica.SetFocus
        Exit Function
    End If
    
    If Me.cmbTipoDireccionado.Text <> "CLABE" Then
        If Not ValidateCardNumber(Me.txtCtrl(1).Text, sError) Then
            MsgBox "Número de tarjeta incorrecto", vbInformation, "Verifique"
            Exit Function
        End If
    End If
    
    DatosCorrectos = True
End Function

Private Sub cmdBorrar_Click()
Dim nOk As Long

'    If (Me.ssdbDireccionados.Rows > 0) Then
'        nOk = MsgBox("¿Desea borrar al vendedor seleccionado?", vbYesNo, "KalaSystems")
'
'        If (nOk = vbYes) Then
'            If (EliminaReg("Direccionados", "idReg=" & Me.ssdbDireccionados.Columns(7).Value, "", Conn)) Then
'                LlenaDirs
'            Else
'                MsgBox "No se eliminó el registro.", vbInformation, "KalaSystems"
'            End If
'        End If
'    End If
End Sub

Private Sub cmdCancelar_Click()
    ClrCtrls
    ActivaCtrls False
    CtrlsEdicion True
    
'    If (Me.ssdbDireccionados.Rows > 0) Then
'        Me.ssdbDireccionados.SetFocus
'    End If
End Sub

Private Sub cmdModificar_Click()
'    If (Me.ssdbDireccionados.Rows > 0) Then
'        bNuevo = False
'        nPosIni = Me.ssdbDireccionados.Bookmark
'
'        LeeDatos
'
'        CtrlsEdicion False
'        ActivaCtrls True
'        InitVars
'
'        Me.txtNombre.SetFocus
'    End If
End Sub

Private Sub cmdNuevo_Click()
    bNuevo = True

    CtrlsEdicion False
    ActivaCtrls True
    ClrCtrls
    
    InitVars

'    Me.txtNombre.SetFocus
End Sub

Private Sub cmdGuarda_Click()
    
'    Dim sqlConn As ADODB.Connection
    
    Dim adocmd As ADODB.Command
    Dim adocmdSQL As ADODB.Command
    
    Dim lNumReg As Long
    Dim sDatosCuenta As String
    Dim sTipoDirecc As String
    Dim sNumVisible As String
    
'    Dim strConn As String
    Dim lRecords As Long
    
    Dim iResp As Integer
    
'    strConn = "Provider=SQLOLEDB.1;Password=password;Persist Security Info=True;User ID=sa;Initial Catalog=KalaClubSQL;Data Source=172.16.2.1"
    
    If Not DatosCorrectos Then
        Exit Sub
    End If
    
    If Not MsgBox("Datos Correctos, ¿proceder con el alta?", vbQuestion + vbOKCancel, "Confirme") = vbOK Then
        Exit Sub
    End If
    
    sDatosCuenta = Trim(Me.txtCtrl(1)) + "-" + Trim(Me.txtCtrl(3)) + "-" + Trim(Me.txtCtrl(4))
    
    If Me.cmbTipoDireccionado.ListIndex = 0 Then
        sTipoDirecc = "TC"
    ElseIf Me.cmbTipoDireccionado.ListIndex = 1 Then
        sTipoDirecc = "TD"
    Else
        sTipoDirecc = "CI"
    End If
    
    sNumVisible = Right$(Me.txtCtrl(1).Text, 4)
    
    'Set sqlConn = New ADODB.Connection
    'sqlConn.Open strConn
    
    'Set adocmdSQL = New ADODB.Command
    
    'adocmdSQL.CommandType = adCmdStoredProc
    
    'adocmdSQL.ActiveConnection = sqlConn
    'adocmdSQL.CommandText = "usp_DireccionadosDatos_Insert"
    'adocmdSQL.Parameters.REFRESH
    
    'adocmdSQL.Parameters(2).Value = CDate(Me.dtpFechaAlta.Value)
    'adocmdSQL.Parameters(3).Value = lNoInscripcion
    'adocmdSQL.Parameters(4).Value = lTitular
    'adocmdSQL.Parameters(5).Value = sTipoDirecc
    'adocmdSQL.Parameters(6).Value = Me.txtCtrl(1).Text
    'adocmdSQL.Parameters(7).Value = Right(Me.txtCtrl(1).Text, 1) & "*******" & Left(Me.txtCtrl(1).Text, 4)
    'adocmdSQL.Parameters(8).Value = Me.ssCmbBancos.Text
    'adocmdSQL.Parameters(9).Value = Me.txtCtrl(0)
    'adocmdSQL.Parameters(10).Value = Me.txtCtrl(2)
    'adocmdSQL.Parameters(11).Value = Me.txtCtrl(3)
    'adocmdSQL.Parameters(12).Value = Me.txtCtrl(4)
    'adocmdSQL.Parameters(13).Value = Me.ssCmbEmisorTarjeta.Text
    'adocmdSQL.Parameters(14).Value = Me.txtCtrl(5)
    'adocmdSQL.Parameters(15).Value = 1
    'adocmdSQL.Parameters(16).Value = Me.txtEmail.Text
    'adocmdSQL.Parameters(17).Value = Me.txtTelefono.Text
    'adocmdSQL.Parameters(18).Value = CDate(Me.dtpFechaAplica.Value)
    'adocmdSQL.Parameters(19).Value = sDB_User
    'adocmdSQL.Parameters(20).Value = Null
    'adocmdSQL.Parameters(21).Value = Null
    
    'adocmdSQL.Execute lRecords
    
    lNumReg = LeeUltReg("DIRECCIONADOS", "IdReg") + 1
    
    #If SqlServer_ Then
        strSQL = ""
        strSQL = strSQL & "INSERT INTO DIRECCIONADOS ("
        strSQL = strSQL & " IdReg,"
        strSQL = strSQL & " IdMember,"
        strSQL = strSQL & " FechaAlta,"
        strSQL = strSQL & " TipoDireccionado,"
        strSQL = strSQL & " Nombre,"
        strSQL = strSQL & " Importe,"
        strSQL = strSQL & " Activo)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lNumReg & ","
        strSQL = strSQL & lTitular & ","
        strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & sTipoDirecc & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(0).Text & " " & Me.txtCtrl(6).Text & " " & Me.txtCtrl(7).Text & "',"
        strSQL = strSQL & Me.txtCtrl(5).Text & ","
        strSQL = strSQL & -1 & ")"
    #Else
        strSQL = ""
        strSQL = strSQL & "INSERT INTO DIRECCIONADOS ("
        strSQL = strSQL & " IdReg,"
        strSQL = strSQL & " IdMember,"
        strSQL = strSQL & " FechaAlta,"
        strSQL = strSQL & " TipoDireccionado,"
        strSQL = strSQL & " Nombre,"
        strSQL = strSQL & " Importe,"
        strSQL = strSQL & " Activo)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lNumReg & ","
        strSQL = strSQL & lTitular & ","
        strSQL = strSQL & "#" & Format(Date, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "'" & sTipoDirecc & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(0).Text & " " & Me.txtCtrl(6).Text & " " & Me.txtCtrl(7).Text & "',"
        strSQL = strSQL & Me.txtCtrl(5).Text & ","
        strSQL = strSQL & -1 & ")"
    #End If
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO DireccionadosDatos ("
        strSQL = strSQL & "IdConvenio, "
        strSQL = strSQL & "FechaAlta, "
        strSQL = strSQL & "NoInscripcion, "
        strSQL = strSQL & "IdMember, "
        strSQL = strSQL & "TipoDireccionado, "
        strSQL = strSQL & "NumeroCuenta, "
        strSQL = strSQL & "NumeroCuentaVisible, "
        strSQL = strSQL & "BancoEmisor, "
        strSQL = strSQL & "Nombre, "
        strSQL = strSQL & "A_Paterno, "
        strSQL = strSQL & "A_Materno, "
        strSQL = strSQL & "FechaExpedicion, "
        strSQL = strSQL & "FechaVencimiento, "
        strSQL = strSQL & "CodigoSeguridad, "
        strSQL = strSQL & "OperadorTarjeta, "
        strSQL = strSQL & "Importe, "
        strSQL = strSQL & "Activo, "
        strSQL = strSQL & "Email, "
        strSQL = strSQL & "Telefono, "
        strSQL = strSQL & "AplicaAPartir, "
        strSQL = strSQL & "UsuarioAlta) "
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & lNumReg & ","
        strSQL = strSQL & "'" & Format(Me.dtpFechaAlta.Value, "yyyymmdd") & "',"
        strSQL = strSQL & lNoInscripcion & ","
        strSQL = strSQL & lTitular & ","
        strSQL = strSQL & "'" & sTipoDirecc & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(1).Text & "',"
        strSQL = strSQL & "'" & Right(Me.txtCtrl(1).Text, 4) & "',"
        strSQL = strSQL & "'" & Me.ssCmbBancos.Text & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(0) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(6) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(7) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(2) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(3) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(4) & "',"
        strSQL = strSQL & "'" & Me.ssCmbEmisorTarjeta.Text & "',"
        strSQL = strSQL & Val(Me.txtCtrl(5)) & ","
        strSQL = strSQL & -1 & ","
        strSQL = strSQL & "'" & LCase(Me.txtEmail.Text) & "',"
        strSQL = strSQL & "'" & Me.txtTelefono.Text & "',"
        strSQL = strSQL & "'" & Format(Me.dtpFechaAplica.Value, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & sDB_User & "')"
    #Else
        strSQL = "INSERT INTO DireccionadosDatos ("
        strSQL = strSQL & "IdConvenio, "
        strSQL = strSQL & "FechaAlta, "
        strSQL = strSQL & "NoInscripcion, "
        strSQL = strSQL & "IdMember, "
        strSQL = strSQL & "TipoDireccionado, "
        strSQL = strSQL & "NumeroCuenta, "
        strSQL = strSQL & "NumeroCuentaVisible, "
        strSQL = strSQL & "BancoEmisor, "
        strSQL = strSQL & "Nombre, "
        strSQL = strSQL & "A_Paterno, "
        strSQL = strSQL & "A_Materno, "
        strSQL = strSQL & "FechaExpedicion, "
        strSQL = strSQL & "FechaVencimiento, "
        strSQL = strSQL & "CodigoSeguridad, "
        strSQL = strSQL & "OperadorTarjeta, "
        strSQL = strSQL & "Importe, "
        strSQL = strSQL & "Activo, "
        strSQL = strSQL & "Email, "
        strSQL = strSQL & "Telefono, "
        strSQL = strSQL & "AplicaAPartir, "
        strSQL = strSQL & "UsuarioAlta) "
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & lNumReg & ","
        strSQL = strSQL & "#" & Format(Me.dtpFechaAlta.Value, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & lNoInscripcion & ","
        strSQL = strSQL & lTitular & ","
        strSQL = strSQL & "'" & sTipoDirecc & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(1).Text & "',"
        strSQL = strSQL & "'" & Right(Me.txtCtrl(1).Text, 4) & "',"
        strSQL = strSQL & "'" & Me.ssCmbBancos.Text & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(0) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(6) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(7) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(2) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(3) & "',"
        strSQL = strSQL & "'" & Me.txtCtrl(4) & "',"
        strSQL = strSQL & "'" & Me.ssCmbEmisorTarjeta.Text & "',"
        strSQL = strSQL & Val(Me.txtCtrl(5)) & ","
        strSQL = strSQL & -1 & ","
        strSQL = strSQL & "'" & LCase(Me.txtEmail.Text) & "',"
        strSQL = strSQL & "'" & Me.txtTelefono.Text & "',"
        strSQL = strSQL & "#" & Format(Me.dtpFechaAplica.Value, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "'" & sDB_User & "')"
    #End If
    
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    Set adocmd = Nothing
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.txtEmail.SetFocus
End Sub

Private Sub Form_Load()
    Dim sFechaInicio As String
    
    Me.dtpFechaAlta.Value = Date
    
    LlenaBancos
    LlenaTipos
    
    Me.cmbTipoDireccionado.AddItem "TARJETA DE CREDITO"
    Me.cmbTipoDireccionado.AddItem "TARJETA DE DEBITO"
    Me.cmbTipoDireccionado.AddItem "CLABE"
    
    sFechaInicio = ObtieneParametro("FECHA_INICIO_CONVENIO")
    
    If sFechaInicio = vbNullString Then
        Me.dtpFechaAplica.Value = UltimoDiaDelMes(Date) + 1
    Else
        Me.dtpFechaAplica.Value = CDate(sFechaInicio)
    End If
    
'    If Month(Date) = 1 Then
'        Me.txtCtrl(2).Text = Format(DateSerial(Year(Date) - 1, 12, 1), "mmyy")
'    Else
'        Me.txtCtrl(2).Text = Format(DateSerial(Year(Date), Month(Date) - 1, 1), "mmyy")
'    End If
    
    'LlenaDatos
    
    nPosIni = 0
    sTextMain = MDIPrincipal.StatusBar1.Panels.Item(1).Text
    
    Me.txtTitular.Text = sTitular
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Direccionar pagos"
    
    Me.cmdGuarda.Enabled = True
    
    dImporteMant = CalculaMantenimientoMes(CLng(lTitular), True, 1)
    
    Me.txtCtrl(5).Text = Format(dImporteMant, "#0.00")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmDireccionarNuevo = Nothing
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextMain
End Sub

Private Sub ActivaCtrls(bValor As Boolean)
    With Me
'        .txtNombre.Enabled = bValor
'        .txtImporte.Enabled = bValor
'        .txtNoTipo.Enabled = bValor
'        .cmdHTipoDir.Enabled = bValor
'        .chkActivo.Enabled = bValor
        .dtpFechaAlta.Enabled = bValor
'        .frmTipoDireccionado.Enabled = bValor
    End With
End Sub

Private Sub CtrlsEdicion(bValor As Boolean)
    With Me
'        .ssdbDireccionados.Enabled = bValor
'        .cmdNuevo.Enabled = bValor
'        .cmdModificar.Enabled = bValor
'        .cmdBorrar.Enabled = bValor
        
'        .cmdAgregar.Enabled = Not bValor
        .cmdCancelar.Enabled = Not bValor
    End With
End Sub

Private Sub ClrCtrls()
    With Me
'        .txtNombre.Text = Me.txtTitular.Text
'        .txtImporte.Text = ""
'        .chkActivo.Value = 0
        .dtpFechaAlta.Value = Date
'        .optTC.Value = True
    End With
End Sub

Private Sub InitVars()
    'sNombre = Trim$(Me.txtNombre.Text)
    
    nImporte = 0
'    If (Val(Me.txtImporte.Text) > 0) Then
'        nImporte = CDbl(Me.txtImporte.Text)
'    End If
    
    'nActivo = Me.chkActivo.Value
    dFechaAlta = Me.dtpFechaAlta.Value
    'bTarjCredito = Me.optTC.Value
End Sub

Private Sub LeeDatos()
'    With Me.ssdbDireccionados
'        Me.txtNombre.Text = .Columns(0).Text
'        Me.txtImporte.Text = .Columns(3).Value
'        Me.chkActivo.Value = Val(.Columns(1).Value)
'        Me.dtpFechaAlta.Value = .Columns(2).Value
'        Me.optTC.Value = IIf(.Columns(4).Text = "TC", True, False)
'        Me.optTD.Value = Not Me.optTC.Value
'    End With
End Sub

Private Sub LlenaDatos()
    Dim adorcs As ADODB.Recordset
    
    strSQL = ""
    strSQL = "SELECT D.TipoDireccionado, D.Nombre, D.NumeroVisible, D.FechaExpedicion, D.FechaVencimiento, D.BancoEmisor, D.EmpresaEmisor, D.Importe, D.FechaAlta, B.Banco, T.NombreEmisor"
    strSQL = strSQL & " FROM (DIRECCIONADOS D LEFT JOIN CT_BANCOS B ON D.BancoEmisor=B.IdBanco)"
    strSQL = strSQL & " LEFT JOIN CT_EMISOR_TARJETA T ON D.EmpresaEmisor = T.IdEmisor "
    strSQL = strSQL & " WHERE (((D.IdMember)=" & lTitular & "))"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        
        Me.txtCtrl(0).Text = adorcs!Nombre
        Me.txtCtrl(1).Text = "XXXX-XXXX-XXXX-" & adorcs!NumeroVisible
        Me.txtCtrl(2).Text = IIf(IsNull(adorcs!FechaExpedicion), "0100", Format(adorcs!FechaExpedicion, "mmyy"))
        Me.txtCtrl(3).Text = IIf(IsNull(adorcs!FechaVencimiento), "", Format(adorcs!FechaVencimiento, "mmyy"))
        Me.txtCtrl(4).Text = "XXX"
        Me.txtCtrl(5).Text = adorcs!Importe
        Me.dtpFechaAlta.Value = adorcs!FechaAlta
        Me.ssCmbBancos.Text = adorcs!banco
        
        Me.ssCmbEmisorTarjeta.Text = adorcs!NombreEmisor
        
        If adorcs!TipoDireccionado = "TC" Then Me.cmbTipoDireccionado.ListIndex = 0
        If adorcs!TipoDireccionado = "TD" Then Me.cmbTipoDireccionado.ListIndex = 1
        If adorcs!TipoDireccionado = "CI" Then Me.cmbTipoDireccionado.ListIndex = 2
    
    Else
    
        Me.txtCtrl(0).Enabled = True
        Me.txtCtrl(0).Text = sTitular
                
        Me.txtCtrl(1).Enabled = True
        Me.txtCtrl(2).Enabled = True
        Me.txtCtrl(3).Enabled = True
        Me.txtCtrl(4).Enabled = True
        Me.txtCtrl(5).Enabled = True
        
    End If
    
    adorcs.Close
    Set adorcs = Nothing
End Sub

Private Sub LlenaBancos()
    strSQL = ""
'    strSQL = "SELECT B.Banco, B.IdBanco"
'    strSQL = strSQL & " FROM CT_BANCOS B"
'    strSQL = strSQL & " ORDER BY B.Banco"
    strSQL = "SELECT B.Banco, B.IdBanco"
    strSQL = strSQL & " FROM CT_BANCOS B"
    strSQL = strSQL & " ORDER BY B.Banco"
    
    LlenaSsCombo Me.ssCmbBancos, Conn, strSQL, 2
End Sub

Private Sub LlenaTipos()
    strSQL = ""
    strSQL = "SELECT T.NombreEmisor, T.IdEmisor"
    strSQL = strSQL & " FROM CT_EMISOR_TARJETA T"
    strSQL = strSQL & " ORDER BY T.NombreEmisor"
    
    LlenaSsCombo Me.ssCmbEmisorTarjeta, Conn, strSQL, 2
End Sub

Private Sub txtCtrl_GotFocus(Index As Integer)
    Me.txtCtrl(Index).SelStart = 0
    Me.txtCtrl(Index).SelLength = Len(Me.txtCtrl(Index).Text)
End Sub

Private Sub txtCtrl_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 6, 7
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case 5
            Select Case KeyAscii
                Case 8 ' Tecla backspace
                    KeyAscii = KeyAscii
                Case 46 'punto decimal
                    If InStr(1, Me.txtCtrl(Index).Text, ".", 1) > 0 Then
                        KeyAscii = 0
                    End If
                    
                    KeyAscii = KeyAscii
                Case 48 To 57 ' Numeros del 0 al 9
                    KeyAscii = KeyAscii
                Case Else
                    KeyAscii = 0
                    'MsgBox "Solo admite números", vbInformation, "Ver"
            End Select
        Case 1, 2, 3, 4
            Select Case KeyAscii
                Case 8 ' Tecla backspace
                    KeyAscii = KeyAscii
                Case 48 To 57 ' Numeros del 0 al 9
                    KeyAscii = KeyAscii
                Case Else
                    KeyAscii = 0
                    'MsgBox "Solo admite números", vbInformation, "Productos"
            End Select
    End Select
End Sub
