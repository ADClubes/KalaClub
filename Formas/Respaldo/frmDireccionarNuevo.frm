VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmDireccionarNuevo 
   Caption         =   "Direccionar pagos"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   Icon            =   "frmDireccionarNuevoNuevo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox txtCtrl 
      Height          =   285
      Index           =   5
      Left            =   240
      MaxLength       =   10
      TabIndex        =   22
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtCtrl 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   240
      MaxLength       =   4
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtCtrl 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   4200
      MaxLength       =   4
      TabIndex        =   17
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtCtrl 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtCtrl 
      Height          =   285
      Index           =   1
      Left            =   240
      MaxLength       =   60
      TabIndex        =   15
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtCtrl 
      Height          =   285
      Index           =   0
      Left            =   240
      MaxLength       =   60
      TabIndex        =   14
      Top             =   960
      Width           =   5175
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbEmisorTarjeta 
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   2415
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
      _ExtentX        =   4260
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbBancos 
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   2160
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
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.ComboBox cmbTipoDireccionado 
      Height          =   315
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2880
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtpFechaAlta 
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   57278465
      CurrentDate     =   38663
   End
   Begin VB.CommandButton cmdCancelar 
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      Picture         =   "frmDireccionarNuevoNuevo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   735
      Left            =   4440
      Picture         =   "frmDireccionarNuevo.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdGuarda 
      Enabled         =   0   'False
      Height          =   735
      Left            =   3240
      Picture         =   "frmDireccionarNuevo.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4560
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
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Tipo tarjeta"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Emitida (mmaa)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Banco emisor"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de cuenta"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblCVV 
      Alignment       =   2  'Center
      Caption         =   "Código seguridad"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblFecVen 
      Caption         =   "Vence (mmaa)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblCta 
      Caption         =   "Numero de tarjeta/CLABE"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblImporte 
      Caption         =   "Importe"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblFechaAlta 
      Caption         =   "Fecha de alta"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre Tarjetahabiente"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblTitular 
      Caption         =   "Titular"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
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



Const DATOSDIRS = 8

Public lTitular As Long
Public sTitular As String

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

    If Me.txtCtrl(0).Text = vbNullString Then
        MsgBox "Capture el nombre del tarjetahabiente", vbExclamation, "Verifique"
        Me.txtCtrl(0).SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(1).Text = vbNullString Then
        MsgBox "Capture el numero de tarjeta", vbExclamation, "Verifique"
        Me.txtCtrl(1).SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(3).Text = vbNullString Then
        MsgBox "Capture la fecha de vencimiento", vbExclamation, "Verifique"
        Me.txtCtrl(3).SetFocus
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
    
    If Me.cmbTipoDireccionado.Text = vbNullString Then
        MsgBox "Indique el tipo de cuenta", vbExclamation, "Verifique"
        Me.cmbTipoDireccionado.SetFocus
        Exit Function
    End If
    
    If Me.txtCtrl(5).Text = vbNullString Then
        MsgBox "Indique el importe", vbExclamation, "Verifique"
        Me.txtCtrl(5).SetFocus
        Exit Function
    End If
    
    
    
    
    If Len(Me.txtCtrl(1).Text) < 16 Then
        MsgBox "Verifique la cantidad de digitos del numero de tarjeta", vbExclamation, "Verifique"
        Me.txtCtrl(1).SetFocus
        Exit Function
    End If
    
    If Len(Me.txtCtrl(3).Text) < 4 Then
        MsgBox "Verifique la cantidad de digitos de la fecha de vencimiento", vbExclamation, "Verifique"
        Me.txtCtrl(3).SetFocus
        Exit Function
    End If
    
    
    
    If Val(Left(Me.txtCtrl(3).Text, 2)) < 0 Or Val(Left(Me.txtCtrl(3).Text, 2)) > 12 Then
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
    Dim adocmd As ADODB.Command
    
    
    Dim lNumReg As Long
    Dim sDatosCuenta As String
    Dim sTipoDirecc As String
    Dim sNumVisible As String
    
    
    If Not DatosCorrectos Then
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
    
    
    lNumReg = LeeUltReg("DIRECCIONADOS", "IdReg") + 1
    
    
    
    
    strSQL = ""
    strSQL = strSQL & "INSERT INTO DIRECCIONADOS ("
    strSQL = strSQL & " IdReg,"
    strSQL = strSQL & " IdMember,"
    strSQL = strSQL & " TipoDireccionado,"
    strSQL = strSQL & " Nombre,"
    strSQL = strSQL & " DatosCuenta,"
    strSQL = strSQL & " NumeroVisible,"
    strSQL = strSQL & " FechaExpedicion,"
    strSQL = strSQL & " FechaVencimiento,"
    strSQL = strSQL & " BancoEmisor,"
    strSQL = strSQL & " EmpresaEmisor,"
    strSQL = strSQL & " Importe,"
    strSQL = strSQL & " FechaAlta,"
    strSQL = strSQL & " HoraAlta,"
    strSQL = strSQL & " UsuarioAlta)"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & lNumReg & ","
    strSQL = strSQL & lTitular & ","
    strSQL = strSQL & "'" & sTipoDirecc & "',"
    strSQL = strSQL & "'" & Me.txtCtrl(0).Text & "',"
    strSQL = strSQL & "'" & sDatosCuenta & "',"
    strSQL = strSQL & "'" & sNumVisible & "',"
    If Me.txtCtrl(2).Text = vbNullString Then
        strSQL = strSQL & "'01/01/2000',"
    Else
        strSQL = strSQL & "'" & CDate("01/" & Left(Trim(Me.txtCtrl(2).Text), 2) & "/" & "20" & Right(Trim(Me.txtCtrl(2).Text), 2)) & "',"
    End If
    strSQL = strSQL & "'" & CDate("01/" & Left(Trim(Me.txtCtrl(3).Text), 2) & "/" & "20" & Right(Trim(Me.txtCtrl(3).Text), 2)) & "',"
    strSQL = strSQL & Me.ssCmbBancos.Columns("IdBanco").Value & ","
    strSQL = strSQL & Me.ssCmbEmisorTarjeta.Columns("IdTipo").Value & ","
    strSQL = strSQL & Me.txtCtrl(5).Text & ","
    strSQL = strSQL & "'" & Format(Date, "dd/mm/yyyy") & "',"
    strSQL = strSQL & "'" & Format(Now, "Hh:Nn") & "',"
    strSQL = strSQL & "'" & sDB_User & "')"
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    
    Set adocmd = Nothing
    
    
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Command1_Click()
'    Me.ssCmbBancos.Bookmark = Me.ss
End Sub

Private Sub Form_Activate()
    Me.txtCtrl(0).SetFocus
End Sub

Private Sub Form_Load()

    Me.dtpFechaAlta.Value = Date
    
    LlenaBancos
    LlenaTipos
    
    Me.cmbTipoDireccionado.AddItem "TARJETA DE CREDITO"
    Me.cmbTipoDireccionado.AddItem "TARJETA DE DEBITO"
    Me.cmbTipoDireccionado.AddItem "CLABE"
    
    
    LlenaDatos
    
    nPosIni = 0
    sTextMain = MDIPrincipal.StatusBar1.Panels.Item(1).Text
    
    Me.txtTitular.Text = sTitular
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Direccionar pagos"
    
    
    Me.cmdGuarda.Enabled = True
    
    
    
    
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
        Me.ssCmbBancos.Text = adorcs!Banco
        
        
        
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
        Case 0
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
