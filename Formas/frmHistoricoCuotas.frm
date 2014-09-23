VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHistoricoCuotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico de Cuotas"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8055
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   6840
      TabIndex        =   17
      Top             =   3720
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpVigenteHasta 
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   59047937
      CurrentDate     =   38665
   End
   Begin VB.CheckBox chkVigentes 
      Caption         =   "Ver sólo cuotas vigentes."
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   6480
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4920
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtPeriodo 
      Height          =   285
      Left            =   360
      MaxLength       =   2
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComCtl2.DTPicker dtpVigenteDesde 
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   59047937
      CurrentDate     =   38647
   End
   Begin VB.TextBox txtMontoDescuento 
      Height          =   285
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox txtMonto 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdInsertar 
      Caption         =   "Agregar"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbTipoUsuario 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5415
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
      Columns(0).Width=   6641
      Columns(0).Caption=   "Descripcion"
      Columns(0).Name =   "Descripcion"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1879
      Columns(1).Caption=   "IdTipoUsuario"
      Columns(1).Name =   "IdTipoUsuario"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   9551
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgCuotas 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   7695
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   7
      AllowUpdate     =   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   1244
      Columns(0).Caption=   "Periodo"
      Columns(0).Name =   "Periodo"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2302
      Columns(1).Caption=   "Monto"
      Columns(1).Name =   "Monto"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2302
      Columns(2).Caption=   "MontoDescuento"
      Columns(2).Name =   "MontoDescuento"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1984
      Columns(3).Caption=   "VigenteDesde"
      Columns(3).Name =   "VigenteDesde"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "VigenteHasta"
      Columns(4).Name =   "VigenteHasta"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "IdReg"
      Columns(5).Name =   "IdReg"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "Anterior"
      Columns(6).Name =   "Anterior"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      _ExtentX        =   13573
      _ExtentY        =   3201
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de usuario"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lblVigenteHasta 
      Caption         =   "Vigente Hasta"
      Height          =   255
      Left            =   6360
      TabIndex        =   16
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblPeriodo 
      Caption         =   "Periodo"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblVigenteDesde 
      Caption         =   "Vigente desde"
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblMontoDesc 
      Caption         =   "Monto Desc."
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblMonto 
      Caption         =   "Monto"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmHistoricoCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boEsNueva As Boolean

Private Sub chkVigentes_Click()
    GridUpdate
End Sub

Private Sub cmdAceptar_Click()
    
    
    Dim dFechaAnterior As Date
    Dim lRegAnt As Long
    
    Dim adocmdHistCuot As ADODB.Command
    Dim adorcsHistCuot As ADODB.Recordset
    Dim strSQL1 As String
    
    
    If Val(Me.txtPeriodo.Text) <= 0 Then
        MsgBox "El periodo debe ser mayor que 0!", vbExclamation, "Error"
        Me.txtPeriodo.SetFocus
        Exit Sub
    End If
    
    
    If Val(Me.txtPeriodo.Text) > 12 Then
        MsgBox "El periodo no puede ser mayor que 12", vbExclamation, "Error"
        Me.txtPeriodo.SetFocus
        Exit Sub
    End If
    
    
    Set adocmdHistCuot = New ADODB.Command
    adocmdHistCuot.ActiveConnection = Conn
    adocmdHistCuot.CommandType = adCmdText
    
    
    If boEsNueva Then
    
        
        'Busca que no exista un registro igual
        #If SqlServer_ Then
            strSQL = "SELECT IdReg"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " IdTipoUsuario=" & Me.sscmbTipoUsuario.Columns("IdTipoUsuario").Value
            strSQL = strSQL & " AND Periodo=" & Trim(Me.txtPeriodo.Text)
            strSQL = strSQL & " AND VigenteDesde=" & "'" & Format(Me.dtpVigenteDesde.Value, "yyyymmdd") & "'"
        #Else
            strSQL = "SELECT IdReg"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " IdTipoUsuario=" & Me.sscmbTipoUsuario.Columns("IdTipoUsuario").Value
            strSQL = strSQL & " AND Periodo=" & Trim(Me.txtPeriodo.Text)
            strSQL = strSQL & " AND VigenteDesde=" & "#" & Format(Me.dtpVigenteDesde.Value, "mm/dd/yyyy") & "#"
        #End If
        
        Set adorcsHistCuot = New ADODB.Recordset
        adorcsHistCuot.CursorLocation = adUseServer
        
        
        
        adorcsHistCuot.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        
        If Not adorcsHistCuot.EOF Then
            adorcsHistCuot.Close
            Set adorcsHistCuot = Nothing
            MsgBox "Ya existe un registro con las mismas características!", vbExclamation, "Error"
            Exit Sub
        End If
        
        #If SqlServer_ Then
            strSQL = "INSERT INTO HISTORICO_CUOTAS ("
            strSQL = strSQL & " IdTipoUsuario,"
            strSQL = strSQL & " Periodo,"
            strSQL = strSQL & " Monto,"
            strSQL = strSQL & " MontoDescuento,"
            strSQL = strSQL & " VigenteDesde,"
            strSQL = strSQL & " VigenteHasta,"
            strSQL = strSQL & " Anterior)"
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & Me.sscmbTipoUsuario.Columns("IdTipoUsuario").Value & ","
            strSQL = strSQL & Trim(Me.txtPeriodo.Text) & ","
            strSQL = strSQL & Trim(Me.txtMonto.Text) & ","
            strSQL = strSQL & Trim(Me.txtMontoDescuento.Text) & ","
            strSQL = strSQL & "'" & Format(Me.dtpVigenteDesde.Value, "yyyymmdd") & "',"
            strSQL = strSQL & "'" & "21001231" & "'"
            strSQL = strSQL & lRegAnt & ")"
        #Else
            strSQL = "INSERT INTO HISTORICO_CUOTAS ("
            strSQL = strSQL & " IdTipoUsuario,"
            strSQL = strSQL & " Periodo,"
            strSQL = strSQL & " Monto,"
            strSQL = strSQL & " MontoDescuento,"
            strSQL = strSQL & " VigenteDesde,"
            strSQL = strSQL & " VigenteHasta,"
            strSQL = strSQL & " Anterior)"
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & Me.sscmbTipoUsuario.Columns("IdTipoUsuario").Value & ","
            strSQL = strSQL & Trim(Me.txtPeriodo.Text) & ","
            strSQL = strSQL & Trim(Me.txtMonto.Text) & ","
            strSQL = strSQL & Trim(Me.txtMontoDescuento.Text) & ","
            strSQL = strSQL & "#" & Format(Me.dtpVigenteDesde.Value, "mm/dd/yyyy") & "#,"
            strSQL = strSQL & "#" & "12/31/2100" & "#"
            strSQL = strSQL & lRegAnt & ")"
        #End If
        
    Else
    
        dFechaAnterior = Me.ssdbgCuotas.Columns("VigenteDesde").Value
        
        'Si se modifica la fecha de inicio de vigencia
        'es necesario actualizar la fecha de VigenteHasta del
        'registro que le precede
        
        If dFechaAnterior <> Me.dtpVigenteDesde.Value And Val(Me.ssdbgCuotas.Columns("Anterior").Value) > 0 Then
            strSQL1 = "UPDATE HISTORICO_CUOTAS SET"
            strSQL1 = strSQL1 & " VigenteHasta=" & "'" & DateAdd("d", -1, Me.dtpVigenteDesde.Value) & "'"
            strSQL1 = strSQL1 & " WHERE"
            strSQL1 = strSQL1 & " IdReg=" & Me.ssdbgCuotas.Columns("Anterior").Value
        End If
    
        #If SqlServer_ Then
            strSQL = "UPDATE HISTORICO_CUOTAS SET"
            strSQL = strSQL & " Monto=" & Trim(Me.txtMonto.Text) & ","
            strSQL = strSQL & " MontoDescuento=" & Trim(Me.txtMontoDescuento.Text) & ","
            strSQL = strSQL & " VigenteDesde=" & "'" & Format(Me.dtpVigenteDesde.Value, "yyyymmdd") & "'"
            strSQL = strSQL & " WHERE IdReg=" & Me.ssdbgCuotas.Columns("IdReg").Value
        #Else
            strSQL = "UPDATE HISTORICO_CUOTAS SET"
            strSQL = strSQL & " Monto=" & Trim(Me.txtMonto.Text) & ","
            strSQL = strSQL & " MontoDescuento=" & Trim(Me.txtMontoDescuento.Text) & ","
            strSQL = strSQL & " VigenteDesde=" & "#" & Format(Me.dtpVigenteDesde.Value, "mm/dd/yyyy") & "#"
            strSQL = strSQL & " WHERE IdReg=" & Me.ssdbgCuotas.Columns("IdReg").Value
        #End If
        
        adocmdHistCuot.CommandText = strSQL
        adocmdHistCuot.Execute
        
        
        
    End If
    
    
    Set adocmdHistCuot = Nothing
    
    ActivaCtrls (False)
    GridUpdate
    
End Sub

Private Sub cmdCancelar_Click()
    ActivaCtrls False
End Sub

Private Sub cmdInsertar_Click()
    
    If Me.sscmbTipoUsuario.Text = "" Then
        MsgBox "Seleccionar un tipo de Usuario!", vbExclamation, "Error"
        Exit Sub
    End If
    
    
    ActivaCtrls True
    
    Me.txtPeriodo.Text = "1"
    Me.txtMonto.Text = ""
    Me.txtMontoDescuento.Text = ""
    Me.dtpVigenteDesde.Value = Date
    Me.dtpVigenteHasta.Value = Date
    
    boEsNueva = True
    
End Sub

Private Sub cmdModificar_Click()
        
    ActivaCtrls True
    
    Me.txtPeriodo.Text = Me.ssdbgCuotas.Columns("Periodo").Value
    Me.txtMonto.Text = Me.ssdbgCuotas.Columns("Monto").Value
    Me.txtMontoDescuento.Text = Me.ssdbgCuotas.Columns("MontoDescuento").Value
    Me.dtpVigenteDesde.Value = Me.ssdbgCuotas.Columns("VigenteDesde").Value
    
    
End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LlenaComboTipos
    
    CentraForma MDIPrincipal, Me
    
    boEsNueva = False
    
End Sub

Private Sub LlenaComboTipos()
    Dim adorcsTipos As ADODB.Recordset
    
    strSQL = "SELECT IdTipoUsuario, Descripcion"
    strSQL = strSQL & " FROM TIPO_USUARIO "
    strSQL = strSQL & " ORDER BY idTipoUsuario"
    
    Set adorcsTipos = New ADODB.Recordset
    adorcsTipos.CursorLocation = adUseServer
    
    adorcsTipos.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.sscmbTipoUsuario.RemoveAll
    
    Do Until adorcsTipos.EOF
        Me.sscmbTipoUsuario.AddItem adorcsTipos!Descripcion & vbTab & adorcsTipos!idtipousuario
        adorcsTipos.MoveNext
    Loop
    
    
    
    adorcsTipos.Close
    
    Set adorcsTipos = Nothing

End Sub



Private Sub GridUpdate()
    
    Dim adorcsHistCuo As ADODB.Recordset

    #If SqlServer_ Then
    '    strSQL = "SELECT IdReg, Periodo, Monto, MontoDescuento, VigenteDesde, VigenteHasta, Anterior"
        strSQL = "SELECT  Periodo, Monto, MontoDescuento, VigenteDesde, VigenteHasta"
        strSQL = strSQL & " FROM HISTORICO_CUOTAS"
        strSQL = strSQL & " WHERE IdTipoUsuario=" & Me.sscmbTipoUsuario.Columns("IdTipoUsuario").Value
        If Me.chkVigentes.Value = 1 Then
            strSQL = strSQL & " AND VigenteDesde <= " & "'" & Format(Date, "yyyymmdd") & "'"
            strSQL = strSQL & " AND VigenteHasta >= " & "'" & Format(Date, "yyyymmdd") & "'"
        End If
        strSQL = strSQL & " ORDER BY Periodo, VigenteDesde"
    #Else
    '    strSQL = "SELECT IdReg, Periodo, Monto, MontoDescuento, VigenteDesde, VigenteHasta, Anterior"
        strSQL = "SELECT  Periodo, Monto, MontoDescuento, VigenteDesde, VigenteHasta"
        strSQL = strSQL & " FROM HISTORICO_CUOTAS"
        strSQL = strSQL & " WHERE IdTipoUsuario=" & Me.sscmbTipoUsuario.Columns("IdTipoUsuario").Value
        If Me.chkVigentes.Value = 1 Then
            strSQL = strSQL & " AND VigenteDesde <= " & "#" & Format(Date, "mm/dd/yyyy") & "#"
            strSQL = strSQL & " AND VigenteHasta >= " & "#" & Format(Date, "mm/dd/yyyy") & "#"
        End If
        strSQL = strSQL & " ORDER BY Periodo, VigenteDesde"
    #End If
    
    Set adorcsHistCuo = New ADODB.Recordset
    adorcsHistCuo.CursorLocation = adUseServer
    
    adorcsHistCuo.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.ssdbgCuotas.RemoveAll
    
    Do Until adorcsHistCuo.EOF
        Me.ssdbgCuotas.AddItem adorcsHistCuo!Periodo & vbTab & adorcsHistCuo!Monto & vbTab & adorcsHistCuo!MontoDescuento & vbTab & adorcsHistCuo!VigenteDesde & vbTab & adorcsHistCuo!VigenteHasta & vbTab & adorcsHistCuo!IdReg & vbTab & adorcsHistCuo!Anterior
        adorcsHistCuo.MoveNext
    Loop
    
    
    adorcsHistCuo.Close
    Set adorcsHistCuo = Nothing

End Sub



Private Sub Label2_Click()

End Sub

Private Sub sscmbTipoUsuario_Click()
    GridUpdate
End Sub

Private Sub ActivaCtrls(boValue As Boolean)

    
    Me.lblPeriodo.Visible = boValue
    Me.lblMonto.Visible = boValue
    Me.lblMontoDesc.Visible = boValue
    Me.lblVigenteDesde.Visible = boValue
    Me.lblVigenteHasta.Visible = boValue
    
    

    Me.txtPeriodo.Visible = boValue
    Me.txtMonto.Visible = boValue
    Me.txtMontoDescuento.Visible = boValue
    Me.dtpVigenteDesde.Visible = boValue
    Me.dtpVigenteHasta.Visible = boValue
    
    Me.cmdAceptar.Visible = boValue
    Me.cmdCancelar.Visible = boValue
    

End Sub




Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtMonto.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtMontoDescuento_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtMontoDescuento.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub
