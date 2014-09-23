VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmFechaPago 
   Caption         =   "Fecha de pago"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CheckBox chkTodos 
      Caption         =   "Actualiza a todos"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdCambia 
      Caption         =   "Actualiza"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpFechaPago 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   58851329
      CurrentDate     =   39288
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgFechasPago 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   4
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   1958
      Columns(0).Caption=   "NoFamilia"
      Columns(0).Name =   "NoFamilia"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   6218
      Columns(1).Caption=   "Nombre"
      Columns(1).Name =   "Nombre"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2064
      Columns(2).Caption=   "Fecha"
      Columns(2).Name =   "Fecha"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1535
      Columns(3).Caption=   "IdMember"
      Columns(3).Name =   "IdMember"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   11668
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
End
Attribute VB_Name = "frmFechaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'20/05/08
Public lNoFamilia As Long
Private Sub LlenaGrid()
    Dim adorcsFechas As ADODB.Recordset
    
    Set adorcsFechas = New ADODB.Recordset
    adorcsFechas.CursorLocation = adUseServer
    
    strSQL = "SELECT USUARIOS_CLUB.NoFamilia, FECHAS_USUARIO.IdConcepto, FECHAS_USUARIO.FechaUltimoPago, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.Nombre, FECHAS_USUARIO.IdMember"
    strSQL = strSQL & " FROM FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember = USUARIOS_CLUB.IdMember"
    strSQL = strSQL & " Where (((USUARIOS_CLUB.NoFamilia) =" & lNoFamilia & "))"
    strSQL = strSQL & " ORDER BY FECHAS_USUARIO.IdMember"
    
    adorcsFechas.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.ssdbgFechasPago.RemoveAll
    
    Do Until adorcsFechas.EOF
        Me.ssdbgFechasPago.AddItem adorcsFechas!NoFamilia & vbTab & adorcsFechas!A_Paterno & " " & adorcsFechas!A_Materno & " " & adorcsFechas!Nombre & vbTab & adorcsFechas!Fechaultimopago & vbTab & adorcsFechas!Idmember
        adorcsFechas.MoveNext
    Loop
    
    adorcsFechas.Close
    Set adorcsFechas = Nothing
    
    If Me.ssdbgFechasPago.Rows > 0 Then
        ActivaControles True
    Else
        ActivaControles False
    End If
    
    
    
End Sub

Private Sub cmdBusca_Click()
    LlenaGrid
End Sub

Private Sub cmdCambia_Click()
    Dim adocmdFechas As ADODB.Command
    
    If Day(CDate(Me.dtpFechaPago.Value)) <> Day(UltimoDiaDelMes(CDate(Me.dtpFechaPago.Value))) Then
        MsgBox "La última fecha de pago debe ser" & vbCrLf & "igual al último día del mes correspondiente", vbExclamation, "Error"
        Me.dtpFechaPago.Value = UltimoDiaDelMes(CDate(Me.dtpFechaPago.Value))
        Me.dtpFechaPago.SetFocus
        Exit Sub
    End If
    
    If Me.chkTodos.Value Then
        #If SqlServer_ Then
            strSQL = "UPDATE FECHAS_USUARIO"
            strSQL = strSQL & " SET FechaUltimoPago='" & Format(Me.dtpFechaPago.Value, "yyyymmdd") & "'"
            strSQL = strSQL & " FROM FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember"
            strSQL = strSQL & " WHERE USUARIOS_CLUB.NoFamilia=" & lNoFamilia
        #Else
            strSQL = "UPDATE FECHAS_USUARIO"
            strSQL = strSQL & " SET FechaUltimoPago='" & Me.dtpFechaPago.Value & "'"
            strSQL = strSQL & " FROM FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember"
            strSQL = strSQL & " WHERE USUARIOS_CLUB.NoFamilia=" & lNoFamilia
        #End If
    Else
        #If SqlServer_ Then
            strSQL = "UPDATE FECHAS_USUARIO "
            strSQL = strSQL & " SET FechaUltimoPago='" & Format(Me.dtpFechaPago.Value, "yyyymmdd") & "'"
            strSQL = strSQL & " FROM FECHAS_USUARIO"
            strSQL = strSQL & " INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember"
            strSQL = strSQL & " WHERE USUARIOS_CLUB.IdMember=" & Me.ssdbgFechasPago.Columns("IdMember").Value
        #Else
            strSQL = "UPDATE FECHAS_USUARIO INNER JOIN USUARIOS_CLUB"
            strSQL = strSQL & " ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember"
            strSQL = strSQL & " SET FechaUltimoPago='" & Me.dtpFechaPago.Value & "'"
            strSQL = strSQL & " WHERE USUARIOS_CLUB.IdMember=" & Me.ssdbgFechasPago.Columns("IdMember").Value
        #End If
    End If
    Set adocmdFechas = New ADODB.Command
    adocmdFechas.ActiveConnection = Conn
    adocmdFechas.CommandType = adCmdText
    adocmdFechas.CommandText = strSQL
    adocmdFechas.Execute
    
    Set adocmdFechas = Nothing
    
    
    
    LlenaGrid
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
     Me.dtpFechaPago.Value = UltimoDiaDelMes(Date)
End Sub

Private Sub ActivaControles(bValor As Boolean)

    Me.dtpFechaPago.Enabled = bValor
    Me.chkTodos.Enabled = bValor
    Me.cmdCambia.Enabled = bValor
    

End Sub

Private Sub Form_Load()
    LlenaGrid
    CentraForma MDIPrincipal, Me
End Sub

