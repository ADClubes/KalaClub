VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmMensajesMovs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensajes"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "frmMensajesMovs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Solo lo puede modifcar quien lo creo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   5640
      Width           =   3015
   End
   Begin VB.CommandButton cmdDesmarcar 
      Height          =   615
      Left            =   2040
      Picture         =   "frmMensajesMovs.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Marcar como no Leído"
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdMarcar 
      Height          =   615
      Left            =   1080
      Picture         =   "frmMensajesMovs.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Marcar como Leído"
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Height          =   615
      Left            =   120
      Picture         =   "frmMensajesMovs.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nuevo Mensaje"
      Top             =   360
      Width           =   735
   End
   Begin VB.CheckBox chkSoloNoLeidos 
      Caption         =   "Ver sólo mensajes No Leídos"
      Height          =   255
      Left            =   5880
      TabIndex        =   0
      Top             =   720
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.TextBox txtMensaje 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3960
      Width           =   8535
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgMensajes 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   8775
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      AllowUpdate     =   0   'False
      AllowGroupMoving=   0   'False
      AllowGroupSwapping=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   1879
      Columns(0).Caption=   "Fecha"
      Columns(0).Name =   "Fecha"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2011
      Columns(1).Caption=   "Hora"
      Columns(1).Name =   "Hora"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2355
      Columns(2).Caption=   "Usuario"
      Columns(2).Name =   "Usuario"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   8308
      Columns(3).Caption=   "Mensaje"
      Columns(3).Name =   "Mensaje"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   512
      Columns(4).Width=   1693
      Columns(4).Caption=   "Bloqueado"
      Columns(4).Name =   "ReadOnly"
      Columns(4).CaptionAlignment=   2
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(4).Locked=   -1  'True
      Columns(4).Style=   2
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "IdMensaje"
      Columns(5).Name =   "IdMensaje"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   15478
      _ExtentY        =   3625
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
      Caption         =   "Mensaje:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmMensajesMovs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lIdTitular As Long

Dim iColIndex As Integer

Private Sub chkSoloNoLeidos_Click()
    ActGridMsg
End Sub

Private Sub cmdCancelar_Click()
    Me.cmdGuardar.Visible = False
    Me.cmdCancelar.Visible = False
        
    Me.txtMensaje.Text = ""
    Me.txtMensaje.Enabled = False
    
    Me.chkReadOnly.Enabled = False
    
    
End Sub

Private Sub cmdDesmarcar_Click()
    Dim adocmdMsg As ADODB.Command
    
    
    If Not ChecaSeguridad(Me.Name, Me.cmdDesmarcar.Name) Then
        Exit Sub
    End If
    
    
    
    If Me.ssdbgMensajes.Rows = 0 Then
        Exit Sub
    End If
    
    If Me.ssdbgMensajes.Columns("ReadOnly").Value = 1 And UCase(sDB_User) <> UCase(Me.ssdbgMensajes.Columns("Usuario").Value) Then
        MsgBox "Este mensaje solo puede ser modificado" & vbCrLf & "por el usuario que lo creo", vbExclamation, "Error"
        Exit Sub
    End If
    
    
    
    strSQL = "UPDATE MENSAJES SET"
    strSQL = strSQL & " Leido=0"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMensaje=" & Me.ssdbgMensajes.Columns("IdMensaje").Value
    
    
    Set adocmdMsg = New ADODB.Command
    adocmdMsg.ActiveConnection = Conn
    adocmdMsg.CommandType = adCmdText
    adocmdMsg.CommandText = strSQL
    
    adocmdMsg.Execute
    
    Set adocmdMsg = Nothing
    
    
    ActGridMsg
End Sub

Private Sub cmdGuardar_Click()
    
    Dim adocmdMsg As ADODB.Command
    
    
    If Me.txtMensaje.Text = vbNullString Then
        MsgBox "El texto del mensaje no puede quedar en blanco", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO Mensajes ("
        strSQL = strSQL & " IdMember,"
        strSQL = strSQL & " FechaAlta,"
        strSQL = strSQL & " HoraAlta,"
        strSQL = strSQL & " IdUsuarioAlta,"
        strSQL = strSQL & " Titulo,"
        strSQL = strSQL & " TextoMensaje,"
        'gpo 16/05/2008 Se agrega el campo ReadOnly
        strSQL = strSQL & " ReadOnly)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lIdTitular & ","
        strSQL = strSQL & "'" & Format(Now, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "',"
        strSQL = strSQL & "'" & Trim(sDB_User) & "',"
        strSQL = strSQL & "'" & "',"
        strSQL = strSQL & "'" & Trim(Me.txtMensaje.Text) & "',"
        strSQL = strSQL & IIf(Me.chkReadOnly.Value, -1, 0) & ")"
    #Else
        strSQL = "INSERT INTO Mensajes ("
        strSQL = strSQL & " IdMember,"
        strSQL = strSQL & " FechaAlta,"
        strSQL = strSQL & " HoraAlta,"
        strSQL = strSQL & " IdUsuarioAlta,"
        strSQL = strSQL & " Titulo,"
        strSQL = strSQL & " TextoMensaje,"
        'gpo 16/05/2008 Se agrega el campo ReadOnly
        strSQL = strSQL & " ReadOnly)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lIdTitular & ","
        strSQL = strSQL & "#" & Format(Now, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "',"
        strSQL = strSQL & "'" & Trim(sDB_User) & "',"
        strSQL = strSQL & "'" & "',"
        strSQL = strSQL & "'" & Trim(Me.txtMensaje.Text) & "',"
        strSQL = strSQL & IIf(Me.chkReadOnly.Value, -1, 0) & ")"
    #End If
    
    Set adocmdMsg = New ADODB.Command
    adocmdMsg.ActiveConnection = Conn
    adocmdMsg.CommandType = adCmdText
    adocmdMsg.CommandText = strSQL
    
    adocmdMsg.Execute
    
    Set adocmdMsg = Nothing
    
    
    ActGridMsg
    
    Me.txtMensaje.Text = ""
    
    Me.cmdGuardar.Visible = False
    Me.cmdCancelar.Visible = False
    
    Me.txtMensaje.Enabled = False
    
    Me.chkReadOnly.Enabled = False
    
End Sub

Private Sub cmdMarcar_Click()
    
    Dim adocmdMsg As ADODB.Command
    
    
    If Not ChecaSeguridad(Me.Name, Me.cmdMarcar.Name) Then
        Exit Sub
    End If
    
    
    If Me.ssdbgMensajes.Rows = 0 Then
        Exit Sub
    End If
    
    If Me.ssdbgMensajes.Columns("ReadOnly").Value = 1 And UCase(sDB_User) <> UCase(Me.ssdbgMensajes.Columns("Usuario").Value) Then
        MsgBox "Este mensaje solo puede ser modificado" & vbCrLf & "por el usuario que lo creo", vbExclamation, "Error"
        Exit Sub
    End If
    
    strSQL = "UPDATE MENSAJES SET"
    strSQL = strSQL & " Leido=-1"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMensaje=" & Me.ssdbgMensajes.Columns("IdMensaje").Value
    
    
    Set adocmdMsg = New ADODB.Command
    adocmdMsg.ActiveConnection = Conn
    adocmdMsg.CommandType = adCmdText
    adocmdMsg.CommandText = strSQL
    
    adocmdMsg.Execute
    
    Set adocmdMsg = Nothing
    
    
    ActGridMsg
End Sub

Private Sub cmdNuevo_Click()
    
     If Not ChecaSeguridad(Me.Name, Me.cmdNuevo.Name) Then
        Exit Sub
    End If
    
    
    
    Me.cmdGuardar.Visible = True
    Me.cmdCancelar.Visible = True
    
    Me.txtMensaje.Enabled = True
    Me.txtMensaje.Text = ""
    
    Me.chkReadOnly.Enabled = True
    
    
    Me.txtMensaje.SetFocus
        
End Sub

Private Sub Form_Load()
    ActGridMsg
End Sub

Private Sub ActGridMsg()
    Dim adorcsMsg As ADODB.Recordset
    
    strSQL = "SELECT FechaAlta, HoraAlta, IdusuarioAlta, TextoMensaje, ReadOnly, idMensaje"
    strSQL = strSQL & " FROM Mensajes"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember=" & lIdTitular
    
    If Me.chkSoloNoLeidos.Value Then
        strSQL = strSQL & " AND Leido=0"
    End If
    
    Select Case iColIndex
        Case 1
            strSQL = strSQL & " ORDER BY HoraAlta"
        Case 2
            strSQL = strSQL & " ORDER BY IdUsuarioAlta"
        Case 3
            strSQL = strSQL & " ORDER BY TextoMensaje"
        Case 4
            strSQL = strSQL & " ORDER BY ReadOnly"
        Case Else
            strSQL = strSQL & " ORDER BY FechaAlta"
    End Select
    
    
    Set adorcsMsg = New ADODB.Recordset
    adorcsMsg.CursorLocation = adUseServer
    
    adorcsMsg.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.ssdbgMensajes.RemoveAll
    Do Until adorcsMsg.EOF
        Me.ssdbgMensajes.AddItem adorcsMsg!FechaAlta & vbTab & Format(adorcsMsg!HoraAlta, "HH:mm:ss") & vbTab & adorcsMsg!IdUsuarioAlta & vbTab & Left$(adorcsMsg!TextoMensaje, 512) & vbTab & IIf(adorcsMsg!ReadOnly, 1, 0) & vbTab & adorcsMsg!IdMensaje
        adorcsMsg.MoveNext
    Loop
    
    adorcsMsg.Close
    Set adorcsMsg = Nothing
    
End Sub







Private Sub ssdbgMensajes_HeadClick(ByVal ColIndex As Integer)
    iColIndex = ColIndex
    ActGridMsg
End Sub

Private Sub ssdbgMensajes_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
    
    If Me.ssdbgMensajes.Rows = 0 Then
        Exit Sub
    End If
    
    Me.txtMensaje.Text = Me.ssdbgMensajes.Columns("Mensaje").Value
    
    
End Sub



Private Sub txtMensaje_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
    End Select
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTitulo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
    End Select
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub
