VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmAyudaClave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ayuda"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelecciona 
      Caption         =   "Selecciona"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdBusca 
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgResultado 
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   6135
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   3
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   8123
      Columns(0).Caption=   "Nombre"
      Columns(0).Name =   "Nombre"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1799
      Columns(1).Caption=   "Familia"
      Columns(1).Name =   "NumeroFamilia"
      Columns(1).Alignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "IdMember"
      Columns(2).Name =   "IdMember"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   10821
      _ExtentY        =   4895
      _StockProps     =   79
      Caption         =   "Resultados de la búsqueda"
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
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6195
   End
   Begin VB.Label lblNoRec 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "frmAyudaClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBusca_Click()
    Dim adorcsBusca As ADODB.Recordset
    Dim sCadCompara
        
    If Me.txtNombre.Text = "" Then
        MsgBox "Indique el nombre a Buscar!", vbExclamation, "Ayuda"
        Exit Sub
    End If
    
    Me.txtNombre.Text = Trim(UCase(Me.txtNombre.Text))
    
    sCadCompara = " LIKE " & "'%" & Trim(Me.txtNombre.Text) & "%'"
    
    #If SqlServer_ Then
        strSQL = "SELECT A_PATERNO + ' ' + A_MATERNO + ' ' + NOMBRE AS NombreCompleto, NoFamilia, IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB"
        strSQL = strSQL & " WHERE A_PATERNO + ' ' + A_MATERNO + ' ' + NOMBRE LIKE " & "'" & Trim(Me.txtNombre.Text) & "%'"
        'strSQL = strSQL & " WHERE ((Nombre & ' ' & A_Paterno & ' ' & A_Materno) LIKE '%" & Trim$(UCase$(Me.txtNombre.Text))
        strSQL = strSQL & " ORDER BY A_PATERNO + ' ' + A_MATERNO + ' ' + NOMBRE"
    #Else
        strSQL = "SELECT A_PATERNO & ' ' & A_MATERNO & ' ' & NOMBRE AS NombreCompleto, NoFamilia, IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB"
        strSQL = strSQL & " WHERE A_PATERNO & ' ' & A_MATERNO & ' ' & NOMBRE LIKE " & "'" & Trim(Me.txtNombre.Text) & "%'"
        'strSQL = strSQL & " WHERE ((Nombre & ' ' & A_Paterno & ' ' & A_Materno) LIKE '%" & Trim$(UCase$(Me.txtNombre.Text))
        strSQL = strSQL & " ORDER BY A_PATERNO & ' ' & A_MATERNO & ' ' & NOMBRE"
    #End If
    
    
    Set adorcsBusca = New ADODB.Recordset
    adorcsBusca.CursorLocation = adUseServer
    adorcsBusca.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.ssdbgResultado.RemoveAll
    
    Do Until adorcsBusca.EOF
        Me.ssdbgResultado.AddItem adorcsBusca!NombreCompleto & vbTab & adorcsBusca!NoFamilia & vbTab & adorcsBusca!IdMember
        adorcsBusca.MoveNext
    Loop
    
    adorcsBusca.Close
    Set adorcsBusca = Nothing
    
    Me.lblNoRec = Format(Me.ssdbgResultado.Rows, "###0") & " Registros Encontrados"
    
    Me.txtNombre.SelStart = 0
    Me.txtNombre.SelLength = Len(Me.txtNombre.Text)
    Me.txtNombre.SetFocus
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelecciona_Click()
        
    Dim frmTarget As Form
        
    If Me.ssdbgResultado.Rows = 0 Then
        Exit Sub
    End If
    
    For Each frmTarget In Forms
         If frmTarget.Caption = "Facturación" Then
            frmTarget.txtClave.Text = Me.ssdbgResultado.Columns("NumeroFamilia").Value
            Exit For
         End If
    Next
    
    
    'Set frmTarget = Forms(Forms.Count - 2)
    'frmTarget.txtClave.Text = Me.ssdbgResultado.Columns("NumeroFamilia").Value
    Set frmTarget = Nothing
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Me.Height = 5745
    Me.Width = 6780
    
    
    CentraForma MDIPrincipal, Me
End Sub

Private Sub ssdbgResultado_DblClick()
    If Me.ssdbgResultado.Rows = 0 Then
        Me.txtNombre.SetFocus
        Exit Sub
    End If
    
    Set frmTarget = Forms(Forms.Count - 2)
    
    
    frmTarget.txtClave.Text = Me.ssdbgResultado.Columns("NumeroFamilia").Value
    
    
    Set frmTarget = Nothing
    
    Unload Me
End Sub


