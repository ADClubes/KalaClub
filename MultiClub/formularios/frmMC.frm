VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmMC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios Multiclub"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3615
      Left            =   5760
      TabIndex        =   3
      Top             =   840
      Width           =   5895
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   12
         Top             =   2400
         Width           =   4575
      End
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   11
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   10
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   9
         Top             =   960
         Width           =   4575
      End
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Status"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   13
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Pagado hasta"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Club origen"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   855
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgResultado 
      Height          =   3615
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5070
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      AllowUpdate     =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   6376
      Columns(0).Caption=   "Nombre"
      Columns(0).Name =   "Nombre"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Inscripcion"
      Columns(1).Name =   "Inscripcion"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Club"
      Columns(2).Name =   "Club"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "FechaPago"
      Columns(3).Name =   "FechaPago"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   4551
      Columns(4).Caption=   "Tipo"
      Columns(4).Name =   "Tipo"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "Status"
      Columns(5).Name =   "Status"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   8943
      _ExtentY        =   6376
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
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdBuscar_Click()
    Dim adorcs As ADODB.Recordset
    
    
    Me.txtNombre.Text = UCase(Me.txtNombre.Text)
        
    strSQL = "SELECT USUARIOS.Nombre & ' ' & USUARIOS.A_Paterno & ' ' & USUARIOS.A_Materno AS Nombre, USUARIOS.NoFamilia, USUARIOS.FechaUltimoPago, USUARIOS.Descripcion, CLUB.NombreClub, iif(USUARIOS.FechaUltimoPago < Date(), 'LLAMAR AL CLUB SEDE', 'OK') AS STATUS"
    strSQL = strSQL & " FROM CLUB INNER JOIN USUARIOS ON CLUB.IdClub = USUARIOS.IdClub"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "(((USUARIOS.Nombre & ' ' & USUARIOS.A_Paterno & ' ' & USUARIOS.A_Materno) Like '%" & Trim(Me.txtNombre.Text) & "%" & "'))"
    strSQL = strSQL & " ORDER BY USUARIOS.Nombre & ' ' & USUARIOS.A_Paterno & ' ' & USUARIOS.A_Materno"
    
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    If adorcs.EOF Then
        adorcs.Close
        Set adorcs = Nothing
        
        MsgBox "No se encontro", vbExclamation, "Error"
        Exit Sub
        
    End If
    
    Me.ssdbgResultado.RemoveAll
    
    Do Until adorcs.EOF
        Me.ssdbgResultado.AddItem adorcs!Nombre & vbTab & adorcs!NoFamilia & vbTab & adorcs!Nombreclub & vbTab & adorcs!FechaUltimopago & vbTab & adorcs!Descripcion & vbTab & adorcs!Status
        adorcs.MoveNext
    Loop
    
    
    adorcs.Close
    Set adorcs = Nothing
    
    
End Sub

Private Sub Form_Activate()
    Me.txtNombre.SetFocus
End Sub

Private Sub Form_Load()
    
    Lee_Ini
    
    If Not Connection_DB() Then
        End
    End If
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Close
    Set Conn = Nothing
End Sub

Private Sub ssdbgResultado_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    Me.txtCtrl(0).Text = Me.ssdbgResultado.Columns("Nombre").Value & " (" & Me.ssdbgResultado.Columns("Inscripcion").Value & ")"
    Me.txtCtrl(1).Text = Me.ssdbgResultado.Columns("Club").Value
    Me.txtCtrl(2).Text = Me.ssdbgResultado.Columns("Tipo").Value
    Me.txtCtrl(3).Text = Me.ssdbgResultado.Columns("FechaPago").Value
    Me.txtCtrl(4).Text = Me.ssdbgResultado.Columns("Status").Value
End Sub

Private Sub txtNombre_GotFocus()
    Me.txtNombre.SelStart = 0
    Me.txtNombre.SelLength = Len(Me.txtNombre.Text)
End Sub
