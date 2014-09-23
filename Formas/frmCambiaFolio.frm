VERSION 5.00
Begin VB.Form frmCambiaFolio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Folios"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Folio Impreso Facturas"
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   4095
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   2280
         TabIndex        =   17
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdActFolInt 
         Caption         =   "Actualizar"
         Height          =   495
         Left            =   600
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtFolioIntNuevo 
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtFolioIntAct 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cmbSerie 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Folio Nuevo"
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Folio Actual"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Serie"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Folio Interno"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtFolioActual 
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton optFac 
         Caption         =   "Facturas"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optRec 
         Caption         =   "Recibos"
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtFolioNuevo 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Actualizar"
         Default         =   -1  'True
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Folio Actual"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Folio Nuevo"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCambiaFolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbSerie_Click()
    Dim adorcsSerie As ADODB.Recordset
    
    Set adorcsSerie = New ADODB.Recordset
    adorcsSerie.CursorLocation = adUseServer
    
    
    sStrSql = "SELECT NumeroFactura"
    sStrSql = sStrSql & " FROM FOLIO_FACTURA_SERIE"
    sStrSql = sStrSql & " WHERE"
    If Me.cmbSerie.Text = "SinSerie" Then
        sStrSql = sStrSql & " SerieFactura Is Null"
    Else
        sStrSql = sStrSql & " SerieFactura=" & "'" & Trim(Me.cmbSerie.Text) & "'"
    End If
    
    adorcsSerie.Open sStrSql, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Me.txtFolioIntAct.Text = adorcsSerie!NumeroFactura + 1
    Me.txtFolioIntNuevo.Text = Val(Me.txtFolioIntAct.Text) + 1
    
    'Me.txtFolioIntNuevo.SelStart = 0
    Me.txtFolioIntNuevo.SelLength = Len(Me.txtFolioIntNuevo.Text)
        
    adorcsSerie.Close
    Set adorcsSerie = Nothing
    
    
End Sub

Private Sub cmdActFolInt_Click()
    Dim adocmdFol As ADODB.Command
    
    If Val(Me.txtFolioIntNuevo.Text) <= 0 Then
        MsgBox "El folio nuevo debe ser mayor que cero!", vbCritical, "Folios"
        Me.txtFolioIntNuevo.SetFocus
        Exit Sub
    End If
    
    strSQL = "UPDATE FOLIO_FACTURA_SERIE SET"
    strSQL = strSQL & " NumeroFactura="
    strSQL = strSQL & Val(Me.txtFolioIntNuevo.Text) - 1
    strSQL = strSQL & " WHERE"
    If Me.cmbSerie.Text = "SinSerie" Then
        strSQL = strSQL & " SerieFacTura Is Null"
    Else
        strSQL = strSQL & " SerieFactura=" & "'" & Trim(Me.cmbSerie.Text) & "'"
    End If
    
    Set adocmdFol = New ADODB.Command
    adocmdFol.ActiveConnection = Conn
    adocmdFol.CommandType = adCmdText
    adocmdFol.CommandText = strSQL
    adocmdFol.Execute
    
    Set adocmd = Nothing
    
    MsgBox "Folio actualizado!", vbInformation, "Folios"
    
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim adocmdFol As ADODB.Command
    
    If Val(Me.txtFolioNuevo.Text) <= 0 Then
        MsgBox "El folio nuevo debe ser mayor que cero!", vbCritical, "Folios"
        Me.txtFolioNuevo.SetFocus
        Exit Sub
    End If
    
    strSQL = "UPDATE FOLIO_FACTURA SET"
    
    If Me.optFac.Value Then
        strSQL = strSQL & " NumeroFactura="
    Else
        strSQL = strSQL & " NumeroRecibo="
    End If
    
    strSQL = strSQL & Val(Me.txtFolioNuevo.Text) - 1
    
    Set adocmdFol = New ADODB.Command
    adocmdFol.ActiveConnection = Conn
    adocmdFol.CommandType = adCmdText
    adocmdFol.CommandText = strSQL
    adocmdFol.Execute
    
    Set adocmd = Nothing
    
    MsgBox "Folio actualizado!", vbInformation, "Folios"
    
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.txtFolioNuevo.SelLength = Len(Me.txtFolioNuevo.Text)
    Me.txtFolioNuevo.SetFocus
End Sub

Private Sub Form_Load()
    
    
    Me.Height = 6480
    Me.Width = 4800
    
    CENTRAFORMA MDIPrincipal, Me
    
    
    LlenaComboSerie
    
    Me.optFac.Value = True
    
    
End Sub

Private Sub optFac_Click()
    CargaDatos
End Sub

Private Sub CargaDatos()
    Dim adorcsFol As ADODB.Recordset
    
    
    strSQL = "SELECT TOP 1 NumeroFactura, NumeroRecibo"
    strSQL = strSQL & " FROM FOLIO_FACTURA"
    
    Set adorcsFol = New ADODB.Recordset
    adorcsFol.CursorLocation = adUseServer
    
    adorcsFol.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsFol.EOF Then
        If Me.optFac.Value Then
            Me.txtFolioActual.Text = adorcsFol!NumeroFactura + 1
        Else
            Me.txtFolioActual.Text = adorcsFol!NumeroRecibo + 1
        End If
    Else
        If Me.optFac.Value Then
            Me.txtFolioActual.Text = 1
        Else
            Me.txtFolioActual.Text = 1
        End If
    End If
    
    adorcsFol.Close
    
    Set adorcsFol = Nothing
    
    Me.txtFolioNuevo.Text = Val(Me.txtFolioActual.Text) + 1
    Me.txtFolioNuevo.SelLength = Len(Me.txtFolioNuevo.Text)
    
End Sub

Private Sub optRec_Click()
    CargaDatos
End Sub


Private Sub txtFolioNuevo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
            SendKeys vbTab
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub LlenaComboSerie()
    
    Dim adorcsSerie As ADODB.Recordset
    
    Set adorcsSerie = New ADODB.Recordset
    adorcsSerie.CursorLocation = adUseServer
    
    
    sStrSql = "SELECT SerieFactura"
    sStrSql = sStrSql & " FROM FOLIO_FACTURA_SERIE"
    sStrSql = sStrSql & " ORDER BY SerieFactura"
    
    adorcsSerie.Open sStrSql, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.cmbSerie.Clear
    
    Do Until adorcsSerie.EOF
        If IsNull(adorcsSerie!SerieFactura) Then
            Me.cmbSerie.AddItem "SinSerie"
        Else
            Me.cmbSerie.AddItem adorcsSerie!SerieFactura
        End If
        adorcsSerie.MoveNext
    Loop
    
    adorcsSerie.Close
    Set adorcsSerie = Nothing
    
End Sub
