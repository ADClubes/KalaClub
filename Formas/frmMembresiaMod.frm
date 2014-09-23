VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmMembresiaMod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificación condiciones Inscripción"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmActual 
      Caption         =   "Condiciones Actuales"
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5055
      Begin VB.TextBox txtMontoPagado 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   30
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtFechaAltaAct 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   23
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtDescMemAct 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtPagosPendientes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtPagosHechos 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtNoPagosActual 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtEngancheActual 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtMontoActual 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   1
         Top             =   2040
         Width           =   1335
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgPagosActual 
         Height          =   2295
         Left            =   240
         TabIndex        =   3
         Top             =   3840
         Width           =   4695
         _Version        =   196616
         DataMode        =   2
         Col.Count       =   6
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         RowHeight       =   423
         Columns.Count   =   6
         Columns(0).Width=   1429
         Columns(0).Caption=   "NoPago"
         Columns(0).Name =   "NoPago"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Vence"
         Columns(1).Name =   "Vence"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2143
         Columns(2).Caption=   "Monto"
         Columns(2).Name =   "Monto"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   2566
         Columns(3).Caption=   "FechaPago"
         Columns(3).Name =   "FechaPago"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Caption=   "Observaciones"
         Columns(4).Name =   "Observaciones"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Caption=   "idReg"
         Columns(5).Name =   "idReg"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         _ExtentX        =   8281
         _ExtentY        =   4048
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
      Begin VB.Label Label8 
         Caption         =   "Monto Pagado"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha Alta"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo Inscripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Pagos pendientes"
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Pagos realizados"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Pagos"
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Enganche"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
      End
   End
   Begin VB.TextBox txtIdMember 
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame frmNuevo 
      Caption         =   "Condiciones Nuevas"
      Height          =   6495
      Left            =   5520
      TabIndex        =   6
      Top             =   600
      Width           =   5055
      Begin VB.CommandButton cmdGenPagos 
         Caption         =   "Genera Pagos"
         Height          =   495
         Left            =   3600
         TabIndex        =   29
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtDiferencia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   27
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cmbTipoMemNuevo 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox txtMotivo 
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   4695
      End
      Begin VB.TextBox txtNoPagosNuevo 
         Height          =   285
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgPagosNuevo 
         Height          =   2295
         Left            =   240
         TabIndex        =   9
         Top             =   3840
         Width           =   4695
         _Version        =   196616
         DataMode        =   2
         Col.Count       =   6
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowDragDrop   =   0   'False
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         RowHeight       =   423
         Columns.Count   =   6
         Columns(0).Width=   1429
         Columns(0).Caption=   "NoPago"
         Columns(0).Name =   "NoPago"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2275
         Columns(1).Caption=   "Vence"
         Columns(1).Name =   "Vence"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2143
         Columns(2).Caption=   "Monto"
         Columns(2).Name =   "Monto"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   2566
         Columns(3).Caption=   "FechaPago"
         Columns(3).Name =   "FechaPago"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Caption=   "Observaciones"
         Columns(4).Name =   "Observaciones"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Caption=   "idReg"
         Columns(5).Name =   "idReg"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         _ExtentX        =   8281
         _ExtentY        =   4048
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
      Begin VB.TextBox txtMontoNuevo 
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Diferencia"
         Height          =   255
         Left            =   1920
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo Inscripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Pagos"
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Monto"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMembresiaMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CargaDatos()
    
    Dim adorsMem As ADODB.Recordset
    
    Dim iPagosHechos As Integer
    Dim dMontoPagado As Double
    
    Dim lIdTipoMembresia
    
    iPagosHechos = 0
    
    strSQL = "SELECT MEMBRESIAS.Monto, MEMBRESIAS.Enganche, MEMBRESIAS.FechaAlta, MEMBRESIAS.NumeroPagos, MEMBRESIAS.Observaciones, MEMBRESIAS.IdTipoMembresia, "
    strSQL = strSQL & " DETALLE_MEM.NoPago, DETALLE_MEM.Monto, DETALLE_MEM.FechaVence, DETALLE_MEM.FechaPago, DETALLE_MEM.Observaciones, TIPO_MEMBRESIA.Descripcion"
    strSQL = strSQL & " FROM (MEMBRESIAS INNER JOIN DETALLE_MEM"
    strSQL = strSQL & " ON MEMBRESIAS.IdMembresia=DETALLE_MEM.IdMembresia)"
    strSQL = strSQL & " LEFT JOIN TIPO_MEMBRESIA"
    strSQL = strSQL & " ON MEMBRESIAS.IdTipoMembresia=TIPO_MEMBRESIA.IdTipoMembresia"
    strSQL = strSQL & " WHERE IdMember=" & Trim(Me.txtIdMember.Text)
    
    Set adorsMem = New ADODB.Recordset
    adorsMem.CursorLocation = adUseServer
    
    adorsMem.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not adorsMem.EOF Then
        Me.txtDescMemAct.Text = adorsMem!Descripcion
        Me.txtFechaAltaAct.Text = Format(adorsMem!FechaAlta, "dd/MMM/yyyy")
        Me.txtMontoActual.Text = adorsMem![Membresias.Monto]
        Me.txtEngancheActual.Text = adorsMem!Enganche
        Me.txtNoPagosActual.Text = adorsMem!NumeroPagos
        
        
        Me.txtMontoNuevo.Text = Me.txtMontoActual.Text
        
        lIdTipoMembresia = adorsMem!IdTipoMembresia
    End If
    
    Me.ssdbgPagosActual.RemoveAll
    Me.ssdbgPagosNuevo.RemoveAll
    
    Do Until adorsMem.EOF
        Me.ssdbgPagosActual.AddItem adorsMem!NoPago & vbTab & adorsMem!FechaVence & vbTab & adorsMem![DETALLE_MEM.Monto] & vbTab & adorsMem!fechapago & vbTab & adorsMem![DETALLE_MEM.Observaciones]
        Me.ssdbgPagosNuevo.AddItem adorsMem!NoPago & vbTab & adorsMem!FechaVence & vbTab & adorsMem![DETALLE_MEM.Monto] & vbTab & adorsMem!fechapago & vbTab & adorsMem![DETALLE_MEM.Observaciones]
        If Not IsNull(adorsMem!fechapago) Then
            If adorsMem!NoPago > 0 Then
                iPagosHechos = iPagosHechos + 1
            End If
            dMontoPagado = dMontoPagado + adorsMem![DETALLE_MEM.Monto]
        End If
        adorsMem.MoveNext
    Loop
    
    
    
    
    adorsMem.Close
    Set adorsMem = Nothing
    
    Me.txtPagosHechos.Text = iPagosHechos
    Me.txtMontoPagado.Text = dMontoPagado
    Me.txtPagosPendientes.Text = Val(Me.txtNoPagosActual.Text) - iPagosHechos
    
    LlenaCmbTipoMem (lIdTipoMembresia)
    
    
End Sub

Private Sub Form_Activate()
    CargaDatos
End Sub

Private Sub Form_Load()
    Me.txtIdMember.Text = frmAltaSocios.txtTitCve.Text
End Sub

Private Sub LlenaCmbTipoMem(lTipoMem As Long)

    Dim adorcsTipoMem As ADODB.Recordset
    Dim lIndex As Long
    
    strSQL = "SELECT IdTipoMembresia, Descripcion"
    strSQL = strSQL & " FROM TIPO_MEMBRESIA"
    strSQL = strSQL & " ORDER BY Descripcion"
    
    Set adorcsTipoMem = New ADODB.Recordset
    adorcsTipoMem.CursorLocation = adUseServer
    
    adorcsTipoMem.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Me.cmbTipoMemNuevo.Clear
    Do Until adorcsTipoMem.EOF
        Me.cmbTipoMemNuevo.AddItem adorcsTipoMem!Descripcion
        Me.cmbTipoMemNuevo.ItemData(Me.cmbTipoMemNuevo.NewIndex) = adorcsTipoMem!IdTipoMembresia
        If adorcsTipoMem!IdTipoMembresia = lTipoMem Then
            lIndex = Me.cmbTipoMemNuevo.NewIndex
        End If
        
        adorcsTipoMem.MoveNext
    Loop
    
    adorcsTipoMem.Close
    Set adorcsTipoMem = Nothing
    
    If lIndex >= 0 And lIndex <= Me.cmbTipoMemNuevo.ListCount - 1 Then
       Me.cmbTipoMemNuevo.ListIndex = lIndex
    End If
    
End Sub

Private Sub txtMontoNuevo_Change()
    Me.txtDiferencia.Text = Val(Me.txtMontoNuevo.Text) - (Val(Me.txtMontoActual.Text))
End Sub

Private Sub txtMontoNuevo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtMontoNuevo.Text, ".") Then
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

Private Sub txtNoPagosNuevo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub
