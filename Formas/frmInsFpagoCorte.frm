VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmInsFpagoCorte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inserta al corte"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbAjuste 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   2295
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
         Columns(0).Width=   3200
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1535
         Columns(1).Caption=   "IdAjuste"
         Columns(1).Name =   "IdAjuste"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtObser 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3840
         Width           =   3735
      End
      Begin VB.TextBox txtImpRecibido 
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txtLote 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtOperacion 
         Height          =   405
         Left            =   1560
         TabIndex        =   3
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbFormaPago 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   2295
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
         Columns.Count   =   6
         Columns(0).Width=   3810
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "IdForma"
         Columns(1).Name =   "IdForma"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   3200
         Columns(2).Visible=   0   'False
         Columns(2).Caption=   "TieneOpcion"
         Columns(2).Name =   "TieneOpcion"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   3200
         Columns(3).Visible=   0   'False
         Columns(3).Caption=   "TieneTPV"
         Columns(3).Name =   "TieneTPV"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Visible=   0   'False
         Columns(4).Caption=   "TieneLote"
         Columns(4).Name =   "TieneLote"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Visible=   0   'False
         Columns(5).Caption=   "TieneOperacion"
         Columns(5).Name =   "TieneOperacion"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbTPV 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
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
         Columns(0).Width=   9313
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "IdTPV"
         Columns(1).Name =   "IdTPV"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label Label4 
         Caption         =   "Motivo del ajuste"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Importe"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Forma de pago"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "# TPV"
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "# Lote"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "# Operaciones"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Insertar"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   4800
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   975
   End
End
Attribute VB_Name = "frmInsFpagoCorte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lIdCorteCaja As Long
Public iNumeroCaja As Long
Public dImporte As Double

Private Sub cmdAceptar_Click()

    Dim iSel As Integer

    If Me.ssCmbAjuste.Text = vbNullString Then
        MsgBox "Seleccionar el motivo del ajuste", vbExclamation, "Verifique"
        Me.ssCmbAjuste.SetFocus
        Exit Sub
    End If
    
    If Me.ssCmbFormaPago.Text = vbNullString Then
        MsgBox "¡Seleccionar una forma de pago!", vbExclamation, "Verifique"
        Me.ssCmbFormaPago.SetFocus
        Exit Sub
    End If
    
    If Me.ssCmbTPV.Visible = True And Me.ssCmbTPV.Text = vbNullString Then
        MsgBox "¡Seleccionar la TPV!", vbExclamation, "Verifique"
        Me.ssCmbTPV.SetFocus
        Exit Sub
    End If
    
    
    If Me.txtLote.Visible And Me.txtLote.Text = vbNullString Then
        MsgBox "Indicar un # de lote", vbExclamation, "Verifique"
        Me.txtLote.SetFocus
        Exit Sub
    End If
    
     If Me.txtOperacion.Visible And Me.txtOperacion.Text = vbNullString Then
        MsgBox "Indicar un # de operacion", vbExclamation, "Verifique"
        Me.txtOperacion.SetFocus
        Exit Sub
    End If
    
    
    
    If Me.txtImpRecibido.Text = vbNullString Then
        MsgBox "Indicar el importe recibido", vbExclamation, "Verifique"
        Me.txtImpRecibido.SetFocus
        Exit Sub
    End If
    
    
    InsPago
    
    
    
    Unload Me
    
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    CentraForma MDIPrincipal, Me
    
    'Llena el combo de formas de pago.
    #If SqlServer_ Then
        strSQL = "SELECT Descripcion, IdFormaPago, TieneOpcion, CONVERT(int, ISNULL(TieneTPV, 0)) AS TieneTPV, CONVERT(int, ISNULL(TieneLote, 0)) AS TieneLote, CONVERT(int, ISNULL(TieneOperacion, 0)) AS TieneOperacion"
        strSQL = strSQL & " FROM FORMA_PAGO"
        strSQL = strSQL & " ORDER BY IdFormaPago"
    #Else
        strSQL = "SELECT Descripcion, IdFormaPago, TieneOpcion, iif(TieneTPV , 1, 0) AS TieneTPV, iif(TieneLote, 1 , 0 ) AS TieneLote, iif(TieneOperacion, 1, 0 ) AS TieneOperacion"
        strSQL = strSQL & " FROM FORMA_PAGO"
        strSQL = strSQL & " ORDER BY IdFormaPago"
    #End If
    
    LlenaSsCombo Me.ssCmbFormaPago, Conn, strSQL, 6
    
    
    'llena el combo de las motivo de ajuste
    strSQL = "SELECT CT_AJUSTES_CORTE.DescripcionAjuste, CT_AJUSTES_CORTE.IdAjuste"
    strSQL = strSQL & " From CT_AJUSTES_CORTE"
    strSQL = strSQL & " Where CT_AJUSTES_CORTE.IdAjuste > 0"
    strSQL = strSQL & " ORDER BY CT_AJUSTES_CORTE.IdAjuste"
    
    LlenaSsCombo Me.ssCmbAjuste, Conn, strSQL, 2
    
    
    Me.txtImpRecibido.Text = dImporte
    
End Sub

Private Sub ssCmbFormaPago_Click()
    
    
    
    
    If Me.ssCmbFormaPago.Columns("TieneTPV").Value = 1 Then
        
        #If SqlServer_ Then
            strSQL = "SELECT '(' + Str(IdInterno) + ') ' + DescripcionTPV AS DescripcionTPV, IdTPV"
            strSQL = strSQL & " FROM CT_TPVS"
            strSQL = strSQL & " WHERE CAJA=" & iNumeroCaja
            strSQL = strSQL & " ORDER BY IdInterno"
        #Else
            strSQL = "SELECT '(' & Str(IdInterno) & ') ' & DescripcionTPV AS DescripcionTPV, IdTPV"
            strSQL = strSQL & " FROM CT_TPVS"
            strSQL = strSQL & " WHERE CAJA=" & iNumeroCaja
            strSQL = strSQL & " ORDER BY IdInterno"
        #End If
    
        LlenaSsCombo Me.ssCmbTPV, Conn, strSQL, 2
        
        
        Me.Label5.Visible = True
        Me.ssCmbTPV.Visible = True
        
        
        
        
        
        
    Else
        
        Me.Label5.Visible = False
        Me.ssCmbTPV.Visible = False
        
         
        
        
        
    End If
    
    If Me.ssCmbFormaPago.Columns("TieneLote").Value = 1 Then
        
        Me.Label6.Visible = True
        Me.txtLote.Visible = True
    Else
        
        Me.Label6.Visible = False
        Me.txtLote.Visible = False
        
    End If
    
    
    If Me.ssCmbFormaPago.Columns("TieneOperacion").Value = 1 Then
        
        Me.Label7.Visible = True
        Me.txtOperacion.Visible = True
    Else
        
        Me.Label7.Visible = False
        Me.txtOperacion.Visible = False
        
    End If
    
    
    
End Sub




Private Sub ssCmbTPV_Click()

            
        'Llena el combo de las opciones  de pago.
     
        strSQL = "SELECT  Modalidad, IdAfiliacion"
        strSQL = strSQL & " FROM CT_AFILIACIONES"
        strSQL = strSQL & " WHERE idTerminal=" & Me.ssCmbTPV.Columns("IdTPV").Value
        strSQL = strSQL & " ORDER BY Modalidad"


        'LlenaSsCombo Me.ssCmbOpcionPago, Conn, strSQL, 2
        
        'Me.ssCmbOpcionPago.Text = vbNullString

End Sub





Private Sub txtImporteaCobrar_Change()
    'CalculaDiferencia
End Sub

Private Sub txtImporteaCobrar_GotFocus()
    'Me.txtImporteaCobrar.SelStart = 0
    'Me.txtImporteaCobrar.SelLength = Len(Me.txtImporteaCobrar.Text)
End Sub




Private Sub txtImpRecibido_GotFocus()
    Me.txtImpRecibido.SelStart = 0
    Me.txtImpRecibido.SelLength = Len(Me.txtImpRecibido.Text)
End Sub

Private Sub txtImpRecibido_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
            
        Case 45 'Signo menos
            If InStr(Me.txtImpRecibido.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 46 'punto decimal
            If InStr(Me.txtImpRecibido.Text, ".") Then
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




Private Sub InsPago()

    Dim adocmd As ADODB.Command
    Dim adorcs As ADODB.Recordset
    
    Dim lRen As Long
    
    strSQL = "SELECT Max(Renglon) As Ultimo"
    strSQL = strSQL & " FROM CORTE_CAJA_DETALLE"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((IdCorteCaja)=" & lIdCorteCaja & ")"
    strSQL = strSQL & ")"
    
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        lRen = adorcs!Ultimo + 1
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    
    strSQL = "INSERT INTO CORTE_CAJA_DETALLE ("
    strSQL = strSQL & "IdCorteCaja" & ","
    strSQL = strSQL & "Renglon" & ","
    strSQL = strSQL & "IdFormaPago" & ","
    strSQL = strSQL & "OpcionPago" & ","
    strSQL = strSQL & "Importe" & ","
    strSQL = strSQL & "Referencia" & ","
    strSQL = strSQL & "IdTPV" & ","
    strSQL = strSQL & "LoteNumero" & ","
    strSQL = strSQL & "NumeroOperaciones" & ","
    strSQL = strSQL & "Nota" & ","
    strSQL = strSQL & "idAjuste" & ")"
    strSQL = strSQL & "VALUES ("
    strSQL = strSQL & lIdCorteCaja & ","
    strSQL = strSQL & lRen & ","
    strSQL = strSQL & Me.ssCmbFormaPago.Columns("IdForma").Value & ","
    strSQL = strSQL & "''" & ","
    strSQL = strSQL & Val(Me.txtImpRecibido.Text) & ","
    strSQL = strSQL & "''" & ","
    strSQL = strSQL & IIf(Me.ssCmbTPV.Visible, Me.ssCmbTPV.Columns("IdTpv").Value, 0) & ","
    strSQL = strSQL & "'" & Trim(Me.txtLote.Text) & "',"
    strSQL = strSQL & 1 & ","
    strSQL = strSQL & "'" & Me.txtObser.Text & "',"
    strSQL = strSQL & Me.ssCmbAjuste.Columns("IdAjuste").Value & ")"

    

    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    Set adocmd = Nothing
     
    
    
    
    
    
    
    
    
    

End Sub



Private Sub txtObser_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 34
            KeyAscii = 0
        Case 39
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub
