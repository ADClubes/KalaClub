VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmFacInsFormaPago2
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forma de pago"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   2880
      Width           =   4095
      Begin VB.CheckBox chkGuarda 
         Caption         =   "Guarda como pago rápido"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Insertar"
      Height          =   495
      Left            =   6600
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtOperacion 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtLote 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   1680
         Visible         =   0   'False
         Width           =   2295
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbFormaPago 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbOpcionPago 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
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
         Columns(0).Width=   8493
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "IdOpcion"
         Columns(1).Name =   "IdOpcion"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbTPV 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   720
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
      Begin VB.Label Label7 
         Caption         =   "# Operación/Ref."
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "# Lote"
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "# TPV"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Opción"
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Forma de pago"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtDiferencia 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtImporteaCobrar 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFechaOpera 
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59768833
         CurrentDate     =   39797
      End
      Begin VB.TextBox txtImpRecibido 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Diferencia"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Importe a aplicar"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha operación"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Importe Recibido"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmFacInsFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public doImporte As Double
Public iModo As Integer


Private Sub cmdAceptar_Click()

    Dim iSel As Integer

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
    
    If Me.ssCmbOpcionPago.Visible And Me.ssCmbOpcionPago.Text = vbNullString Then
        MsgBox "¡Seleccionar la opción!", vbExclamation, "Verifique"
        Me.ssCmbOpcionPago.SetFocus
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
    
    If Me.txtImporteaCobrar.Text = vbNullString Then
        MsgBox "Indicar el importe a aplicar", vbExclamation, "Verifique"
        Me.txtImporteaCobrar.SetFocus
        Exit Sub
    End If
    
    If Val(Me.txtImpRecibido.Text) < Val(Me.txtImporteaCobrar.Text) Then
        iSel = MsgBox("El importe recibido no puede ser menor" & vbCrLf & "que el importe a aplicar" & vbCrLf & "¿desea hacerlos iguales?", vbYesNo + vbQuestion, "Confirme")
        
        If iSel = vbYes Then
            Me.txtImporteaCobrar.Text = Me.txtImpRecibido.Text
        Else
            Me.txtImpRecibido.SetFocus
            Exit Sub
        End If
        
        
        
    End If
    
    
    
    If Me.dtpFechaOpera.Value > Date Then
        MsgBox "¡La fecha no puede ser mayor a la fecha actual!", vbExclamation, "Verifique"
        Me.dtpFechaOpera.SetFocus
        Exit Sub
    End If
    
    If iModo = 0 Then
        frmFacturacion.ssdbgPagos.AddItem Me.ssCmbFormaPago.Columns("IdForma").Value & vbTab & Me.ssCmbFormaPago.Columns("Descripcion").Value & vbTab & Me.ssCmbOpcionPago.Columns("Descripcion").Value & vbTab & Me.txtImporteaCobrar.Text & vbTab & vbNullString & vbTab & Me.txtImporteaCobrar.Text & vbTab & Me.ssCmbOpcionPago.Columns("IdOpcion").Value & vbTab & Trim(Me.txtLote.Text) & vbTab & Trim(Me.txtOperacion.Text) & vbTab & Me.txtImpRecibido.Text & vbTab & Me.dtpFechaOpera.Value
    Else
        InsPago
    End If
    
    If Me.chkGuarda.Value Then
        GuardaPagoRapido
    End If
    
    
    Unload Me
    
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    CentraForma MDIPrincipal, Me
    
    'Llena el combo de formas de pago.
    strSQL = "SELECT Descripcion, IdFormaPago, TieneOpcion, iif(TieneTPV , 1, 0) AS TieneTPV, iif(TieneLote, 1 , 0 ) AS TieneLote, iif(TieneOperacion, 1, 0 ) AS TieneOperacion"
    strSQL = strSQL & " FROM FORMA_PAGO"
    strSQL = strSQL & " ORDER BY IdFormaPago"
    
    LlenaSsCombo Me.ssCmbFormaPago, Conn, strSQL, 6
    
    Me.txtImpRecibido.Text = doImporte
    Me.txtImporteaCobrar.Text = doImporte
    Me.txtDiferencia.Text = 0
    
    Me.dtpFechaOpera.Value = Date
    
End Sub

Private Sub ssCmbFormaPago_Click()
    
    
    
    
    If Me.ssCmbFormaPago.Columns("TieneTPV").Value = 1 Then
        
        strSQL = "SELECT '(' & Str(IdInterno) & ') ' & DescripcionTPV AS DescripcionTPV, IdTPV"
        strSQL = strSQL & " FROM CT_TPVS"
        strSQL = strSQL & " WHERE CAJA=" & iNumeroCaja
        strSQL = strSQL & " ORDER BY IdInterno"
    
    
        LlenaSsCombo Me.ssCmbTPV, Conn, strSQL, 2
        
        
        Me.Label5.Visible = True
        Me.ssCmbTPV.Visible = True
        
        Me.Label4.Visible = True
        Me.ssCmbOpcionPago.Visible = True
        
        
        
    Else
        
        Me.Label5.Visible = False
        Me.ssCmbTPV.Visible = False
        
         Me.Label4.Visible = False
        Me.ssCmbOpcionPago.Visible = False
        
        
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


        LlenaSsCombo Me.ssCmbOpcionPago, Conn, strSQL, 2
        
        Me.ssCmbOpcionPago.Text = vbNullString

End Sub





Private Sub txtImporteaCobrar_Change()
    CalculaDiferencia
End Sub

Private Sub txtImporteaCobrar_GotFocus()
    Me.txtImporteaCobrar.SelStart = 0
    Me.txtImporteaCobrar.SelLength = Len(Me.txtImporteaCobrar.Text)
End Sub

Private Sub txtImporteaCobrar_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtImporteaCobrar.Text, ".") Then
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

Private Sub txtImpRecibido_Change()
    CalculaDiferencia
End Sub

Private Sub txtImpRecibido_GotFocus()
    Me.txtImpRecibido.SelStart = 0
    Me.txtImpRecibido.SelLength = Len(Me.txtImpRecibido.Text)
End Sub

Private Sub txtImpRecibido_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
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

Private Sub CalculaDiferencia()
    Me.txtDiferencia.Text = Round((Val(Me.txtImporteaCobrar.Text) - Val(Me.txtImpRecibido.Text)), 2)
End Sub

Private Sub GuardaPagoRapido()
    Dim adocmdGuarda As ADODB.Command
    
    
    Set adocmdGuarda = New ADODB.Command
    adocmdGuarda.ActiveConnection = Conn
    adocmdGuarda.CommandType = adCmdText
    
    
    strSQL = "DELETE *"
    strSQL = strSQL & " FROM CFG_PAGO_RAPIDO"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & " ((IdCaja)=" & iNumeroCaja & ")"
    strSQL = strSQL & ")"
    
    adocmdGuarda.CommandText = strSQL
    adocmdGuarda.Execute
    
    
    strSQL = "INSERT INTO CFG_PAGO_RAPIDO ("
    strSQL = strSQL & " IdCaja" & ","
    strSQL = strSQL & " IdFormaPago" & ","
    strSQL = strSQL & " IdAfiliacion" & ","
    strSQL = strSQL & " LoteNumero" & ","
    strSQL = strSQL & " OperacionNumero" & ","
    strSQL = strSQL & " FechaOperacion" & ")"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & iNumeroCaja & ","
    strSQL = strSQL & Me.ssCmbFormaPago.Columns("IdForma").Value & ","
    strSQL = strSQL & Me.ssCmbOpcionPago.Columns("IdOpcion").Value & ","
    strSQL = strSQL & "'" & Trim(Me.txtLote.Text) & "',"
    strSQL = strSQL & "'" & Trim(Me.txtOperacion.Text) & "',"
    strSQL = strSQL & "'" & Format(Me.dtpFechaOpera.Value, "dd/mm/yyyy") & "')"
    
    adocmdGuarda.CommandText = strSQL
    adocmdGuarda.Execute
    
    
    Set adocmdGuarda = Nothing
    
End Sub
