VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmPreVentaDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos de inscripcion"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mantenimiento"
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   7695
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbTipoMant 
         Height          =   375
         Left            =   1320
         TabIndex        =   15
         Top             =   840
         Width           =   3495
         DataFieldList   =   "Column 0"
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
         Row.Count       =   3
         Col.Count       =   2
         Row(0).Col(0)   =   "Mensual Direccionado"
         Row(0).Col(1)   =   "1"
         Row(1).Col(0)   =   "Mensual Convencional"
         Row(1).Col(1)   =   "2"
         Row(2).Col(0)   =   "Anual"
         Row(2).Col(1)   =   "3"
         RowHeight       =   423
         Columns.Count   =   2
         Columns(0).Width=   4921
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "IdTipoMant"
         Columns(1).Name =   "IdTipoMant"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6165
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbFormaPagoMant 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   1320
         Width           =   3495
         DataFieldList   =   "Column 0"
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
         Columns(0).Width=   4948
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2223
         Columns(1).Caption=   "IdFormaPago"
         Columns(1).Name =   "IdFormaPago"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   6165
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtCtrl 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbOpcionMant 
         Height          =   375
         Left            =   5640
         TabIndex        =   20
         Top             =   1320
         Width           =   1815
         DataFieldList   =   "Column 0"
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
         Columns(0).Width=   3651
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "IdFormadepagoOpcion"
         Columns(1).Name =   "IdFormadepagoOpcion"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Opcion"
         Height          =   255
         Index           =   8
         Left            =   5040
         TabIndex        =   21
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Forma de pago"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Importe"
         Height          =   255
         Index           =   5
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Inscripción"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbOpcionInsc 
         Height          =   375
         Left            =   5520
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
         DataFieldList   =   "Column 0"
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
         Columns(0).Width=   3651
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "IdFormadepagoOpcion"
         Columns(1).Name =   "IdFormadepagoOpcion"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdbcmbFormaPagoInsc 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   1320
         Width           =   3375
         DataFieldList   =   "Column 0"
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
         Columns(0).Width=   5477
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1931
         Columns(1).Caption=   "IdFormaPago"
         Columns(1).Name =   "IdFormaPago"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   5953
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtCtrl 
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtCtrl 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Width           =   1815
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbTipoInsc 
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   6015
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
         Columns(0).Width=   6350
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Caption=   "IdTipoInsc"
         Columns(1).Name =   "IdTipoInsc"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   10610
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Opcion"
         Height          =   255
         Index           =   4
         Left            =   4800
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Forma de pago"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Enganche"
         Height          =   255
         Index           =   2
         Left            =   4440
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Precio"
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmPreVentaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public lRecNum As Long
Public bReadOnly As Boolean

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
    
    Dim adocmd As ADODB.Command
    
    
    If Me.sscmbTipoInsc.Text = vbNullString Then
        MsgBox "Seleccione un tipo de Inscripción", vbExclamation, "Verifique"
        Me.sscmbTipoInsc.SetFocus
        Exit Sub
    End If
    
    If Me.txtCtrl(0).Text = vbNullString Then
        MsgBox "Indique el precio de la inscripcion", vbExclamation, "Verifique"
        Me.txtCtrl(0).SetFocus
        Exit Sub
    End If
    
    If Me.txtCtrl(1).Text = vbNullString Then
        MsgBox "Indique el monto del enganche", vbExclamation, "Verifique"
        Me.txtCtrl(1).SetFocus
        Exit Sub
    End If
    
    
    If Val(Me.txtCtrl(1).Text) > Val(Me.txtCtrl(0).Text) Then
        MsgBox "El monto del enganche no puede ser mayor" & vbCrLf & "que el precio de la inscripción!", vbExclamation, "Verifique"
        Me.txtCtrl(1).SetFocus
        Exit Sub
    End If
    
    If Me.ssdbcmbFormaPagoInsc.Text = vbNullString Then
        MsgBox "Seleccione el tipo de pago de la inscripción", vbExclamation, "Verifique"
        Me.ssdbcmbFormaPagoInsc.SetFocus
        Exit Sub
    End If
    
    If Me.sscmbOpcionInsc.Rows > 0 And Me.sscmbOpcionInsc.Text = vbNullString Then
        MsgBox "Seleccione la opción de pago de la inscripción", vbExclamation, "Verifique"
        Me.sscmbOpcionInsc.SetFocus
        Exit Sub
    End If
    
    
    If Me.sscmbFormaPagoMant.Text = vbNullString Then
        MsgBox "Seleccione el tipo de pago del mantenimiento", vbExclamation, "Verifique"
        Me.sscmbFormaPagoMant.SetFocus
        Exit Sub
    End If
    
    If Me.sscmbOpcionMant.Rows > 0 And Me.sscmbOpcionMant.Text = vbNullString Then
        MsgBox "Seleccione la opción de pago de la inscripción", vbExclamation, "Verifique"
        Me.sscmbOpcionMant.SetFocus
        Exit Sub
    End If
    
    
    
    strSQL = ""
    strSQL = "UPDATE PROSPECTOS SET"
    strSQL = strSQL & " IdTipoInscripcion=" & Me.sscmbTipoInsc.Columns("IdTipoInsc").Value & ","
    strSQL = strSQL & " PrecioInscripcion=" & Trim(Me.txtCtrl(0)) & ","
    strSQL = strSQL & " MontoEnganche=" & Trim(Me.txtCtrl(1)) & ","
    strSQL = strSQL & " IdFormaPagoInsc=" & Me.ssdbcmbFormaPagoInsc.Columns("IdFormapago").Value & ","
    If Me.sscmbOpcionInsc.Text <> vbNullString Then
        strSQL = strSQL & " IdOPcionPagoInsc=" & Me.sscmbOpcionInsc.Columns("IdFormadepagoOpcion").Value & ","
    End If
    strSQL = strSQL & " ImporteMantenimiento=" & Trim(Me.txtCtrl(2)) & ","
    strSQL = strSQL & " TipoMantenimiento=" & Me.sscmbTipoMant.Columns("IdTipoMant").Value & ","
    strSQL = strSQL & " IdFormaPagoMant=" & Me.sscmbFormaPagoMant.Columns("IdFormapago").Value & ","
    If Me.sscmbOpcionMant.Text <> vbNullString Then
        strSQL = strSQL & " IdOPcionPagoMant=" & Me.sscmbOpcionMant.Columns("IdFormadepagoOpcion").Value & ","
    End If
    #If SqlServer_ Then
        strSQL = strSQL & " FechaPreAlta=" & "'" & Format(Date, "yyyymmdd") & "'" & ","
    #Else
        strSQL = strSQL & " FechaPreAlta=" & "#" & Format(Date, "mm/dd/yyyy") & "#" & ","
    #End If
    strSQL = strSQL & " HoraPreAlta=" & "'" & Format(Now, "Hh:Nn:Ss") & "'" & ","
    strSQL = strSQL & " StatusProspecto=" & 2
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdProspecto=" & lRecNum
    strSQL = strSQL & " And StatusProspecto = 1"
    
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    
    adocmd.Execute
    
    
    Set adocmd = Nothing
    
    MsgBox "Datos de preventa registrados", vbInformation, "Ok"
    
    
    
End Sub

Private Sub Form_Load()
    strSQL = ""
    strSQL = "SELECT Descripcion, IdTipoMembresia"
    strSQL = strSQL & " FROM TIPO_MEMBRESIA"
    strSQL = strSQL & " ORDER BY IdTipoMembresia"
    
    LlenaSsCombo Me.sscmbTipoInsc, Conn, strSQL, 2
    
    
    strSQL = ""
    strSQL = "SELECT Descripcion, IdFormaPago"
    strSQL = strSQL & " FROM FORMA_PAGO"
    strSQL = strSQL & " ORDER BY IdFormaPago"
    
    LlenaSsCombo Me.ssdbcmbFormaPagoInsc, Conn, strSQL, 2
    LlenaSsCombo Me.sscmbFormaPagoMant, Conn, strSQL, 2
    
    CargaDatos
        
    If bReadOnly Then
        Me.cmdGuardar.Enabled = False
    End If
    
    CentraForma MDIPrincipal, Me
    
End Sub

Private Sub sscmbFormaPagoMant_Click()
    LlenaComboOpcionPago Me.sscmbOpcionMant, Me.sscmbFormaPagoMant.Columns("IdFormaPago").Value
    Me.sscmbOpcionMant.Text = vbNullString
End Sub

Private Sub ssdbcmbFormaPagoInsc_Click()
    LlenaComboOpcionPago Me.sscmbOpcionInsc, Me.ssdbcmbFormaPagoInsc.Columns("IdFormaPago").Value
    Me.sscmbOpcionInsc.Text = vbNullString
End Sub

Private Sub LlenaComboOpcionPago(sscmb As SSOleDBCombo, lIdFormaPago As Long)
    Dim adorcsOpcionPago As ADODB.Recordset
    
    strSQL = "SELECT IdFormadePagoOpcion, Descripcion"
    strSQL = strSQL & " FROM FORMA_PAGO_OPCION"
    strSQL = strSQL & " WHERE idFormaPago=" & lIdFormaPago
    strSQL = strSQL & " ORDER BY IdFormadePagoOpcion"
    
    
    Set adorcsOpcionPago = New ADODB.Recordset
    adorcsOpcionPago.CursorLocation = adUseServer
    
    adorcsOpcionPago.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    sscmb.RemoveAll
    
    Do While Not adorcsOpcionPago.EOF
        sscmb.AddItem adorcsOpcionPago!Descripcion & vbTab & adorcsOpcionPago!idformadepagoopcion
        adorcsOpcionPago.MoveNext
    Loop
    
    adorcsOpcionPago.Close
    Set adorcsOpcionPago = Nothing
    
End Sub

Private Sub CargaDatos()
    Dim adorcs As ADODB.Recordset
    Dim sTipoMant As String
        
'    strSQL = "SELECT P.Nombre, P.A_Paterno, P.A_Materno, P.PrecioInscripcion, P.MontoEnganche,  P.ImporteMantenimiento, P.TipoMantenimiento,  P.IdOpcionPagoMant, T.Descripcion, FPI.Descripcion, P.IdOpcionPagoInsc, FPOI.Descripcion, FPM.Descripcion, FPOM.Descripcion"
'    strSQL = strSQL & " FROM ((((PROSPECTOS P INNER JOIN TIPO_MEMBRESIA T ON P.IdTipoInscripcion = T.idTipoMembresia) INNER JOIN FORMA_PAGO FPI ON P.IdFormaPagoInsc = FPI.IdFormaPago) LEFT JOIN FORMA_PAGO_OPCION FPOI ON (P.IdFormaPagoInsc = FPOI.IdFormaPago) AND (P.IdOpcionPagoInsc = FPOI.IdFormadePagoOpcion)) INNER JOIN FORMA_PAGO AS FPM ON P.IdFormaPagoMant = FPM.IdFormaPago) LEFT JOIN FORMA_PAGO_OPCION AS FPOM ON (P.IdFormaPagoMant = FPOM.IdFormaPago) AND (P.IdOpcionPagoMant = FPOM.IdFormadePagoOpcion)"
'    strSQL = strSQL & " Where (((P.idProspecto) = " & lRecNum & "))"
    
    strSQL = "SELECT P.Nombre, P.A_Paterno, P.A_Materno, P.PrecioInscripcion, P.MontoEnganche, P.ImporteMantenimiento, P.TipoMantenimiento, P.IdOpcionPagoMant,  T.Descripcion AS T_Descripcion,"
    strSQL = strSQL & " FPI.Descripcion AS FPI_Descripcion, P.IdOpcionPagoInsc, FPOI.Descripcion AS FPOI_Descripcion, FPM.Descripcion AS FPM_Descripcion, FPOM.Descripcion as FPOM_Descripcion"
    strSQL = strSQL & " FROM PROSPECTOS P"
    strSQL = strSQL & " INNER JOIN TIPO_MEMBRESIA T ON P.IdTipoInscripcion = T.idTipoMembresia"
    strSQL = strSQL & " INNER JOIN FORMA_PAGO FPI ON P.IdFormaPagoInsc = FPI.IdFormaPago"
    strSQL = strSQL & " LEFT JOIN FORMA_PAGO_OPCION FPOI ON P.IdFormaPagoInsc = FPOI.IdFormaPago AND P.IdOpcionPagoInsc = FPOI.IdFormadePagoOpcion"
    strSQL = strSQL & " INNER JOIN FORMA_PAGO AS FPM ON P.IdFormaPagoMant = FPM.IdFormaPago"
    strSQL = strSQL & " LEFT JOIN FORMA_PAGO_OPCION AS FPOM ON P.IdFormaPagoMant = FPOM.IdFormaPago AND P.IdOpcionPagoMant = FPOM.IdFormadePagoOpcion"
    strSQL = strSQL & " WHERE P.idProspecto = " & lRecNum
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
    
        sTipoMant = GetTipoMantenimiento(adorcs!Tipomantenimiento)
    
        If Not IsNull(adorcs![T_Descripcion]) Then Me.sscmbTipoInsc.Text = adorcs![T_Descripcion]
        Me.txtCtrl(0).Text = adorcs!PrecioInscripcion
        Me.txtCtrl(1).Text = adorcs!MontoEnganche
        If Not IsNull(adorcs![FPI_Descripcion]) Then
            Me.ssdbcmbFormaPagoInsc.Text = adorcs![FPI_Descripcion]
            BuscaSSCombo Me.ssdbcmbFormaPagoInsc, adorcs![FPI_Descripcion], 0
            ssdbcmbFormaPagoInsc_Click
            If Not IsNull(adorcs![FPOI_Descripcion]) Then Me.sscmbOpcionInsc.Text = adorcs![FPOI_Descripcion]
        End If
        
        If Not IsNull(sTipoMant) Then Me.sscmbTipoMant.Text = sTipoMant
        Me.txtCtrl(2).Text = adorcs!ImporteMantenimiento
        If Not IsNull(adorcs![FPM_Descripcion]) Then
            Me.sscmbFormaPagoMant.Text = adorcs![FPM_Descripcion]
            If adorcs![FPOI_Descripcion] > 0 Then
                BuscaSSCombo Me.sscmbFormaPagoMant, adorcs![FPOI_Descripcion], 0
            End If
            sscmbFormaPagoMant_Click
            If Not IsNull(adorcs![FPOM_Descripcion]) Then Me.sscmbOpcionMant.Text = adorcs![FPOM_Descripcion]
        End If
    End If
End Sub

