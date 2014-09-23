VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCuponDesc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insertar Cupon de Descuento"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Datos del cupon"
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   2
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Folio"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Insertar"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Importe"
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   2040
         MaxLength       =   15
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optCtrl 
         Caption         =   "Porcentaje %"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optCtrl 
         Caption         =   "Importe"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optCtrl 
         Caption         =   "Total"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Label lblImporteTotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Cupon Descuento"
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbTipoCupon 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2895
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         AllowNull       =   0   'False
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
         Columns(0).Width=   5265
         Columns(0).Caption=   "TipoCupon"
         Columns(0).Name =   "TipoCupon"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2249
         Columns(1).Caption=   "IdConcepto"
         Columns(1).Name =   "IdConcepto"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   5106
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
   End
End
Attribute VB_Name = "frmCuponDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dMontoTotal As Double

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim dMonto As Double
    Dim lConceptoIngreso As Long
    Dim sDescripcion As String
    Dim dIvaPor As Double
    Dim dImporte As Double
    Dim sCadGrid As String
    Dim sUnidad As String
    
    Dim sFacORec As String
    
    If Me.ssCmbTipoCupon.Text = vbNullString Then
        MsgBox "Seleccionar el tipo de cupón", vbExclamation, "Verifique"
        Me.ssCmbTipoCupon.SetFocus
        Exit Sub
    End If

    If Me.optCtrl(1).Value = True Then
        
        If Me.txtCtrl(0).Text = vbNullString Then
            MsgBox "Indicar el importe", vbExclamation, "Verifique"
            Exit Sub
        End If
        If Val(Me.txtCtrl(0).Text) <= 0 Then
            MsgBox "El importe debe ser mayor a 0", vbExclamation, "Verifique"
            Exit Sub
        End If
        If Val(Me.txtCtrl(0).Text) > dMontoTotal Then
            MsgBox "El importe no puede ser mayor al total", vbExclamation, "Verifique"
            Exit Sub
        End If

    End If
    If Me.optCtrl(2).Value = True Then
        If Me.txtCtrl(1).Text = vbNullString Then
            MsgBox "Indicar el porcentaje de descuento", vbExclamation, "Verifique"
            Exit Sub
        End If
        If Val(Me.txtCtrl(1).Text) <= 0 Then
            MsgBox "El porcentaje debe ser mayor a 0", vbExclamation, "Verifique"
            Exit Sub
        End If
        If Val(Me.txtCtrl(1).Text) > 100 Then
            MsgBox "El porcentaje no puede ser mayor a 100", vbExclamation, "Verifique"
            Exit Sub
        End If
    End If
    If Me.ssCmbTipoCupon.AddItemRowIndex(Me.ssCmbTipoCupon.Bookmark) = 0 Then
        If Me.txtCtrl(2).Text = vbNullString Then
            MsgBox "Indicar el folio del cupón", vbExclamation, "Verifique"
            Exit Sub
        End If
    End If
    
        'Columnas del grid
    '0  Concepto
    '1  Nombre
    '2  Periodo
    '3  Cantidad
    '4  Importe
    '5  Intereses
    '6  Descuento
    '7  Total
    '8  Clave
    '9  IvaPor
    '10 IVA
    '11 IvaDescuento
    '12 IvaInteres
    '13 DescMonto
    '14 IdMember
    '15 NoFamiliar
    '16 FormaPago
    '17 IdTipoUsuario
    '18 TipoCargo
    '19 Auxiliar
    '20 FacoRec
    '21 IdInstructor
    
    
    
    If Me.optCtrl(0).Value = True Then
        dImporte = dMontoTotal
    ElseIf Me.optCtrl(1).Value = True Then
        dImporte = Round(Val(Me.txtCtrl(0)), 2)
    Else
        dImporte = Round(dMontoTotal * (Val(Me.txtCtrl(1).Text) / 100), 2)
    End If
    
    lConceptoIngreso = Val(Me.ssCmbTipoCupon.Columns("IdConcepto").Value)
    
    If lConceptoIngreso = 0 Then
        MsgBox "No hay concepto de ingresos configurado", vbExclamation, "Error"
        Exit Sub
    End If
    
    If (Not ObtieneDatosConceptoIngresos(lConceptoIngreso, sDescripcion, dMonto, dIvaPor, sFacORec, sUnidad)) Then
        sDescripcion = "DESCUENTO"
        dIvaPor = 0
    End If
    
    sCadGrid = vbNullString
    
    sCadGrid = sCadGrid & sDescripcion & vbTab
    sCadGrid = sCadGrid & "" & vbTab
    sCadGrid = sCadGrid & Format(Date, "dd/mm/yyyy") & vbTab
    sCadGrid = sCadGrid & 1 & vbTab
    sCadGrid = sCadGrid & dImporte * -1 & vbTab
    sCadGrid = sCadGrid & 0 & vbTab
    sCadGrid = sCadGrid & 0 & vbTab
    sCadGrid = sCadGrid & dImporte * -1 & vbTab
    sCadGrid = sCadGrid & lConceptoIngreso & vbTab
    sCadGrid = sCadGrid & dIvaPor / 100 & vbTab
    sCadGrid = sCadGrid & Round(dImporte - (dImporte / ((100 + dIvaPor) / 100)), 2) * -1 & vbTab
    sCadGrid = sCadGrid & 0 & vbTab
    sCadGrid = sCadGrid & 0 & vbTab
    sCadGrid = sCadGrid & 0 & vbTab
    sCadGrid = sCadGrid & 0 & vbTab
    sCadGrid = sCadGrid & 1 & vbTab
    sCadGrid = sCadGrid & 1 & vbTab
    sCadGrid = sCadGrid & 0 & vbTab
    sCadGrid = sCadGrid & 5 & vbTab
    sCadGrid = sCadGrid & Trim(Me.txtCtrl(2).Text) & vbTab
    sCadGrid = sCadGrid & "'" & sFacORec & "'" & vbTab
    sCadGrid = sCadGrid & 0 & vbTab
    sCadGrid = sCadGrid & "'" & sUnidad & "'"
    
    frmFacturacion.ssdbgFactura.AddItem sCadGrid

    Unload Me
    
    
End Sub

Private Sub Form_Activate()
    Me.lblImporteTotal.Caption = Format(dMontoTotal, "$#,#0.00")
End Sub

Private Sub Form_Load()
    
    strSQL = "SELECT CONCEPTO_INGRESOS.Descripcion, CT_CONCEPTO_DESCUENTO.IdConcepto"
    strSQL = strSQL & " FROM CT_CONCEPTO_DESCUENTO INNER JOIN CONCEPTO_INGRESOS ON CT_CONCEPTO_DESCUENTO.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((CT_CONCEPTO_DESCUENTO.IniciaVigencia)<=Date())"
    strSQL = strSQL & " AND ((CT_CONCEPTO_DESCUENTO.TerminaVigencia)>=Date())"
    strSQL = strSQL & " AND ((CT_CONCEPTO_DESCUENTO.Status)='A')"
    strSQL = strSQL & ")"
    
    LlenaSsCombo Me.ssCmbTipoCupon, Conn, strSQL, 2
    
    
    
    CentraForma MDIPrincipal, Me
End Sub

Private Sub optCtrl_Click(Index As Integer)
    Select Case Index
        Case 0
            Me.txtCtrl(0).Enabled = False
            Me.txtCtrl(1).Enabled = False
        Case 1
            Me.txtCtrl(0).Enabled = True
            Me.txtCtrl(1).Enabled = False
        Case 2
            Me.txtCtrl(0).Enabled = False
            Me.txtCtrl(1).Enabled = True
    End Select
End Sub

Private Sub ssCmbTipoCupon_Click()
    If Me.ssCmbTipoCupon.AddItemRowIndex(Me.ssCmbTipoCupon.Bookmark) = 0 Then
        Me.Frame3.Visible = True
    Else
        Me.Frame3.Visible = False
    End If
End Sub

Private Sub txtCtrl_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        
        Case 46 'punto decimal
            If InStr(Me.txtCtrl(Index).Text, ".") Then
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
