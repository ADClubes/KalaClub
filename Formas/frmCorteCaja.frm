VERSION 5.00
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCorteCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Corte de caja"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgValida 
      Height          =   1455
      Left            =   4800
      TabIndex        =   38
      Top             =   1920
      Visible         =   0   'False
      Width           =   5775
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   3
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   3200
      Columns(0).Caption=   "IdFormaPago"
      Columns(0).Name =   "IdFormaPago"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "IdTerminal"
      Columns(1).Name =   "IdTerminal"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Lote"
      Columns(2).Name =   "Lote"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   10186
      _ExtentY        =   2566
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
   Begin VB.Frame Frame3 
      Caption         =   "Corte"
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   5280
      Width           =   11055
      Begin VB.TextBox txtFondoaDejar 
         Height          =   375
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   34
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtSerieFin 
         Height          =   375
         Left            =   5520
         MaxLength       =   10
         TabIndex        =   32
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtFolioFin 
         Height          =   375
         Left            =   4320
         MaxLength       =   10
         TabIndex        =   31
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtFolioIni 
         Enabled         =   0   'False
         Height          =   375
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   30
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   9600
         TabIndex        =   37
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdGuarda 
         Caption         =   "Realizar Corte"
         Height          =   495
         Left            =   8040
         TabIndex        =   36
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Fondo en efectivo por dejar"
         Height          =   495
         Left            =   6240
         TabIndex        =   35
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Serie"
         Height          =   255
         Left            =   5520
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "IdTurno"
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblIdTurno 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Folio Final"
         Height          =   255
         Left            =   4320
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Folios usuados"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblTurno 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblCaja 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Turno"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Caja"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Total"
      Height          =   855
      Left            =   8160
      TabIndex        =   3
      Top             =   3960
      Width           =   3135
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdBorra 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Formas de pago"
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtOperaciones 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   13
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtImporte 
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   11
         Top             =   2640
         Width           =   1935
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbLote 
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
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
         Columns(0).Width=   3466
         Columns(0).Caption=   "Lote"
         Columns(0).Name =   "Lote"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtOperaciones 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   12
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txtImporte 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtLote 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdInserta 
         Caption         =   "Inserta"
         Height          =   495
         Left            =   2280
         TabIndex        =   15
         Top             =   4320
         Width           =   1335
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbFormaPago 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   360
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
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
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
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label Label12 
         Caption         =   "# Oper. (Confirme)"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Importe (Confirme)"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "# Oper."
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Forma de pago"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Importe"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "# Lote"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "# TPV"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbgPagos 
      Height          =   3495
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   6975
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   10
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   10
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "IdTipoPago"
      Columns(0).Name =   "IdTipoPago"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3387
      Columns(1).Caption=   "Pago"
      Columns(1).Name =   "DescripPago"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2196
      Columns(2).Caption=   "Importe"
      Columns(2).Name =   "Importe"
      Columns(2).Alignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   6
      Columns(2).NumberFormat=   "CURRENCY"
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "IdAfiliacion"
      Columns(3).Name =   "IdAfiliacion"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "IdTPV"
      Columns(4).Name =   "IdTPV"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2196
      Columns(5).Caption=   "Operaciones"
      Columns(5).Name =   "Operaciones"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "TPV"
      Columns(6).Name =   "TPV"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3201
      Columns(7).Caption=   "LoteNumero"
      Columns(7).Name =   "LoteNumero"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "OperacionNumero"
      Columns(8).Name =   "OperacionNumero"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Visible=   0   'False
      Columns(9).Caption=   "ImporteRecibido"
      Columns(9).Name =   "ImporteRecibido"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   12303
      _ExtentY        =   6165
      _StockProps     =   79
      Caption         =   "Pagos"
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
Attribute VB_Name = "frmCorteCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dTotal As Double



Private Sub cmdBorra_Click()
    If Me.ssdbgPagos.Rows < 1 Then
        Exit Sub
    End If
    
        
    dTotal = dTotal - Me.ssdbgPagos.Columns("Importe").Value
    
    Me.ssdbgPagos.RemoveItem Me.ssdbgPagos.AddItemRowIndex(Me.ssdbgPagos.Bookmark)
    
    Me.lblTotal.Caption = Format(dTotal, "#,#0.00")
    
End Sub

Private Sub cmdCancelar_Click()
    
    Dim iResp As Integer
    
    If Me.ssdbgPagos.Rows > 0 Then
    
        iResp = MsgBox("¿Cancelar el corte?", vbYesNo + vbQuestion, "Confirme")
        
        If iResp = vbNo Then
            Exit Sub
        End If
        
    End If
    
    
    Unload Me
End Sub

Private Sub cmdGuarda_Click()

    Dim adocmdCorte As ADODB.Command
    Dim adorcsCorte As ADODB.Recordset
    
    Dim frmreporte As frmReportViewer
    
    Dim lI As Long
    
    Dim lIdCorteCaja As Long
    
    Dim iResp As Integer
    
    Dim sPrefijo As String
    Dim sReferencia As String
    Dim iDigVer As Integer
    
    Dim sFolioIni As String
    Dim sSerieIni As String
    Dim dFondoIni As Double
    
    Dim sMensajeVal As String
    
    
    iResp = 0
    
    'Valida los datos del cierre del turno
    If ClosedShift(Me.lblIdTurno.Caption) Then
        MsgBox "Este turno ya esta cerrado!", vbCritical, "Verifique"
        Exit Sub
    End If
    
    
    '
    If Me.ssdbgPagos.Rows = 0 Then
        MsgBox "Hay que registrar formas de pago", vbExclamation, "Verifique"
        Me.ssCmbFormaPago.SetFocus
        Exit Sub
    End If
    
    If RecibosPendientes(Date, Val(Me.lblCaja.Caption), Val(Me.lblTurno.Caption)) > 0 Then
        MsgBox "¡Aun quedan por facturar recibos!", vbCritical, "Verifique"
        Exit Sub
    End If
    
    If Not ValidaCierre(Date, Val(Me.lblCaja.Caption), Val(Me.lblTurno.Caption), Date, CDate(Format(Now, "Hh:Nn"))) Then
        MsgBox "No se puede cerrar el turno!", vbCritical, "Verifique"
        Exit Sub
    End If
    
    
'    If Me.txtFolioIni.Text = vbNullString Then
'        MsgBox "Indicar el # de facturas usadas", vbExclamation, "Verifique"
'        Me.txtFolioIni.SetFocus
'        Exit Sub
'    End If
    
    If Me.txtFolioFin.Text = vbNullString Then
        MsgBox "Indicar el # de folio final", vbExclamation, "Verifique"
        Me.txtFolioFin.SetFocus
        Exit Sub
    End If
    
    If Me.txtSerieFin.Text = vbNullString Then
        MsgBox "Indicar la serie del folio final", vbExclamation, "Verifique"
        Me.txtSerieFin.SetFocus
        Exit Sub
    End If
    
    If Me.txtFondoaDejar.Text = vbNullString Then
        MsgBox "Indicar el fondo de afectivo a dejar", vbExclamation, "Verifique"
        Me.txtFondoaDejar.SetFocus
        Exit Sub
    End If
    
    If Not DatosTurno(Date, Val(Me.lblCaja.Caption), Val(Me.lblTurno.Caption), sFolioIni, sSerieIni, dFondoIni) Then
        MsgBox "No se pudieron obtener los datos iniciales del turno", vbExclamation, "Error"
        Exit Sub
    End If
    
    If Not ValidaCorte() Then
        MsgBox "Falta registrar las siguientes operaciones:" + vbCrLf + ValidaMensaje, vbExclamation, "Error"
        Exit Sub
    End If
    
    iResp = MsgBox("¿Proceder con el corte?", vbOKCancel + vbQuestion, "Confirme")
    
    If iResp = vbCancel Then
        Exit Sub
    End If
    
    
    sPrefijo = ObtieneParametro("PREFIJO_UNIDAD")
    sReferencia = UCase(Trim(sPrefijo)) & Format(Now, "yymmdd") & "C" & Format(Trim(Me.lblCaja.Caption), "00") & "T" & Format(Trim(Me.lblTurno.Caption), "00")
    iDigVer = dvAlgoritmo35(sReferencia)
    sReferencia = sReferencia & iDigVer
    
    

    
    
    Set adocmdCorte = New ADODB.Command
    adocmdCorte.ActiveConnection = Conn
    adocmdCorte.CommandType = adCmdText
    
    
    
    'Inserta el encabezado
    #If SqlServer_ Then
        strSQL = "INSERT INTO CORTE_CAJA ("
        strSQL = strSQL & "FechaCorte" & ","
        strSQL = strSQL & "HoraCorte" & ","
        strSQL = strSQL & "UsuarioCorte" & ","
        strSQL = strSQL & "FechaOperacion" & ","
        strSQL = strSQL & "Caja" & ","
        strSQL = strSQL & "Turno" & ","
        strSQL = strSQL & "FolioInicial" & ","
        strSQL = strSQL & "SerieInicial" & ","
        strSQL = strSQL & "FolioFinal" & ","
        strSQL = strSQL & "SerieFinal" & ","
        strSQL = strSQL & "NoFoliosUsados" & ","
        strSQL = strSQL & "FondoInicial" & ","
        strSQL = strSQL & "FondoDejado" & ","
        strSQL = strSQL & "Referencia" & ")"
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & "'" & Format(Now, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Format(Now, "Hh:Nn") & "',"
        strSQL = strSQL & "'" & sDB_User & "',"
        strSQL = strSQL & "'" & Format(Now, "yyyymmdd") & "',"
        strSQL = strSQL & "" & Me.lblCaja.Caption & ","
        strSQL = strSQL & "" & Me.lblTurno.Caption & ","
        strSQL = strSQL & "'" & sFolioIni & "',"
        strSQL = strSQL & "'" & sSerieIni & "',"
        strSQL = strSQL & "'" & Trim(Me.txtFolioFin.Text) & "',"
        strSQL = strSQL & "'" & Trim(Me.txtSerieFin.Text) & "',"
        strSQL = strSQL & Trim(Me.txtFolioIni.Text) & ","
        strSQL = strSQL & dFondoIni & ","
        strSQL = strSQL & Trim(Me.txtFondoaDejar.Text) & ","
        strSQL = strSQL & "'" & sReferencia & "')"
    #Else
        strSQL = "INSERT INTO CORTE_CAJA ("
        strSQL = strSQL & "FechaCorte" & ","
        strSQL = strSQL & "HoraCorte" & ","
        strSQL = strSQL & "UsuarioCorte" & ","
        strSQL = strSQL & "FechaOperacion" & ","
        strSQL = strSQL & "Caja" & ","
        strSQL = strSQL & "Turno" & ","
        strSQL = strSQL & "FolioInicial" & ","
        strSQL = strSQL & "SerieInicial" & ","
        strSQL = strSQL & "FolioFinal" & ","
        strSQL = strSQL & "SerieFinal" & ","
        strSQL = strSQL & "NoFoliosUsados" & ","
        strSQL = strSQL & "FondoInicial" & ","
        strSQL = strSQL & "FondoDejado" & ","
        strSQL = strSQL & "Referencia" & ")"
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & "#" & Format(Now, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "'" & Format(Now, "Hh:Nn") & "',"
        strSQL = strSQL & "'" & sDB_User & "',"
        strSQL = strSQL & "#" & Format(Now, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "" & Me.lblCaja.Caption & ","
        strSQL = strSQL & "" & Me.lblTurno.Caption & ","
        strSQL = strSQL & "'" & sFolioIni & "',"
        strSQL = strSQL & "'" & sSerieIni & "',"
        strSQL = strSQL & "'" & Trim(Me.txtFolioFin.Text) & "',"
        strSQL = strSQL & "'" & Trim(Me.txtSerieFin.Text) & "',"
        strSQL = strSQL & Trim(Me.txtFolioIni.Text) & ","
        strSQL = strSQL & dFondoIni & ","
        strSQL = strSQL & Trim(Me.txtFondoaDejar.Text) & ","
        strSQL = strSQL & "'" & sReferencia & "')"
    #End If
    
    adocmdCorte.CommandText = strSQL
    adocmdCorte.Execute
    
    #If SqlServer_ Then
        strSQL = "SELECT Max(IdCorteCaja) as IdCorteCaja"
        strSQL = strSQL & " FROM CORTE_CAJA"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FechaCorte)='" & Format(Now, "yyyymmdd") & "')"
        strSQL = strSQL & " AND ((UsuarioCorte)='" & sDB_User & "')"
        strSQL = strSQL & " AND ((Caja)=" & Me.lblCaja.Caption & ")"
        strSQL = strSQL & " AND ((Turno)=" & Me.lblTurno.Caption & ")"
        strSQL = strSQL & ")"
    #Else
        strSQL = "SELECT Max(IdCorteCaja) as IdCorteCaja"
        strSQL = strSQL & " FROM CORTE_CAJA"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FechaCorte)=#" & Format(Now, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((UsuarioCorte)='" & sDB_User & "')"
        strSQL = strSQL & " AND ((Caja)=" & Me.lblCaja.Caption & ")"
        strSQL = strSQL & " AND ((Turno)=" & Me.lblTurno.Caption & ")"
        strSQL = strSQL & ")"
    #End If
    
    Set adorcsCorte = New ADODB.Recordset
    
    
    adorcsCorte.CursorLocation = adUseServer
    
    adorcsCorte.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsCorte.EOF Then
        lIdCorteCaja = adorcsCorte!IdCorteCaja
    End If
    
    adorcsCorte.Close
    Set adorcsCorte = Nothing
    
    
    For lI = 0 To Me.ssdbgPagos.Rows - 1
        Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.AddItemBookmark(lI)
        
        strSQL = "INSERT INTO CORTE_CAJA_DETALLE ("
        strSQL = strSQL & "IdCorteCaja" & ","
        strSQL = strSQL & "Renglon" & ","
        strSQL = strSQL & "IdFormaPago" & ","
        strSQL = strSQL & "OpcionPago" & ","
        strSQL = strSQL & "Importe" & ","
        strSQL = strSQL & "Referencia" & ","
        strSQL = strSQL & "IdTPV" & ","
        strSQL = strSQL & "LoteNumero" & ","
        strSQL = strSQL & "NumeroOperaciones" & ")"
        strSQL = strSQL & "VALUES ("
        strSQL = strSQL & lIdCorteCaja & ","
        strSQL = strSQL & lI + 1 & ","
        strSQL = strSQL & Me.ssdbgPagos.Columns("IdTipoPago").Value & ","
        strSQL = strSQL & "''" & ","
        If Me.ssdbgPagos.Columns("IdTipoPago").Value = 1 Then
            strSQL = strSQL & Me.ssdbgPagos.Columns("Importe").Value + -Val(Me.txtFondoaDejar.Text) & ","
        Else
            strSQL = strSQL & Me.ssdbgPagos.Columns("Importe").Value & ","
        End If
        strSQL = strSQL & "''" & ","
        strSQL = strSQL & IIf(Me.ssdbgPagos.Columns("IdTPV").Value = vbNullString, "0", Me.ssdbgPagos.Columns("IdTPV").Value) & ","
        strSQL = strSQL & "'" & Me.ssdbgPagos.Columns("LoteNumero").Value & "',"
        strSQL = strSQL & Me.ssdbgPagos.Columns("Operaciones").Value & ")"
        
        
        
        adocmdCorte.CommandText = strSQL
        adocmdCorte.Execute
    Next
    
    
    
    'Actualiza las facturas que pertencen al corte
    #If SqlServer_ Then
        strSQL = "UPDATE FACTURAS SET"
        strSQL = strSQL & " IdCorteCaja=" & lIdCorteCaja
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FechaFactura)='" & Format(Now, "yyyymmdd") & "')"
        strSQL = strSQL & " AND ((Caja)=" & Me.lblCaja.Caption & ")"
        strSQL = strSQL & " AND ((Turno)=" & Me.lblTurno.Caption & ")"
        strSQL = strSQL & ")"
    #Else
        strSQL = "UPDATE FACTURAS SET"
        strSQL = strSQL & " IdCorteCaja=" & lIdCorteCaja
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FechaFactura)=#" & Format(Now, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((Caja)=" & Me.lblCaja.Caption & ")"
        strSQL = strSQL & " AND ((Turno)=" & Me.lblTurno.Caption & ")"
        strSQL = strSQL & ")"
    #End If
    
    adocmdCorte.CommandText = strSQL
    adocmdCorte.Execute
    
    
    Set adocmdCorte = Nothing
    
    
    If Not CloseShift(Val(Me.lblIdTurno.Caption)) Then
    
    End If
    
    
    ReportesCorte 0, lIdCorteCaja
    
    
    
    
    
    MsgBox "Corte efectuado", vbExclamation, "Ok"
    
    Unload Me
    
End Sub

Private Sub cmdInserta_Click()

    Dim sCadenaIns As String
    
    If Me.ssCmbFormaPago.Text = vbNullString Then
        MsgBox "¡Seleccionar una forma de pago!", vbExclamation, "Verifique"
        Me.ssCmbFormaPago.SetFocus
        Exit Sub
    End If
    
    
    If Me.ssCmbFormaPago.Columns("IdForma").Value = 1 Then
        Me.txtOperaciones(0).Text = 1
        Me.txtOperaciones(1).Text = 1
    End If
    
    
    If Me.ssCmbTPV.Visible = True And Me.ssCmbTPV.Text = vbNullString Then
        MsgBox "¡Seleccionar la TPV!", vbExclamation, "Verifique"
        Me.ssCmbTPV.SetFocus
        Exit Sub
    End If
    
    
    If Me.sscmbLote.Visible And Me.sscmbLote.Text = vbNullString Then
        MsgBox "Indicar un # de lote", vbExclamation, "Verifique"
        Me.txtLote.SetFocus
        Exit Sub
    End If
    
    If Me.txtImporte(0).Text = vbNullString Then
        MsgBox "Indicar el importe recibido", vbExclamation, "Verifique"
        Me.txtImporte(0).SetFocus
        Exit Sub
    End If
    
    If Me.txtImporte(1).Text = vbNullString Then
        MsgBox "Confirme el importe recibido", vbExclamation, "Verifique"
        Me.txtImporte(1).SetFocus
        Exit Sub
    End If
    
    If Me.txtImporte(0).Text <> Me.txtImporte(1).Text Then
        MsgBox "El importe y su confirmación no son iguales", vbExclamation, "Verifique"
        Me.txtImporte(0).SetFocus
        Exit Sub
    End If
    
     If Me.txtOperaciones(0).Text = vbNullString Then
        MsgBox "Indicar un # de operaciones", vbExclamation, "Verifique"
        Me.txtOperaciones(0).SetFocus
        Exit Sub
    End If
    
    If Me.txtOperaciones(1).Text = vbNullString Then
        MsgBox "Confirme el # de operaciones", vbExclamation, "Verifique"
        Me.txtOperaciones(1).SetFocus
        Exit Sub
    End If
    
    If Me.txtOperaciones(0).Text <> Me.txtOperaciones(1).Text Then
        MsgBox "El # de operaciones y su confirmación no son iguales", vbExclamation, "Verifique"
        Me.txtOperaciones(0).SetFocus
        Exit Sub
    End If
    
    
    
    
    sCadenaIns = Me.ssCmbFormaPago.Columns("IdForma").Value & vbTab & Me.ssCmbFormaPago.Columns("Descripcion").Value & vbTab & Me.txtImporte(0).Text & vbTab & vbNullString & vbTab
    
    If Me.ssCmbTPV.Visible Then
        sCadenaIns = sCadenaIns & Me.ssCmbTPV.Columns("IdTPV").Value & vbTab
    Else
        sCadenaIns = sCadenaIns & "" & vbTab
    End If
    
    sCadenaIns = sCadenaIns & Me.txtOperaciones(0).Text & vbTab
    
    If Me.ssCmbTPV.Visible Then
        sCadenaIns = sCadenaIns & Me.ssCmbTPV.Text & vbTab
    Else
        sCadenaIns = sCadenaIns & "" & vbTab
    End If
    
    If Me.sscmbLote.Visible Then
        sCadenaIns = sCadenaIns & Trim(Me.sscmbLote.Text)
    Else
        sCadenaIns = sCadenaIns & ""
    End If
    
    
    
    Me.ssdbgPagos.AddItem sCadenaIns
    
    'Se ubica en el ultimo renglon
    Me.ssdbgPagos.Bookmark = Me.ssdbgPagos.AddItemBookmark(Me.ssdbgPagos.Rows - 1)
    
    dTotal = dTotal + Me.ssdbgPagos.Columns("Importe").Value
    
    
    Me.lblTotal.Caption = Format(dTotal, "#,#0.00")
    
        
        
        
        
    
    Me.ssCmbFormaPago.Text = vbNullString
    Me.ssCmbTPV.Text = vbNullString
    
    Me.txtImporte(0).Text = vbNullString
    Me.txtImporte(1).Text = vbNullString
    Me.txtLote.Text = vbNullString
    Me.sscmbLote.Text = vbNullString
    Me.txtOperaciones(0).Text = vbNullString
    Me.txtOperaciones(1).Text = vbNullString
    
    
        
    
    
End Sub

Private Sub Form_Activate()
    If RecibosPendientes(Date, Val(Me.lblCaja.Caption), Val(Me.lblTurno.Caption)) > 0 Then
        MsgBox "¡Aun quedan por facturar recibos!", vbCritical, "Verifique"
        Unload Me
    End If
End Sub

Private Sub GetUltimaFactura(dFecha As Date, iCaja As Integer, lTurno As Long)
    Dim adorcsTurno As ADODB.Recordset
    
    #If SqlServer_ Then
        strSQL = "select top 1 FolioCFD,SerieCFD from Facturas "
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & "FechaFactura='" & Format(dFecha, "yyyymmdd") & "' "
        strSQL = strSQL & " AND Caja=" & iCaja & " "
        strSQL = strSQL & " AND Turno=" & lTurno & " "
        strSQL = strSQL & "order by NumeroFactura desc "
        
    #End If

    Set adorcsTurno = New ADODB.Recordset
    adorcsTurno.CursorLocation = adUseServer
    adorcsTurno.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsTurno.EOF Then
        'GetIdCorte = adorcsTurno!IdTurno
         Me.txtFolioFin.Text = adorcsTurno!FolioCFD
    
         Me.txtSerieFin.Text = adorcsTurno!SerieCFD
    End If
    
    adorcsTurno.Close
    Set adorcsTurno = Nothing
    
End Sub

Private Sub Form_Load()
    CentraForma MDIPrincipal, Me
    
    'Llena el combo de formas de pago.
    #If SqlServer_ Then
        strSQL = "SELECT Descripcion, IdFormaPago, TieneOpcion, CONVERT(int, ISNULL(TieneTPV,0)) AS TieneTPV, CONVERT(int, ISNULL(TieneLote,0)) AS TieneLote, CONVERT(int, ISNULL(TieneOperacion,0)) AS TieneOperacion"
        strSQL = strSQL & " FROM FORMA_PAGO"
        strSQL = strSQL & " ORDER BY IdFormaPago"
    #Else
        strSQL = "SELECT Descripcion, IdFormaPago, TieneOpcion, iif(TieneTPV , 1, 0) AS TieneTPV, iif(TieneLote, 1 , 0 ) AS TieneLote, iif(TieneOperacion, 1, 0 ) AS TieneOperacion"
        strSQL = strSQL & " FROM FORMA_PAGO"
        strSQL = strSQL & " ORDER BY IdFormaPago"
    #End If
    
    LlenaSsCombo Me.ssCmbFormaPago, Conn, strSQL, 6
    
    Me.lblCaja.Caption = iNumeroCaja
    
    Me.lblTurno.Caption = OpenShiftF()
    
    Me.lblIdTurno.Caption = GetIdCorte(Date, iNumeroCaja, Val(Me.lblTurno.Caption))
    
    GetUltimaFactura Date, iNumeroCaja, Val(Me.lblTurno.Caption)
    
    'Llena el grid con los datos para validar
    #If SqlServer_ Then
        strSQL = "SELECT PAGOS_FACTURA.IdFormaPago, CT_AFILIACIONES.IdTerminal, PAGOS_FACTURA.LoteNumero"
        strSQL = strSQL & " FROM (PAGOS_FACTURA LEFT JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion) INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura"
        strSQL = strSQL & " Where ("
        strSQL = strSQL & "((FACTURAS.FechaFactura)='" & Format(Date, "yyyymmdd") & "')"
        strSQL = strSQL & " And ((FACTURAS.Caja)=" & iNumeroCaja & ")"
        strSQL = strSQL & " And ((FACTURAS.Turno)=" & Me.lblTurno.Caption & ")"
        strSQL = strSQL & " And ((FACTURAS.Cancelada) = 0)"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY PAGOS_FACTURA.IdFormaPago, CT_AFILIACIONES.IdTerminal, PAGOS_FACTURA.LoteNumero;"
    #Else
        strSQL = "SELECT PAGOS_FACTURA.IdFormaPago, CT_AFILIACIONES.IdTerminal, PAGOS_FACTURA.LoteNumero"
        strSQL = strSQL & " FROM (PAGOS_FACTURA LEFT JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion) INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura"
        strSQL = strSQL & " Where ("
        strSQL = strSQL & "((FACTURAS.FechaFactura)=#" & Format(Date, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " And ((FACTURAS.Caja)=" & iNumeroCaja & ")"
        strSQL = strSQL & " And ((FACTURAS.Turno)=" & Me.lblTurno.Caption & ")"
        strSQL = strSQL & " And ((FACTURAS.Cancelada) = False)"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY PAGOS_FACTURA.IdFormaPago, CT_AFILIACIONES.IdTerminal, PAGOS_FACTURA.LoteNumero;"
    #End If
        
    LlenaSsDbGrid Me.ssdbgValida, Conn, strSQL, 3
    
    
    
    
End Sub




Private Sub ssCmbFormaPago_Click()
    If Me.ssCmbFormaPago.Columns("TieneTPV").Value = 1 Then
        
        #If SqlServer_ Then
            strSQL = "SELECT '(' + CONVERT(varchar,IdInterno) + ') ' + DescripcionTPV AS DescripcionTPV, IdTPV"
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
        'Me.txtLote.Visible = True
        Me.sscmbLote.Visible = True
    Else
        
        Me.Label6.Visible = False
        Me.txtLote.Visible = False
        Me.sscmbLote.Visible = False
    End If
    
    If Me.ssCmbFormaPago.Columns("IdForma").Value = 1 Then
        Me.txtOperaciones(0).Text = 1
        Me.txtOperaciones(1).Text = 1
    Else
        Me.txtOperaciones(0).Text = vbNullString
        Me.txtOperaciones(1).Text = vbNullString
    End If
    
End Sub




Private Sub ssCmbTPV_Click()
    
    #If SqlServer_ Then
        strSQL = "SELECT DISTINCT PAGOS_FACTURA.LoteNumero"
        strSQL = strSQL & " FROM (PAGOS_FACTURA INNER JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion) INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((FACTURAS.FechaFactura)='" & Format(Date, "yyyymmdd") & "')"
        strSQL = strSQL & " AND ((CT_AFILIACIONES.IdTerminal)=" & Me.ssCmbTPV.Columns("IdTPV").Value & ")"
        strSQL = strSQL & " AND ((FACTURAS.Caja)=" & Val(Me.lblCaja.Caption) & ")"
        strSQL = strSQL & " AND ((FACTURAS.Turno)=" & Val(Me.lblTurno.Caption) & ")"
        strSQL = strSQL & ")"
    #Else
        strSQL = "SELECT DISTINCT PAGOS_FACTURA.LoteNumero"
        strSQL = strSQL & " FROM (PAGOS_FACTURA INNER JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion) INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((FACTURAS.FechaFactura)=#" & Format(Date, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((CT_AFILIACIONES.IdTerminal)=" & Me.ssCmbTPV.Columns("IdTPV").Value & ")"
        strSQL = strSQL & " AND ((FACTURAS.Caja)=" & Val(Me.lblCaja.Caption) & ")"
        strSQL = strSQL & " AND ((FACTURAS.Turno)=" & Val(Me.lblTurno.Caption) & ")"
        strSQL = strSQL & ")"
    #End If
    
    Screen.MousePointer = vbHourglass
    
    LlenaSsCombo Me.sscmbLote, Conn, strSQL, 1
    
    Screen.MousePointer = vbDefault
    
    Me.sscmbLote.Text = vbNullString

End Sub





Private Sub txtFolioFin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub



Private Sub txtFolioIni_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub



Private Sub txtFondoaDejar_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtFondoaDejar.Text, ".") Then
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

Private Sub txtImporte_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtImporte(Index).Text, ".") Then
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




Private Sub txtOperaciones_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSerieFin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 65 To 90 ' Letras de la A a la Z
            KeyAscii = KeyAscii
        Case 97 To 122 ' Letras de la a a la z
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Function ValidaCorte() As Boolean
    Dim vBookMark As Variant
    Dim lI As Integer
    
    ValidaCorte = False
    
    For lI = 0 To Me.ssdbgPagos.Rows - 1
        vBookMark = Me.ssdbgPagos.AddItemBookmark(lI)
        If ValidaBuscaRen(Me.ssdbgPagos.Columns("IdTipoPago").CellValue(vBookMark), Me.ssdbgPagos.Columns("IdTPV").CellValue(vBookMark), Me.ssdbgPagos.Columns("LoteNumero").CellValue(vBookMark)) Then
        End If
    Next
    
    If Me.ssdbgValida.Rows = 0 Then
        ValidaCorte = True
    End If
    
End Function


Private Function ValidaBuscaRen(IdFormaPago As String, IdTerminal As String, sLote As String) As Boolean
    Dim vBookMark As Variant
    Dim lI As Integer
    
    ValidaBuscaRen = True
    
    For lI = 0 To Me.ssdbgValida.Rows - 1
        vBookMark = Me.ssdbgValida.AddItemBookmark(lI)
        If Me.ssdbgValida.Columns("IdFormaPago").CellValue(vBookMark) = IdFormaPago And Me.ssdbgValida.Columns("IdTerminal").CellValue(vBookMark) = IdTerminal And Me.ssdbgValida.Columns("Lote").CellValue(vBookMark) = sLote Then
            
            Me.ssdbgValida.RemoveItem (Me.ssdbgValida.AddItemRowIndex(vBookMark))
            
            Exit Function
        End If
    Next
End Function

Private Function ValidaMensaje() As String
    Dim vBookMark As Variant
    Dim lI As Integer
    
    ValidaMensaje = ""
    
    For lI = 0 To Me.ssdbgValida.Rows - 1
        vBookMark = Me.ssdbgValida.AddItemBookmark(lI)
            
        ValidaMensaje = ValidaMensaje & "Forma de pago: " & Me.ssdbgValida.Columns("IdFormaPago").CellValue(vBookMark)
        ValidaMensaje = ValidaMensaje & " TPV: " & Me.ssdbgValida.Columns("IdTerminal").CellValue(vBookMark)
        ValidaMensaje = ValidaMensaje & " Lote: " & Me.ssdbgValida.Columns("Lote").CellValue(vBookMark)
        ValidaMensaje = ValidaMensaje & vbCrLf
        
    Next
End Function

