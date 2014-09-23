VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmValidaCorteDia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validación del corte del día"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdValida 
      Caption         =   "Valida"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9000
      TabIndex        =   5
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Reporte"
      Enabled         =   0   'False
      Height          =   615
      Left            =   7320
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdBusca 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   59113473
      CurrentDate     =   40287
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgCorte 
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   9975
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   5
      AllowColumnMoving=   0
      AllowColumnSwapping=   0
      SelectTypeRow   =   1
      RowHeight       =   423
      CaptionAlignment=   0
      Columns.Count   =   5
      Columns(0).Width=   3200
      Columns(0).Caption=   "IdCorte"
      Columns(0).Name =   "IdCorte"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "Caja"
      Columns(1).Name =   "Caja"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Turno"
      Columns(2).Name =   "Turno"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Status"
      Columns(3).Name =   "Status"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "Importe"
      Columns(4).Name =   "Importe"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      _ExtentX        =   17595
      _ExtentY        =   2990
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
End
Attribute VB_Name = "frmValidaCorteDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBusca_Click()
    #If SqlServer_ Then
        strSQL = "SELECT CORTE_CAJA.IdCorteCaja, CORTE_CAJA.Caja, CORTE_CAJA.Turno, CORTE_CAJA.Status, Sum(CORTE_CAJA_DETALLE.Importe) AS SumaDeImporte"
        strSQL = strSQL & " FROM CORTE_CAJA INNER JOIN CORTE_CAJA_DETALLE ON CORTE_CAJA.IdCorteCaja = CORTE_CAJA_DETALLE.IdCorteCaja"
        strSQL = strSQL & " WHERE CORTE_CAJA.FechaCorte = '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "'"
        strSQL = strSQL & " GROUP BY CORTE_CAJA.IdCorteCaja, CORTE_CAJA.Caja, CORTE_CAJA.Turno, CORTE_CAJA.Status"
    #Else
        strSQL = "SELECT CORTE_CAJA.IdCorteCaja, CORTE_CAJA.Caja, CORTE_CAJA.Turno, CORTE_CAJA.Status, Sum(CORTE_CAJA_DETALLE.Importe) AS SumaDeImporte"
        strSQL = strSQL & " FROM CORTE_CAJA INNER JOIN CORTE_CAJA_DETALLE ON CORTE_CAJA.IdCorteCaja = CORTE_CAJA_DETALLE.IdCorteCaja"
        strSQL = strSQL & " Where (((CORTE_CAJA.FechaCorte) = #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#))"
        strSQL = strSQL & " GROUP BY CORTE_CAJA.IdCorteCaja, CORTE_CAJA.Caja, CORTE_CAJA.Turno, CORTE_CAJA.Status"
    #End If
    
    LlenaSsDbGrid Me.ssdbgCorte, Conn, strSQL, 5
    
    If Me.ssdbgCorte.Rows = 0 Then
        MsgBox "No se encontró información", vbInformation, "Verifique"
    End If
    
    Me.cmdReporte.Enabled = True
    Me.cmdValida.Enabled = True
    
End Sub

Private Sub cmdReporte_Click()
    Dim adorcs As ADODB.Recordset
    Dim adocmd As ADODB.Command
    
    
    Dim sIdent As String
    Dim sPrefijo As String
    Dim sReferencia As String
    Dim iDigVer As Integer
    Dim sCuenta As String
    
    Dim frmRep As frmReportViewer
    
    
    sIdent = Format(Now, "ddmmHhNnSs")
    sPrefijo = ObtieneParametro("PREFIJO_UNIDAD")
      
    
    sCuenta = sCuenta & ObtieneParametro("CUENTA_DEPOSITO")
    
    #If SqlServer_ Then
        strSQL = "SELECT FORMA_PAGO.Descripcion, FORMA_PAGO.IdFormaPago, SUM(CORTE_CAJA_DETALLE.Importe) AS Importe"
        strSQL = strSQL & " FROM CORTE_CAJA_DETALLE INNER JOIN CORTE_CAJA ON CORTE_CAJA_DETALLE.IdCorteCaja = CORTE_CAJA.IdCorteCaja INNER JOIN FORMA_PAGO ON CORTE_CAJA_DETALLE.IdFormaPago = FORMA_PAGO.IdFormaPago"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " CORTE_CAJA.FechaCorte='" & Format(Me.dtpFecha.Value, "yyyymmdd") & "'"
        strSQL = strSQL & " AND CORTE_CAJA_DETALLE.IdAjuste=0"
        strSQL = strSQL & " GROUP BY FORMA_PAGO.Descripcion, FORMA_PAGO.IdFormaPago"
        strSQL = strSQL & " ORDER BY FORMA_PAGO.IdFormaPago"
    #Else
        strSQL = "SELECT FORMA_PAGO.Descripcion, FORMA_PAGO.IdFormaPago, Sum(CORTE_CAJA_DETALLE.Importe) AS Importe"
        strSQL = strSQL & " FROM (CORTE_CAJA_DETALLE INNER JOIN CORTE_CAJA ON CORTE_CAJA_DETALLE.IdCorteCaja = CORTE_CAJA.IdCorteCaja) INNER JOIN FORMA_PAGO ON CORTE_CAJA_DETALLE.IdFormaPago = FORMA_PAGO.IdFormaPago"
        strSQL = strSQL & " Where ("
        strSQL = strSQL & "((CORTE_CAJA.FechaCorte)=#" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((CORTE_CAJA_DETALLE.IdAjuste)=0)"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY FORMA_PAGO.Descripcion, FORMA_PAGO.IdFormaPago"
        strSQL = strSQL & " ORDER BY FORMA_PAGO.IdFormaPago"
    #End If
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    
    
    Do While Not adorcs.EOF
    
        
        sReferencia = UCase(Trim(sPrefijo)) & Format(Me.dtpFecha.Value, "yymmdd") & Format(adorcs!IdFormaPago, "00")
        iDigVer = dvAlgoritmo35(sReferencia)
        sReferencia = sReferencia & iDigVer
    
    
    
        strSQL = "INSERT INTO TMP_VALIDA_CORTE_DIA ("
        strSQL = strSQL & "Identificador" & ","
        strSQL = strSQL & "FormaPago" & ","
        strSQL = strSQL & "Importe" & ","
        strSQL = strSQL & "Referencia" & ")"
        strSQL = strSQL & "Values" & "("
        strSQL = strSQL & "'" & sIdent & "',"
        strSQL = strSQL & "'" & adorcs!Descripcion & "',"
        strSQL = strSQL & adorcs!Importe & ","
        strSQL = strSQL & "'" & sReferencia & "')"
        
        adocmd.CommandText = strSQL
        adocmd.Execute
        
        adorcs.MoveNext
    Loop
    
    adorcs.Close
    
    
    
    strSQL = "SELECT FormaPago, Importe, Referencia, '" & sCuenta & "' AS CUENTA, " & "'" & Me.dtpFecha.Value & "' As Fecha"
    strSQL = strSQL & " FROM TMP_VALIDA_CORTE_DIA"
    strSQL = strSQL & " Where ("
    strSQL = strSQL & "((Identificador)='" & sIdent & "')"
    strSQL = strSQL & ")"
    
    
    Set frmRep = New frmReportViewer
    
    
    frmRep.sQuery = strSQL
    
    frmRep.sNombreReporte = sDB_ReportSource & "\" & "cc_dia.rpt"
    
    frmRep.Show vbModal
    
    #If SqlServer_ Then
        strSQL = "DELETE FROM TMP_VALIDA_CORTE_DIA"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " Identificador='" & sIdent & "'"
    #Else
        strSQL = "DELETE * FROM TMP_VALIDA_CORTE_DIA"
        strSQL = strSQL & " Where ("
        strSQL = strSQL & "((Identificador)='" & sIdent & "')"
        strSQL = strSQL & ")"
    #End If
    
    adocmd.CommandText = strSQL
    adocmd.Execute
    Set adocmd = Nothing
    
    
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    CentraForma MDIPrincipal, Me
    
    Me.dtpFecha.Value = Date
End Sub
