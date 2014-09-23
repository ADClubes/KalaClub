VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCCajaValida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cortes de caja"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdValidaCorte 
      Caption         =   "Valida Corte"
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCorteCajaSup 
      Caption         =   "Reporte Supervisor"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCorteCajaRep 
      Caption         =   "Reporte cajero"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgCortes 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   2540
      Columns(0).Caption=   "IdCorteCaja"
      Columns(0).Name =   "IdCorteCaja"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "FechaCorte"
      Columns(1).Name =   "FechaCorte"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "HoraCorte"
      Columns(2).Name =   "HoraCorte"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "UsuarioCorte"
      Columns(3).Name =   "UsuarioCorte"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1640
      Columns(4).Caption=   "Caja"
      Columns(4).Name =   "Caja"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1773
      Columns(5).Caption=   "Turno"
      Columns(5).Name =   "Turno"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   17383
      _ExtentY        =   4260
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
Attribute VB_Name = "frmCCajaValida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCorteCajaRep_Click()
    
    Dim frmreporte As frmReportViewer
    
    
    
    strSQL = "SELECT FORMA_PAGO.Descripcion, CT_TPVS.IdTPV, CT_TPVS.IdInterno, CT_TPVS.DescripcionTPV, CORTE_CAJA_DETALLE.LoteNumero, CORTE_CAJA_DETALLE.Importe, CORTE_CAJA_DETALLE.NumeroOperaciones, CORTE_CAJA_DETALLE.Renglon, CORTE_CAJA.IdCorteCaja, CORTE_CAJA.FechaCorte, CORTE_CAJA.HoraCorte, CORTE_CAJA.UsuarioCorte, CORTE_CAJA.FechaOperacion, CORTE_CAJA.Caja, CORTE_CAJA.Turno, CORTE_CAJA.FolioInicial, CORTE_CAJA.FolioFinal"
    strSQL = strSQL & " FROM ((CORTE_CAJA_DETALLE INNER JOIN CORTE_CAJA ON CORTE_CAJA_DETALLE.IdCorteCaja = CORTE_CAJA.IdCorteCaja) INNER JOIN FORMA_PAGO ON CORTE_CAJA_DETALLE.IdFormaPago = FORMA_PAGO.IdFormaPago) LEFT JOIN CT_TPVS ON CORTE_CAJA_DETALLE.IdTPV = CT_TPVS.IdTPV"
    strSQL = strSQL & " Where (((CORTE_CAJA.IdCorteCaja) = " & Me.ssdbgCortes.Columns("IdCorteCaja").Value & "))"
    strSQL = strSQL & " ORDER BY CORTE_CAJA_DETALLE.Renglon"

    
    Set frmreporte = New frmReportViewer
    
    
    frmreporte.sNombreReporte = sDB_ReportSource & "\" & "cc_cajero.rpt"
    frmreporte.sQuery = strSQL
    
    frmreporte.Show vbModal
End Sub

Private Sub cmdCorteCajaSup_Click()
    Dim adocmdCorte As ADODB.Command
    Dim adorcsCorte As ADODB.Recordset
    Dim adorcsCompara As ADODB.Recordset
    
    
    Dim sIdent As String
    
    
    
    Dim frmRep As frmReportViewer
    
    
    sIdent = Format(Now, "ddmmHhNnSs")
    
    
    Set adocmdCorte = New ADODB.Command
    adocmdCorte.ActiveConnection = Conn
    adocmdCorte.CommandType = adCmdText
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO TMP_CORTE_CAJA ("
        strSQL = strSQL & " IdCorteCaja,"
        strSQL = strSQL & " IdFormaPago,"
        strSQL = strSQL & " IdTPV,"
        strSQL = strSQL & " LoteNumero,"
        strSQL = strSQL & " ImporteOperado,"
        strSQL = strSQL & " ImporteCorte,"
        strSQL = strSQL & " Identificador)"
        strSQL = strSQL & " SELECT FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, ISNULL(CT_AFILIACIONES.IdTerminal, 0) AS IdTPV, PAGOS_FACTURA.LoteNumero, Sum(PAGOS_FACTURA.Importe) AS ImporteOperado, 0 AS ImporteCorte, '" & sIdent & "' AS Ident"
        strSQL = strSQL & " FROM (PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura) LEFT JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FACTURAS.IdCorteCaja)=" & Me.ssdbgCortes.Columns("IdCorteCaja").Value & "))"
        strSQL = strSQL & " GROUP BY FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, CT_AFILIACIONES.IdTerminal, PAGOS_FACTURA.LoteNumero"
    #Else
        strSQL = "INSERT INTO TMP_CORTE_CAJA ("
        strSQL = strSQL & " IdCorteCaja,"
        strSQL = strSQL & " IdFormaPago,"
        strSQL = strSQL & " IdTPV,"
        strSQL = strSQL & " LoteNumero,"
        strSQL = strSQL & " ImporteOperado,"
        strSQL = strSQL & " ImporteCorte,"
        strSQL = strSQL & " Identificador)"
        strSQL = strSQL & " SELECT FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, iif(isnull(CT_AFILIACIONES.IdTerminal),0,CT_AFILIACIONES.IdTerminal) AS IdTPV, PAGOS_FACTURA.LoteNumero, Sum(PAGOS_FACTURA.Importe) AS ImporteOperado, 0 AS ImporteCorte, '" & sIdent & "' AS Ident"
        strSQL = strSQL & " FROM (PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura) LEFT JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FACTURAS.IdCorteCaja)=" & Me.ssdbgCortes.Columns("IdCorteCaja").Value & "))"
        strSQL = strSQL & " GROUP BY FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, CT_AFILIACIONES.IdTerminal, PAGOS_FACTURA.LoteNumero"
    #End If
    
    adocmdCorte.CommandText = strSQL
    adocmdCorte.Execute
    
    
    
    
    strSQL = "SELECT CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, CORTE_CAJA_DETALLE.LoteNumero, CORTE_CAJA_DETALLE.Importe"
    strSQL = strSQL & " FROM CORTE_CAJA_DETALLE INNER JOIN CORTE_CAJA ON CORTE_CAJA_DETALLE.IdCorteCaja = CORTE_CAJA.IdCorteCaja"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((CORTE_CAJA.IdCorteCaja)=" & Me.ssdbgCortes.Columns("IdCorteCaja").Value & "))"
    
    
    Set adorcsCorte = New ADODB.Recordset
    adorcsCorte.CursorLocation = adUseServer
    
    adorcsCorte.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Set adorcsCompara = New ADODB.Recordset
    adorcsCompara.CursorLocation = adUseServer
    
    Do While Not adorcsCorte.EOF
        
        strSQL = "SELECT ImporteCorte"
        strSQL = strSQL & " FROM TMP_CORTE_CAJA"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((IdFormaPago)=" & adorcsCorte!IdFormaPago & ")"
        strSQL = strSQL & " AND ((IdTPV)=" & adorcsCorte!IdTPV & ")"
        strSQL = strSQL & " AND ((LoteNumero)='" & adorcsCorte!LoteNumero & "')"
        strSQL = strSQL & " AND ((Identificador)='" & sIdent & "')"
        strSQL = strSQL & ")"
        
        adorcsCompara.Open strSQL, Conn, adOpenForwardOnly, adLockOptimistic
        
        
        If Not adorcsCompara.EOF Then
            adorcsCompara!ImporteCorte = adorcsCorte!Importe
            adorcsCompara.Update
            
        Else
            strSQL = "INSERT INTO TMP_CORTE_CAJA ("
            strSQL = strSQL & " IdCorteCaja,"
            strSQL = strSQL & " IdFormaPago,"
            strSQL = strSQL & " IdTPV,"
            strSQL = strSQL & " LoteNumero,"
            strSQL = strSQL & " ImporteCorte,"
            strSQL = strSQL & " Identificador)"
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & Me.ssdbgCortes.Columns("IdCorteCaja").Value & ","
            strSQL = strSQL & adorcsCorte!IdFormaPago & ","
            strSQL = strSQL & adorcsCorte!IdTPV & ","
            strSQL = strSQL & "'" & adorcsCorte!LoteNumero & "',"
            strSQL = strSQL & adorcsCorte!Importe & ","
            strSQL = strSQL & "'" & sIdent & "')"
            
            
            adocmdCorte.CommandText = strSQL
            adocmdCorte.Execute
            
        End If
        
        adorcsCompara.Close
        
        adorcsCorte.MoveNext
    Loop
    
    
    Set adorcsCompara = Nothing
    
    adorcsCorte.Close
    Set adorcsCorte = Nothing
    
    
    
   
    
    
    strSQL = "SELECT TMP_CORTE_CAJA.IdCorteCaja, CORTE_CAJA.FechaCorte, CORTE_CAJA.HoraCorte, CORTE_CAJA.UsuarioCorte, CORTE_CAJA.Caja, CORTE_CAJA.Turno, CORTE_CAJA.FolioInicial, CORTE_CAJA.FolioFinal, CORTE_CAJA.FondoInicial, TMP_CORTE_CAJA.IdFormaPago, FORMA_PAGO.Descripcion, TMP_CORTE_CAJA.IdTPV, CT_TPVS.DescripcionTPV, TMP_CORTE_CAJA.LoteNumero, TMP_CORTE_CAJA.ImporteOperado, TMP_CORTE_CAJA.ImporteCorte"
    strSQL = strSQL & " FROM ((TMP_CORTE_CAJA INNER JOIN CORTE_CAJA ON TMP_CORTE_CAJA.IdCorteCaja = CORTE_CAJA.IdCorteCaja) INNER JOIN FORMA_PAGO ON TMP_CORTE_CAJA.IdFormaPago = FORMA_PAGO.IdFormaPago) LEFT JOIN CT_TPVS ON TMP_CORTE_CAJA.IdTPV = CT_TPVS.IdTPV"
    strSQL = strSQL & " WHERE (((TMP_CORTE_CAJA.Identificador) = '" & sIdent & "'))"
    strSQL = strSQL & " ORDER BY TMP_CORTE_CAJA.IdFormaPago, TMP_CORTE_CAJA.IdTPV, TMP_CORTE_CAJA.LoteNumero;"

    
    Set frmRep = New frmReportViewer
    
    
    frmRep.sQuery = strSQL
    frmRep.sNombreReporte = sDB_ReportSource & "/" & "cc_supervisor.rpt"
    
    frmRep.Show vbModal
    
    #If SqlServer_ Then
        strSQL = "DELETE FROM TMP_CORTE_CAJA"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((TMP_CORTE_CAJA.Identificador)='" & sIdent & "')"
        strSQL = strSQL & ")"
    #Else
        strSQL = "DELETE * FROM TMP_CORTE_CAJA"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((TMP_CORTE_CAJA.Identificador)='" & sIdent & "')"
        strSQL = strSQL & ")"
    #End If
    
    adocmdCorte.CommandText = strSQL
    adocmdCorte.Execute
    
    Set adocmdCorte = Nothing
    
    
End Sub

Private Sub Form_Load()
    
    strSQL = "SELECT CORTE_CAJA.IdCorteCaja, CORTE_CAJA.FechaCorte, CORTE_CAJA.HoraCorte, CORTE_CAJA.Caja, CORTE_CAJA.Turno"
    strSQL = strSQL & " From CORTE_CAJA"
    strSQL = strSQL & " WHERE (((CORTE_CAJA.Status)=0))"
    
    
    LlenaSsDbGrid Me.ssdbgCortes, Conn, strSQL, 5

End Sub
