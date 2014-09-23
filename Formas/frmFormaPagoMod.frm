VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmFormaPagoMod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Validación de cortes de caja por turno"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   13890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReporteSuper 
      Caption         =   "Reporte supervisor"
      Height          =   495
      Left            =   12360
      TabIndex        =   22
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdReporteCajero 
      Caption         =   "Reporte cajero"
      Height          =   495
      Left            =   10800
      TabIndex        =   21
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCambiaDetalle 
      Caption         =   "Modifica Operacion"
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   6240
      Width           =   1095
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgPagosDetalle 
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   13455
      _Version        =   196616
      DataMode        =   2
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   15
      Columns(0).Width=   1799
      Columns(0).Caption=   "NumeroFactura"
      Columns(0).Name =   "NumeroFactura"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1746
      Columns(1).Caption=   "FolioFactura"
      Columns(1).Name =   "FolioFactura"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   979
      Columns(2).Caption=   "Serie"
      Columns(2).Name =   "Serie"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1296
      Columns(3).Caption=   "Renglon"
      Columns(3).Name =   "Renglon"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1852
      Columns(4).Caption=   "IdFormaPago"
      Columns(4).Name =   "IdFormaPago"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3387
      Columns(5).Caption=   "Descripcion"
      Columns(5).Name =   "Descripcion"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   1217
      Columns(6).Caption=   "IdTPV"
      Columns(6).Name =   "IdTPV"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1879
      Columns(7).Caption=   "Afiliacion"
      Columns(7).Name =   "Afiliacion"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   1508
      Columns(8).Caption=   "TPV"
      Columns(8).Name =   "TPV"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   1852
      Columns(9).Caption=   "Modalidad"
      Columns(9).Name =   "Modalidad"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      Columns(10).Width=   2117
      Columns(10).Caption=   "LoteNumero"
      Columns(10).Name=   "LoteNumero"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      Columns(11).Width=   1931
      Columns(11).Caption=   "OperacionNumero"
      Columns(11).Name=   "OperacionNumero"
      Columns(11).DataField=   "Column 11"
      Columns(11).DataType=   8
      Columns(11).FieldLen=   256
      Columns(12).Width=   3200
      Columns(12).Caption=   "FechaOperacion"
      Columns(12).Name=   "FechaOperacion"
      Columns(12).DataField=   "Column 12"
      Columns(12).DataType=   8
      Columns(12).FieldLen=   256
      Columns(13).Width=   3200
      Columns(13).Caption=   "Importe"
      Columns(13).Name=   "Importe"
      Columns(13).DataField=   "Column 13"
      Columns(13).DataType=   8
      Columns(13).FieldLen=   256
      Columns(14).Width=   2646
      Columns(14).Caption=   "ImporteRecibido"
      Columns(14).Name=   "ImporteRecibido"
      Columns(14).DataField=   "Column 14"
      Columns(14).DataType=   8
      Columns(14).FieldLen=   256
      _ExtentX        =   23733
      _ExtentY        =   2778
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
   Begin VB.CommandButton cmdValidar 
      Caption         =   "Validar corte"
      Height          =   495
      Left            =   7800
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgregarCorte 
      Caption         =   "Agregar al corte"
      Height          =   495
      Left            =   5400
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminarCorte 
      Caption         =   "Anular renglon"
      Height          =   495
      Left            =   4200
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdVerDetalle 
      Caption         =   "Ver detalle"
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Funcion"
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   7815
      Begin VB.CommandButton cmdModifica 
         Caption         =   "Modificar"
         Height          =   495
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Totales"
      Height          =   855
      Left            =   8280
      TabIndex        =   8
      Top             =   7080
      Width           =   5415
      Begin VB.Label lblTotal 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblRenglones 
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de busqueda"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4560
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFechaBusca 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   64290817
         CurrentDate     =   39879
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Turno"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Caja"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgPagos 
      Height          =   1695
      Left            =   0
      TabIndex        =   13
      Top             =   1320
      Width           =   13575
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   9
      AllowUpdate     =   0   'False
      AllowColumnMoving=   0
      AllowColumnSwapping=   0
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   9
      Columns(0).Width=   1905
      Columns(0).Caption=   "IdFormaPago"
      Columns(0).Name =   "IdFormaPago"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3836
      Columns(1).Caption=   "Descripcion"
      Columns(1).Name =   "Descripcion"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1376
      Columns(2).Caption=   "IdTPV"
      Columns(2).Name =   "IdTPV"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Descripcion TPV"
      Columns(3).Name =   "Descripcion TPV"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2725
      Columns(4).Caption=   "LoteNumero"
      Columns(4).Name =   "LoteNumero"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "ImporteOperado"
      Columns(5).Name =   "ImporteOperado"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   6
      Columns(5).NumberFormat=   "CURRENCY"
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "ImporteCorte"
      Columns(6).Name =   "ImporteCorte"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   6
      Columns(6).NumberFormat=   "CURRENCY"
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Caption=   "ImporteAjuste"
      Columns(7).Name =   "ImporteAjuste"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   6
      Columns(7).NumberFormat=   "CURRENCY"
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "Diferencia"
      Columns(8).Name =   "Diferencia"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   6
      Columns(8).NumberFormat=   "CURRENCY"
      Columns(8).FieldLen=   256
      _ExtentX        =   23945
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
   Begin VB.Label lblDiferenciaTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   23
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblCorteNumero 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   19
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmFormaPagoMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lIdCorte As Long
Dim dDiferenciaTotal As Double

Private Sub cmdAgregaDetalle_Click()
    

    
    
End Sub

Private Sub cmdAgregarCorte_Click()
    
    Dim frmInsCorte As frmInsFpagoCorte
    
    
    If Me.ssdbgPagos.Columns("Diferencia").Value = 0 Then
        If MsgBox("Este renglón no tiene diferencias" & vbCrLf & "¿proceder de cualquier forma?", vbQuestion + vbOKCancel, "Confirme") = vbCancel Then
            Exit Sub
        End If
    End If
    
    
    Set frmInsCorte = New frmInsFpagoCorte
    
    frmInsCorte.lIdCorteCaja = lIdCorte
    frmInsCorte.iNumeroCaja = Val(Me.txtCtrl(0).Text)
    frmInsCorte.dImporte = Val(Me.ssdbgPagos.Columns("Diferencia").Value)
    
    frmInsCorte.Show vbModal
    
    ActualizaGrid lIdCorte
    
    
End Sub

Private Sub cmdBusca_Click()
    
    If Me.txtCtrl(0).Text = vbNullString Then
        MsgBox "Indicar número de caja", vbExclamation, "Verifique"
        Me.txtCtrl(0).SetFocus
        Exit Sub
    End If
    
    If Me.txtCtrl(1).Text = vbNullString Then
        MsgBox "Indicar número de turno", vbExclamation, "Verifique"
        Me.txtCtrl(1).SetFocus
        Exit Sub
    End If
    
    lIdCorte = BuscaIdCorte(Me.dtpFechaBusca.Value, Val(Me.txtCtrl(0).Text), Val(Me.txtCtrl(1).Text))
    
    
    If lIdCorte <= 0 Then
        MsgBox "No existen datos con los parámetros establecidos", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    ActualizaGrid lIdCorte
    
    If Me.ssdbgPagos.Rows = 0 Then
        MsgBox "No existen datos con los parámetros establecidos", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    'Me.Frame2.Visible = True
    'Me.Frame3.Visible = True
    
    
End Sub

Private Sub ActualizaGrid(lIdCorte As Long)

    Dim sIdent As String
    
    Dim adocmdCorte As ADODB.Command
    Dim adorcsCorte As ADODB.Recordset
    Dim adorcsCompara As ADODB.Recordset
    
    Dim dTotalOperado As Double
    Dim dTotalCorte As Double
    
    
    Screen.MousePointer = vbHourglass
    
    sIdent = Format(Now, "ddmmHhNnSs")
    
        
    Me.lblCorteNumero.Caption = "Corte # " & lIdCorte
    Me.lblCorteNumero.Visible = True
    
    
    Set adocmdCorte = New ADODB.Command
    adocmdCorte.ActiveConnection = Conn
    adocmdCorte.CommandType = adCmdText
    
    
    'Procesa el corte conforme a los datos indicados
    'Llena la tabla temporal con los datos operados
    #If SqlServer_ Then
        strSQL = "INSERT INTO TMP_CORTE_CAJA ("
        strSQL = strSQL & " IdCorteCaja,"
        strSQL = strSQL & " IdFormaPago,"
        strSQL = strSQL & " IdTPV,"
        strSQL = strSQL & " LoteNumero,"
        strSQL = strSQL & " ImporteOperado,"
        strSQL = strSQL & " ImporteCorte,"
        strSQL = strSQL & " Identificador)"
        strSQL = strSQL & " SELECT FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, ISNULL(CT_AFILIACIONES.IdTerminal, '0') AS IdTPV, ISNULL(PAGOS_FACTURA.LoteNumero, 0), Sum(PAGOS_FACTURA.Importe) AS ImporteOperado, 0 AS ImporteCorte, '" & sIdent & "' AS Ident"
        strSQL = strSQL & " FROM (PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura) LEFT JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FACTURAS.IdCorteCaja)=" & lIdCorte & ")"
        strSQL = strSQL & " AND ((FACTURAS.Cancelada)=0" & ")"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, ISNULL(CT_AFILIACIONES.IdTerminal, '0'), ISNULL(PAGOS_FACTURA.LoteNumero, 0)"
    #Else
        strSQL = "INSERT INTO TMP_CORTE_CAJA ("
        strSQL = strSQL & " IdCorteCaja,"
        strSQL = strSQL & " IdFormaPago,"
        strSQL = strSQL & " IdTPV,"
        strSQL = strSQL & " LoteNumero,"
        strSQL = strSQL & " ImporteOperado,"
        strSQL = strSQL & " ImporteCorte,"
        strSQL = strSQL & " Identificador)"
        strSQL = strSQL & " SELECT FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, iif(isnull(CT_AFILIACIONES.IdTerminal),'0',CT_AFILIACIONES.IdTerminal) AS IdTPV, iif(IsNull(PAGOS_FACTURA.LoteNumero),0,PAGOS_FACTURA.LoteNumero), Sum(PAGOS_FACTURA.Importe) AS ImporteOperado, 0 AS ImporteCorte, '" & sIdent & "' AS Ident"
        strSQL = strSQL & " FROM (PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura) LEFT JOIN CT_AFILIACIONES ON PAGOS_FACTURA.IdAfiliacion = CT_AFILIACIONES.IdAfiliacion"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FACTURAS.IdCorteCaja)=" & lIdCorte & ")"
        strSQL = strSQL & " AND ((FACTURAS.Cancelada)=0" & ")"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY FACTURAS.IdCorteCaja, PAGOS_FACTURA.IdFormaPago, iif(isnull(CT_AFILIACIONES.IdTerminal),'0',CT_AFILIACIONES.IdTerminal), iif(IsNull(PAGOS_FACTURA.LoteNumero),0,PAGOS_FACTURA.LoteNumero)"
    #End If
    
    adocmdCorte.CommandText = strSQL
    adocmdCorte.Execute
    
    
    
    'Selecciona lo capturado en el corte, solo lo capturado por el cajero
    #If SqlServer_ Then
        strSQL = "SELECT CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, ISNULL(CORTE_CAJA_DETALLE.LoteNumero, '0') AS LoteNumero, CORTE_CAJA.Status, Sum(CORTE_CAJA_DETALLE.Importe) AS Importe "
        strSQL = strSQL & " FROM CORTE_CAJA_DETALLE INNER JOIN CORTE_CAJA ON CORTE_CAJA_DETALLE.IdCorteCaja = CORTE_CAJA.IdCorteCaja"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((CORTE_CAJA.IdCorteCaja)=" & lIdCorte & ")"
        strSQL = strSQL & " AND ((CORTE_CAJA_DETALLE.IdAjuste)=" & "0)"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, ISNULL(CORTE_CAJA_DETALLE.LoteNumero, '0'), CORTE_CAJA.Status"
    #Else
        strSQL = "SELECT CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, iif(IsNull(CORTE_CAJA_DETALLE.LoteNumero),'0',CORTE_CAJA_DETALLE.LoteNumero) AS LoteNumero, CORTE_CAJA.Status, Sum(CORTE_CAJA_DETALLE.Importe) AS Importe "
        strSQL = strSQL & " FROM CORTE_CAJA_DETALLE INNER JOIN CORTE_CAJA ON CORTE_CAJA_DETALLE.IdCorteCaja = CORTE_CAJA.IdCorteCaja"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((CORTE_CAJA.IdCorteCaja)=" & lIdCorte & ")"
        strSQL = strSQL & " AND ((CORTE_CAJA_DETALLE.IdAjuste)=" & "0)"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, iif(IsNull(CORTE_CAJA_DETALLE.LoteNumero),'0',CORTE_CAJA_DETALLE.LoteNumero), CORTE_CAJA.Status"
    #End If
    
    Set adorcsCorte = New ADODB.Recordset
    adorcsCorte.CursorLocation = adUseServer
    
    adorcsCorte.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsCorte.EOF Then
        If adorcsCorte!Status = 0 Then
            Me.cmdValidar.Enabled = True
        Else
            Me.cmdValidar.Enabled = False
        End If
    End If
    
    Set adorcsCompara = New ADODB.Recordset
    adorcsCompara.CursorLocation = adUseServer
    
    Dim cmdUpdate As Object
    'Compara lo operado con lo capturado en el corte
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
            #If SqlServer_ Then
                Set cmdUpdate = New ADODB.Command
                cmdUpdate.ActiveConnection = Conn
                cmdUpdate.CommandType = adCmdText
                
                strSQL = "UPDATE TMP_CORTE_CAJA SET ImporteCorte = " & adorcsCorte!Importe
                strSQL = strSQL & " WHERE"
                strSQL = strSQL & " IdFormaPago=" & adorcsCorte!IdFormaPago
                strSQL = strSQL & " AND IdTPV=" & adorcsCorte!IdTPV
                strSQL = strSQL & " AND LoteNumero='" & adorcsCorte!LoteNumero & "'"
                strSQL = strSQL & " AND Identificador='" & sIdent & "'"
                
                cmdUpdate.CommandText = strSQL
                cmdUpdate.Execute
            #Else
                adorcsCompara!ImporteCorte = adorcsCorte!Importe
                adorcsCompara.Update
            #End If
        Else
            strSQL = "INSERT INTO TMP_CORTE_CAJA ("
            strSQL = strSQL & " IdCorteCaja,"
            strSQL = strSQL & " IdFormaPago,"
            strSQL = strSQL & " IdTPV,"
            strSQL = strSQL & " LoteNumero,"
            strSQL = strSQL & " ImporteCorte,"
            strSQL = strSQL & " Identificador)"
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & lIdCorte & ","
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
    
    adorcsCorte.Close
    
    
    'Selecciona lo capturado en el corte, solo lo capturado por el supervisor
    #If SqlServer_ Then
        strSQL = "SELECT CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, ISNULL(CORTE_CAJA_DETALLE.LoteNumero, '0') AS LoteNumero, CORTE_CAJA.Status, Sum(CORTE_CAJA_DETALLE.Importe) AS Importe "
        strSQL = strSQL & " FROM CORTE_CAJA_DETALLE INNER JOIN CORTE_CAJA ON CORTE_CAJA_DETALLE.IdCorteCaja = CORTE_CAJA.IdCorteCaja"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((CORTE_CAJA.IdCorteCaja)=" & lIdCorte & ")"
        strSQL = strSQL & " AND ((CORTE_CAJA_DETALLE.IdAjuste) <>" & "0)"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, ISNULL(CORTE_CAJA_DETALLE.LoteNumero, '0'), CORTE_CAJA.Status"
    #Else
        strSQL = "SELECT CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, iif(IsNull(CORTE_CAJA_DETALLE.LoteNumero),'0',CORTE_CAJA_DETALLE.LoteNumero) AS LoteNumero, CORTE_CAJA.Status, Sum(CORTE_CAJA_DETALLE.Importe) AS Importe "
        strSQL = strSQL & " FROM CORTE_CAJA_DETALLE INNER JOIN CORTE_CAJA ON CORTE_CAJA_DETALLE.IdCorteCaja = CORTE_CAJA.IdCorteCaja"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((CORTE_CAJA.IdCorteCaja)=" & lIdCorte & ")"
        strSQL = strSQL & " AND ((CORTE_CAJA_DETALLE.IdAjuste) <>" & "0)"
        strSQL = strSQL & ")"
        strSQL = strSQL & " GROUP BY CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.IdTPV, iif(IsNull(CORTE_CAJA_DETALLE.LoteNumero),'0',CORTE_CAJA_DETALLE.LoteNumero), CORTE_CAJA.Status"
    #End If
    
    adorcsCorte.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    'Compara lo operado con lo capturado en el corte
    Do While Not adorcsCorte.EOF
        
        strSQL = "SELECT ImporteAjuste"
        strSQL = strSQL & " FROM TMP_CORTE_CAJA"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((IdFormaPago)=" & adorcsCorte!IdFormaPago & ")"
        strSQL = strSQL & " AND ((IdTPV)=" & adorcsCorte!IdTPV & ")"
        strSQL = strSQL & " AND ((LoteNumero)='" & adorcsCorte!LoteNumero & "')"
        strSQL = strSQL & " AND ((Identificador)='" & sIdent & "')"
        strSQL = strSQL & ")"
        
        adorcsCompara.Open strSQL, Conn, adOpenForwardOnly, adLockOptimistic
        
        
        If Not adorcsCompara.EOF Then
            #If SqlServer_ Then
                Set cmdUpdate = New ADODB.Command
                cmdUpdate.ActiveConnection = Conn
                cmdUpdate.CommandType = adCmdText
                
                strSQL = "UPDATE TMP_CORTE_CAJA SET ImporteCorte = " & adorcsCorte!Importe
                strSQL = strSQL & " WHERE"
                strSQL = strSQL & " IdFormaPago=" & adorcsCorte!IdFormaPago
                strSQL = strSQL & " AND IdTPV=" & adorcsCorte!IdTPV
                strSQL = strSQL & " AND LoteNumero='" & adorcsCorte!LoteNumero & "'"
                strSQL = strSQL & " AND Identificador='" & sIdent & "'"
                
                cmdUpdate.CommandText = strSQL
                cmdUpdate.Execute
            #Else
                adorcsCompara!ImporteAjuste = adorcsCorte!Importe
                adorcsCompara.Update
            #End If
        End If
        adorcsCompara.Close
        adorcsCorte.MoveNext
    Loop
    
    
    
    
    
    Set adorcsCompara = Nothing
    
    adorcsCorte.Close
    Set adorcsCorte = Nothing
    
    
    
   
    
    
    strSQL = "SELECT TMP_CORTE_CAJA.IdFormaPago, FORMA_PAGO.Descripcion, TMP_CORTE_CAJA.IdTPV, CT_TPVS.DescripcionTPV, TMP_CORTE_CAJA.LoteNumero, TMP_CORTE_CAJA.ImporteOperado, TMP_CORTE_CAJA.ImporteCorte, TMP_CORTE_CAJA.ImporteAjuste , TMP_CORTE_CAJA.ImporteOperado - (TMP_CORTE_CAJA.ImporteCorte + TMP_CORTE_CAJA.ImporteAjuste) AS Diferencia ,TMP_CORTE_CAJA.IdCorteCaja ,CORTE_CAJA.FechaCorte, CORTE_CAJA.HoraCorte, CORTE_CAJA.UsuarioCorte, CORTE_CAJA.Caja, CORTE_CAJA.Turno, CORTE_CAJA.FolioInicial, CORTE_CAJA.FolioFinal, CORTE_CAJA.FondoInicial "
    strSQL = strSQL & " FROM ((TMP_CORTE_CAJA INNER JOIN CORTE_CAJA ON TMP_CORTE_CAJA.IdCorteCaja = CORTE_CAJA.IdCorteCaja) INNER JOIN FORMA_PAGO ON TMP_CORTE_CAJA.IdFormaPago = FORMA_PAGO.IdFormaPago) LEFT JOIN CT_TPVS ON TMP_CORTE_CAJA.IdTPV = CT_TPVS.IdTPV"
    strSQL = strSQL & " WHERE (((TMP_CORTE_CAJA.Identificador) = '" & sIdent & "'))"
    strSQL = strSQL & " ORDER BY TMP_CORTE_CAJA.IdFormaPago, TMP_CORTE_CAJA.IdTPV, TMP_CORTE_CAJA.LoteNumero;"
    
    
    LlenaSsDbGrid Me.ssdbgPagos, Conn, strSQL, 9
    
    
    Me.lblRenglones.Caption = Me.ssdbgPagos.Rows & " Renglon(es)"
    ActualizaTotal
    ActualizaDiferencia
    
    #If SqlServer_ Then
        strSQL = "DELETE FROM TMP_CORTE_CAJA"
        strSQL = strSQL & " WHERE TMP_CORTE_CAJA.Identificador = " & "'" & sIdent & "'"
    #Else
        strSQL = "DELETE * FROM TMP_CORTE_CAJA"
        strSQL = strSQL & " WHERE TMP_CORTE_CAJA.Identificador = " & "'" & sIdent & "'"
    #End If
    
    adocmdCorte.CommandText = strSQL
    adocmdCorte.Execute
    
    
    Set adocmdCorte = Nothing
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdEliminaDetalle_Click()
    
    Dim adocmd As ADODB.Command
    
    strSQL = "DELETE FROM PAGOS_FACTURA"
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " NumeroFactura =" & Me.ssdbgPagosDetalle.Columns("NumeroFactura").Value
    strSQL = strSQL & " AND Renglon =" & Me.ssdbgPagosDetalle.Columns("Renglon").Value
    
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    Set adocmd = Nothing
    
    Me.cmdVerDetalle.Value = True
    ActualizaGrid lIdCorte
    
End Sub

Private Sub cmdCambiaDetalle_Click()
    
    Dim frmModPago As frmInsFpago
    
    
    Set frmModPago = New frmInsFpago
    
    
    If Me.ssdbgPagosDetalle.Rows = 0 Then
        Exit Sub
    End If
    
    
    frmModPago.lNumeroFactura = Me.ssdbgPagosDetalle.Columns("NumeroFactura").Value
    frmModPago.lNumeroRenglon = Me.ssdbgPagosDetalle.Columns("Renglon").Value
    
    frmModPago.iNumeroCaja = Val(Me.txtCtrl(0).Text)
    
    frmModPago.lIdFormaPago = Me.ssdbgPagosDetalle.Columns("IdFormaPago").Value
    frmModPago.lIdTPV = IIf(Me.ssdbgPagosDetalle.Columns("IdTPV").Value = vbNullString, 0, Me.ssdbgPagosDetalle.Columns("IdTPV").Value)
    frmModPago.lIdAfiliacion = IIf(Me.ssdbgPagosDetalle.Columns("Afiliacion").Value = vbNullString, 0, Me.ssdbgPagosDetalle.Columns("Afiliacion").Value)
    
    frmModPago.sLoteNumero = Me.ssdbgPagosDetalle.Columns("LoteNumero").Value
    frmModPago.sOperacion = Me.ssdbgPagosDetalle.Columns("OperacionNumero").Value
    frmModPago.dImporte = Me.ssdbgPagosDetalle.Columns("Importe").Value
    frmModPago.dImporteRecibido = Me.ssdbgPagosDetalle.Columns("ImporteRecibido").Value
    frmModPago.dFechaOPeracion = Me.ssdbgPagosDetalle.Columns("FechaOperacion").Value
    
    frmModPago.Show vbModal
    
    Me.cmdVerDetalle.Value = True
    
    ActualizaGrid lIdCorte
    Me.ssdbgPagosDetalle.RemoveAll
    
End Sub

Private Sub cmdEliminarCorte_Click()
    strSQL = strSQL & "INSERT INTO CORTE_CAJA_DETALLE"
    strSQL = " SELECT CORTE_CAJA_DETALLE.IdCorteCaja, 10, CORTE_CAJA_DETALLE.IdFormaPago, CORTE_CAJA_DETALLE.OpcionPago, CORTE_CAJA_DETALLE.Importe * -1, CORTE_CAJA_DETALLE.Referencia, CORTE_CAJA_DETALLE.IdTPV, CORTE_CAJA_DETALLE.LoteNumero, CORTE_CAJA_DETALLE.NumeroOperaciones *-1"
    strSQL = strSQL & " FROM CORTE_CAJA_DETALLE"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "(CORTE_CAJA_DETALLE.IdCorteCaja) =" & Me.ssdbgPagos.Columns("IdCorteCaja")
    strSQL = strSQL & " AND CORTE_CAJA_DETALLE.Renglon ="
End Sub

Private Sub cmdReporteCajero_Click()
    '06/Dic/2011 UCM
    If ssdbgPagos.Rows < 1 Then
        Exit Sub
    End If
    
    ReportesCorte 0, lIdCorte
End Sub

Private Sub cmdReporteSuper_Click()
    ReportesCorte 1, lIdCorte
End Sub

Private Sub cmdValidar_Click()
    
    If dDiferenciaTotal > 0 Then
        If MsgBox("Existe diferencia en el corte!" & vbLf & "¿Continuar?", vbInformation + vbOKCancel, "Confirme") = vbCancel Then
            Exit Sub
        End If
    End If
    
    
    If ValidaCorte() Then
        MsgBox "Corte validado", vbInformation, "Ok"
    End If
End Sub

Private Sub cmdVerDetalle_Click()
    
    If Me.ssdbgPagos.Columns("IdFormaPago").Value = "" Or Me.ssdbgPagos.Columns("IdTPV").Value = "" Then Exit Sub
    
    #If SqlServer_ Then
        strSQL = "SELECT PAGOS_FACTURA.NumeroFactura, FACTURAS.FolioCFD, FACTURAS.SerieCFD, PAGOS_FACTURA.Renglon, PAGOS_FACTURA.IdFormaPago, FORMA_PAGO.Descripcion, TPV.IdTerminal, PAGOS_FACTURA.IdAfiliacion, TPV.DescripcionTPV, TPV.Modalidad, PAGOS_FACTURA.LoteNumero, PAGOS_FACTURA.OperacionNumero, PAGOS_FACTURA.FechaOperacion, PAGOS_FACTURA.Importe, PAGOS_FACTURA.ImporteRecibido"
        strSQL = strSQL & " FROM PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura INNER JOIN FORMA_PAGO ON PAGOS_FACTURA.IdFormaPago = FORMA_PAGO.IdFormaPago LEFT JOIN (SELECT CT_Afiliaciones.IdTerminal, CT_Afiliaciones.IdAfiliacion,CT_TPVS.DescripcionTPV, CT_AFILIACIONES.Modalidad"
        strSQL = strSQL & " FROM CT_AFILIACIONES INNER JOIN CT_TPVS ON CT_AFILIACIONES.IdTerminal = CT_TPVS.IdTPV) AS TPV ON PAGOS_FACTURA.IdAfiliacion = TPV.IdAfiliacion"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " FACTURAS.FechaFactura = '" & Format(Me.dtpFechaBusca.Value, "yyyymmdd") & "'"
        strSQL = strSQL & " And FACTURAS.Cancelada = 0"
        strSQL = strSQL & " And FACTURAS.Caja =" & Trim(Me.txtCtrl(0).Text)
        strSQL = strSQL & " And FACTURAS.Turno =" & Trim(Me.txtCtrl(1).Text)
        strSQL = strSQL & " And PAGOS_FACTURA.IdFormaPago =" & Me.ssdbgPagos.Columns("IdFormaPago").Value
        If Me.ssdbgPagos.Columns("IdTPV").Value <> 0 Then
            strSQL = strSQL & " And TPV.IdTerminal =" & Me.ssdbgPagos.Columns("IdTPV").Value
        End If
        If Me.ssdbgPagos.Columns("LoteNumero").Value <> vbNullString Then
            strSQL = strSQL & " And PAGOS_FACTURA.LoteNumero =" & "'" & Me.ssdbgPagos.Columns("LoteNumero").Value & "'"
        End If
        
        strSQL = strSQL & " ORDER BY PAGOS_FACTURA.NumeroFactura, PAGOS_FACTURA.Renglon"
    #Else
        strSQL = "SELECT PAGOS_FACTURA.NumeroFactura, FACTURAS.FolioCFD, FACTURAS.SerieCFD,  PAGOS_FACTURA.Renglon, PAGOS_FACTURA.IdFormaPago, FORMA_PAGO.Descripcion, TPV.IdTerminal, PAGOS_FACTURA.IdAfiliacion, TPV.DescripcionTPV, TPV.Modalidad, PAGOS_FACTURA.LoteNumero, PAGOS_FACTURA.OperacionNumero, PAGOS_FACTURA.FechaOperacion, PAGOS_FACTURA.Importe, PAGOS_FACTURA.ImporteRecibido"
        strSQL = strSQL & " FROM ((PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura) INNER JOIN FORMA_PAGO ON PAGOS_FACTURA.IdFormaPago = FORMA_PAGO.IdFormaPago) LEFT JOIN [SELECT CT_Afiliaciones.IdTerminal, CT_Afiliaciones.IdAfiliacion,CT_TPVS.DescripcionTPV, CT_AFILIACIONES.Modalidad"
        strSQL = strSQL & " FROM CT_AFILIACIONES INNER JOIN CT_TPVS ON CT_AFILIACIONES.IdTerminal = CT_TPVS.IdTPV]. AS TPV ON PAGOS_FACTURA.IdAfiliacion = TPV.IdAfiliacion"
        strSQL = strSQL & " WHERE("
        strSQL = strSQL & " ((FACTURAS.FechaFactura) = #" & Format(Me.dtpFechaBusca.Value, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " And ((FACTURAS.Cancelada) = False)"
        strSQL = strSQL & " And ((FACTURAS.Caja) =" & Trim(Me.txtCtrl(0).Text) & ")"
        strSQL = strSQL & " And ((FACTURAS.Turno) =" & Trim(Me.txtCtrl(1).Text) & ")"
        strSQL = strSQL & " And ((PAGOS_FACTURA.IdFormaPago) =" & Me.ssdbgPagos.Columns("IdFormaPago").Value & ")"
        If Me.ssdbgPagos.Columns("IdTPV").Value <> 0 Then
            strSQL = strSQL & " And ((CT_AFILIACIONES.IdTerminal) =" & Me.ssdbgPagos.Columns("IdTPV").Value & ")"
        End If
        If Me.ssdbgPagos.Columns("LoteNumero").Value <> vbNullString Then
            strSQL = strSQL & " And ((PAGOS_FACTURA.LoteNumero) =" & "'" & Me.ssdbgPagos.Columns("LoteNumero").Value & "')"
        End If
        
        strSQL = strSQL & ")"
        strSQL = strSQL & " ORDER BY PAGOS_FACTURA.NumeroFactura, PAGOS_FACTURA.Renglon"
    #End If
    
    LlenaSsDbGrid Me.ssdbgPagosDetalle, Conn, strSQL, 15
    
    
    
End Sub



Private Sub Form_Load()
    Me.dtpFechaBusca.Value = Date
    
    '06/Dic/2011 UCM
    cmdVerDetalle.Enabled = (sDB_NivelUser = 0)
    cmdAgregarCorte.Enabled = (sDB_NivelUser = 0)
    cmdValidar.Enabled = (sDB_NivelUser = 0)
    cmdReporteSuper.Enabled = (sDB_NivelUser = 0)
    cmdCambiaDetalle.Enabled = (sDB_NivelUser = 0)
End Sub

Private Sub ActualizaTotal()
    Dim adorcs As ADODB.Recordset
    
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    
    #If SqlServer_ Then
        strSQL = "SELECT Sum(PAGOS_FACTURA.Importe) AS Total"
        strSQL = strSQL & " FROM PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " FACTURAS.FechaFactura = '" & Format(Me.dtpFechaBusca.Value, "yyyymmdd") & "'"
        strSQL = strSQL & " AND FACTURAS.Cancelada=0"
        strSQL = strSQL & " AND FACTURAS.Caja=" & Trim(Me.txtCtrl(0).Text)
        strSQL = strSQL & " AND FACTURAS.Turno=" & Trim(Me.txtCtrl(1).Text)
    #Else
        strSQL = "SELECT Sum(PAGOS_FACTURA.Importe) AS Total"
        strSQL = strSQL & " FROM PAGOS_FACTURA INNER JOIN FACTURAS ON PAGOS_FACTURA.NumeroFactura = FACTURAS.NumeroFactura"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((FACTURAS.FechaFactura) = #" & Format(Me.dtpFechaBusca.Value, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((FACTURAS.Cancelada)=False)"
        strSQL = strSQL & " AND ((FACTURAS.Caja)=" & Trim(Me.txtCtrl(0).Text) & ")"
        strSQL = strSQL & " AND ((FACTURAS.Turno)=" & Trim(Me.txtCtrl(1).Text) & ")"
        strSQL = strSQL & ")"
    #End If
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not adorcs.EOF Then
        Me.lblTotal.Caption = "Total: " & Format(adorcs!Total, "$#,##0.00")
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    
  
    
End Sub




Private Function ValidaCorte() As Boolean
    
    Dim adoCmdValida As ADODB.Command
        
        
    ValidaCorte = False
        
    #If SqlServer_ Then
        strSQL = "UPDATE CORTE_CAJA SET"
        strSQL = strSQL & " FechaValidacion ='" & Format(Date, "yyyymmdd") & "',"
        strSQL = strSQL & " HoraValidacion =" & "'" & Format(Now(), "Hh:Nn:Ss") & "',"
        strSQL = strSQL & " UsuarioValidacion =" & "'" & sDB_User & "'" & ","
        strSQL = strSQL & " Status=1"
        strSQL = strSQL & " WHERE IdCorteCaja = " & lIdCorte
    #Else
        strSQL = "UPDATE CORTE_CAJA SET"
        strSQL = strSQL & " FechaValidacion =#" & Format(Date, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & " HoraValidacion =" & "'" & Format(Now(), "Hh:Nn:Ss") & "',"
        strSQL = strSQL & " UsuarioValidacion =" & "'" & sDB_User & "'" & ","
        strSQL = strSQL & " Status=1"
        strSQL = strSQL & " WHERE IdCorteCaja = " & lIdCorte
    #End If
    
    Set adoCmdValida = New ADODB.Command
    adoCmdValida.CommandType = adCmdText
    adoCmdValida.ActiveConnection = Conn
    adoCmdValida.CommandText = strSQL
    
    adoCmdValida.Execute
    
    
    ValidaCorte = True
    
    Set adoCmdValida = Nothing
    
    
End Function

Private Sub ssdbgPagos_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    If Val(LastRow) <> Me.ssdbgPagos.row Then
        Me.ssdbgPagosDetalle.RemoveAll
    End If
End Sub

Private Sub ActualizaDiferencia()
    Dim vBookMark As Variant
    Dim lI As Long
    
    dDiferenciaTotal = 0
    
    For lI = 0 To Me.ssdbgPagos.Rows - 1
        vBookMark = Me.ssdbgPagos.AddItemBookmark(lI)
        dDiferenciaTotal = dDiferenciaTotal + CDbl(Me.ssdbgPagos.Columns("Diferencia").CellValue(vBookMark))
    Next
    
    Me.lblDiferenciaTotal = "Diferencia: " & Format(dDiferenciaTotal, "$#,0.00")
    
    
End Sub
