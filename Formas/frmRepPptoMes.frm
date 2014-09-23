VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmRepPptoMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Comparativo presupuesto"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdbgReporte 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   480
      Width           =   3975
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
      Columns.Count   =   3
      Columns(0).Width=   7832
      Columns(0).Caption=   "Reporte"
      Columns(0).Name =   "Reporte"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "IdReporte"
      Columns(1).Name =   "IdReporte"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "ReporteRPT"
      Columns(2).Name =   "ReporteRPT"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   7011
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1035
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14993
            MinWidth        =   14993
            Key             =   "Concepto"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Generar"
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpFechaRep 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   3604481
      CurrentDate     =   40400
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Reporte"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmRepPptoMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReporte_Click()


    Dim frmrpt As frmReportViewer

    Dim adoRcsCfg As ADODB.Recordset
    Dim adoCmdCfg As ADODB.Command
    
    Dim adoParamIdent As ADODB.Parameter
    Dim adoParamConcepPpto As ADODB.Parameter
    Dim adoParamConcep As ADODB.Parameter
    Dim adoParamFechaInicial As ADODB.Parameter
    Dim adoParamFechaFinal As ADODB.Parameter
    
    Dim sIdent As String
    
    Dim sUnidad As String
    
    If Me.ssdbgReporte.Text = vbNullString Then
        MsgBox "Seleccione un reporte", vbExclamation, "Corrija"
        Me.ssdbgReporte.SetFocus
        Exit Sub
    End If
    
    
    
    Me.cmdReporte.Enabled = False
    Me.cmdSalir.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
    sIdent = Format(Now, "ddmmHhNnSs")
    
    sUnidad = ObtieneParametro("NOMBRE DEL CLUB")
    
    Set adoRcsCfg = New ADODB.Recordset
    
    strSQL = "SELECT CFG_Reporte_Ppto.NombreSP, CFG_Reporte_Ppto.ConceptoPpto, CFG_Reporte_Ppto.ConceptoIngreso, CFG_Reporte_Ppto.IdGrupoPpto"
    strSQL = strSQL & " From CFG_Reporte_Ppto"
    strSQL = strSQL & " Where CFG_Reporte_Ppto.IdReporte=" & Me.ssdbgReporte.Columns("idReporte").Value
    strSQL = strSQL & " ORDER BY CFG_Reporte_Ppto.IdSP"
    
    
    adoRcsCfg.CursorLocation = adUseServer
Debug.Print strSQL
    adoRcsCfg.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    
    
    
    
    
    Do While Not adoRcsCfg.EOF
        
        Me.StatusBar1.Panels("Concepto").Text = adoRcsCfg!ConceptoPpto
        
        DoEvents

        #If SqlServer_ Then
            Set adoCmdCfg = New ADODB.Command
            adoCmdCfg.ActiveConnection = Conn
            adoCmdCfg.CommandType = adCmdStoredProc
            adoCmdCfg.CommandText = adoRcsCfg!NombreSP
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter(Name:="@Ident", Type:=adVarChar, Direction:=adParamInput, Size:=255, Value:=sIdent)
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter(Name:="@IdGrupoPpto", Type:=adVarChar, Direction:=adParamInput, Size:=255, Value:=adoRcsCfg!IdGrupoPpto)
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter(Name:="@Concep", Type:=adVarChar, Direction:=adParamInput, Size:=255, Value:=adoRcsCfg!ConceptoIngreso)
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter(Name:="@FechaIni", Type:=adVarChar, Direction:=adParamInput, Size:=8, Value:=Format(Me.dtpFechaRep.Value, "yyyymm01"))
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter(Name:="@FechaFin", Type:=adVarChar, Direction:=adParamInput, Size:=8, Value:=Format(Me.dtpFechaRep.Value, "yyyymmdd"))
            
            Debug.Print adoRcsCfg!NombreSP & _
            vbNewLine & "@Ident=" & sIdent & _
            vbNewLine & "@IdGrupoPpto=" & adoRcsCfg!IdGrupoPpto & _
            vbNewLine & "@Concep=" & adoRcsCfg!ConceptoIngreso & _
            vbNewLine & "@FechaIni=" & Format(Me.dtpFechaRep.Value, "yyyymm01") & _
            vbNewLine & "@FechaFin=" & Format(Me.dtpFechaRep.Value, "yyyymmdd")
        #Else
            Set adoCmdCfg = New ADODB.Command
            adoCmdCfg.ActiveConnection = Conn
            adoCmdCfg.CommandType = adCmdStoredProc
            adoCmdCfg.CommandText = adoRcsCfg!NombreSP
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter("Ident", adVarChar, adParamInput, 255, sIdent)
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter("IdGrupoPpto", adInteger, adParamInput, 255, adoRcsCfg!IdGrupoPpto)
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter("Concep", adVarChar, adParamInput, 255, adoRcsCfg!ConceptoIngreso)
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter("FechaIni", adDate, adParamInput, , DateSerial(Year(Me.dtpFechaRep.Value), Month(Me.dtpFechaRep.Value), 1))
            adoCmdCfg.Parameters.Append adoCmdCfg.CreateParameter("FechaFin", adDate, adParamInput, , Me.dtpFechaRep.Value)
        #End If
        
        adoCmdCfg.Execute
        
        Set adoCmdCfg = Nothing
        
        
        
                
        
        adoRcsCfg.MoveNext
    Loop
    
    
    adoRcsCfg.Close
    Set adoRcsCfg = Nothing
    
    
       
    
    
    Me.StatusBar1.Panels("Concepto").Text = "Terminado"
    
    
    strSQL = "SELECT GRUPO, SUBGRUPO, R.SumaDeImporteP AS ImporteP, R.SumaDeCantidadP AS CantidadP, R.SumaDeImporteR AS ImporteR, R.SumaDeCantidadR As CantidadR, R.SumaDeImporteRAA AS ImporteRAA, R.SumaDeCantidadRAA As CantidadRAA, '" & Me.dtpFechaRep.Value & "' AS Fecha" & ", '" & sUnidad & "' AS Unidad"
    strSQL = strSQL & " FROM CFG_GRU_PPTO_OP LEFT JOIN (SELECT TMP_REP_PPTO_OP.IdGrupoPpto, Sum(TMP_REP_PPTO_OP.ImporteP) AS SumaDeImporteP, Sum(TMP_REP_PPTO_OP.CantidadP) AS SumaDeCantidadP, Sum(TMP_REP_PPTO_OP.ImporteR) AS SumaDeImporteR, Sum(TMP_REP_PPTO_OP.CantidadR) AS SumaDeCantidadR, Sum(TMP_REP_PPTO_OP.ImporteRAA) AS SumaDeImporteRAA, Sum(TMP_REP_PPTO_OP.CantidadRAA) AS SumaDeCantidadRAA"
    strSQL = strSQL & " From TMP_REP_PPTO_OP"
    strSQL = strSQL & " WHERE TMP_REP_PPTO_OP.IdSesion='" & sIdent & "'"
    strSQL = strSQL & " GROUP BY TMP_REP_PPTO_OP.IdGrupoPpto) AS R ON CFG_GRU_PPTO_OP.IdGrupoPpto = R.IdGrupoPpto ORDER BY CFG_GRU_PPTO_OP.IdGrupoPpto"
    
    
    'strSQL = "SELECT PresupuestoIngresos.Concepto, PresupuestoIngresos.ImportePpto, TMP_REPORTE_PPTO.Importe," & "'" & Me.dtpFechaRep.Value & "' AS Fecha"
    'strSQL = strSQL & " FROM PresupuestoIngresos LEFT JOIN TMP_REPORTE_PPTO ON PresupuestoIngresos.Concepto = TMP_REPORTE_PPTO.Concepto"
    'strSQL = strSQL & " WHERE ("
    'strSQL = strSQL & "((PresupuestoIngresos.Mes)=" & Month(Me.dtpFechaRep.Value) & ")"
    'strSQL = strSQL & " AND ((PresupuestoIngresos.Anio)=" & Year(Me.dtpFechaRep.Value) & ")"
    'strSQL = strSQL & ")"
    
    Set frmrpt = New frmReportViewer
    
    frmrpt.sQuery = strSQL
    frmrpt.sNombreReporte = sDB_ReportSource & "\" & Me.ssdbgReporte.Columns("ReporteRPT").Value
    
    
    
    Screen.MousePointer = vbDefault
    
    frmrpt.Show vbModal
    
    Set frmrpt = Nothing
    
    'Borra los registros en la tabla temporal
    #If SqlServer_ Then
        strSQL = "DELETE FROM TMP_REP_PPTO_OP WHERE IdSesion='" & sIdent & "'"
    #Else
        strSQL = "DELETE * FROM TMP_REP_PPTO_OP WHERE IdSesion='" & sIdent & "'"
    #End If
    
    Set adoCmdCfg = New ADODB.Command
    adoCmdCfg.ActiveConnection = Conn
    adoCmdCfg.CommandText = strSQL
    adoCmdCfg.CommandType = adCmdText
    adoCmdCfg.Execute
    Set adoCmdCfg = Nothing
    
    Me.cmdReporte.Enabled = True
    Me.cmdSalir.Enabled = True
    
    
    
    
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.dtpFechaRep.Value = Date
    
    strSQL = "SELECT CT_REP_PPTO.ReporteNombre, CT_REP_PPTO.idReporte, CT_REP_PPTO.ReporteRPT"
    strSQL = strSQL & " From CT_REP_PPTO"
    strSQL = strSQL & " WHERE (((CT_REP_PPTO.Status)='A'))"
    strSQL = strSQL & " ORDER BY CT_REP_PPTO.idReporte"

    
    CentraForma MDIPrincipal, Me
    
    LlenaSsCombo Me.ssdbgReporte, Conn, strSQL, 3
    
End Sub
