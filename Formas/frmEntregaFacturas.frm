VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmEntregaFacturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entrega de Facturas"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   Icon            =   "frmEntregaFacturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImpTicket 
      Caption         =   "Imprime Ticket"
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCopiaCFD 
      Caption         =   "Copia CFD"
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "Imprime CFD"
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprime 
      Caption         =   "Imprime acuse"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdDesmarcaTodos 
      Caption         =   "Desmarca Todas"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdMarcaTodos 
      Caption         =   "Marca Todas"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgFacturas 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   4
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   2037
      Columns(0).Caption=   "Folio"
      Columns(0).Name =   "Folio"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   3200
      Columns(1).Caption=   "Numero"
      Columns(1).Name =   "Numero"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2090
      Columns(2).Caption=   "Fecha"
      Columns(2).Name =   "Fecha"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(3).Width=   1138
      Columns(3).Name =   "Marca"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   11
      Columns(3).FieldLen=   256
      Columns(3).Style=   2
      _ExtentX        =   10398
      _ExtentY        =   7223
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
Attribute VB_Name = "frmEntregaFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lIdTitular As Long



Private Sub cmdCopiaCFD_Click()

    

     If Me.ssdbgFacturas.Rows = 0 Then
        Exit Sub
    End If
    
    If Me.ssdbgFacturas.Columns("Folio").Value = "" Then
        MsgBox "Esta factura no tiene CFD", vbInformation, "Verifique"
        Exit Sub
    End If
    
    
    Dim sNombreArcConRuta As String
    Dim sNombrePDF As String
    Dim sNombreXML As String
    Dim sNombreArc As String
    
    
    
    sNombreArcConRuta = ObtieneParametro("RUTA_CFD") & "\" & NombreArchivoCFD(Me.ssdbgFacturas.Columns("Numero").Value, "F", 0)
    
    
    sNombrePDF = sNombreArcConRuta & ".pdf"
    sNombreXML = sNombreArcConRuta & ".xml"
    
    sNombreArc = NombreArchivoCFD(Me.ssdbgFacturas.Columns("Numero").Value, "F", 1)
    
    
    Dim fsObjectCopy As FileSystemObject

    
    Me.CommonDialog1.DialogTitle = "Nombre de archivo PDF"
    Me.CommonDialog1.CancelError = True
    On Error GoTo ErrCommonDialog
    
    
    
    Me.CommonDialog1.Filter = "Archivos PDF (*.pdf)|*.pdf|"
    
    Me.CommonDialog1.FilterIndex = 1
    
    
    Me.CommonDialog1.FileName = Trim(sNombreArc)
    
    'Me.CommonDialog1.InitDir
    
    Me.CommonDialog1.ShowSave
    
    
    
    Set fsObjectCopy = New FileSystemObject
    
    
    
    fsObjectCopy.CopyFile sNombrePDF, Trim(Me.CommonDialog1.FileName), True
    
    Set fsObjectCopy = Nothing
    
    
    Me.CommonDialog1.DialogTitle = "Nombre de archivo XML"
    Me.CommonDialog1.CancelError = True
    On Error GoTo ErrCommonDialog
    
    
    
    Me.CommonDialog1.Filter = "Archivos XML (*.xml)|*.xml|"
    
    Me.CommonDialog1.FilterIndex = 1
    
    
    Me.CommonDialog1.FileName = Trim(sNombreArc)
    
    'Me.CommonDialog1.InitDir
    
    Me.CommonDialog1.ShowSave
    
    
    
    Set fsObjectCopy = New FileSystemObject
    
    
    
    fsObjectCopy.CopyFile sNombreXML, Trim(Me.CommonDialog1.FileName), True
    
    Set fsObjectCopy = Nothing
    
    
    
    
    Exit Sub
ErrCommonDialog:
    Exit Sub
    
    
    
    
    
End Sub

Private Sub cmdDesmarcaTodos_Click()
    MarcaRen False
End Sub

Private Sub cmdImp_Click()
    
    If Me.ssdbgFacturas.Rows = 0 Then
        Exit Sub
    End If
    
    If Me.ssdbgFacturas.Columns("Folio").Value = "" Then
        MsgBox "Esta factura no tiene CFD", vbInformation, "Verifique"
        Exit Sub
    End If
    
    
    MuestraCFD Me.ssdbgFacturas.Columns("Numero").Value
    
End Sub

Private Sub cmdImprime_Click()

    Dim sStrSql As String
    
    Dim sCadenaFac As String
    
    Me.ssdbgFacturas.Update
    
    sCadenaFac = CadenaFacturas()
    
    If sCadenaFac = vbNullString Then
        MsgBox "No hay documentos marcados!", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    
    ActualizaFechaEntrega sCadenaFac
    
    #If SqlServer_ Then
        sStrSql = "SELECT FACTURAS.FechaFactura, FACTURAS.Serie + CONVERT(nvarchar, FACTURAS.Folio) As Folio"
        sStrSql = sStrSql & " FROM FACTURAS"
        sStrSql = sStrSql & " WHERE"
        sStrSql = sStrSql & " FACTURAS.NumeroFactura IN ("
        sStrSql = sStrSql & sCadenaFac & ")"
        sStrSql = sStrSql & " ORDER BY FACTURAS.NumeroFactura"
    #Else
        sStrSql = "SELECT FACTURAS.FechaFactura, FACTURAS.Serie & FACTURAS.Folio As Folio"
        sStrSql = sStrSql & " FROM FACTURAS"
        sStrSql = sStrSql & " WHERE"
        sStrSql = sStrSql & " FACTURAS.NumeroFactura IN ("
        sStrSql = sStrSql & sCadenaFac & ")"
        sStrSql = sStrSql & " ORDER BY FACTURAS.NumeroFactura"
    #End If

    frmReportViewer.sNombreReporte = sDB_ReportSource & "\acusefactura.rpt"
    frmReportViewer.sQuery = sStrSql
    
    frmReportViewer.Show vbModal
    
    Unload Me
    
    
End Sub

'06/Dic/2011 UCM
Private Sub cmdImpTicket_Click()
    Dim sStrSql As String
    Dim sCadenaFac As String
    
    Me.ssdbgFacturas.Update
    
    Dim lI As Long
    Dim vBm As Variant
    
    sCadenaFac = vbNullString
    
    For lI = 0 To Me.ssdbgFacturas.Rows - 1
        vBm = Me.ssdbgFacturas.AddItemBookmark(lI)
        If Me.ssdbgFacturas.Columns(3).CellValue(vBm) = True Then
            sCadenaFac = Me.ssdbgFacturas.Columns(1).CellValue(vBm)
            Exit For
        End If
    Next
    
    If sCadenaFac = vbNullString Then
        MsgBox "No hay documentos marcados!", vbExclamation, "Verifique"
        Exit Sub
    End If
    
'    If CBool(ssdbgFacturas.Columns("Cancelada").Text) Then
'        MsgBox "No se puede Imprimir este documento por que está cancelado!", vbExclamation, "Documentos"
'        Exit Sub
'    End If

    If ssdbgFacturas.Columns("Folio").CellValue(vBm) = "" Then
        MsgBox "No se puede Imprimir este documento por que no tiene asignado Folio!", vbExclamation, "Documentos"
        Exit Sub
    End If
    
    ActualizaFechaEntrega sCadenaFac

    Dim frmImpF As New frmImpFac

    lNumFacIniImp = ssdbgFacturas.Columns("Numero").CellValue(vBm)
    lNumFacFinImp = ssdbgFacturas.Columns("Numero").CellValue(vBm)

    lNumFolioFacIniImp = ssdbgFacturas.Columns("Folio").CellValue(vBm)
    lNumFolioFacFinImp = ssdbgFacturas.Columns("Folio").CellValue(vBm)
    frmImpF.Tag = "F"
    frmImpF.cModo = "F"
    frmImpF.lNumeroInicial = lNumFacIniImp
    frmImpF.lNumeroFinal = lNumFacFinImp
    frmImpF.Show 1
    
    Set frmImpF = Nothing
End Sub

Private Sub cmdMarcaTodos_Click()
    MarcaRen True
End Sub

Private Sub Form_Activate()
    'LlenaGridFacturas
End Sub

Private Sub LlenaGridFacturas()
    Dim adoRcsFac As ADODB.Recordset
    
    #If SqlServer_ Then
        strSQL = "SELECT FACTURAS.SerieCFD + CONVERT(varchar,FACTURAS.FolioCFD) AS Folio, FACTURAS.NumeroFactura, FACTURAS.FechaFactura"
        strSQL = strSQL & " FROM FACTURAS"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " IdTitular=" & lIdTitular
        strSQL = strSQL & " AND Cancelada=0"
        strSQL = strSQL & " AND FechaEntrega Is Null"
        strSQL = strSQL & ")"
        strSQL = strSQL & " ORDER BY"
        strSQL = strSQL & " NumeroFactura"
    #Else
        strSQL = "SELECT FACTURAS.SerieCFD & FACTURAS.FolioCFD AS Folio, FACTURAS.NumeroFactura, FACTURAS.FechaFactura"
        strSQL = strSQL & " FROM FACTURAS"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " IdTitular=" & lIdTitular
        strSQL = strSQL & " AND Cancelada=0"
        strSQL = strSQL & " AND FechaEntrega Is Null"
        strSQL = strSQL & ")"
        strSQL = strSQL & " ORDER BY"
        strSQL = strSQL & " NumeroFactura"
    #End If
    
    
    Set adoRcsFac = New ADODB.Recordset
    adoRcsFac.CursorLocation = adUseServer
    
    adoRcsFac.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    
    Do Until adoRcsFac.EOF
        Me.ssdbgFacturas.AddItem adoRcsFac!Folio & vbTab & adoRcsFac!NumeroFactura & vbTab & adoRcsFac!FechaFactura
        adoRcsFac.MoveNext
    Loop

    
    
    adoRcsFac.Close
    Set adoRcsFac = Nothing
    
End Sub


Private Sub MarcaRen(boValor As Boolean)
    
    Dim lI As Long
    Dim vBm As Variant
    
    For lI = 0 To Me.ssdbgFacturas.Rows - 1
         Me.ssdbgFacturas.Bookmark = Me.ssdbgFacturas.AddItemBookmark(lI)
         Me.ssdbgFacturas.Columns("Marca").Value = boValor
    Next
    
    Me.ssdbgFacturas.Update


End Sub

Private Function CadenaFacturas() As String
    
    Dim lI As Long
    Dim vBm As Variant
    
    CadenaFacturas = vbNullString
    
    For lI = 0 To Me.ssdbgFacturas.Rows - 1
        vBm = Me.ssdbgFacturas.AddItemBookmark(lI)
        If Me.ssdbgFacturas.Columns(3).CellValue(vBm) = True Then
            CadenaFacturas = CadenaFacturas & Me.ssdbgFacturas.Columns(1).CellValue(vBm) & ","
        End If
         
    Next
    
    If CadenaFacturas <> vbNullString Then
        CadenaFacturas = Mid$(CadenaFacturas, 1, Len(CadenaFacturas) - 1)
    End If
    
    
End Function

Private Sub ActualizaFechaEntrega(sCadenaFac As String)

    Dim adoCmdFac As ADODB.Command
    
    #If SqlServer_ Then
        strSQL = "UPDATE FACTURAS SET"
        strSQL = strSQL & " FechaEntrega=" & "'" & Format(Date, "yyyymmdd") & "',"
        strSQL = strSQL & " HoraEntrega=" & "'" & Format(Now, "Hh:Nn") & "'"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " NumeroFactura IN (" & sCadenaFac & ")"
    #Else
        strSQL = "UPDATE FACTURAS SET"
        strSQL = strSQL & " FechaEntrega=" & "#" & Format(Date, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & " HoraEntrega=" & "'" & Format(Now, "Hh:Nn") & "'"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " NumeroFactura IN (" & sCadenaFac & ")"
    #End If
    
    Set adoCmdFac = New ADODB.Command
    adoCmdFac.ActiveConnection = Conn
    adoCmdFac.CommandType = adCmdText
    adoCmdFac.CommandText = strSQL
    adoCmdFac.Execute
    
    Set adoCmdFac = Nothing
    

End Sub

Private Sub Form_Load()
    LlenaGridFacturas
End Sub

