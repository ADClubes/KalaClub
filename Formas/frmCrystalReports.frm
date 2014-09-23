VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportesCrystal 
   Caption         =   "Pre-Visualización del Reporte Seleccionado"
   ClientHeight    =   6540
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmCrystalReports.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   11400
   Begin VB.Frame frmExporta 
      BorderStyle     =   0  'None
      Caption         =   "Exporta"
      Height          =   2055
      Left            =   3120
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   6255
      Begin VB.OptionButton optExcel 
         Caption         =   "Excel"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton optCsv 
         Caption         =   "CSV"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtNomArchivo 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   5655
      End
      Begin VB.CommandButton cmdNomArchivo 
         Caption         =   "Examinar"
         Height          =   375
         Left            =   4800
         TabIndex        =   2
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Formato"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   3735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6225
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   11160
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmReportesCrystal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' * * * * * * REPORTES GENERADOS EN CRYSTAL REPORTS * * * * * * * * *
'
' Objetivo: GENERA LOS REPORTES PARA SU IMPRESIÓN                   *
'    Autor:
'    Fecha: DICIEMBRE de 2002                                       *
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Dim AdoRcs1 As ADODB.Recordset
Dim crxApplication As New CRAXDRT.Application
Dim CrxReport As CRAXDRT.Report
Dim CrxFormulaFields As CRAXDRT.FormulaFieldDefinitions
Dim CrxFormulaField As CRAXDRT.FormulaFieldDefinition
Dim sqlQuery As String

Private Sub cmdCancelar_Click()
    Me.frmExporta.Visible = False
End Sub

Private Sub cmdExportar_Click()
    
    If Me.optCsv.Value Then
        CrxReport.ExportOptions.DestinationType = crEDTDiskFile
        CrxReport.ExportOptions.FormatType = crEFTCommaSeparatedValues
        CrxReport.ExportOptions.DiskFileName = Trim(Me.txtNomArchivo.Text)
    Else
        CrxReport.ExportOptions.DestinationType = crEDTDiskFile
        CrxReport.ExportOptions.FormatType = crEFTExcel80
        CrxReport.ExportOptions.DiskFileName = Trim(Me.txtNomArchivo.Text)
    End If
    CrxReport.Export (False)
    
    Me.frmExporta.Visible = False
End Sub

Private Sub cmdNomArchivo_Click()

    Me.CommonDialog1.DialogTitle = "Nombre de archivo de salida"
    Me.CommonDialog1.CancelError = True
    On Error GoTo ErrCommonDialog
    If Me.optExcel Then
        Me.CommonDialog1.Filter = "Archivos de Microsoft Excel (*.xls)|*.xls|"
    Else
        Me.CommonDialog1.Filter = "Archivos de Texto (*.csv)|*.csv|"
    End If
    
    Me.CommonDialog1.FilterIndex = 1
    
    
    Me.CommonDialog1.FileName = Trim(sReportes.Nombre)
    
    Me.CommonDialog1.ShowOpen
    
    Me.txtNomArchivo.Text = Trim(Me.CommonDialog1.FileName)
    If Len(Me.txtNomArchivo.Text) > 0 Then
        Me.cmdExportar.Enabled = True
    End If
    
    Exit Sub
ErrCommonDialog:
    Exit Sub
End Sub

Private Sub CRViewer1_ExportButtonClicked(UseDefault As Boolean)
    
    Me.frmExporta.Visible = True
    Me.optExcel.Value = True
    
    Me.cmdExportar.Enabled = False

    UseDefault = False
End Sub

Private Sub CRViewer1_OnReportSourceError(ByVal errorMsg As String, ByVal errorCode As Long, UseDefault As Boolean)
    Beep
    MsgBox errorMsg & " " & errorCode
End Sub

Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)

    Exit Sub
    UseDefault = False
    On Error GoTo Cancel:

    Me.CommonDialog1.CancelError = True
    Me.CommonDialog1.ShowPrinter
    CrxReport.PrintOut False
Cancel:
End Sub

Private Sub Form_Activate()
    Dim Reporte As String, intIdRep As Integer
    
    
    Conn.Errors.Clear
    Err.Clear
    On Error GoTo CATCH_ERROR
    
    If Me.Tag = "LOADED" Then Exit Sub
    
    
    Me.Tag = "LOADED"
        
    Screen.MousePointer = vbHourglass
        
        
    Reporte = sDB_ReportSource & "\" & UCase(Trim(sReportes.Nombre)) & ".rpt"
            
    Call Obtiene_Query
    
    Set AdoRcs1 = New ADODB.Recordset
    Conn.CommandTimeout = 600
    AdoRcs1.ActiveConnection = Conn
    AdoRcs1.CursorType = adOpenStatic
    AdoRcs1.LockType = adLockReadOnly
    AdoRcs1.CursorLocation = adUseServer
    
              
    If strSQL <> "" Then
        On Error Resume Next
        AdoRcs1.Open strSQL
        If Err.Number <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Error al ejecutar la consulta SQL " & vbLf & Err.Description, vbCritical, "Reportes"
            Unload Me
            Exit Sub
        End If
    End If
    
    On Error GoTo CATCH_ERROR
    
    
    
    If AdoRcs1.EOF Then
        MsgBox "¡ No Se Encontró Información Para Este Reporte !", vbExclamation, "Reporte Vacío"
        Screen.MousePointer = vbDefault
        Unload Me
        Exit Sub
    End If
    
    Set CrxReport = crxApplication.OpenReport(Reporte, 1)
    CrxReport.Database.Tables(1).SetDataSource AdoRcs1, 3
    CrxReport.DiscardSavedData
    CrxReport.ReadRecords
    
    
    PasaParametrosReportes frmSelecReportes.cboReportes.ItemData(frmSelecReportes.cboReportes.ListIndex), CrxReport
    
    CRViewer1.ReportSource = CrxReport
    CRViewer1.ViewReport
    CRViewer1.Zoom 100
    
    Me.Tag = "LOADED"

    Screen.MousePointer = vbDefault
    
    Exit Sub
CATCH_ERROR:
    Screen.MousePointer = vbDefault
    MsgBox "Ocurrió un error!" & vbLf & Err.Number & vbLf & Err.Description, vbCritical, "Reportes"
End Sub

Private Sub Form_Initialize()
    Me.Caption = sReportes.Titulo
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 7000
    Me.Width = 13000
    Me.CRViewer1.EnableGroupTree = True
    Me.CRViewer1.DisplayGroupTree = False
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If AdoRcs1.State Then
        AdoRcs1.Close
        Set AdoRcs1 = Nothing
    End If
    
    Set CrxReport = Nothing
End Sub

Sub Obtiene_Query()
    Dim AdoRcsReportes As ADODB.Recordset
    Dim strParametro, strProc, strvar As String
    Dim lngPos As Long
    
    On Error GoTo Err_Obtiene_Query
        
    sqlQuery = "SELECT * FROM reportes WHERE idreporte = " & _
                    frmSelecReportes.cboReportes.ItemData(frmSelecReportes.cboReportes.ListIndex)
    
    Set AdoRcsReportes = New ADODB.Recordset
    AdoRcsReportes.ActiveConnection = Conn
    AdoRcsReportes.CursorLocation = adUseClient
    AdoRcsReportes.CursorType = adOpenDynamic
    AdoRcsReportes.LockType = adLockReadOnly
    AdoRcsReportes.Open sqlQuery
    
    strParametro = sReportes.xImpPara(0)
       
    'Arma el query para el reporte
    strSQL = ""
    If Not AdoRcsReportes.EOF Then
        #If SqlServer_ Then
            strProc = Trim(AdoRcsReportes!sqlQuerySQL)
        #Else
            strProc = Trim(AdoRcsReportes!sqlQuery)
        #End If
        strSQL = ""
        Do While 1
            lngPos = InStr(strProc, "<@")
            If lngPos = 0 Then
                strSQL = strSQL & strProc
                Exit Do
            Else
                strSQL = strSQL & Left$(strProc, lngPos - 1)
            End If
            strProc = Mid$(strProc, lngPos)
            lngPos = InStr(strProc, ">")
            strvar = Left$(strProc, lngPos)
            strProc = Mid$(strProc, lngPos + 1)
            Select Case strvar
                Case "<@Fecha>"
                    #If SqlServer_ Then
                        strSQL = strSQL & " '" & Format(frmSelecReportes.dtpFecha(0).Value, "yyyymmdd") & "' "
                    #Else
                        strSQL = strSQL & " #" & Format(frmSelecReportes.dtpFecha(0).Value, "mm/dd/yyyy") & "# "
                    #End If
                Case "<@FechaInicial>"
                    #If SqlServer_ Then
                        strSQL = strSQL & " '" & Format(frmSelecReportes.dtpFecha(1).Value, "yyyymmdd") & "' "
                    #Else
                        strSQL = strSQL & " #" & Format(frmSelecReportes.dtpFecha(1).Value, "mm/dd/yyyy") & "# "
                    #End If
                Case "<@FechaFinal>"
                    #If SqlServer_ Then
                        strSQL = strSQL & " '" & Format(frmSelecReportes.dtpFecha(2).Value, "yyyymmdd") & "' "
                    #Else
                        strSQL = strSQL & " #" & Format(frmSelecReportes.dtpFecha(2).Value, "mm/dd/yyyy") & "# "
                    #End If
                Case "<@Num>"
                    strSQL = strSQL & Val(frmSelecReportes.txtNumero(0).Text)
                Case "<@NumInicial>"
                    strSQL = strSQL & Val(frmSelecReportes.txtNumero(1).Text)
                Case "<@NumFinal>"
                    strSQL = strSQL & Val(frmSelecReportes.txtNumero(2).Text)
                Case "<@IdAccionista>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(0).ItemData(frmSelecReportes.cboDatos(0).ListIndex)
                Case "<@Anio>"
                    strSQL = strSQL & Val(frmSelecReportes.cboDatos(1).Text)
                Case "<@IdMembresia>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(2).ItemData(frmSelecReportes.cboDatos(2).ListIndex)
                Case "<@Mes>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(3).ListIndex + 1
                Case "<@Serie>"
                    strSQL = strSQL & "'" & Trim(frmSelecReportes.cboDatos(4).Text) & "'"
                Case "<@Sexo>"
                    strSQL = strSQL & "'" & Trim(frmSelecReportes.cboDatos(5).Text) & "'"
                Case "<@TipoPago>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(6).ItemData(frmSelecReportes.cboDatos(6).ListIndex)
                Case "<@TipoRentable>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(7).ItemData(frmSelecReportes.cboDatos(7).ListIndex)
                Case "<@TipoUsuario>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(8).ItemData(frmSelecReportes.cboDatos(8).ListIndex)
                Case "<@Socio>"
                    strSQL = strSQL & frmSelecReportes.cboDatos(9).ItemData(frmSelecReportes.cboDatos(9).ListIndex)
                Case "<@Turno>"
                    If Val(frmSelecReportes.txtNumero(3).Text) > 0 Then
                        strSQL = strSQL & Trim(frmSelecReportes.txtNumero(3).Text)
                    Else
                        #If SqlServer_ Then
                            strSQL = Mid(strSQL, 1, Len(strSQL) - 1)
                            strSQL = strSQL & ">0"
                        #Else
                            strSQL = strSQL & "True"
                        #End If
                    End If
                Case "<@Caja>"
                    If Val(frmSelecReportes.txtNumero(4).Text) > 0 Then
                        strSQL = strSQL & Trim(frmSelecReportes.txtNumero(4).Text)
                    Else
                        #If SqlServer_ Then
                            strSQL = Mid(strSQL, 1, Len(strSQL) - 1)
                            strSQL = strSQL & ">0"
                        #Else
                            strSQL = strSQL & "True"
                        #End If
                    End If
            End Select
        Loop
    End If
    
    Exit Sub
Err_Obtiene_Query:
        GeneraMensajeError Err.Number
End Sub

Sub PasaParametrosReportes(lIdReporte As Long, Reporte As CRAXDRT.Report)

    Dim AdoRcsPara As ADODB.Recordset

    Set CrxFormulaFields = Reporte.FormulaFields
    
    
    strSQL = "SELECT * FROM REPORTES_PARAMETROS"
    strSQL = strSQL & " WHERE IdReporte=" & lIdReporte
    
    Set AdoRcsPara = New ADODB.Recordset
    AdoRcsPara.ActiveConnection = Conn
    AdoRcsPara.CursorLocation = adUseClient
    AdoRcsPara.CursorType = adOpenForwardOnly
    AdoRcsPara.LockType = adLockReadOnly
    AdoRcsPara.Open strSQL
    
    If Not AdoRcsPara.EOF Then
        Do While Not AdoRcsPara.EOF
    
        For Each CrxFormulaField In CrxFormulaFields
            'Encuentra la fórmula requerida con el nombre especificado
            If CrxFormulaField.Name = Trim(AdoRcsPara!FormulaField) Then
                'Asigna el texto que aparecerá en el campo fórmula
                Select Case Trim(AdoRcsPara!Parametro)
                    Case "<@NombreReporte>"
                        CrxFormulaField.Text = "'" & frmSelecReportes.cboReportes.Text & "'"
                    Case "<@FechaInicial>"
                        CrxFormulaField.Text = "'" & Format(frmSelecReportes.dtpFecha(1).Value, "dd/mmm/yyyy") & "'"
                    Case "<@FechaFinal>"
                        CrxFormulaField.Text = "'" & Format(frmSelecReportes.dtpFecha(2).Value, "dd/mmm/yyyy") & "'"
                    Case "<@Turno>"
                        CrxFormulaField.Text = "'" & Trim(frmSelecReportes.txtNumero(3)) & "'"
                     Case "<@Caja>"
                        CrxFormulaField.Text = "'" & Trim(frmSelecReportes.txtNumero(4)) & "'"
                End Select
                Exit For
            End If
        Next
        AdoRcsPara.MoveNext
        Loop
    End If
    
    AdoRcsPara.Close
    Set AdoRcsPara = Nothing
    
End Sub

