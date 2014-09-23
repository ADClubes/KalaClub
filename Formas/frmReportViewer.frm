VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReportViewer 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   1080
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmExporta 
      BorderStyle     =   0  'None
      Caption         =   "Exporta"
      Height          =   2055
      Left            =   2640
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdNomArchivo 
         Caption         =   "Examinar"
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtNomArchivo 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   5655
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optCsv 
         Caption         =   "CSV"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optExcel 
         Caption         =   "Excel"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Formato"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   120
         Width           =   735
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public sNombreReporte As String
Public sQuery As String
Public strValor1 As String
Public strValor2 As String
Public strValor3 As String
Public strValor4 As String
Public strValor5 As String

Dim adorcs As ADODB.Recordset
Dim crxApplication As New CRAXDRT.Application
Dim CrxReport As CRAXDRT.Report
Dim CrxFormulaFields As CRAXDRT.FormulaFieldDefinitions
Dim CrxFormulaField As CRAXDRT.FormulaFieldDefinition

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

Private Sub Form_Activate()
    Dim intIdRep As Integer
    
    Conn.Errors.Clear
    Err.Clear
    On Error GoTo CATCH_ERROR
    
    If Me.Tag = "LOADED" Then Exit Sub
    
    Me.Caption = sNombreReporte
    
    Me.Tag = "LOADED"
    
    Screen.MousePointer = vbHourglass
    
    Set adorcs = New ADODB.Recordset
    adorcs.ActiveConnection = Conn
    adorcs.CursorType = adOpenStatic
    adorcs.LockType = adLockReadOnly
    adorcs.CursorLocation = adUseServer
              
    If sQuery <> vbNullString Then
        On Error Resume Next
        adorcs.Open sQuery
        If Err.Number <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Error al ejecutar la consulta SQL " & vbLf & Err.Description, vbCritical, "Reportes"
            Unload Me
            Exit Sub
        End If
    End If
    
    On Error GoTo CATCH_ERROR
    
    If sQuery <> vbNullString Then
        If adorcs.EOF Then
            MsgBox "¡ No Se Encontró Información Para Este Reporte !", vbExclamation, "Reporte Vacío"
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
        End If
    End If
    
    Set CrxReport = crxApplication.OpenReport(sNombreReporte, 1)
    
    If sQuery <> vbNullString Then
        CrxReport.Database.Tables(1).SetDataSource adorcs, 3
        CrxReport.DiscardSavedData
        CrxReport.ReadRecords
    End If
    
    If strValor1 <> vbNullString Then
        AsignaFormulas CrxReport
    End If
    
    CRViewer1.ReportSource = CrxReport
    CRViewer1.ViewReport
    CRViewer1.Zoom 100

    Screen.MousePointer = vbDefault
    
    Exit Sub
CATCH_ERROR:
    Screen.MousePointer = vbDefault
    MsgBox "Ocurrió un error!" & vbLf & Err.Number & vbLf & Err.Description, vbCritical, "Reportes"
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 7000
    Me.Width = 13000
    Me.CRViewer1.EnableGroupTree = True
    Me.CRViewer1.DisplayGroupTree = False
    CentraForma MDIPrincipal, Me
End Sub
Private Sub mnuProcEjecutarQry_Click()
    Dim frmExec As frmSQLExec
    
    Set frmExec = New frmSQLExec
    
    frmExec.Show vbModal
    
    
End Sub
Private Sub CRViewer1_ExportButtonClicked(UseDefault As Boolean)
    
    Me.frmExporta.Visible = True
    Me.optExcel.Value = True
    
    Me.cmdExportar.Enabled = False

    UseDefault = False
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If adorcs.State Then
        adorcs.Close
        Set adorcs = Nothing
    End If
End Sub

Private Sub AsignaFormulas(Reporte As CRAXDRT.Report)
    Dim lI As Long
    
    Set CrxFormulaFields = Reporte.FormulaFields
    
    For lI = 1 To CrxFormulaFields.Count
        Select Case lI
            Case 1
                CrxFormulaFields(lI).Text = "'" & strValor1 & "'"
             Case 2
                CrxFormulaFields(lI).Text = "'" & strValor2 & "'"
            Case 3
                CrxFormulaFields(lI).Text = "'" & strValor3 & "'"
            Case 4
                CrxFormulaFields(lI).Text = "'" & strValor4 & "'"
            Case 5
                CrxFormulaFields(lI).Text = "'" & strValor5 & "'"
        End Select
    Next
End Sub
