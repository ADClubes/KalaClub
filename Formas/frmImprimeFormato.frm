VERSION 5.00
Begin VB.Form frmImprimeFormato 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de formato"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   2813
      TabIndex        =   3
      Top             =   1178
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprime 
      Caption         =   "Imprimir"
      Height          =   735
      Left            =   653
      TabIndex        =   2
      Top             =   1178
      Width           =   1215
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Seleccionar Impresora"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmImprimeFormato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lIdFormato As Long
Dim crxApplication As New CRAXDRT.Application
Dim CrxReport As CRAXDRT.Report
Dim CrxFormulaFields As CRAXDRT.FormulaFieldDefinitions
Dim CrxFormulaField As CRAXDRT.FormulaFieldDefinition
Dim CrxParameterFields As CRAXDRT.ParameterFieldDefinitions

Private Sub cmdImprime_Click()
    Dim adorcs As ADODB.Recordset
    Dim lPos As Long
    Dim lIndex As Long
    
    Dim sCurrentPrinter As String
    
    sNombreReporte = ""
    
    #If SqlServer_ Then
        strSQL = "SELECT CTF.ArchivoRpt, FD.IdItem, CTFC.NoItemInReport, CASE WHEN FD.IdItem = 1 THEN CONVERT(varchar,CONVERT(datetime,FD.Valor),103) ELSE FD.Valor END AS Valor"
        strSQL = strSQL & " FROM ((FORMATOS AS F INNER JOIN FORMATOS_DETALLE AS FD ON F.IdFormato = FD.IdFormato) INNER JOIN CT_Formatos AS CTF ON F.IdTipoFormato = CTF.IdTipoFormato) INNER JOIN CT_Formatos_Campos CTFC ON (FD.IdItem = CTFC.IdItem) AND (F.IdTipoFormato = CTFC.IdTipoFormato)"
        strSQL = strSQL & " WHERE (((F.IdFormato)=" & lIdFormato & "))"
        strSQL = strSQL & " ORDER BY FD.IdItem"
    #Else
        strSQL = "SELECT CTF.ArchivoRpt, FD.IdItem, CTFC.NoItemInReport, FD.Valor"
        strSQL = strSQL & " FROM ((FORMATOS AS F INNER JOIN FORMATOS_DETALLE AS FD ON F.IdFormato = FD.IdFormato) INNER JOIN CT_Formatos AS CTF ON F.IdTipoFormato = CTF.IdTipoFormato) INNER JOIN CT_Formatos_Campos CTFC ON (FD.IdItem = CTFC.IdItem) AND (F.IdTipoFormato = CTFC.IdTipoFormato)"
        strSQL = strSQL & " WHERE (((F.IdFormato)=" & lIdFormato & "))"
        strSQL = strSQL & " ORDER BY FD.IdItem"
    #End If
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenStatic, adLockReadOnly
    
    
    sCurrentPrinter = SelectPrinter(Me.cmbPrinter.Text)
    
    
    
    sNombreReporte = sDB_ReportSource & "\" & Trim(adorcs!ArchivoRpt)
    
    
    
    Set CrxReport = crxApplication.OpenReport(sNombreReporte, 1)
    CrxReport.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    
    CrxReport.EnableParameterPrompting = False
    
    'Set CrxFormulaFields = CrxReport.FormulaFields
    Set CrxParameterFields = CrxReport.ParameterFields
    
    lPos = 1
    
    Do Until adorcs.EOF
        
        If lPos > CrxParameterFields.Count Then
            Exit Do
            End If
        
        lIndex = adorcs!NoItemInReport
        
        If lIndex <= CrxParameterFields.Count Then
            
        
        
        
            Select Case CrxParameterFields(lIndex).ValueType
                
                Case crNumberField
                    CrxParameterFields(lIndex).AddCurrentValue (Val(Trim(adorcs!Valor)))
                Case crStringField
                    CrxParameterFields(lIndex).AddCurrentValue (Trim(adorcs!Valor))
                Case crDateField
                    CrxParameterFields(lIndex).AddCurrentValue (CDate(Trim(adorcs!Valor)))
            End Select
        
        End If
        
        lPos = lPos + 1
        adorcs.MoveNext
    Loop
    
    
    
    
    CrxReport.PrintOut False
    
    
    Set CrxReport = Nothing
    
    adorcs.Close
    Set adorcs = Nothing
    
    
    SelectPrinter sCurrentPrinter
    
    Unload Me
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    LlenaComboImpresoras Me.cmbPrinter, True
    
    CentraForma MDIPrincipal, Me
    
End Sub
