VERSION 5.00
Begin VB.Form frmReporteCFD 
   Caption         =   "Reporte CFD"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   6
      Left            =   3120
      MaxLength       =   15
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   5
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   11
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   4
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Generar"
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   3
      Left            =   3120
      MaxLength       =   15
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   2
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   1
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   0
      Left            =   1320
      MaxLength       =   18
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblCtrl 
      Alignment       =   2  'Center
      Caption         =   "No. Aprobación"
      Height          =   375
      Index           =   6
      Left            =   2160
      TabIndex        =   15
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Serie Direc."
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Serie Caja"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblCtrl 
      Alignment       =   2  'Center
      Caption         =   "No. Aprobación"
      Height          =   375
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Año"
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "Mes"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblCtrl 
      Alignment       =   1  'Right Justify
      Caption         =   "RFC"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmReporteCFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim adorcs As ADODB.Recordset
    
    Dim dFechaIni As Date
    Dim dFechaFin As Date
    
    Dim sNombreArc As String
    
    
    Dim fs As Object
    Dim outputFile As Object
    
    Dim sRenglon As String
    
    Dim sRFCPG As String
    sRFCPG = "XAXX010101000"
    
    
    dFechaIni = DateSerial(Val(Me.txtCtrl(2).Text), Val(Me.txtCtrl(1).Text), 1)
    dFechaFin = DateSerial(Val(Me.txtCtrl(2).Text), Val(Me.txtCtrl(1).Text) + 1, 1) - 1
    
    sNombreArc = "C:\" & "2" & Trim(Me.txtCtrl(0).Text) & Trim(Me.txtCtrl(1).Text) & Trim(Me.txtCtrl(2).Text) & ".txt"
    
    #If SqlServer_ Then
        strSQL = "SELECT FACTURAS.NumeroFactura, FACTURAS.Folio, FACTURAS.Serie, FACTURAS.FechaFactura, FACTURAS.HoraFactura, FACTURAS.RFC, FACTURAS.TipoPersona, FACTURAS.Cancelada, Sum(FACTURAS_DETALLE.Total) AS Total, Sum(FACTURAS_DETALLE.Iva) AS Iva, Sum(FACTURAS_DETALLE.IvaIntereses) AS IvaIntereses, Sum(FACTURAS_DETALLE.IvaDescuento) AS IvaDescuento"
        strSQL = strSQL & " FROM FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura"
        strSQL = strSQL & " WHERE FACTURAS.FolioCFD Is Null"
        strSQL = strSQL & " GROUP BY FACTURAS.NumeroFactura, FACTURAS.Folio, FACTURAS.Serie, FACTURAS.FechaFactura, FACTURAS.HoraFactura, FACTURAS.RFC, FACTURAS.TipoPersona, FACTURAS.Cancelada"
        strSQL = strSQL & " HAVING"
        strSQL = strSQL & " FACTURAS.FechaFactura Between " & "'" & Format(dFechaIni, "yyyymmdd") & "' And '" & Format(dFechaFin, "yyyymmdd") & "'"
    #Else
        strSQL = "SELECT FACTURAS.NumeroFactura, FACTURAS.Folio, FACTURAS.Serie, FACTURAS.FechaFactura, FACTURAS.HoraFactura, FACTURAS.RFC, FACTURAS.TipoPersona, FACTURAS.Cancelada, Sum(FACTURAS_DETALLE.Total) AS Total, Sum(FACTURAS_DETALLE.Iva) AS Iva, Sum(FACTURAS_DETALLE.IvaIntereses) AS IvaIntereses, Sum(FACTURAS_DETALLE.IvaDescuento) AS IvaDescuento"
        strSQL = strSQL & " FROM FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura"
        strSQL = strSQL & " WHERE (((FACTURAS.FolioCFD) Is Null))"
        strSQL = strSQL & " GROUP BY FACTURAS.NumeroFactura, FACTURAS.Folio, FACTURAS.Serie, FACTURAS.FechaFactura, FACTURAS.HoraFactura, FACTURAS.RFC, FACTURAS.TipoPersona, FACTURAS.Cancelada"
        strSQL = strSQL & " HAVING ("
        strSQL = strSQL & "((FACTURAS.FechaFactura) Between " & "#" & Format(dFechaIni, "mm/dd/yyyy") & "# And #" & Format(dFechaFin, "mm/dd/yyyy") & "#))"
    #End If

    Set adorcs = New ADODB.Recordset
    
    
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set outputFile = fs.CreateTextFile(sNombreArc)
    
    
    Do While Not adorcs.EOF
    
        
        sRfc = Trim(adorcs!rfc)
    
        If sRfc = vbNullString Then
            sRfc = sRFCPG
        End If
        
        If Len(sRfc) > 13 Then
            sRfc = sRFCPG
        End If
        
        If Len(sRfc) < 12 Then
            sRfc = sRFCPG
        End If
        
        
    
        sRenglon = vbNullString
        sRenglon = sRenglon & "|"
        sRenglon = sRenglon & sRfc & "|"
        sRenglon = sRenglon & Trim(adorcs!Serie) & "|"      'Serie factura
        sRenglon = sRenglon & Trim(adorcs!Folio) & "|"      'Folio factura
        If Trim(adorcs!Serie) = Trim(Me.txtCtrl(4).Text) Then 'Si es de caja
            sRenglon = sRenglon & Trim(Me.txtCtrl(3)) & "|"     'No aprobacion
        Else
            sRenglon = sRenglon & Trim(Me.txtCtrl(6)) & "|"     'No aprobacion
        End If
        sRenglon = sRenglon & Format(adorcs!FechaFactura, "dd/mm/yyyy") & " 00:00:00" & "|"    'Fecha y hora
        sRenglon = sRenglon & Format(adorcs!Total, "#########0.00") & "|"    'Total
        sRenglon = sRenglon & Format(adorcs!Iva + adorcs!IvaIntereses - adorcs!IvaDescuento, "#########0.00") & "|" 'IVA
        sRenglon = sRenglon & IIf(adorcs!Cancelada, "0", "1") & "|"  ' Activa o cancelada
        sRenglon = sRenglon & "I" & "|" 'Efecto del comprobante
        sRenglon = sRenglon & "" & "|" 'Pedimento
        sRenglon = sRenglon & "" & "|" 'Fecha de pedimento
        sRenglon = sRenglon & "" & "|" 'aduana
        
        
        outputFile.WriteLine sRenglon
        adorcs.MoveNext
    Loop
    
    adorcs.Close
    
    Set adorcs = Nothing
    
    Set fs = Nothing
    
    MsgBox "Finalizo " & sNombreArc
    
End Sub

Private Sub Form_Load()
    Me.txtCtrl(0).Text = ObtieneParametro("RFC FISCAL")
End Sub


