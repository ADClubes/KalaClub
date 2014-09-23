VERSION 5.00
Begin VB.Form frmImpNCred 
   Caption         =   "Imprime Nota de crédito"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLines 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Margen superior"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "frmImpNCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lNumNotaIni As Long

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.lblMensaje.Caption = "Imprimir Nota de crédito " & lNumNotaIni
    Me.txtLines.Text = 500
End Sub

Private Sub Form_Load()
    
    CentraForma MDIPrincipal, Me
    
    
    
End Sub

Private Sub Ok_Click()

    Dim lCad As String
    
    Dim lMargenDer As Long
    Dim lYOffset As Long
    Dim lI As Integer

    Dim adorcs As ADODB.Recordset
    
    Dim sFolioFactura
    Dim dFechaFactura
    
    Dim dImporte As Double
    Dim dIva As Double
    Dim dSubtotal As Double
    Dim dTotal As Double
    
    Dim dDecimal As Double
    
    Dim iRen As Integer
    Dim iRenMax As Integer
    
    
    lMargenDer = Printer.Width - 500
    
    iRenMax = 11
    
    strSQL = "SELECT NOTAS_CRED.NumeroNota, NOTAS_CRED.Serie & Format(NOTAS_CRED.Folio,'0000') AS Folio, NOTAS_CRED.NoFamilia, NOTAS_CRED.FechaNota, NOTAS_CRED.NombreNota, NOTAS_CRED.CalleNota, NOTAS_CRED.ColoniaNota, NOTAS_CRED.DelNota, NOTAS_CRED.CiudadNota, NOTAS_CRED.EstadoNota, NOTAS_CRED.CodPos, NOTAS_CRED.RFC, NOTAS_CRED.Tel1, NOTAS_CRED.Observaciones, NOTAS_CRED.Total, NOTAS_CRED_DETALLE.Periodo, NOTAS_CRED_DETALLE.Concepto, NOTAS_CRED_DETALLE.Cantidad, NOTAS_CRED_DETALLE.Importe, NOTAS_CRED_DETALLE.Intereses, NOTAS_CRED_DETALLE.Descuento, NOTAS_CRED_DETALLE.Iva, NOTAS_CRED_DETALLE.IvaIntereses, NOTAS_CRED_DETALLE.IvaDescuento, USUARIOS_CLUB.Inscripcion, FACTURAS.Serie & FACTURAS.Folio As FolioFactura, FACTURAS.FechaFactura"
    strSQL = strSQL & " FROM ((NOTAS_CRED INNER JOIN NOTAS_CRED_DETALLE ON NOTAS_CRED.NumeroNota = NOTAS_CRED_DETALLE.NumeroNota) LEFT JOIN USUARIOS_CLUB ON NOTAS_CRED.IdTitular=USUARIOS_CLUB.IdMember)"
    strSQL = strSQL & " LEFT JOIN FACTURAS ON NOTAS_CRED.NumeroFactura=FACTURAS.NumeroFactura"
    strSQL = strSQL & " WHERE NOTAS_CRED.NumeroNota=" & lNumNotaIni
    
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    
    lYOffset = Val(Me.txtLines.Text)
    
    
    For lI = 1 To 2
    
        dImporte = 0
        dIva = 0
        dSubtotal = 0
        
        iRen = 0
        
        adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        sFolioFactura = IIf(IsNull(adorcs!foliofactura), vbNullString, adorcs!foliofactura)
        dFechaFactura = adorcs!FechaFactura
    
        Printer.FontName = "Arial Narrow"
    
    
    
        Printer.FontSize = 8
        
        Printer.CurrentY = lYOffset + 1000
        Printer.CurrentX = 7500
        
        
        
        Printer.CurrentX = 10000
        Printer.FontBold = True
        Printer.Print vbNullString
        Printer.FontBold = False
        
        
        Printer.CurrentY = lYOffset + 1600
        Printer.CurrentX = 6000
        Printer.FontBold = True
        Printer.Print adorcs!Folio;
        Printer.FontBold = False
        Printer.Print " (" & adorcs!NumeroNota & ")";
        
        lCad = Format(adorcs!FechaNota, "Long Date")
        Printer.CurrentX = (lMargenDer - Printer.TextWidth(lCad))
        Printer.Print lCad
        
        
        Printer.CurrentY = lYOffset + 2000
        
        Printer.FontSize = 10
        Printer.CurrentX = 4000
        Printer.Print adorcs!NoFamilia & " " & adorcs!Nombrenota
        
        Printer.FontSize = 8
        Printer.CurrentX = 4000
        Printer.Print adorcs!CalleNota;
        
        Printer.CurrentX = 9000
        Printer.Print "R.F.C " & adorcs!rfc
        
        Printer.CurrentX = 4000
        Printer.Print adorcs!ColoniaNota;
        
        Printer.CurrentX = 9000
        Printer.Print adorcs!EstadoNota
        
        Printer.CurrentX = 4000
        Printer.Print adorcs!CiudadNota, "C.P. " & Format(adorcs!Codpos, "00000"), "Tel: " & adorcs!Tel1
        
        Printer.FontSize = 7
        
        Printer.CurrentY = lYOffset + 3050
        
        Do Until adorcs.EOF
            Printer.CurrentX = 4000
            Printer.Print adorcs!Periodo;
            
            Printer.CurrentX = 5000
            Printer.Print adorcs!Concepto;
            
            dImporte = adorcs!Importe - adorcs!Iva
            dIva = dIva + adorcs!Iva
            dSubtotal = dSubtotal + dImporte
            
            
            Printer.CurrentX = 8000
            lCad = Format(dImporte, "$#,0.00")
            Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
            
            Printer.Print lCad
            
            
            adorcs.MoveNext
            iRen = iRen + 1
        Loop
        
        lCad = "1"
        
        Printer.CurrentY = Printer.CurrentY + Printer.TextHeight(lCad) * (iRenMax - iRen)
        
        Printer.CurrentX = 4000
        Printer.Print "FACTURA NO. ";
        Printer.Print sFolioFactura;
        Printer.Print " CON FECHA ";
        Printer.Print dFechaFactura;
        
    
        
        Printer.CurrentY = Printer.CurrentY + 300
        
        lCad = Format(dSubtotal, "$#,0.00")
        Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
        Printer.Print lCad
        
        
                
        
        lCad = Format(dIva, "$#,0.00")
        Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
        Printer.Print lCad
        
        
        dTotal = dSubtotal + dIva
        dDecimal = (dTotal - Int(dTotal)) * 100
        lCad = UCase(Num2Txt(Int(dTotal))) & " PESOS "
        lCad = lCad & Format(dDecimal, "00") & "/100 M.N."
        lCad = "(" & lCad & ")"
        Printer.CurrentX = 4000
        Printer.Print lCad;
        
        
        
        lCad = Format(dTotal, "$#,0.00")
        Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
        Printer.Print lCad
        
        
        adorcs.Close
        lYOffset = lYOffset + 7920
    Next
    
    Printer.EndDoc
    Set adorcs = Nothing
    Unload Me
    
End Sub
