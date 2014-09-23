VERSION 5.00
Begin VB.Form frmImpFac 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imprime Factura"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   713
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtNumFacFinal 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2633
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtNumFacIni 
      Enabled         =   0   'False
      Height          =   405
      Left            =   593
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Imprimir"
      Default         =   -1  'True
      Height          =   615
      Left            =   2633
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Se generaron las siguientes facturas:"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "# Final"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "# Inicial"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmImpFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lNumeroInicial As Long
Public lNumeroFinal As Long
Public cModo As String

Dim AdoRcs1 As ADODB.Recordset
Dim crxApplication As New CRAXDRT.Application
Dim CrxReport As CRAXDRT.Report
Dim CrxFormulaFields As CRAXDRT.FormulaFieldDefinitions
Dim CrxFormulaField As CRAXDRT.FormulaFieldDefinition

Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section

Dim crxReportObjs As CRAXDRT.ReportObjects

Dim crxSubreportObj As CRAXDRT.SubreportObject

Dim CrxSubReport As CRAXDRT.Report

Dim sqlQuery As String

Dim AdorcsPrint As ADODB.Recordset
Dim adorcsPrint2 As ADODB.Recordset
Dim adorcsFormaPago As ADODB.Recordset

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim Reporte As String, intIdRep As Integer
    
    
    
    Dim lI As Long
    Dim lX As Long
    Dim lY As Long
    
    Dim dTempInt As Double
    Dim sTempStr As String
    
    Dim sConLetra As String
    
    On Error GoTo CatchError
    
    Err.Clear
    
    
    Me.cmdOk.Enabled = False
    
    If Val(Me.txtNumFacIni.Text) > Val(Me.txtNumFacFinal.Text) Then
        MsgBox "La factura inicial debe ser menor que la final!", vbInformation, "Facturación"
        Me.txtNumFacIni.SelLength = Len(Me.txtNumFacIni.Text)
        Me.txtNumFacIni.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set AdorcsPrint = New ADODB.Recordset
    AdorcsPrint.ActiveConnection = Conn
    AdorcsPrint.CursorType = adOpenStatic
    AdorcsPrint.LockType = adLockReadOnly
    AdorcsPrint.CursorLocation = adUseServer
    
    
    
    'For lI = Val(Me.txtNumFacIni.Text) To Val(Me.txtNumFacFinal.Text)
    'For lI = Val(lNumFacIniImp) To Val(lNumFacFinImp)
    For lI = lNumeroInicial To lNumeroFinal
        If Me.Tag = "F" Then
            Reporte = App.Path & "\rptlocal\" & "facturaSportium" & ".rpt"
        Else
            Reporte = App.Path & "\rptlocal\" & "ReciboSportium" & ".rpt"
        End If
        
        MDIPrincipal.StatusBar1.Panels(1).Text = "Ejecutando Query"
        
        If Me.Tag = "F" Then
            #If SqlServer_ Then
                strSQL = "SELECT FACTURAS.NumeroFactura, FACTURAS.Serie + Convert(varchar, 8,FACTURAS.Folio) AS Folio, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Cantidad, FACTURAS_DETALLE.Importe, FACTURAS_DETALLE.Intereses, FACTURAS_DETALLE.Descuento, FACTURAS_DETALLE.Iva, FACTURAS_DETALLE.IvaIntereses, FACTURAS_DETALLE.IvaDescuento, USUARIOS_CLUB.Inscripcion"
                strSQL = strSQL & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) LEFT JOIN USUARIOS_CLUB ON FACTURAS.IdTitular=USUARIOS_CLUB.IdMember"
                strSQL = strSQL & " WHERE FACTURAS.NumeroFactura=" & lI
            #Else
                strSQL = "SELECT FACTURAS.NumeroFactura, FACTURAS.Serie & FACTURAS.Folio AS Folio, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Cantidad, FACTURAS_DETALLE.Importe, FACTURAS_DETALLE.Intereses, FACTURAS_DETALLE.Descuento, FACTURAS_DETALLE.Iva, FACTURAS_DETALLE.IvaIntereses, FACTURAS_DETALLE.IvaDescuento, USUARIOS_CLUB.Inscripcion"
                strSQL = strSQL & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) LEFT JOIN USUARIOS_CLUB ON FACTURAS.IdTitular=USUARIOS_CLUB.IdMember"
                strSQL = strSQL & " WHERE FACTURAS.NumeroFactura=" & lI
            #End If
        Else
            #If SqlServer_ Then
                strSQL = "SELECT FACTURAS.NumeroFactura, FACTURAS.Serie + Convert(varchar, 8,FACTURAS.Folio) AS Folio, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Cantidad, FACTURAS_DETALLE.Importe, FACTURAS_DETALLE.Intereses, FACTURAS_DETALLE.Descuento, FACTURAS_DETALLE.Iva, FACTURAS_DETALLE.IvaIntereses, FACTURAS_DETALLE.IvaDescuento, USUARIOS_CLUB.Inscripcion"
                strSQL = strSQL & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) LEFT JOIN USUARIOS_CLUB ON FACTURAS.IdTitular=USUARIOS_CLUB.IdMember"
                strSQL = strSQL & " WHERE FACTURAS.NumeroFactura=" & lI
            #Else
                strSQL = "SELECT FACTURAS.NumeroFactura, FACTURAS.Serie & FACTURAS.Folio AS Folio, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Cantidad, FACTURAS_DETALLE.Importe, FACTURAS_DETALLE.Intereses, FACTURAS_DETALLE.Descuento, FACTURAS_DETALLE.Iva, FACTURAS_DETALLE.IvaIntereses, FACTURAS_DETALLE.IvaDescuento, USUARIOS_CLUB.Inscripcion"
                strSQL = strSQL & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) LEFT JOIN USUARIOS_CLUB ON FACTURAS.IdTitular=USUARIOS_CLUB.IdMember"
                strSQL = strSQL & " WHERE FACTURAS.NumeroFactura=" & lI
            #End If
        End If
                
              
        If strSQL <> "" Then
            AdorcsPrint.Open strSQL
        End If
    
        If AdorcsPrint.EOF Then
            MsgBox "¡ No Se Encontró Información Para Este Reporte !", vbExclamation, "Reporte Vacío"
            Screen.MousePointer = vbDefault
            Unload Me
            Exit Sub
        End If
        
        If Me.Tag = "R" Then
            strSQL = "SELECT FORMA_PAGO.Descripcion, PAGOS_FACTURA.OpcionPago, PAGOS_FACTURA.Referencia, PAGOS_FACTURA.Importe"
            strSQL = strSQL & " FROM PAGOS_FACTURA INNER JOIN FORMA_PAGO ON PAGOS_FACTURA.IdFormaPago = FORMA_PAGO.IdFormaPago"
            strSQL = strSQL & " WHERE PAGOS_FACTURA.NumeroFactura = " & lI
        Else
            strSQL = "SELECT FORMA_PAGO.Descripcion, PAGOS_FACTURA.OpcionPago, PAGOS_FACTURA.Referencia, PAGOS_FACTURA.Importe"
            strSQL = strSQL & " FROM PAGOS_FACTURA INNER JOIN FORMA_PAGO ON PAGOS_FACTURA.IdFormaPago = FORMA_PAGO.IdFormaPago"
            strSQL = strSQL & " WHERE PAGOS_FACTURA.NumeroFactura = " & lI
        End If
        
        Set adorcsFormaPago = New ADODB.Recordset
        adorcsFormaPago.CursorLocation = adUseServer
            
        adorcsFormaPago.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        
        
        
        If Me.Tag = "F" Then
            'ImprimeFactura
            'ImprimeCFD22
            ImprimeComprobanteCajero
        End If
        If Me.Tag = "R" Then
            ImprimeRecibo
        End If
    Next
    
    
    adorcsFormaPago.Close
    Set adorcsFormaPago = Nothing
    
    
    Set AdorcsPrint = Nothing
    
    Screen.MousePointer = vbDefault
    
    MDIPrincipal.StatusBar1.Panels(1).Text = ""
    
    Unload Me
    
    Exit Sub
    
CatchError:

    MsgBox "Error: " & Err.Description, vbCritical, "Imprime"
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Activate()
    
    If Me.Tag = "F" Then
        Me.lblTitulo.Caption = "Se generaron las siguientes facturas:"
        Me.txtNumFacIni.Text = lNumFolioFacIniImp 'lNumFacIniImp
        Me.txtNumFacFinal.Text = lNumFolioFacFinImp 'lNumFacFinImp
    Else
        Me.lblTitulo.Caption = "Se generaron los siguientes recibos:"
        Me.txtNumFacIni.Text = lNumFolioFacIniImp
        Me.txtNumFacFinal.Text = lNumFolioFacFinImp
    End If

    Me.cmdOk.SetFocus

End Sub

Private Sub Form_Load()
    
    Me.Height = 3450
    Me.Width = 4770
    
    
    CentraForma MDIPrincipal, Me
    
    
End Sub



Private Sub txtNumFacIni_KeyPress(KeyAscii As Integer)
        
        
        KeyAscii = 0
        
        Exit Sub
        
        Select Case KeyAscii
        Case 13
            KeyAscii = 0
            SendKeys vbTab
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txtNumFacFinal_KeyPress(KeyAscii As Integer)
        
        KeyAscii = 0
        
        Exit Sub

        
        
        Select Case KeyAscii
        Case 13
            KeyAscii = 0
            SendKeys vbTab
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select

End Sub


Private Sub ImprimeRecibo()
    
    
    Dim sCadena As String
    Dim sConLetra As String
    Dim dTempInt As Double
    Dim dTotal As Double
    Dim lTurno As Long


    Dim lCol1 As Long
    Dim lCol2 As Long
    Dim lCol3 As Long
    Dim lCol4 As Long


    Dim sImpresoraActual As String
    Dim sImpresoraCupon As String


    Dim sAdPicture As String
    Dim picLogo As Picture


    lCol1 = 0
    lCol2 = 500
    lCol3 = 3000
    lCol4 = 4000

    Dim iRespuesta As Integer


If sTicket <> "" Then
    sImpresoraCupon = sTicket '"Apos Premium"
    Else
    sImpresoraCupon = ObtieneParametro("IMPRESORA CUPON")
    End If

    If sImpresoraCupon = vbNullString Then
        iRespuesta = MsgBox("No se ha definido una impresora" & vbLf & "para cupones, desea continuar?", vbQuestion + vbOKCancel, "Confirme")

        If iRespuesta = vbCancel Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

    End If


    If sImpresoraCupon <> Printer.DeviceName Then
        sImpresoraActual = SelectPrinter(sImpresoraCupon)

        'No se encontró la impresora configurada
        If sImpresoraActual = Printer.DeviceName Then
            iRespuesta = MsgBox("No se ha encontrado la impresora" & vbLf & "para cupones, desea continuar?", vbQuestion + vbOKCancel, "Confirme")

            If iRespuesta = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If




    Printer.Font = "Arial"
    Printer.FontSize = 7

    sCadena = Trim(ObtieneParametro("EMPRESA FISCAL"))
    ImprimeTexto sCadena, 1, 0, Printer.Width

    sCadena = Trim(ObtieneParametro("CALLE FISCAL"))
    ImprimeTexto sCadena, 1, 0, Printer.Width


    sCadena = "COL. " & ObtieneParametro("COLONIA FISCAL")
    ImprimeTexto sCadena, 1, 0, Printer.Width



    sCadena = ObtieneParametro("DELEGACION FISCAL")
    'ImprimeTexto sCadena, 1, 0, Printer.Width

    sCadena = sCadena & " C.P. " & ObtieneParametro("CP FISCAL")
    ImprimeTexto sCadena, 1, 0, Printer.Width

    sCadena = ObtieneParametro("CIUDAD FISCAL")
    ImprimeTexto sCadena, 1, 0, Printer.Width

    sCadena = ObtieneParametro("RFC FISCAL")
    ImprimeTexto sCadena, 1, 0, Printer.Width

    'lTurno = AdorcsPrint!Turno
    dTempInt = AdorcsPrint!Total - Int(AdorcsPrint!Total)


    If dTempInt > 0 Then
        dTempInt = dTempInt * 100
    End If

    sCadena = "PESOS " & Format(dTempInt, "#00") & "/100 M.N.)"
    sConLetra = Trim("(" & UCase(Num2Txt(Int(AdorcsPrint!Total))) & sCadena)

    dTotal = AdorcsPrint!Total



    Printer.Print
    Printer.Print Format(AdorcsPrint!FechaFactura, "Long Date")

    Printer.Print
    Printer.Print "Folio: ";
    Printer.Print AdorcsPrint!NumeroFactura

    Printer.Print
    Printer.Print "Inscripción: ";
    Printer.Print AdorcsPrint!NoFamilia

    Printer.Print
    Printer.Print "Nombre: ";
    Printer.Print AdorcsPrint!NombreFactura



    Printer.Print



    Printer.Font = "Courier New"
    Printer.FontSize = 8

    Do Until AdorcsPrint.EOF
        Printer.CurrentX = lCol1
        Printer.Print AdorcsPrint!Cantidad;

        Printer.CurrentX = lCol3
        sCadena = Format(AdorcsPrint!Importe + AdorcsPrint!Intereses - AdorcsPrint!Descuento, "$#,0.00")
        Printer.CurrentX = lCol4 - Printer.TextWidth(sCadena)
        Printer.Print sCadena;



        Printer.CurrentX = lCol2

        sCadena = Trim(AdorcsPrint!Concepto)

        ImprimeTexto sCadena, 0, lCol2, lCol3

        AdorcsPrint.MoveNext

    Loop



    Printer.Print


    Printer.Print "Total:";

    Printer.CurrentX = lCol3
    sCadena = Format(dTotal, "$#,0.00")
    Printer.CurrentX = lCol4 - Printer.TextWidth(sCadena)
    Printer.Print sCadena

    Printer.Print
    ImprimeTexto sConLetra, 0, 0, Printer.Width

    Printer.Print


    'Imprime la forma de pago
    Do Until adorcsFormaPago.EOF
        Printer.Print adorcsFormaPago!Descripcion
        adorcsFormaPago.MoveNext
    Loop

    Printer.Print


    Printer.Font = "Arial"
    Printer.FontSize = 7
    Printer.FontBold = True
    Printer.Print
    Printer.Print "ESTE RECIBO NO ES VALIDO PARA CANJEARSE"
    Printer.Print "POR NINGUN PRODUCTO O SERVICIO"
    Printer.FontBold = False



    Printer.Print
    Printer.Print "ESTE RECIBO NO ES UN COMPROBANTE FISCAL"

    Printer.Print
    ImprimeTexto "COPIA PARA EL CLIENTE", 1, 0, Printer.Width

    'Para la impresión de publicidad
    sAdPicture = App.Path & "\" & "adpicture.jpg"

    If sAdPicture <> vbNullString Then
        If Dir(sAdPicture) <> vbNullString Then
            Set picLogo = LoadPicture(sAdPicture)
            Printer.PaintPicture picLogo, 0, Printer.CurrentY + 100, Printer.ScaleWidth, Printer.ScaleWidth
        End If
    End If

    Printer.EndDoc

    sImpresoraActual = SelectPrinter(sImpresoraActual)
    
End Sub
Private Sub ImprimeFactura()
'    Dim lCad As String
'
'    Dim lMargenDer As Long
'    Dim lYOffset As Long
'    Dim lI As Integer
'
'    Dim adorcs As ADODB.Recordset
'
'    Dim sFolioFactura
'    Dim dFechaFactura
'
'    Dim dImporte As Double
'    Dim dIva As Double
'    Dim dSubtotal As Double
'    Dim dTotal As Double
'
'    Dim dDecimal As Double
'
'    Dim iRen As Integer
'    Dim iRenMax As Integer
'
'    Dim sObservaciones As String
'
'    Dim iTipoFactura As Integer 'Tipo de factura 0 = normal, 1 = Varios
'
'    Dim sCadFormaPago As String
'
'    '04/06/10
'    Dim lDesglose As Boolean 'True cuando hay que desglosar la factura
'
'    Dim sStrQry As String
'
'    lMargenDer = Printer.Width - 1000
'
'    iRenMax = 11
'
'
'    iTipoFactura = GetTipoFactura(lNumeroInicial)
'    lDesglose = True
'
'
'    If iTipoFactura = 0 Then
'        Do While Not adorcsFormaPago.EOF
'            sCadFormaPago = sCadFormaPago & adorcsFormaPago!Descripcion & " " & adorcsFormaPago!OpcionPago & " " & adorcsFormaPago!Referencia & ","
'            adorcsFormaPago.MoveNext
'        Loop
'        If sCadFormaPago <> vbNullString Then
'            sCadFormaPago = Left$(sCadFormaPago, Len(sCadFormaPago) - 1)
'        End If
'    Else
'        sCadFormaPago = "FORMAS DE PAGO VARIAS"
'    End If
'
'
'
'    If iTipoFactura = 0 Then
'        sStrQry = "SELECT FACTURAS.NumeroFactura,   FACTURAS.Folio & FACTURAS.Serie AS Folio, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Cantidad, FACTURAS_DETALLE.Importe, FACTURAS_DETALLE.Intereses, FACTURAS_DETALLE.Descuento, FACTURAS_DETALLE.Iva, FACTURAS_DETALLE.IvaIntereses, FACTURAS_DETALLE.IvaDescuento, USUARIOS_CLUB.Inscripcion, iif(isNull(FACTURAS.TipoPersona),'F',FACTURAS.TipoPersona) AS TipoPersona"
'        sStrQry = sStrQry & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) LEFT JOIN USUARIOS_CLUB ON FACTURAS.IdTitular=USUARIOS_CLUB.IdMember"
'        sStrQry = sStrQry & " WHERE FACTURAS.NumeroFactura=" & lNumeroInicial
'    Else
'        sStrQry = "SELECT FACTURAS.NumeroFactura, FACTURAS.Folio & FACTURAS.Serie As Folio, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS.FechaFactura AS Periodo, 'VARIOS' AS Concepto, 1 AS Cantidad, SUM(FACTURAS_DETALLE.Importe*FACTURAS_DETALLE.Cantidad) AS Importe, Sum(FACTURAS_DETALLE.Intereses) AS Intereses, Sum(FACTURAS_DETALLE.Descuento) AS Descuento, Sum(FACTURAS_DETALLE.Iva) AS Iva, Sum(FACTURAS_DETALLE.IvaIntereses) AS IvaIntereses, Sum(FACTURAS_DETALLE.IvaDescuento) AS IvaDescuento, USUARIOS_CLUB.Inscripcion, iif(isNull(FACTURAS.TipoPersona),'F',FACTURAS.TipoPersona) AS TipoPersona"
'        sStrQry = sStrQry & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) INNER JOIN USUARIOS_CLUB ON FACTURAS.IdTitular = USUARIOS_CLUB.IdMember"
'        sStrQry = sStrQry & " GROUP BY FACTURAS.NumeroFactura, FACTURAS.Folio & FACTURAS.Serie, FACTURAS.FechaFactura, FACTURAS.NoFamilia, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total,  'VARIOS', 1,  USUARIOS_CLUB.Inscripcion, iif(isNull(FACTURAS.TipoPersona),'F',FACTURAS.TipoPersona)"
'        sStrQry = sStrQry & " HAVING FACTURAS.NumeroFactura=" & lNumeroInicial
'    End If
'
'    Set adorcs = New ADODB.Recordset
'    adorcs.CursorLocation = adUseServer
'
'
'    lYOffset = 0
'
'
'    For lI = 1 To 2
'
'        dImporte = 0
'        dIva = 0
'        dSubtotal = 0
'
'        iRen = 0
'
'
'
'
'
'        adorcs.Open sStrQry, Conn, adOpenForwardOnly, adLockReadOnly
'
'        If adorcs!TipoPersona = "F" And Len(adorcs!rfc) <> 13 Then
'            lDesglose = False
'        End If
'
'        If adorcs!TipoPersona = "M" And Len(adorcs!rfc) <> 12 Then
'            lDesglose = False
'        End If
'
'
'        sFolioFactura = IIf(IsNull(adorcs!Folio), vbNullString, adorcs!Folio)
'        dFechaFactura = adorcs!FechaFactura
'        sObservaciones = adorcs!Observaciones
'
'        Printer.FontName = "Arial Narrow"
'
'
'
'        Printer.FontSize = 8
'
'        Printer.CurrentY = lYOffset + 1500
'        Printer.CurrentX = 6500
'
'        Printer.FontBold = True
'        Printer.Print adorcs!Folio;
'        Printer.FontBold = False
'        Printer.Print " (" & adorcs!Numerofactura & ")";
'
'
'        Printer.CurrentX = 10000
'        Printer.FontBold = True
'        Printer.Print vbNullString
'        Printer.FontBold = False
'
'
'        Printer.CurrentY = lYOffset + 2000
'        Printer.CurrentX = 6000
'
'
'        lCad = Format(adorcs!FechaFactura, "Long Date")
'        Printer.CurrentX = (lMargenDer - Printer.TextWidth(lCad))
'        Printer.Print lCad
'
'
'        Printer.CurrentY = lYOffset + 2500
'
'        Printer.FontSize = 10
'        Printer.CurrentX = 4000
'        Printer.Print adorcs!NoFamilia & " " & adorcs!NombreFactura
'        'Printer.Print adorcs!NombreFactura
'
'        Printer.FontSize = 8
'        Printer.CurrentX = 4000
'        Printer.Print adorcs!CalleFactura;
'
'        Printer.CurrentX = 9000
'        Printer.Print "R.F.C " & adorcs!rfc
'
'        Printer.CurrentX = 4000
'        Printer.Print "COL. " & adorcs!ColoniaFactura;
'
'        Printer.CurrentX = 9000
'        Printer.Print adorcs!EstadoFactura
'
'        Printer.CurrentX = 4000
'        Printer.Print adorcs!CiudadFactura, "C.P. " & Format(adorcs!Codpos, "00000"), "Tel: " & adorcs!Tel1
'
'        Printer.FontSize = 7
'
'        Printer.CurrentY = lYOffset + 3450
'
'        Do Until adorcs.EOF
'            Printer.CurrentX = 4000
'            Printer.Print adorcs!Periodo;
'
'            Printer.CurrentX = 5000
'            Printer.Print adorcs!Concepto;
'
'
'            If lDesglose Then
'                dImporte = ((adorcs!Importe * adorcs!Cantidad) + adorcs!Intereses - adorcs!Descuento) - adorcs!Iva - adorcs!IvaIntereses + adorcs!IvaDescuento
'                dIva = dIva + adorcs!Iva + adorcs!IvaIntereses - adorcs!IvaDescuento
'                dSubtotal = dSubtotal + dImporte
'            Else
'                dImporte = ((adorcs!Importe * adorcs!Cantidad) + adorcs!Intereses - adorcs!Descuento)
'                dIva = 0
'                dSubtotal = dSubtotal + dImporte
'            End If
'
'            Printer.CurrentX = 8000
'            lCad = Format(dImporte, "$#,0.00")
'            Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
'
'            Printer.Print lCad
'
'
'            adorcs.MoveNext
'            iRen = iRen + 1
'        Loop
'
'
'        Printer.CurrentX = 4000
'        Printer.CurrentY = lYOffset + 5300
'        Printer.Print sObservaciones
'
'        lCad = "1"
'
'        Printer.CurrentY = lYOffset + 5400
'
'
'        dTotal = dSubtotal + dIva
'
'
'        If (lDesglose) Then
'            lCad = Format(dSubtotal, "$#,0.00")
'            Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
'            Printer.Print lCad
'        End If
'
'        Printer.CurrentY = lYOffset + 5740
'
'
'        If (lDesglose) Then
'
'            lCad = ObtieneParametro("IVA_GENERAL") & "%"
'            Printer.CurrentX = (lMargenDer - 2300 - Printer.TextWidth(lCad))
'            Printer.Print lCad;
'
'
'            lCad = Format(dIva, "$#,0.00")
'            Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
'            Printer.Print lCad
'        End If
'
'        Printer.CurrentY = lYOffset + 6080
'
'        lCad = Format(dTotal, "$#,0.00")
'        Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
'        Printer.Print lCad
'
'
'
'        dDecimal = (dTotal - Int(dTotal)) * 100
'        lCad = UCase(Num2Txt(Int(dTotal))) & " PESOS "
'        lCad = lCad & Format(dDecimal, "00") & "/100 M.N."
'        Printer.CurrentX = 4000
'        lCad = "(" & lCad & ")"
'        Printer.CurrentY = lYOffset + 5900
'        Printer.Print lCad;
'
'
'        Printer.CurrentX = 8000
'        Printer.CurrentY = lYOffset + 7250
'        Printer.Print sCadFormaPago
'
'
'
'
'        adorcs.Close
'        lYOffset = lYOffset + 7950
'    Next
'
'    Printer.EndDoc
'    Set adorcs = Nothing
'    Unload Me


End Sub

Private Sub ImprimeCFD()

    Dim parser As DOMDocument
    Dim oXlst As DOMDocument
    Dim sRutaCFD As String
    Dim sRutaXLST As String
    Dim sNombreCFD As String
    
    
    Dim comprobante As IXMLDOMNode
    Dim emisor As IXMLDOMNode
    Dim receptor As IXMLDOMNode
    Dim conceptos As IXMLDOMNode
    Dim impuestos As IXMLDOMNode
    Dim addenda As IXMLDOMNode
    
    Dim nodoInsc As IXMLDOMNode
    
    
    Dim lCol1 As Long
    Dim lCol2 As Long
    Dim lCol3 As Long
    Dim lCol4 As Long
    Dim lCol5 As Long
    
    Dim sCadPrint As String
    Dim dTempInt As Double
    
    lCol1 = 0
    lCol2 = 300
    lCol3 = 3000
    lCol4 = 4000
    lCol5 = 5000
    
    Dim iRespuesta As Integer
    Dim sImpresoraCupon As String
    Dim sImpresoraActual As String
    
    
    If sTicket <> "" Then
    sImpresoraCupon = sTicket '"Apos Premium"
    Else
    sImpresoraCupon = ObtieneParametro("IMPRESORA CUPON")
    End If

    If sImpresoraCupon = vbNullString Then
        iRespuesta = MsgBox("No se ha definido una impresora" & vbLf & "para cupones, desea continuar?", vbQuestion + vbOKCancel, "Confirme")

        If iRespuesta = vbCancel Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

    End If


    If sImpresoraCupon <> Printer.DeviceName Then
        sImpresoraActual = SelectPrinter(sImpresoraCupon)

        'No se encontró la impresora configurada
        If sImpresoraActual = Printer.DeviceName Then
            iRespuesta = MsgBox("No se ha encontrado la impresora" & vbLf & "para cupones, desea continuar?", vbQuestion + vbOKCancel, "Confirme")

            If iRespuesta = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    
    

    Set parser = New DOMDocument
    Set oXlst = New DOMDocument
    
    parser.async = False
    oXlst.async = False
    
    
    sRutaCFD = ObtieneParametro("RUTA_CFD")
    sRutaXLST = ObtieneParametro("RUTA_XLST")
    
    sNombreCFD = NombreArchivoCFD(lNumeroInicial, "F", 0)
    
    If Not parser.Load(sRutaCFD & "\" & sNombreCFD & ".xml") Then
        MsgBox "Error no se pudo cargar el archivo" & vbCrLf & sNombreCFD, vbExclamation, "Error"
        Exit Sub
    End If
    
   If Not oXlst.Load(sRutaXLST & "\cadenaoriginal_2_0.xslt") Then
        MsgBox "Error no se pudo cargar el archivo XLST", vbExclamation, "Error"
        Exit Sub
   End If
    
    
    Set comprobante = parser.selectSingleNode("Comprobante")
    Set emisor = parser.selectSingleNode("Comprobante/Emisor")
    Set receptor = parser.selectSingleNode("Comprobante/Receptor")
    Set conceptos = parser.selectSingleNode("Comprobante/Conceptos")
    Set impuestos = parser.selectSingleNode("Comprobante/Impuestos")
    Set addenda = parser.selectSingleNode("Comprobante/Addenda")
    
    Set nodoInsc = addenda.selectSingleNode("md:NoInscripcion")
    
    Printer.FontSize = 8
    
    
    'Datos CFD
    Printer.Print "Serie: " & comprobante.Attributes.getNamedItem("serie").Text
    Printer.Print "Folio: " & comprobante.Attributes.getNamedItem("folio").Text
    Printer.Print "Fecha: " & comprobante.Attributes.getNamedItem("fecha").Text
    Printer.Print "No. Aprobacion: " & comprobante.Attributes.getNamedItem("noAprobacion").Text
    Printer.Print "Año de aprobación: " & comprobante.Attributes.getNamedItem("anoAprobacion").Text
    Printer.Print "Certificado: " & comprobante.Attributes.getNamedItem("noCertificado").Text
    Printer.Print
    
    'Emisor
    Printer.FontBold = True
    Printer.Print "Emisor"
    Printer.FontBold = False
    Printer.Print emisor.Attributes.getNamedItem("rfc").Text
    ImprimeTexto emisor.Attributes.getNamedItem("nombre").Text, 0, 0, 4000
    ImprimeDireccion emisor.selectSingleNode("DomicilioFiscal")
    Printer.Print
    
    'Expedida en
    If emisor.childNodes.Length = 2 Then
        Printer.FontBold = True
        Printer.Print "Expedida en"
        Printer.FontBold = False
        ImprimeDireccion emisor.selectSingleNode("ExpedidoEn")
        Printer.Print
    End If
    
    'Receptor
    Printer.FontBold = True
    Printer.Print "Cliente"
    Printer.FontBold = False
    
    If Not nodoInsc Is Nothing Then
        Printer.Print nodoInsc.Text
    End If
    
    Printer.Print receptor.Attributes.getNamedItem("rfc").Text
    ImprimeTexto receptor.Attributes.getNamedItem("nombre").Text, 0, 0, 4000
    ImprimeDireccion receptor.selectSingleNode("Domicilio")
    Printer.Print
    
    'Conceptos
    Printer.FontSize = 7
    ImprimeConceptos conceptos
    
    Printer.Print
    
    'Subttotal
    sCadPrint = "Subtotal:" & Format(comprobante.Attributes.getNamedItem("subTotal").Text, "$#,0.00")
    Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
    Printer.Print sCadPrint
    Printer.Print
    
    'Descuento
    If Not comprobante.Attributes.getNamedItem("descuento") Is Nothing Then
        sCadPrint = "Descuento:" & Format(comprobante.Attributes.getNamedItem("descuento").Text, "$#,0.00")
        Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
        Printer.Print sCadPrint
        Printer.Print
    End If
    
    'Impuestos
    ImprimeImpuestos impuestos
    Printer.Print
    
    'Total
    sCadPrint = "Total: " & Format(comprobante.Attributes.getNamedItem("total").Text, "$#,0.00")
    Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
    Printer.Print sCadPrint
    Printer.Print
    
    
    
    dTempInt = CDbl(comprobante.Attributes.getNamedItem("total").Text) - Int(CDbl(comprobante.Attributes.getNamedItem("total").Text))
    
    
    If dTempInt > 0 Then
        dTempInt = dTempInt * 100
    End If
        
    sCadPrint = "PESOS " & Format(dTempInt, "#00") & "/100 M.N.)"
    sCadPrint = Trim("(" & UCase(Num2Txt(Int(CDbl(comprobante.Attributes(8).Text)))) & sCadPrint)
    
    ImprimeTexto sCadPrint, 0, 0, Printer.Width
    
    Printer.Print
    
    Printer.Font = "Arial"
    Printer.FontSize = 7
    Printer.FontBold = True
    Printer.Print "Sello Digital"
    Printer.FontBold = False
    Printer.FontSize = 6
    MultipleLinea comprobante.Attributes.getNamedItem("sello").Text
    
    
    
    Printer.Print
    
    Printer.FontSize = 6
    Printer.FontBold = True
    Printer.Print "Cadena Original"
    Printer.FontBold = False
    Printer.FontSize = 6
    MultipleLinea parser.transformNode(oXlst)
    
    Printer.Print
    ImprimeTexto "Pago " & comprobante.Attributes.getNamedItem("formaDePago").Text, 0, 0, Printer.Width
    
    Printer.Print
    ImprimeTexto "ESTE DOCUMENTO ES UNA REPRESENTACION IMPRESA DE UN CFD", 0, 1, 4000
    
    Printer.EndDoc
    
    sImpresoraActual = SelectPrinter(sImpresoraActual)
    
End Sub
'----------------------------------------------------
Private Sub ImprimeCFD22()

    Dim parser As DOMDocument
    Dim oXlst As DOMDocument
    Dim sRutaCFD As String
    Dim sRutaXLST As String
    Dim sNombreCFD As String
    
    
    Dim comprobante As IXMLDOMNode
    Dim emisor As IXMLDOMNode
    Dim receptor As IXMLDOMNode
    Dim conceptos As IXMLDOMNode
    Dim impuestos As IXMLDOMNode
    Dim addenda As IXMLDOMNode
    
    Dim nodoInsc As IXMLDOMNode
    
    
    Dim lCol1 As Long
    Dim lCol2 As Long
    Dim lCol3 As Long
    Dim lCol4 As Long
    Dim lCol5 As Long
    
    Dim sCadPrint As String
    Dim dTempInt As Double
    
    lCol1 = 0
    lCol2 = 300
    lCol3 = 3000
    lCol4 = 4000
    lCol5 = 5000
    
    Dim iRespuesta As Integer
    Dim sImpresoraCupon As String
    Dim sImpresoraActual As String
    
    
    If sTicket <> "" Then
    sImpresoraCupon = sTicket '"Apos Premium"
    Else
    sImpresoraCupon = ObtieneParametro("IMPRESORA CUPON")
    End If

    If sImpresoraCupon = vbNullString Then
        iRespuesta = MsgBox("No se ha definido una impresora" & vbLf & "para cupones, desea continuar?", vbQuestion + vbOKCancel, "Confirme")

        If iRespuesta = vbCancel Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

    End If


    If sImpresoraCupon <> Printer.DeviceName Then
        sImpresoraActual = SelectPrinter(sImpresoraCupon)

        'No se encontró la impresora configurada
        If sImpresoraActual = Printer.DeviceName Then
            iRespuesta = MsgBox("No se ha encontrado la impresora" & vbLf & "para cupones, desea continuar?", vbQuestion + vbOKCancel, "Confirme")

            If iRespuesta = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If
    
    

    Set parser = New DOMDocument
    Set oXlst = New DOMDocument
    
    parser.async = False
    oXlst.async = False
    
    
    sRutaCFD = ObtieneParametro("RUTA_CFD")
    sRutaXLST = ObtieneParametro("RUTA_XLST")
    
    sNombreCFD = NombreArchivoCFD(lNumeroInicial, "F", 0)
    
    If Not parser.Load(sRutaCFD & "\" & sNombreCFD & ".xml") Then
        MsgBox "Error no se pudo cargar el archivo" & vbCrLf & sNombreCFD, vbExclamation, "Error"
        Exit Sub
    End If
    
   If Not oXlst.Load(sRutaXLST & "\cadenaoriginal_2_2.xslt") Then
        MsgBox "Error no se pudo cargar el archivo XLST", vbExclamation, "Error"
        Exit Sub
   End If
    
    
    Set comprobante = parser.selectSingleNode("Comprobante")
    Set emisor = parser.selectSingleNode("Comprobante/Emisor")
    Set receptor = parser.selectSingleNode("Comprobante/Receptor")
    Set conceptos = parser.selectSingleNode("Comprobante/Conceptos")
    Set impuestos = parser.selectSingleNode("Comprobante/Impuestos")
    Set addenda = parser.selectSingleNode("Comprobante/Addenda")
    
    Set nodoInsc = addenda.selectSingleNode("md:NoInscripcion")
    
    Printer.FontSize = 8
    
    
    'Datos CFD
    Printer.Print "Serie: " & comprobante.Attributes.getNamedItem("serie").Text
    Printer.Print "Folio: " & comprobante.Attributes.getNamedItem("folio").Text
    Printer.Print "Fecha: " & comprobante.Attributes.getNamedItem("fecha").Text
    Printer.Print "No. Aprobacion: " & comprobante.Attributes.getNamedItem("noAprobacion").Text
    Printer.Print "Año de aprobación: " & comprobante.Attributes.getNamedItem("anoAprobacion").Text
    Printer.Print "Certificado: " & comprobante.Attributes.getNamedItem("noCertificado").Text
    Printer.Print
    
    'Emisor
    Printer.FontBold = True
    Printer.Print "Emisor"
    Printer.FontBold = False
    Printer.Print emisor.Attributes.getNamedItem("rfc").Text
    ImprimeTexto emisor.Attributes.getNamedItem("nombre").Text, 0, 0, 4000
    ImprimeDireccion emisor.selectSingleNode("DomicilioFiscal")
    ImprimeTexto emisor.selectSingleNode("RegimenFiscal").Attributes.getNamedItem("Regimen").Text, 0, 0, 4000
    Printer.Print
    
    'Expedida en
    If emisor.childNodes.Length = 2 Then
        Printer.FontBold = True
        Printer.Print "Expedida en"
        Printer.FontBold = False
        ImprimeDireccion emisor.selectSingleNode("ExpedidoEn")
        Printer.Print
    End If
    
    'Receptor
    Printer.FontBold = True
    Printer.Print "Cliente"
    Printer.FontBold = False
    
    If Not nodoInsc Is Nothing Then
        Printer.Print nodoInsc.Text
    End If
    
    Printer.Print receptor.Attributes.getNamedItem("rfc").Text
    ImprimeTexto receptor.Attributes.getNamedItem("nombre").Text, 0, 0, 4000
    ImprimeDireccion receptor.selectSingleNode("Domicilio")
    Printer.Print
    
    'Conceptos
    Printer.FontSize = 7
    ImprimeConceptos22 conceptos
    
    Printer.Print
    
    'Subttotal
    sCadPrint = "Subtotal:" & Format(comprobante.Attributes.getNamedItem("subTotal").Text, "$#,0.00")
    Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
    Printer.Print sCadPrint
    Printer.Print
    
    'Descuento
    If Not comprobante.Attributes.getNamedItem("descuento") Is Nothing Then
        sCadPrint = "Descuento:" & Format(comprobante.Attributes.getNamedItem("descuento").Text, "$#,0.00")
        Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
        Printer.Print sCadPrint
        Printer.Print
    End If
    
    'Impuestos
    ImprimeImpuestos22 impuestos
    Printer.Print
    
    'Total
    sCadPrint = "Total: " & Format(comprobante.Attributes.getNamedItem("total").Text, "$#,0.00")
    Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
    Printer.Print sCadPrint
    Printer.Print
    
    
    
    dTempInt = CDbl(comprobante.Attributes.getNamedItem("total").Text) - Int(CDbl(comprobante.Attributes.getNamedItem("total").Text))
    
    
    If dTempInt > 0 Then
        dTempInt = dTempInt * 100
    End If
        
    sCadPrint = "PESOS " & Format(dTempInt, "#00") & "/100 M.N.)"
    sCadPrint = Trim("(" & UCase(Num2Txt(Int(CDbl(comprobante.Attributes(8).Text)))) & sCadPrint)
    
    ImprimeTexto sCadPrint, 0, 0, Printer.Width
    
    Printer.Print
    
    Printer.Font = "Arial"
    Printer.FontSize = 7
    Printer.FontBold = True
    Printer.Print "Sello Digital"
    Printer.FontBold = False
    Printer.FontSize = 6
    MultipleLinea comprobante.Attributes.getNamedItem("sello").Text
    
    
    
    Printer.Print
    
    Printer.FontSize = 6
    Printer.FontBold = True
    Printer.Print "Cadena Original"
    Printer.FontBold = False
    Printer.FontSize = 6
    MultipleLinea parser.transformNode(oXlst)
    
    Printer.Print
    ImprimeTexto "Pago " & comprobante.Attributes.getNamedItem("formaDePago").Text, 0, 0, Printer.Width
    
    Printer.Print
    ImprimeTexto "ESTE DOCUMENTO ES UNA REPRESENTACION IMPRESA DE UN CFD", 0, 1, 4000
    
    Printer.EndDoc
    
    sImpresoraActual = SelectPrinter(sImpresoraActual)
    
End Sub



'----------------------------------------------------

Private Sub ImprimeDireccion(nodoDir As IXMLDOMNode)
    
    On Error Resume Next
    
    ImprimeTexto nodoDir.Attributes.getNamedItem("calle").Text, 0, 0, 4000
    Printer.Print nodoDir.Attributes.getNamedItem("noExterior").Text
    Printer.Print nodoDir.Attributes.getNamedItem("noInterior").Text
    Printer.Print nodoDir.Attributes.getNamedItem("colonia").Text
    Printer.Print nodoDir.Attributes.getNamedItem("municipio").Text
    Printer.Print nodoDir.Attributes.getNamedItem("codigoPostal").Text
    Printer.Print nodoDir.Attributes.getNamedItem("estado").Text
    Printer.Print nodoDir.Attributes.getNamedItem("pais").Text
    
End Sub
Private Sub ImprimeConceptos(nodoConceptos As IXMLDOMNode)
    Dim iNodos As Integer
    Dim iConceptos As Integer
    Dim nodoCon As IXMLDOMNode
    
    Dim lCol1 As Long
    Dim lCol2 As Long
    Dim lCol3 As Long
    Dim lCol4 As Long
    Dim lCol5 As Long
    
    Dim lY As Long
    Dim lYNext As Long
    
    Dim sCadPrint As String
    
    lCol1 = 0
    lCol2 = 300
    lCol3 = 2000
    lCol4 = 4000
    lCol5 = 5000
    
    Printer.Font = "Courier New"
    Printer.FontSize = 6
    
    Printer.Print String((Printer.Width / Printer.TextWidth("-")), "-")
    
    Printer.CurrentX = lCol1
    Printer.Print "Concepto"
    
    Printer.CurrentX = lCol1
    Printer.Print "Cant.";
    
    sCadPrint = "P.Unitario"
    Printer.CurrentX = lCol3 - Printer.TextWidth(sCadPrint)
    Printer.Print sCadPrint;
    
    
    sCadPrint = "Importe"
    Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
    Printer.Print sCadPrint
    
    
    Printer.Print String((Printer.Width / Printer.TextWidth("-")), "-")
    
    
    For Each nodoCon In nodoConceptos.childNodes
        
        lY = Printer.CurrentY
        
        
        'Descripcion
        Printer.CurrentX = lCol1
        ImprimeTexto Trim(nodoCon.Attributes(1).Text), 0, lCol1, Printer.Width
        lYNext = Printer.CurrentY
        
        'Cantidad
        Printer.CurrentX = lCol1
        'Printer.CurrentY = lY
        'Printer.Print
        Printer.Print Format(nodoCon.Attributes(3).Text, "00");
        
        'Precio unitario
        sCadPrint = Format(nodoCon.Attributes(2).Text, "$#,0.00")
        Printer.CurrentX = lCol3 - Printer.TextWidth(sCadPrint)
        'Printer.CurrentY = lY
        Printer.Print sCadPrint;
        
        
        sCadPrint = Format(nodoCon.Attributes(4).Text, "$#,0.00")
        'Printer.CurrentY = lY
        Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
        Printer.Print sCadPrint
        
        'Printer.CurrentY = lYNext
        
    Next
    
    
End Sub
'---------------------------------------------------------
Private Sub ImprimeConceptos22(nodoConceptos As IXMLDOMNode)
    Dim iNodos As Integer
    Dim iConceptos As Integer
    Dim nodoCon As IXMLDOMNode
    
    Dim lCol1 As Long
    Dim lCol2 As Long
    Dim lCol3 As Long
    Dim lCol4 As Long
    Dim lCol5 As Long
    
    Dim lY As Long
    Dim lYNext As Long
    
    Dim sCadPrint As String
    
    lCol1 = 0
    lCol2 = 300
    lCol3 = 2000
    lCol4 = 4000
    lCol5 = 5000
    
    Printer.Font = "Courier New"
    Printer.FontSize = 6
    
    Printer.Print String((Printer.Width / Printer.TextWidth("-")), "-")
    
    Printer.CurrentX = lCol1
    Printer.Print "Concepto"
    
    Printer.CurrentX = lCol1
    Printer.Print "Unidad"
    
    Printer.CurrentX = lCol1
    Printer.Print "Cant.";
    
    sCadPrint = "P.Unitario"
    Printer.CurrentX = lCol3 - Printer.TextWidth(sCadPrint)
    Printer.Print sCadPrint;
    
    
    sCadPrint = "Importe"
    Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
    Printer.Print sCadPrint
    
    
    Printer.Print String((Printer.Width / Printer.TextWidth("-")), "-")
    
    
    For Each nodoCon In nodoConceptos.childNodes
        
        lY = Printer.CurrentY
        
        
        'Descripcion
        Printer.CurrentX = lCol1
        ImprimeTexto Trim(nodoCon.Attributes(1).Text), 0, lCol1, Printer.Width
        lYNext = Printer.CurrentY
        
        'Unidad
        Printer.CurrentX = lCol1
        ImprimeTexto Trim(nodoCon.Attributes(5).Text), 0, lCol1, Printer.Width
        lYNext = Printer.CurrentY
        
        
        'Cantidad
        Printer.CurrentX = lCol1
        'Printer.CurrentY = lY
        'Printer.Print
        Printer.Print Format(nodoCon.Attributes(3).Text, "00");
        
        'Precio unitario
        sCadPrint = Format(nodoCon.Attributes(2).Text, "$#,0.00")
        Printer.CurrentX = lCol3 - Printer.TextWidth(sCadPrint)
        'Printer.CurrentY = lY
        Printer.Print sCadPrint;
        
        
        sCadPrint = Format(nodoCon.Attributes(4).Text, "$#,0.00")
        'Printer.CurrentY = lY
        Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
        Printer.Print sCadPrint
        
        'Printer.CurrentY = lYNext
        
    Next
    
    
End Sub

'---------------------------------------------------------
Private Sub ImprimeImpuestos(nodoImpuestos As IXMLDOMNode)
    Dim iNodos As Integer
    Dim nodoImp As IXMLDOMNode
    Dim nodo As IXMLDOMNode
    
    Dim lCol4 As Long
    Dim lCol5 As Long
    Dim sCadPrint As String
    
    lCol4 = 4000
    lCol5 = 5000
    
    For Each nodoImp In nodoImpuestos.childNodes
        For Each nodo In nodoImp.childNodes
            
            'Impuesto
            sCadPrint = nodo.Attributes(0).Text
            
            
            'Porcentaje
            sCadPrint = sCadPrint & " " & nodo.Attributes(1).Text & "%: "
            
            'Importe
            sCadPrint = sCadPrint & Format(nodo.Attributes(2).Text, "$#,0.00")
            Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
            Printer.Print sCadPrint
        Next
    Next
End Sub
'----------------------------------------------------------------------------------
Private Sub ImprimeImpuestos22(nodoImpuestos As IXMLDOMNode)
    Dim iNodos As Integer
    Dim nodoImp As IXMLDOMNode
    Dim nodo As IXMLDOMNode
    
    Dim lCol4 As Long
    Dim lCol5 As Long
    Dim sCadPrint As String
    
    lCol4 = 4000
    lCol5 = 5000
    
    For Each nodoImp In nodoImpuestos.childNodes
        For Each nodo In nodoImp.childNodes
            
            'Impuesto
            sCadPrint = nodo.Attributes(0).Text
            
            
            'Porcentaje
            sCadPrint = sCadPrint & " " & nodo.Attributes(2).Text & "%: "
            
            'Importe
            sCadPrint = sCadPrint & Format(nodo.Attributes(1).Text, "$#,0.00")
            Printer.CurrentX = lCol4 - Printer.TextWidth(sCadPrint)
            Printer.Print sCadPrint
        Next
    Next
End Sub

'----------------------------------------------------------------------------------

Private Sub MultipleLinea(sCadenaOri)

    Dim lCarAct  As Long
    Dim lCarActTot As Long
    Dim sCadenaResta As String
    Dim sCadenaImprime As String
    Dim lAncho As Long
    
    lAncho = Int(Printer.Width * 0.8)
    
    
    If Printer.TextWidth(sCadenaOri) <= lAncho Then
        Printer.Print sCadenaOri
        Exit Sub
    End If
    
    sCadenaResta = sCadenaOri
    sCadenaImprime = ""
    
    lCarAct = 1
    lCarActTot = 0
    
    Do While lCarActTot <= Len(sCadenaOri)
        sCadenaImprime = sCadenaImprime & Mid$(sCadenaResta, lCarAct, 1)
        Select Case Printer.TextWidth(sCadenaImprime)
            Case Is < lAncho
                lCarAct = lCarAct + 1
                lCarActTot = lCarActTot + 1
            Case Is >= lAncho
                Printer.Print sCadenaImprime
                sCadenaResta = Mid$(sCadenaResta, lCarAct + 1)
                sCadenaImprime = vbNullString
                lCarAct = 1
        End Select
    Loop
    
    If sCadenaImprime <> vbNullString Then
        Printer.Print sCadenaImprime
    End If
    
    

End Sub

Private Sub ImprimeComprobanteCajero()
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

    Dim sObservaciones As String

    Dim iTipoFactura As Integer 'Tipo de factura 0 = normal, 1 = Varios

    Dim sCadFormaPago As String

    '04/06/10
    Dim lDesglose As Boolean 'True cuando hay que desglosar la factura

    Dim sStrQry As String

    lMargenDer = Printer.Width - 1000

    iRenMax = 11


    iTipoFactura = GetTipoFactura(lNumeroInicial)
    lDesglose = True


    If iTipoFactura = 0 Then
        Do While Not adorcsFormaPago.EOF
            sCadFormaPago = sCadFormaPago & adorcsFormaPago!Descripcion & " " & adorcsFormaPago!OpcionPago & " " & adorcsFormaPago!Referencia & ","
            adorcsFormaPago.MoveNext
        Loop
        If sCadFormaPago <> vbNullString Then
            sCadFormaPago = Left$(sCadFormaPago, Len(sCadFormaPago) - 1)
        End If
    Else
        sCadFormaPago = "FORMAS DE PAGO VARIAS"
    End If



    If iTipoFactura = 0 Then
        #If SqlServer_ Then
            sStrQry = "SELECT FACTURAS.NumeroFactura, CONVERT(varchar, FACTURAS.FolioCFD) + CONVERT(varchar, FACTURAS.SerieCFD) AS Folio, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Cantidad, FACTURAS_DETALLE.Importe, FACTURAS_DETALLE.Intereses, FACTURAS_DETALLE.Descuento, FACTURAS_DETALLE.Iva, FACTURAS_DETALLE.IvaIntereses, FACTURAS_DETALLE.IvaDescuento, USUARIOS_CLUB.Inscripcion, ISNULL(FACTURAS.TipoPersona,'F') AS TipoPersona"
            sStrQry = sStrQry & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) LEFT JOIN USUARIOS_CLUB ON FACTURAS.IdTitular=USUARIOS_CLUB.IdMember"
            sStrQry = sStrQry & " WHERE FACTURAS.NumeroFactura=" & lNumeroInicial
        #Else
            sStrQry = "SELECT FACTURAS.NumeroFactura,   FACTURAS.FolioCFD & FACTURAS.SerieCFD AS Folio, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS_DETALLE.Periodo, FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Cantidad, FACTURAS_DETALLE.Importe, FACTURAS_DETALLE.Intereses, FACTURAS_DETALLE.Descuento, FACTURAS_DETALLE.Iva, FACTURAS_DETALLE.IvaIntereses, FACTURAS_DETALLE.IvaDescuento, USUARIOS_CLUB.Inscripcion, iif(isNull(FACTURAS.TipoPersona),'F',FACTURAS.TipoPersona) AS TipoPersona"
            sStrQry = sStrQry & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) LEFT JOIN USUARIOS_CLUB ON FACTURAS.IdTitular=USUARIOS_CLUB.IdMember"
            sStrQry = sStrQry & " WHERE FACTURAS.NumeroFactura=" & lNumeroInicial
        #End If
    Else
        #If SqlServer_ Then
            sStrQry = "SELECT FACTURAS.NumeroFactura, CONVERT(varchar, FACTURAS.FolioCFD) + CONVERT(varchar, FACTURAS.SerieCFD) As FolioCFD, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS.FechaFactura AS Periodo, 'VARIOS' AS Concepto, 1 AS Cantidad, SUM(FACTURAS_DETALLE.Importe*FACTURAS_DETALLE.Cantidad) AS Importe, Sum(FACTURAS_DETALLE.Intereses) AS Intereses, Sum(FACTURAS_DETALLE.Descuento) AS Descuento, Sum(FACTURAS_DETALLE.Iva) AS Iva, Sum(FACTURAS_DETALLE.IvaIntereses) AS IvaIntereses, Sum(FACTURAS_DETALLE.IvaDescuento) AS IvaDescuento, USUARIOS_CLUB.Inscripcion, ISNULL(FACTURAS.TipoPersona,'F') AS TipoPersona"
            sStrQry = sStrQry & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) INNER JOIN USUARIOS_CLUB ON FACTURAS.IdTitular = USUARIOS_CLUB.IdMember"
            sStrQry = sStrQry & " GROUP BY FACTURAS.NumeroFactura, CONVERT(varchar, FACTURAS.FolioCFD) + CONVERT(varchar, FACTURAS.SerieCFD), FACTURAS.FechaFactura, FACTURAS.NoFamilia, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total,  'VARIOS', 1,  USUARIOS_CLUB.Inscripcion, ISNULL(FACTURAS.TipoPersona,'F')"
            sStrQry = sStrQry & " HAVING FACTURAS.NumeroFactura=" & lNumeroInicial
        #Else
            sStrQry = "SELECT FACTURAS.NumeroFactura, FACTURAS.Folio & FACTURAS.SerieCFD As FolioCFD, FACTURAS.NoFamilia, FACTURAS.FechaFactura, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total, FACTURAS.FechaFactura AS Periodo, 'VARIOS' AS Concepto, 1 AS Cantidad, SUM(FACTURAS_DETALLE.Importe*FACTURAS_DETALLE.Cantidad) AS Importe, Sum(FACTURAS_DETALLE.Intereses) AS Intereses, Sum(FACTURAS_DETALLE.Descuento) AS Descuento, Sum(FACTURAS_DETALLE.Iva) AS Iva, Sum(FACTURAS_DETALLE.IvaIntereses) AS IvaIntereses, Sum(FACTURAS_DETALLE.IvaDescuento) AS IvaDescuento, USUARIOS_CLUB.Inscripcion, iif(isNull(FACTURAS.TipoPersona),'F',FACTURAS.TipoPersona) AS TipoPersona"
            sStrQry = sStrQry & " FROM (FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura) INNER JOIN USUARIOS_CLUB ON FACTURAS.IdTitular = USUARIOS_CLUB.IdMember"
            sStrQry = sStrQry & " GROUP BY FACTURAS.NumeroFactura, FACTURAS.Folio & FACTURAS.Serie, FACTURAS.FechaFactura, FACTURAS.NoFamilia, FACTURAS.NombreFactura, FACTURAS.CalleFactura, FACTURAS.ColoniaFactura, FACTURAS.DelFactura, FACTURAS.CiudadFactura, FACTURAS.EstadoFactura, FACTURAS.CodPos, FACTURAS.RFC, FACTURAS.Tel1, FACTURAS.Observaciones, FACTURAS.Total,  'VARIOS', 1,  USUARIOS_CLUB.Inscripcion, iif(isNull(FACTURAS.TipoPersona),'F',FACTURAS.TipoPersona)"
            sStrQry = sStrQry & " HAVING FACTURAS.NumeroFactura=" & lNumeroInicial
        #End If
    End If

    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer





    dImporte = 0
    dIva = 0
    dSubtotal = 0

    iRen = 0





    adorcs.Open sStrQry, Conn, adOpenForwardOnly, adLockReadOnly

   

    sFolioFactura = IIf(IsNull(adorcs!Folio), vbNullString, adorcs!Folio)
    dFechaFactura = adorcs!FechaFactura
    sObservaciones = adorcs!Observaciones

    Printer.FontName = "Arial Narrow"



    Printer.FontSize = 8


    Printer.FontBold = True
    Printer.Print adorcs!Folio;
    Printer.FontBold = False
    Printer.Print " (" & adorcs!NumeroFactura & ")";


    Printer.CurrentX = 10000
    Printer.FontBold = True
    Printer.Print vbNullString
    Printer.FontBold = False

    lCad = Format(adorcs!FechaFactura, "Short Date")
    Printer.Print lCad



    Printer.Print adorcs!NoFamilia & " " & adorcs!NombreFactura


    Do Until adorcs.EOF
        'Printer.CurrentX = 4000
        'Printer.Print adorcs!Periodo;

        'Printer.CurrentX = 5000
        'Printer.Print adorcs!Concepto;


        If lDesglose Then
            dImporte = ((adorcs!Importe * adorcs!Cantidad) + adorcs!Intereses - adorcs!Descuento) - adorcs!Iva - adorcs!IvaIntereses + adorcs!IvaDescuento
            dIva = dIva + adorcs!Iva + adorcs!IvaIntereses - adorcs!IvaDescuento
            dSubtotal = dSubtotal + dImporte
        Else
            dImporte = ((adorcs!Importe * adorcs!Cantidad) + adorcs!Intereses - adorcs!Descuento)
            dIva = 0
            dSubtotal = dSubtotal + dImporte
        End If

        'Printer.CurrentX = 8000
        'lCad = Format(dImporte, "$#,0.00")
        'Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))

        'Printer.Print lCad


        adorcs.MoveNext
        iRen = iRen + 1
    Loop


    'Printer.CurrentX = 4000
    'Printer.CurrentY = lYOffset + 5300
    'Printer.Print sObservaciones

    lCad = "1"

    'Printer.CurrentY = lYOffset + 5400


    dTotal = dSubtotal + dIva


    If (lDesglose) Then
        lCad = Format(dSubtotal, "$#,0.00")
        'Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
        Printer.Print lCad
    End If

    'Printer.CurrentY = lYOffset + 5740


    If (lDesglose) Then

        lCad = ObtieneParametro("IVA_GENERAL") & "%"
        'Printer.CurrentX = (lMargenDer - 2300 - Printer.TextWidth(lCad))
        Printer.Print lCad;


        lCad = Format(dIva, "$#,0.00")
        'Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
        Printer.Print lCad
    End If

    'Printer.CurrentY = lYOffset + 6080

    lCad = Format(dTotal, "$#,0.00")
    'Printer.CurrentX = (lMargenDer - 1000 - Printer.TextWidth(lCad))
    Printer.Print lCad



    'dDecimal = (dTotal - Int(dTotal)) * 100
    'lCad = UCase(Num2Txt(Int(dTotal))) & " PESOS "
    'lCad = lCad & Format(dDecimal, "00") & "/100 M.N."
    'Printer.CurrentX = 4000
    'lCad = "(" & lCad & ")"
    'Printer.CurrentY = lYOffset + 5900
    'Printer.Print lCad;


    'Printer.CurrentX = 8000
    'Printer.CurrentY = lYOffset + 7250
     MultipleLinea sCadFormaPago




    adorcs.Close

    Printer.EndDoc
    Set adorcs = Nothing
    Unload Me


End Sub

