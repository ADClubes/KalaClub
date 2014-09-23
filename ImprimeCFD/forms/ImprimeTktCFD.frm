VERSION 5.00
Begin VB.Form ImprimeTktCFD 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "90"
      Top             =   960
      Width           =   495
   End
   Begin VB.ComboBox cmbImpresora 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprime"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "ImprimeTktCFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim parser As DOMDocument
Dim oXls As DOMDocument
Dim sCurPrn As String


Private Sub cmbImpresora_Click()
    sCurPrn = SelectImpresora(Me.cmbImpresora.Text)
End Sub

Private Sub Command1_Click()
    
    Dim comprobante As IXMLDOMNode
    Dim emisor As IXMLDOMNode
    Dim receptor As IXMLDOMNode
    Dim conceptos As IXMLDOMNode
    Dim impuestos As IXMLDOMNode
    Dim addenda As IXMLDOMNode
    
    
    
    
    Set comprobante = parser.selectSingleNode("Comprobante")
    Set emisor = parser.selectSingleNode("Comprobante/Emisor")
    Set receptor = parser.selectSingleNode("Comprobante/Receptor")
    Set conceptos = parser.selectSingleNode("Comprobante/Conceptos")
    Set impuestos = parser.selectSingleNode("Comprobante/Impuestos")
    Set addenda = parser.selectSingleNode("Comprobante/Addenda")
    
    Printer.FontSize = 8
    
    
    'Datos CFD
    Printer.Print "Serie: " & comprobante.Attributes(1).Text
    Printer.Print "Folio: " & comprobante.Attributes(2).Text
    Printer.Print "Fecha: " & comprobante.Attributes(3).Text
    Printer.Print "No. Aprobacion: " & comprobante.Attributes(4).Text
    Printer.Print "Año de aprobación: " & comprobante.Attributes(5).Text
    Printer.Print "Certificado: " & comprobante.Attributes(11).Text
    Printer.Print
    
    'Emisor
    Printer.FontBold = True
    Printer.Print "Emisor"
    Printer.FontBold = False
    Printer.Print emisor.Attributes(0).Text
    Printer.Print emisor.Attributes(1).Text
    ImprimeDireccion emisor.selectSingleNode("DomicilioFiscal")
    Printer.Print
    
    'Receptor
    Printer.FontBold = True
    Printer.Print "Cliente"
    Printer.FontBold = False
    Printer.Print receptor.Attributes(0).Text
    Printer.Print receptor.Attributes(1).Text
    ImprimeDireccion receptor.selectSingleNode("Domicilio")
    Printer.Print
    
    'Conceptos
    Printer.FontSize = 6
    ImprimeConceptos conceptos
    
    'Subttotal
    Printer.Print "Subtotal:" & comprobante.Attributes(7).Text
    
    'Impuestos
    ImprimeImpuestos impuestos
    
    'Total
    Printer.Print "Total: " & comprobante.Attributes(8).Text
    Printer.Print
    
    Printer.FontSize = 6
    Printer.FontBold = True
    Printer.Print "Sello Digital"
    Printer.FontBold = False
    Printer.FontSize = 6
    MultipleLinea comprobante.Attributes(13).Text
    
    
    
    Printer.Print
    
    Printer.FontSize = 6
    Printer.FontBold = True
    Printer.Print "Cadena Original"
    Printer.FontBold = False
    Printer.FontSize = 6
    MultipleLinea parser.transformNode(oXls)
    
    Printer.Print
    MultipleLinea "PAGO EN UNA SOLA EXHIBICION"
    
    Printer.Print
    MultipleLinea "ESTE DOCUMENTO ES UNA IMPRESION DE UN COMPROBANTE FISCAL DIGITAL"
    
    Printer.EndDoc
    
End Sub
Private Sub ImprimeDireccion(nodoDir As IXMLDOMNode)
    Dim iI As Integer
    For iI = 0 To nodoDir.Attributes.length - 1
        Printer.Print nodoDir.Attributes(iI).Text
    Next
End Sub
Private Sub ImprimeConceptos(nodoConceptos As IXMLDOMNode)
    Dim iNodos As Integer
    Dim iConceptos As Integer
    Dim nodoCon As IXMLDOMNode
    For Each nodoCon In nodoConceptos.childNodes
        Printer.Print nodoCon.Attributes(3).Text, nodoCon.Attributes(1).Text, nodoCon.Attributes(2).Text, nodoCon.Attributes(4).Text
    Next
End Sub
Private Sub ImprimeImpuestos(nodoImpuestos As IXMLDOMNode)
    Dim iNodos As Integer
    Dim nodoImp As IXMLDOMNode
    Dim nodo As IXMLDOMNode
    For Each nodoImp In nodoImpuestos.childNodes
        For Each nodo In nodoImp.childNodes
            Printer.Print nodo.Attributes(0).Text, nodo.Attributes(1).Text, nodo.Attributes(2).Text
        Next
    Next
End Sub

Private Sub Form_Load()
    Dim cadena As String
    Set parser = New DOMDocument
    Set oXls = New DOMDocument
    
    parser.async = False
    oXls.async = False
    
    If Not parser.Load("\\172.16.2.1\Aplicaciones\KalaClub\CFD\PRUEBA\OCS051012B29SANCU00014.xml") Then
        MsgBox "Error"
    End If
    
    oXls.Load ("\\172.16.2.1\Aplicaciones\KalaClub\CFD\XSLT\cadenaoriginal_2_0.xslt")
    
    LlenaImpresoras
    
    
End Sub

Private Sub MultipleLinea(sCadenaOri)

    Dim lCarAct  As Long
    Dim lCarActTot As Long
    Dim sCadenaResta As String
    Dim sCadenaImprime As String
    Dim lAncho As Long
    
    lAncho = Int(Printer.Width * Val(Me.Text1.Text) / 100)
    
    
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

Private Sub LlenaImpresoras()
    Dim prn As Printer
    
    For Each prn In Printers
        Me.cmbImpresora.AddItem prn.DeviceName
    Next
    
End Sub

Private Function SelectImpresora(sPrnName As String)
    Dim prn As Printer
    
    SelectImpresora = Printer.DeviceName
    
    For Each prn In Printers
        If prn.DeviceName = sPrnName Then
            Set Printer = prn
        End If
    Next
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim sText As String
    
    sText = SelectImpresora(sCurPrn)
    
End Sub
