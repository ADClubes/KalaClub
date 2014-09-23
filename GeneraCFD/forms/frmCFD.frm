VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCFD 
   Caption         =   "Generar CFD"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3480
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtSerieFactura 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtSOAP 
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   4095
   End
   Begin VB.CommandButton cmdGenera 
      Caption         =   "Generar CFD"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtFolioCFD 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtNoFactura 
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblCtrl 
      Alignment       =   2  'Center
      Caption         =   "Serie"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblCtrl 
      Alignment       =   2  'Center
      Caption         =   "Folio"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmCFD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
      "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As _
      String, ByVal lpszFile As String, ByVal lpszParams As String, _
      ByVal lpszDir As String, ByVal FsShowCmd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Const SW_SHOWNORMAL = 1

Const SE_ERR_FNF = 2&
Const SE_ERR_PNF = 3&
Const SE_ERR_ACCESSDENIED = 5&
Const SE_ERR_OOM = 8&
Const SE_ERR_DLLNOTFOUND = 32&
Const SE_ERR_SHARE = 26&
Const SE_ERR_ASSOCINCOMPLETE = 27&
Const SE_ERR_DDETIMEOUT = 28&
Const SE_ERR_DDEFAIL = 29&
Const SE_ERR_DDEBUSY = 30&
Const SE_ERR_NOASSOC = 31&
Const ERROR_BAD_FORMAT = 11&






Private Const sSOAPGeneraFactura = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
    "<soap:Body>" & _
    "<GeneraFactura xmlns=""http://tempuri.org/"">" & _
        "<IdFacturaKala>int</IdFacturaKala>" & _
        "<serie>string</serie>" & _
        "</GeneraFactura>" & _
    "</soap:Body>" & _
    "</soap:Envelope>"
    
Private Const sURLWS = "http://172.16.2.5/WSFacturadigital/FacturaElectronica.asmx"






Private Sub cmdAbrir_Click()
    Dim r As Long, msg As String
    
    Dim sFileName As String
    
    sFileName = "\\172.16.2.5\Pruebas\OCS051012B29" & Me.txtSerieFactura.Text & Me.txtFolioCFD & ".pdf"
    
    r = StartDoc(sFileName)
    
    If r <= 32 Then
        'There was an error
        Select Case r
            Case SE_ERR_FNF
                      msg = "File not found"
            Case SE_ERR_PNF
                      msg = "Path not found"
            Case SE_ERR_ACCESSDENIED
                      msg = "Access denied"
            Case SE_ERR_OOM
                      msg = "Out of memory"
            Case SE_ERR_DLLNOTFOUND
                      msg = "DLL not found"
            Case SE_ERR_SHARE
                      msg = "A sharing violation occurred"
            Case SE_ERR_ASSOCINCOMPLETE
                      msg = "Incomplete or invalid file association"
            Case SE_ERR_DDETIMEOUT
                      msg = "DDE Time out"
            Case SE_ERR_DDEFAIL
                      msg = "DDE transaction failed"
            Case SE_ERR_DDEBUSY
                      msg = "DDE busy"
            Case SE_ERR_NOASSOC
                      msg = "No association for file extension"
            Case ERROR_BAD_FORMAT
                      msg = "Invalid EXE file or error in EXE image"
            Case Else
                      msg = "Unknown error"
        End Select
              MsgBox msg
    End If
    
End Sub

Private Sub cmdGenera_Click()
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    
    Me.cmdGenera.Enabled = False
    Me.cmdAbrir.Enabled = False
    Me.txtFolioCFD.Text = vbNullString
    
    
    parser.loadXML sSOAPGeneraFactura
    
    'Asigna los valores para el servicio
    'Folio
    parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraFactura/IdFacturaKala").Text = Trim(Me.txtNoFactura)
    
    'Serie
    parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraFactura/serie").Text = Trim(Me.txtSerieFactura)
    
    
    Me.txtSOAP = parser.xml
    DoEvents
    
    'Inet1.Execute sURLWS, "POST", parser.xml, "Content-Type: text/xml; charset=utf-8" & vbCrLf & "SOAPAction: http://tempuri.org/GeneraFactura"
    
    
    enviarComando parser.xml, "http://tempuri.org/GeneraFactura"

    
    
    
    
End Sub


Private Sub Inet1_StateChanged(ByVal State As Integer)
    '
    If (State = icResponseCompleted) Then ' icResponseCompleted = 12
        Dim s As String
        '
        ' Leer los datos devueltos por el servidor
        s = Inet1.GetChunk(4096)
        Me.txtSOAP.Text = s
        '
        ' Poner los datos en el analizador de XML
        Dim parser As DOMDocument
        Set parser = New DOMDocument
        parser.loadXML s
        '
        On Error Resume Next
        '
        Me.txtFolioCFD.Text = parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraFacturaResponse/GeneraFacturaResult").Text
        '
        If Err.Number > 0 Then
            Me.txtSOAP.SetFocus
        End If
        
    Me.txtNoFactura.Text = Val(Me.txtNoFactura.Text) + 1
    Me.cmdGenera.Enabled = True
    
    End If
End Sub

Private Sub enviarComando(ByVal sXml As String, ByVal sSoapAction As String)
    ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    Dim oHttReq As XMLHTTP60
    Set oHttReq = New XMLHTTP60
    
    
    
    '
    ' Enviar el comando de forma síncrona (se espera a que se reciba la respuesta)
    oHttReq.open "POST", sURLWS, False
    ' las cabeceras a enviar al servicio Web
    ' (no incluir los dos puntos en el nombre de la cabecera)
    oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    oHttReq.setRequestHeader "SOAPAction", sSoapAction
    ' enviar el comando
    oHttReq.send sXml
    DoEvents
    
    Do While True
        DoEvents
        If oHttReq.readyState = 4 Then Exit Do
    Loop

    
    '
    ' este será el texto recibido del servicio Web
    procesarRespuesta oHttReq.responseText
    '
    Me.txtNoFactura.Text = Val(Me.txtNoFactura.Text) + 1
    Me.cmdGenera.Enabled = True
    Me.cmdAbrir.Enabled = True
    
End Sub

Private Sub procesarRespuesta(ByVal s As String)
    ' procesar la respuesta recibida del servicio Web
    Me.txtSOAP.Text = s
    '
    ' Poner los datos en el analizador de XML
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    parser.loadXML s
    '
    On Error Resume Next
    '
    Me.txtFolioCFD.Text = parser.selectSingleNode("/soap:Envelope/soap:Body/GeneraFacturaResponse/GeneraFacturaResult").Text
    '
    If Err.Number > 0 Then
        Me.txtSOAP.SetFocus
    End If
End Sub


Function StartDoc(DocName As String) As Long
    Dim Scr_hDC As Long
    Scr_hDC = GetDesktopWindow()
    StartDoc = ShellExecute(Scr_hDC, "Open", DocName, _
    "", "C:\", SW_SHOWNORMAL)
End Function

