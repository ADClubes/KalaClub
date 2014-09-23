VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCtrlCUpon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Entrega de Cupones"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   9840
      TabIndex        =   12
      Top             =   480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cupones"
      Height          =   3255
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   10935
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir  Cupones"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar Cupones"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgCupones 
         Height          =   1815
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   10455
         _Version        =   196616
         DataMode        =   2
         Col.Count       =   7
         AllowUpdate     =   0   'False
         AllowColumnMoving=   0
         AllowColumnSwapping=   0
         SelectTypeCol   =   0
         SelectTypeRow   =   0
         RowHeight       =   423
         Columns.Count   =   7
         Columns(0).Width=   1270
         Columns(0).Caption=   "Folio"
         Columns(0).Name =   "Folio"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1535
         Columns(1).Caption=   "Consecutivo"
         Columns(1).Name =   "Consecutivo"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   2514
         Columns(2).Caption=   "Fecha"
         Columns(2).Name =   "Fecha"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   2170
         Columns(3).Caption=   "Hora"
         Columns(3).Name =   "Hora"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         Columns(4).Width=   3200
         Columns(4).Caption=   "Concepto"
         Columns(4).Name =   "Concepto"
         Columns(4).DataField=   "Column 4"
         Columns(4).DataType=   8
         Columns(4).FieldLen=   256
         Columns(5).Width=   3200
         Columns(5).Caption=   "Vigencia"
         Columns(5).Name =   "Vigencia"
         Columns(5).DataField=   "Column 5"
         Columns(5).DataType=   8
         Columns(5).FieldLen=   256
         Columns(6).Width=   1429
         Columns(6).Caption=   "Impreso"
         Columns(6).Name =   "Impreso"
         Columns(6).DataField=   "Column 6"
         Columns(6).DataType=   8
         Columns(6).FieldLen=   256
         _ExtentX        =   18441
         _ExtentY        =   3201
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
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   10935
      Begin VB.TextBox txtFechaPago 
         Height          =   375
         Left            =   6240
         TabIndex        =   13
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtStatus 
         Height          =   375
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox txtNombreTit 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label5 
         Caption         =   "Pagado Hasta"
         Height          =   255
         Left            =   6360
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Status"
         Height          =   255
         Left            =   8280
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre titular inscripción"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtInscripcion 
      Height          =   375
      Left            =   240
      MaxLength       =   6
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Status"
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Inscripción"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmCtrlCUpon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBuscar_Click()
    
    Dim bTieneCupones As Boolean
    Dim iStatusCupones As Integer
    
    If Me.txtInscripcion.Text = vbNullString Then
        MsgBox "Indicar número de inscripción", vbExclamation, "Verifique"
        Me.txtInscripcion.SetFocus
        Exit Sub
    End If
    
    'Checa que la inscripción tenga cupones
    'Si no tiene cupones asignados
    bTieneCupones = BuscaInsc(CLng(Me.txtInscripcion.Text))
    If Not bTieneCupones Then
        MsgBox "Esta inscripción no tiene cupones asignados", vbExclamation, "Error"
        Me.txtInscripcion.SetFocus
        Exit Sub
    End If
    
    'Trae el dato de la inscripcion
    Me.txtNombreTit.Text = DatosInscripcion(CLng(Me.txtInscripcion.Text))
    
    'Trae la última fecha de pago
    Me.txtFechaPago.Text = DatosPago(CLng(Me.txtInscripcion.Text))
    
    'Checa el status de los cupones
    iStatusCupones = StatusCupones(CLng(Me.txtInscripcion.Text))
    
    Select Case iStatusCupones
        Case 0
            Me.txtStatus.Text = "NO ASIGNADOS"
        Case 1
            Me.txtStatus.Text = "ASIGNADOS"
    End Select
    
    If CDate(Me.txtFechaPago.Text) < Date Then
        MsgBox "Verifique que el usuario no tenga adeudo de mantenimiento", vbExclamation, "Verifique"
    End If
    
    
    'Si no se han creado los cupones
    If iStatusCupones = 0 And bTieneCupones Then
        Me.cmdGenerar.Enabled = True
        Me.cmdGenerar.SetFocus
        Me.txtInscripcion.Enabled = False
        Me.cmdBuscar.Enabled = False
        Exit Sub
    End If
    
    If iStatusCupones = 1 And bTieneCupones Then
        DisplayCupones CLng(Me.txtInscripcion.Text)
        Me.cmdImprimir.Enabled = True
        Me.txtInscripcion.Enabled = False
        Me.cmdBuscar.Enabled = False
    End If
    
    
    
End Sub

Private Sub cmdCancelar_Click()
    Me.ssdbgCupones.RemoveAll
    Me.cmdGenerar.Enabled = False
    Me.cmdImprimir.Enabled = False
    Me.txtNombreTit.Text = vbNullString
    Me.txtFechaPago.Text = vbNullString
    Me.txtStatus.Text = vbNullString
    
    Me.txtInscripcion.Enabled = True
    Me.cmdBuscar.Enabled = True
    
    
    Me.txtInscripcion.SetFocus
End Sub

Private Sub cmdGenerar_Click()
    
    Dim sInter As String
    Dim lDiasVigencia As Long
    Dim iCupones As Integer
    
    Me.cmdGenerar.Enabled = False
    
    sInter = ObtieneParametro("DIAS_VIGENCIA_CUPON_REGALO")
    
    If sInter = vbNullString Then
        lDiasVigencia = 30
    Else
        lDiasVigencia = Val(sInter)
    End If
    
    sInter = ObtieneParametro("NUMERO_CUPON_REGALO")
    
    If sInter = vbNullString Then
        iCupones = 1
    Else
        iCupones = 3
    End If
    
    If GeneraCupones(CLng(Me.txtInscripcion.Text), iCupones, lDiasVigencia, "PASE DE INVITADO POR UN DIA") Then
        MsgBox "Cupones Creados", vbInformation, "Correcto"
        DisplayCupones CLng(Me.txtInscripcion.Text)
        If Me.ssdbgCupones.Rows > 0 Then
            Me.cmdImprimir.Enabled = True
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    
    Dim sLogoFile As String 'Path donde se buscará el archivo con el logo
    Dim picLogo As Picture 'Almacena el logo que será impreso
    
    Dim sImpresoraActual As String
    Dim sImpresoraCupon As String
    
    Dim adoRcsCupones As ADODB.Recordset
    
    Dim sNombreClub As String
    
    Dim iRespuesta As Integer
    
    Dim iTotCupones As Integer
    Dim lNumInsc As Long
    
    On Error GoTo ERROR_SUB
    
    
    
    
    Screen.MousePointer = vbHourglass
    
    sNombreClub = ObtieneParametro("NOMBRE DEL CLUB")
    sLogoFile = ObtieneParametro("GRAFICO CUPON") '"d:\kalaclub\recursos\logo_sportium_bn.jpg"
    sImpresoraCupon = ObtieneParametro("IMPRESORA CUPON") '"Apos Premium"
    
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
    
    
    strSQL = "SELECT Cuponesregalo.NoInscripcion, CuponesRegalo.Folio, CuponesRegalo.TotalCupones, CuponesRegaloDetalle.Consecutivo, CuponesRegalo.FechaCreacion, CuponesRegalo.HoraCreacion, CuponesRegaloDetalle.Concepto, CuponesRegaloDetalle.Vigencia, CuponesRegalo.Impresiones"
    strSQL = strSQL & " FROM CuponesRegalo INNER JOIN CuponesRegaloDetalle ON CuponesRegalo.Folio = CuponesRegaloDetalle.Folio"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((CuponesRegalo.Folio)=" & Me.ssdbgCupones.Columns("Folio").Value & ")"
    strSQL = strSQL & ")"
    strSQL = strSQL & "ORDER BY CuponesRegaloDetalle.Consecutivo"
    
    
    Set adoRcsCupones = New ADODB.Recordset
    adoRcsCupones.ActiveConnection = Conn
    adoRcsCupones.CursorLocation = adUseServer
    adoRcsCupones.CursorType = adOpenForwardOnly
    adoRcsCupones.LockType = adLockReadOnly
    
    adoRcsCupones.Open strSQL
    
    
    If sLogoFile <> vbNullString Then
        If Dir(sLogoFile) <> vbNullString Then
            Set picLogo = LoadPicture(sLogoFile)
        Else
            Set picLogo = LoadPicture("")
        End If
    Else
        Set picLogo = LoadPicture("")
    End If
    
    
    
    Do Until adoRcsCupones.EOF
    
        iTotCupones = adoRcsCupones!TotalCupones
        lNumInsc = adoRcsCupones!NoInscripcion
        
        On Error Resume Next
        Printer.PaintPicture picLogo, 600, 0, 3000, 1000
        On Error GoTo ERROR_SUB
        
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        Printer.Print
        
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.Print sNombreClub
        Printer.Print
        Printer.FontSize = 8
        Printer.FontBold = False
        Printer.Print "INSCRIPCION: " & adoRcsCupones!NoInscripcion
        Printer.Print
        Printer.FontBold = True
        Printer.Print "CUPON DE CORTESIA"
        Printer.Print
        Printer.FontBold = False
        Printer.Print "VALIDO POR UN " & adoRcsCupones!Concepto
        Printer.Print
            
        Printer.Print "CUPON " & adoRcsCupones!Consecutivo; " DE "; adoRcsCupones!TotalCupones
        Printer.Print
        Printer.Print "EMITIDO: " & Format(adoRcsCupones!FechaCreacion, "dd/mmm/yy") & " VALIDO HASTA: " & Format(adoRcsCupones!Vigencia, "dd/mmm/yy")
        Printer.Print
        Printer.Print "FOLIO #: "; adoRcsCupones!Folio
        Printer.Print
        Printer.Print
        Printer.Print "EL USO DEL PRESENTE CUPON ESTA SUJETO"
        Printer.Print "AL REGLAMENTO SPORTIUM"
        Printer.Print "PUEDEN APLICAR RESTRICCIONES"
        
        
        
        If adoRcsCupones!Impresiones > 0 Then
            Printer.FontSize = 6
            Printer.Print
            Printer.Print "ESTA ES UNA REIMPRESION DEL CUPON ORIGINAL"
            Printer.Print "PARA QUE SEA VALIDO DEBE ESTAR FIRMADO POR"
            Printer.Print "EL CAJERO"
            Printer.Print
            Printer.Print
            Printer.Print
            Printer.Print "___________________________________________"
            Printer.Print
            Printer.Print "REIMPRESION # " & adoRcsCupones!Impresiones
        End If
        
        
        
        
        
        Printer.EndDoc
        
        
        
        
        
        adoRcsCupones.MoveNext
    Loop
    
    'Imprime el acuse de recibo
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.Print sNombreClub
    Printer.Print
    Printer.Print
    Printer.FontSize = 8
    Printer.Print "INSCRIPCION: " & lNumInsc
    Printer.Print
    Printer.Print "RECIBI " & iTotCupones & " CUPONES"
    Printer.Print
    Printer.Print Now
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print
    Printer.Print "___________________________________________"
    Printer.Print
    Printer.Print "                    Nombre y Firma"
    Printer.Print
    Printer.Print
    
    Printer.EndDoc
    
    
    adoRcsCupones.Close
    Set adoRcsCupones = Nothing
    
    If ActualContImp(Me.ssdbgCupones.Columns("Folio").Value) Then
    End If
    
    Set picLogo = Nothing
    
    
    sImpresoraActual = SelectPrinter(sImpresoraActual)
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Cupones Impresos", vbInformation, "Correcto"
    
    
    Me.cmdCancelar.Value = True
    
    
    
    
    Exit Sub
    
ERROR_SUB:
    
    'Restablece la impresora
    sImpresoraActual = SelectPrinter(sImpresoraActual)
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Ocurrio un error", vbCritical, "Error"
    
    'MsjError




End Sub

Private Sub Form_Load()
    If Lee_Ini() <> 0 Then
        Exit Sub
    End If
    
    If Not Connection_DB() Then
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndConn_DB
End Sub

Private Sub txtInscripcion_GotFocus()
    Me.txtInscripcion.SelStart = 0
    Me.txtInscripcion.SelLength = Len(Me.txtInscripcion.Text)
End Sub

Private Sub txtInscripcion_KeyPress(KeyAscii As Integer)
      Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub
