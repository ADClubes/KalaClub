VERSION 5.00
Begin VB.Form frmImpCupones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imprime Cupones"
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
   Begin VB.TextBox txtNumFacFinal 
      Enabled         =   0   'False
      Height          =   405
      Left            =   2633
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtNumFacIni 
      Enabled         =   0   'False
      Height          =   405
      Left            =   593
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Default         =   -1  'True
      Height          =   615
      Left            =   2633
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   713
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "# Final"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "# Inicial"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "frmImpCupones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()

    Dim sLogoFile As String 'Path donde se buscará el archivo con el logo
    Dim picLogo As Picture 'Almacena el logo que será impreso
    
    Dim sImpresoraActual As String
    Dim sImpresoraCupon As String
    
    Dim adoRcsCupones As ADODB.Recordset
    
    Dim sNombreClub As String
    
    Dim iRespuesta As Integer
    
    On Error GoTo ERROR_SUB
    
    
    
    
    Screen.MousePointer = vbHourglass
    
    sNombreClub = ObtieneParametro("NOMBRE DEL CLUB")
    sLogoFile = ObtieneParametro("GRAFICO CUPON") '"d:\kalaclub\recursos\logo_sportium_bn.jpg"
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
    
    If Me.Tag = "R" Then
        #If SqlServer_ Then
            strSQL = "SELECT CUPONES.FolioCupon, CUPONES.NumeroCupon, CUPONES.TotalCupones, CUPONES.TipoDocumento, CUPONES.NumeroDocumento, CUPONES.FechaAlta, CUPONES.FechaVigencia, CUPONES.Impresiones, CUPONES.DatosAdicionales, CUPONES.ImporteCupon, INSTRUCTORES.Apellido_Paterno + ' ' + INSTRUCTORES.Apellido_Materno + ' ' + INSTRUCTORES.Nombre AS Instructor, USUARIOS_CLUB.NoFamilia, CONCEPTO_INGRESOS.DescripcionCupon, CONCEPTO_INGRESOS.ObservacionesCupon, CONCEPTO_INGRESOS.RequiereUsuario, CONCEPTO_INGRESOS.RequiereInstructor, CONCEPTO_INGRESOS.ImprimeImporte, USUARIOS_CLUB.A_Paterno + ' ' + USUARIOS_CLUB.A_Materno + ' ' + USUARIOS_CLUB.Nombre AS Usuario"
            strSQL = strSQL & " FROM ((CUPONES INNER JOIN CONCEPTO_INGRESOS ON CUPONES.IdConcepto = CONCEPTO_INGRESOS.IdConcepto) LEFT JOIN INSTRUCTORES ON CUPONES.IdInstructor = INSTRUCTORES.IdInstructor) INNER JOIN USUARIOS_CLUB ON CUPONES.IdMember = USUARIOS_CLUB.IdMember"
            strSQL = strSQL & " WHERE (((CUPONES.TipoDocumento)='R') AND ((CUPONES.NumeroDocumento) Between " & Me.txtNumFacIni.Text & " AND " & Me.txtNumFacFinal.Text & "))"
            strSQL = strSQL & " ORDER BY CUPONES.NumeroDocumento, CUPONES.IdConcepto, CUPONES.NumeroCupon"
        #Else
            strSQL = "SELECT CUPONES.FolioCupon, CUPONES.NumeroCupon, CUPONES.TotalCupones, CUPONES.TipoDocumento, CUPONES.NumeroDocumento, CUPONES.FechaAlta, CUPONES.FechaVigencia, CUPONES.Impresiones, CUPONES.DatosAdicionales, CUPONES.ImporteCupon, INSTRUCTORES.Apellido_Paterno & ' ' & INSTRUCTORES.Apellido_Materno & ' ' & INSTRUCTORES.Nombre AS Instructor, USUARIOS_CLUB.NoFamilia, CONCEPTO_INGRESOS.DescripcionCupon, CONCEPTO_INGRESOS.ObservacionesCupon, CONCEPTO_INGRESOS.RequiereUsuario, CONCEPTO_INGRESOS.RequiereInstructor, CONCEPTO_INGRESOS.ImprimeImporte, USUARIOS_CLUB.A_Paterno & ' ' & USUARIOS_CLUB.A_Materno & ' ' & USUARIOS_CLUB.Nombre AS Usuario"
            strSQL = strSQL & " FROM ((CUPONES INNER JOIN CONCEPTO_INGRESOS ON CUPONES.IdConcepto = CONCEPTO_INGRESOS.IdConcepto) LEFT JOIN INSTRUCTORES ON CUPONES.IdInstructor = INSTRUCTORES.IdInstructor) INNER JOIN USUARIOS_CLUB ON CUPONES.IdMember = USUARIOS_CLUB.IdMember"
            strSQL = strSQL & " WHERE (((CUPONES.TipoDocumento)='R') AND ((CUPONES.NumeroDocumento) Between " & Me.txtNumFacIni.Text & " AND " & Me.txtNumFacFinal.Text & "))"
            strSQL = strSQL & " ORDER BY CUPONES.NumeroDocumento, CUPONES.IdConcepto, CUPONES.NumeroCupon"
        #End If
    Else
        #If SqlServer_ Then
            strSQL = "SELECT CUPONES.FolioCupon, CUPONES.NumeroCupon, CUPONES.TotalCupones, CUPONES.TipoDocumento, CUPONES.NumeroDocumento, CUPONES.FechaAlta, CUPONES.FechaVigencia, CUPONES.Impresiones, CUPONES.DatosAdicionales, CUPONES.ImporteCupon, INSTRUCTORES.Apellido_Paterno + ' ' + INSTRUCTORES.Apellido_Materno + ' ' + INSTRUCTORES.Nombre AS Instructor, USUARIOS_CLUB.NoFamilia, CONCEPTO_INGRESOS.DescripcionCupon, CONCEPTO_INGRESOS.ObservacionesCupon, CONCEPTO_INGRESOS.RequiereUsuario, CONCEPTO_INGRESOS.RequiereInstructor, CONCEPTO_INGRESOS.ImprimeImporte, USUARIOS_CLUB.A_Paterno + ' ' + USUARIOS_CLUB.A_Materno + ' ' + USUARIOS_CLUB.Nombre AS Usuario"
            strSQL = strSQL & " FROM ((CUPONES INNER JOIN CONCEPTO_INGRESOS ON CUPONES.IdConcepto = CONCEPTO_INGRESOS.IdConcepto) LEFT JOIN INSTRUCTORES ON CUPONES.IdInstructor = INSTRUCTORES.IdInstructor) INNER JOIN USUARIOS_CLUB ON CUPONES.IdMember = USUARIOS_CLUB.IdMember"
            strSQL = strSQL & " WHERE (((CUPONES.TipoDocumento)='F') AND ((CUPONES.NumeroDocumento) Between " & Me.txtNumFacIni.Text & " AND " & Me.txtNumFacFinal.Text & "))"
            strSQL = strSQL & " ORDER BY CUPONES.NumeroDocumento, CUPONES.IdConcepto, CUPONES.NumeroCupon"
        #Else
            strSQL = "SELECT CUPONES.FolioCupon, CUPONES.NumeroCupon, CUPONES.TotalCupones, CUPONES.TipoDocumento, CUPONES.NumeroDocumento, CUPONES.FechaAlta, CUPONES.FechaVigencia, CUPONES.Impresiones, CUPONES.DatosAdicionales, CUPONES.ImporteCupon, INSTRUCTORES.Apellido_Paterno & ' ' & INSTRUCTORES.Apellido_Materno & ' ' & INSTRUCTORES.Nombre AS Instructor, USUARIOS_CLUB.NoFamilia, CONCEPTO_INGRESOS.DescripcionCupon, CONCEPTO_INGRESOS.ObservacionesCupon, CONCEPTO_INGRESOS.RequiereUsuario, CONCEPTO_INGRESOS.RequiereInstructor, CONCEPTO_INGRESOS.ImprimeImporte, USUARIOS_CLUB.A_Paterno & ' ' & USUARIOS_CLUB.A_Materno & ' ' & USUARIOS_CLUB.Nombre AS Usuario"
            strSQL = strSQL & " FROM ((CUPONES INNER JOIN CONCEPTO_INGRESOS ON CUPONES.IdConcepto = CONCEPTO_INGRESOS.IdConcepto) LEFT JOIN INSTRUCTORES ON CUPONES.IdInstructor = INSTRUCTORES.IdInstructor) INNER JOIN USUARIOS_CLUB ON CUPONES.IdMember = USUARIOS_CLUB.IdMember"
            strSQL = strSQL & " WHERE (((CUPONES.TipoDocumento)='F') AND ((CUPONES.NumeroDocumento) Between " & Me.txtNumFacIni.Text & " AND " & Me.txtNumFacFinal.Text & "))"
            strSQL = strSQL & " ORDER BY CUPONES.NumeroDocumento, CUPONES.IdConcepto, CUPONES.NumeroCupon"
        #End If
    End If
    
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
        Printer.Print "INSCRIPCION: " & adoRcsCupones!NoFamilia
        Printer.Print
        Printer.Print "VALIDO POR " & adoRcsCupones!DescripcionCupon
        If adoRcsCupones!ImprimeImporte Then
            Printer.Print
            Printer.Print Format(adoRcsCupones!ImporteCupon, "$#,##0.00")
        End If
        Printer.Print
        If (adoRcsCupones!RequiereUsuario) And (Trim(adoRcsCupones!Usuario) <> vbNullString) Then
            Printer.Print "Alumno(a): " & adoRcsCupones!Usuario
        End If
        
        If (adoRcsCupones!RequiereUsuario) And (Trim(adoRcsCupones!Instructor) <> vbNullString) Then
            Printer.Print "Instructor(a): " & adoRcsCupones!Instructor
        End If
        
        If Not IsNull(adoRcsCupones!DATOSADICIONALES) Then
            Printer.Print adoRcsCupones!DATOSADICIONALES
        End If
            
        Printer.Print "Vale " & adoRcsCupones!NumeroCupon; " de " & adoRcsCupones!TotalCupones
        Printer.Print "Emitido: " & Format(adoRcsCupones!FechaAlta, "dd/mmm/yy") & " Válido hasta: " & Format(adoRcsCupones!FechaVigencia, "dd/mmm/yy")
        Printer.Print
        Printer.Print "Recibo #: "; adoRcsCupones!NumeroDocumento; " Folio: "; adoRcsCupones!FolioCupon
        
        If Not IsNull(adoRcsCupones!ObservacionesCupon) Then
            Printer.Print
            Printer.FontSize = 7
            Printer.Print adoRcsCupones!ObservacionesCupon
        End If
        
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
    
    adoRcsCupones.Close
    Set adoRcsCupones = Nothing
    
    Set picLogo = Nothing
    
    
    sImpresoraActual = SelectPrinter(sImpresoraActual)
    
    Screen.MousePointer = vbDefault
    
    MDIPrincipal.StatusBar1.Panels(1).Text = ""
    
    
    ActImpCupon Val(Me.txtNumFacIni.Text), Val(Me.txtNumFacFinal.Text)
    
    
    Unload Me
    Exit Sub
    
ERROR_SUB:
    
    'Restablece la impresora
    sImpresoraActual = SelectPrinter(sImpresoraActual)
    
    Screen.MousePointer = vbDefault
    
    MsjError



End Sub

Private Sub Form_Activate()
    If Me.Tag = "F" Then
        'Me.lblTitulo.Caption = "Se generaron las siguientes facturas:"
        Me.txtNumFacIni.Text = lNumFacIniImp
        Me.txtNumFacFinal.Text = lNumFacFinImp
    Else
        'Me.lblTitulo.Caption = "Se generaron los siguientes recibos:"
        Me.txtNumFacIni.Text = lNumRecIniImp
        Me.txtNumFacFinal.Text = lNumRecFinImp
    End If
    
    Me.cmdImprimir.Default = True
    Me.cmdImprimir.SetFocus
    
End Sub

Private Sub Form_Load()


    Me.Height = 3450
    Me.Width = 4770
    
    
    CentraForma MDIPrincipal, Me

    
End Sub
Private Sub ActImpCupon(lRecIni As Long, lRecFin As Long)
    
    Dim adocmdActCupon As ADODB.Command
    
    strSQL = ""
    strSQL = "UPDATE CUPONES SET"
    strSQL = strSQL & " Impresiones=Impresiones + 1"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "(" & "NumeroDocumento Between " & lRecIni & " AND " & lRecFin & ")"
    strslq = strslq & "(" & " AND TipoDocumento=" & "'R'" & ")"
    
    
    Set adocmdActCupon = New ADODB.Command
    
    adocmdActCupon.ActiveConnection = Conn
    adocmdActCupon.CommandType = adCmdText
    adocmdActCupon.CommandText = strSQL

    
    adocmdActCupon.Execute


    Set adocmdActCupon = Nothing
    
    Exit Sub
    
ERROR_SUB:
    
    
    MsjError
    

End Sub
