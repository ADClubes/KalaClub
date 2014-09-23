VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmDatosSocios 
   Caption         =   "Información de los socios"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   Icon            =   "frmDatosSocios.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmDatosSocios.frx":1002
   ScaleHeight     =   6000
   ScaleWidth      =   10125
   Begin VB.OptionButton optTodos 
      Caption         =   "Todos"
      Height          =   255
      Left            =   9120
      TabIndex        =   26
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton optTitular 
      Caption         =   "Titulares"
      Height          =   255
      Left            =   7920
      TabIndex        =   25
      Top             =   840
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpIngreso 
      Height          =   285
      Left            =   7920
      TabIndex        =   24
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   102236161
      CurrentDate     =   38334
   End
   Begin VB.TextBox txtTipo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7920
      TabIndex        =   21
      Top             =   4680
      Width           =   2055
   End
   Begin VB.TextBox txtId 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9360
      TabIndex        =   19
      Top             =   5280
      Width           =   615
   End
   Begin VB.TextBox txtSecuen 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      TabIndex        =   16
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtFoto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7920
      TabIndex        =   15
      Top             =   4080
      Width           =   855
   End
   Begin VB.Frame frmFoto 
      Height          =   2655
      Left            =   7920
      TabIndex        =   14
      Top             =   1080
      Width           =   2055
      Begin VB.Image imgFoto 
         Height          =   2295
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdDesactiva 
      Height          =   615
      Left            =   5880
      Picture         =   "frmDatosSocios.frx":2004
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Desactivar credencial "
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdActiva 
      Height          =   615
      Left            =   5280
      Picture         =   "frmDatosSocios.frx":230E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Activar credencial "
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtRegs 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   9120
      TabIndex        =   12
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton cmdBaja 
      Height          =   615
      Left            =   7920
      Picture         =   "frmDatosSocios.frx":2618
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Dar de baja "
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtBuscar 
      Height          =   300
      Left            =   840
      MaxLength       =   50
      TabIndex        =   11
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdLocaliza 
      Default         =   -1  'True
      Height          =   615
      Left            =   4560
      Picture         =   "frmDatosSocios.frx":2922
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   " Buscar "
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox chkDesdeInicio 
      Caption         =   "&Buscar siempre desde el primer registro"
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   480
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CommandButton cmdNuevo 
      Height          =   615
      Left            =   6480
      Picture         =   "frmDatosSocios.frx":31EC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Nuevo socio "
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdModificar 
      Height          =   615
      Left            =   7200
      Picture         =   "frmDatosSocios.frx":3AB6
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Modificar datos "
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   615
      Left            =   8640
      Picture         =   "frmDatosSocios.frx":3EF8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Refrescar datos "
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   9360
      Picture         =   "frmDatosSocios.frx":4202
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Salir "
      Top             =   120
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dgSocios 
      Bindings        =   "frmDatosSocios.frx":4BF8
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoSocios 
      Height          =   330
      Left            =   5520
      Top             =   5520
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Socios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblFechaIngreso 
      Caption         =   "Fec. Ingreso"
      Height          =   255
      Left            =   7920
      TabIndex        =   23
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblTipo 
      Caption         =   "Tipo de usuario"
      Height          =   255
      Left            =   7920
      TabIndex        =   22
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblId 
      Caption         =   "# usuario"
      Height          =   255
      Left            =   9360
      TabIndex        =   20
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label lblSecuen 
      Caption         =   "# Sec."
      Height          =   255
      Left            =   9000
      TabIndex        =   18
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblFoto 
      Caption         =   "Foto"
      Height          =   255
      Left            =   7920
      TabIndex        =   17
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblRegs 
      Caption         =   "Número de registros"
      Height          =   255
      Left            =   7560
      TabIndex        =   13
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblNotas 
      Caption         =   "Haga click sobre el encabezado de la columna para cambiar el orden"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   5295
   End
End
Attribute VB_Name = "frmDatosSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*  Formulario para mostrar los datos de los titulares y familiares *
'*  Daniel Hdez                                                     *
'*  27 / Septiembre / 2004                                          *
'*  Ultima actualizacion: 31 / Octubre / 2005                       *
'********************************************************************


'Numero de columnas mostradas
Const TOTCOLUMNS = 26

'Encabezados para la consulta
Dim mTitSocios(TOTCOLUMNS) As String

'Ancho de las columnas mostradas
Dim mAncSocios(TOTCOLUMNS) As Integer

'Nombre de los campos en la consulta
Dim mCamSocios(TOTCOLUMNS) As String

'Matriz de los encabezados del formulario
Dim mTitForma(TOTCOLUMNS) As String

'Columna activa o seleccionada para establecer el orden
Public nColActiva As Integer

'Cadena de las tablas que se utilizan en el DbGrid
Dim sTablaSocios As String

'Posicion del cursor dentro del dbGrid
Dim nPos As Variant


Private Sub cmdActiva_Click()
Dim nErrCode As Long
Dim nTor As Long

    With Me.adoSocios.Recordset
        #If SqlServer_ Then
            ActivaCredSQL 1, .Fields("Secuencial"), 1, .Fields("IdMember"), True, True
        #Else
            ActivaCred 1, .Fields("Secuencial"), 1, .Fields("IdMember"), True, True
        #End If
        'nTor = (68 * 16777216) + Fields("Secuencial")
      'nErrCode = AgregaAcceso("68" & .Fields("IdMember"), .Fields("Nombre"), nTor)
    End With
End Sub


Private Sub cmdBaja_Click()
Dim nAnswer As Integer

    nPos = Me.adoSocios.Recordset.Bookmark
    
    If (Me.adoSocios.Recordset.RecordCount > 0) Then
        'Baja de familiar
        If (Me.adoSocios.Recordset.Fields("IdMember") <> Me.adoSocios.Recordset.Fields("IdTitular")) Then
            nAnswer = MsgBox("¿Desea borrar al familiar seleccionado?", vbYesNo, "Baja de familiares")
        
            If (nAnswer = vbYes) Then
                QuitaFamiliar (Me.adoSocios.Recordset.Fields("IdMember")), 0
                Refresca
            End If
        Else
            'Baja de titular
            nAnswer = MsgBox("¿Desea borrar al titular y a sus familiares?", vbYesNo, "Baja de titulares")
        
            If (nAnswer = vbYes) Then
                QuitaTitular (Me.adoSocios.Recordset.Fields("IdMember"))
                Refresca
            End If
        End If
    End If
End Sub


Private Sub cmdDesActiva_Click()
    With Me.adoSocios.Recordset
        #If SqlServer_ Then
            ActivaCredSQL 1, .Fields("Secuencial"), 1, .Fields("IdMember"), False, True
        #Else
            ActivaCred 1, .Fields("Secuencial"), 1, .Fields("IdMember"), False, True
        #End If
    End With
End Sub


Private Sub cmdRefresh_Click()
    Refresca
    
    If (Me.dgSocios.Enabled) Then
        Me.dgSocios.SetFocus
    End If
End Sub


Private Sub cmdSalir_Click()
    'Cierra el formulario
    Unload Me
End Sub


Private Sub Command1_Click()
    Dim adorcsTitulares As ADODB.Recordset
    Dim adorcsCambio As ADODB.Recordset
    
    Dim iContador As Integer
    
    Screen.MousePointer = vbHourglass
    
    strSQL = "SELECT IdTitular, NumeroFamiliar"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " WHERE IdMember=IdTitular"
    strSQL = strSQL & " ORDER BY IdMember"
    
    
    Set adorcsTitulares = New ADODB.Recordset
    
    adorcsTitulares.CursorLocation = adUseServer
    adorcsTitulares.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not adorcsTitulares.EOF
        
        strSQL = "SELECT IdMember, NumeroFamiliar, Status"
        strSQL = strSQL & " FROM USUARIOS_CLUB"
        strSQL = strSQL & " WHERE IdTitular=" & adorcsTitulares!IdTitular
        strSQL = strSQL & " ORDER BY IdMember"
        
        Set adorcsCambio = New ADODB.Recordset
        adorcsCambio.CursorLocation = adUseServer
        
        adorcsCambio.Open strSQL, Conn, adOpenDynamic, adLockOptimistic
        
        iContador = 1
        Do While Not adorcsCambio.EOF
            
            adorcsCambio!NumeroFamiliar = iContador
'            adorcsCambio!Status = "ACTIVO"
            adorcsCambio.Update
            
            adorcsCambio.MoveNext
            iContador = iContador + 1
        Loop
        
        adorcsCambio.Close
        Set adorcsCambio = Nothing
        
        
        adorcsTitulares.MoveNext
    Loop
    
    adorcsTitulares.Close
    Set adorcsTitulares = Nothing
    
    Screen.MousePointer = vbDefault
    MsgBox "Terminado", vbInformation, "Cambio"
    
End Sub


Private Sub dgSocios_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With Me.adoSocios.Recordset
        If ((.RecordCount > 0) And (Not .EOF)) Then
            Me.txtFoto.Text = .Fields("FotoFile")
            Me.txtSecuen.Text = .Fields("Secuencial")
            Me.txtTipo.Text = .Fields("Descripcion")
            Me.dtpIngreso.Value = Format(.Fields("FechaIngreso"), "dd/mm/yyyy")
            Me.txtId.Text = .Fields("idMember")
            
            If (Dir(sG_RutaFoto & "\" & Trim(.Fields("FotoFile")) & ".jpg") <> "") Then
                Me.imgFoto.Picture = LoadPicture(sG_RutaFoto & "\" & Trim(.Fields("FotoFile")) & ".jpg")
            Else
                Me.imgFoto.Picture = LoadPicture("")
            End If
        End If
    End With
End Sub


Private Sub Form_Activate()
    Me.dgSocios.Visible = False

    'Encabezados del formulario
    mTitForma(0) = " Clave"
    mTitForma(1) = " # Fam."
    mTitForma(2) = " Nombre"
    mTitForma(3) = " parentesco"
    mTitForma(4) = " # secuencial"
    mTitForma(5) = " fecha de nacimiento"
    mTitForma(6) = " fecha de ingreso"
    mTitForma(7) = " status"
    mTitForma(8) = " sexo"
    mTitForma(9) = " país de origen"
    mTitForma(10) = " tipo de usuario"
    mTitForma(11) = " celular"
    mTitForma(12) = " email"
    mTitForma(13) = " profesión"
    mTitForma(14) = " email"
    mTitForma(15) = " profesión"
    mTitForma(16) = " serie"
    mTitForma(17) = " tipo"
    mTitForma(18) = " acción"
    mTitForma(19) = " foto"
    mTitForma(20) = " clave"
    mTitForma(21) = " titular"
    mTitForma(22) = " apellido paterno"
    mTitForma(23) = " apellido materno"
    mTitForma(24) = " nombre(s)"
    mTitForma(25) = " No.Inscripcion"
    
    
    'Campos que aparecen en dgSocios
    
    mCamSocios(0) = "Usuarios_Club.NoFamilia"
    mCamSocios(1) = "Usuarios_Club.NumeroFamiliar"
    mCamSocios(2) = "Usuarios_Club.A_Paterno + ' ' + Usuarios_Club.A_Materno + ' ' + Usuarios_Club.Nombre AS Nombre"
    mCamSocios(3) = "Parentesco.Parentesco"
    mCamSocios(4) = "Secuencial.Secuencial"
    mCamSocios(5) = "Usuarios_Club.FechaNacio"
    mCamSocios(6) = "Usuarios_Club.FechaIngreso"
    mCamSocios(7) = "Usuarios_Club.Status"
    mCamSocios(8) = "Usuarios_Club.Sexo"
    mCamSocios(9) = "Paises.Pais"
    mCamSocios(10) = "Tipo_Usuario.Descripcion"
    mCamSocios(11) = "Usuarios_Club.Celular"
    mCamSocios(12) = "Usuarios_Club.Email"
    mCamSocios(13) = "Usuarios_Club.Profesion"
    mCamSocios(14) = "Tipo_Pago.Descripcion"
    mCamSocios(15) = "Tipo_Uso_Accion.Descripcion"
    mCamSocios(16) = "Usuarios_Titulo.Serie"
    mCamSocios(17) = "Usuarios_Titulo.Tipo"
    mCamSocios(18) = "Usuarios_Titulo.Numero"
    mCamSocios(19) = "Usuarios_Club.FotoFile"
    mCamSocios(20) = "Usuarios_Club.IdMember"
    mCamSocios(21) = "Usuarios_Club.IdTitular"
    mCamSocios(22) = "Usuarios_Club.A_Paterno"
    mCamSocios(23) = "Usuarios_Club.A_Materno"
    mCamSocios(24) = "Usuarios_Club.Nombre"
    mCamSocios(25) = "Usuarios_Club.Inscripcion"
    
    'Cadena para la union de las tablas en el DbGrid de los padres
    sTablaSocios = "((((((Usuarios_Club LEFT JOIN Paises ON Usuarios_Club.IdPais=Paises.IdPais) "
    sTablaSocios = sTablaSocios & "LEFT JOIN Tipo_Usuario ON Usuarios_Club.IdTipoUsuario=Tipo_Usuario.IdTipoUsuario) "
    sTablaSocios = sTablaSocios & "LEFT JOIN Parentesco ON Tipo_Usuario.Parentesco=Parentesco.Clave) "
    sTablaSocios = sTablaSocios & "LEFT JOIN Usuarios_Titulo ON Usuarios_Club.IdMember=Usuarios_Titulo.IdMember) "
    sTablaSocios = sTablaSocios & "LEFT JOIN Tipo_Uso_Accion ON Usuarios_Titulo.IdTipoUsoAccion=Tipo_Uso_Accion.IdTipoUsoAccion) "
    sTablaSocios = sTablaSocios & "LEFT JOIN Tipo_Pago ON Usuarios_Titulo.IdTipoPago=Tipo_Pago.IdTipoPago) "
    sTablaSocios = sTablaSocios & "LEFT JOIN Secuencial ON Usuarios_Club.IdMember=Secuencial.IdMember "
    
    Refresca
    
    Me.WindowState = 2
    
    Me.dgSocios.Visible = True
End Sub


Private Sub Form_Load()
    'Propiedades del formulario
    With Me
        .txtBuscar.Text = ""
        '.chkDesdeInicio.Value = 0
        .Top = 0
        .Left = 0
        .Width = 10245
        .Height = 6510
    End With
    
    nPos = 0
    
    'Encabezados de los DbGrid
    Me.dgSocios.Caption = "Información de titulares"
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Información de los usuarios del club"
End Sub




Private Sub Form_Unload(Cancel As Integer)
    Set frmDatosSocios = Nothing
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "KALACLUB"
End Sub


Private Sub cmdLocaliza_Click()
Dim strBuscar As String
Dim nPos As Integer

    If ((Me.adoSocios.Recordset.RecordCount > 0) And (Trim(Me.txtBuscar.Text) <> "")) Then
        'Cambia el texto a mayusculas
        Me.txtBuscar.Text = UCase$(txtBuscar.Text)
        Me.txtBuscar.REFRESH
        
        strBuscar = Me.dgSocios.Columns.Item(nColActiva).DataField & " like "
        
        Select Case nColActiva
            Case 0, 4, 19 To 24
                strBuscar = strBuscar & Val(Me.txtBuscar.Text)
                
            Case 1 To 2, 8 To 17
                'nPos = InStr(1, Me.txtBuscar.Text, " ")
                'If (nPos > 1) Then
                '    Me.txtBuscar.Text = Mid(Me.txtBuscar.Text, 1, (nPos - 1))
                '    Me.txtBuscar.REFRESH
                'End If
                strBuscar = strBuscar & "'" & Trim(Me.txtBuscar.Text) & "%'"
                
            Case 5, 6
                strBuscar = strBuscar & "#" & Format(Me.txtBuscar.Text, "dd/mm/yyyy") & "#"
        End Select
        
        'Busca el primer registro que coincida con el texto escrito
        BuscarEnDG Me.adoSocios, chkDesdeInicio, strBuscar
    End If
End Sub


Private Sub cmdNuevo_Click()
    If (Me.adoSocios.Recordset.RecordCount > 0) Then
        nPos = Me.adoSocios.Recordset.Bookmark
    End If
    
    Me.WindowState = 1

    frmAltaSocios.sFormaAnterior = "frmDatosSocios"
    frmAltaSocios.bSocioNvo = True
    Load frmAltaSocios
    frmAltaSocios.Show (1)
End Sub


Private Sub cmdModificar_Click()
    If (Me.adoSocios.Recordset.RecordCount > 0) Then
        If (Me.adoSocios.Recordset.Fields("IdMember") = Me.adoSocios.Recordset.Fields("IdTitular")) Then
            ModificaDatos
        Else
            nPos = Me.adoSocios.Recordset.Bookmark
            
            Me.WindowState = 1
            
            frmAltaFam.bNvoFam = False
            frmAltaFam.nCveFam = Me.adoSocios.Recordset.Fields("IdMember")
            frmAltaFam.Show (1)
            
'            Me.WindowState = 2
        End If
    End If
End Sub


Private Sub dgSocios_KeyPress(KeyAscii As Integer)
Dim nLen As Byte
Dim lBuscar As Boolean

    lBuscar = False
    If (Me.adoSocios.Recordset.RecordCount > 0) Then
        Select Case KeyAscii
            Case 65 To 90
                Me.txtBuscar.Text = Me.txtBuscar.Text + Chr(KeyAscii)
                lBuscar = True
                
            Case 97 To 122
                Me.txtBuscar.Text = Me.txtBuscar.Text + UCase(Chr(KeyAscii))
                lBuscar = True
            
            Case 8
                nLen = Len(Trim(Me.txtBuscar.Text))
                If (nLen > 0) Then
                    Me.txtBuscar.Text = Mid(Me.txtBuscar.Text, 1, nLen - 1)
                    lBuscar = True
                End If
        End Select
        
        If ((Trim(Me.txtBuscar.Text) <> "") And (lBuscar)) Then
            Me.adoSocios.Recordset.Find Me.adoSocios.Recordset.Fields(nColActiva).Name & " LIKE '*" & LTrim(Me.txtBuscar.Text) & "*'"
            If (Me.adoSocios.Recordset.EOF) Then
                Me.adoSocios.Recordset.MoveFirst
            End If
        End If
    End If
End Sub


'Cambia el orden en que aparecen los datos en pantalla
Private Sub dgSocios_HeadClick(ByVal ColIndex As Integer)
    nColActiva = ColIndex
    Refresca
End Sub


Private Sub InitdgSocios()
    'Asigna valores a la matriz de encabezados
    mTitSocios(0) = "Clave"
    mTitSocios(1) = "# Fam."
    mTitSocios(2) = "Nombre"
    mTitSocios(3) = "Parentesco"
    mTitSocios(4) = "# Sec."
    mTitSocios(5) = "Nació"
    mTitSocios(6) = "Ingresó"
    mTitSocios(7) = "Status"
    mTitSocios(8) = "Sexo"
    mTitSocios(9) = "País de origen"
    mTitSocios(10) = "Tipo de usuario"
    mTitSocios(11) = "Celular"
    mTitSocios(12) = "Email"
    mTitSocios(13) = "Profesión"
    mTitSocios(14) = "Como paga"
    mTitSocios(15) = "Uso de la acción"
    mTitSocios(16) = "Serie"
    mTitSocios(17) = "Tipo"
    mTitSocios(18) = "Número"
    mTitSocios(19) = "Foto"
    mTitSocios(20) = "Cve"
    mTitSocios(21) = "Titular"
    mTitSocios(22) = "A. Paterno"
    mTitSocios(23) = "A. Materno"
    mTitSocios(24) = "Nombre(s)"
    mTitSocios(25) = "No. Inscripcion"
    
    'Asigna los encabezados de las columnas
    DefHeadersDBGrid dgSocios, mTitSocios
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncSocios(0) = 700
    mAncSocios(1) = 700
    mAncSocios(2) = 5000
    mAncSocios(3) = 1500
    mAncSocios(4) = 800
    mAncSocios(5) = 1100
    mAncSocios(6) = 1100
    mAncSocios(7) = 800
    mAncSocios(8) = 700
    mAncSocios(9) = 1400
    mAncSocios(10) = 600
    mAncSocios(11) = 2500
    mAncSocios(12) = 3000
    mAncSocios(13) = 3000
    mAncSocios(14) = 3500
    mAncSocios(15) = 2500
    mAncSocios(16) = 1500
    mAncSocios(17) = 1700
    mAncSocios(18) = 900
    mAncSocios(19) = 900
    mAncSocios(20) = 900
    mAncSocios(21) = 700
    mAncSocios(22) = 2400
    mAncSocios(23) = 2400
    mAncSocios(24) = 2400
    mAncSocios(25) = 1600

    'Asigna el ancho de cada columna
    DefAnchoDBGrid dgSocios, mAncSocios
    
    'Evita que se puedan modificar los datos de la consulta
    Me.dgSocios.AllowUpdate = False
End Sub


'Llama a la forma que permite modificar datos
Private Sub ModificaDatos()
    If (Me.adoSocios.Recordset.RecordCount > 0) Then
        nPos = Me.adoSocios.Recordset.Bookmark
        
        Me.WindowState = 1
        
        frmAltaSocios.sFormaAnterior = "frmDatosSocios"
        frmAltaSocios.bSocioNvo = False
        Load frmAltaSocios
        frmAltaSocios.Show (1)
        
'        Me.WindowState = 2
    End If
End Sub


Public Sub Refresca()
    Dim sCondicion As String

    '11/10/05 gpo
    Screen.MousePointer = vbHourglass

    sCondicion = "Secuencial.Temporal=0"
    
    If (Me.optTitular.Value) Then
        sCondicion = sCondicion & " AND Usuarios_Club.idTitular=Usuarios_Club.idMember"
    End If

    'Escribe el encabezado del formulario
    frmDatosSocios.Caption = "Información de los socios ordenada por: " & mTitForma(nColActiva)

    'Inicializa un ctrlAdo de los padres con una lista de campos
    InitCtrlAdoSel Me.adoSocios, sTablaSocios, mCamSocios, TOTCOLUMNS, nColActiva, sCondicion, Conn
    
    'Configura el ctrl DataGrid
    InitdgSocios
    
    'Quita el efecto que deja la columna en color negro
    Me.dgSocios.ClearSelCols
    
    Me.txtRegs.Text = Me.adoSocios.Recordset.RecordCount
    
    If (Val(Me.txtRegs.Text) <= 0) Then
        Me.dgSocios.Enabled = False
    End If
    
    If (nPos > Val(Me.txtRegs.Text)) Then
        nPos = Val(Me.txtRegs.Text)
    End If
    
    If ((Not Me.adoSocios.Recordset.EOF) And (nPos > 0)) Then
        Me.adoSocios.Recordset.Bookmark = nPos
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub Form_Resize()
Const MARGEN_01 = 2200
Const MARGEN_02 = 1000

    If ((ScaleWidth - 250) > 0) Then
        Me.dgSocios.Width = ScaleWidth - MARGEN_01 - 250
        Me.lblRegs.Left = ScaleWidth - 2600
        Me.txtRegs.Left = ScaleWidth - MARGEN_02

        Me.optTitular.Left = ScaleWidth - MARGEN_01
        Me.frmFoto.Left = ScaleWidth - MARGEN_01
        Me.lblFoto.Left = ScaleWidth - MARGEN_01
        Me.txtFoto.Left = ScaleWidth - MARGEN_01
        Me.lblTipo.Left = ScaleWidth - MARGEN_01
        Me.txtTipo.Left = ScaleWidth - MARGEN_01
        Me.lblFechaIngreso.Left = ScaleWidth - MARGEN_01
        Me.dtpIngreso.Left = ScaleWidth - MARGEN_01

        Me.optTodos.Left = ScaleWidth - MARGEN_02

        Me.lblSecuen.Left = ScaleWidth - MARGEN_02 - 120
        Me.txtSecuen.Left = ScaleWidth - MARGEN_02 - 120

        Me.lblId.Left = ScaleWidth - MARGEN_02 + 240
        Me.txtId.Left = ScaleWidth - MARGEN_02 + 240
    End If

    If ((ScaleHeight - 1100) > 0) Then
        Me.dgSocios.Height = ScaleHeight - 1250
        Me.lblNotas.Top = ScaleHeight - 290
        Me.lblRegs.Top = ScaleHeight - 290
        Me.txtRegs.Top = ScaleHeight - 290
    End If
End Sub


Private Sub optTitular_Click()
    'Encabezados de los DbGrid
    Me.dgSocios.Caption = "Información de titulares"

    Refresca
End Sub


Private Sub optTodos_Click()
    'Encabezados de los DbGrid
    Me.dgSocios.Caption = "Información de los titulares y familiares"
    
    Refresca
End Sub
