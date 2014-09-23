VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmConsSocios 
   Caption         =   "Consultas"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   Icon            =   "frmConsSocios.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   10575
   Begin VB.CommandButton cmdRefresh 
      Height          =   615
      Left            =   9000
      Picture         =   "frmConsSocios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   9840
      Picture         =   "frmConsSocios.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin MSAdodcLib.Adodc adoDatos 
      Height          =   375
      Left            =   360
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
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
      Caption         =   "Datos"
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
   Begin MSDataGridLib.DataGrid dgDatos 
      Bindings        =   "frmConsSocios.frx":0A56
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10398
      _Version        =   393216
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
End
Attribute VB_Name = "frmConsSocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*  Formulario para mostrar las consultas                           *
'*  Daniel Hdez                                                     *
'*  09 / Septiembre / 2004                                          *
'*  Ultima actualización: 15 / Noviembre / 2004                     *
'********************************************************************


'Constantes
Const OCHO_COLS = 8
Const DIEZ_COLS = 10
Const ONCE_COLS = 11
Const CATORCE_COLS = 14

'Encabezados para la consulta
Dim mTitCons() As String

'Ancho de las columnas mostradas
Dim mAncCons() As Integer

'Nombre de los campos en la consulta
Dim mCamCons() As String

'Matriz de los encabezados del formulario
Dim mTitForma() As String

'Columna activa o seleccionada para establecer el orden
Dim nColActiva As Integer

'Tipo de consulta
Public nCons As Byte

'Cadena de las tablas que se utilizan en el DbGrid
Dim sTablaCons As String

'Numero de columnas
Dim nColumns As Byte

'Condicion
Dim sCond As String

'Titulo del dbGrid
Dim sTit  As String


Private Sub cmdRefresh_Click()
    Refresca

    If (Me.dgDatos.Enabled) Then
        Me.dgDatos.SetFocus
    End If
End Sub


Private Sub cmdSalir_Click()
    'Cierra el formulario
    Unload Me
End Sub


Private Sub Form_Activate()
Dim sTitulo As String

    'Centra el formulario
    CENTRAFORMA MDIPrincipal, frmConsSocios

    Select Case nCons
        Case 1
            sTitulo = "Consulta de ausencias"
            
        Case 2
            sTitulo = "Consulta de autos"
        
        Case 3
            sTitulo = "Consulta de certificados médicos"
        
        Case 4
            sTitulo = "Consulta de domicilios"
        
        Case 5
            sTitulo = "Consulta de horarios"
        
        Case 6
            sTitulo = "Consulta de pases temporales"
        
        Case 7
            sTitulo = "Consulta de rentables"
        
    End Select
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTitulo
End Sub


Private Sub Form_Load()
    Select Case nCons
        Case 1
            ReDim mTitForma(OCHO_COLS)
            ReDim mTitCons(OCHO_COLS)
            ReDim mCamCons(OCHO_COLS)
            ReDim mAncCons(OCHO_COLS)
        
            nColumns = OCHO_COLS
            nColActiva = 2
            sTit = "Ausencias ordenadas por: "
        
            'Encabezados del formulario
            mTitForma(0) = " # reg."
            mTitForma(1) = " # usuario"
            mTitForma(2) = " A. paterno"
            mTitForma(3) = " A. materno"
            mTitForma(4) = " nombre"
            mTitForma(5) = " inicia"
            mTitForma(6) = " termina"
            mTitForma(7) = " % Desc."
            
            'Asigna valores a la matriz de encabezados
            mTitCons(0) = "# Reg."
            mTitCons(1) = "# usuario"
            mTitCons(2) = "A. paterno"
            mTitCons(3) = "A. materno"
            mTitCons(4) = "Nombre"
            mTitCons(5) = "Inicia"
            mTitCons(6) = "Termina"
            mTitCons(7) = "% Desc."
            
            'Campos que aparecen en dgDatos
            mCamCons(0) = "Ausencias!IdAusencia"
            mCamCons(1) = "Ausencias!IdMember"
            mCamCons(2) = "Usuarios_Club!A_Paterno"
            mCamCons(3) = "Usuarios_Club!A_Materno"
            mCamCons(4) = "Usuarios_Club!Nombre"
            mCamCons(5) = "Ausencias!FechaInicial"
            mCamCons(6) = "Ausencias!FechaFinal"
            mCamCons(7) = "Ausencias!Porcentaje"
            
            'Asigna valores a la matriz que define el ancho de cada columna
            mAncCons(0) = 700
            mAncCons(1) = 900
            mAncCons(2) = 1800
            mAncCons(3) = 1800
            mAncCons(4) = 2800
            mAncCons(5) = 1100
            mAncCons(6) = 1100
            mAncCons(7) = 900
            
            'Cadena para la union de las tablas en el DbGrid de los padres
            sTablaCons = "Ausencias LEFT JOIN Usuarios_Club ON Ausencias.IdMember=Usuarios_Club.IdMember "
        
        Case 2
            ReDim mTitForma(OCHO_COLS)
            ReDim mTitCons(OCHO_COLS)
            ReDim mCamCons(OCHO_COLS)
            ReDim mAncCons(OCHO_COLS)
            
            nColumns = OCHO_COLS
            nColActiva = 2
            sTit = "Autos ordenados por: "
        
            'Encabezados del formulario
            mTitForma(0) = " # reg."
            mTitForma(1) = " # usuario"
            mTitForma(2) = " A. paterno"
            mTitForma(3) = " A. materno"
            mTitForma(4) = " nombre"
            mTitForma(5) = " # calc."
            mTitForma(6) = " placas"
            mTitForma(7) = " descripción"
            
            'Asigna valores a la matriz de encabezados
            mTitCons(0) = "# Reg."
            mTitCons(1) = "# usuario"
            mTitCons(2) = "A. paterno"
            mTitCons(3) = "A. materno"
            mTitCons(4) = "Nombre"
            mTitCons(5) = "# Calc."
            mTitCons(6) = "Placas"
            mTitCons(7) = "Descripción"
            
            'Campos que aparecen en dgDatos
            mCamCons(0) = "Calcomanias!Id"
            mCamCons(1) = "Calcomanias!IdMember"
            mCamCons(2) = "Usuarios_Club!A_Paterno"
            mCamCons(3) = "Usuarios_Club!A_Materno"
            mCamCons(4) = "Usuarios_Club!Nombre"
            mCamCons(5) = "Calcomanias!Numero"
            mCamCons(6) = "Calcomanias!Placa"
            mCamCons(7) = "Calcomanias!Descripcion"
            
            'Asigna valores a la matriz que define el ancho de cada columna
            mAncCons(0) = 700
            mAncCons(1) = 900
            mAncCons(2) = 1800
            mAncCons(3) = 1800
            mAncCons(4) = 2800
            mAncCons(5) = 1100
            mAncCons(6) = 1100
            mAncCons(7) = 3500
            
            'Cadena para la union de las tablas en el DbGrid de los padres
            sTablaCons = "Calcomanias LEFT JOIN Usuarios_Club ON Calcomanias.IdMember=Usuarios_Club.IdMember "
        
        Case 3
            ReDim mTitForma(ONCE_COLS)
            ReDim mTitCons(ONCE_COLS)
            ReDim mCamCons(ONCE_COLS)
            ReDim mAncCons(ONCE_COLS)
            
            nColumns = ONCE_COLS
            nColActiva = 2
            sTit = "Certificados médicos ordenados por: "
        
            'Encabezados del formulario
            mTitForma(0) = " # reg."
            mTitForma(1) = " # usuario"
            mTitForma(2) = " A. paterno"
            mTitForma(3) = " A. materno"
            mTitForma(4) = " nombre"
            mTitForma(5) = " peso"
            mTitForma(6) = " estatura"
            mTitForma(7) = " fecha"
            mTitForma(8) = " # médico"
            mTitForma(9) = " apellido"
            mTitForma(10) = " nombre del médico"
            
            'Asigna valores a la matriz de encabezados
            mTitCons(0) = "# Reg."
            mTitCons(1) = "# usuario"
            mTitCons(2) = "A. paterno"
            mTitCons(3) = "A. materno"
            mTitCons(4) = "Nombre"
            mTitCons(5) = "Peso"
            mTitCons(6) = "Estatura"
            mTitCons(7) = "Fecha"
            mTitCons(8) = "# Médico"
            mTitCons(9) = "Apellido"
            mTitCons(10) = "Nombre"
            
            'Campos que aparecen en dgDatos
            mCamCons(0) = "Certificados!IdCertificado"
            mCamCons(1) = "Certificados!IdMember"
            mCamCons(2) = "Usuarios_Club!A_Paterno"
            mCamCons(3) = "Usuarios_Club!A_Materno"
            mCamCons(4) = "Usuarios_Club!Nombre"
            mCamCons(5) = "Certificados!Peso"
            mCamCons(6) = "Certificados!Estatura"
            mCamCons(7) = "Certificados!Fecha"
            mCamCons(8) = "Certificados!IdMedico"
            mCamCons(9) = "Medicos!A_Paterno"
            mCamCons(10) = "Medicos!Nombre"
            
            'Asigna valores a la matriz que define el ancho de cada columna
            mAncCons(0) = 700
            mAncCons(1) = 900
            mAncCons(2) = 1800
            mAncCons(3) = 1800
            mAncCons(4) = 2800
            mAncCons(5) = 1100
            mAncCons(6) = 1100
            mAncCons(7) = 1100
            mAncCons(8) = 900
            mAncCons(9) = 1800
            mAncCons(10) = 2800
            
            'Cadena para la union de las tablas en el DbGrid de los padres
            sTablaCons = "(Certificados LEFT JOIN Usuarios_Club ON Certificados.IdMember=Usuarios_Club.IdMember) "
            sTablaCons = sTablaCons & "LEFT JOIN Medicos ON Certificados.IdMedico=Medicos.IdMedico "
            
        Case 4
            ReDim mTitForma(CATORCE_COLS)
            ReDim mTitCons(CATORCE_COLS)
            ReDim mCamCons(CATORCE_COLS)
            ReDim mAncCons(CATORCE_COLS)
            
            nColumns = CATORCE_COLS
            nColActiva = 2
            sTit = "Domicilios ordenados por: "
        
            'Encabezados del formulario
            mTitForma(0) = " # reg."
            mTitForma(1) = " # usuario"
            mTitForma(2) = " A. paterno"
            mTitForma(3) = " A. materno"
            mTitForma(4) = " nombre"
            mTitForma(5) = " calle"
            mTitForma(6) = " colonia"
            mTitForma(7) = " Del. o municipio"
            mTitForma(8) = " estado"
            mTitForma(9) = " CP"
            mTitForma(10) = " Tel. 1"
            mTitForma(11) = " Tel. 2"
            mTitForma(12) = " fax"
            mTitForma(13) = " tipo"
            
            'Asigna valores a la matriz de encabezados
            mTitCons(0) = "# Reg."
            mTitCons(1) = "# usuario"
            mTitCons(2) = "A. paterno"
            mTitCons(3) = "A. materno"
            mTitCons(4) = "Nombre"
            mTitCons(5) = "Calle"
            mTitCons(6) = "Colonia"
            mTitCons(7) = "Del. o municipio"
            mTitCons(8) = "Estado"
            mTitCons(9) = "CP"
            mTitCons(10) = "Tel. 1"
            mTitCons(11) = "Tel. 2"
            mTitCons(12) = "Fax"
            mTitCons(13) = "Tipo Dir."
            
            'Campos que aparecen en dgDatos
            mCamCons(0) = "Direcciones!IdDireccion"
            mCamCons(1) = "Direcciones!IdMember"
            mCamCons(2) = "Usuarios_Club!A_Paterno"
            mCamCons(3) = "Usuarios_Club!A_Materno"
            mCamCons(4) = "Usuarios_Club!Nombre"
            mCamCons(5) = "Direcciones!Calle"
            mCamCons(6) = "Direcciones!Colonia"
            mCamCons(7) = "DelgaMunici!NomDeloMuni"
            mCamCons(8) = "EntFederativa!NomEntFederativa"
            mCamCons(9) = "Direcciones!CodPos"
            mCamCons(10) = "Direcciones!Tel1"
            mCamCons(11) = "Direcciones!Tel2"
            mCamCons(12) = "Direcciones!Fax"
            mCamCons(13) = "Tipo_Direccion!Descripcion"
            
            'Asigna valores a la matriz que define el ancho de cada columna
            mAncCons(0) = 700
            mAncCons(1) = 900
            mAncCons(2) = 1800
            mAncCons(3) = 1800
            mAncCons(4) = 2800
            mAncCons(5) = 4500
            mAncCons(6) = 3500
            mAncCons(7) = 3500
            mAncCons(8) = 3500
            mAncCons(9) = 900
            mAncCons(10) = 1100
            mAncCons(11) = 1100
            mAncCons(12) = 1100
            mAncCons(13) = 3500
            
            'Cadena para la union de las tablas en el DbGrid de los padres
            sTablaCons = "(((Direcciones LEFT JOIN Usuarios_Club ON Direcciones.IdMember=Usuarios_Club.IdMember) "
            sTablaCons = sTablaCons & "LEFT JOIN Tipo_Direccion ON Direcciones.IdTipoDireccion=Tipo_Direccion.IdTipoDireccion) "
            sTablaCons = sTablaCons & "LEFT JOIN DelgaMunici ON Direcciones.CveDeloMuni=DelgaMunici.CveDeloMuni) "
            sTablaCons = sTablaCons & "LEFT JOIN EntFederativa ON DelgaMunici.EntidadFed=EntFederativa.CveEntFederativa "
        
        Case 5
            ReDim mTitForma(OCHO_COLS)
            ReDim mTitCons(OCHO_COLS)
            ReDim mCamCons(OCHO_COLS)
            ReDim mAncCons(OCHO_COLS)
            
            nColumns = OCHO_COLS
            nColActiva = 3
            sTit = "Horarios ordenados por: "
        
            'Encabezados del formulario
            mTitForma(0) = " # reg."
            mTitForma(1) = " # familia"
            mTitForma(2) = " # usuario"
            mTitForma(3) = " A. paterno"
            mTitForma(4) = " A. materno"
            mTitForma(5) = " nombre"
            mTitForma(6) = " # horario"
            mTitForma(7) = " descripción"
            
            'Asigna valores a la matriz de encabezados
            mTitCons(0) = "# Reg."
            mTitCons(1) = "# Familia"
            mTitCons(2) = "# usuario"
            mTitCons(3) = "A. paterno"
            mTitCons(4) = "A. materno"
            mTitCons(5) = "Nombre"
            mTitCons(6) = "# horario"
            mTitCons(7) = "Descripción"
            
            'Campos que aparecen en dgDatos
            mCamCons(0) = "Time_Zone_Users!IdReg"
            mCamCons(1) = "Time_Zone_Users!NoFamilia"
            mCamCons(2) = "Time_zone_Users!IdMember"
            mCamCons(3) = "Usuarios_Club!A_Paterno"
            mCamCons(4) = "Usuarios_Club!A_Materno"
            mCamCons(5) = "Usuarios_Club!Nombre"
            mCamCons(6) = "Time_Zone_Users!IdTimeZone"
            mCamCons(7) = "Time_Zone!Descripcion"
            
            'Asigna valores a la matriz que define el ancho de cada columna
            mAncCons(0) = 700
            mAncCons(1) = 700
            mAncCons(2) = 900
            mAncCons(3) = 1800
            mAncCons(4) = 1800
            mAncCons(5) = 2800
            mAncCons(6) = 900
            mAncCons(7) = 3500
            
            'Cadena para la union de las tablas en el DbGrid de los padres
            sTablaCons = "(Time_Zone_Users LEFT JOIN Usuarios_Club ON Time_Zone_Users.IdMember=Usuarios_Club.IdMember) "
            sTablaCons = sTablaCons & "LEFT JOIN Time_Zone ON Time_Zone_Users.IdTimeZone=Time_Zone.IdTimeZone "
        
        Case 6
            ReDim mTitForma(ONCE_COLS)
            ReDim mTitCons(ONCE_COLS)
            ReDim mCamCons(ONCE_COLS)
            ReDim mAncCons(ONCE_COLS)
            
            nColumns = ONCE_COLS
            nColActiva = 2
            sTit = "Pases temporales ordenados por: "
        
            'Encabezados del formulario
            mTitForma(0) = " # reg."
            mTitForma(1) = " # usuario"
            mTitForma(2) = " A. paterno"
            mTitForma(3) = " A. materno"
            mTitForma(4) = " nombre"
            mTitForma(5) = " # Sec."
            mTitForma(6) = " se otorgó a"
            mTitForma(7) = " causa"
            mTitForma(8) = " inicia"
            mTitForma(9) = " termina"
            mTitForma(10) = " horario"
            
            'Asigna valores a la matriz de encabezados
            mTitCons(0) = "# Reg."
            mTitCons(1) = "# usuario"
            mTitCons(2) = "A. paterno"
            mTitCons(3) = "A. materno"
            mTitCons(4) = "Nombre"
            mTitCons(5) = "# sec."
            mTitCons(6) = "Se otorgó a"
            mTitCons(7) = "Causa"
            mTitCons(8) = "Inicia"
            mTitCons(9) = "Termina"
            mTitCons(10) = "Horario"
            
            'Campos que aparecen en dgDatos
            mCamCons(0) = "Pases_Temporales!IdPase"
            mCamCons(1) = "Secuencial!IdMember"
            mCamCons(2) = "Usuarios_Club!A_Paterno"
            mCamCons(3) = "Usuarios_Club!A_Materno"
            mCamCons(4) = "Usuarios_Club!Nombre"
            mCamCons(5) = "Pases_Temporales!Secuencial"
            mCamCons(6) = "Pases_Temporales!QuienRecibe"
            mCamCons(7) = "Causas_Pase!Descripcion"
            mCamCons(8) = "Pases_Temporales!FechaInicio"
            mCamCons(9) = "Pases_Temporales!FechaFinal"
            mCamCons(10) = "Time_Zone!Descripcion"
            
            'Asigna valores a la matriz que define el ancho de cada columna
            mAncCons(0) = 700
            mAncCons(1) = 900
            mAncCons(2) = 1800
            mAncCons(3) = 1800
            mAncCons(4) = 2800
            mAncCons(5) = 700
            mAncCons(6) = 3500
            mAncCons(7) = 3500
            mAncCons(8) = 1100
            mAncCons(9) = 1100
            mAncCons(10) = 3500
            
            'Cadena para la union de las tablas en el DbGrid de los padres
            sTablaCons = "(((Pases_Temporales LEFT JOIN Secuencial ON Pases_Temporales.Secuencial=Secuencial.Secuencial) "
            sTablaCons = sTablaCons & "LEFT JOIN Causas_Pase ON Pases_Temporales.IdCausa=Causas_Pase.IdCausa) "
            sTablaCons = sTablaCons & "LEFT JOIN Usuarios_Club ON Secuencial.IdMember=Usuarios_Club.IdMember) "
            sTablaCons = sTablaCons & "LEFT JOIN Time_Zone ON Pases_Temporales.IdTimeZone=Time_Zone.IdTimeZone "
        
        Case 7
            ReDim mTitForma(DIEZ_COLS)
            ReDim mTitCons(DIEZ_COLS)
            ReDim mCamCons(DIEZ_COLS)
            ReDim mAncCons(DIEZ_COLS)
            
            nColumns = DIEZ_COLS
            nColActiva = 1
            sTit = "Rentables ordenados por: "
        
            'Encabezados del formulario
            mTitForma(0) = " número"
            mTitForma(1) = " tipo"
            mTitForma(2) = " ubicación"
            mTitForma(3) = " sexo"
            mTitForma(4) = " pagado"
            mTitForma(5) = " propiedad"
            mTitForma(6) = " # usuario"
            mTitForma(7) = " A. paterno"
            mTitForma(8) = " A. materno"
            mTitForma(9) = " nombre"
            
            'Asigna valores a la matriz de encabezados
            mTitCons(0) = "Número"
            mTitCons(1) = "Tipo"
            mTitCons(2) = "Ubicación"
            mTitCons(3) = "Sexo"
            mTitCons(4) = "Fecha pago"
            mTitCons(5) = "Propiedad"
            mTitCons(6) = "# usuario"
            mTitCons(7) = "A. paterno"
            mTitCons(8) = "A. materno"
            mTitCons(9) = "Nombre"
            
            'Campos que aparecen en dgDatos
            mCamCons(0) = "Rentables!Numero"
            mCamCons(1) = "Tipo_Rentables!Descripcion"
            mCamCons(2) = "Rentables!Ubicacion"
            mCamCons(3) = "Rentables!Sexo"
            mCamCons(4) = "Rentables!FechaPago"
            mCamCons(5) = "Rentables!Propiedad"
            mCamCons(6) = "Rentables!IdUsuario"
            mCamCons(7) = "Usuarios_Club!A_Paterno"
            mCamCons(8) = "Usuarios_Club!A_Materno"
            mCamCons(9) = "Usuarios_Club!Nombre"
            
            'Asigna valores a la matriz que define el ancho de cada columna
            mAncCons(0) = 900
            mAncCons(1) = 2000
            mAncCons(2) = 2000
            mAncCons(3) = 900
            mAncCons(4) = 1100
            mAncCons(5) = 700
            mAncCons(6) = 900
            mAncCons(7) = 1800
            mAncCons(8) = 1800
            mAncCons(9) = 2800
            
            'Cadena para la union de las tablas en el DbGrid de los padres
            sTablaCons = "(Rentables LEFT JOIN Usuarios_Club ON Rentables.IdUsuario=Usuarios_Club.IdMember) "
            sTablaCons = sTablaCons & "LEFT JOIN Tipo_Rentables ON Rentables.IdTipoRentable=Tipo_Rentables.IdTipoRentable "
            
            'Condicion
            sCond = "Rentables.IdUsuario>0"
    End Select
    
    Refresca
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmConsSocios = Nothing
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "KALACLUB"
End Sub


'Cambia el orden en que aparecen los datos en pantalla
Private Sub dgDatos_HeadClick(ByVal ColIndex As Integer)
    nColActiva = ColIndex
    Refresca
End Sub


Private Sub InitdgDatos()
Dim sHead As String

    'Asigna los encabezados de las columnas
    DefHeadersDBGrid Me.dgDatos, mTitCons
    
    'Asigna el ancho de cada columna
    DefAnchoDBGrid Me.dgDatos, mAncCons
    
    'Evita que se puedan modificar los datos de la consulta
    Me.dgDatos.AllowUpdate = False
    
    'Encabezados de los DbGrid
    Select Case nCons
        Case 1
            sHead = "Información de ausencias"
            
        Case 2
            sHead = "Información de autos"
        
        Case 3
            sHead = "Información de certiifcados médicos"
        
        Case 4
            sHead = "Información de domicilios"
            
        Case 5
            sHead = "Información de horarios"
        
        Case 6
            sHead = "Información de pases temporales"
        
        Case 7
            sHead = "Información de rentables"
    End Select
    
    Me.dgDatos.Caption = sHead
End Sub


Public Sub Refresca()
    'Escribe el encabezado del formulario
    frmConsSocios.Caption = sTit & mTitForma(nColActiva)

    'Inicializa un ctrlAdo de los padres con una lista de campos
    InitCtrlAdoSel Me.adoDatos, sTablaCons, mCamCons, nColumns, nColActiva, sCond, Conn
    
    'Configura el ctrl DataGrid
    InitdgDatos
    
    'Quita el efecto que deja la columna en color negro
    Me.dgDatos.ClearSelCols
    
'    If ((Not Me.adoDatos.Recordset.EOF) And (nPos > 0)) Then
'        Me.adoDatos.Recordset.Bookmark = nPos
'    End If
End Sub


Private Sub Form_Resize()
    If ((ScaleWidth - 250) > 0) Then
        Me.dgDatos.Width = ScaleWidth - 250
        Me.cmdSalir.Left = ScaleWidth - 750
        Me.cmdRefresh.Left = ScaleWidth - 1500
    End If

    If ((ScaleHeight - 1100) > 0) Then
        Me.dgDatos.Height = ScaleHeight - 1100
'        Me.lblNotas.Top = ScaleHeight - 210
'        Me.lblRegs.Top = ScaleHeight - 210
'        Me.txtRegs.Top = ScaleHeight - 210
    End If
End Sub

