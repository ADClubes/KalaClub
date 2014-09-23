VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00FFFCFF&
   Caption         =   "KalaClub"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11280
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   5503
            Object.ToolTipText     =   "Proceso Actual"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "07/08/2014"
            Object.ToolTipText     =   "Fecha Actual del Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "01:47 p.m."
            Object.ToolTipText     =   "Hora Actual del Sistema"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "Base de datos"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Caja"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1755
      Top             =   2205
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":0C58
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   1588
      ButtonWidth     =   2090
      ButtonHeight    =   1429
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Usuarios"
            Key             =   "socios"
            Description     =   "Datos"
            Object.ToolTipText     =   "Títulos"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "compra"
                  Object.Tag             =   "1"
                  Text            =   "&Compra"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "traspaso"
                  Text            =   "&Traspaso"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Facturación"
            Key             =   "facturacion"
            Description     =   "Facturación"
            Object.ToolTipText     =   "Reportes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reportes"
            Key             =   "reportes"
            Description     =   "Reportes"
            Object.ToolTipText     =   "Acerca de KalaClub"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&MultiClub"
            Key             =   "MultiClub"
            Description     =   "MultiClub"
            Object.ToolTipText     =   "MultiClub"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salir"
            Key             =   "Salir"
            Description     =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   6480
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuSocios 
      Caption         =   "&Usuarios"
      Begin VB.Menu altas 
         Caption         =   "&Altas"
      End
      Begin VB.Menu mnuDatosSocios 
         Caption         =   "&Datos"
      End
      Begin VB.Menu socios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuSociosProspectos 
         Caption         =   "&Prospectos"
      End
      Begin VB.Menu mnuSociosCotiza 
         Caption         =   "&Cotiza Cuota"
      End
      Begin VB.Menu mnuSociosUsuMC 
         Caption         =   "Usuarios Multiclub"
      End
   End
   Begin VB.Menu mnuFac 
      Caption         =   "&Facturación"
      Begin VB.Menu mnufacPagar 
         Caption         =   "&Pagar"
      End
      Begin VB.Menu mnufacConsDoc 
         Caption         =   "&Consulta Documentos"
      End
      Begin VB.Menu mnufacTurnos 
         Caption         =   "&Turnos"
      End
      Begin VB.Menu mnuFacFacRec 
         Caption         =   "Facturación de recibos"
      End
      Begin VB.Menu mnuFacRepCFD 
         Caption         =   "Reporte CFD"
      End
      Begin VB.Menu Separador 
         Caption         =   "-"
      End
      Begin VB.Menu menuFacCorteCaja 
         Caption         =   "&Corte de caja"
      End
      Begin VB.Menu menuFacCorteCajaVal 
         Caption         =   "&Validación de cortes de caja"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFacPagos 
         Caption         =   "&Validacion Cortes de caja por turno"
      End
      Begin VB.Menu mnufacccdia 
         Caption         =   "Validación de corte por dia"
      End
      Begin VB.Menu Separador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufacSustfac 
         Caption         =   "&Sustitución de facturas"
      End
      Begin VB.Menu mnuFacGenNotaCred 
         Caption         =   "G&eneración de notas de crédito"
      End
      Begin VB.Menu mnufacGenAde 
         Caption         =   "&Genera Adeudos"
      End
      Begin VB.Menu mnuFacEstacionamiento 
         Caption         =   "Calculadora estacionamiento"
      End
      Begin VB.Menu mnuFacCobranza 
         Caption         =   "Cobran&za"
         Begin VB.Menu mnuFacValArc 
            Caption         =   "Valida archivo"
         End
         Begin VB.Menu mnuFacCreaArc 
            Caption         =   "Crea Archivo"
         End
      End
   End
   Begin VB.Menu mnuVentas 
      Caption         =   "&Ventas"
      Begin VB.Menu mnuVentasPreVenta 
         Caption         =   "&PreVenta"
      End
      Begin VB.Menu mnuVentasCatOrigenVenta 
         Caption         =   "Catálogo de origenes de venta"
      End
   End
   Begin VB.Menu mnuCtrlAcceso 
      Caption         =   "C&ontrol de Acceso"
      Begin VB.Menu mnuCtrl 
         Caption         =   "&Actualizar Fechas de Acceso"
         Index           =   1
      End
      Begin VB.Menu mnuCtrlBloq 
         Caption         =   "Bloquear usuarios con adeudo"
      End
      Begin VB.Menu mnuCtrlAct 
         Caption         =   "A&ctiva por código"
      End
      Begin VB.Menu mnuCtrlActMC 
         Caption         =   "&MultiClub"
      End
      Begin VB.Menu mnuCtrlActSQL 
         Caption         =   "&Acceso por consulta"
      End
   End
   Begin VB.Menu mnuGgral 
      Caption         =   "&G. General"
      Begin VB.Menu mnuGgralRPptoMes 
         Caption         =   "&Reporte comparativo presupuesto por mes"
      End
   End
   Begin VB.Menu mnuCupones 
      Caption         =   "&Cupones"
      Begin VB.Menu mnuCuponesAdm 
         Caption         =   "&Administración de cupones"
      End
      Begin VB.Menu mnuCuponesNomina 
         Caption         =   "&Generar Nómina"
      End
   End
   Begin VB.Menu mnuOperaciones 
      Caption         =   "&Operaciones"
      Begin VB.Menu mnuOperacionesLockers 
         Caption         =   "&Lockers"
      End
   End
   Begin VB.Menu mnuUtilerias 
      Caption         =   "&Utilerías"
      Begin VB.Menu mnuCatalogos 
         Caption         =   "&Catálogos"
         Begin VB.Menu mnuCatal 
            Caption         =   "&Bancos"
            Index           =   1
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "Clases y Curs&os"
            Index           =   2
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "&Conceptos de Ingresos"
            Index           =   3
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "&Forma de Pago"
            Index           =   4
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "Histórico de Cuotas"
            Index           =   5
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "&Horarios de Clases"
            Index           =   6
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "&Instructores"
            Index           =   7
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "&Membresias"
            Index           =   8
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "&Países"
            Index           =   9
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "R&eglas Tipos de Usuario"
            Index           =   10
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "&Rentables"
            Index           =   11
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "Tipos de &Usuarios"
            Index           =   12
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "Tipos Re&ntables"
            Index           =   13
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "&Títulos"
            Index           =   14
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "Usuarios del &Sistema"
            Index           =   15
         End
         Begin VB.Menu mnuCatal 
            Caption         =   "&Vendedores"
            Index           =   16
         End
      End
      Begin VB.Menu mnuSPC01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones 
         Caption         =   "&Opciones"
         Begin VB.Menu mnuOpc 
            Caption         =   "&Parámetros Globales"
            Index           =   1
         End
         Begin VB.Menu mnuOpc 
            Caption         =   "&Seguridad"
            Index           =   2
         End
         Begin VB.Menu mnuOpc 
            Caption         =   "&Folios"
            Index           =   3
         End
      End
      Begin VB.Menu SPC02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Procesos"
         Begin VB.Menu mnuProc 
            Caption         =   "Membresias"
            Index           =   1
         End
         Begin VB.Menu mnuProc 
            Caption         =   "Códigos"
            Index           =   2
         End
         Begin VB.Menu mnuProcEjecutarQry 
            Caption         =   "Ejecutar Query"
         End
         Begin VB.Menu mnuProcPruebaServCFD 
            Caption         =   "Prueba Servicio CFD"
         End
      End
      Begin VB.Menu SPC03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivos 
         Caption         =   "&Usuarios Activos"
      End
   End
   Begin VB.Menu mnuVentanas 
      Caption         =   "&Ventanas"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuAcerca 
         Caption         =   "&Acerca de Kala Club"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ************************************************************************
' PANTALLA: PRINCIPAL
' Objetivo: MDI DE LA APLICACIÓN
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit

Private Sub altas_Click()
    
    If Not ChecaSeguridad(Me.Name, "ALTAS") Then
        Exit Sub
    End If


    frmAltaSocios.bSocioNvo = True
    frmAltaSocios.sFormaAnterior = ""
    Load frmAltaSocios
    frmAltaSocios.Show (1)
End Sub



Private Sub Command1_Click()

    'Dim frmNC As frmBulkNC
    
    'Set frmNC = New frmBulkNC
    
    'frmNC.Show vbModal
    

'    Dim adorcs As ADODB.Recordset
'    Dim iDigito As Integer
'
'    Dim aUnidades(6) As String
'
'    Dim dFechaIni As Date
'    Dim dFechaFin As Date
'    Dim dFecha As Date
'
'    Dim lI As Integer
'
'    Dim sCadena As String
'
'    aUnidades(0) = "COY"
'    aUnidades(1) = "SAN"
'    aUnidades(2) = "DLN"
'    aUnidades(3) = "SFE"
'    aUnidades(4) = "ARB"
'    aUnidades(5) = "LVE"
'
'    dFechaIni = DateSerial(2011, 4, 1)
'    dFechaFin = DateSerial(2011, 4, 30)
'
'    strSQL = "SELECT * FROM REFER"
'
'
'    Set adorcs = New ADODB.Recordset
'
'    adorcs.Open strSQL, Conn, adOpenDynamic, adLockOptimistic
'
'    For lI = 0 To 5
'
'        dFecha = dFechaIni
'
'        Do While dFecha <= dFechaFin
'            sCadena = aUnidades(lI) & Format(dFecha, "yymmdd") & "02"
'            iDigito = dvAlgoritmo35(sCadena)
'            adorcs.AddNew
'            adorcs!Unidad = aUnidades(lI)
'            adorcs!Fecha = dFecha
'            adorcs!Referencia = sCadena & Trim(Str(iDigito))
'            adorcs.Update
'            dFecha = dFecha + 1
'        Loop
'
'    Next
'    adorcs.Close
'
'    Set adorcs = Nothing
'
'    MsgBox "Hecho", vbOKOnly, ""
    
End Sub

Private Sub MDIForm_Activate()
    'Habilita_Seguridad Me
End Sub

Private Sub MDIForm_Load()
    'Dim Valor As Byte
    'Dim Bandera As Boolean
    
    'If App.PrevInstance Then
    '    MsgBox "¡ El programa ya está en ejecución !", vbInformation, "¡ Imposible ejecutar !"
    '    Unload Me
    '    End
    '    Exit Sub
    'End If
    
    Me.Height = 8550
    Me.Width = 11950
    Me.Top = 0
    Me.Left = 0
    'Bandera = True
    'Valor = Lee_Ini()
    'Select Case Valor
    '    Case 1  'No existe INI
    '        If Not CREA_INI Then
    '           End
    '        End If
    '     Case 0    ' Existe INI y todo esta bien
    '        If Not Connection_DB() Then
    '            Bandera = False
    '            Unload Me
    '        Else
    '            frmPresentacion.Show vbModal, Me
    '            If LoginOk = True Then
                    MDIPrincipal.Visible = True
                    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "KalaClub"
                    MDIPrincipal.StatusBar1.Panels.Item(4).Text = sDB
                    MDIPrincipal.StatusBar1.Panels.Item(5).Text = sDB_User
                    '29/06/2007
                    MDIPrincipal.StatusBar1.Panels.Item(6).Text = "Caja: " & iNumeroCaja
                    
                    'Seguridad (sDB_NivelUser)
                          
                    '29/12/2008
                    MDIPrincipal.StatusBar1.Panels.Item(4).ToolTipText = sDB_DataSource
                    
    '            Else
    '                Unload Me
    '            End If
    '        End If
    'End Select
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim byResp As Byte
    
    If Not LoginOk Then
        Cancel = 0
        Exit Sub
    End If
    
    
    byResp = MsgBox("¿Salir del programa?", vbQuestion Or vbYesNo, "Salir")
    
    If byResp <> vbYes Then
        Cancel = 1
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim i As Integer
    
    
    For i = 0 To Forms.Count - 1
        Unload Forms(i)
    Next
        Call EndConn_DB
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub

Private Sub menuFacCorteCaja_Click()
    Dim frmCorte As frmCorteCaja
    
    If Not ChecaSeguridad(Me.Name, Me.menuFacCorteCaja.Name) Then
        Exit Sub
    End If
    
    If OpenShiftF() = 0 Then
        MsgBox "No hay turno abierto", vbExclamation, "Error"
        Exit Sub
    End If
    
    Set frmCorte = New frmCorteCaja
    
    frmCorte.Show vbModal
    
End Sub

Private Sub menuFacCorteCajaVal_Click()
    Dim frmCCajaVal As frmCCajaValida
    
    
    
    If Not ChecaSeguridad(Me.Name, Me.menuFacCorteCajaVal.Name) Then
        Exit Sub
    End If
    
    
    Set frmCCajaVal = New frmCCajaValida
    
    frmCCajaVal.Show vbModal
    
End Sub

Private Sub mnuCatalBancos_Click(Index As Integer)

End Sub

Private Sub mnuCtrlAct_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.mnuCtrlAct.Name) Then
        Exit Sub
    End If

    frmActiva.Show vbModal
End Sub

Private Sub mnuCtrlActMC_Click()
    Dim frmMC As frmAccesoMC
    
    
    If Not ChecaSeguridad(Me.Name, Me.mnuCtrlActMC.Name) Then
        Exit Sub
    End If
    
    
    Set frmMC = New frmAccesoMC
    
    
    frmMC.Show vbModal
    
End Sub

Private Sub mnuCtrlActSQL_Click()
    
    
    If Not ChecaSeguridad(Me.Name, Me.mnuCtrlActSQL.Name) Then
        Exit Sub
    End If
    
    Dim frmAccSQL As frmAccListaSQL
    Set frmAccSQL = New frmAccListaSQL
    frmAccSQL.Show vbModal
    
End Sub

Private Sub mnuCuponesAdm_Click()
    If Not ChecaSeguridad(Me.Name, Me.mnuCuponesAdm.Name) Then
        Exit Sub
    End If

    frmCupones.Show vbModal
End Sub

Private Sub mnuCuponesNomina_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.mnuCuponesNomina.Name) Then
        Exit Sub
    End If
    
    frmCuponesNom.Show vbModal
End Sub

Private Sub mnufacccdia_Click()
    '06/12/2011 UCM
    If sDB_NivelUser <> 0 And Not ChecaSeguridad(Me.Name, Me.mnufacccdia.Name) Then
        Exit Sub
    End If
    
    Dim frmCorteDia As frmValidaCorteDia
    
    Set frmCorteDia = New frmValidaCorteDia
    
    frmCorteDia.Show vbModal
    
End Sub

Private Sub mnuFacCreaArc_Click()
    '06/12/2011 UCM
    If sDB_NivelUser <> 0 And Not ChecaSeguridad(Me.Name, Me.mnuFacCreaArc.Name) Then
        Exit Sub
    End If
        
    Dim adorcs As ADODB.Recordset
    
    Dim fs As Object
    Dim outputFile As Object
    Dim sDir As String
    Dim sFileName As String
    Dim sRenglon As String

    strSQL = "SELECT * FROM ENVIOBANAMEX"
    strSQL = strSQL & " ORDER BY CLIENTE"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenKeyset, adLockReadOnly
    
    
    sFileName = "c:\clientes.cli"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set outputFile = fs.CreateTextFile(sFileName)
    
    Do While Not adorcs.EOF
        sRenglon = vbNullString
        sRenglon = sRenglon & Format(Format(adorcs!Cliente, "00000000"), "!@@@@@@@@@@@@@@@@@@@@@@@")
        sRenglon = sRenglon & Format(Left(Trim(adorcs!Nombre), 20), "!@@@@@@@@@@@@@@@@@@@@")
        sRenglon = sRenglon & Format(Left(Trim(adorcs!A_Paterno), 20), "!@@@@@@@@@@@@@@@@@@@@")
        sRenglon = sRenglon & Format(Left(Trim(adorcs!A_Materno), 20), "!@@@@@@@@@@@@@@@@@@@@")
        sRenglon = sRenglon & Format(adorcs!Monto, "00000000") & "00"
        sRenglon = sRenglon & adorcs!Tarjeta
        sRenglon = sRenglon & "31052009"
        sRenglon = sRenglon & "31122009"
        sRenglon = sRenglon & adorcs!Vence
        
        outputFile.WriteLine sRenglon
        adorcs.MoveNext
    Loop
    
    adorcs.Close
    Set adorcs = Nothing
    
    
    MsgBox "Archivo creado", vbInformation, "Ok"
    
End Sub

Private Sub mnuFacEstacionamiento_Click()
    
    Dim frmPark As frmEstacionamiento
    
    If Not ChecaSeguridad(Me.Name, Me.mnuFacEstacionamiento.Name) Then
        Exit Sub
    End If
    
    
    Set frmPark = New frmEstacionamiento
    
    frmPark.Show
    
End Sub

Private Sub mnuFacFacRec_Click()

    If Not ChecaSeguridad(Me.Name, Me.mnuFacFacRec.Name) Then
        Exit Sub
    End If
    
    frmFacturaRec.Show vbModal
    
    
End Sub

Private Sub mnuFacGenNotaCred_Click()
    '06/12/2011 UCM
    If sDB_NivelUser <> 0 And Not ChecaSeguridad(Me.Name, Me.mnuFacGenNotaCred.Name) Then
        Exit Sub
    End If
    
    Dim frmSF As frmSustFac
    
    Set frmSF = New frmSustFac
    
    frmSF.iModo = 1
    
    frmSF.Show vbModal
End Sub

Private Sub mnuFacPagos_Click()
    
    Dim frmPagos As frmFormaPagoMod
    
    
    If Not ChecaSeguridad(Me.Name, Me.mnufacPagar.Name) Then
        Exit Sub
    End If
    
    
    Set frmPagos = New frmFormaPagoMod
    
    frmPagos.Show vbModal
    
    
End Sub

Private Sub mnuFacRepCFD_Click()
    Dim frmRep As frmReporteCFD
    
    If Not ChecaSeguridad(Me.Name, Me.mnuFacRepCFD.Name) Then
        Exit Sub
    End If
    
    Set frmRep = New frmReporteCFD
    
    
    frmRep.Show vbModal
    
End Sub

Private Sub mnufacSustfac_Click()
    '06/12/2011 UCM
    If sDB_NivelUser <> 0 And Not ChecaSeguridad(Me.Name, Me.mnufacSustfac.Name) Then
        Exit Sub
    End If
    
    Dim frmSF As frmSustFac
    
    Set frmSF = New frmSustFac
    
    frmSF.iModo = 0
    
    
    frmSF.Show vbModal
    
End Sub

Private Sub mnufacTurnos_Click()
    If Not ChecaSeguridad(Me.Name, Me.mnufacTurnos.Name) Then
        Exit Sub
    End If
    
    frmTurnos.Show vbModal
End Sub




Private Sub mnuAcerca_Click()
    frmAcerca.Show
End Sub

Private Sub mnuActivos_Click()
    Load FrmUsuaActiv
    FrmUsuaActiv.Show
End Sub





Private Sub mnuCatal_Click(Index As Integer)
    Unload frmSelecReportes
    

    
    If Not ChecaSeguridad(Me.Name, "mnuCatal_" & Format(Index, "00")) Then
        Exit Sub
    End If
    
    
    Select Case Index
        Case 1
            StatusBar1.Panels.Item(1).Text = "BANCOS"
        Case 2
            StatusBar1.Panels.Item(1).Text = "TIPOS DE CLASES"
        Case 3
            StatusBar1.Panels.Item(1).Text = "CONCEPTOS DE INGRESOS"
        Case 4
            StatusBar1.Panels.Item(1).Text = "PAGOS"
        Case 5
            StatusBar1.Panels.Item(1).Text = "HISTORICO DE CUOTAS"
        Case 6
            StatusBar1.Panels.Item(1).Text = "HORARIOS DE CLASES"
        Case 7
            StatusBar1.Panels.Item(1).Text = "INSTRUCTORES"
        Case 8
            StatusBar1.Panels.Item(1).Text = "MEMBRESÍAS"
        Case 9
            StatusBar1.Panels.Item(1).Text = "PAISES"
        Case 10
            StatusBar1.Panels.Item(1).Text = "REGLAS TIPO USUARIO"
        Case 11
            StatusBar1.Panels.Item(1).Text = "RENTABLES"
        Case 12
            StatusBar1.Panels.Item(1).Text = "TIPO USUARIO"
        Case 13
            StatusBar1.Panels.Item(1).Text = "TIPO RENTABLES"
        Case 14
            StatusBar1.Panels.Item(1).Text = "TITULOS"
        Case 15
            StatusBar1.Panels.Item(1).Text = "USUARIOS DEL SISTEMA"
    End Select
    
    
    Select Case Index
        Case 5
            frmHistoricoCuotas.Show
        Case 16
             Load frmVendedores
            frmVendedores.Show
    Case Else
            Load frmCatalogos
            frmCatalogos.Show
    End Select
    
End Sub



Private Sub mnuCtrl_Click(Index As Integer)
    
    Dim frmFechas As frmFechasAcceso
    
    If Not ChecaSeguridad(Me.Name, Me.mnuCtrlAcceso.Name) Then
        Exit Sub
    End If
    
    Set frmFechas = New frmFechasAcceso
    
    frmFechas.Show vbModal
    
End Sub

Private Sub mnuCtrlBloq_Click()
    If Not ChecaSeguridad(Me.Name, mnuCtrlBloq.Name) Then
        Exit Sub
    End If
    
    frmBloqueaMorosos.Show vbModal
    
End Sub

Private Sub mnuDatosSocios_Click()
    
    If Not ChecaSeguridad(Me.Name, mnuDatosSocios.Name) Then
        Exit Sub
    End If

    frmDatosSocios.nColActiva = 2
    frmDatosSocios.Show
    
End Sub



Private Sub mnufacConsDoc_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.mnufacConsDoc.Name) Then
        Exit Sub
    End If
    
    frmConsDocs.Show
    
End Sub

Private Sub mnufacGenAde_Click()
    
    '06/12/2011 UCM
    If sDB_NivelUser <> 0 And Not ChecaSeguridad(Me.Name, Me.mnufacGenAde.Name) Then
        Exit Sub
    End If
    
    FrmGeneraAdeudos.Show 1
    
End Sub

Private Sub mnufacPagar_Click()
    If Not ChecaSeguridad(Me.Name, "FACTURACION") Then
        Exit Sub
    End If

    frmFacturacion.Show
End Sub



Private Sub mnuFacValArc_Click()
    '06/12/2011 UCM
    If sDB_NivelUser <> 0 And Not ChecaSeguridad(Me.Name, Me.mnuFacValArc.Name) Then
        Exit Sub
    End If
    
    Dim adorcs As ADODB.Recordset
    
    Dim sRespuesta As String
    
    Dim fs As FileSystemObject
    Dim outfile As TextStream

    strSQL = "SELECT * FROM EnvioBanamex"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenKeyset, adLockReadOnly
    
    Set fs = New FileSystemObject
    
    Set outfile = fs.CreateTextFile("c:\output.txt")
    
    Do While Not adorcs.EOF
        
        If Not ValidateCardNumber(adorcs!Tarjeta, sRespuesta) Then
            MsgBox "Error " & adorcs!Cliente & " " & sRespuesta
            outfile.WriteLine adorcs!Cliente & "," & sRespuesta
        End If
        
        If Not ValidateCardName(adorcs!Nombre & adorcs!A_Paterno & adorcs!A_Materno, sRespuesta) Then
            MsgBox "Error " & adorcs!Cliente & " " & sRespuesta
            outfile.WriteLine adorcs!Cliente & "," & sRespuesta
        End If
        
        If (Val(Mid$(adorcs!Vence, 1, 2)) + Val(Mid$(adorcs!Vence, 3, 2)) * 12) < (Month(Date) + (Year(Date) - 2000) * 12) Then
            MsgBox "Error " & adorcs!Cliente & " " & "Tarjeta vencida " & adorcs!Vence
            outfile.WriteLine adorcs!Cliente & "," & "Tarjeta vencida " & adorcs!Vence
        End If
        adorcs.MoveNext
    Loop
    
    
    adorcs.Close
    Set adorcs = Nothing
    
    outfile.Close
    
    Set outfile = Nothing
    Set fs = Nothing
    
    
    MsgBox "Finalizado"
End Sub



Private Sub mnuGgralRPptoMes_Click()
    Dim frmRep As frmRepPptoMes
    
    
    If Not ChecaSeguridad(Me.Name, mnuGgralRPptoMes.Name) Then
        Exit Sub
    End If
    
    
    
    Set frmRep = New frmRepPptoMes
    
    
    frmRep.Show vbModal
    
End Sub

Private Sub mnuOpc_Click(Index As Integer)
    
     If Not ChecaSeguridad(Me.Name, "mnuOpc_" & Format(Index, "00")) Then
        Exit Sub
    End If


    Select Case Index
        Case 1
            Load frmOpciones
            frmOpciones.Show
        Case 2
            Load frmSeguridad
            frmSeguridad.Show
        Case 3
             frmCambiaFolio.Show vbModal
    End Select
End Sub





Private Sub mnuOperacionesLockers_Click()
    Dim frmLockers As frmOperLockers
    
    If Not ChecaSeguridad(Me.Name, Me.mnuOperacionesLockers.Name) Then
        Exit Sub
    End If
    
    
    Set frmLockers = New frmOperLockers
    
    
    frmLockers.Show vbModal
    
End Sub





Private Sub mnuProc_Click(Index As Integer)
     If Not ChecaSeguridad(Me.Name, "mnuProc_" & Format(Index, "00")) Then
        Exit Sub
    End If


    Select Case Index
        Case 1
            frmProcMem.Show vbModal
        Case 2
            frmProcCredencial.Show vbModal
    End Select
End Sub

Private Sub mnuProcEjecutarQry_Click()
    Dim frmExec As frmSQLExec
    
    Set frmExec = New frmSQLExec
    
    frmExec.Show vbModal
    
    
End Sub

Private Sub mnuProcPruebaServCFD_Click()
    Dim sRespuesta As String
    
    
    sRespuesta = PruebaCFD()
    
    MsgBox "Respuesta recibida:" + vbCrLf + sRespuesta, vbInformation, "Resultado"
    
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub


Private Sub mnuSociosCotiza_Click()
    If Not ChecaSeguridad(Me.Name, mnuSociosCotiza.Name) Then
        Exit Sub
    End If
    
    frmCotiza.Show vbModal
    
End Sub



Private Sub mnuSociosProspectos_Click()
    If Not ChecaSeguridad(Me.Name, mnuSociosProspectos.Name) Then
        Exit Sub
    End If
    frmProspectos.Show vbModal
End Sub



Private Sub mnuSociosUsuMC_Click()
    Dim frmMC As frmUsuariosMC
    
    
    If Not ChecaSeguridad(Me.Name, mnuSociosUsuMC.Name) Then
        Exit Sub
    End If
    
    
    
    Set frmMC = New frmUsuariosMC
    
    frmMC.Show vbModal
    
End Sub

Private Sub mnuVentasCatOrigenVenta_Click()
    Dim frmCtOriVen As frmCtOrigenVenta
    
    Set frmCtOriVen = New frmCtOrigenVenta
    
    
    
    frmCtOriVen.Show vbModal
    
    
End Sub

Private Sub mnuVentasPreVenta_Click()
    
    Dim frmPV As frmPreVenta
    
    If Not ChecaSeguridad(Me.Name, mnuVentasPreVenta.Name) Then
        Exit Sub
    End If
    
    Set frmPV = New frmPreVenta
    
    frmPV.Show
    
End Sub

Private Sub socios_Click()
    Load frmSocios
    frmSocios.Show
End Sub




Private Sub StatusBar1_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
'    Dim Resp As VbMsgBoxResult
'
'    Dim sForma As String
'    Dim sNombreObjeto As String
'
'
'    Dim adorcsSeg As ADODB.Recordset
'    Dim adocmdSeg As ADODB.Command
'
'    If Panel.index <> 7 Then
'        Exit Sub
'    End If
'
'    If Left(Panel, 11) <> "Seguridad: " Then
'        Exit Sub
'    End If
'
'    Resp = MsgBox("¿Proceder?", vbQuestion + vbOKCancel, "Confirme")
'
'
'    If Resp = vbCancel Then
'        Exit Sub
'    End If
'
'
'    sForma = Mid$(Panel, 12, InStr(Panel, ",") - 1 - 11)
'    sNombreObjeto = Mid$(Panel, InStr(Panel, ",") + 1)
'
'
'
'    Set adocmdSeg = New ADODB.Command
'    adocmdSeg.ActiveConnection = Conn
'    adocmdSeg.CommandType = adCmdText
'
'
'    Set adorcsSeg = New ADODB.Recordset
'
'    adorcsSeg.CursorLocation = adUseServer
'
'
'    strSQL = "SELECT IdObjeto"
'    strSQL = strSQL & " FROM CT_Objetos_Seguridad"
'    strSQL = strSQL & " WHERE ("
'    strSQL = strSQL & " ((FormaNombre)='" & UCase(sForma) & "')"
'    strSQL = strSQL & " AND ((ObjetoNombre)='" & UCase(sNombreObjeto) & "')"
'    strSQL = strSQL & ")"
'
'
'    adorcsSeg.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
'
'    If adorcsSeg.EOF Then
'
'        adorcsSeg.Close
'
'
'        strSQL = "INSERT INTO CT_Objetos_Seguridad ("
'        strSQL = strSQL & "FormaNombre,"
'        strSQL = strSQL & "ObjetoNombre)"
'        strSQL = strSQL & " VALUES ("
'        strSQL = strSQL & "'" & sForma & "',"
'        strSQL = strSQL & "'" & sNombreObjeto & "')"
'
'        adocmdSeg.CommandText = strSQL
'        adocmdSeg.Execute
'
'
'    End If
'
'
'    strSQL = "INSERT INTO SEGURIDAD_USUARIO "
'    strSQL = strSQL & " SELECT "
'    strSQL = strSQL & iDB_IdUser & " AS IdUsuario,"
'    strSQL = strSQL & " IdObjeto"
'    strSQL = strSQL & " FROM"
'    strSQL = strSQL & " CT_Objetos_Seguridad"
'    strSQL = strSQL & " WHERE ("
'    strSQL = strSQL & " ((FormaNombre)='" & UCase(sForma) & "')"
'    strSQL = strSQL & " AND ((ObjetoNombre)='" & UCase(sNombreObjeto) & "')"
'    strSQL = strSQL & ")"
'
'    On Error GoTo Error_Catch
'
'    adocmdSeg.CommandText = strSQL
'    adocmdSeg.Execute
'
'    On Error GoTo 0
'
'
'    Set adorcsSeg = Nothing
'    Set adocmdSeg = Nothing
'
'
'    MDIPrincipal.StatusBar1.Panels(7).Text = ""
'
'    Exit Sub
'
'
'Error_Catch:
'
'    MsgError
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case "SOCIOS"
            If ExistWindow("frmSocios") Then
                frmSocios.SetFocus
                If frmSocios.WindowState = 1 Then
                    frmSocios.WindowState = 0
                End If
            Else
                frmSocios.Show
            End If
        Case "FACTURACION"
            If Not ChecaSeguridad(Me.Name, "FACTURACION") Then
                Exit Sub
            End If
            
            If ExistWindow("frmFacturacion") Then
                frmFacturacion.SetFocus
                If frmFacturacion.WindowState = 1 Then
                    frmFacturacion.WindowState = 0
                End If
            Else
                frmFacturacion.Show
            End If
        Case "ACERCADE"
            Load frmAcerca
            frmAcerca.Show
        Case "REPORTES"
            If Not ChecaSeguridad(Me.Name, "REPORTES") Then
                Exit Sub
            End If
            Load frmSelecReportes
            frmSelecReportes.Show
        Case "SALIR"
            Unload Me
        Case "MULTICLUB"
            If ExistWindow("frmUsuariosMC") Then
                frmUsuariosMC.SetFocus
                If frmUsuariosMC.WindowState = 1 Then
                    frmUsuariosMC.WindowState = 0
                End If
            Else
                frmUsuariosMC.Show
            End If
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case UCase(ButtonMenu.Key)
        Case "COMPRA"
            Load frmCompra
            frmCompra.Show
        Case "TRASPASO"
            Load frmTraspaso
            frmTraspaso.Show
    End Select
End Sub

Private Sub vendedores_Click()
    Load frmVendedores
    frmVendedores.Show
End Sub


Private Sub Seguridad(iNivel As Integer)
    
    If iNivel = 0 Then
        Exit Sub
    End If
    
    
    Me.mnuSocios.Enabled = False
    Me.mnuFac.Enabled = False
    Me.mnuVentas.Enabled = False
    Me.mnuCtrlAcceso.Enabled = False
    Me.mnuCupones.Enabled = False
    Me.mnuUtilerias.Enabled = False
    
    Select Case iNivel
        
        Case 1 'Consulta socios y puede modificarlos y facturar
            Me.mnuSocios.Enabled = True
            Me.mnuFac.Enabled = True
        Case 2 'Consulta socios y puede modificarlos
            Me.mnuSocios.Enabled = True
            Me.mnuDatosSocios.Enabled = False
            Me.Toolbar1.Buttons(2).Enabled = False
        Case 3 'Puede consulta socios pero no puede modificarlos
            Me.Toolbar1.Buttons(2).Enabled = False
        Case 4 'Totalmente restringido
            Me.Toolbar1.Buttons(1).Enabled = False
            Me.Toolbar1.Buttons(2).Enabled = False
    End Select
End Sub
