VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmOtrosDatos 
   Caption         =   "Datos adicionales"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "frmOtrosDatos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTitCve 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   5
      TabIndex        =   11
      Top             =   300
      Width           =   615
   End
   Begin VB.TextBox txtFamilia 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      MaxLength       =   5
      TabIndex        =   9
      Top             =   300
      Width           =   975
   End
   Begin VB.TextBox txtTitular 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   300
      Width           =   5295
   End
   Begin TabDlg.SSTab sstabOtrosDatos 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Pases temporales"
      TabPicture(0)   =   "frmOtrosDatos.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdModPase"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDelPase"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAddPase"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ssdbPases"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Autos"
      TabPicture(1)   =   "frmOtrosDatos.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ssdbAutos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdDelAuto"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdModAuto"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdAddAuto"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Certificados médicos"
      TabPicture(2)   =   "frmOtrosDatos.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdAddExamen"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdModExamen"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdDelExamen"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtObsCert"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "ssdbExamMed"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Zonas"
      TabPicture(3)   =   "frmOtrosDatos.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdRefresca"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdActiva"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdDesactiva"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdAddTZone"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdModTZone"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdDelTZone"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "ssdbTZone"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Referencias"
      TabPicture(4)   =   "frmOtrosDatos.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdQuita3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "cmdQuita2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdQuita1"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtNoReg3"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "txtNoReg2"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtNoReg1"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "cmdCancelRefs"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "cmdGuardaRefs"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "txtTel3"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "txtAMaterno3"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "txtAPaterno3"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "txtNombre3"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "txtTel2"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "txtAMaterno2"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "txtAPaterno2"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "txtNombre2"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "txtTel1"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "txtAMaterno1"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "txtAPaterno1"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "txtNombre1"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "lblTel"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "lblAMaterno"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "lblAPaterno"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "lblNombre"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "lblTres"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "lblDos"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "lblUno"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).ControlCount=   27
      Begin VB.CommandButton cmdQuita3 
         Caption         =   "Quitar"
         Height          =   255
         Left            =   -72240
         TabIndex        =   41
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton cmdQuita2 
         Caption         =   "Quitar"
         Height          =   255
         Left            =   -72240
         TabIndex        =   36
         Top             =   3240
         Width           =   615
      End
      Begin VB.CommandButton cmdQuita1 
         Caption         =   "Quitar"
         Height          =   255
         Left            =   -72240
         TabIndex        =   31
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtNoReg3 
         Height          =   285
         Left            =   -68280
         TabIndex        =   46
         Top             =   4680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNoReg2 
         Height          =   285
         Left            =   -68280
         TabIndex        =   45
         Top             =   3240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtNoReg1 
         Height          =   285
         Left            =   -68280
         TabIndex        =   44
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelRefs 
         Height          =   615
         Left            =   -66120
         Picture         =   "frmOtrosDatos.frx":04CE
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "  Cancelar  "
         Top             =   4920
         Width           =   615
      End
      Begin VB.CommandButton cmdGuardaRefs 
         Height          =   615
         Left            =   -66960
         Picture         =   "frmOtrosDatos.frx":0A58
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "  Guardar referencias  "
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox txtTel3 
         Height          =   285
         Left            =   -74280
         TabIndex        =   40
         Top             =   4680
         Width           =   1935
      End
      Begin VB.TextBox txtAMaterno3 
         Height          =   285
         Left            =   -68280
         TabIndex        =   39
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtAPaterno3 
         Height          =   285
         Left            =   -71280
         TabIndex        =   38
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtNombre3 
         Height          =   285
         Left            =   -74280
         TabIndex        =   37
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtTel2 
         Height          =   285
         Left            =   -74280
         TabIndex        =   35
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox txtAMaterno2 
         Height          =   285
         Left            =   -68280
         TabIndex        =   34
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtAPaterno2 
         Height          =   285
         Left            =   -71280
         TabIndex        =   33
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtNombre2 
         Height          =   285
         Left            =   -74280
         TabIndex        =   32
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtTel1 
         Height          =   285
         Left            =   -74280
         TabIndex        =   30
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtAMaterno1 
         Height          =   285
         Left            =   -68280
         TabIndex        =   29
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtAPaterno1 
         Height          =   285
         Left            =   -71280
         TabIndex        =   28
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtNombre1 
         Height          =   285
         Left            =   -74280
         TabIndex        =   27
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton cmdRefresca 
         Height          =   615
         Left            =   -70200
         Picture         =   "frmOtrosDatos.frx":0E9A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   " Refrescar datos  "
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton cmdActiva 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -68520
         Picture         =   "frmOtrosDatos.frx":11A4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   " Activar en la zona elegida "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdDesactiva 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -67680
         Picture         =   "frmOtrosDatos.frx":14AE
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   " Desactiva de la zona elegida "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdAddTZone 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -69360
         Picture         =   "frmOtrosDatos.frx":17B8
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   " Agregar usuario a una zona "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdModTZone 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -66840
         Picture         =   "frmOtrosDatos.frx":2082
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   " Modificar datos "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdDelTZone 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -66000
         Picture         =   "frmOtrosDatos.frx":24C4
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   " Borrar "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdAddExamen 
         Height          =   615
         Left            =   -67680
         Picture         =   "frmOtrosDatos.frx":27CE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   " Agregar certificado médico "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdModExamen 
         Height          =   615
         Left            =   -66840
         Picture         =   "frmOtrosDatos.frx":3098
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   " Modificar "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdDelExamen 
         Height          =   615
         Left            =   -66000
         Picture         =   "frmOtrosDatos.frx":34DA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   " Borrar "
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtObsCert 
         Enabled         =   0   'False
         Height          =   735
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   5040
         Width           =   9375
      End
      Begin SSDataWidgets_B.SSDBGrid ssdbAutos 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   14
         Top             =   1320
         Width           =   9375
         _Version        =   196616
         DataMode        =   2
         Cols            =   5
         Col.Count       =   5
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   16536
         _ExtentY        =   7858
         _StockProps     =   79
         Caption         =   "Información de los autos del socio"
      End
      Begin SSDataWidgets_B.SSDBGrid ssdbPases 
         Height          =   4455
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   9375
         _Version        =   196616
         DataMode        =   2
         Cols            =   9
         Col.Count       =   9
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   16536
         _ExtentY        =   7858
         _StockProps     =   79
         Caption         =   "Pases temporales"
      End
      Begin VB.CommandButton cmdDelAuto 
         Height          =   615
         Left            =   -66000
         Picture         =   "frmOtrosDatos.frx":37E4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Borrar "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdModAuto 
         Height          =   615
         Left            =   -66840
         Picture         =   "frmOtrosDatos.frx":3AEE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Modificar datos "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdAddAuto 
         Height          =   615
         Left            =   -67680
         Picture         =   "frmOtrosDatos.frx":3F30
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " Agregar autos "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdAddPase 
         Height          =   615
         Left            =   7320
         Picture         =   "frmOtrosDatos.frx":47FA
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "  Agregar pase  "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdDelPase 
         Height          =   615
         Left            =   9000
         Picture         =   "frmOtrosDatos.frx":50C4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "  Borrar  "
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdModPase 
         Height          =   615
         Left            =   8160
         Picture         =   "frmOtrosDatos.frx":53CE
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "  Modificar datos del pase  "
         Top             =   600
         Width           =   615
      End
      Begin SSDataWidgets_B.SSDBGrid ssdbExamMed 
         Height          =   3615
         Left            =   -74760
         TabIndex        =   16
         Top             =   1320
         Width           =   9375
         _Version        =   196616
         DataMode        =   2
         Cols            =   13
         Col.Count       =   13
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   16536
         _ExtentY        =   6376
         _StockProps     =   79
         Caption         =   "Certificados médicos entregados"
      End
      Begin SSDataWidgets_B.SSDBGrid ssdbTZone 
         Height          =   4455
         Left            =   -74760
         TabIndex        =   20
         Top             =   1320
         Width           =   9375
         _Version        =   196616
         DataMode        =   2
         Cols            =   9
         Col.Count       =   9
         AllowUpdate     =   0   'False
         AllowRowSizing  =   0   'False
         AllowGroupSizing=   0   'False
         AllowColumnSizing=   0   'False
         AllowGroupMoving=   0   'False
         AllowGroupSwapping=   0   'False
         AllowGroupShrinking=   0   'False
         AllowColumnShrinking=   0   'False
         AllowDragDrop   =   0   'False
         BackColorOdd    =   14737632
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   16536
         _ExtentY        =   7858
         _StockProps     =   79
         Caption         =   "Información de zonas u horarios"
      End
      Begin VB.Label lblTel 
         Caption         =   "Teléfono"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74280
         TabIndex        =   53
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblAMaterno 
         Caption         =   "Apellido materno"
         Height          =   255
         Left            =   -68280
         TabIndex        =   52
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblAPaterno 
         Caption         =   "Apellido paterno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -71280
         TabIndex        =   51
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre (s)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -74280
         TabIndex        =   50
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblTres 
         Caption         =   "3."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   49
         Top             =   4080
         Width           =   255
      End
      Begin VB.Label lblDos 
         Caption         =   "2."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   48
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label lblUno 
         Caption         =   "1."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   47
         Top             =   1200
         Width           =   255
      End
   End
   Begin VB.Label lblTitCve 
      Caption         =   "# Reg."
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   615
   End
   Begin VB.Label lblFamilia 
      Caption         =   "# Familia"
      Height          =   255
      Left            =   6600
      TabIndex        =   10
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblTitular 
      Caption         =   "Titular"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   60
      Width           =   2775
   End
End
Attribute VB_Name = "frmOtrosDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*  Formulario para capturar de los datos adicionales               *
'*  Daniel Hdez                                                     *
'*  11 / Agosto / 2005                                              *
'*  Ultima actualización: 10 / Septiembre / 2005                    *
'********************************************************************


Dim sTextToolBar As String

Dim sNombre1 As String
Dim sAPaterno1 As String
Dim sAMaterno1 As String
Dim sTel1 As String
Dim bRef1 As Boolean

Dim sNombre2 As String
Dim sAPaterno2 As String
Dim sAMaterno2 As String
Dim sTel2 As String
Dim bRef2 As Boolean

Dim sNombre3 As String
Dim sAPaterno3 As String
Dim sAMaterno3 As String
Dim sTel3 As String
Dim bRef3 As Boolean



Private Sub cmdActiva_Click()
    If (Me.ssdbTZone.Rows > 0) Then
        ActivaCred 0, Me.ssdbTZone.Columns(2).Value, Me.ssdbTZone.Columns(6).Value, Me.ssdbTZone.Columns(2).Value, True, False
    End If

'    With Me.adoTimeZone.Recordset
'        ActivaCred 0, .Fields("Secuencial!Secuencial"), .Fields("Time_Zone_Users!IdTimeZone"), .Fields("Time_Zone_Users!IdMember"), True
'    End With
End Sub


Private Sub cmdAddAuto_Click()
    frmAutos.bNvaCal = True
    frmAutos.Show (1)
    
    frmOtrosDatos.ssdbAutos.SetFocus
End Sub


Private Sub cmdAddExamen_Click()
    frmCertificados.bNvoCertif = True
    frmCertificados.Show (1)
    
    frmOtrosDatos.ssdbExamMed.SetFocus
End Sub


Private Sub cmdAddPase_Click()
    frmPases.bNvoPase = True
    frmPases.Show (1)
    
    frmOtrosDatos.ssdbPases.SetFocus
End Sub


Private Sub cmdAddTZone_Click()
    frmTZoneUsers.bNvaZona = True
    frmTZoneUsers.Show (1)
    
    frmOtrosDatos.ssdbTZone.SetFocus
End Sub


Private Sub cmdCancelRefs_Click()
    Me.txtNombre1.Text = sNombre1
    Me.txtAPaterno1.Text = sAPaterno1
    Me.txtAMaterno1.Text = sAMaterno1
    Me.txtTel1.Text = sTel1
    
    Me.txtNombre2.Text = sNombre2
    Me.txtAPaterno2.Text = sAPaterno2
    Me.txtAMaterno2.Text = sAMaterno2
    Me.txtTel2.Text = sTel2
    
    Me.txtNombre3.Text = sNombre3
    Me.txtAPaterno3.Text = sAPaterno3
    Me.txtAMaterno3.Text = sAMaterno3
    Me.txtTel3.Text = sTel3
End Sub


Private Sub cmdDelAuto_Click()
Dim nAnswer As Integer

    If (Me.ssdbAutos.Rows > 0) Then
        nAnswer = MsgBox("¿Desea borrar los datos del auto?", vbYesNo, "Registro de autos")
        
        If (nAnswer = vbYes) Then
            If (EliminaReg("Calcomanias", "Id=" & Me.ssdbAutos.Columns(0).Value, "", Conn)) Then
                frmAutos.LlenaAutos
            End If
        End If
    End If
    
    frmOtrosDatos.ssdbAutos.SetFocus
End Sub


Private Sub cmdDelExamen_Click()
Dim nAnswer As Integer

    If (Me.ssdbExamMed.Rows > 0) Then
        nAnswer = MsgBox("¿Desea borrar los datos del certificado?", vbYesNo, "Registro de certificados")
        
        If (nAnswer = vbYes) Then
            If (EliminaReg("Certificados", "idCertificado=" & Me.ssdbExamMed.Columns(9).Value, "", Conn)) Then
                frmCertificados.LlenaCertificados
            End If
        End If
    End If
    
    frmOtrosDatos.ssdbExamMed.SetFocus
End Sub


Private Sub cmdDelPase_Click()
    If (Me.ssdbPases.Rows > 0) Then
        frmPases.QuitaPase
        frmPases.LlenaPases
        
        frmOtrosDatos.ssdbPases.SetFocus
    End If
End Sub


Private Sub cmdDelTZone_Click()
Dim nAnswer As Integer

    If (Me.ssdbTZone.Rows > 0) Then
        nAnswer = MsgBox("¿Desea borrar el acceso a la zona?", vbYesNo, "Registro de zonas")
        
        If (nAnswer = vbYes) Then
            With Me.ssdbTZone
                If (EliminaReg("Time_Zone_Users", "idReg=" & .Columns(0).Value, "", Conn)) Then
                    'Desactiva la credencial de la zona seleccionada
                    ActivaCred 0, Me.ssdbTZone.Columns(2).Value, Me.ssdbTZone.Columns(6).Value, Me.ssdbTZone.Columns(1).Value, True, False
'                    ActivaCred 0, .Fields("Secuencial!Secuencial"), .Fields("Time_Zone_Users!IdTimeZone"), .Fields("Time_Zone_Users!IdMember"), False

                    frmTZoneUsers.LlenaTZoneUsers
                End If
            End With
        End If
    End If
    
    Me.ssdbTZone.SetFocus
End Sub


Private Sub cmdDesactiva_Click()
    If (Me.ssdbTZone.Rows > 0) Then
        ActivaCred 0, Me.ssdbTZone.Columns(2).Value, Me.ssdbTZone.Columns(6).Value, Me.ssdbTZone.Columns(1).Value, True, False
'        ActivaCred 0, .Fields("Secuencial!Secuencial"), .Fields("Time_Zone_Users!IdTimeZone"), .Fields("Time_Zone_Users!IdMember"), False
    End If
    
    Me.ssdbTZone.SetFocus
End Sub


Private Sub cmdGuardaRefs_Click()
Dim i As Byte

    For i = 1 To 3
        If (CambioRefs(i)) Then
            If (Not GuardaRefs(i)) Then
                MsgBox "No se realizaron los cambios, ver referencia: " & i, vbExclamation, "KalaSystems"
                Exit For
            End If
        End If
    Next i
End Sub


Private Function CambioRefs(nRef As Byte) As Boolean
    CambioRefs = True
    
    Select Case nRef
        
        Case 1
            bRef1 = True
        
            If (sNombre1 <> Trim$(Me.txtNombre1.Text)) Then
                Exit Function
            End If
            
            If (sAPaterno1 <> Trim$(Me.txtAPaterno1.Text)) Then
                Exit Function
            End If
            
            If (sAMaterno1 <> Trim$(Me.txtAMaterno1.Text)) Then
                Exit Function
            End If
            
            If (sTel1 <> Trim$(Me.txtTel1.Text)) Then
                Exit Function
            End If
            
            bRef1 = False
    
        Case 2
            bRef2 = True
    
            If (sNombre2 <> Trim$(Me.txtNombre2.Text)) Then
                Exit Function
            End If
            
            If (sAPaterno2 <> Trim$(Me.txtAPaterno2.Text)) Then
                Exit Function
            End If
            
            If (sAMaterno2 <> Trim$(Me.txtAMaterno2.Text)) Then
                Exit Function
            End If
            
            If (sTel2 <> Trim$(Me.txtTel2.Text)) Then
                Exit Function
            End If
            
            bRef2 = False
    
        Case 3
            bRef3 = True
    
            If (sNombre3 <> Trim$(Me.txtNombre3.Text)) Then
                Exit Function
            End If
            
            If (sAPaterno3 <> Trim$(Me.txtAPaterno3.Text)) Then
                Exit Function
            End If
            
            If (sAMaterno3 <> Trim$(Me.txtAMaterno3.Text)) Then
                Exit Function
            End If
            
            If (sTel3 <> Trim$(Me.txtTel3.Text)) Then
                Exit Function
            End If
            
            bRef3 = False
    End Select
    
    CambioRefs = False
End Function


Private Function GuardaRefs(nRef As Byte) As Boolean
Const DATOSREF = 6
Dim i As Byte
Dim mFieldsRef(DATOSREF) As String
Dim mValuesRef(DATOSREF) As Variant

    If (RefCorrectas(nRef)) Then
        GuardaRefs = False
        
        mFieldsRef(0) = "idRef"
        mFieldsRef(1) = "idMember"
        mFieldsRef(2) = "Nombre"
        mFieldsRef(3) = "A_Paterno"
        mFieldsRef(4) = "A_Materno"
        mFieldsRef(5) = "Telefono"
        
        Select Case nRef
            Case 1
                mValuesRef(0) = LeeUltReg("Referencias", "idRef") + 1
                mValuesRef(1) = Val(Me.txtTitCve.Text)
                mValuesRef(2) = UCase$(Trim$(Me.txtNombre1.Text))
                mValuesRef(3) = UCase$(Trim$(Me.txtAPaterno1.Text))
                mValuesRef(4) = UCase$(Trim$(Me.txtAMaterno1.Text))
                mValuesRef(5) = UCase$(Trim$(Me.txtTel1.Text))
            
            Case 2
                mValuesRef(0) = LeeUltReg("Referencias", "idRef") + 1
                mValuesRef(1) = Val(Me.txtTitCve.Text)
                mValuesRef(2) = UCase$(Trim$(Me.txtNombre2.Text))
                mValuesRef(3) = UCase$(Trim$(Me.txtAPaterno2.Text))
                mValuesRef(4) = UCase$(Trim$(Me.txtAMaterno2.Text))
                mValuesRef(5) = UCase$(Trim$(Me.txtTel2.Text))
            
            Case 3
                mValuesRef(0) = LeeUltReg("Referencias", "idRef") + 1
                mValuesRef(1) = Val(Me.txtTitCve.Text)
                mValuesRef(2) = UCase$(Trim$(Me.txtNombre3.Text))
                mValuesRef(3) = UCase$(Trim$(Me.txtAPaterno3.Text))
                mValuesRef(4) = UCase$(Trim$(Me.txtAMaterno3.Text))
                mValuesRef(5) = UCase$(Trim$(Me.txtTel3.Text))
        End Select
        
        If (AgregaRegistro("Referencias", mFieldsRef, DATOSREF, mValuesRef, Conn)) Then
            Select Case nRef
                Case 1
                    Me.txtNoReg1.Text = mValuesRef(0)
                    
                Case 2
                    Me.txtNoReg2.Text = mValuesRef(0)
                
                Case 3
                    Me.txtNoReg3.Text = mValuesRef(0)
            End Select
            
            AMayusculas
            
            InitVarRefs nRef
            
            GuardaRefs = True
        End If
    End If
End Function


Private Function RefCorrectas(nNoRef As Byte) As Boolean
    RefCorrectas = False
    
    Select Case nNoRef
        Case 1
'            If ((Trim$(Me.txtNombre1.Text) = "") And (Trim$(Me.txtAPaterno1.Text) = "") And (Trim$(Me.txtTel1.Text) = "")) Then
'                MsgBox "Faltan datos en la referencia 1.", vbExclamation, "KalaSystems"
'                Me.txtNombre1.SetFocus
'                Exit Function
'            End If
            
            If (Trim$(Me.txtNombre1.Text) <> "") Then
                If (Trim$(Me.txtAPaterno1.Text) = "") Then
                    MsgBox "Falta el apellido paterno de la referencia 1.", vbExclamation, "KalaSystems"
                    Me.txtAPaterno1.SetFocus
                    Exit Function
                End If
                
                If (Trim$(Me.txtTel1.Text) = "") Then
                    MsgBox "Falta el número telefónico en la referencia 1.", vbExclamation, "KalaSystems"
                    Me.txtTel1.SetFocus
                    Exit Function
                End If
                
                RefCorrectas = True
                Exit Function
            End If
            
            If (Trim$(Me.txtAPaterno1.Text) <> "") Then
                If (Trim$(Me.txtTel1.Text) = "") Then
                    MsgBox "Falta el número telefónico en la referencia 1.", vbExclamation, "KalaSystems"
                    Me.txtTel1.SetFocus
                    Exit Function
                End If
            
                If (Trim$(Me.txtNombre1.Text) = "") Then
                    MsgBox "Falta el nombre de la referencia 1.", vbExclamation, "KalaSystems"
                    Me.txtNombre1.SetFocus
                    Exit Function
                End If
                
                RefCorrectas = True
                Exit Function
            End If
            
            If (Trim$(Me.txtTel1.Text) <> "") Then
                If (Trim$(Me.txtNombre1.Text) = "") Then
                    MsgBox "Falta el nombre de la referencia 1.", vbExclamation, "KalaSystems"
                    Me.txtNombre1.SetFocus
                    Exit Function
                End If
                
                If (Trim$(Me.txtAPaterno1.Text) = "") Then
                    MsgBox "Falta el apellido paterno de la referencia 1.", vbExclamation, "KalaSystems"
                    Me.txtAPaterno1.SetFocus
                    Exit Function
                End If
                
                RefCorrectas = True
                Exit Function
            End If
    
        Case 2
'            If (Not (Trim$(Me.txtNombre2.Text) = Trim$(Me.txtAPaterno2.Text) = Trim$(Me.txtTel2.Text) = "")) Then
'                MsgBox "Faltan datos en la referencia 2.", vbExclamation, "KalaSystems"
'                Me.txtNombre1.SetFocus
'                Exit Function
'            End If
            
            If (Trim$(Me.txtNombre2.Text) <> "") Then
                If (Trim$(Me.txtAPaterno2.Text) = "") Then
                    MsgBox "Falta el apellido paterno de la referencia 2.", vbExclamation, "KalaSystems"
                    Me.txtAPaterno2.SetFocus
                    Exit Function
                End If
                
                If (Trim$(Me.txtTel2.Text) = "") Then
                    MsgBox "Falta el número telefónico en la referencia 2.", vbExclamation, "KalaSystems"
                    Me.txtTel2.SetFocus
                    Exit Function
                End If
                
                RefCorrectas = True
                Exit Function
            End If
            
            If (Trim$(Me.txtAPaterno2.Text) <> "") Then
                If (Trim$(Me.txtTel2.Text) = "") Then
                    MsgBox "Falta el número telefónico en la referencia 2.", vbExclamation, "KalaSystems"
                    Me.txtTel2.SetFocus
                    Exit Function
                End If
            
                If (Trim$(Me.txtNombre2.Text) = "") Then
                    MsgBox "Falta el nombre de la referencia 2.", vbExclamation, "KalaSystems"
                    Me.txtNombre2.SetFocus
                    Exit Function
                End If
                
                RefCorrectas = True
                Exit Function
            End If
            
            If (Trim$(Me.txtTel2.Text) <> "") Then
                If (Trim$(Me.txtNombre2.Text) = "") Then
                    MsgBox "Falta el nombre de la referencia 2.", vbExclamation, "KalaSystems"
                    Me.txtNombre2.SetFocus
                    Exit Function
                End If
                
                If (Trim$(Me.txtAPaterno2.Text) = "") Then
                    MsgBox "Falta el apellido paterno de la referencia 2.", vbExclamation, "KalaSystems"
                    Me.txtAPaterno2.SetFocus
                    Exit Function
                End If
                
                RefCorrectas = True
                Exit Function
            End If
    
    
        Case 3
'            If (Not (Trim$(Me.txtNombre3.Text) = Trim$(Me.txtAPaterno3.Text) = Trim$(Me.txtTel3.Text) = "")) Then
'                MsgBox "Faltan datos en la referencia 3.", vbExclamation, "KalaSystems"
'                Me.txtNombre1.SetFocus
'                Exit Function
'            End If
            
            If (Trim$(Me.txtNombre3.Text) <> "") Then
                If (Trim$(Me.txtAPaterno3.Text) = "") Then
                    MsgBox "Falta el apellido paterno de la referencia 3.", vbExclamation, "KalaSystems"
                    Me.txtAPaterno3.SetFocus
                    Exit Function
                End If
                
                If (Trim$(Me.txtTel3.Text) = "") Then
                    MsgBox "Falta el número telefónico en la referencia 3.", vbExclamation, "KalaSystems"
                    Me.txtTel3.SetFocus
                    Exit Function
                End If
                
                RefCorrectas = True
                Exit Function
            End If
            
            If (Trim$(Me.txtAPaterno3.Text) <> "") Then
                If (Trim$(Me.txtTel3.Text) = "") Then
                    MsgBox "Falta el número telefónico en la referencia 3.", vbExclamation, "KalaSystems"
                    Me.txtTel3.SetFocus
                    Exit Function
                End If
            
                If (Trim$(Me.txtNombre3.Text) = "") Then
                    MsgBox "Falta el nombre de la referencia 3.", vbExclamation, "KalaSystems"
                    Me.txtNombre3.SetFocus
                    Exit Function
                End If
                
                RefCorrectas = True
                Exit Function
            End If
            
            If (Trim$(Me.txtTel3.Text) <> "") Then
                If (Trim$(Me.txtNombre3.Text) = "") Then
                    MsgBox "Falta el nombre de la referencia 3.", vbExclamation, "KalaSystems"
                    Me.txtNombre3.SetFocus
                    Exit Function
                End If
                
                If (Trim$(Me.txtAPaterno3.Text) = "") Then
                    MsgBox "Falta el apellido paterno de la referencia 3.", vbExclamation, "KalaSystems"
                    Me.txtAPaterno3.SetFocus
                    Exit Function
                End If
                
                RefCorrectas = True
                Exit Function
            End If
    End Select
    
    RefCorrectas = True
End Function


Private Sub AMayusculas()
    With Me
        .txtNombre1.Text = UCase$(Trim$(Me.txtNombre1.Text))
        .txtAPaterno1.Text = UCase$(Trim$(Me.txtAPaterno1.Text))
        .txtAMaterno1.Text = UCase$(Trim$(Me.txtAMaterno1))
        .txtTel1.Text = UCase$(Trim$(Me.txtTel1.Text))
        
        .txtNombre2.Text = UCase$(Trim$(Me.txtNombre2.Text))
        .txtAPaterno2.Text = UCase$(Trim$(Me.txtAPaterno2.Text))
        .txtAMaterno2.Text = UCase$(Trim$(Me.txtAMaterno2))
        .txtTel2.Text = UCase$(Trim$(Me.txtTel2.Text))
        
        .txtNombre3.Text = UCase$(Trim$(Me.txtNombre3.Text))
        .txtAPaterno3.Text = UCase$(Trim$(Me.txtAPaterno3.Text))
        .txtAMaterno3.Text = UCase$(Trim$(Me.txtAMaterno3))
        .txtTel3.Text = UCase$(Trim$(Me.txtTel3.Text))
    End With
End Sub


Private Sub ClrTxtRefs()
    With Me
        .txtNombre1.Text = ""
        .txtAPaterno1.Text = ""
        .txtAMaterno1.Text = ""
        .txtTel1.Text = ""
        .txtNoReg1.Text = ""
        
        .txtNombre2.Text = ""
        .txtAPaterno2.Text = ""
        .txtAMaterno2.Text = ""
        .txtTel2.Text = ""
        .txtNoReg2.Text = ""
        
        .txtNombre3.Text = ""
        .txtAPaterno3.Text = ""
        .txtAMaterno3.Text = ""
        .txtTel3.Text = ""
        .txtNoReg3.Text = ""
    End With
End Sub


Private Sub cmdModAuto_Click()
    If (Me.ssdbAutos.Rows > 0) Then
        frmAutos.bNvaCal = False
        frmAutos.Show (1)
    End If
    
    frmOtrosDatos.ssdbAutos.SetFocus
End Sub


Private Sub cmdModExamen_Click()
    If (Me.ssdbExamMed.Rows > 0) Then
        frmCertificados.bNvoCertif = False
        frmCertificados.Show (1)
    End If
    
    frmOtrosDatos.ssdbExamMed.SetFocus
End Sub


Private Sub cmdModPase_Click()
    If (Me.ssdbPases.Rows > 0) Then
        frmPases.bNvoPase = False
        frmPases.Show (1)
    End If
    
    frmOtrosDatos.ssdbPases.SetFocus
End Sub


Private Sub cmdModTZone_Click()
    If (Me.ssdbTZone.Rows > 0) Then
        frmTZoneUsers.bNvaZona = False
        frmTZoneUsers.Show (1)
    End If
    
    frmOtrosDatos.ssdbTZone.SetFocus
End Sub


Private Sub cmdQuita1_Click()
    If (Not EliminaReg("Referencias", "idRef=" & Val(Me.txtNoReg1.Text), "", Conn)) Then
        MsgBox "No se borró la referencia seleccionada.", vbExclamation, "KalaSystems"
    End If
    
    LeeRefs
    
    InitVarRefs 0
End Sub


Private Sub cmdQuita2_Click()
    If (Not EliminaReg("Referencias", "idRef=" & Val(Me.txtNoReg2.Text), "", Conn)) Then
        MsgBox "No se borró la referencia seleccionada.", vbExclamation, "KalaSystems"
    End If
    
    LeeRefs
    
    InitVarRefs 0
End Sub


Private Sub cmdQuita3_Click()
    If (Not EliminaReg("Referencias", "idRef=" & Val(Me.txtNoReg3.Text), "", Conn)) Then
        MsgBox "No se borró la referencia seleccionada.", vbExclamation, "KalaSystems"
    End If
    
    LeeRefs
    
    InitVarRefs 0
End Sub


Private Sub cmdRefresca_Click()
    frmTZoneUsers.LlenaTZoneUsers
End Sub


Private Sub Form_Load()
    sTextToolBar = Trim(MDIPrincipal.StatusBar1.Panels.Item(1).Text)
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Pases, autos, ausencias, certificados y horarios"

    Me.txtTitCve.Text = frmAltaSocios.txtTitCve.Text
    Me.txtFamilia.Text = frmAltaSocios.txtFamilia.Text
    Me.txtTitular.Text = Trim$(frmAltaSocios.txtTitPaterno.Text) & " " & Trim$(frmAltaSocios.txtTitMaterno.Text) & " " & Trim$(frmAltaSocios.txtTitNombre.Text)
    
    LeeRefs
    
    InitVarRefs 0
    
    frmPases.LlenaPases
    frmAutos.LlenaAutos
    frmCertificados.LlenaCertificados
    frmTZoneUsers.LlenaTZoneUsers
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub


Private Sub LeeRefs()
Dim rsRefs As ADODB.Recordset
Dim i As Byte

    ClrTxtRefs

    InitRecordSet rsRefs, "idRef, Nombre, A_Paterno, A_Materno, Telefono", "Referencias", "idMember=" & Val(Me.txtTitCve.Text), "idRef", Conn
    With rsRefs
        If (.RecordCount > 0) Then
            i = 1
            
            .MoveFirst
            Do While (Not .EOF)
                Select Case i
                    Case 1
                        Me.txtNoReg1.Text = .Fields("idRef")
                        Me.txtNombre1.Text = .Fields("Nombre")
                        Me.txtAPaterno1.Text = .Fields("A_Paterno")
                        Me.txtAMaterno1.Text = .Fields("A_Materno")
                        Me.txtTel1.Text = .Fields("Telefono")
                    
                    Case 2
                        Me.txtNoReg2.Text = .Fields("idRef")
                        Me.txtNombre2.Text = .Fields("Nombre")
                        Me.txtAPaterno2.Text = .Fields("A_Paterno")
                        Me.txtAMaterno2.Text = .Fields("A_Materno")
                        Me.txtTel2.Text = .Fields("Telefono")
                    
                    Case 3
                        Me.txtNoReg3.Text = .Fields("idRef")
                        Me.txtNombre3.Text = .Fields("Nombre")
                        Me.txtAPaterno3.Text = .Fields("A_Paterno")
                        Me.txtAMaterno3.Text = .Fields("A_Materno")
                        Me.txtTel3.Text = .Fields("Telefono")
                End Select
            
                i = i + 1
                
                .MoveNext
            Loop
        End If
        
        .Close
    End With
    Set rsRefs = Nothing
End Sub


Private Sub InitVarRefs(nOpcion As Byte)
    Select Case nOpcion
        Case 0
            sNombre1 = Trim$(Me.txtNombre1.Text)
            sAPaterno1 = Trim$(Me.txtAPaterno1.Text)
            sAMaterno1 = Trim$(Me.txtAMaterno1.Text)
            sTel1 = Trim$(Me.txtTel1.Text)
            
            sNombre2 = Trim$(Me.txtNombre2.Text)
            sAPaterno2 = Trim$(Me.txtAPaterno2.Text)
            sAMaterno2 = Trim$(Me.txtAMaterno2.Text)
            sTel2 = Trim$(Me.txtTel2.Text)
            
            sNombre3 = Trim$(Me.txtNombre3.Text)
            sAPaterno3 = Trim$(Me.txtAPaterno3.Text)
            sAMaterno3 = Trim$(Me.txtAMaterno3.Text)
            sTel3 = Trim$(Me.txtTel3.Text)
    
        Case 1
            sNombre1 = Trim$(Me.txtNombre1.Text)
            sAPaterno1 = Trim$(Me.txtAPaterno1.Text)
            sAMaterno1 = Trim$(Me.txtAMaterno1.Text)
            sTel1 = Trim$(Me.txtTel1.Text)
            
        Case 2
            sNombre2 = Trim$(Me.txtNombre2.Text)
            sAPaterno2 = Trim$(Me.txtAPaterno2.Text)
            sAMaterno2 = Trim$(Me.txtAMaterno2.Text)
            sTel2 = Trim$(Me.txtTel2.Text)
    
        Case 3
            sNombre3 = Trim$(Me.txtNombre3.Text)
            sAPaterno3 = Trim$(Me.txtAPaterno3.Text)
            sAMaterno3 = Trim$(Me.txtAMaterno3.Text)
            sTel3 = Trim$(Me.txtTel3.Text)
    End Select
End Sub


Private Sub ssdbExamMed_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    If (Me.ssdbExamMed.Rows) Then
        Me.txtObsCert.Text = Me.ssdbExamMed.Columns(12).Text
    End If
End Sub
