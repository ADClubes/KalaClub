VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAltas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Altas"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8880
   Icon            =   "frmAltas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAltas.frx":030A
   ScaleHeight     =   6120
   ScaleWidth      =   8880
   Begin VB.TextBox txtCurp 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   5760
      TabIndex        =   37
      Top             =   1530
      Width           =   2385
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2955
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAltas.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAltas.frx":0934
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAltas.frx":0D88
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtaltas 
      Height          =   315
      Index           =   2
      Left            =   6795
      TabIndex        =   7
      Top             =   2220
      Width           =   1900
   End
   Begin VB.TextBox txtaltas 
      Height          =   315
      Index           =   1
      Left            =   4845
      TabIndex        =   5
      Top             =   2220
      Width           =   1900
   End
   Begin VB.TextBox txtaltas 
      Height          =   315
      Index           =   0
      Left            =   2910
      TabIndex        =   3
      Top             =   2220
      Width           =   1900
   End
   Begin VB.Frame frmLaborales 
      Caption         =   "Datos Laborales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   1515
      TabIndex        =   32
      Top             =   4845
      Width           =   7260
      Begin VB.ComboBox cboAltas 
         Height          =   315
         Index           =   4
         Left            =   5160
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   645
         Width           =   1935
      End
      Begin VB.ComboBox cboAltas 
         Height          =   315
         Index           =   3
         Left            =   900
         Sorted          =   -1  'True
         TabIndex        =   27
         Top             =   645
         Width           =   3495
      End
      Begin VB.ComboBox cboAltas 
         Height          =   315
         Index           =   2
         Left            =   1425
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   255
         Width           =   5655
      End
      Begin VB.Label lblLaboral 
         Alignment       =   1  'Right Justify
         Caption         =   "T&urno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   4500
         TabIndex        =   28
         Top             =   705
         Width           =   615
      End
      Begin VB.Label lblLaboral 
         Alignment       =   1  'Right Justify
         Caption         =   "&Puesto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   26
         Top             =   705
         Width           =   690
      End
      Begin VB.Label lblLaboral 
         Alignment       =   1  'Right Justify
         Caption         =   "Depa&rtamento:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   165
         TabIndex        =   24
         Top             =   315
         Width           =   1260
      End
   End
   Begin VB.TextBox txtRFC 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   3105
      TabIndex        =   1
      Top             =   1560
      Width           =   1725
   End
   Begin VB.Frame frmLogo 
      Height          =   5220
      Left            =   120
      TabIndex        =   31
      Top             =   825
      Width           =   1275
      Begin VB.Image imgKala 
         Height          =   4680
         Left            =   135
         Picture         =   "frmAltas.frx":10A8
         Stretch         =   -1  'True
         Top             =   270
         Width           =   975
      End
   End
   Begin VB.Frame frmGenerales 
      Caption         =   "Domicilio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   1485
      TabIndex        =   30
      Top             =   2565
      Width           =   7245
      Begin MSComCtl2.DTPicker DTPickerAltas 
         Height          =   360
         Left            =   5685
         TabIndex        =   23
         Top             =   1695
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   635
         _Version        =   393216
         Format          =   24444929
         CurrentDate     =   37781
      End
      Begin VB.TextBox txtaltas 
         Height          =   315
         Index           =   3
         Left            =   165
         TabIndex        =   9
         Top             =   465
         Width           =   3195
      End
      Begin VB.TextBox txtaltas 
         Height          =   315
         Index           =   5
         Left            =   4395
         TabIndex        =   13
         Top             =   465
         Width           =   2700
      End
      Begin VB.ComboBox cboAltas 
         Height          =   315
         Index           =   0
         Left            =   165
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   2265
      End
      Begin VB.ComboBox cboAltas 
         Height          =   315
         Index           =   1
         Left            =   2775
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   1080
         Width           =   3090
      End
      Begin VB.TextBox txtaltas 
         Height          =   315
         Index           =   7
         Left            =   165
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   1695
         Width           =   5175
      End
      Begin VB.TextBox txtaltas 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   3517
         TabIndex        =   11
         Top             =   465
         Width           =   720
      End
      Begin VB.TextBox txtaltas 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   6
         Left            =   6105
         TabIndex        =   19
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lblGral 
         Caption         =   "In&greso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   5700
         TabIndex        =   22
         Top             =   1470
         Width           =   1245
      End
      Begin VB.Label lblGral 
         Caption         =   "Númer&o:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   3532
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblGral 
         Caption         =   "&Teléfono(s):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   180
         TabIndex        =   20
         Top             =   1470
         Width           =   1245
      End
      Begin VB.Label lblGral 
         Caption         =   "&C. P.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   6120
         TabIndex        =   18
         Top             =   855
         Width           =   585
      End
      Begin VB.Label lblGral 
         Caption         =   "&Entidad Federativa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   14
         Top             =   855
         Width           =   1770
      End
      Begin VB.Label lblGral 
         Caption         =   "&Delegación o Mununicipio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2790
         TabIndex        =   16
         Top             =   855
         Width           =   2625
      End
      Begin VB.Label lblGral 
         Caption         =   "Colon&ia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4410
         TabIndex        =   12
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lblGral 
         Caption         =   "Ca&lle:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   240
         Width           =   510
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   690
      Left            =   570
      TabIndex        =   34
      Top             =   15
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1217
      BandCount       =   1
      _CBWidth        =   1860
      _CBHeight       =   690
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   630
      Width1          =   1800
      NewRow1         =   0   'False
      BandStyle1      =   1
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   630
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1111
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "NUEVO"
               Object.ToolTipText     =   "Nuevo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "GUARDAR"
               Object.ToolTipText     =   "Guardar"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SALIR"
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label lblGral 
      Alignment       =   1  'Right Justify
      Caption         =   "C.&U.R.P.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   5685
      TabIndex        =   36
      Top             =   1320
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   1530
      Left            =   1530
      Top             =   975
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Foto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1980
      TabIndex        =   33
      Top             =   1635
      Width           =   420
   End
   Begin VB.Label lblGral 
      Alignment       =   1  'Right Justify
      Caption         =   "R. &F. C.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   3090
      TabIndex        =   0
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label lblNombre 
      Caption         =   "&Nombre(s):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6810
      TabIndex        =   6
      Top             =   2010
      Width           =   1200
   End
   Begin VB.Label ApMaterno 
      Caption         =   "Ap. &Materno:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4860
      TabIndex        =   4
      Top             =   2010
      Width           =   1200
   End
   Begin VB.Label lblApPaterno 
      Caption         =   "&Ap. Paterno:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2925
      TabIndex        =   2
      Top             =   1995
      Width           =   1200
   End
   Begin VB.Image imgFoto 
      Height          =   1590
      Left            =   1530
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1320
   End
End
Attribute VB_Name = "frmAltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AdoRcsAltas As ADODB.Recordset, AdoComAltas As ADODB.Command
Dim blnGuardado As Boolean

Private Sub cmdCancelar_Click()
   Unload Me
End Sub

Private Sub cboAltas_Click(Index As Integer)
   Dim lngCveDelomuni As Long
   Dim Controles As TextBox, strCampo1 As String, strCampo2 As String
   
   Select Case Index
      Case 0
         'Llena combo Delegación o Municipio
         lngCveDelomuni = Me.cboAltas(0).ItemData(Me.cboAltas(0).ListIndex)
         strSQL = "SELECT cvedelomuni, nomdelomuni FROM delomuni WHERE entidadfed = " & lngCveDelomuni
         strCampo1 = "nomdelomuni"
         strCampo2 = "cvedelomuni"
         Call LlenaCombos(cboAltas(1), strSQL, strCampo1, strCampo2)
      End Select
End Sub

Private Sub Form_Activate()
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "ALTAS DE EMPLEADOS"
    With frmAltas
        .Height = 6600
        .Left = 0
        .Top = 0
        .Width = 9000
    End With
End Sub

Private Sub Form_Load()
   Dim Controles As TextBox, strCampo1 As String, strCampo2 As String
   blnGuardado = False
   DTPickerAltas.Value = Now()
   
   'Llena combo Entidad Federativa
   strSQL = "SELECT cveentfederativa, nomentfederativa FROM entfederativa"
   strCampo1 = "nomentfederativa"
   strCampo2 = "cveentfederativa"
   Call LlenaCombos(cboAltas(0), strSQL, strCampo1, strCampo2)
   cboAltas(0).ListIndex = 8
   
   'Llena combo Departamento
   strSQL = "SELECT cvedepartamento, nomdepartamento FROM departamentos"
   strCampo1 = "nomdepartamento"
   strCampo2 = "cvedepartamento"
   Call LlenaCombos(cboAltas(2), strSQL, strCampo1, strCampo2)

   'Llena combo Puesto
   strSQL = "SELECT cvepuesto, nompuesto FROM puestos"
   strCampo1 = "nompuesto"
   strCampo2 = "cvepuesto"
   Call LlenaCombos(cboAltas(3), strSQL, strCampo1, strCampo2)

   'Llena combo Turno
   strSQL = "SELECT cveturno, nomturno FROM turnos"
   strCampo1 = "nomturno"
   strCampo2 = "cveturno"
   Call LlenaCombos(cboAltas(4), strSQL, strCampo1, strCampo2)
   Call DesactivaControles
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim Respuesta As Integer
   On Error GoTo Barra_Herramientas_err
   Select Case UCase(Button.Key)
      Case "NUEVO"
         If (blnGuardado = True) Or (txtRFC = "") Then
            Call LimpiaControles
         Else
            Beep
            Respuesta = MsgBox("¿ Desea Gravar el presente registro antes ?", vbYesNo + vbQuestion, "Altas")
            If Respuesta = vbNo Then
               Call LimpiaControles
            Else
               Call GuardaAlta
            End If
         End If
      Case "GUARDAR"
         Call GuardaAlta
      Case "SALIR"
         If txtRFC.Text = "" Then
            Unload Me
            Exit Sub
         End If
         If blnGuardado = False Then
            Beep
            Respuesta = MsgBox("¿ Desea Salir ?", vbYesNo + vbQuestion, "Altas")
            If Respuesta = vbYes Then
               Unload Me
            Else
               Exit Sub
            End If
         Else
            Unload Me
         End If
   End Select
   Exit Sub
Barra_Herramientas_err:
End Sub

Private Sub txtaltas_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCurp_KeyPress(KeyAscii As Integer)
 Dim strNombre As String, strPaterno As String, strMaterno As String, strBusca As String
      If KeyAscii = 13 Then
         If Len(txtCurp) < 18 Then
            MsgBox "¡ El RFC es incorrecto o está incompleto !", vbCritical + vbOKOnly, "Altas"
            txtCurp.SetFocus
            Exit Sub
         End If
         'strBusca = Left(txtRFC, 10)
         strBusca = txtCurp.Text
         strSQL = "SELECT apellidopat, apellidomat, nombre FROM empleados WHERE curp = '"
         strSQL = strSQL & Trim(strBusca) & "'"
         
         Set AdoRcsAltas = New ADODB.Recordset
         AdoRcsAltas.ActiveConnection = Conn
         AdoRcsAltas.CursorLocation = adUseClient
         AdoRcsAltas.CursorType = adOpenDynamic
         AdoRcsAltas.LockType = adLockReadOnly
         AdoRcsAltas.Open strSQL
         
         If Not AdoRcsAltas.EOF Then
            strNombre = AdoRcsAltas!Nombre
            strPaterno = AdoRcsAltas!apellidopat
            strMaterno = AdoRcsAltas!apellidomat
            MsgBox "¡ No se puede registrar a " & strNombre & " " & strPaterno & " " & strMaterno & _
            ", pues YA se encuentra registrado(a) en la Base de Datos !", vbInformation + vbOKOnly, "Altas"
            txtCurp.Text = ""
            txtCurp.SetFocus
         Else
            txtCurp.Enabled = False
            txtaltas(0).Enabled = True
            txtaltas(1).Enabled = True
            txtaltas(2).Enabled = True
            frmGenerales.Enabled = True
            frmLaborales.Enabled = True
            txtaltas(0).SetFocus
         End If
         Exit Sub
      Else
           KeyAscii = Asc(UCase(Chr(KeyAscii)))
      End If
End Sub

Private Sub txtrfc_KeyPress(KeyAscii As Integer)
  Dim strNombre As String, strPaterno As String, strMaterno As String, strBusca As String
      If KeyAscii = 13 Then
         If Len(txtRFC) < 10 Then
            MsgBox "¡ El RFC es incorrecto o está incompleto !", vbCritical + vbOKOnly, "Altas"
            txtRFC.SetFocus
            Exit Sub
         End If
         'strBusca = Left(txtRFC, 10)
         strBusca = txtRFC.Text
         strSQL = "SELECT apellidopat, apellidomat, nombre FROM empleados WHERE rfc = '"
         strSQL = strSQL & Trim(strBusca) & "'"
         
         Set AdoRcsAltas = New ADODB.Recordset
         AdoRcsAltas.ActiveConnection = Conn
         AdoRcsAltas.CursorLocation = adUseClient
         AdoRcsAltas.CursorType = adOpenDynamic
         AdoRcsAltas.LockType = adLockReadOnly
         AdoRcsAltas.Open strSQL
         
         If Not AdoRcsAltas.EOF Then
            strNombre = AdoRcsAltas!Nombre
            strPaterno = AdoRcsAltas!apellidopat
            strMaterno = AdoRcsAltas!apellidomat
            MsgBox "¡ El Empleado " & strNombre & " " & strPaterno & " " & strMaterno & _
            " Ya Se Encuentra Registrado en la Base de Datos !", vbInformation + vbOKOnly, "Altas"
            txtRFC.Text = ""
            txtRFC.SetFocus
         Else
            txtRFC.Enabled = False
            txtaltas(0).Enabled = True
            txtaltas(1).Enabled = True
            txtaltas(2).Enabled = True
            frmGenerales.Enabled = True
            frmLaborales.Enabled = True
            txtaltas(0).SetFocus
            Toolbar1.Buttons(2).Enabled = True
         End If
         Exit Sub
      Else
           KeyAscii = Asc(UCase(Chr(KeyAscii)))
      End If
End Sub

Private Sub txtRFC_LostFocus()
    Dim strNombre As String, strPaterno As String, strMaterno As String, strBusca As String
    If Len(txtRFC) < 10 Then
       MsgBox "¡ El RFC es incorrecto o está incompleto !", vbCritical + vbOKOnly, "Altas"
       txtRFC.SetFocus
       Exit Sub
    End If
    'strBusca = Left(txtRFC, 10)
    strBusca = txtRFC.Text
    strSQL = "SELECT apellidopat, apellidomat, nombre FROM empleados WHERE rfc = '"
    strSQL = strSQL & Trim(strBusca) & "'"
    
    Set AdoRcsAltas = New ADODB.Recordset
    AdoRcsAltas.ActiveConnection = Conn
    AdoRcsAltas.CursorLocation = adUseClient
    AdoRcsAltas.CursorType = adOpenDynamic
    AdoRcsAltas.LockType = adLockReadOnly
    AdoRcsAltas.Open strSQL
    
    If Not AdoRcsAltas.EOF Then
       strNombre = AdoRcsAltas!Nombre
       strPaterno = AdoRcsAltas!apellidopat
       strMaterno = AdoRcsAltas!apellidomat
       MsgBox "¡ La Matrícula ( " & txtRFC & " ) Ya Está Registrado Con El Nombre De: ( " & _
                    strNombre & " " & strPaterno & " " & strMaterno & _
                    " )", vbInformation + vbOKOnly, "Altas"
       txtRFC.Text = ""
       txtRFC.SetFocus
    Else
       txtRFC.Enabled = False
       txtaltas(0).Enabled = True
       txtaltas(1).Enabled = True
       txtaltas(2).Enabled = True
       frmGenerales.Enabled = True
       frmLaborales.Enabled = True
       txtaltas(0).SetFocus
       Toolbar1.Buttons(2).Enabled = True
    End If
End Sub

Sub GuardaAlta()
    Dim lgFolio, IniTrans As Long
    On Error GoTo err_Altas
    Screen.MousePointer = vbHourglass
    
    If Not VerificaDatos Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ' Obtenemos el Número de Folio para el asignarlo al Número de Empleado
    strSQL = "SELECT folioempleados FROM Folios"
    AdoRcsAltas.ActiveConnection = Conn
    AdoRcsAltas.LockType = adLockOptimistic
    AdoRcsAltas.CursorType = adOpenKeyset
    AdoRcsAltas.CursorLocation = adUseServer
    AdoRcsAltas.Open strSQL
    If Not AdoRcsAltas.EOF Then
        AdoRcsAltas!folioempleados = AdoRcsAltas!folioempleados + 1
        lgFolio = AdoRcsAltas!folioempleados
    Else
        AdoRcsAltas.AddNew
        AdoRcsAltas!folioempleados = 1
        lgFolio = 1
    End If
    AdoRcsAltas.Update
    AdoRcsAltas.Close
    Set AdoRcsAltas = Nothing
    
    strSQL = "INSERT INTO empleados (numepleado, nombre, apellidopat, apellidomat, " & _
                "calle, numero, colonia, entfederativa, delomuni, rfc, curp, departamento, " & _
                "puesto, turno, ingreso) VALUES(" & lgFolio & ", '" & Trim(txtaltas(2).Text) & _
                "', '" & Trim(txtaltas(0).Text) & "', '" & Trim(txtaltas(1).Text) & "', '" & _
                Trim(txtaltas(3).Text) & "', '" & Trim(txtaltas(4).Text) & "', '" & _
                Trim(txtaltas(5).Text) & "', '" & Trim(cboAltas(0).Text) & "', '" & _
                Trim(cboAltas(1).Text) & "', '" & Trim(txtRFC.Text)
    Conn.CommitTrans                          'Termina transacción
    Screen.MousePointer = vbDefault
    MsgBox "Todo bien "
    Exit Sub
    
err_Altas:
    Screen.MousePointer = Default
    If IniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Sub LimpiaControles()
   Dim I As Integer
   txtRFC.Text = ""
   For I = 0 To 7
      txtaltas(I).Text = ""
   Next I
   For I = 0 To 4
      cboAltas(I).Text = ""
   Next I
   Call DesactivaControles
   txtRFC.SetFocus
End Sub

Sub DesactivaControles()
   'Desactiva todos los controles
   txtaltas(0).Enabled = False
   txtaltas(1).Enabled = False
   txtaltas(2).Enabled = False
   frmGenerales.Enabled = False
   frmLaborales.Enabled = False
   txtRFC.Enabled = True
   Toolbar1.Buttons(2).Enabled = False
End Sub

Private Function VerificaDatos() As Boolean
   Dim I As Integer, strDato As String
      strDato = ""
      For I = 0 To 6
         If txtaltas(I).Text = "" Then
            Select Case I
               Case 0
                  strDato = "Apellido Paterno"
               Case 1
                  strDato = "Apellido Materno"
               Case 2
                  strDato = "Nombre"
               Case 3
                  strDato = "Calle del Domicilio"
               Case 4
                  strDato = "Número del Domicilio"
               Case 5
                  strDato = "Colonia del Domicilio"
               Case 6
                  strDato = "Código Postal del Domicilio"
               End Select
            MsgBox "¡ Favor De Llenar La Información Para La Casilla De " & strDato & " !", vbOKOnly, "Altas"
            VerificaDatos = False
            txtaltas(I).SetFocus
            Exit Function
         End If
      Next I
      For I = 0 To 4
         If cboAltas(I).Text = "" Then
            Select Case I
               Case 0
                  strDato = "Entidad Federativa Del Domicilio"
               Case 1
                  strDato = "Delegación o Municipio Del Domicilio"
               Case 2
                  strDato = "Departamento Al Que Ingresa"
               Case 3
                  strDato = "Puesto Al Que Ingresa"
               Case 4
                  strDato = "Turno Al Que Ingresa"
               End Select
            MsgBox "¡ Favor De Llenar La Información Para La Casilla De " & strDato & " !", vbOKOnly, "Altas"
            VerificaDatos = False
            cboAltas(I).SetFocus
            Exit Function
         End If
         VerificaDatos = True
      Next I
End Function
