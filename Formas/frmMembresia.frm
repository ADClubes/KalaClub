VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmMembresia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de inscripciones"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8415
   Icon            =   "frmMembresia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMantIni 
      Caption         =   "Mantenimiento"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   48
      Top             =   5160
      Width           =   8055
      Begin VB.ComboBox cmbMantIni 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   240
         Width           =   3015
      End
      Begin VB.OptionButton optDirec 
         Caption         =   "Direccionado"
         Height          =   255
         Left            =   6480
         TabIndex        =   50
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optConv 
         Caption         =   "Convencional"
         Height          =   255
         Left            =   4080
         TabIndex        =   49
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame frmVendedor 
      Caption         =   "  Vendida por  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   8055
      Begin VB.TextBox txtVendedor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   480
         Width           =   6375
      End
      Begin VB.CommandButton cmdHVendedor 
         Height          =   255
         Left            =   960
         Picture         =   "frmMembresia.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtCveVendedor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblVendedor 
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   1440
         TabIndex        =   47
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblCveVendedor 
         Caption         =   "Clave"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdGeneraRen 
      Caption         =   "Genera Pagos"
      Height          =   735
      Left            =   4200
      TabIndex        =   19
      ToolTipText     =   "  Generar pagos  "
      Top             =   6120
      Width           =   735
   End
   Begin VB.CheckBox chkPrimerDia 
      Caption         =   "Pagos a partir del día primero de cada mes"
      Height          =   615
      Left            =   1920
      TabIndex        =   18
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Frame frameNoMem 
      Caption         =   " # Inscripción "
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   6000
      Width           =   1455
      Begin VB.TextBox txtNoMem 
         Alignment       =   2  'Center
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
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame frameDetalle 
      Height          =   975
      Left            =   5280
      TabIndex        =   26
      Top             =   7080
      Width           =   2895
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpFechaVenc 
         Height          =   285
         Left            =   1440
         TabIndex        =   28
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   106299393
         CurrentDate     =   38531
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblFechaVenc 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   1440
         TabIndex        =   39
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frameGrales 
      Caption         =   "  Datos de la inscripción "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8055
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbOrigenVenta 
         Height          =   375
         Left            =   4200
         TabIndex        =   53
         Top             =   1800
         Width           =   3615
         DataFieldList   =   "Column 0"
         _Version        =   196616
         DataMode        =   2
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns.Count   =   4
         Columns(0).Width=   5212
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   1693
         Columns(1).Caption=   "CveOrigen"
         Columns(1).Name =   "CveOrigen"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         Columns(2).Width=   1270
         Columns(2).Caption=   "SinCosto"
         Columns(2).Name =   "SinCosto"
         Columns(2).DataField=   "Column 2"
         Columns(2).DataType=   8
         Columns(2).FieldLen=   256
         Columns(3).Width=   1508
         Columns(3).Caption=   "EsVenta"
         Columns(3).Name =   "EsVenta"
         Columns(3).DataField=   "Column 3"
         Columns(3).DataType=   8
         Columns(3).FieldLen=   256
         _ExtentX        =   6376
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   405
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3240
         Width           =   7695
      End
      Begin VB.TextBox txtNombreProp 
         Height          =   285
         Left            =   120
         MaxLength       =   80
         TabIndex        =   4
         Top             =   1200
         Width           =   7695
      End
      Begin VB.ComboBox cmbTipoMembresia 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1800
         Width           =   3855
      End
      Begin VB.CommandButton cmdHTitular 
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         Picture         =   "frmMembresia.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   6375
      End
      Begin VB.TextBox txtCveTit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         MaxLength       =   5
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtNoPagos 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtEnganche 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   8
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtMontoTotal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox txtDuracion 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   6
         Top             =   2640
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFechaAlta 
         Height          =   285
         Left            =   6480
         TabIndex        =   10
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   106299393
         CurrentDate     =   38531
      End
      Begin VB.Label Label6 
         Caption         =   "Origen del Alta"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4200
         TabIndex        =   52
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre del Propietario"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de la Inscripción"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   1440
         TabIndex        =   38
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblCveTit 
         Caption         =   "# Titular"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblFechaAlta 
         Caption         =   "Fecha de alta"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6480
         TabIndex        =   36
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblNoPagos 
         Alignment       =   2  'Center
         Caption         =   "No. de pagos (máximo 4)"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4800
         TabIndex        =   35
         Top             =   2235
         Width           =   1335
      End
      Begin VB.Label lblEnganche 
         Caption         =   "Enganche"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblMontoTotal 
         Caption         =   "Monto total"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblDuracion 
         Caption         =   "Duración de la inscripción"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   2235
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdDeshacer 
      Caption         =   "Elimina Ren."
      Enabled         =   0   'False
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "  Elimina renglón  "
      Top             =   8280
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Acepta datos"
      Enabled         =   0   'False
      Height          =   735
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "  Efectuar cambio  "
      Top             =   8280
      Width           =   735
   End
   Begin VB.CommandButton cmdAceptar 
      Height          =   735
      Left            =   5280
      Picture         =   "frmMembresia.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "  Guardar datos  "
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   735
      Left            =   7440
      Picture         =   "frmMembresia.frx":0B18
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "  Salir  "
      Top             =   6120
      Width           =   735
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Inserta Ren."
      Enabled         =   0   'False
      Height          =   735
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "  Insertar renglón  "
      Top             =   8280
      Width           =   735
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   735
      Left            =   6360
      Picture         =   "frmMembresia.frx":150E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "  Cancelar  "
      Top             =   6120
      Width           =   735
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbPagos 
      Height          =   2055
      Left            =   120
      TabIndex        =   25
      Top             =   7080
      Width           =   4935
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      AllowUpdate     =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      BackColorOdd    =   12640511
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   1402
      Columns(0).Caption=   "# Pago"
      Columns(0).Name =   "NoPago"
      Columns(0).CaptionAlignment=   1
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2170
      Columns(1).Caption=   "Vence"
      Columns(1).Name =   "Vence"
      Columns(1).CaptionAlignment=   1
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1931
      Columns(2).Caption=   "Monto"
      Columns(2).Name =   "Monto"
      Columns(2).CaptionAlignment=   1
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   6
      Columns(2).NumberFormat=   "CURRENCY"
      Columns(2).FieldLen=   256
      Columns(3).Width=   1931
      Columns(3).Caption=   "Fec.Pago"
      Columns(3).Name =   "FechaPago"
      Columns(3).CaptionAlignment=   1
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3519
      Columns(4).Caption=   "Observaciones"
      Columns(4).Name =   "Observaciones"
      Columns(4).CaptionAlignment=   1
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1931
      Columns(5).Caption=   "IdReg"
      Columns(5).Name =   "IdReg"
      Columns(5).CaptionAlignment=   1
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   8705
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Relación de pagos por realizar"
      Enabled         =   0   'False
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
   Begin VB.Label lblPorCubrir 
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label lblCubierto 
      Height          =   255
      Left            =   960
      TabIndex        =   23
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Cubierto:"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Por Cubrir:"
      Height          =   375
      Left            =   2280
      TabIndex        =   42
      Top             =   6840
      Width           =   855
   End
End
Attribute VB_Name = "frmMembresia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'*  Formulario para registro de membresias  *
'*  Daniel Hdez                             *
'*  28 / Junio / 2005                       *
'********************************************

'*  Ult actualización:  14 / Noviembre / 2005


Public sFormaAnterior As String
Public nAyuda As Byte

'Número de columnas mostradas
Const TOTCOLUMNS = 6

'Número máximo de pagos permitidos
Const MAXPAGOS = 4

'Encabezados para la consulta
Dim mTitPagos(TOTCOLUMNS) As String

'Ancho de las columnas mostradas
Dim mAncPagos(TOTCOLUMNS) As Integer

'Nombre de los campos en la consulta
Dim mCamPagos(TOTCOLUMNS) As String

'Matriz de los encabezados del formulario
Dim mTitForma(TOTCOLUMNS) As String

'Columna activa o seleccionada para establecer el orden
Dim nColActiva As Integer

'Posicion del cursor dentro del dbGrid
Dim nPos As Variant

Dim bSave As Boolean
Dim bNvaMem As Boolean

Dim nDuracion As Integer
Dim nMontoTotal As Double
Dim nEnganche As Double
Dim nPagos As Integer

Dim nMonto As Double
Dim dFechaVence As Date

'12/10/2005 gpo
Dim sNombrePropietario As String
Dim nTipoMem As Integer
Dim dCubierto As Double
Dim dPorCubrir As Double
Dim bYaPagada As Boolean

Private Sub cmdAceptar_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdAceptar.Name) Then
        MsgBox "No tiene acceso a esta opción", vbCritical, "Error"
        Exit Sub
    End If
    
    '20111219 UCM SE RESTRINGE DESDE BASE DE DATOS
'    If Not sDB_NivelUser = 6 Or Not sDB_NivelUser = 4 Then
'        MsgBox "Función disponible sólo para ejecutivos de ventas."
'        Exit Sub
'    End If
    
'    If Me.dtpFechaAlta.Value < Date Then
'        MsgBox "No es posible modificar los datos!", vbInformation, "Error"
'        Exit Sub
'    End If
    
    Screen.MousePointer = vbHourglass
    If (Cambios) Then
        If (Not bSave) Then
            If (Not GuardaDatos) Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
        
        If (sFormaAnterior = "frmConsMembers") Then
            bSave = False
            bNvaMem = True
            
            ClrCtrlsSocios
            CtrlsSocios True
            
            ClrCtrlsTxt
            ActivaCtrls True
        End If
        
        AsignaVars
        
        RefreshSSGrid
    End If
    
    If (Me.txtCveTit.Enabled) Then
        Me.txtCveTit.SetFocus
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub ActivaBotones(bValor As Boolean)
    With Me
        .ssdbPagos.Enabled = bValor
        .cmdOk.Enabled = bValor
        .cmdModificar.Enabled = bValor
        .cmdDeshacer.Enabled = bValor
    End With
End Sub


Private Sub AsignaVars()
    nDuracion = Val(Me.txtDuracion.Text)
    nMontoTotal = Val(Me.txtMontoTotal.Text)
    nEnganche = Val(Me.txtEnganche.Text)
    nPagos = Val(Me.txtNoPagos.Text)
    
    If Me.cmbTipoMembresia.ListIndex >= 0 Then
        nTipoMem = Me.cmbTipoMembresia.ItemData(Me.cmbTipoMembresia.ListIndex)
    Else
        nTipoMem = 0
    End If
    
End Sub


Private Sub cmdCancelar_Click()
    If ((Not bSave) And (Not bNvaMem)) Then
        ClrCtrlsTxt
    End If
End Sub


Private Sub cmdDeshacer_Click()

    Dim vBookMark As Variant
    Dim lRenAct As Long
    
    If Me.ssdbPagos.Rows = 0 Then
        MsgBox "No hay renglones para eliminar!", vbExclamation, "Membresias"
        Exit Sub
    End If
    
    If Me.ssdbPagos.Columns("Fec.Pago").Value <> "" Then
        MsgBox "NO se pueden eliminar los pagos ya efectuados!", vbExclamation, "Membresias"
        Exit Sub
    End If
    
    
    If Val(Me.ssdbPagos.Columns("# Pago").Value) = 0 And Me.ssdbPagos.Rows > 1 Then
        MsgBox "Se deben eliminar primero los pagos mayores a 0", vbExclamation, "Membresias"
        Exit Sub
    End If
    
        
    lRenAct = Me.ssdbPagos.AddItemRowIndex(Me.ssdbPagos.Bookmark)
    
    
    Select Case lRenAct
        'Si es el ultimo renglon
        Case Is = Me.ssdbPagos.Rows - 1
            vBookMark = Me.ssdbPagos.AddItemBookmark(Me.ssdbPagos.Rows - 2)
        Case Is < Me.ssdbPagos.Rows - 1
            vBookMark = Me.ssdbPagos.AddItemBookmark(lRenAct + 1)
        Case Else
            vBookMark = Me.ssdbPagos.Bookmark
    End Select
    
    
    Me.ssdbPagos.RemoveItem Me.ssdbPagos.AddItemRowIndex(Me.ssdbPagos.Bookmark)
    Me.ssdbPagos.Update
    
    If Me.ssdbPagos.Rows > 1 Then
        Me.txtNoPagos.Text = Me.ssdbPagos.Rows - 1
    Else
        Me.txtNoPagos.Text = 0
    End If
    
    Me.ssdbPagos.Bookmark = vBookMark
    
    CalculaTotalGrid True
        
    Me.lblCubierto = Format(dCubierto, "$#,##0.00")
    Me.lblPorCubrir = Format(dPorCubrir, "$#,##0.00")
    
    If Me.ssdbPagos.Rows = 0 Then
        Me.txtMontoTotal.Enabled = True
        Me.txtEnganche.Enabled = True
        Me.txtNoPagos.Enabled = True
        Me.dtpFechaAlta.Enabled = True
        Me.cmdGeneraRen.Enabled = True
        'gpo
        '31/08/09
        Me.sscmbOrigenVenta.Enabled = True
    End If
    
    
End Sub

Private Sub cmdGeneraRen_Click()
    
    If Me.sscmbOrigenVenta.Text = vbNullString Then
        MsgBox "¡Seleccionar el origen de la venta!", vbInformation, "Verifique"
        Me.sscmbOrigenVenta.SetFocus
    End If
    
    
    If Me.txtMontoTotal.Text = "" Then
        MsgBox "Indicar el monto total!", vbExclamation, "Membresias"
        Me.txtMontoTotal.SetFocus
        Exit Sub
    End If
    
    If Me.sscmbOrigenVenta.Columns("SinCosto").Value = "N" And CDbl(Me.txtMontoTotal.Text) <= 0 Then
        MsgBox "El monto total debe ser mayor que 0!", vbExclamation, "Membresias"
        Me.txtMontoTotal.SetFocus
        Exit Sub
    End If
    
    If Me.txtEnganche.Text = "" Then
        MsgBox "Indicar el enganche!", vbExclamation, "Membresias"
        Me.txtEnganche.SetFocus
        Exit Sub
    End If
    
    If CDbl(Me.txtEnganche.Text) < 0 Then
        MsgBox "El enganche debe ser mayor que 0!", vbExclamation, "Membresias"
        Me.txtEnganche.SetFocus
        Exit Sub
    End If
    
    If Me.sscmbOrigenVenta.Columns("SinCosto").Value = "N" And CDbl(Me.txtEnganche.Text) < 0 Then
        MsgBox "El enganche debe ser mayor que 0!", vbExclamation, "Membresias"
        Me.txtEnganche.SetFocus
        Exit Sub
    End If
    
    If Me.txtNoPagos.Text = "" Then
        MsgBox "Indicar el número de pagos!", vbExclamation, "Membresias"
        Me.txtNoPagos.SetFocus
        Exit Sub
    End If
    
    If CDbl(Me.txtEnganche.Text) > CDbl(Me.txtMontoTotal.Text) Then
        MsgBox "El enganche es mayor que el monto total", vbInformation, "Membresias"
        Exit Sub
    End If
    
    
    If CDbl(Me.txtEnganche.Text) < CDbl(Me.txtMontoTotal.Text) And CDbl(Me.txtNoPagos.Text) = 0 Then
        MsgBox "El enganche no cubre el monto total" & vbLf & "y no se indicaron pagos parciales", vbInformation, "Membresias"
        Exit Sub
    End If
    
    If CDbl(Me.txtEnganche.Text) = CDbl(Me.txtMontoTotal.Text) And CDbl(Me.txtNoPagos.Text) > 0 Then
        Me.txtNoPagos.Text = 0
    End If
    
    If Val(Me.txtNoPagos.Text) > MAXPAGOS Then
        MsgBox "El número de pagos no puede ser mayor de " & MAXPAGOS & "!", vbExclamation, "Membresias"
        Me.txtNoPagos.SetFocus
        Exit Sub
    End If
    
    GeneraRenglones
    
    Me.txtMontoTotal.Enabled = False
    Me.txtEnganche.Enabled = False
    Me.txtNoPagos.Enabled = False
    Me.dtpFechaAlta.Enabled = False
    'gpo
    '31/08/09
    Me.sscmbOrigenVenta.Enabled = False
    
    Me.txtMonto.Enabled = True
    Me.ssdbPagos.Enabled = True
    Me.cmdDeshacer.Enabled = True
    Me.cmdModificar.Enabled = True
    Me.cmdOk.Enabled = True
    
    Me.lblCubierto = Format(dCubierto, "$#,##0.00")
    Me.lblPorCubrir = Format(dPorCubrir, "$#,##0.00")
    
    If Me.ssdbPagos.Rows > 0 Then
        Me.ssdbPagos.Bookmark = Me.ssdbPagos.AddItemBookmark(0)
    End If
    
End Sub

Private Sub cmdModificar_Click()
    
    If Me.ssdbPagos.Rows = 0 Then
        MsgBox "Generar primero los pagos!", vbExclamation, "Membresias"
        Exit Sub
    End If
    
    If Me.ssdbPagos.Rows >= MAXPAGOS + 1 Then
        MsgBox "Ya no se pueden insertar más renglones", vbExclamation, "Membresias"
        Exit Sub
    End If
    If Me.txtMonto.Text = "" Then
        MsgBox "Indicar el monto!", vbExclamation, "Membresias"
        Exit Sub
    End If
    
    If CDbl(Me.txtMonto.Text) = 0 Then
        MsgBox "El monto debe ser mayor de 0!", vbExclamation, "Membresias"
        Exit Sub
    End If
    
    Me.ssdbPagos.MoveLast
    If CDate(Me.ssdbPagos.Columns("Vence").Value) > CDate(Me.dtpFechaVenc.Value) Then
        MsgBox "La fecha de vencimiento no puede ser menor que la fecha del pago anterior!", vbExclamation, "Membresias"
        Exit Sub
    End If
    
    Me.ssdbPagos.AddItem 0 & vbTab & Me.dtpFechaVenc.Value & vbTab & CDbl(Me.txtMonto.Text) & vbTab & "" & vbTab & "" & vbTab & 0
    
    Me.txtNoPagos.Text = Me.ssdbPagos.Rows - 1
    
    CalculaTotalGrid True
    
    Me.lblCubierto = Format(dCubierto, "$#,##0.00")
    Me.lblPorCubrir = Format(dPorCubrir, "$#,##0.00")
    
    Me.txtMonto.Text = ""
    Me.dtpFechaVenc.Value = Date
End Sub

Private Sub cmdOk_Click()
    If Me.ssdbPagos.Rows = 0 Then
        Exit Sub
    End If
    
    If Me.txtMonto.Text = "" Then
        Exit Sub
    End If
    
    If Val(Me.ssdbPagos.Columns("# Pago").Value) = 0 Then
        MsgBox "No se puede modificar el monto del enganche!", vbExclamation, "Membresias"
        Exit Sub
    End If
    
    If Me.ssdbPagos.Columns("Fec.Pago").Value <> "" Then
        MsgBox "NO se pueden eliminar los pagos ya efectuados!", vbExclamation, "Membresias"
        Exit Sub
    End If
    
    If Me.ssdbPagos.AddItemRowIndex(Me.ssdbPagos.Bookmark) > 0 Then
        Me.ssdbPagos.MovePrevious
        If CDate(Me.ssdbPagos.Columns("Vence").Value) > CDate(Me.dtpFechaVenc.Value) Then
            MsgBox "La fecha de vencimiento no puede ser menor que la fecha del pago anterior!", vbExclamation, "Membresias"
            Me.ssdbPagos.MoveNext
            Exit Sub
        End If
    End If
    
    Me.ssdbPagos.MoveNext
    
    Me.ssdbPagos.Columns("Vence").Value = Me.dtpFechaVenc.Value
    Me.ssdbPagos.Columns("Monto").Value = CDbl(Me.txtMonto.Text)
    Me.ssdbPagos.Update
    
    CalculaTotalGrid False
    
    Me.lblCubierto = Format(dCubierto, "$#,##0.00")
    Me.lblPorCubrir = Format(dPorCubrir, "$#,##0.00")
    
    Me.txtMonto.Text = ""
    
End Sub

Private Sub ssdbPagos_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
    If Me.ssdbPagos.Rows Then
        Me.txtMonto.Text = CDbl(Me.ssdbPagos.Columns(2).Value)
        Me.dtpFechaVenc.Value = Me.ssdbPagos.Columns(1).Value
    End If
End Sub

Private Sub txtCveTit_LostFocus()
    If (Trim$(Me.txtCveTit.Text) <> "") Then
        If (IsNumeric(Me.txtCveTit.Text)) Then
            LlenaDatosSocio
            
            If (Trim(Me.txtNombre.Text) = "VACIO") Then
                MsgBox "El socio seleccionado no existe en la base de" & Chr(13) & "datos o ya tiene asignada una membresía.", vbExclamation, "KalaSystems"
                Me.txtNombre.Text = ""
                Me.txtCveTit.Text = ""
                Me.txtCveTit.SetFocus
            End If
        Else
            MsgBox "La clave del socio es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtNombre.Text = ""
            Me.txtCveTit.Text = ""
            Me.txtCveTit.SetFocus
        End If
    End If

    Me.txtCveTit.REFRESH
End Sub

Private Sub txtCveVendedor_LostFocus()
    Dim sCond As String

    If (Trim$(Me.txtCveVendedor.Text) <> "") Then
        If (IsNumeric(Me.txtCveVendedor.Text)) Then
            Me.txtVendedor.Text = LeeXValor("Nombre", "Vendedores", "idVendedor=" & Val(Me.txtCveVendedor.Text), "Nombre", "s", Conn)
            
            If (Trim(Me.txtVendedor.Text) = "VACIO") Then
                MsgBox "El vendedor seleccionado no existe en la base de datos.", vbExclamation, "KalaSystems"
                Me.txtVendedor.Text = ""
                Me.txtCveVendedor.Text = ""
                'Me.txtCveVendedor.SetFocus
            End If
        Else
            MsgBox "La clave del vendedor es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtVendedor.Text = ""
            Me.txtCveVendedor.Text = ""
            Me.txtCveVendedor.SetFocus
        End If
    End If

    Me.txtCveVendedor.REFRESH
End Sub

Private Sub LlenaDatosSocio()
    Dim rsDatos As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String
    Dim sCond As String

    sCampos = "Usuarios_Club!idMember, Usuarios_Club!Nombre, "
    sCampos = sCampos & "Usuarios_Club!A_Paterno, Usuarios_Club!A_Materno "

    sTablas = "Usuarios_Club"

    sCond = "idMember=" & Val(Me.txtCveTit.Text) & " AND (" & Val(Me.txtCveTit.Text) & " NOT IN(SELECT idMember FROM Membresias))"

    InitRecordSet rsDatos, sCampos, sTablas, sCond, "", Conn

    With rsDatos
        If (.RecordCount > 0) Then
            If (Not IsNull(.Fields("Usuarios_Club!Nombre"))) Then
                Me.txtNombre.Text = .Fields("Usuarios_Club!Nombre")
            End If
        
            If (Not IsNull(.Fields("Usuarios_Club!A_Paterno"))) Then
                Me.txtNombre.Text = Trim$(Me.txtNombre.Text) & " " & .Fields("Usuarios_Club!A_Paterno")
            End If
        
            If (Not IsNull(.Fields("Usuarios_Club!A_Materno"))) Then
                Me.txtNombre.Text = Trim$(Me.txtNombre.Text) & " " & .Fields("Usuarios_Club!A_Materno")
            End If
        Else
            Me.txtNombre.Text = "VACIO"
        End If

        .Close
    End With

    Set rsDatos = Nothing
End Sub

Private Sub ActivaCtrls(bValor As Boolean)
    With Me
        .txtDuracion.Enabled = bValor
        .txtMontoTotal.Enabled = bValor
        .txtEnganche.Enabled = bValor
        .txtNoPagos.Enabled = bValor
        .dtpFechaAlta.Enabled = bValor
        .cmdAceptar.Enabled = bValor
        .cmdCancelar.Enabled = bValor
        .chkPrimerDia.Enabled = bValor
    
        .txtMonto.Enabled = bValor
        .dtpFechaVenc.Enabled = bValor
        
        'gpo 12/10/2005
        .cmbTipoMembresia.Enabled = bValor
        .txtNombreProp.Enabled = bValor
        .txtObservaciones.Enabled = bValor
        
        
        .ssdbPagos.Enabled = bValor
        .cmdGeneraRen.Enabled = bValor
        .cmdModificar.Enabled = bValor
        .cmdOk.Enabled = bValor
        .cmdDeshacer.Enabled = bValor
        
    End With
End Sub

Private Sub ClrCtrlsTxt()
    With Me
        .txtDuracion.Text = ""
        .txtMontoTotal.Text = ""
        .txtEnganche.Text = ""
        .txtNoPagos.Text = ""
        .dtpFechaAlta.Value = Format(Date, "dd/mm/yyyy")
        'gpo 13/10/2005
        .txtNombreProp.Text = vbNullString
        .txtObservaciones.Text = vbNullString
    End With
End Sub


Private Sub CtrlsSocios(bValor As Boolean)
    Me.txtCveTit.Enabled = bValor
    Me.cmdHTitular.Enabled = bValor
End Sub


Private Sub ClrCtrlsSocios()
    Me.txtCveTit.Text = ""
    Me.txtNombre.Text = ""
End Sub


Private Sub cmdSalir_Click()
    'Cierra el formulario
    Unload Me
End Sub


Private Sub Form_Activate()
    'Me.ssdbPagos.Visible = False

    nColActiva = 3

    If Me.Tag = vbNullString Then
        'InitssdbGrid
        RefreshSSGrid
        Me.Tag = "LOADED"
    End If
    
    
    
End Sub


Private Sub Form_Load()
    'Propiedades del formulario
    With Me
        .Top = 0
        .Left = 0
        .Height = 9765
        .Width = 8445
    End With

    nPos = 0
    bSave = False
    bNvaMem = True
    
    If (sFormaAnterior = "frmAltaSocios") Then
        Me.txtCveTit.Text = frmAltaSocios.txtTitCve.Text
        Me.txtNombre.Text = Trim$(frmAltaSocios.txtTitNombre.Text) & " " & Trim$(frmAltaSocios.txtTitPaterno.Text) & " " & Trim$(frmAltaSocios.txtTitMaterno.Text)
    Else
        CtrlsSocios True
    End If
    
    AsignaVars
    
    Me.dtpFechaAlta.Value = Format(Date, "dd/mm/yyyy")
    Me.dtpFechaVenc.Value = Format(Date, "dd/mm/yyyy")
    
    LlenaComboTipos
    
    strSQL = "SELECT CT_ORIGEN_VENTA.DescripcionOrigenVenta, CT_ORIGEN_VENTA.idOrigenVenta, CT_ORIGEN_VENTA.SinCosto, CT_ORIGEN_VENTA.EsVenta"
    strSQL = strSQL & " FROM CT_ORIGEN_VENTA"
    
    #If SqlServer_ Then
        strSQL = strSQL & " WHERE CT_ORIGEN_VENTA.FechaInicial <= getDate() And CT_ORIGEN_VENTA.FechaFinal >= getDate()"
    #Else
        strSQL = strSQL & " WHERE CT_ORIGEN_VENTA.FechaInicial <= Date() And CT_ORIGEN_VENTA.FechaFinal >= Date()"
    #End If
    
    LlenaSsCombo Me.sscmbOrigenVenta, Conn, strSQL, 4
    
    txtCveVendedor.Text = iDB_IdUser
    txtVendedor.Text = NombreVendedor(iDB_IdUser)
    
    'Modificacion para que tome datos de tabla PERIODO_PAGO
    Me.cmbMantIni.Clear
    LlenaComboPeriodo
    
    'Me.cmbMantIni.AddItem "MENSUAL"
    'Me.cmbMantIni.AddItem "ANUAL"
    
    'Me.cmbMantIni.Text = "MENSUAL"
    'Me.optConv.Value = True

    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Registro de membresías"
    
    
End Sub
''Carga el periodo del Mantenimiento
Private Sub LlenaComboPeriodo()
    Dim adorcsPeriodoPago As ADODB.Recordset
    
    strSQL = "SELECT PeriodoPago, Descripcion"
    strSQL = strSQL & " FROM PERIODO_PAGO WHERE Tipo= 1 "
    strSQL = strSQL & " ORDER BY PeriodoPago"
    
    
    Set adorcsPeriodoPago = New ADODB.Recordset
    adorcsPeriodoPago.CursorLocation = adUseServer
    
    adorcsPeriodoPago.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not adorcsPeriodoPago.EOF
        Me.cmbMantIni.AddItem adorcsPeriodoPago!Descripcion
        Me.cmbMantIni.ItemData(Me.cmbMantIni.NewIndex) = adorcsPeriodoPago!PeriodoPago
        adorcsPeriodoPago.MoveNext
    Loop
    
    adorcsPeriodoPago.Close
    Set adorcsPeriodoPago = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMembresia = Nothing
    
    If (sFormaAnterior = "frmConsMembers") Then
        frmConsMembers.RefreshssMem
    End If

    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "KALACLUB"
End Sub


Private Sub InitssdbGrid()
    'Asigna valores a la matriz de encabezados
    mTitPagos(0) = "# Pago"
    mTitPagos(1) = "Vence"
    mTitPagos(2) = "Monto"
    mTitPagos(3) = "Fec.Pago"
    mTitPagos(4) = "Observaciones"
    mTitPagos(5) = "IdReg"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid Me.ssdbPagos, mTitPagos

    'Asigna valores a la matriz que define el ancho de cada columna
    mAncPagos(0) = 800
    mAncPagos(1) = 1225
    mAncPagos(2) = 1100
    mAncPagos(3) = 1100
    mAncPagos(4) = 2000
    mAncPagos(5) = 1100
    

    'Asigna el ancho de cada columna
    DefAnchossGrid Me.ssdbPagos, mAncPagos
    
    'Alinea a la derecha las columnas que contienen números
    Me.ssdbPagos.Columns(0).Alignment = ssCaptionAlignmentRight
    Me.ssdbPagos.Columns(1).Alignment = ssCaptionAlignmentRight
    Me.ssdbPagos.Columns(2).Alignment = ssCaptionAlignmentRight
    Me.ssdbPagos.Columns(3).Alignment = ssCaptionAlignmentRight
    Me.ssdbPagos.Columns(4).Alignment = ssCaptionAlignmentRight
    Me.ssdbPagos.Columns(5).Alignment = ssCaptionAlignmentRight
    
    Me.ssdbPagos.Columns(2).NumberFormat = "CURRENCY"
    
    Me.ssdbPagos.Columns(5).Visible = False
    
    Me.ssdbPagos.AllowColumnMoving = ssRelocateNotAllowed
    Me.ssdbPagos.AllowColumnSwapping = ssRelocateNotAllowed
    Me.ssdbPagos.AllowColumnShrinking = False
    Me.ssdbPagos.AllowColumnSizing = False
    
    'Evita que se puedan modificar los datos de la consulta
    Me.ssdbPagos.AllowUpdate = False
End Sub


Private Sub RefreshSSGrid()
Dim rsPagos As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String

Dim iIndex As Integer

    sCampos = "Detalle_Mem.NoPago, Detalle_Mem.FechaVence, Detalle_Mem.Monto, "
    sCampos = sCampos & " Detalle_Mem.FechaPago, Detalle_Mem.Observaciones, Detalle_Mem.IdReg, Membresias.idMember, "
    sCampos = sCampos & "Membresias.Duracion, Membresias.Monto AS MontoTotal, Membresias.Enganche, "
    sCampos = sCampos & "Membresias.FechaAlta, Membresias.NumeroPagos, Membresias.idMembresia, "
    sCampos = sCampos & "Membresias.IdTipoMembresia, Membresias.NombrePropietario, Membresias.Observaciones AS ObservacionesMem, "
    sCampos = sCampos & "Membresias.idVendedor" & ", "
    sCampos = sCampos & "Membresias.MantenimientoIni" & ","
    sCampos = sCampos & "Membresias.IdOrigenVenta"
    sTablas = "Detalle_Mem LEFT JOIN Membresias ON Detalle_Mem.idMembresia=Membresias.idMembresia"

    InitRecordSet rsPagos, sCampos, sTablas, "Membresias.idMember=" & Val(Me.txtCveTit.Text), "", Conn
    
    Me.ssdbPagos.RemoveAll
    
    With rsPagos
        If (.RecordCount > 0) Then
            .MoveFirst
            
            Me.txtDuracion.Text = .Fields("Duracion")
            Me.txtMontoTotal.Text = .Fields("MontoTotal")
            Me.txtEnganche.Text = .Fields("Enganche")
            Me.txtNoPagos.Text = .Fields("NumeroPagos")
            Me.dtpFechaAlta.Value = Format(.Fields("FechaAlta"), "dd/mm/yyyy")
            Me.txtNoMem.Text = .Fields("idMembresia")
            
            'gpo 12/10/05
            Me.txtNombreProp.Text = IIf(IsNull(.Fields("NombrePropietario")), "", Trim(.Fields("NombrePropietario")))
            Me.txtObservaciones.Text = IIf(IsNull(.Fields("ObservacionesMem")), "", Trim(.Fields("ObservacionesMem")))
            
            '20111219 UCM
            txtCveVendedor.Text = iDB_IdUser
            txtVendedor.Text = NombreVendedor(iDB_IdUser)
            '20111219 UCM
            
            Me.txtCveVendedor.Text = IIf(IsNull(.Fields("idVendedor")), "", .Fields("idVendedor"))
            'If .Fields("idMembresia") > CInt(ObtieneParametro("MAX_ULTIMA_MEMBRESIA")) Then
                If IsNumeric(txtCveVendedor.Text) Then
                    txtVendedor.Text = NombreVendedor(txtCveVendedor.Text)
                End If
            'Else
                txtCveVendedor_LostFocus
            'End If
            
            'gpo 12/10/05
            For lI = 0 To Me.cmbTipoMembresia.ListCount - 1
                If Me.cmbTipoMembresia.ItemData(lI) = .Fields("IdTipoMembresia") Then
                    Me.cmbTipoMembresia.ListIndex = lI
                    Exit For
                End If
            Next
            
            'gpo 23/01/2005
            ''Se modifica para permitir y cargar los diferentes periodos de pagos
            Select Case .Fields("MantenimientoIni")
                Case "MC"
                    'Me.cmbMantIni.Text = "MENSUAL"
                    Me.cmbMantIni.ListIndex = 0
                    Me.optConv.Value = True
                Case "MD"
                    'Me.cmbMantIni.Text = "MENSUAL"
                    Me.cmbMantIni.ListIndex = 0
                    Me.optDirec.Value = True
                Case "BC"
                    Me.cmbMantIni.ListIndex = 1
                    Me.optConv.Value = True
                Case "BD"
                    Me.cmbMantIni.ListIndex = 1
                    Me.optDirec.Value = True
                Case "TC"
                    Me.cmbMantIni.ListIndex = 2
                    Me.optConv.Value = True
                Case "TD"
                    Me.cmbMantIni.ListIndex = 2
                    Me.optDirec.Value = True
                Case "SC"
                    Me.cmbMantIni.ListIndex = 3
                    Me.optConv.Value = True
                Case "SD"
                    Me.cmbMantIni.ListIndex = 3
                    Me.optDirec.Value = True
                Case "AC"
                    'Me.cmbMantIni.Text = "ANUAL"
                    Me.cmbMantIni.ListIndex = 4
                    Me.optConv.Value = True
                Case "AD"
                    Me.cmbMantIni.ListIndex = 4
                    Me.optDirec.Value = True
                Case Else
                    'Me.cmbMantIni.Text = "MENSUAL"
                    Me.cmbMantIni.ListIndex = 1
                    Me.optConv.Value = True
            End Select
            
            
            If (.Fields("IdOrigenVenta") <> 0) Then
                BuscaSSCombo Me.sscmbOrigenVenta, .Fields("IdOrigenVenta"), 1
            
            
                If Not IsNull(Me.sscmbOrigenVenta.Columns("Descripcion").Value) Then
                    Me.sscmbOrigenVenta.Text = Me.sscmbOrigenVenta.Columns("Descripcion").Value
                End If
            End If
            
            ActivaCtrls False
            CtrlsSocios False
            
            dCubierto = 0
            
            Do While (Not .EOF)
                Me.ssdbPagos.AddItem .Fields("NoPago") & vbTab & _
                Format(.Fields("FechaVence"), "dd/mm/yyyy") & vbTab & _
                Format(.Fields("Monto"), "###,##0.000") & vbTab & _
                Format(.Fields("FechaPago"), "dd/mm/yyyy") & vbTab & _
                Trim(.Fields("Observaciones")) & vbTab & _
                .Fields("IdReg")
                
                'Si ya existe algun pago
                'cambia la bandera bYaPagada
                If Not IsNull(.Fields("FechaPago")) Then
                    bYaPagada = True
                End If
                
                dCubierto = dCubierto + .Fields("Monto")
                
                .MoveNext
            Loop
            
            dPorCubrir = CDbl(Me.txtMontoTotal.Text) - dCubierto
            
            bNvaMem = False
            
        End If
        
        

    End With
    
    rsPagos.Close
    Set rsPagos = Nothing

    'InitssdbGrid
    
    If (Me.ssdbPagos.Rows <= 0) Then
        Me.ssdbPagos.Enabled = False
        
        If bNvaMem Then
            Me.txtNombreProp.Text = Trim(Me.txtNombre.Text)
        End If
    Else
        Me.ssdbPagos.Enabled = True
        
        Me.txtNombreProp.Enabled = True
        Me.cmbTipoMembresia.Enabled = True
        Me.txtDuracion.Enabled = True
        Me.txtObservaciones.Enabled = True
        
        Me.txtMonto.Enabled = True
        Me.dtpFechaVenc.Enabled = True
        
        Me.cmdDeshacer.Enabled = True
        Me.cmdOk.Enabled = True
        Me.cmdModificar.Enabled = True
        
        Me.cmdAceptar.Enabled = True
        
    End If
    
    Me.ssdbPagos.Bookmark = nPos
    
    Me.ssdbPagos.Visible = True
    If (Me.ssdbPagos.Enabled) Then
        Me.ssdbPagos.SetFocus
    End If
    
    Me.lblCubierto = Format(dCubierto, "$#,##0.00")
    Me.lblPorCubrir = Format(dPorCubrir, "$#,##0.00")
    
End Sub


Private Function ChecaDatos()
    ChecaDatos = False
    
    If (Val(Me.txtCveTit.Text) <= 0) Then
        MsgBox "Se debe seleccionar un socio para asignar la membresía.", vbExclamation, DEVELOPER
        Me.txtCveTit.SetFocus
        Exit Function
    End If
    
    If (bNvaMem And ExisteXValor("idMember", "Membresias", "idMember=" & Val(Me.txtCveTit.Text), Conn, "")) Then
        MsgBox "Ya existe una membresía asignada a este socio, seleccione otro.", vbExclamation, DEVELOPER
        'Me.txtCveTit.SetFocus
        Exit Function
    End If
    
    If dPorCubrir <> 0 Then
        MsgBox "La suma de los pagos más el enganche" & vbLf & "no coincide con el monto total!", vbExclamation, "Membresias"
        Exit Function
    End If
    
    If Me.txtNombreProp.Text = "" Then
        MsgBox "Se debe capturar un nombre de Propietario!", vbExclamation, "Membresias"
        Exit Function
    End If
    
    If (Me.cmbTipoMembresia.Text = "") Then
        MsgBox "Se debe seleccionar un tipo de membresía.", vbExclamation, DEVELOPER
        Me.cmbTipoMembresia.SetFocus
        Exit Function
    End If
    
    'gpo
    '29/08/09
    If Me.sscmbOrigenVenta.Text = vbNullString Then
        MsgBox "Se debe seleccionar un origen del alta", vbExclamation, "Verifique"
        Me.sscmbOrigenVenta.SetFocus
        Exit Function
    End If
    

    If (IsNumeric(Me.txtDuracion.Text)) Then
        If (Val(Me.txtDuracion.Text) <= 0) Then
            MsgBox "Al menos se debe escribir un apellido.", vbExclamation, "KalaSystems"
            'Me.txtDuracion.SetFocus
            Exit Function
        End If
    Else
        MsgBox "El valor escrito en la duración es incorrecto.", vbExclamation, "KalaSystems"
        'Me.txtDuracion.SetFocus
        Exit Function
    End If
    
    If (IsNumeric(Me.txtMontoTotal.Text)) Then
        If (CDbl(Me.txtMontoTotal.Text) <= 0) Then
            MsgBox "El monto de la membresía debe ser mayor a cero.", vbExclamation, "KalaSystems"
            'Me.txtMontoTotal.SetFocus
            Exit Function
        End If
    Else
        MsgBox "El valor escrito en el monto es incorrecto.", vbExclamation, "KalaSystems"
        'Me.txtMontoTotal.SetFocus
        Exit Function
    End If
    
    If (IsNumeric(Me.txtEnganche.Text)) Then
        If (CDbl(Me.txtEnganche.Text) < 0) Then
            MsgBox "El enganche no puede ser un valor negativo.", vbExclamation, "KalaSystems"
            'Me.txtEnganche.SetFocus
            Exit Function
        Else
            If (CDbl(Me.txtEnganche.Text) > CDbl(Me.txtMontoTotal.Text)) Then
                MsgBox "El enganche no debe ser mayor que el monto de la membresía.", vbExclamation, "KalaSystems"
                'Me.txtEnganche.SetFocus
                Exit Function
            End If
        End If
    Else
        MsgBox "El valor escrito en el enganche es incorrecto.", vbExclamation, "KalaSystems"
        'Me.txtEnganche.SetFocus
        Exit Function
    End If
    
    If (IsNumeric(Me.txtNoPagos.Text)) Then
        If (CDbl(Me.txtEnganche.Text) > 0) Then
            If (Val(Me.txtNoPagos.Text) <= 0) Then
                'MsgBox "El número de pagos debe ser mayor que cero.", vbExclamation, "KalaSystems"
                'Me.txtNoPagos.SetFocus
                'Exit Function
            Else
                If (Val(Me.txtNoPagos.Text) > MAXPAGOS) Then
                    MsgBox "No se aceptan más de " & MAXPAGOS & ".", vbInformation, "KalaSystems"
                    'Me.txtNoPagos.SetFocus
                    Exit Function
                End If
            End If
        End If
    Else
        If (CDbl(Me.txtMontoTotal.Text) <> CDbl(Me.txtEnganche.Text)) Then
            MsgBox "El valor escrito en el número de pagos es incorrecto.", vbExclamation, "KalaSystems"
            'Me.txtNoPagos.SetFocus
            Exit Function
        End If
    End If
    
    'gpo 19/01/2006
    If Me.txtCveVendedor.Text = "" Then
        MsgBox "Seleccione un Vendedor!", vbExclamation, "Verifique"
        Me.txtCveVendedor.SetFocus
        Exit Function
    End If
    
    
    If Me.ssdbPagos.Rows = 0 Then
        MsgBox "Se deben generar los pagos!", vbExclamation, "Membresias"
        Exit Function
    End If
    
    
    '20111130
    If Me.cmbMantIni.Text = "" Then
        MsgBox "Seleccione un tipo de mantenimiento.", vbExclamation, "Verifique"
        Me.cmbMantIni.SetFocus
        Exit Function
    End If
    
    If Not (optConv.Value Xor optDirec.Value) Then
        MsgBox "Seleccione la forma de cobro del mantenimiento.", vbExclamation, "Verifique"
        Exit Function
    End If

    ChecaDatos = True
End Function


Private Function GuardaDatos() As Boolean
Const DATOSMEM = 14
Const DATOSDET = 7
Dim bCreado As Boolean
Dim mFieldsMem(DATOSMEM) As String
Dim mValuesMem(DATOSMEM) As Variant
Dim mFieldsDet(DATOSDET) As String
Dim mValuesDet(DATOSDET) As Variant

Dim lI As Long
Dim nParcial As Double

Dim InitTrans As Long

Dim adocmdMem As ADODB.Command


Dim lNumMem As Long
Dim lIdReg As Long

Dim sStrSql1 As String
Dim sStrSql2 As String

Dim sMantIni As String

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If
    
    
    On Error GoTo CatchError
    
    
    Err.Clear
    Conn.Errors.Clear
    
    InitTrans = Conn.BeginTrans

    'Campos de la tabla
    mFieldsMem(0) = "idMembresia"
    mFieldsMem(1) = "idMember"
    mFieldsMem(2) = "Descripcion"
    mFieldsMem(3) = "Monto"
    mFieldsMem(4) = "Enganche"
    mFieldsMem(5) = "FechaAlta"
    mFieldsMem(6) = "NumeroPagos"
    mFieldsMem(7) = "Duracion"
    mFieldsMem(8) = "IdTipoMembresia"
    mFieldsMem(9) = "NombrePropietario"
    mFieldsMem(10) = "Observaciones"
    mFieldsMem(11) = "idVendedor"
    mFieldsMem(12) = "MantenimientoIni"
    mFieldsMem(12) = "OrigenVenta"
    
    If (bNvaMem) Then
        mValuesMem(0) = LeeUltReg("Membresias", "idMembresia") + 1
    Else
        mValuesMem(0) = Val(Me.txtNoMem.Text)
    End If
    
    #If SqlServer_ Then
        mValuesMem(1) = Val(Me.txtCveTit.Text)
        mValuesMem(2) = ""
        mValuesMem(3) = Val(Me.txtMontoTotal.Text)
        mValuesMem(4) = Val(Me.txtEnganche.Text)
        mValuesMem(5) = Format(Me.dtpFechaAlta.Value, "yyyymmdd")
        mValuesMem(6) = Val(Me.txtNoPagos.Text)
        mValuesMem(7) = Val(Me.txtDuracion.Text)
        mValuesMem(8) = Me.cmbTipoMembresia.ItemData(Me.cmbTipoMembresia.ListIndex)
        mValuesMem(9) = UCase(Me.txtNombreProp.Text)
        mValuesMem(10) = UCase(RemoveLF(Me.txtObservaciones.Text))
        mValuesMem(11) = Val(Me.txtCveVendedor.Text)
    #Else
        mValuesMem(1) = Val(Me.txtCveTit.Text)
        mValuesMem(2) = ""
        mValuesMem(3) = Val(Me.txtMontoTotal.Text)
        mValuesMem(4) = Val(Me.txtEnganche.Text)
        mValuesMem(5) = Format(Me.dtpFechaAlta.Value, "dd/mm/yyyy")
        mValuesMem(6) = Val(Me.txtNoPagos.Text)
        mValuesMem(7) = Val(Me.txtDuracion.Text)
        mValuesMem(8) = Me.cmbTipoMembresia.ItemData(Me.cmbTipoMembresia.ListIndex)
        mValuesMem(9) = UCase(Me.txtNombreProp.Text)
        mValuesMem(10) = UCase(RemoveLF(Me.txtObservaciones.Text))
        mValuesMem(11) = Val(Me.txtCveVendedor.Text)
    #End If
    
    sMantIni = Left(Me.cmbMantIni, 1)
    sMantIni = sMantIni & IIf(Me.optConv.Value, "C", "D")
    
    
    mFieldsDet(0) = "idReg"
    mFieldsDet(1) = "idMembresia"
    mFieldsDet(2) = "NoPago"
    mFieldsDet(3) = "Monto"
    mFieldsDet(4) = "FechaVence"
    mFieldsDet(5) = "FechaPago"
    mFieldsDet(6) = "Observaciones"
    
    Set adocmdMem = New ADODB.Command
    adocmdMem.ActiveConnection = Conn
    adocmdMem.CommandType = adCmdText

    
    
    If (bNvaMem) Then
    
        lNumMem = LeeUltReg("Membresias", "idMembresia") + 1
    
        #If SqlServer_ Then
            sStrSql1 = "INSERT INTO MEMBRESIAS ("
            sStrSql1 = sStrSql1 & " IdMembresia,"
            sStrSql1 = sStrSql1 & " IdMember,"
            sStrSql1 = sStrSql1 & " Descripcion,"
            sStrSql1 = sStrSql1 & " Monto,"
            sStrSql1 = sStrSql1 & " Enganche,"
            sStrSql1 = sStrSql1 & " FechaAlta,"
            sStrSql1 = sStrSql1 & " NumeroPagos,"
            sStrSql1 = sStrSql1 & " Duracion,"
            sStrSql1 = sStrSql1 & " IdTipoMembresia,"
            sStrSql1 = sStrSql1 & " NombrePropietario,"
            sStrSql1 = sStrSql1 & " Observaciones,"
            sStrSql1 = sStrSql1 & " idVendedor,"
            sStrSql1 = sStrSql1 & " MantenimientoIni,"
            sStrSql1 = sStrSql1 & " IdOrigenVenta)" '29/10/09 gpo
            sStrSql1 = sStrSql1 & " VALUES ("
            sStrSql1 = sStrSql1 & lNumMem & ","
            sStrSql1 = sStrSql1 & Val(Me.txtCveTit.Text) & ","
            sStrSql1 = sStrSql1 & "'" & "',"
            sStrSql1 = sStrSql1 & CDbl(Me.txtMontoTotal.Text) & ","
            sStrSql1 = sStrSql1 & CDbl(Me.txtEnganche.Text) & ","
            sStrSql1 = sStrSql1 & "'" & Format(Me.dtpFechaAlta.Value, "yyyymmdd") & "',"
            sStrSql1 = sStrSql1 & Val(Me.txtNoPagos.Text) & ","
            sStrSql1 = sStrSql1 & Val(Me.txtDuracion.Text) & ","
            sStrSql1 = sStrSql1 & Me.cmbTipoMembresia.ItemData(Me.cmbTipoMembresia.ListIndex) & ","
            sStrSql1 = sStrSql1 & "'" & UCase(Me.txtNombreProp.Text) & "',"
            sStrSql1 = sStrSql1 & "'" & UCase(RemoveLF(Me.txtObservaciones.Text)) & "',"
            sStrSql1 = sStrSql1 & Val(Me.txtCveVendedor.Text) & ","
            sStrSql1 = sStrSql1 & "'" & sMantIni & "',"
            sStrSql1 = sStrSql1 & "'" & Me.sscmbOrigenVenta.Columns("CveOrigen").Value & "')"
        #Else
            sStrSql1 = "INSERT INTO MEMBRESIAS ("
            sStrSql1 = sStrSql1 & " IdMembresia,"
            sStrSql1 = sStrSql1 & " IdMember,"
            sStrSql1 = sStrSql1 & " Descripcion,"
            sStrSql1 = sStrSql1 & " Monto,"
            sStrSql1 = sStrSql1 & " Enganche,"
            sStrSql1 = sStrSql1 & " FechaAlta,"
            sStrSql1 = sStrSql1 & " NumeroPagos,"
            sStrSql1 = sStrSql1 & " Duracion,"
            sStrSql1 = sStrSql1 & " IdTipoMembresia,"
            sStrSql1 = sStrSql1 & " NombrePropietario,"
            sStrSql1 = sStrSql1 & " Observaciones,"
            sStrSql1 = sStrSql1 & " idVendedor,"
            sStrSql1 = sStrSql1 & " MantenimientoIni,"
            sStrSql1 = sStrSql1 & " IdOrigenVenta)" '29/10/09 gpo
            sStrSql1 = sStrSql1 & " VALUES ("
            sStrSql1 = sStrSql1 & lNumMem & ","
            sStrSql1 = sStrSql1 & Val(Me.txtCveTit.Text) & ","
            sStrSql1 = sStrSql1 & "'" & "',"
            sStrSql1 = sStrSql1 & CDbl(Me.txtMontoTotal.Text) & ","
            sStrSql1 = sStrSql1 & CDbl(Me.txtEnganche.Text) & ","
            sStrSql1 = sStrSql1 & "#" & Format(Me.dtpFechaAlta.Value, "mm/dd/yyyy") & "#,"
            sStrSql1 = sStrSql1 & Val(Me.txtNoPagos.Text) & ","
            sStrSql1 = sStrSql1 & Val(Me.txtDuracion.Text) & ","
            sStrSql1 = sStrSql1 & Me.cmbTipoMembresia.ItemData(Me.cmbTipoMembresia.ListIndex) & ","
            sStrSql1 = sStrSql1 & "'" & UCase(Me.txtNombreProp.Text) & "',"
            sStrSql1 = sStrSql1 & "'" & UCase(RemoveLF(Me.txtObservaciones.Text)) & "',"
            sStrSql1 = sStrSql1 & Val(Me.txtCveVendedor.Text) & ","
            sStrSql1 = sStrSql1 & "'" & sMantIni & "',"
            sStrSql1 = sStrSql1 & "'" & Me.sscmbOrigenVenta.Columns("CveOrigen").Value & "')"
        #End If
        
        'Registra los datos de la nueva Membresia
        adocmdMem.CommandText = sStrSql1
        adocmdMem.Execute
        
        For lI = 0 To Me.ssdbPagos.Rows - 1
        
            Me.ssdbPagos.Bookmark = lI
            
            #If SqlServer_ Then
                sStrSql2 = "INSERT INTO DETALLE_MEM ("
                sStrSql2 = sStrSql2 & " IdReg,"
                sStrSql2 = sStrSql2 & " IdMembresia,"
                sStrSql2 = sStrSql2 & " NoPago,"
                sStrSql2 = sStrSql2 & " Monto,"
                sStrSql2 = sStrSql2 & " FechaVence,"
                sStrSql2 = sStrSql2 & " FechaPago,"
                sStrSql2 = sStrSql2 & " Observaciones)"
                sStrSql2 = sStrSql2 & " VALUES ("
                sStrSql2 = sStrSql2 & LeeUltReg("Detalle_Mem", "idReg") + 1 & ","
                sStrSql2 = sStrSql2 & lNumMem & ","
                sStrSql2 = sStrSql2 & lI & ","
                sStrSql2 = sStrSql2 & Round(CDbl(Me.ssdbPagos.Columns("Monto").Value), 2) & ","
                sStrSql2 = sStrSql2 & "'" & Format(Me.ssdbPagos.Columns("Vence").Value, "yyyymmdd") & "',"
                sStrSql2 = sStrSql2 & IIf(Me.ssdbPagos.Columns("Fec.Pago").Value = "", "Null", "'" & Format(Me.ssdbPagos.Columns("Fec.Pago").Value, "yyyymmdd") & "'") & ","
                sStrSql2 = sStrSql2 & "'" & Me.ssdbPagos.Columns("Observaciones").Value & "')"
            #Else
                sStrSql2 = "INSERT INTO DETALLE_MEM ("
                sStrSql2 = sStrSql2 & " IdReg,"
                sStrSql2 = sStrSql2 & " IdMembresia,"
                sStrSql2 = sStrSql2 & " NoPago,"
                sStrSql2 = sStrSql2 & " Monto,"
                sStrSql2 = sStrSql2 & " FechaVence,"
                sStrSql2 = sStrSql2 & " FechaPago,"
                sStrSql2 = sStrSql2 & " Observaciones)"
                sStrSql2 = sStrSql2 & " VALUES ("
                sStrSql2 = sStrSql2 & LeeUltReg("Detalle_Mem", "idReg") + 1 & ","
                sStrSql2 = sStrSql2 & lNumMem & ","
                sStrSql2 = sStrSql2 & lI & ","
                sStrSql2 = sStrSql2 & Round(CDbl(Me.ssdbPagos.Columns("Monto").Value), 2) & ","
                sStrSql2 = sStrSql2 & "#" & Format(Me.ssdbPagos.Columns("Vence").Value, "mm/dd/yyyy") & "#,"
                sStrSql2 = sStrSql2 & IIf(Me.ssdbPagos.Columns("Fec.Pago").Value = "", "Null", "'" & Format(Me.ssdbPagos.Columns("Fec.Pago").Value, "dd/mm/yyyy") & "'") & ","
                sStrSql2 = sStrSql2 & "'" & Me.ssdbPagos.Columns("Observaciones").Value & "')"
            #End If
            
            adocmdMem.CommandText = sStrSql2
            adocmdMem.Execute
        
        Next lI
                    
        Conn.CommitTrans
        GuardaDatos = True
            
        bSave = True
        bNvaMem = False
        
        MsgBox "Se dieron de alta los datos correctamente", vbExclamation, "Membresias"
        
    Else
        If (Val(Me.txtNoMem.Text) > 0) Then
        
            lNumMem = Val(Me.txtNoMem.Text)
        
            #If SqlServer_ Then
                sStrSql1 = "UPDATE MEMBRESIAS SET "
                sStrSql1 = sStrSql1 & " Descripcion=" & "''" & ","
                sStrSql1 = sStrSql1 & " Monto=" & CDbl(Me.txtMontoTotal.Text) & ","
                sStrSql1 = sStrSql1 & " Enganche=" & CDbl(Me.txtEnganche.Text) & ","
                sStrSql1 = sStrSql1 & " FechaAlta=" & "'" & Format(Me.dtpFechaAlta.Value, "yyyymmdd") & "',"
                sStrSql1 = sStrSql1 & " NumeroPagos=" & Val(Me.txtNoPagos.Text) & ","
                sStrSql1 = sStrSql1 & " Duracion=" & Val(Me.txtDuracion.Text) & ","
                sStrSql1 = sStrSql1 & " IdTipoMembresia=" & Me.cmbTipoMembresia.ItemData(Me.cmbTipoMembresia.ListIndex) & ","
                sStrSql1 = sStrSql1 & " NombrePropietario=" & "'" & UCase(Me.txtNombreProp.Text) & "',"
                sStrSql1 = sStrSql1 & " Observaciones=" & "'" & UCase(RemoveLF(Me.txtObservaciones.Text)) & "',"
                sStrSql1 = sStrSql1 & " idVendedor=" & Val(Me.txtCveVendedor.Text) & ","
                sStrSql1 = sStrSql1 & " MantenimientoIni=" & "'" & sMantIni & "',"
                sStrSql1 = sStrSql1 & " IdOrigenVenta=" & "'" & Me.sscmbOrigenVenta.Columns("CveOrigen").Value & "'" 'gpo 29/08/2009
                sStrSql1 = sStrSql1 & " WHERE "
                sStrSql1 = sStrSql1 & " (IdMembresia=" & Trim(Me.txtNoMem.Text) & ")"
                sStrSql1 = sStrSql1 & " AND (IdMember=" & Trim(Me.txtCveTit.Text) & ")"
                
                
                strSQL = "DELETE FROM DETALLE_MEM"
                strSQL = strSQL & " WHERE IdMembresia=" & lNumMem
            #Else
                sStrSql1 = "UPDATE MEMBRESIAS SET "
                sStrSql1 = sStrSql1 & " Descripcion=" & "''" & ","
                sStrSql1 = sStrSql1 & " Monto=" & CDbl(Me.txtMontoTotal.Text) & ","
                sStrSql1 = sStrSql1 & " Enganche=" & CDbl(Me.txtEnganche.Text) & ","
                sStrSql1 = sStrSql1 & " FechaAlta=" & "'" & Format(Me.dtpFechaAlta.Value, "mm/dd/yyyy") & "',"
                sStrSql1 = sStrSql1 & " NumeroPagos=" & Val(Me.txtNoPagos.Text) & ","
                sStrSql1 = sStrSql1 & " Duracion=" & Val(Me.txtDuracion.Text) & ","
                sStrSql1 = sStrSql1 & " IdTipoMembresia=" & Me.cmbTipoMembresia.ItemData(Me.cmbTipoMembresia.ListIndex) & ","
                sStrSql1 = sStrSql1 & " NombrePropietario=" & "'" & UCase(Me.txtNombreProp.Text) & "',"
                sStrSql1 = sStrSql1 & " Observaciones=" & "'" & UCase(RemoveLF(Me.txtObservaciones.Text)) & "',"
                sStrSql1 = sStrSql1 & " idVendedor=" & Val(Me.txtCveVendedor.Text) & ","
                sStrSql1 = sStrSql1 & " MantenimientoIni=" & "'" & sMantIni & "',"
                sStrSql1 = sStrSql1 & " IdOrigenVenta=" & "'" & Me.sscmbOrigenVenta.Columns("CveOrigen").Value & "'" 'gpo 29/08/2009
                sStrSql1 = sStrSql1 & " WHERE "
                sStrSql1 = sStrSql1 & " (IdMembresia=" & Trim(Me.txtNoMem.Text) & ")"
                sStrSql1 = sStrSql1 & " AND (IdMember=" & Trim(Me.txtCveTit.Text) & ")"
                

                strSQL = "DELETE * FROM DETALLE_MEM"
                strSQL = strSQL & " WHERE IdMembresia=" & lNumMem
            #End If

            
            'Actualiza los datos
            adocmdMem.CommandText = sStrSql1
            adocmdMem.Execute
                
            adocmdMem.CommandText = strSQL
            adocmdMem.Execute
                
            For lI = 0 To Me.ssdbPagos.Rows - 1
                Me.ssdbPagos.Bookmark = lI
                'Registro existente
                If Val(Me.ssdbPagos.Columns("IdReg").Value) > 0 Then
                    lIdReg = Val(Me.ssdbPagos.Columns("IdReg").Value)
                Else
                    lIdReg = LeeUltReg("Detalle_Mem", "idReg") + 1
                End If
                    
                #If SqlServer_ Then
                    sStrSql2 = "INSERT INTO DETALLE_MEM ("
                    sStrSql2 = sStrSql2 & " IdReg,"
                    sStrSql2 = sStrSql2 & " IdMembresia,"
                    sStrSql2 = sStrSql2 & " NoPago,"
                    sStrSql2 = sStrSql2 & " Monto,"
                    sStrSql2 = sStrSql2 & " FechaVence,"
                    sStrSql2 = sStrSql2 & " FechaPago,"
                    sStrSql2 = sStrSql2 & " Observaciones)"
                    sStrSql2 = sStrSql2 & " VALUES ("
                    sStrSql2 = sStrSql2 & lIdReg & ","
                    sStrSql2 = sStrSql2 & lNumMem & ","
                    sStrSql2 = sStrSql2 & lI & ","
                    sStrSql2 = sStrSql2 & Round(CDbl(Me.ssdbPagos.Columns("Monto").Value), 2) & ","
                    sStrSql2 = sStrSql2 & "'" & Format(Me.ssdbPagos.Columns("Vence").Value, "yyyymmdd") & "',"
                    sStrSql2 = sStrSql2 & IIf(Me.ssdbPagos.Columns("Fec.Pago").Value = "", "Null", "'" & Format(Me.ssdbPagos.Columns("Fec.Pago").Value, "yyyymmdd") & "'") & ","
                    sStrSql2 = sStrSql2 & "'" & Me.ssdbPagos.Columns("Observaciones").Value & "')"
                #Else
                    sStrSql2 = "INSERT INTO DETALLE_MEM ("
                    sStrSql2 = sStrSql2 & " IdReg,"
                    sStrSql2 = sStrSql2 & " IdMembresia,"
                    sStrSql2 = sStrSql2 & " NoPago,"
                    sStrSql2 = sStrSql2 & " Monto,"
                    sStrSql2 = sStrSql2 & " FechaVence,"
                    sStrSql2 = sStrSql2 & " FechaPago,"
                    sStrSql2 = sStrSql2 & " Observaciones)"
                    sStrSql2 = sStrSql2 & " VALUES ("
                    sStrSql2 = sStrSql2 & lIdReg & ","
                    sStrSql2 = sStrSql2 & lNumMem & ","
                    sStrSql2 = sStrSql2 & lI & ","
                    sStrSql2 = sStrSql2 & Round(CDbl(Me.ssdbPagos.Columns("Monto").Value), 2) & ","
                    sStrSql2 = sStrSql2 & "#" & Format(Me.ssdbPagos.Columns("Vence").Value, "mm/dd/yyyy") & "#,"
                    sStrSql2 = sStrSql2 & IIf(Me.ssdbPagos.Columns("Fec.Pago").Value = "", "Null", "#" & Format(Me.ssdbPagos.Columns("Fec.Pago").Value, "mm/dd/yyyy") & "#") & ","
                    sStrSql2 = sStrSql2 & "'" & Me.ssdbPagos.Columns("Observaciones").Value & "')"
                #End If
                
                adocmdMem.CommandText = sStrSql2
                adocmdMem.Execute
                    
            Next
            Conn.CommitTrans
            GuardaDatos = True
            MsgBox "Se modificaron los datos correctamente", vbExclamation, "Membresias"
            
        End If
    End If
    
    Set adocmdMem = Nothing
    
    Exit Function
CatchError:
    If InitTrans > 0 Then
        Conn.RollbackTrans
    End If
    
    Dim sErrores As String
    
    If Conn.Errors.Count > 0 Then
        sErrores = Conn.Errors.Item(0).Description
    End If
    
    
    MsgBox "Ocurrio el error: " & sErrores, vbCritical, "Membresias"

End Function


Private Function CalculaFecha(dFecha As Date, nIncMeses As Byte, nDiaInicial As Byte) As Date
Dim nDia As String
Dim nAnio As Integer
Dim nMes As Byte

    If (nDiaInicial = 1) Then
        nDia = 1
    Else
        nDia = Day(dFecha)
    End If
    
    nAnio = Year(dFecha)
    
    nMes = Month(dFecha) + nIncMeses
    If (nMes > 12) Then
        nMes = nMes - 12
        nAnio = nAnio + 1
    End If
    
    'Calcula el día para los meses que no tienen 31 días
    If (nDia > 28) Then
        Select Case nMes
            'Meses con 30 días
            Case 4, 6, 9, 11
                If (nDia > 30) Then
                    nDia = 30
                End If
            
            'Febrero
            Case 2
                If ((nAnio Mod 4) = 0) Then
                    If (nDia > 29) Then
                        nDia = 29
                    End If
                Else
                    If (nDia > 28) Then
                        nDia = 28
                    End If
                End If
        End Select
    End If
    
    CalculaFecha = Format(CDate(Str(nDia) & "/" & Str(nMes) & "/" & Str(nAnio)), "dd/mm/yyyy")
End Function


'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    Cambios = True
    
    If (nDuracion <> Val(Me.txtDuracion.Text)) Then
        Exit Function
    End If
    
    If Me.txtMontoTotal.Text = "" Then Me.txtMontoTotal.Text = 0
    
    If (nMontoTotal <> CDbl(Me.txtMontoTotal.Text)) Then
        Exit Function
    End If
    
    If Me.txtEnganche.Text = "" Then Me.txtEnganche.Text = 0
    
    If (nEnganche <> CDbl(Me.txtEnganche.Text)) Then
        Exit Function
    End If
    
    If Me.txtNoPagos.Text = "" Then Me.txtNoPagos.Text = 0
    
    If (nPagos <> CDbl(Me.txtNoPagos.Text)) Then
        Exit Function
    End If

    Cambios = False
End Function



'************************************************************
'*                          Ayudas                          *
'************************************************************


Private Sub cmdHTitular_Click()
Const DATOSTIT = 4
Dim sCadena As String
Dim mFAyuda(DATOSTIT) As String
Dim mAAyuda(DATOSTIT) As Integer
Dim mCAyuda(DATOSTIT) As String
Dim mEAyuda(DATOSTIT) As String

    nAyuda = 1

    Set frmHTit = New frmayuda

    mFAyuda(0) = "Titulares ordenados por clave"
    mFAyuda(1) = "Titulares ordenados por nombre"
    mFAyuda(2) = "Titulares ordenados por A. paterno"
    mFAyuda(3) = "Titulares ordenados por A. materno"

    mAAyuda(0) = 800
    mAAyuda(1) = 2500
    mAAyuda(2) = 2500
    mAAyuda(3) = 2500

    mCAyuda(0) = "idMember"
    mCAyuda(1) = "Nombre"
    mCAyuda(2) = "A_Paterno"
    mCAyuda(3) = "A_Materno"

    mEAyuda(0) = "# Tit."
    mEAyuda(1) = "Nombre"
    mEAyuda(2) = "A. Paterno"
    mEAyuda(3) = "A. Materno"

    With frmHTit
        .nColActiva = 1
        .nColsAyuda = DATOSTIT
        .sTabla = "Usuarios_Club"

        .sCondicion = "idMember=idTitular AND (" & Val(Me.txtCveTit.Text) & "NOT IN(SELECT idMember FROM Membresias))"
        .sTitAyuda = "Titulares"
        .lAgregar = False

        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda

        .Show (1)
    End With

    Me.txtCveTit.SetFocus
End Sub


Private Sub cmdHVendedor_Click()
Const DATOSVEN = 2
Dim sCadena As String
Dim mFAyuda(DATOSVEN) As String
Dim mAAyuda(DATOSVEN) As Integer
Dim mCAyuda(DATOSVEN) As String
Dim mEAyuda(DATOSVEN) As String

    nAyuda = 2

    Set frmHven = New frmayuda

    mFAyuda(0) = "Vendedores ordenados por clave"
    mFAyuda(1) = "Vendedores ordenados por nombre"

    mAAyuda(0) = 800
    mAAyuda(1) = 3500

    mCAyuda(0) = "idVendedor"
    mCAyuda(1) = "Nombre"

    mEAyuda(0) = "# Ven."
    mEAyuda(1) = "Nombre"

    With frmHven
        .nColActiva = 1
        .nColsAyuda = DATOSVEN
        .sTabla = "Vendedores"

        .sCondicion = ""
        .sTitAyuda = "Vendedores"
        .lAgregar = False

        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda

        .Show (1)
    End With

    Me.txtCveVendedor.SetFocus
End Sub


Private Sub LlenaComboTipos()
    Dim adorcsTipoMem As ADODB.Recordset
    
    strSQL = "SELECT IdTipoMembresia, Descripcion"
    strSQL = strSQL & " FROM TIPO_MEMBRESIA"
    strSQL = strSQL & " ORDER BY IdTipoMembresia"
    
    Set adorcsTipoMem = New ADODB.Recordset
    adorcsTipoMem.CursorLocation = adUseServer
   
    adorcsTipoMem.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.cmbTipoMembresia.Clear
    Do Until adorcsTipoMem.EOF
        Me.cmbTipoMembresia.AddItem Trim(adorcsTipoMem!Descripcion)
        Me.cmbTipoMembresia.ItemData(Me.cmbTipoMembresia.NewIndex) = adorcsTipoMem!IdTipoMembresia
        adorcsTipoMem.MoveNext
    Loop

    adorcsTipoMem.Close
    Set adorcsTipoMem = Nothing
    
End Sub

Private Sub GeneraRenglones()
    Dim byI As Byte
    Dim dParcial As Double
    
    
    
    
    Me.ssdbPagos.RemoveAll
    
    'InitssdbGrid
    
    dCubierto = 0
    
    
    Me.ssdbPagos.AddItem 0 & vbTab & Me.dtpFechaAlta.Value & vbTab & Me.txtEnganche
    
    dCubierto = CDbl(Me.txtEnganche.Text)
    
    If CDbl(Me.txtMontoTotal.Text) = CDbl(Me.txtEnganche.Text) Then
        dPorCubrir = CDbl(Me.txtMontoTotal.Text) - dCubierto
        Exit Sub
    End If
    
    dParcial = Round((CDbl(Me.txtMontoTotal.Text) - CDbl(Me.txtEnganche.Text)) / CDbl(Me.txtNoPagos.Text), 2)
    
    For byI = 1 To CDbl(Me.txtNoPagos.Text)
        If byI = CDbl(Me.txtNoPagos.Text) And dCubierto + dParcial <> CDbl(Me.txtMontoTotal.Text) Then
            dParcial = CDbl(Me.txtMontoTotal.Text) - dCubierto
        End If
        Me.ssdbPagos.AddItem byI & vbTab & CalculaFecha(Me.dtpFechaAlta.Value, byI, Me.chkPrimerDia.Value) & vbTab & dParcial
        dCubierto = dCubierto + dParcial
    Next
    
    dPorCubrir = CDbl(Me.txtMontoTotal.Text) - dCubierto
    
End Sub

Private Sub CalculaTotalGrid(boReenumera As Boolean)
    
    Dim lI As Long
    Dim vBookMark As Variant
    
    
    dCubierto = 0
    
    'Guarda el bookmark actual
    vBookMark = Me.ssdbPagos.Bookmark
    
    
    For lI = 0 To Me.ssdbPagos.Rows - 1
        Me.ssdbPagos.Bookmark = Me.ssdbPagos.AddItemBookmark(lI)
        dCubierto = dCubierto + CDbl(Me.ssdbPagos.Columns("Monto").Value)
        If boReenumera Then
            Me.ssdbPagos.Columns("# Pago").Value = Me.ssdbPagos.AddItemRowIndex(Me.ssdbPagos.Bookmark)
        End If
    Next
    Me.ssdbPagos.Update
    dPorCubrir = CDbl(Me.txtMontoTotal.Text) - dCubierto
    
    
    Me.ssdbPagos.Bookmark = vBookMark
    
    
End Sub


Private Sub txtDuracion_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtEnganche_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtEnganche.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtMonto.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub



Private Sub txtMontoTotal_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtMontoTotal.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtNoPagos_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select

End Sub
