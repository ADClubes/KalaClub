VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmConsMembers 
   Caption         =   "Membresías"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   Icon            =   "frmConsMembers.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   10125
   Begin VB.CommandButton cmdFrente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      Picture         =   "frmConsMembers.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   615
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   405
      Left            =   1680
      TabIndex        =   13
      Top             =   5520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   57475073
      CurrentDate     =   38532
   End
   Begin VB.TextBox txtNoPagos 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox txtEnganche 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   11
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   10
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox txtDuracion 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmdBaja 
      Enabled         =   0   'False
      Height          =   615
      Left            =   6480
      Picture         =   "frmConsMembers.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Baja "
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtBuscar 
      Height          =   300
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdBuscar 
      Default         =   -1  'True
      Height          =   615
      Left            =   3360
      Picture         =   "frmConsMembers.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Buscar "
      Top             =   120
      Width           =   615
   End
   Begin VB.CheckBox chkDesdeInicio 
      Caption         =   "&Buscar desde el primer registro"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmdNuevo 
      Enabled         =   0   'False
      Height          =   615
      Left            =   7920
      Picture         =   "frmConsMembers.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Nuevos datos "
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdModificar 
      Height          =   615
      Left            =   5760
      Picture         =   "frmConsMembers.frx":1BEA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Modificar "
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdRefresh 
      Height          =   615
      Left            =   8640
      Picture         =   "frmConsMembers.frx":202C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Actualizar "
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   9360
      Picture         =   "frmConsMembers.frx":2336
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Salir "
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtNoRegs 
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
      Left            =   4920
      TabIndex        =   15
      Top             =   5640
      Width           =   615
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbMembresia 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9855
      _Version        =   196616
      DataMode        =   2
      Cols            =   11
      Col.Count       =   11
      SelectTypeRow   =   1
      BackColorOdd    =   16777152
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   17383
      _ExtentY        =   5741
      _StockProps     =   79
      Caption         =   "Membresias"
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbDetalle 
      Height          =   1695
      Left            =   5760
      TabIndex        =   14
      Top             =   4200
      Width           =   4215
      _Version        =   196616
      DataMode        =   2
      Cols            =   3
      Col.Count       =   3
      SelectTypeRow   =   1
      BackColorOdd    =   12648384
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   7435
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Detalle de pagos"
   End
   Begin VB.Label lblFechaAlta 
      Caption         =   "Fecha de alta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   22
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblNoPagos 
      Caption         =   "No. de pagos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblEnganche 
      Caption         =   "Enganche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   20
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label lblMonto 
      Caption         =   "Total de la membresía"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   19
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label lblDuracion 
      Caption         =   "Duración (años)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblRegs 
      Caption         =   "No. de membresías"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblNotas 
      Caption         =   "Haga click sobre el encabezado de la columna para cambiar el orden"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   4120
      Width           =   5295
   End
End
Attribute VB_Name = "frmConsMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'*  Formulario de membresias                *
'*  Daniel Hdez                             *
'*  29 / Junio / 2005                       *
'********************************************

'*  Ult actualización:  09 / Septiembre / 2005


'Numero de columnas mostradas
Const TOTCOLUMNS = 11

'Encabezados para la consulta
Dim mTitMembers(TOTCOLUMNS) As String

'Ancho de las columnas mostradas
Dim mAncMembers(TOTCOLUMNS) As Integer

'Nombre de los campos en la consulta
Dim mCamMembers(TOTCOLUMNS) As String

'Matriz de los encabezados del formulario
Dim mTitForma(TOTCOLUMNS) As String

'Columna activa o seleccionada para establecer el orden
Dim nColActiva As Integer

'Posicion del cursor dentro del dbGrid
Dim nPos As Variant



Private Sub cmdBuscar_Click()
Dim i As Variant
Dim nPosIni As Integer
Dim sBuscar As String

    If (Trim$(Me.txtBuscar.Text) = "") Then
        Me.txtBuscar.SetFocus
        Exit Sub
    End If

    If (Me.chkDesdeInicio.Value) Then
        Me.ssdbMembresia.MoveFirst
    Else
        If (Me.ssdbMembresia.Row < Me.ssdbMembresia.Rows) Then
            Me.ssdbMembresia.MoveNext
        Else
            Me.ssdbMembresia.MoveFirst
        End If
    End If
    
    sBuscar = UCase$(Me.txtBuscar.Text)

    nPosIni = Me.ssdbMembresia.Row
    i = Me.ssdbMembresia.Row
    
    Do
        If (InStr(1, Trim$(Me.ssdbMembresia.Columns(nColActiva).CellValue(i)), sBuscar)) Then
            Me.ssdbMembresia.Bookmark = i
            Exit Do
        End If
        
        i = i + 1
    Loop Until i > Me.ssdbMembresia.Rows
    
    If (i > Me.ssdbMembresia.Rows) Then
        MsgBox "No se encontraron más ocurrencias del texto buscado.", vbExclamation, DEVELOPER
        Me.ssdbMembresia.Bookmark = (nPosIni - 1)
    End If

End Sub


Private Sub cmdFrente_Click()
    nPos = Me.ssdbMembresia.Bookmark

    
    frmRepMembresia.nidTitular = CSng(Me.ssdbMembresia.Columns("# Titular").Value)
    frmRepMembresia.nTotalMem = CDbl(Me.ssdbDetalle.Columns("Monto").Value)
    
    Load frmRepMembresia
    frmRepMembresia.Show
    
    Me.ssdbMembresia.Bookmark = nPos
End Sub


Private Sub cmdModificar_Click()
    If (Me.ssdbMembresia.Rows > 0) Then
        nPos = Me.ssdbMembresia.Bookmark
    End If

'    frmAltaPac.bNvoPaciente = False
'    Load frmAltaPac
'    frmAltaPac.Show (1)
End Sub


'Private Sub cmdBorrar_Click()
'Dim nOk As Long
''Dim InitTrans As Long
'
'    If (Me.ssdbMembresia.Rows > 0) Then
'
'        nOk = MsgBox("¿Desea borrar el producto seleccionado?", vbYesNo, DEVELOPER)
'
'        If (nOk = vbYes) Then
'
''            InitTrans = cnConexion.BeginTrans
'
''            If (borrareceta(Me.ssdbMembresia.Columns(0).Text)) Then
'                If (EliminaReg("Productos", "idProducto=" & Me.ssdbMembresia.Columns(0).Text, "", cnConexion)) Then
''                    cnConexion.CommitTrans
'
'                    RefreshssMem
'                Else
'                    MsgBox "No se borró el registro, intentelo de nuevo.", vbInformation, DEVELOPER
'                End If
''            End If
'        End If
'
'        Me.ssdbMembresia.SetFocus
'    End If
'End Sub


Private Sub cmdNuevo_Click()
    If (Me.ssdbMembresia.Rows > 0) Then
        nPos = Me.ssdbMembresia.Bookmark
    End If

    frmMembresia.sFormaAnterior = "frmConsMembers"
    Load frmMembresia
    frmMembresia.Show
    
'    RefreshssMem
End Sub


Private Sub cmdRefresh_Click()
    RefreshssMem
End Sub


Private Sub cmdSalir_Click()
    'Cierra el formulario
    Unload Me
End Sub


Private Sub Form_Activate()
    Me.ssdbMembresia.Visible = False

    nColActiva = 0

    'Encabezados del formulario
    mTitForma(0) = " # Mem."
    mTitForma(1) = " # fam."
    mTitForma(2) = " nombre"
    mTitForma(3) = " A. paterno"
    mTitForma(4) = " A. materno"
    mTitForma(5) = " # titular"
    mTitForma(6) = " duración"
    mTitForma(7) = " monto"
    mTitForma(8) = " enganche"
    mTitForma(9) = " fecha alta"
    mTitForma(10) = " No. pagos"
    
    mCamMembers(0) = "Membresias!idMembresia"
    mCamMembers(1) = "Usuarios_Club!NoFamilia"
    mCamMembers(2) = "Usuarios_Club!Nombre"
    mCamMembers(3) = "Usuarios_Club!A_Paterno"
    mCamMembers(4) = "Usuarios_Club!A_Materno"
    mCamMembers(5) = "Membresias!idMember"
    mCamMembers(6) = "Membresias!Duracion"
    mCamMembers(7) = "Membresias!Monto"
    mCamMembers(8) = "Membresias!Enganche"
    mCamMembers(9) = "Membresias!FechaAlta"
    mCamMembers(10) = "Membresias!NumeroPagos"
    
    RefreshssMem
    
    Me.ssdbMembresia.Visible = True
    
    Me.WindowState = 0
End Sub


Private Sub Form_Load()
    'Propiedades del formulario
    With Me
        .Top = 0
        .Left = 0
        .Height = 6510
        .Width = 10245
    End With

    nPos = 0
    bSave = False
    bNvoDato = True

    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Consulta de membresías"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmConsMembers = Nothing

    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "KALACLUB"
End Sub


Private Sub ssdbMembresia_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    If (Me.ssdbMembresia.Rows > 0) Then
        Me.txtDuracion.Text = Me.ssdbMembresia.Columns(6).Value
        Me.txtMonto.Text = Format(Me.ssdbMembresia.Columns(7).Value, "###,##0.000")
        Me.txtEnganche.Text = Format(Me.ssdbMembresia.Columns(8).Value, "###,##0.000")
        Me.dtpFecha.Value = Format(Me.ssdbMembresia.Columns(9).Value, "dd/mm/yyyy")
        Me.txtNoPagos.Text = Me.ssdbMembresia.Columns(10).Value
    
        'If (Me.ssdbMembresia.Columns(10).Value > 0) Then
            VistaDet
        'Else
            'Me.ssdbDetalle.RemoveAll
        'End If
    End If
End Sub


Private Sub InitssMembers()
    'Asigna valores a la matriz de encabezados
    mTitMembers(0) = "# Memb."
    mTitMembers(1) = "# familia"
    mTitMembers(2) = "Nombre"
    mTitMembers(3) = "A. Paterno"
    mTitMembers(4) = "A. Materno"
    mTitMembers(5) = "# Titular"
    mTitMembers(6) = "Duración"
    mTitMembers(7) = "Monto"
    mTitMembers(8) = "Enganche"
    mTitMembers(9) = "Fecha alta"
    mTitMembers(10) = "# pagos"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid Me.ssdbMembresia, mTitMembers

    'Asigna valores a la matriz que define el ancho de cada columna
    mAncMembers(0) = 900
    mAncMembers(1) = 900
    mAncMembers(2) = 2500
    mAncMembers(3) = 2500
    mAncMembers(4) = 2500
    mAncMembers(5) = 800
    mAncMembers(6) = 900
    mAncMembers(7) = 1300
    mAncMembers(8) = 1300
    mAncMembers(9) = 1100
    mAncMembers(10) = 900

    'Asigna el ancho de cada columna
    DefAnchossGrid Me.ssdbMembresia, mAncMembers
    
    'Alinea a la derecha las columnas que contienen números
    Me.ssdbMembresia.Columns(0).Alignment = ssCaptionAlignmentRight
    Me.ssdbMembresia.Columns(1).Alignment = ssCaptionAlignmentRight
    Me.ssdbMembresia.Columns(5).Alignment = ssCaptionAlignmentRight
    Me.ssdbMembresia.Columns(6).Alignment = ssCaptionAlignmentRight
    Me.ssdbMembresia.Columns(7).Alignment = ssCaptionAlignmentRight
    Me.ssdbMembresia.Columns(8).Alignment = ssCaptionAlignmentRight
    Me.ssdbMembresia.Columns(9).Alignment = ssCaptionAlignmentRight
    Me.ssdbMembresia.Columns(10).Alignment = ssCaptionAlignmentRight
    
    Me.ssdbMembresia.AllowColumnMoving = ssRelocateNotAllowed
    Me.ssdbMembresia.AllowColumnSwapping = ssRelocateNotAllowed
    Me.ssdbMembresia.AllowColumnShrinking = False
    Me.ssdbMembresia.AllowColumnSizing = False
    
    'Evita que se puedan modificar los datos de la consulta
    Me.ssdbMembresia.AllowUpdate = False
    
    'Encabezados de los DbGrid
    Me.ssdbMembresia.Caption = "Membresías (" & mTitMembers(nColActiva) & ")"
End Sub


Private Sub VistaDet()
Dim rsDetalle As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String

    sCampos = "NoPago, FechaVence, Monto "
    
    sTablas = "Detalle_Mem"

    Me.ssdbDetalle.RemoveAll

    InitRecordSet rsDetalle, sCampos, sTablas, "idMembresia=" & Me.ssdbMembresia.Columns(0).Value, "", Conn

    With rsDetalle
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                Me.ssdbDetalle.AddItem .Fields("NoPago") & vbTab & _
                Format(.Fields("FechaVence"), "dd/mm/yyyy") & vbTab & _
                Format(.Fields("Monto"), "###,##0.000")
                
                .MoveNext
            Loop
        
        End If
        
        .Close
    End With
    
    Set rsDetalle = Nothing

    InitssDetalle
End Sub


Private Sub InitssDetalle()
Const DATOSDET = 3
Dim mTitDet(DATOSDET) As String
Dim mAncDet(DATOSDET) As Integer

    'Asigna valores a la matriz de encabezados
    mTitDet(0) = "# Pago"
    mTitDet(1) = "Vence"
    mTitDet(2) = "Monto"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid Me.ssdbDetalle, mTitDet

    'Asigna valores a la matriz que define el ancho de cada columna
    mAncDet(0) = 800
    mAncDet(1) = 1225
    mAncDet(2) = 1575

    'Asigna el ancho de cada columna
    DefAnchossGrid Me.ssdbDetalle, mAncDet
    
    'Alinea a la derecha las columnas que contienen números
    Me.ssdbDetalle.Columns(0).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(1).Alignment = ssCaptionAlignmentRight
    Me.ssdbDetalle.Columns(2).Alignment = ssCaptionAlignmentRight
    
    Me.ssdbDetalle.AllowColumnMoving = ssRelocateNotAllowed
    Me.ssdbDetalle.AllowColumnSwapping = ssRelocateNotAllowed
    Me.ssdbDetalle.AllowColumnShrinking = False
    Me.ssdbDetalle.AllowColumnSizing = False
    
    'Evita que se puedan modificar los datos de la consulta
    Me.ssdbDetalle.AllowUpdate = False
End Sub


Public Sub RefreshssMem()
Dim rsMember As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String

    sCampos = "Membresias!idMembresia, Usuarios_Club!NoFamilia, "
    sCampos = sCampos & "Usuarios_Club!Nombre, Usuarios_Club!A_Paterno, Usuarios_Club!A_Materno, "
    sCampos = sCampos & "Membresias!idMember, Membresias!Duracion, Membresias!Monto, "
    sCampos = sCampos & "Membresias!Enganche, Membresias!FechaAlta, Membresias!NumeroPagos "
    
    sTablas = "Membresias LEFT JOIN Usuarios_Club ON Membresias.idMember=Usuarios_Club.idMember "

    InitRecordSet rsMember, sCampos, sTablas, "", mCamMembers(nColActiva), Conn
    
    Me.ssdbMembresia.RemoveAll
    
    With rsMember
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                Me.ssdbMembresia.AddItem .Fields("Membresias!idMembresia") & vbTab & _
                .Fields("Usuarios_Club!NoFamilia") & vbTab & _
                .Fields("Usuarios_Club!Nombre") & vbTab & _
                .Fields("Usuarios_Club!A_Paterno") & vbTab & _
                .Fields("Usuarios_Club!A_Materno") & vbTab & _
                .Fields("Membresias!idMember") & vbTab & _
                .Fields("Membresias!Duracion") & vbTab & _
                Format(.Fields("Membresias!Monto"), "###,##0.000") & vbTab & _
                IIf(Not IsNull(.Fields("Membresias!Enganche")), Format(.Fields("Membresias!Enganche"), "###,##0.000"), 0) & vbTab & _
                Format(.Fields("Membresias!FechaAlta"), "dd/mm/yyyy") & vbTab & _
                IIf(Not IsNull(.Fields("Membresias!NumeroPagos")), .Fields("Membresias!NumeroPagos"), 0)

                .MoveNext
            Loop
        End If
    End With
    
    rsMember.Close
    Set rsMember = Nothing

    InitssMembers
    
    If (Me.ssdbMembresia.Rows <= 0) Then
        Me.ssdbMembresia.Enabled = False
    Else
        Me.txtNoRegs.Text = Me.ssdbMembresia.Rows
        Me.txtNoRegs.REFRESH
        
        Me.ssdbMembresia.Enabled = True
    End If
    
    If (Me.ssdbMembresia.Rows > 0) Then
        Me.ssdbMembresia.Bookmark = nPos
    End If
    
    Me.ssdbMembresia.Visible = True
    If (Me.ssdbMembresia.Enabled) Then
        Me.ssdbMembresia.SetFocus
    End If
End Sub


Private Sub ssdbMembresia_HeadClick(ByVal ColIndex As Integer)
    nColActiva = ColIndex
    nPos = Me.ssdbMembresia.Row
    
    RefreshssMem
End Sub
