VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmVendedores 
   Caption         =   "Vendedores"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8295
   Icon            =   "frmVendedores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBorrar 
      Height          =   615
      Left            =   7560
      Picture         =   "frmVendedores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "  Borrar datos  "
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton cmdAgregar 
      Enabled         =   0   'False
      Height          =   615
      Left            =   7560
      Picture         =   "frmVendedores.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "  Guardar y agregar  "
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   7560
      Picture         =   "frmVendedores.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "  Salir  "
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Enabled         =   0   'False
      Height          =   615
      Left            =   7560
      Picture         =   "frmVendedores.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "  Cancelar  "
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton cmdModificar 
      Height          =   615
      Left            =   7560
      Picture         =   "frmVendedores.frx":12EA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "  Modificar datos  "
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdNuevo 
      Height          =   615
      Left            =   7560
      Picture         =   "frmVendedores.frx":172C
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "  Agregar registro  "
      Top             =   240
      Width           =   615
   End
   Begin MSComCtl2.DTPicker dtpFechaAlta 
      Height          =   285
      Left            =   5880
      TabIndex        =   3
      Top             =   4200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   58720257
      CurrentDate     =   38668
   End
   Begin VB.TextBox txtNombre 
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox txtCve 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   615
   End
   Begin SSDataWidgets_B.SSDBGrid ssdbVendedores 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _Version        =   196616
      DataMode        =   2
      Cols            =   3
      Col.Count       =   3
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowColumnSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      BackColorOdd    =   14737632
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   12726
      _ExtentY        =   6588
      _StockProps     =   79
      Caption         =   "Vendedores"
   End
   Begin VB.Label lblFecAlta 
      Caption         =   "Fecha de alta"
      Height          =   255
      Left            =   5880
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblCve 
      Caption         =   "Cve."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   615
   End
End
Attribute VB_Name = "frmVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************
'*  Formulario para vendedores      *
'*  Daniel Hdez                     *
'*  12 / Noviembre / 2005           *
'************************************

'*  Ult actualización:  14 / Noviembre / 2005


'Numero de columnas mostradas
Const TOTCOLUMNS = 3

'Encabezados para la consulta
Dim mTitVen(TOTCOLUMNS) As String

'Ancho de las columnas mostradas
Dim mAncVen(TOTCOLUMNS) As Integer

'Nombre de los campos en la consulta
Dim mCamVen(TOTCOLUMNS) As String

'Matriz de los encabezados del formulario
Dim mTitForma(TOTCOLUMNS) As String

'Columna activa o seleccionada para establecer el orden
Dim nColActiva As Integer

'Posicion del cursor dentro del dbGrid
Dim nPos As Variant

Dim sTextMain As String

Dim bSave As Boolean
Dim bNvoDato As Boolean

Dim sNombre As String
Dim nId As Integer
Dim nActivo  As Byte
Dim dFecAlta As Date



Private Sub cmdCancelar_Click()
    Me.cmdAgregar.Enabled = False
    Me.cmdCancelar.Enabled = False
    Me.cmdNuevo.Enabled = True
    Me.cmdModificar.Enabled = True
    Me.cmdBorrar.Enabled = True
    Me.cmdSalir.Enabled = True
    
    Me.ssdbVendedores.Enabled = True

    ActivaCtrlsTxt False
    
    ClrCtrlsTxt
    
    InitVars
    
    bNvoDato = False

    If (Me.ssdbVendedores.Enabled) Then
        If (Me.ssdbVendedores.Rows > 0) Then
            Me.ssdbVendedores.Bookmark = Me.ssdbVendedores.AddItemBookmark(Me.ssdbVendedores.AddItemRowIndex(nPos))
        End If
        
        Me.ssdbVendedores.SetFocus
    End If
End Sub


Private Sub cmdBorrar_Click()
Dim nOk As Long

    If (Me.ssdbVendedores.Rows > 0) Then
        nOk = MsgBox("¿Desea borrar al vendedor seleccionado?", vbYesNo, "KalaSystems")

        If (nOk = vbYes) Then
            If (EliminaReg("Vendedores", "idVendedor=" & Me.ssdbVendedores.Columns(0).Value, "", Conn)) Then
                RefreshSSGrid
            Else
                MsgBox "No se borraron los datos del usuario.", vbInformation, DEVELOPER
            End If
        End If

        If (Me.ssdbVendedores.Enabled) Then
            Me.ssdbVendedores.SetFocus
        End If
    End If
End Sub


'Modificar datos
Private Sub cmdModificar_Click()
    If (Me.ssdbVendedores.Rows <= 0) Then
        Exit Sub
    End If
    
    bNvoDato = False
    nPos = Me.ssdbVendedores.Bookmark

    Me.cmdNuevo.Enabled = False
    Me.cmdModificar.Enabled = False
    Me.cmdBorrar.Enabled = False
    Me.cmdCancelar.Enabled = True
    Me.cmdAgregar.Enabled = True
    Me.cmdSalir.Enabled = False
    
    Me.ssdbVendedores.Enabled = False
    
    ActivaCtrlsTxt True
    
    ClrCtrlsTxt
    
    LeeDatos
    
    InitVars
    
    Me.txtNombre.SetFocus
End Sub


Private Sub ActivaCtrlsTxt(bValor As Boolean)
    With Me
        .txtNombre.Enabled = bValor
        .dtpFechaAlta.Enabled = bValor
    End With
End Sub


Private Sub ClrCtrlsTxt()
    With Me
        .txtNombre.Text = ""
        .txtCve.Text = ""
        .dtpFechaAlta.Value = Format(Date, "dd/mm/yyyy")
    End With
End Sub


Private Sub LeeDatos()
    With Me
        .txtCve.Text = Val(.ssdbVendedores.Columns(0).Text)
        .txtNombre.Text = Trim$(.ssdbVendedores.Columns(1).Text)
        .dtpFechaAlta.Value = Format(.ssdbVendedores.Columns(2).Value, "dd/mm/yyyy")
    End With
End Sub


Private Sub InitVars()
    sNombre = Trim$(Me.txtNombre.Text)
    nId = Val(Me.txtCve.Text)
    dFecAlta = Me.dtpFechaAlta.Value
End Sub


Private Sub cmdAgregar_Click()
    If (Cambios) Then
        If (Not bSave) Then
            If (Not GuardaDatos) Then
                Exit Sub
            End If
        End If
        
        ClrCtrlsTxt
        
        InitVars
        
        RefreshSSGrid

        'Inicializa los valores de la forma
        bSave = False
        bNvoDato = True

        Me.txtNombre.SetFocus
'    Else
'        Me.cmdAgregar.Enabled = False
'        Me.cmdCancelar.Enabled = False
'        Me.cmdNuevo.Enabled = True
'        Me.cmdModificar.Enabled = True
'        Me.cmdBorrar.Enabled = True
'        Me.cmdSalir.Enabled = True
'
'        Me.ssdbVendedores.Enabled = True
'
'        ClrCtrlsTxt
'
'        InitVars
'
'        ActivaCtrlsTxt False
'
'        If (Me.ssdbVendedores.Enabled) Then
'            If (Me.ssdbVendedores.Rows > 0) Then
'                Me.ssdbVendedores.Bookmark = Me.ssdbVendedores.AddItemBookmark(Me.ssdbVendedores.AddItemRowIndex(nPos))
'            End If
'
'            Me.ssdbVendedores.SetFocus
'        End If
    End If
End Sub


Private Sub cmdNuevo_Click()
    bNvoDato = True
    
    Me.cmdNuevo.Enabled = False
    Me.cmdModificar.Enabled = False
    Me.cmdBorrar.Enabled = False
    Me.cmdSalir.Enabled = False
    Me.cmdAgregar.Enabled = True
    Me.cmdCancelar.Enabled = True
    
    Me.ssdbVendedores.Enabled = False
    
    ActivaCtrlsTxt True
    
    Me.txtNombre.Enabled = True
    Me.txtNombre.SetFocus
End Sub


Private Sub cmdSalir_Click()
    'Cierra el formulario
    Unload Me
End Sub


Private Sub Form_Activate()
    'Encabezados del formulario
    mTitForma(0) = " clave"
    mTitForma(1) = " nombre"
    mTitForma(2) = " Fec. alta"
    
    mCamVen(0) = "idVendedor"
    mCamVen(1) = "Nombre"
    mCamVen(2) = "FechaAlta"
    
    InitssGridVend
    
    RefreshSSGrid
End Sub


Private Sub Form_Load()
    nPos = 0
    bSave = False
    bNvoDato = True
    
    sTextMain = MDIPrincipal.StatusBar1.Panels.Item(1).Text

    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Vendedores"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmVendedores = Nothing
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextMain
End Sub


Private Sub InitssGridVend()
    'Asigna valores a la matriz de encabezados
    mTitVen(0) = "Cve."
    mTitVen(1) = "Nombre"
    mTitVen(2) = "Fec. alta"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid Me.ssdbVendedores, mTitVen

    'Asigna valores a la matriz que define el ancho de cada columna
    mAncVen(0) = 700
    mAncVen(1) = 4170
    mAncVen(2) = 2000

    'Asigna el ancho de cada columna
    DefAnchossGrid Me.ssdbVendedores, mAncVen
    
    'Alinea a la derecha las columnas que contienen números
    Me.ssdbVendedores.Columns(0).Alignment = ssCaptionAlignmentRight
    Me.ssdbVendedores.Columns(1).Alignment = ssCaptionAlignmentLeft
    Me.ssdbVendedores.Columns(2).Alignment = ssCaptionAlignmentCenter
End Sub


Private Sub RefreshSSGrid()
Dim rsVendedor As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String

    sCampos = "idVendedor, Nombre, FechaAlta"
    
    sTablas = "Vendedores"

    InitRecordSet rsVendedor, sCampos, sTablas, "", mCamVen(1), Conn
    
    Me.ssdbVendedores.RemoveAll
    
    With rsVendedor
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                Me.ssdbVendedores.AddItem .Fields("idVendedor") & vbTab & _
                .Fields("Nombre") & vbTab & _
                Format(.Fields("FechaAlta"), "dd / mmm / yyyy")
                
                .MoveNext
            Loop
        End If
    End With
    
    rsVendedor.Close
    Set rsVendedor = Nothing

    If (Me.ssdbVendedores.Rows <= 0) Then
        Me.ssdbVendedores.Enabled = False
    Else
        Me.ssdbVendedores.Bookmark = Me.ssdbVendedores.AddItemBookmark(Me.ssdbVendedores.AddItemRowIndex(nPos))
    End If
    
    If (Me.ssdbVendedores.Enabled) Then
        Me.ssdbVendedores.SetFocus
    End If
End Sub


Private Function ChecaDatos()
Dim sCond As String
Dim sCamp As String

    ChecaDatos = False

    If (Trim$(Me.txtNombre.Text) = "") Then
        MsgBox "Se debe escribir el nombre del empleado.", vbExclamation, DEVELOPER
        Me.txtNombre.SetFocus
        Exit Function
    End If
    
    ChecaDatos = True
End Function


Private Function GuardaDatos() As Boolean
Const DATOSVEND = 3
Dim bCreado As Boolean
Dim mFieldsVen(DATOSVEND) As String
Dim mValuesVen(DATOSVEND) As Variant

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla
    mFieldsVen(0) = "idVendedor"
    mFieldsVen(1) = "Nombre"
    mFieldsVen(2) = "FechaAlta"

    If (bNvoDato) Then
        mValuesVen(0) = LeeUltReg("Vendedores", "idVendedor") + 1
    Else
        mValuesVen(0) = Val(Me.txtCve.Text)
    End If

    mValuesVen(1) = Trim$(UCase$(Me.txtNombre.Text))
    #If SqlServer_ Then
        mValuesVen(2) = Format(Me.dtpFechaAlta.Value, "yyyymmdd")
    #Else
        mValuesVen(2) = Format(Me.dtpFechaAlta.Value, "dd/mm/yyyy")
    #End If

    If (bNvoDato) Then
        'Registra los datos de la nueva direccion
        If (AgregaRegistro("Vendedores", mFieldsVen, DATOSVEND, mValuesVen, Conn)) Then
            GuardaDatos = True
        Else
            MsgBox "El registro no fue completado.", vbCritical, DEVELOPER
        End If
    Else

        If (Val(Me.txtCve.Text) > 0) Then

            'Actualiza los datos
            If (CambiaReg("Vendedores", mFieldsVen, DATOSVEND, mValuesVen, "idVendedor=" & Val(Me.txtCve.Text), Conn)) Then
                GuardaDatos = True
            Else
                MsgBox "No se realizaron los cambios.", vbCritical, DEVELOPER
            End If
        End If

    End If
End Function


'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    Cambios = True

    If (sNombre <> Trim$(Me.txtNombre.Text)) Then
        Exit Function
    End If
    
    If (dFecAlta <> Format(Me.dtpFechaAlta.Value, "dd/mm/yyyy")) Then
        Exit Function
    End If

    Cambios = False
End Function
