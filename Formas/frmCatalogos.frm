VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCatalogos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogos"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10860
   Icon            =   "frmCatalogos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   10860
   Begin VB.Frame frmBusca 
      ForeColor       =   &H000000FF&
      Height          =   1560
      Left            =   2205
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   6075
      Begin VB.CommandButton cmdAceptaBusca 
         Caption         =   "&Buscar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   870
         Width           =   1200
      End
      Begin VB.CommandButton cmdCancelaBusca 
         Caption         =   "C&ancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3435
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   885
         Width           =   1200
      End
      Begin VB.TextBox txtBusca 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         TabIndex        =   13
         Top             =   480
         Width           =   4350
      End
      Begin VB.OptionButton optBusca 
         Caption         =   "&Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   975
         TabIndex        =   11
         Top             =   210
         Value           =   -1  'True
         Width           =   1845
      End
      Begin VB.OptionButton optBusca 
         Caption         =   "&Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3690
         TabIndex        =   12
         Top             =   210
         Width           =   1710
      End
   End
   Begin VB.Frame frmConfirmaAdmin 
      BackColor       =   &H00800000&
      Height          =   1500
      Left            =   3210
      TabIndex        =   20
      Top             =   1605
      Visible         =   0   'False
      Width           =   3465
      Begin VB.CommandButton cmdCancelaConfirmaAdmin 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   10
         Top             =   960
         Width           =   1110
      End
      Begin VB.CommandButton cmdAceptaConfirmaAdmin 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   615
         TabIndex        =   9
         Top             =   960
         Width           =   1110
      End
      Begin VB.TextBox txtConfirmaAdmin 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   735
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   555
         Width           =   2085
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Usted es el Administrador, Ingrese su Contraseña de Acceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   300
         TabIndex        =   21
         Top             =   135
         Width           =   2925
      End
   End
   Begin VB.CommandButton cmdPassword 
      Enabled         =   0   'False
      Height          =   600
      Left            =   5745
      Picture         =   "frmCatalogos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Cambio de Contraseña "
      Top             =   0
      Width           =   945
   End
   Begin VB.CommandButton cmdMuchos 
      Enabled         =   0   'False
      Height          =   600
      Left            =   1005
      Picture         =   "frmCatalogos.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   " Agregar Múltiples "
      Top             =   0
      Width           =   945
   End
   Begin VB.CommandButton cmdModificar 
      Height          =   600
      Left            =   3855
      Picture         =   "frmCatalogos.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " Modificar "
      Top             =   0
      Width           =   945
   End
   Begin VB.CommandButton cmdRefresca 
      Height          =   600
      Left            =   4800
      Picture         =   "frmCatalogos.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Refrescar "
      Top             =   0
      Width           =   945
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   600
      Left            =   7260
      Picture         =   "frmCatalogos.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Salir "
      Top             =   0
      Width           =   945
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   600
      Left            =   2895
      Picture         =   "frmCatalogos.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Buscar "
      Top             =   0
      Width           =   945
   End
   Begin VB.CommandButton cmdEliminar 
      Height          =   600
      Left            =   1950
      Picture         =   "frmCatalogos.frx":1546
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Eliminar "
      Top             =   0
      Width           =   945
   End
   Begin VB.CommandButton cmdAgregar 
      Height          =   600
      Left            =   45
      Picture         =   "frmCatalogos.frx":1850
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " Agregar "
      Top             =   0
      Width           =   945
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid grdCatalogos 
      Height          =   5295
      Left            =   240
      TabIndex        =   23
      Top             =   840
      Width           =   10455
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   979
      Columns(0).Caption=   "IdUnidad"
      Columns(0).Name =   "IdUnidad"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2011
      Columns(1).Caption=   "NoInscripcion"
      Columns(1).Name =   "NoInscripcion"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4577
      Columns(2).Caption=   "Nombre"
      Columns(2).Name =   "Nombre"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2117
      Columns(3).Caption=   "FechaUPago"
      Columns(3).Name =   "FechaUPago"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1773
      Columns(4).Caption=   "Secuencial"
      Columns(4).Name =   "Secuencial"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "Fotofile"
      Columns(5).Name =   "Fotofile"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   18441
      _ExtentY        =   9340
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
   Begin VB.Label lblModoDescrip 
      Height          =   300
      Left            =   12000
      TabIndex        =   19
      Top             =   4080
      Width           =   285
   End
   Begin VB.Label lblModo 
      Height          =   300
      Left            =   12000
      TabIndex        =   18
      Top             =   4080
      Width           =   195
   End
   Begin VB.Label lblTotReg 
      BackStyle       =   0  'Transparent
      Caption         =   "Registros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8955
      TabIndex        =   17
      Top             =   195
      Width           =   810
   End
   Begin VB.Label LblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
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
      Left            =   8955
      TabIndex        =   16
      Top             =   420
      Width           =   810
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCatalogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA CATÁLOGOS
' Objetivo: CONTROLA LOS CATÁLOGOS DEL SISTEMA
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim intIntentos As Integer
    Dim blnEncontrado As Boolean
    Dim adoRcsCatal As ADODB.Recordset
    Dim strTabla As String, strCampoClave As String, strCampoDescrip As String
    Dim strDato1 As String, strDato2 As String, strConsulta As String

Private Sub cmdAceptaBusca_Click()
    Dim Posicion As Variant
    Dim strCadena As String
    
    If Me.Tag = "TITULOS" Then
        If optBusca(0).Value = True Then
            strCadena = grdCatalogos.Columns.Item(1).DataField & " = " & Val(Trim(txtBusca.Text))
        Else
            strCadena = grdCatalogos.Columns.Item(0).DataField & " like %" & Trim(txtBusca.Text) & "%"
        End If
    ElseIf (Me.Tag = "HORARIOS DE CLASES") Or (Me.Tag = "RENTABLES") Then
        If optBusca(0).Value = True Then
            strCadena = grdCatalogos.Columns.Item(0).DataField & " like %" & Trim(txtBusca.Text) & "%"
        Else
            strCadena = grdCatalogos.Columns.Item(1).DataField & " like %" & Trim(txtBusca.Text) & "%"
        End If
    Else
        If optBusca(0).Value = True Then
            strCadena = grdCatalogos.Columns.Item(1).DataField & " like %" & Trim(txtBusca.Text) & "%"
        Else
            strCadena = grdCatalogos.Columns.Item(0).DataField & " = " & Val(Trim(txtBusca.Text))
        End If
    End If
    Posicion = AdoDcCatal.Recordset.Bookmark
    AdoDcCatal.Recordset.MoveFirst
    AdoDcCatal.Recordset.Find (strCadena)
    If AdoDcCatal.Recordset.EOF Then
        AdoDcCatal.Recordset.Bookmark = Posicion
        blnEncontrado = False
        MsgBox "Elemento NO encontrado."
        txtBusca.SetFocus
        Exit Sub
    Else
        Posicion = AdoDcCatal.Recordset.Bookmark
        frmBusca.Visible = False
    End If
    AdoDcCatal.Recordset.Bookmark = Posicion
    Call BotonesOn
End Sub

Private Sub cmdAceptaConfirmaAdmin_Click()
    Dim strAcredita As String
    strAcredita = Trim(txtConfirmaAdmin.Text)
    If strAcredita <> sDB_PW Then
        MsgBox "¡ Contraseña Incorrecta !", vbCritical, "Error"
        txtConfirmaAdmin.SetFocus
        intIntentos = intIntentos + 1
        If intIntentos > 2 Then Unload Me
        Exit Sub
    End If
    txtConfirmaAdmin.Text = ""
    cmdAgregar.Enabled = True
    cmdEliminar.Enabled = True
    cmdBuscar.Enabled = True
    cmdModificar.Enabled = True
    cmdRefresca.Enabled = True
    cmdPassword.Enabled = True
    frmConfirmaAdmin.Visible = False
    frmCambioPassword.Show
End Sub

Private Sub cmdAgregar_Click()
    lblModo.Caption = "A"
    Call CargaForma
    Actualiza_Grid (MDIPrincipal.StatusBar1.Panels.Item(1).Text)
End Sub

Private Sub cmdBuscar_Click()
    Call BotonesOff
    frmBusca.Visible = True
    If Me.Tag = "REGLAS" Then
        optBusca(1).Value = True
        optBusca(0).Enabled = False
    End If
    If Me.Tag = "USUARIOS_SISTEMA" Then
        optBusca(0).Value = True
        optBusca(1).Enabled = False
    End If
        
    Select Case Me.Tag
        Case "HORARIOS DE CLASES"
            optBusca(0).Caption = "Instructor"
            optBusca(1).Caption = "Clase"
        Case "INSTRUCTORES"
            optBusca(0).Caption = "&Apellido Paterno"
        Case "RENTABLES"
            optBusca(0).Caption = "&Número"
            optBusca(1).Caption = "&Rentable"
        Case "TITULOS"
            optBusca(0).Caption = "&Número"
            optBusca(1).Caption = "&Tipo"
        Case "USUARIOS_SISTEMA"
            optBusca(0).Caption = "Login"
    End Select
    txtBusca.Text = ""
    txtBusca.SetFocus
End Sub

Private Sub cmdCancelaBusca_Click()
    Call BotonesOn
    frmBusca.Visible = False
End Sub

Private Sub cmdCancelaConfirmaAdmin_Click()
    cmdAgregar.Enabled = True
    cmdEliminar.Enabled = True
    cmdBuscar.Enabled = True
    cmdModificar.Enabled = True
    cmdRefresca.Enabled = True
    cmdPassword.Enabled = True
    frmConfirmaAdmin.Visible = False
    txtConfirmaAdmin.Text = ""
End Sub

Private Sub cmdEliminar_Click()
    Dim Respuesta As Integer
    Dim rs As New ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim str_empid, strconnect As String
    
    If Me.grdCatalogos.Rows = 0 Then
        Exit Sub
    End If
    
    str_empid = Me.grdCatalogos.Columns("IdUnidad").Value
    If Not IsNull(Me.grdCatalogos.Columns("IdUnidad").Value) Then
'        strDato1 = AdoDcCatal.Recordset.Fields(strCampoClave)
'        strDato2 = AdoDcCatal.Recordset.Fields(strCampoDescrip)
        Respuesta = MsgBox("¿ Desea Eliminar el Registro " & Me.grdCatalogos.Columns("Nombre").Value & _
                                        " ?", vbQuestion + vbYesNo, "Catálogos")
        If Respuesta = vbYes Then
            'Call EliminaRegistro
       

             Set cmd = New ADODB.Command
             cmd.ActiveConnection = Conn
             cmd.CommandType = adCmdStoredProc
             cmd.CommandText = "usp_Catalogos_Elimina"
            
             cmd.Parameters.Append cmd.CreateParameter("Id", adVarChar, adParamInput, 6, Me.grdCatalogos.Columns("IdUnidad").Value)
            
             Set rs = cmd.Execute
            
             If Not rs.EOF Then
              'txt_firstname = rs.Fields(0)
              'txt_title = rs.Fields(1)
              'txt_address = rs.Fields(2)
             End If
            
             Set cmd.ActiveConnection = Nothing
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    strDato1 = Me.grdCatalogos.Columns("IdUnidad").Value
    'strDato2 = AdoDcCatal.Recordset.Fields(strCampoDescrip)
    lblModo.Caption = strDato1
    lblModoDescrip.Caption = strDato2
    Call CargaForma
End Sub

Private Sub cmdMuchos_Click()
    lblModo.Caption = "AA"
    Call CargaForma
End Sub

Private Sub cmdPassword_Click()
    If (sDB_NivelUser = 0) Then
'        cmdAgregar.Enabled = False
'        cmdEliminar.Enabled = False
'        cmdBuscar.Enabled = False
'        cmdModificar.Enabled = False
'        cmdRefresca.Enabled = False
'        cmdPassword.Enabled = False
'        frmConfirmaAdmin.Visible = True
'        txtConfirmaAdmin.SetFocus
frmCambioPassword.Show
    Else
        frmCambioPassword.Show
    End If
End Sub

Private Sub cmdRefresca_Click()
    Actualiza_Grid (MDIPrincipal.StatusBar1.Panels.Item(1).Text)
End Sub

Private Sub cmdSalir_Click()
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "KALACLUB"
    Unload Me
End Sub

Private Sub Form_Activate()
    With frmCatalogos
        .Left = 0
        .Top = 0
        '.Height = 6270
        '.Width = 10000
    End With
   ' Actualiza_Grid (MDIPrincipal.StatusBar1.Panels.Item(1).Text)
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    'grdCatalogos.ClearFields
    intIntentos = 0
    For i = 1 To 14
        MDIPrincipal.mnuCatal(i).Enabled = False        'Deshabilita menú
    Next i
 Actualiza_Grid (MDIPrincipal.StatusBar1.Panels.Item(1).Text)
End Sub

Private Sub Actualiza_Grid(Opicion As String)
       Select Case Opicion
        Case "CONCEPTOS DE INGRESOS"
            Me.Tag = "INGRESOS"
            Me.Caption = "CATÁLOGOS - CONCEPTOS DE INGRESOS"
            strTabla = "concepto_ingresos"
            Call LlenaGrid("concepto_ingresos", "idconcepto")
        Case "PAGOS"
            Me.Tag = "PAGOS"
            Me.Caption = "CATÁLOGOS - PAGOS"
            strTabla = "forma_pago"
            Call LlenaGrid("forma_pago", "idformapago")
        Case "BANCOS"
            Me.Tag = "BANCOS"
            Me.Caption = "CATÁLOGOS - BANCOS"
            strTabla = "bancos"
            Call LlenaGrid("bancos", "idbanco")
        Case "TIPO RENTABLES"
            Me.Tag = "TIPO RENTABLES"
            Me.Caption = "CATÁLOGOS - TIPO RENTABLES"
            strTabla = "tipo_rentables"
            Call LlenaGrid("tipo_rentables", "idtiporentable")
        Case "TIPO USUARIO"
            Me.Tag = "USUARIO"
            Me.Caption = "CATÁLOGOS - TIPOS DE USUARIOS"
            strTabla = "tipo_usuario"
            Call LlenaGrid("tipo_usuario", "idtipousuario")
        Case "REGLAS TIPO USUARIO"
            Me.Tag = "REGLAS"
            Me.Caption = "CATÁLOGOS - REGLAS TIPOS DE USUARIOS"
            strTabla = "reglas_tipo"
            Call LlenaGrid("reglas_tipo", "idtipoactual")
        Case "RENTABLES"
            Me.Tag = "RENTABLES"
            Me.Caption = "CATÁLOGOS - RENTABLES"
            strTabla = "rentables"
            cmdMuchos.Enabled = True
            Call LlenaGrid("rentables", "numero")
        Case "TITULOS"
            Me.Tag = "TITULOS"
            Me.Caption = "CATÁLOGOS - TÍTULOS"
            strTabla = "titulos"
            cmdAgregar.Enabled = False
            cmdMuchos.Enabled = True
            Call LlenaGrid("titulos", "numero")
        Case "MEMBRESÍAS"
            Me.Tag = "MEMBRESÍAS"
            Me.Caption = "CATÁLOGOS - MEMBRESÍAS"
            strTabla = "membresias"
            Call LlenaGrid("membresias", "descripcion")
        Case "USUARIOS DEL SISTEMA"
            Me.Tag = "USUARIOS_SISTEMA"
            Me.Caption = "CATÁLOGOS - USUARIOS DEL SISTEMA"
            strTabla = "usuarios_sistema"
            cmdPassword.Enabled = True
            Call LlenaGrid("usuarios_sistema", "idusuario")
        Case "INSTRUCTORES"
            Me.Tag = "INSTRUCTORES"
            Me.Caption = "CATÁLOGOS - INSTRUCTORES"
            strTabla = "instructores"
            Call LlenaGrid("instructores", "apellido_paterno")
        Case "HORARIOS DE CLASES"
            Me.Tag = "HORARIOS DE CLASES"
            Me.Caption = "CATÁLOGOS - HORARIOS DE CLASES"
            strTabla = "horarios_clases"
            Call LlenaGrid("horarios_clases", "idtipoclase")
        Case "TIPOS DE CLASES"
            Me.Tag = "TIPOS DE CLASES"
            Me.Caption = "CATÁLOGOS - TIPOS DE CLASES"
            strTabla = "tipo_clase"
            Call LlenaGrid("tipo_clase", "idtipoclase")
        Case "PAISES"
            Me.Tag = "PAISES"
            Me.Caption = "CATÁLOGOS - PAÍSES"
            strTabla = "paises"
            Call LlenaGrid("paises", "pais")
    End Select
    Call ActualizaGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    For i = 1 To 14
        MDIPrincipal.mnuCatal(i).Enabled = True         'Rehabilia Menú
    Next i
    Me.Caption = "CATÁLOGOS"
End Sub

Private Sub optBusca_Click(Index As Integer)
    txtBusca.SetFocus
End Sub

Private Sub txtBusca_GotFocus()
    txtBusca.SelStart = 0
    txtBusca.SelLength = Len(txtBusca)
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 32 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Else
        KeyAscii = 0
    End If
End Sub

Sub ActualizaGrid()
    'AdoDcCatal.REFRESH
    grdCatalogos.REFRESH
'    If AdoDcCatal.Recordset.RecordCount = 0 Then
'        cmdEliminar.Enabled = False
'        cmdBuscar.Enabled = False
'        cmdModificar.Enabled = False
'        LblTotal.Caption = "0"
'    Else
'        cmdEliminar.Enabled = True
'        cmdBuscar.Enabled = True
'        cmdModificar.Enabled = True
'    End If
    Select Case Me.Tag
        Case "INGRESOS"
            strCampoClave = "idconcepto"
            strCampoDescrip = "descripcion"
            grdCatalogos.Caption = "CONCEPTOS DE INGRESOS"
            grdCatalogos.Columns.Item(0).Width = 950
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(1).Width = 3300
            grdCatalogos.Columns.Item(1).Caption = "Descripción"
            grdCatalogos.Columns.Item(2).Width = 1500
            grdCatalogos.Columns.Item(2).Alignment = dbgCenter
            grdCatalogos.Columns.Item(2).Caption = "Cuenta Contable"
            grdCatalogos.Columns.Item(3).Width = 700
            grdCatalogos.Columns.Item(3).Alignment = dbgCenter
            grdCatalogos.Columns.Item(3).Caption = "Monto"
            grdCatalogos.Columns.Item(4).Width = 1000
            grdCatalogos.Columns.Item(4).Alignment = dbgCenter
            grdCatalogos.Columns.Item(4).Caption = "Impuesto 1"
            grdCatalogos.Columns.Item(5).Width = 1000
            grdCatalogos.Columns.Item(5).Alignment = dbgCenter
            grdCatalogos.Columns.Item(5).Caption = "Impuesto 2"
            grdCatalogos.Columns.Item(6).Width = 1100
            grdCatalogos.Columns.Item(6).Caption = "¿Periódico?"
            grdCatalogos.Columns.Item(6).Alignment = dbgCenter
            grdCatalogos.Columns.Item(7).Width = 1100
            grdCatalogos.Columns.Item(7).Caption = "Tipo Doc."
            grdCatalogos.Columns.Item(7).Alignment = dbgCenter
        Case "PAGOS"
            strCampoClave = "idformapago"
            strCampoDescrip = "descripcion"
            grdCatalogos.Caption = "FORMA DE PAGO"
            grdCatalogos.Columns.Item(0).Width = 1200
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(1).Width = 3500
            grdCatalogos.Columns.Item(1).Caption = "Descripción"
        Case "BANCOS"
            strCampoClave = "idbanco"
            strCampoDescrip = "banco"
            grdCatalogos.Caption = "BANCOS"
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(0).Width = 800
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(1).Caption = "Descripción"
            grdCatalogos.Columns.Item(1).Width = 4500
        Case "TIPO RENTABLES"
            strCampoClave = "idtiporentable"
            strCampoDescrip = "descripcion"
            grdCatalogos.Caption = "TIPO RENTABLES"
            grdCatalogos.Columns.Item(0).Width = 800
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(1).Width = 4500
            grdCatalogos.Columns.Item(1).Caption = "Descripción"
        Case "USUARIO"
            strCampoClave = "idtipousuario"
            strCampoDescrip = "descripcion"
            grdCatalogos.Caption = "TIPO DE USUARIOS"
            grdCatalogos.Columns.Item(0).Width = 800
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(1).Width = 2500
            grdCatalogos.Columns.Item(1).Caption = "Descripción"
            grdCatalogos.Columns.Item(2).Width = 1200
            grdCatalogos.Columns.Item(2).Alignment = dbgCenter
            grdCatalogos.Columns.Item(2).Caption = "Edad Mínima"
            grdCatalogos.Columns.Item(3).Width = 1250
            grdCatalogos.Columns.Item(3).Alignment = dbgCenter
            grdCatalogos.Columns.Item(3).Caption = "Edad Máxima"
            grdCatalogos.Columns.Item(4).Width = 1100
            grdCatalogos.Columns.Item(4).Alignment = dbgCenter
            grdCatalogos.Columns.Item(4).Caption = "Parentesco"
            grdCatalogos.Columns.Item(5).Width = 1200
            grdCatalogos.Columns.Item(5).Caption = "Tipo"
            grdCatalogos.Columns.Item(6).Width = 1200
            grdCatalogos.Columns.Item(6).Alignment = dbgCenter
            grdCatalogos.Columns.Item(6).Caption = "¿Es familiar?"
        Case "REGLAS"
            strCampoClave = "idtipoactual"
            strCampoDescrip = "accion"
            grdCatalogos.Caption = "REGLAS TIPO USUARIO"
            grdCatalogos.Columns.Item(0).Width = 600
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(1).Width = 1100
            grdCatalogos.Columns.Item(1).Alignment = dbgCenter
            grdCatalogos.Columns.Item(1).Caption = "Tipo Actual"
            grdCatalogos.Columns.Item(2).Width = 1100
            grdCatalogos.Columns.Item(2).Alignment = dbgCenter
            grdCatalogos.Columns.Item(2).Caption = "Tipo Nuevo"
            grdCatalogos.Columns.Item(3).Width = 1300
            grdCatalogos.Columns.Item(3).Caption = "Sexo"
            grdCatalogos.Columns.Item(3).Alignment = dbgCenter
            grdCatalogos.Columns.Item(4).Width = 1300
            grdCatalogos.Columns.Item(4).Caption = "Acción"
            grdCatalogos.Columns.Item(4).Alignment = dbgCenter
            cmdBuscar.Enabled = False
        Case "RENTABLES"
            strCampoClave = "IdTipoRentable"
            strCampoDescrip = "Numero"
            grdCatalogos.Caption = "RENTABLES"
            grdCatalogos.Columns.Item(0).Width = 700
            grdCatalogos.Columns.Item(0).Caption = "Número"
            grdCatalogos.Columns.Item(1).Width = 1900
            grdCatalogos.Columns.Item(1).Alignment = dbgLeft
            grdCatalogos.Columns.Item(1).Caption = "Tipo"
            grdCatalogos.Columns.Item(2).Width = 600
            grdCatalogos.Columns.Item(2).Alignment = dbgCenter
            grdCatalogos.Columns.Item(2).Caption = "Sexo"
            grdCatalogos.Columns.Item(3).Width = 1200
            grdCatalogos.Columns.Item(3).Caption = "Ubicación"
            grdCatalogos.Columns.Item(4).Width = 700
            grdCatalogos.Columns.Item(4).Caption = "Clave"
            grdCatalogos.Columns.Item(5).Width = 3000
            grdCatalogos.Columns.Item(5).Caption = "Nombre"
            grdCatalogos.Columns.Item(6).Caption = "Fecha Pago"
            grdCatalogos.Columns.Item(6).Width = 1100
            grdCatalogos.Columns.Item(7).Alignment = dbgCenter
            grdCatalogos.Columns.Item(7).Caption = "UsoDiario"
            grdCatalogos.Columns.Item(7).Width = 900
            
        Case "TITULOS"
            strCampoClave = "numero"
            strCampoDescrip = "serie"
            grdCatalogos.Caption = "TITULOS"
            cmdModificar.Enabled = False
            grdCatalogos.Columns.Item(0).Width = 600
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(0).Caption = "Tipo"
            grdCatalogos.Columns.Item(1).Width = 800
            grdCatalogos.Columns.Item(1).Alignment = dbgCenter
            grdCatalogos.Columns.Item(1).Caption = "Número"
            grdCatalogos.Columns.Item(2).Width = 600
            grdCatalogos.Columns.Item(2).Alignment = dbgCenter
            grdCatalogos.Columns.Item(2).Caption = "Serie"
            grdCatalogos.Columns.Item(3).Width = 4000
            grdCatalogos.Columns.Item(3).Caption = "Propietario"
            grdCatalogos.Columns.Item(4).Width = 1100
            grdCatalogos.Columns.Item(4).Alignment = dbgCenter
            grdCatalogos.Columns.Item(4).Caption = "Creación"
            grdCatalogos.Columns.Item(5).Width = 1100
            grdCatalogos.Columns.Item(5).Alignment = dbgCenter
            grdCatalogos.Columns.Item(5).Caption = "Asignación"
        Case "MEMBRESÍAS"
            strCampoClave = "idmembresia"
            strCampoDescrip = "descripcion"
            grdCatalogos.Caption = "MEMBRESÍAS"
            grdCatalogos.Columns.Item(0).Width = 800
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(1).Width = 2800
            grdCatalogos.Columns.Item(1).Caption = "Descripción"
            grdCatalogos.Columns.Item(2).Width = 950
            grdCatalogos.Columns.Item(2).Alignment = dbgCenter
            grdCatalogos.Columns.Item(2).Caption = "Contado"
            grdCatalogos.Columns.Item(3).Width = 1100
            grdCatalogos.Columns.Item(3).Alignment = dbgCenter
            grdCatalogos.Columns.Item(3).Caption = "Inscripción"
            grdCatalogos.Columns.Item(4).Width = 1050
            grdCatalogos.Columns.Item(4).Caption = "Alta"
            grdCatalogos.Columns.Item(4).Alignment = dbgCenter
            grdCatalogos.Columns.Item(5).Width = 1050
            grdCatalogos.Columns.Item(5).Caption = "VIgencia"
            grdCatalogos.Columns.Item(5).Alignment = dbgCenter
            grdCatalogos.Columns.Item(6).Width = 1450
            grdCatalogos.Columns.Item(6).Alignment = dbgCenter
            grdCatalogos.Columns.Item(6).Caption = "Duración (años)"
        Case "USUARIOS_SISTEMA"
            strCampoClave = "idusuario"
            strCampoDescrip = "login_name"
            grdCatalogos.Caption = "USUARIOS DEL SISTEMA"
            grdCatalogos.Columns.Item(0).Width = 800
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(1).Width = 1500
            grdCatalogos.Columns.Item(1).Caption = "Login"
            grdCatalogos.Columns.Item(2).Width = 4000
            grdCatalogos.Columns.Item(2).Caption = "Nombre"
            grdCatalogos.Columns.Item(3).Width = 1200
            grdCatalogos.Columns.Item(3).Alignment = dbgCenter
            grdCatalogos.Columns.Item(3).Caption = "Alta"
            grdCatalogos.Columns.Item(4).Width = 1200
            grdCatalogos.Columns.Item(4).Alignment = dbgCenter
            grdCatalogos.Columns.Item(4).Caption = "Vigencia"
            grdCatalogos.Columns.Item(5).Width = 1200
            grdCatalogos.Columns.Item(5).Alignment = dbgCenter
            grdCatalogos.Columns.Item(5).Caption = "IdPerfil"
        Case "INSTRUCTORES"
            strCampoClave = "idinstructor"
            strCampoDescrip = "apellido_paterno"
            grdCatalogos.Caption = "INSTRUCTORES"
            grdCatalogos.Columns.Item(0).Width = 700
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(1).Width = 2220
            grdCatalogos.Columns.Item(2).Width = 2220
            grdCatalogos.Columns.Item(3).Width = 2010
            grdCatalogos.Columns.Item(3).Caption = "Nombre(s)"
            grdCatalogos.Columns.Item(4).Width = 1000
            grdCatalogos.Columns.Item(5).Width = 4000
            grdCatalogos.Columns.Item(6).Width = 4000
            grdCatalogos.Columns.Item(7).Width = 3000
            grdCatalogos.Columns.Item(7).Caption = "Entidad Federativa"
            grdCatalogos.Columns.Item(8).Width = 2415
            grdCatalogos.Columns.Item(8).Caption = "Delegación o Municipio"
            grdCatalogos.Columns.Item(9).Width = 1500
            grdCatalogos.Columns.Item(10).Width = 1500
            grdCatalogos.Columns.Item(11).Width = 1620
        Case "HORARIOS DE CLASES"
            strCampoClave = "idtipoclase"
            strCampoDescrip = "idinstructor"
            grdCatalogos.Caption = "HORARIOS DE CLASES"
            grdCatalogos.Columns.Item(0).Width = 2200
            grdCatalogos.Columns.Item(0).Caption = "Clase"
            grdCatalogos.Columns.Item(1).Width = 3700
            grdCatalogos.Columns.Item(1).Caption = "Instructor"
            grdCatalogos.Columns.Item(2).Width = 900
            grdCatalogos.Columns.Item(2).Alignment = dbgCenter
            grdCatalogos.Columns.Item(2).Caption = "Días"
            grdCatalogos.Columns.Item(3).Width = 1200
            grdCatalogos.Columns.Item(3).Caption = "Inicia"
            grdCatalogos.Columns.Item(3).Alignment = dbgCenter
            grdCatalogos.Columns.Item(4).Width = 1200
            grdCatalogos.Columns.Item(4).Caption = "Termina"
            grdCatalogos.Columns.Item(4).Alignment = dbgCenter
            grdCatalogos.Columns.Item(5).Width = 0
            grdCatalogos.Columns.Item(6).Width = 0
        Case "TIPOS DE CLASES"
            strCampoClave = "idtipoclase"
            strCampoDescrip = "descripcion"
            grdCatalogos.Caption = "TIPOS DE CLASES"
            grdCatalogos.Columns.Item(0).Width = 1095
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(1).Width = 3500
            grdCatalogos.Columns.Item(1).Caption = "Descripción"
        Case "PAISES"
            strCampoClave = "idpais"
            strCampoDescrip = "pais"
            grdCatalogos.Caption = "PAISES"
            grdCatalogos.Columns.Item(0).Width = 800
            grdCatalogos.Columns.Item(0).Alignment = dbgCenter
            grdCatalogos.Columns.Item(0).Caption = "Clave"
            grdCatalogos.Columns.Item(1).Width = 2500
            grdCatalogos.Columns.Item(1).Caption = "País"
    End Select
    'LblTotal.Caption = Format(AdoDcCatal.Recordset.RecordCount, "#######")
End Sub

Sub BotonesOff()
    cmdAgregar.Enabled = False
    cmdMuchos.Enabled = False
    cmdEliminar.Enabled = False
    cmdBuscar.Enabled = False
    cmdModificar.Enabled = False
    cmdRefresca.Enabled = False
    cmdPassword.Enabled = False
    cmdSalir.Enabled = False
End Sub

Sub BotonesOn()
    If Me.Tag <> "TITULOS" Then cmdAgregar.Enabled = True
        
    If (Me.Tag = "RENTABLES") Or (Me.Tag = "TITULOS") Then
        cmdMuchos.Enabled = True
    Else
        cmdMuchos.Enabled = False
    End If
    cmdEliminar.Enabled = True
    If (AdoDcCatal.Recordset.RecordCount = 0) Or (Me.Tag = "REGLAS") Then
        cmdBuscar.Enabled = False
    Else
        cmdBuscar.Enabled = True
    End If
    If Me.Tag <> "TITULOS" Then cmdModificar.Enabled = True
    cmdRefresca.Enabled = True
    If Me.Tag = "USUARIOS_SISTEMA" Then cmdPassword.Enabled = True
    cmdSalir.Enabled = True
End Sub

Sub CargaForma()
    Select Case Me.Tag
        Case "INGRESOS"
            frmConceptoIngresos.Show
        Case "PAGOS"
            frmPagos.Show
        Case "BANCOS"
            frmBancos.Show
        Case "TIPO RENTABLES"
            frmTipoRentable.Show
        Case "USUARIO"
            frmTipoUsuario.Show
        Case "REGLAS"
            frmReglasTipo.Show
        Case "RENTABLES"
            If lblModo.Caption = "AA" Then
                frmMultiplesRentables.Show
            Else
                frmRentables.Show
            End If
        Case "TITULOS"
            frmMultiplesTitulos.Show
'        Case "MEMBRESÍAS"
'            frmMembresias.Show
        Case "USUARIOS_SISTEMA"
            frmUsuariosSistema.Show
        Case "INSTRUCTORES"
            frmInstructores.Show
'        Case "HORARIOS DE CLASES"
'            frmHorarios.Show
'        Case "TIPOS DE CLASES"
'            frmTipoClase.Show
        Case "PAISES"
            frmPaises.Show
    End Select
End Sub

Sub EliminaRegistro()
    Dim AdoCmdElimina As ADODB.Command
    On Error GoTo err_Elimina
    Screen.MousePointer = vbHourglass
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    strSQL = "DELETE FROM " & strTabla & " WHERE (" & strCampoClave & _
                    " = " & Val(strDato1) & ") "
    Select Case Me.Tag
        Case "RENTABLES"
            strSQL = strSQL & "AND (numero = '" & strDato2 & "')"
        Case "HORARIOS DE CLASES"
            strSQL = strSQL & "AND (idinstructor = " & _
                                            AdoDcCatal.Recordset.Fields("idinstructor") & ") AND (dias = '" & _
                                            AdoDcCatal.Recordset.Fields("dias") & "') AND (hora_inicio = format('" & _
                                            Format(AdoDcCatal.Recordset.Fields("hora_inicio"), "hh:mm") & _
                                            "', 'hh:mm')) AND (hora_fin = format('" & _
                                            Format(AdoDcCatal.Recordset.Fields("hora_fin"), "hh:mm") & "', 'hh:mm'))"
        Case "REGLAS"
            strSQL = strSQL & "AND (idtiponuevo = " & AdoDcCatal.Recordset.Fields("idtiponuevo") & _
                                            ") AND (sexo = '" & AdoDcCatal.Recordset.Fields("sexo") & "') AND (accion = '" & _
                                            AdoDcCatal.Recordset.Fields("accion") & "')"
    End Select
    Set AdoCmdElimina = New ADODB.Command
    AdoCmdElimina.ActiveConnection = Conn
    AdoCmdElimina.CommandText = strSQL
    AdoCmdElimina.Execute
    Conn.CommitTrans      'Termina transacción
    Call ActualizaGrid
    Screen.MousePointer = Default
   ' MsgBox "Registro Eliminado", vbOKOnly, "Catálogos"
    Exit Sub
    
err_Elimina:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Sub LlenaGrid(strTabla1, strOrden As String)
    Dim iColumnas As Integer
    
    Select Case Me.Tag
        Case "INGRESOS"
            #If SqlServer_ Then
                strSQL = "SELECT idconcepto, descripcion, cuentacontable, monto, impuesto1, impuesto2, " & _
                            "CASE WHEN ISNULL(esperiodico,0)=0 THEN 'No' ELSE 'Si' END AS esperiodico, FacORec FROM concepto_ingresos ORDER BY descripcion"
            #Else
                strSQL = "SELECT idconcepto, descripcion, cuentacontable, monto, impuesto1, impuesto2, " & _
                            "iif(esperiodico=false, 'No', 'Si'), FacORec FROM concepto_ingresos ORDER BY descripcion"
            #End If
        Case "HORARIOS DE CLASES"
            strSQL = "SELECT descripcion, (apellido_paterno + ' ' + apellido_materno + ' ' + nombre) as " & _
                            "instructor, dias, hora_inicio, hora_fin, horarios_clases.idtipoclase, " & _
                            "horarios_clases.idinstructor FROM (horarios_clases LEFT JOIN instructores ON " & _
                            "horarios_clases.idinstructor = instructores.idinstructor) LEFT JOIN tipo_clase ON " & _
                            "horarios_clases.idtipoclase = tipo_clase.idtipoclase ORDER BY descripcion"
        Case "RENTABLES"
            #If SqlServer_ Then
                strSQL = "SELECT  numero,descripcion, rentables.sexo, ubicacion, " & _
                            "usuarios_club.NoFamilia, (a_paterno + ' ' + a_materno + ' ' + nombre) as usuario, fechapago, " & _
                            "CASE WHEN ISNULL(propiedad,0)=0 THEN 'No' ELSE 'Si' END AS propiedad, rentables.idtiporentable FROM (rentables LEFT JOIN tipo_rentables ON " & _
                            "rentables.idtiporentable = tipo_rentables.idtiporentable) LEFT JOIN usuarios_club " & _
                            "ON rentables.idusuario = usuarios_club.idmember ORDER BY numero"
            #Else
                strSQL = "SELECT  numero,descripcion, rentables.sexo, ubicacion, " & _
                            "usuarios_club.NoFamilia, (a_paterno + ' ' + a_materno + ' ' + nombre) as usuario, fechapago, " & _
                            "IIF(propiedad=false, 'No', 'Si'), rentables.idtiporentable FROM (rentables LEFT JOIN tipo_rentables ON " & _
                            "rentables.idtiporentable = tipo_rentables.idtiporentable) LEFT JOIN usuarios_club " & _
                            "ON rentables.idusuario = usuarios_club.idmember ORDER BY numero"
            #End If
        Case "TITULOS"
            strSQL = "SELECT tipo, numero, serie, (a_paterno + ' ' + a_materno + ' ' + nombre) as " & _
                            "propietario, fecha_creacion, fecha_asignacion FROM titulos LEFT JOIN " & _
                            "accionistas ON titulos.idpropietario = accionistas.idproptitulo ORDER BY tipo, numero"
        Case "USUARIO"
            #If SqlServer_ Then
                strSQL = "SELECT idtipousuario, descripcion, edadminima, edadmaxima, parentesco, " & _
                            "tipo, CASE WHEN ISNULL(familiar,0)=0 THEN 'No' ELSE 'Si' END AS familiar FROM tipo_usuario ORDER BY descripcion"
            #Else
                strSQL = "SELECT idtipousuario, descripcion, edadminima, edadmaxima, parentesco, " & _
                            "tipo, iif(familiar=false, 'No', 'Si') FROM tipo_usuario ORDER BY descripcion"
            #End If
        Case "USUARIOS_SISTEMA"
            strSQL = "SELECT idusuario, login_name, nombre, fecha_alta,FechaVencePass,IdPerfil FROM usuarios_sistema " & _
                            "ORDER BY nombre"
            iColumnas = 6
        Case Else
            strSQL = "SELECT * FROM " & strTabla1 & " ORDER BY " & strOrden
    End Select
    Screen.MousePointer = vbHourglass
'    AdoDcCatal.ConnectionString = Conn
'    AdoDcCatal.CursorLocation = adUseClient
'    AdoDcCatal.LockType = adLockReadOnly
'    AdoDcCatal.CursorType = adOpenStatic
'    AdoDcCatal.RecordSource = strSQL
    
    LlenaSsDbGrid Me.grdCatalogos, Conn, strSQL, iColumnas
    
    
    Call ActualizaGrid
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtConfirmaAdmin_GotFocus()
    txtConfirmaAdmin.SelStart = 0
    txtConfirmaAdmin.SelLength = Len(txtConfirmaAdmin)
End Sub
