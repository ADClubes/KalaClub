VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmayuda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar registros"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNRegistros 
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
      Left            =   8160
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtBuscar 
      Height          =   300
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.CheckBox chkDesdeInicio 
      Caption         =   "&Buscar siempre desde el primer registro"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   435
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.CommandButton cmdElegir 
      Default         =   -1  'True
      Height          =   615
      Left            =   5400
      Picture         =   "frmayuda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "  Seleccionar  "
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdCancelar 
      Height          =   615
      Left            =   6120
      Picture         =   "frmayuda.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "  Cancelar  "
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdLocaliza 
      Height          =   615
      Left            =   3360
      Picture         =   "frmayuda.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "  Localizar  "
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdNuevo 
      Height          =   615
      Left            =   4680
      Picture         =   "frmayuda.frx":1296
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "  Nuevo dato  "
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid dgAyuda 
      Bindings        =   "frmayuda.frx":1B60
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4895
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
   Begin VB.Label lblNRegistros 
      Alignment       =   1  'Right Justify
      Caption         =   "No. de opciones"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmayuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************
'*  Formulario para mostrar la ayuda        *
'*  Daniel Hdez                             *
'*  04 / Septiembre / 2004                  *
'********************************************

'Ult actualización: 14 / Noviembre / 2005

'A estas variables se les asigna su valor en la forma que solicita la ayuda
Public nColsAyuda As Byte       'Total de columnas de la ayuda
Public sCondicion As String     'Condicion para filtrar la ayuda
Public nColActiva As Integer    'Columna activa
Public sTabla As String         'Nombre de la tabla en la ayuda
Public sTitAyuda As String      'Titulo del cuadro de errores o mensajes
Public lAgregar As Boolean      'Bandera para mostrar el boton de nuevos datos

'Estas variables toman su valores en esta forma
Dim mFormAyuda() As String      'Titulos en la forma
Dim mAnchAyuda() As Integer     'Ancho de las columnas
Dim mCampAyuda() As String      'Campos de la ayuda
Dim mEncaAyuda() As String      'Encabezados de las columnas
Dim sFormaPadre As String

Private Sub cmdLocaliza_Click()
    Dim rsDatos As ADODB.Recordset
    Set rsDatos = dgAyuda.DataSource
    
    If rsDatos.RecordCount > 0 And txtBuscar.Text <> "" Then
        
        If chkDesdeInicio.Value = 1 Then rsDatos.MoveFirst
        
        Me.txtBuscar.Text = UCase$(txtBuscar.Text)
        Me.txtBuscar.REFRESH
        
        rsDatos.Find rsDatos(nColActiva).Name & " LIKE '" & txtBuscar.Text & "*'"
        dgAyuda.REFRESH
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdElegir_Click()
    LeeRegistro
End Sub

Private Sub cmdNuevo_Click()
    Dim bUsaFrmGral As Boolean

    bUsaFrmGral = True

    Select Case sFormaPadre
        Case "frmAltaDir"
            frmAgregaOps.nOpcion = 1
            
        Case "frmPases"
            If (frmPases.nAyuda = 1) Then
                frmAgregaOps.nOpcion = 2
            ElseIf (frmPases.nAyuda = 2) Then
                frmAgregaOps.nOpcion = 6
            End If
            
        Case "frmAltaSocios"
            If (frmAltaSocios.nAyuda = 2) Then
                frmAgregaOps.nOpcion = 5
            ElseIf (frmAltaSocios.nAyuda = 3) Then
                frmAgregaOps.nOpcion = 3
            End If
            
        Case "frmAltaFam"
            If (frmAltaFam.nAyuda = 1) Then
                frmAgregaOps.nOpcion = 5
            ElseIf (frmAltaFam.nAyuda = 2) Then
                frmAgregaOps.nOpcion = 4
            End If
            
        Case "frmTZoneUsers"
            frmAgregaOps.nOpcion = 6
            
        Case "frmCertificados"
            bUsaFrmGral = False
            frmMedico.bNvoDr = True
            frmMedico.Show (1)
            
    End Select
    
    'Muestra el formulario de formato Gral.
    If (bUsaFrmGral) Then
        frmAgregaOps.Show (1)
    End If

    ActualizaVista nColActiva

    Me.dgAyuda.SetFocus
End Sub

Private Sub dgAyuda_DblClick()
    LeeRegistro
End Sub

Private Sub dgAyuda_KeyUp(KeyCode As Integer, Shift As Integer)
'    If (KeyCode = vbKeyReturn) Then
'        LeeRegistro
'    End If
End Sub

Private Sub dgAyuda_KeyPress(KeyAscii As Integer)
'    Dim nLen As Byte
'
'    If (Me.adoAyuda.Recordset.RecordCount > 0) Then
'        Select Case KeyAscii
'            Case 65 To 90
'                Me.txtBuscar.Text = Me.txtBuscar.Text + Chr(KeyAscii)
'
'            Case 97 To 122
'                Me.txtBuscar.Text = Me.txtBuscar.Text + UCase(Chr(KeyAscii))
'
'            Case 8
'                nLen = Len(Trim(Me.txtBuscar.Text))
'                If (nLen > 0) Then
'                    Me.txtBuscar.Text = Mid(Me.txtBuscar.Text, 1, nLen - 1)
'                End If
'        End Select
'
'        If (Trim(Me.txtBuscar.Text) <> "") Then
'            Me.adoAyuda.Recordset.Find mCampAyuda(nColActiva) & " LIKE '" & LTrim(Me.txtBuscar.Text) & "*'"
'            If (Me.adoAyuda.Recordset.EOF) Then
'                Me.adoAyuda.Recordset.MoveFirst
'            End If
'        End If
'    End If
End Sub

Private Sub Form_Load()
    'Coloca el formulario en el centro de la pantalla
    With Me
        If (Not lAgregar) Then
            .cmdNuevo.Enabled = False
        End If
    End With
    
    sFormaPadre = Forms(Forms.Count - 2).Name
    
    If (LlenaAyuda <= 0) Then
        Me.cmdElegir.Enabled = False
        MsgBox "No existen datos por mostrar.", vbExclamation, sTitAyuda
    Else
        Me.cmdElegir.Enabled = True
    End If
End Sub

'Cambia el orden en que aparecen los datos en pantalla
Private Sub dgAyuda_HeadClick(ByVal ColIndex As Integer)
'    ActualizaVista ColIndex
End Sub

'Asigna las propiedades a la forma
Public Function ConfigAyuda(mFAyuda() As String, mAAyuda() As Integer, mCAyuda() As String, mEAyuda() As String)
    Dim i As Byte
    
    ReDim mFormAyuda(nColsAyuda)
    ReDim mAnchAyuda(nColsAyuda)
    ReDim mCampAyuda(nColsAyuda)
    ReDim mEncaAyuda(nColsAyuda)
    
    For i = 0 To (nColsAyuda - 1)
        mFormAyuda(i) = mFAyuda(i)
        mAnchAyuda(i) = mAAyuda(i)
        mCampAyuda(i) = mCAyuda(i)
        mEncaAyuda(i) = mEAyuda(i)
    Next
End Function

'Llena el DBGrid de la forma
Private Function LlenaAyuda() As Integer
    Dim nRegistros As Integer
    
    nRegistros = 0
    
    'Inicializa los Ctrls Ado
    'InitCtrlAdoSel adoAyuda, sTabla, mCampAyuda, nColsAyuda, nColActiva, sCondicion, Conn
    
    Dim strSQL As String
    strSQL = QueryString(sTabla, mCampAyuda, nColsAyuda, sCondicion, nColActiva)
    
    Dim rsDatos As New ADODB.Recordset
    rsDatos.CursorLocation = adUseServer
    rsDatos.Open strSQL, Conn, adOpenStatic, adLockReadOnly
    
    'nRegistros = adoAyuda.Recordset.RecordCount
    nRegistros = rsDatos.RecordCount
    
    If (nRegistros > 0) Then
        Me.txtNRegistros.Text = nRegistros
        
        'Asigna los encabezados de las columnas
        DefHeadersDBGrid dgAyuda, mEncaAyuda
        
        'Asigna el ancho de cada columna
        DefAnchoDBGrid dgAyuda, mAnchAyuda
        
        'Evita que se puedan modificar los datos de la consulta
        dgAyuda.AllowUpdate = False
        dgAyuda.Caption = sTitAyuda
        
        Me.Caption = mFormAyuda(nColActiva)
    End If

    LlenaAyuda = nRegistros
    
    Set dgAyuda.DataSource = rsDatos
    
'    rsDatos.Close
'    Set rsDatos = Nothing
End Function

Private Sub LeeRegistro()
    Dim rsDatos As ADODB.Recordset
    Set rsDatos = dgAyuda.DataSource
    
    With rsDatos
        If .RecordCount > 0 Then
            Me.cmdElegir.Enabled = True
        Else
            Me.cmdElegir.Enabled = False
            Exit Sub
        End If
    
        Select Case sFormaPadre
        
            Case "frmAltaSocios"
                Select Case frmAltaSocios.nAyuda
                    Case 1
                        frmAltaSocios.txtSerie = .Fields("Serie")
                        frmAltaSocios.txtTipo = .Fields("Tipo")
                        frmAltaSocios.txtNumero = .Fields("Numero")
                        frmAltaSocios.txtTitNombre = Trim(.Fields("Nombre"))
                        frmAltaSocios.sTitPaterno = Trim(.Fields("A_Paterno"))
                        frmAltaSocios.sTitMaterno = Trim(.Fields("A_Materno"))
                        frmAltaSocios.txtCveAccionista.Text = .Fields("IdPropTitulo")
                        frmAltaSocios.txtTel1 = .Fields("Telefono_1")
                        frmAltaSocios.txtTel2 = .Fields("Telefono_2")
                        
                    Case 2
                        frmAltaSocios.txtCvePais.Text = .Fields("IdPais")
                        frmAltaSocios.txtPaisTit.Text = .Fields("Pais")
                    
                    Case 3
                        frmAltaSocios.txtCveTipo.Text = .Fields("IdTipoUsuario")
                        frmAltaSocios.txtTipoTit.Text = .Fields("Descripcion")
                End Select
                
            Case "frmAltaRenta"
                frmAltaRenta.cbTipoRenta.Text = .Fields("Descripcion")
                
                Select Case .Fields("Sexo")
                    Case "F"
                        frmAltaRenta.cbSexo.Text = "FEMENINO"
                    Case "M"
                        frmAltaRenta.cbSexo.Text = "MASCULINO"
                    Case "X"
                        frmAltaRenta.cbSexo.Text = "INDISTINTO"
                End Select
                
                frmAltaRenta.txtNoRenta.Text = .Fields("Numero")
                
            Case "frmAltaDir"
                frmAltaDir.txtCveDir.Text = .Fields("IdTipoDireccion")
                
            Case "frmAltaFam"
                Select Case frmAltaFam.nAyuda
                    Case 1
                        frmAltaFam.txtCvePaisFam.Text = .Fields("IdPais")
                        
                    Case 2
                        frmAltaFam.txtCveTipoFam.Text = .Fields("IdTipoUsuario")
                End Select
                
            Case "frmTZoneUsers"
                frmTZoneUsers.txtNoZona.Text = .Fields("IdTimeZone")
                
            Case "frmPases"
                Select Case frmPases.nAyuda
                    Case 1
                        frmPases.txtCve.Text = .Fields("IdCausa")
                        
                    Case 2
                        frmPases.txtCveZona.Text = .Fields("IdTimeZone")
                End Select
                
            Case "frmCertificados"
                frmCertificados.txtCveDr.Text = .Fields("IdMedico")
                frmCertificados.txtMedico.Text = Trim(.Fields("A_Paterno")) & " " & Trim(.Fields("A_Materno")) & " " & Trim(.Fields("Nombre"))
                
            Case "frmMembresia"
                Select Case frmMembresia.nAyuda
                    Case 1
                        frmMembresia.txtCveTit.Text = .Fields("idMember")
                        frmMembresia.txtNombre.Text = Trim(.Fields("Nombre")) & " " & Trim(.Fields("A_Paterno")) & " " & Trim(.Fields("A_Materno"))
                        
                    Case 2
                        frmMembresia.txtCveVendedor.Text = .Fields("idVendedor")
                        frmMembresia.txtVendedor.Text = Trim$(.Fields("Nombre"))
                End Select
            
        End Select
    End With
    
    Unload Me
End Sub

Private Sub ActualizaVista(nColumna As Integer)
    nColActiva = nColumna
    
    'Escribe el encabezado del formulario
    Me.Caption = mFormAyuda(nColActiva)
    
    'Asigna los encabezados de las columnas
    DefHeadersDBGrid dgAyuda, mEncaAyuda
    
    'Asigna el ancho de cada columna
    DefAnchoDBGrid dgAyuda, mAnchAyuda
    
    'Evita que se puedan modificar los datos de la consulta
    dgAyuda.AllowUpdate = False
    
    'Quita el efecto que deja la columna en color negro
    dgAyuda.ClearSelCols
    
    Dim strSQL As String
    strSQL = QueryString(sTabla, mCampAyuda, nColsAyuda, sCondicion, nColActiva)
    
    Dim rsDatos As New ADODB.Recordset
    rsDatos.CursorLocation = adUseServer
    rsDatos.Open strSQL, Conn, adOpenStatic, adLockReadOnly
    
    'If (adoAyuda.Recordset.RecordCount > 0) Then
    If (rsDatos.RecordCount > 0) Then
        Me.cmdElegir.Enabled = True
    Else
        Me.cmdElegir.Enabled = False
    End If
    
    Set dgAyuda.DataSource = rsDatos
    
    rsDatos.Close
    Set rsDatos = Nothing
End Sub

Public Function QueryString(ByVal sTablas As String, mCampos() As String, ByVal nTotCols As Byte, ByVal sCondicion As String, ByVal nColumnaActiva As Integer) As String
    Dim sListaCampos As String
    Dim nPointer As Integer
    Dim sRecSource As String

    sListaCampos = ""

    'Crea la lista de los campos que se deben mostrar
    For nPointer = 0 To (nTotCols - 2)
        sListaCampos = sListaCampos & mCampos(nPointer) & ", "
    Next nPointer
    sListaCampos = sListaCampos & mCampos(nTotCols - 1)
    
    sRecSource = "Select " & sListaCampos & " From " & sTabla
    
    If (sCondicion <> "") Then
        sRecSource = sRecSource & " Where " & sCondicion
    End If
    
    If InStr(mCampos(nColumnaActiva), "AS") Then
        sRecSource = sRecSource & " Order by " & Left(mCampos(nColumnaActiva), InStr(mCampos(nColumnaActiva), "AS") - 1)
    Else
        sRecSource = sRecSource & " Order by " & mCampos(nColumnaActiva)
    End If
    
    QueryString = sRecSource
End Function
