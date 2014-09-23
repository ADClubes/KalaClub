VERSION 5.00
Begin VB.Form frmAgregaOps 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmAgregaOps.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbTipo 
      Height          =   315
      ItemData        =   "frmAgregaOps.frx":0442
      Left            =   1080
      List            =   "frmAgregaOps.frx":044F
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox cbParentesco 
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtEdadMax 
      Height          =   285
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtEdadMin 
      Height          =   285
      Left            =   120
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtClave 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtDescrip 
      Height          =   285
      Left            =   120
      MaxLength       =   60
      TabIndex        =   2
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton cmdGuardar 
      Height          =   615
      Left            =   3120
      Picture         =   "frmAgregaOps.frx":0471
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   3960
      Picture         =   "frmAgregaOps.frx":08B3
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblTipo 
      Caption         =   "Tipo"
      Height          =   255
      Left            =   1080
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblParentesco 
      Caption         =   "Parentesco"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblEdadMax 
      Caption         =   "Edad máxima"
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblEdadMin 
      Caption         =   "Edad mínima"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblClave 
      Caption         =   "Clave"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblDescrip 
      Caption         =   "Descripción"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmAgregaOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para registrar mas opciones en los catalogos     *
'*  Daniel Hdez                                                 *
'*  20 / Septiembre / 2004                                      *
'****************************************************************

Dim bGuardados As Boolean               'Controla si los datos han sido guardados

'Datos iniciales
Dim sDescrip As String
Dim nCve As Integer
Dim sTexto As String
Public nOpcion As Byte                  '1  Tipos de direccion
                                        '2  Opciones para pases temporales
                                        '3  Tipos de titular
                                        '4  Tipos de familiares
                                        '5  Paises
                                        '6  Time Zone

Private Sub cmdSalir_Click()
Dim Respuesta As Integer

    If (Cambios) Then
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "KalaSystems")
        
        If (Respuesta = vbYes) Then
            If (GuardaDatos) Then
                Unload Me
            Else
                Exit Sub
            End If
        End If
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()
Dim sCadena As String

    sTexto = Trim(MDIPrincipal.StatusBar1.Panels.Item(1).Text)
    Me.Height = 1830
    
    Select Case nOpcion
        Case 1
            MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Tipos de direcciones"
            Me.Caption = "Nuevas opciones para direcciones"
        
        Case 2
            MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Opciones para pases temporales"
            Me.Caption = "Nuevas opciones para pases temporales"
            
        Case 3
            MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Opciones para titulares"
            Me.Caption = "Nuevas opciones para tipos de titular"
            Me.Height = 2430
            Me.lblTipo.Visible = True
            Me.cbTipo.Visible = True
            Me.lblParentesco.Visible = False
            Me.cbParentesco.Visible = False
            
        Case 4
            MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Opciones para familiares"
            Me.Caption = "Nuevas opciones para tipos de familiar"
            Me.Height = 2430
            
            LlenaCbParentesco
            
        Case 5
            MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Opciones para países"
            Me.Caption = "Nuevas opciones para países"
            
        Case 6
            MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Opciones para zonas"
            Me.Caption = "Nuevas opciones para zonas"
    End Select
    
    'Datos guardados
    bGuardados = False
    
    'Limpia el contenido de las cajas de texto
    LimpiaVar
    
    'Inicializa las variables
    InitVar
End Sub


'Limpia las cajas de texto de la forma
Private Sub LimpiaVar()
    With Me
        .txtClave.Text = ""
        .txtDescrip.Text = ""
    End With
End Sub


'Inicializa las variables
Private Sub InitVar()
    With Me
        nCve = Val(.txtClave.Text)
        sDescrip = Trim(.txtDescrip.Text)
    End With
End Sub


'Llama al procedimiento que guarda los datos en la base
Private Sub cmdGuardar_Click()
    If (Cambios) Then
        If (GuardaDatos) Then
            Unload Me
        End If
    End If
End Sub


'Verifica que los haya datos en las cajas de texto
Public Function DatosCorrectos() As Boolean
Dim sTabla As String
Dim sMensaje As String

    Select Case nOpcion
        Case 1
            sTabla = "Tipo_Direccion"
            sMensaje = "Ya existe éste tipo de dirección."
            
        Case 2
            sTabla = "Causas_Pase"
            sMensaje = "Ya existe ésta opción."
            
        Case 3, 4
            sTabla = "Tipo_Usuario"
            
            If (nOpcion = 3) Then
                sMensaje = "Ya existe éste tipo de titular."
            Else
                sMensaje = "Ya existe éste tipo de familiar."
                
                If (Not ExisteXValor("Parentesco", "Parentesco", "Parentesco='" & Trim(Me.cbParentesco.Text) & "'", Conn, "")) Then
                    MsgBox "El parentesco selecionado es incorrecto.", vbExclamation, "KalaSystems"
                    Me.cbParentesco.Text = ""
                    Me.cbParentesco.SetFocus
                    Exit Function
                End If
            End If
            
            If (Not IsNumeric(Me.txtEdadMin.Text)) Then
                MsgBox "La edad mínima debe ser un valor numérico.", vbCritical, "KalaSystems"
                Me.txtEdadMin.Text = ""
                Me.txtEdadMin.SetFocus
                Exit Function
            End If
            
            If (Not IsNumeric(Me.txtEdadMax.Text)) Then
                MsgBox "La edad máxima debe ser un valor numérico.", vbCritical, "KalaSystems"
                Me.txtEdadMax.Text = ""
                Me.txtEdadMax.SetFocus
                Exit Function
            End If
            
            If (Val(Me.txtEdadMax.Text) <= Val(Me.txtEdadMin.Text)) Then
                MsgBox "La edad máxima debe ser mayor que la edad mínima.", vbCritical, "KalaSystems"
                Me.txtEdadMin.Text = ""
                Me.txtEdadMax.Text = ""
                Me.txtEdadMin.SetFocus
                Exit Function
            End If
            
            If (Val(Me.txtEdadMin.Text) < 0) Then
                MsgBox "La edad mínima debe ser mayor o igual a 1.", vbCritical, "KalaSystems"
                Me.txtEdadMin.Text = ""
                Me.txtEdadMin.SetFocus
                Exit Function
            End If
            
            If (Val(Me.txtEdadMax.Text) < 0) Then
                MsgBox "La edad máxima es 150.", vbCritical, "KalaSystems"
                Me.txtEdadMax.Text = ""
                Me.txtEdadMax.SetFocus
                Exit Function
            End If
            
        Case 5
            sTabla = "Paises"
            sMensaje = "El país ya está dado de alta."
            
        Case 6
            sTabla = "Time_Zone"
            sMensaje = "La zona ya existe."
    End Select
    
    If (Me.txtDescrip.Text = "") Then
        MsgBox "No se pueden dejar en blanco la descripción.", vbCritical, "KalaSystems"
        Me.txtDescrip.SetFocus
        DatosCorrectos = False
        Exit Function
    End If

    If (ExisteXValor(IIf(nOpcion = 5, "Pais", "Descripcion"), sTabla, IIf(nOpcion = 5, "Pais='", "Descripcion='") & Trim(UCase(Me.txtDescrip.Text)) & "'", Conn, "")) Then
        MsgBox sMensaje, vbCritical, "KalaSystems"
        Me.txtDescrip.SetFocus
        DatosCorrectos = False
        Exit Function
    End If

    DatosCorrectos = True
End Function


'Guarda la informacion en la base de datos
Private Function GuardaDatos() As Boolean
Const DATOSOPS = 2
Const DATOSTIPO = 7
Dim bCreado As Boolean
Dim mCampos() As String
Dim mValores() As Variant
Dim sTabla As String
Dim nCuantosDatos As Byte
    
    
    If (Not DatosCorrectos) Then
        GuardaDatos = False
        Exit Function
    End If
    
    nCuantosDatos = DATOSOPS
        
    Select Case nOpcion
        Case 1
            ReDim mCampos(DATOSOPS) As String
            ReDim mValores(DATOSOPS) As Variant
        
            mCampos(0) = "IdTipoDireccion"
            sTabla = "Tipo_Direccion"
            
        Case 2
            ReDim mCampos(DATOSOPS) As String
            ReDim mValores(DATOSOPS) As Variant
        
            mCampos(0) = "IdCausa"
            sTabla = "Causas_Pase"
            
        Case 3, 4
            ReDim mCampos(DATOSTIPO) As String
            ReDim mValores(DATOSTIPO) As Variant
            
            nCuantosDatos = DATOSTIPO
            sTabla = "Tipo_Usuario"
        
            mCampos(0) = "IdTipoUsuario"
            mCampos(2) = "EdadMinima"
            mCampos(3) = "EdadMaxima"
            mCampos(4) = "Parentesco"
            mCampos(5) = "Familiar"
            mCampos(6) = "Tipo"
            
            mValores(2) = Val(Me.txtEdadMin.Text)
            mValores(3) = Val(Me.txtEdadMax.Text)
            
            If (nOpcion = 3) Then
                mValores(4) = "TI"
                mValores(5) = 0             'Titular
                mValores(6) = Trim(UCase(Me.cbTipo.Text))
            Else
                mValores(4) = LeeXValor("Clave", "Parentesco", "Parentesco='" & Trim(Me.cbParentesco.Text) & "'", "Clave", "s", Conn)
                mValores(5) = 1             'Conyuge, hijo o dependiente
                mValores(6) = "INDISTINTO"
            End If
            
        Case 5
            ReDim mCampos(DATOSOPS) As String
            ReDim mValores(DATOSOPS) As Variant
        
            mCampos(0) = "IdPais"
            sTabla = "Paises"
            
        Case 6
            ReDim mCampos(DATOSOPS) As String
            ReDim mValores(DATOSOPS) As Variant
        
            mCampos(0) = "IdTimeZone"
            sTabla = "Time_Zone"
    End Select
    
    mCampos(1) = IIf(nOpcion = 5, "Pais", "Descripcion")
    
    mValores(0) = LeeUltReg(sTabla, mCampos(0)) + 1
    mValores(1) = UCase(Trim(Me.txtDescrip.Text))
    
    'Agrega un nuevo registro
    If (AgregaRegistro(sTabla, mCampos, nCuantosDatos, mValores, Conn)) Then
        MsgBox "Los datos se dieron de alta con la clave: " & mValores(0), vbInformation, "KalaSystems"
            
        bGuardados = True
        GuardaDatos = True
    Else
        MsgBox "El registro no fué completado.", vbCritical, "KalaSystems"
        Me.txtDescrip.SetFocus
    End If
End Function


Private Sub Form_Unload(Cancel As Integer)
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTexto
End Sub


'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    Cambios = True

    If (sDescrip <> Trim(Me.txtDescrip.Text)) Then
        Exit Function
    End If

    Cambios = False
End Function


'Llena el combo con las opciones del parentesco
Private Sub LlenaCbParentesco()
Dim rsParentesco As ADODB.Recordset
Dim i As Byte

    InitRecordSet rsParentesco, "Parentesco", "Parentesco", "Clave<>'TI'", "Parentesco", Conn
    If (rsParentesco.RecordCount > 0) Then
    
        i = 1
        Do While (Not rsParentesco.EOF)
            Me.cbParentesco.AddItem rsParentesco.Fields("Parentesco")
            Me.cbParentesco.ItemData(Me.cbParentesco.NewIndex) = i
            
            i = i + 1
            rsParentesco.MoveNext
        Loop
    End If
    
    rsParentesco.Close
    Set rsParentesco = Nothing
End Sub
