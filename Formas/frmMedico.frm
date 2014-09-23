VERSION 5.00
Begin VB.Form frmMedico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de médicos"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "frmMedico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDatos 
      Caption         =   " Datos generales "
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.CommandButton cmdSalir 
         Height          =   550
         Left            =   6000
         Picture         =   "frmMedico.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Salir "
         Top             =   1080
         Width           =   550
      End
      Begin VB.CommandButton cmdGuardar 
         Height          =   550
         Left            =   5280
         Picture         =   "frmMedico.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Guardar "
         Top             =   1080
         Width           =   550
      End
      Begin VB.TextBox txtCve 
         Height          =   285
         Left            =   3480
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtMaterno 
         Height          =   285
         Left            =   3480
         MaxLength       =   60
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox txtPaterno 
         Height          =   285
         Left            =   240
         MaxLength       =   60
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   240
         MaxLength       =   60
         TabIndex        =   3
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Label lblCve 
         Caption         =   "Cve"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblPaterno 
         Caption         =   "Apellido paterno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblMaterno 
         Caption         =   "Apellido materno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblNombre 
         Caption         =   "Nombre(s)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de familiares                   *
'*  Daniel Hdez                                                 *
'*  02 / Septiembre / 2004                                      *
'****************************************************************


Public bNvoDr As Boolean

Dim sFormaAnt As String
Dim sTextToolBar As String
Dim sPaterno As String
Dim sMaterno As String
Dim sNombre As String


Private Sub cmdSalir_Click()
Dim Respuesta As Integer

    If (Cambios) Then
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Registro de médicos")
        
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
    sTextToolBar = Trim(MDIPrincipal.StatusBar1.Panels.Item(1).Text)
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Registro de médicos"
    
    'Guarda el nombre de la forma inmediata anterior
    sFormaAnt = Forms(Forms.Count - 2).Name
    
    ClearCtrls
    
    InitVar
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    If (sFormaAnt = "frmAltaSocios") Then
'        frmAltaFam.LlenaFam
'    Else
'        frmDatosSocios.Refresca
'    End If
    
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub


Private Sub ClearCtrls()
    With frmMedico
        .txtCve.Text = ""
        .txtPaterno.Text = ""
        .txtMaterno.Text = ""
        .txtNombre.Text = ""
    End With
End Sub


Private Sub InitVar()
    sPaterno = Trim(Me.txtPaterno.Text)
    sMaterno = Trim(Me.txtMaterno.Text)
    sNombre = Trim(Me.txtNombre.Text)
End Sub


Private Sub cmdGuardar_Click()
    If (Cambios) Then
        If (GuardaDatos) Then
            Unload Me
        Else
            MsgBox "No se registraron los datos, verifique la información.", vbCritical, "KalaSystems"
        End If
    End If
End Sub


Private Function ChecaDatos()
    ChecaDatos = False
    
    If (Trim(Me.txtPaterno.Text) = "") Then
        MsgBox "El apellido paterno no puede quedar en blanco.", vbExclamation, "KalaSystems"
        Me.txtPaterno.SetFocus
        Exit Function
    End If
    
    If ((Trim(Me.txtMaterno.Text) = "") And (Trim(Me.txtNombre.Text) = "")) Then
        MsgBox "Al menos se debe escribir uno de los apellidos.", vbExclamation, "KalaSystems"
        Me.txtPaterno.SetFocus
        Exit Function
    End If
    
    ChecaDatos = True
End Function


Private Function GuardaDatos() As Boolean
Const DATOSDOCTOR = 4
Dim mFieldsDr(DATOSDOCTOR) As String
Dim mValuesDr(DATOSDOCTOR) As Variant

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Usuarios_Club
    mFieldsDr(0) = "IdMedico"
    mFieldsDr(1) = "Nombre"
    mFieldsDr(2) = "A_Paterno"
    mFieldsDr(3) = "A_Materno"
    
    'Valores para la tabla Usuarios_Club
    mValuesDr(0) = LeeUltReg("Medicos", "IdMedico") + 1
    mValuesDr(1) = Trim(UCase(Me.txtNombre.Text))
    mValuesDr(2) = Trim(UCase(Me.txtPaterno.Text))
    mValuesDr(3) = Trim(UCase(Me.txtMaterno.Text))

    If (bNvoDr) Then
        'Registra los datos de la nueva direccion
        If (AgregaRegistro("Medicos", mFieldsDr, DATOSDOCTOR, mValuesDr, Conn)) Then
        
            MsgBox "Los datos se dieron de alta correctamente.", vbInformation, "KalaSystems"
            
            'Muestra el numero del registro de la direccion
            Me.txtCve.Text = mValuesDr(0)
            Me.txtCve.Refresh
                            
            bNvoDr = False
            GuardaDatos = True
        Else
            MsgBox "El registro no fue completado.", vbCritical, "KalaSystems"
        End If
    End If
End Function


'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    Cambios = True

    If (sPaterno <> Trim(Me.txtPaterno.Text)) Then
        Exit Function
    End If
    
    If (sMaterno <> Trim(Me.txtMaterno.Text)) Then
        Exit Function
    End If
    
    If (sNombre <> Trim(Me.txtNombre.Text)) Then
        Exit Function
    End If

    Cambios = False
End Function
