VERSION 5.00
Begin VB.Form frmAutos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de autos"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   Icon            =   "frmAutos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDir 
      Height          =   1575
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6825
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1080
         Width           =   6495
      End
      Begin VB.TextBox txtIdPlaca 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtPlaca 
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtCalcomania 
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   615
         Left            =   6000
         Picture         =   "frmAutos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Salir "
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdGuardar 
         Height          =   615
         Left            =   5280
         Picture         =   "frmAutos.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Guardar registro "
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblDescrip 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblIdPlaca 
         Caption         =   "# Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblCalcomania 
         Caption         =   "# de calcomania"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblPlaca 
         Caption         =   "Placas"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmAutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de calcomanias                  *
'*  Daniel Hdez                                                 *
'*  19 / Julio / 2004                                           *
'*  Ult Act: 16 / Agosto / 2005                                 *
'****************************************************************


Public bNvaCal As Boolean

Dim sTextToolBar As String
Dim sCalco As String
Dim sPlaca As String
Dim sDecrip As String
Dim nCveAuto As Integer


Private Sub cmdSalir_Click()
Dim Respuesta As Integer

    If (Cambios) Then
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Registro de autos")
        
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
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Registro de autos"
    
    ClearCtrls
    
    'Clave para el nuevo registro
    nCveAuto = 0
    
    If (Not bNvaCal) Then
        LeeDatos
        
        Me.txtCalcomania.Enabled = False
    End If
    
    InitVar
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmAutos.LlenaAutos
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub


Private Sub ClearCtrls()
    With Me
        .txtIdPlaca.Text = ""
        .txtPlaca.Text = ""
        .txtCalcomania.Text = ""
        .txtDescripcion.Text = ""
    End With
End Sub


Private Sub InitVar()
    sPlaca = Trim(Me.txtPlaca.Text)
    sCalco = Trim(Me.txtCalcomania.Text)
    sDecrip = Trim(Me.txtDescripcion.Text)
End Sub


Private Sub cmdGuardar_Click()
Dim i As Byte

    If (Cambios) Then
        If (GuardaDatos) Then
            'Inicializa las variables
            InitVar
        Else
            MsgBox "No se registraron los datos, verifique la información.", vbCritical, "KalaSystems"
        End If
    End If
    
    Me.txtPlaca.SetFocus
End Sub


Private Function ChecaDatos()
Dim sCond As String
Dim sCamp As String

    ChecaDatos = False
    
    If (Trim(Me.txtPlaca.Text) = "") Then
        MsgBox "Se debe registrar un número de placa.", vbExclamation, "KalaSystems"
        Me.txtPlaca.SetFocus
        Exit Function
    End If
    
    If (Trim(Me.txtCalcomania.Text) = "") Then
        MsgBox "Se debe registrar el número de calcomania entregada.", vbExclamation, "KalaSystems"
        Me.txtCalcomania.SetFocus
        Exit Function
    End If
    
    If (bNvaCal) Then
        If (ExisteXValor("Numero", "Calcomanias", "Numero='" & Trim(Me.txtCalcomania.Text) & "'", Conn, "")) Then
            MsgBox "El número de calcomania ya se encuentra en uso.", vbCritical, "KalaSystems"
            Me.txtCalcomania.Text = ""
            Me.txtCalcomania.SetFocus
            Exit Function
        End If
    End If
    
    ChecaDatos = True
End Function


Private Function GuardaDatos() As Boolean
Const DATOSAUTO = 6
Dim bCreado As Boolean
Dim mFieldsAuto(DATOSAUTO) As String
Dim mValuesAuto(DATOSAUTO) As Variant

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Usuarios_Club
    mFieldsAuto(0) = "Id"
    mFieldsAuto(1) = "IdMember"
    mFieldsAuto(2) = "Numero"
    mFieldsAuto(3) = "Fecha"
    mFieldsAuto(4) = "Placa"
    mFieldsAuto(5) = "Descripcion"

    If (bNvaCal) Then
        mValuesAuto(0) = LeeUltReg("Calcomanias", "Id") + 1
    Else
        mValuesAuto(0) = Val(Me.txtIdPlaca.Text)
    End If
    
    mValuesAuto(1) = Val(frmOtrosDatos.txtTitCve.Text)
    mValuesAuto(2) = Trim(Me.txtCalcomania.Text)
    mValuesAuto(3) = Format(Date, "dd/mm/yyyy")
    mValuesAuto(4) = Trim(UCase(Me.txtPlaca.Text))
    mValuesAuto(5) = Trim(UCase(Me.txtDescripcion.Text))

    If (bNvaCal) Then
        'Registra los datos de la nueva direccion
        If (AgregaRegistro("Calcomanias", mFieldsAuto, DATOSAUTO, mValuesAuto, Conn)) Then
            MsgBox "Los datos se dieron de alta correctamente.", vbInformation, "KalaSystems"
            
            'Muestra el numero del registro de la direccion
            Me.txtIdPlaca.Text = mValuesAuto(0)
            Me.txtIdPlaca.REFRESH

            bNvaCal = False
            GuardaDatos = True
        Else
            MsgBox "El registro no fue completado.", vbCritical, "KalaSystems"
        End If
    Else

        If (Val(Me.txtIdPlaca.Text) > 0) Then

            'Actualiza los datos del titular
            If (CambiaReg("Calcomanias", mFieldsAuto, DATOSAUTO, mValuesAuto, "Id=" & Val(Me.txtIdPlaca.Text), Conn)) Then
                MsgBox "Los datos se actualizaron correctamente.", vbInformation, "KalaSystems"
                GuardaDatos = True
            Else
                MsgBox "No se realizaron los cambios.", vbCritical, "KalaSystems"
            End If
        End If

    End If

    'Pasa a mayusculas el contenido de las cajas de texto
    CambiaAMayusculas
End Function


'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    If (sPlaca <> Trim(Me.txtPlaca.Text)) Then
        Cambios = True
        Exit Function
    End If

    If (sCalco <> Trim(Me.txtCalcomania.Text)) Then
        Cambios = True
        Exit Function
    End If

    Cambios = False
End Function


Private Sub CambiaAMayusculas()
    With Me
        .txtCalcomania.Text = UCase(.txtCalcomania.Text)
        .txtPlaca.Text = UCase(.txtPlaca.Text)
        .txtDescripcion.Text = UCase(.txtDescripcion.Text)
    End With
End Sub


Public Sub LlenaAutos()
Const DATOSAUTO = 5
Dim rsAutos As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String
Dim mAncAuto(DATOSAUTO) As Integer
Dim mEncAuto(DATOSAUTO) As String

    
    frmOtrosDatos.ssdbAutos.RemoveAll

    sCampos = "id, Placa, Numero, Fecha, Descripcion"
    
    sTablas = "Calcomanias "
    
    InitRecordSet rsAutos, sCampos, sTablas, "idMember=" & Val(frmOtrosDatos.txtTitCve.Text), "", Conn
    With rsAutos
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                frmOtrosDatos.ssdbAutos.AddItem .Fields("id") & vbTab & _
                .Fields("Placa") & vbTab & _
                .Fields("Numero") & vbTab & _
                Format(.Fields("Fecha"), "dd / mmm / yyyy") & vbTab & _
                .Fields("Descripcion")
                
                .MoveNext
            Loop
        End If
    
        .Close
    End With
    Set rsAutos = Nothing
    
    'Asigna valores a la matriz de encabezados
    mEncAuto(0) = "# Reg."
    mEncAuto(1) = "Placas"
    mEncAuto(2) = "# Cal."
    mEncAuto(3) = "Fecha"
    mEncAuto(4) = "Descripción"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid frmOtrosDatos.ssdbAutos, mEncAuto
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncAuto(0) = 800
    mAncAuto(1) = 1500
    mAncAuto(2) = 1000
    mAncAuto(3) = 1500
    mAncAuto(4) = 3900

    'Asigna el ancho de cada columna
    DefAnchossGrid frmOtrosDatos.ssdbAutos, mAncAuto
    
    frmOtrosDatos.ssdbAutos.Columns(0).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbAutos.Columns(2).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbAutos.Columns(3).Alignment = ssCaptionAlignmentCenter
End Sub


Private Sub LeeDatos()
Dim rsAutos As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String


    sCampos = "id, Placa, Numero, Fecha, Descripcion"
    
    sTablas = "Calcomanias "
    
    InitRecordSet rsAutos, sCampos, sTablas, "id=" & frmOtrosDatos.ssdbAutos.Columns(0).Value, "", Conn
    With rsAutos
        If (.RecordCount > 0) Then
            nCveAuto = .Fields("id")
            frmAutos.txtIdPlaca.Text = nCveAuto
        
            If (.Fields("Placa") <> "") Then
                frmAutos.txtPlaca.Text = .Fields("Placa")
            End If
            
            If (.Fields("Numero") <> "") Then
                frmAutos.txtCalcomania.Text = .Fields("Numero")
            End If
            
            If (.Fields("Descripcion") <> "") Then
                frmAutos.txtDescripcion.Text = .Fields("Descripcion")
            End If
        End If
        
        .Close
    End With
    Set rsAutos = Nothing
End Sub
