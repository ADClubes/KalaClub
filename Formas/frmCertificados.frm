VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCertificados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de certificados médicos"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmCertificados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotas 
      Height          =   1575
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2880
      Width           =   5415
   End
   Begin VB.TextBox txtCveUser 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   4920
      Picture         =   "frmCertificados.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   " Salir "
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdGuardar 
      Height          =   615
      Left            =   4080
      Picture         =   "frmCertificados.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " Guardar registro "
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtMedico 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   2280
      Width           =   4095
   End
   Begin VB.CommandButton cmdAyuda 
      Height          =   305
      Left            =   940
      Picture         =   "frmCertificados.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Lista  de médicos registrados "
      Top             =   2280
      Width           =   425
   End
   Begin VB.TextBox txtCveDr 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.ComboBox cbNombre 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   5415
   End
   Begin VB.TextBox txtTalla 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtPeso 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   67567617
      CurrentDate     =   38224
   End
   Begin VB.TextBox txtReg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblNotas 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lblMts 
      Caption         =   "metros"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   1600
      Width           =   615
   End
   Begin VB.Label lblKg 
      Caption         =   "kilogramos"
      Height          =   255
      Left            =   960
      TabIndex        =   20
      Top             =   1600
      Width           =   855
   End
   Begin VB.Label lblCveSocio 
      Caption         =   "# usuario"
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblMedico 
      Caption         =   "Nombre del médico"
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblCveDr 
      Caption         =   "# médico"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblSocio 
      Caption         =   "Nombre"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblTalla 
      Caption         =   "Estatura"
      Height          =   255
      Left            =   2040
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblPeso 
      Caption         =   "Peso"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblFecha 
      Caption         =   "Fecha"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblReg 
      Caption         =   "# Reg. "
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmCertificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de certificados medicos         *
'*  Daniel Hdez                                                 *
'*  27 / Agosto / 2004                                          *
'*  Ult Act: 16 / Agosto / 2005                                 *
'****************************************************************


Public bNvoCertif As Boolean

Dim sTextToolBar As String
Dim nRegCertif As Integer
Dim sNombre As String
Dim nPeso As Single
Dim nTalla As Single
Dim dFecha As Date
Dim sNota As String
Dim frmHDoctor As frmayuda



Private Sub cbNombre_LostFocus()
    If (Trim(Me.cbNombre.Text) <> "") Then
        Me.txtCveUser.Text = LeeXValor("IdMember", "Usuarios_Club", "(Trim(Nombre) & Chr(32) & Trim(A_Paterno) & Chr(32) & Trim(A_Materno))='" & Trim(Me.cbNombre.Text) & "'", "IdMember", "n", Conn)
        
        If (Val(Me.txtCveUser.Text) < 0) Then
            MsgBox "El usuario seleccionado no existe en la base de datos.", vbExclamation, "KalaSystems"
            Me.txtCveUser.Text = ""
            Me.cbNombre.Text = ""
            Me.cbNombre.SetFocus
        End If
    End If
End Sub


Private Sub cmdGuardar_Click()
    If (Cambios) Then
        If (GuardaDatos) Then
            InitVar
        Else
            MsgBox "No se registraron los datos, verifique la información.", vbCritical, "KalaSystems"
        End If
    End If
    
    Me.txtPeso.SetFocus
End Sub


Private Function ChecaDatos()
Dim sCond As String
Dim sCamp As String

    ChecaDatos = False
    
    If (bNvoCertif) Then
        If (Trim(Me.cbNombre.Text) = "") Then
            MsgBox "Se debe seleccionar un usuario.", vbExclamation, "KalaSystems"
            Me.cbNombre.SetFocus
            Exit Function
        End If
    End If
    
    If (Trim(Me.txtMedico.Text) = "") Then
        MsgBox "Se debe registrar el nombre del médico que expide el certificado.", vbInformation, "KalaSystems"
        Me.txtCveDr.SetFocus
        Exit Function
    End If
    
    If (Not IsNumeric(Me.txtPeso.Text)) Then
        MsgBox "El peso debe ser un valor numérico.", vbExclamation, "KalaSystems"
        Me.txtPeso.Text = ""
        Me.txtPeso.SetFocus
        Exit Function
    End If
    
    If ((Val(Me.txtPeso.Text) <= 0) Or (Val(Me.txtPeso.Text) > 250)) Then
        MsgBox "El peso debe ser un valor entre 1 y 250.", vbExclamation, "KalaSystems"
        Me.txtPeso.Text = ""
        Me.txtPeso.SetFocus
        Exit Function
    End If
    
    If (Not IsNumeric(Me.txtTalla.Text)) Then
        MsgBox "La estatura debe ser un valor numérico.", vbExclamation, "KalaSystems"
        Me.txtTalla.Text = ""
        Me.txtTalla.SetFocus
        Exit Function
    End If
    
    If ((Val(Me.txtTalla.Text) < 0) Or (Val(Me.txtTalla.Text) > 2.5)) Then
        MsgBox "La estatura debe ser un valor entre 0 y 2.5.", vbExclamation, "KalaSystems"
        Me.txtTalla.Text = ""
        Me.txtTalla.SetFocus
        Exit Function
    End If
    
    ChecaDatos = True
End Function


Private Function GuardaDatos() As Boolean
Const DATOSCERTIF = 7
Dim mFieldsCertif(DATOSCERTIF) As String
Dim mValuesCertif(DATOSCERTIF) As Variant

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Usuarios_Club
    mFieldsCertif(0) = "IdCertificado"
    mFieldsCertif(1) = "IdMember"
    mFieldsCertif(2) = "Fecha"
    mFieldsCertif(3) = "IdMedico"
    mFieldsCertif(4) = "Peso"
    mFieldsCertif(5) = "Estatura"
    mFieldsCertif(6) = "Observaciones"

    If (bNvoCertif) Then
        mValuesCertif(0) = LeeUltReg("Certificados", "IdCertificado") + 1
    Else
        mValuesCertif(0) = Val(Me.txtReg.Text)
    End If
    
    mValuesCertif(1) = Val(Me.txtCveUser.Text)
    mValuesCertif(2) = Format(Me.dtpFecha.Value, "dd/mm/yyyy")
    mValuesCertif(3) = Val(Me.txtCveDr.Text)
    mValuesCertif(4) = Val(Me.txtPeso.Text)
    mValuesCertif(5) = Val(Me.txtTalla.Text)
    mValuesCertif(6) = Trim(UCase(Me.txtNotas.Text))

    If (bNvoCertif) Then
        'Registra los datos de la nueva direccion
        If (AgregaRegistro("Certificados", mFieldsCertif, DATOSCERTIF, mValuesCertif, Conn)) Then
            MsgBox "Los datos se dieron de alta correctamente.", vbInformation, "KalaSystems"
            
            'Muestra el numero del registro de la direccion
            Me.txtReg.Text = mValuesCertif(0)
            Me.txtReg.REFRESH

            bNvoCertif = False
            GuardaDatos = True
        Else
            MsgBox "El registro no fue completado.", vbCritical, "KalaSystems"
        End If
    Else

        If (Val(Me.txtReg.Text) > 0) Then

            'Actualiza los datos de la ausencia
            If (CambiaReg("Certificados", mFieldsCertif, DATOSCERTIF, mValuesCertif, "IdCertificado=" & Val(Me.txtReg.Text), Conn)) Then
            
                MsgBox "Los datos se actualizaron correctamente.", vbInformation, "KalaSystems"
                GuardaDatos = True
            Else
                MsgBox "No se realizaron los cambios.", vbCritical, "KalaSystems"
            End If
        End If
    End If
End Function


Private Sub cmdSalir_Click()
Dim Respuesta As Integer

    If (Cambios) Then
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Certificados")
        
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
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Registro de certificados médicos"
    
    If (bNvoCertif) Then
        ClearCtrls
    Else
        LeeDatos
        
        Me.cbNombre.Enabled = False
    End If
    
    InitVar
End Sub


Private Sub LeeDatos()
Dim rsExamen As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String

    sCampos = "Certificados!Fecha, Certificados!Peso, Certificados!Estatura, "
    sCampos = sCampos & "Usuarios_Club!Nombre, Usuarios_Club!A_Paterno, Usuarios_Club!A_Materno, "
    sCampos = sCampos & "Medicos!Nombre, Medicos!A_Paterno, Medicos!A_Materno, "
    sCampos = sCampos & "Certificados!idCertificado, Certificados!idMember, Certificados!idMedico, "
    sCampos = sCampos & "Certificados!Observaciones "
    
    sTablas = "(Certificados LEFT JOIN Usuarios_Club ON Certificados.idMember=Usuarios_Club.idMember) "
    sTablas = sTablas & "LEFT JOIN Medicos ON Certificados.IdMedico=Medicos.IdMedico "
    
    InitRecordSet rsExamen, sCampos, sTablas, "Certificados.idCertificado=" & frmOtrosDatos.ssdbExamMed.Columns(9).Value, "", Conn
    With rsExamen
        If (.RecordCount > 0) Then
            Me.dtpFecha.Value = .Fields("Certificados!Fecha")
            Me.txtPeso.Text = .Fields("Certificados!Peso")
            Me.txtTalla.Text = .Fields("Certificados!Estatura")
            Me.cbNombre.Text = Trim$(.Fields("Usuarios_Club!Nombre")) & " " & Trim$(.Fields("Usuarios_Club!A_Paterno")) & " " & Trim$(.Fields("Usuarios_Club!A_Materno"))
            Me.txtMedico.Text = Trim$(.Fields("Medicos!Nombre")) & " " & Trim$(.Fields("Medicos!A_Paterno")) & " " & Trim$(.Fields("Medicos!A_Materno"))
            Me.txtReg.Text = .Fields("Certificados!idCertificado")
            Me.txtCveUser.Text = .Fields("Certificados!idMember")
            Me.txtCveDr.Text = .Fields("Certificados!idMedico")
            Me.txtNotas.Text = IIf(IsNull(.Fields("Certificados!Observaciones")), "", Trim(.Fields("Certificados!Observaciones")))
        End If
        
        .Close
    End With
    Set rsExamen = Nothing
End Sub


Private Sub ClearCtrls()
    With Me
        .txtReg.Text = ""
        .txtCveUser.Text = ""
        .txtPeso.Text = "0"
        .txtTalla.Text = "0"
        .dtpFecha.Value = Format(Date, "dd/mm/yyyy")
        .txtCveDr.Text = ""
        .txtMedico.Text = ""
        .txtNotas.Text = ""
        
        'Llena el combo con la lista de los Estados de la Republica
        sSql = "SELECT (Trim(Nombre) & chr(32) & Trim(A_Paterno) & chr(32) & Trim(A_Materno)) AS Nombre, IdMember FROM Usuarios_Club WHERE NoFamilia=" & Val(frmOtrosDatos.txtFamilia.Text)
        LlenaCombos Me.cbNombre, sSql, "Nombre", "IdMember"
    End With
End Sub


Private Sub InitVar()
    With Me
        nRegCertif = Val(.txtReg.Text)
        sNombre = Trim(.cbNombre.Text)
        nPeso = Val(.txtPeso.Text)
        nTalla = Val(.txtTalla.Text)
        dFecha = .dtpFecha.Value
        sNota = Trim(.txtNotas.Text)
    End With
End Sub


Public Sub LlenaCertificados()
Const DATOSCERTIF = 13
Dim rsExamen As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String
Dim mAncCertif(DATOSCERTIF) As Integer
Dim mEncCertif(DATOSCERTIF) As String


    frmOtrosDatos.ssdbExamMed.RemoveAll

    sCampos = "Certificados!Fecha, Certificados!Peso, Certificados!Estatura, "
    sCampos = sCampos & "Usuarios_Club!Nombre, Usuarios_Club!A_Paterno, Usuarios_Club!A_Materno, "
    sCampos = sCampos & "Medicos!Nombre, Medicos!A_Paterno, Medicos!A_Materno, "
    sCampos = sCampos & "Certificados!idCertificado, Certificados!idMember, Certificados!idMedico, "
    sCampos = sCampos & "Certificados!Observaciones "
    
    sTablas = "(Certificados LEFT JOIN Usuarios_Club ON Certificados.idMember=Usuarios_Club.idMember) "
    sTablas = sTablas & "LEFT JOIN Medicos ON Certificados.IdMedico=Medicos.IdMedico "
    
    InitRecordSet rsExamen, sCampos, sTablas, "Usuarios_Club.NoFamilia=" & Val(frmOtrosDatos.txtFamilia.Text), "", Conn
    With rsExamen
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                frmOtrosDatos.ssdbExamMed.AddItem Format(.Fields("Certificados!Fecha"), "dd / mmm / yyyy") & vbTab & _
                .Fields("Certificados!Peso") & vbTab & _
                .Fields("Certificados!Estatura") & vbTab & _
                .Fields("Usuarios_Club!Nombre") & vbTab & _
                .Fields("Usuarios_Club!A_Paterno") & vbTab & _
                .Fields("Usuarios_Club!A_Materno") & vbTab & _
                .Fields("Medicos!Nombre") & vbTab & _
                .Fields("Medicos!A_Paterno") & vbTab & _
                .Fields("Medicos!A_Materno") & vbTab & _
                .Fields("Certificados!idCertificado") & vbTab & _
                .Fields("Certificados!idMember") & vbTab & _
                .Fields("Certificados!idMedico") & vbTab & _
                .Fields("Certificados!Observaciones")
            
                .MoveNext
            Loop
        End If
        
        .Close
    End With
    Set rsExamen = Nothing
    
    'Asigna valores a la matriz de encabezados
    mEncCertif(0) = "Fecha"
    mEncCertif(1) = "Peso"
    mEncCertif(2) = "Estatura"
    mEncCertif(3) = "Nombre"
    mEncCertif(4) = "A. paterno"
    mEncCertif(5) = "A. materno"
    mEncCertif(6) = "Médico (nombre)"
    mEncCertif(7) = "Médico (A. paterno)"
    mEncCertif(8) = "Médico (A. Materno)"
    mEncCertif(9) = "# Reg."
    mEncCertif(10) = "# Usuario"
    mEncCertif(11) = "# Médico"
    mEncCertif(12) = "Observaciones"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid frmOtrosDatos.ssdbExamMed, mEncCertif
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncCertif(0) = 1500
    mAncCertif(1) = 900
    mAncCertif(2) = 900
    mAncCertif(3) = 2500
    mAncCertif(4) = 2500
    mAncCertif(5) = 2500
    mAncCertif(6) = 2500
    mAncCertif(7) = 2500
    mAncCertif(8) = 2500
    mAncCertif(9) = 900
    mAncCertif(10) = 900
    mAncCertif(11) = 900
    mAncCertif(12) = 3500

    'Asigna el ancho de cada columna
    DefAnchossGrid frmOtrosDatos.ssdbExamMed, mAncCertif
    
    frmOtrosDatos.ssdbExamMed.Columns(0).Alignment = ssCaptionAlignmentCenter
    frmOtrosDatos.ssdbExamMed.Columns(1).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbExamMed.Columns(2).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbExamMed.Columns(9).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbExamMed.Columns(10).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbExamMed.Columns(11).Alignment = ssCaptionAlignmentRight
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmCertificados.LlenaCertificados
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub


'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    Cambios = True
    
    If (nRegCertif <> Val(Me.txtReg.Text)) Then
        Exit Function
    End If
    
    If (sNombre <> Trim(Me.cbNombre.Text)) Then
        Exit Function
    End If
    
    If (nPeso <> Val(Me.txtPeso.Text)) Then
        Exit Function
    End If
    
    If (nTalla <> Val(Me.txtTalla.Text)) Then
        Exit Function
    End If

    If (dFecha <> Format(Me.dtpFecha.Value, "dd/mm/yyyy")) Then
        Exit Function
    End If
    
    If (sNota <> Trim(Me.txtNotas.Text)) Then
        Exit Function
    End If

    Cambios = False
End Function


Private Sub txtCveDr_LostFocus()
Dim rsMedico As ADODB.Recordset

    If (Trim(Me.txtCveDr.Text) <> "") Then
        If (IsNumeric(Me.txtCveDr.Text)) Then
            InitRecordSet rsMedico, "A_Paterno, A_Materno, Nombre", "Medicos", "IdMedico=" & Val(Me.txtCveDr.Text), "", Conn
            
            With rsMedico
                If (rsMedico.RecordCount > 0) Then
                    Me.txtMedico.Text = Trim(.Fields("A_Paterno")) & " " & Trim(.Fields("A_Materno")) & " " & Trim(.Fields("Nombre"))
                Else
                    MsgBox "La clave del médico es incorrecta.", vbExclamation, "KalaSystems"
                    Me.txtMedico.Text = ""
                    Me.txtCveDr.Text = ""
                    Me.txtCveDr.SetFocus
                End If
                
                .Close
            End With
            
            Set rsMedico = Nothing
        Else
            MsgBox "La clave del médico debe ser un valor numérico.", vbExclamation, "KalaSystems"
            Me.txtMedico.Text = ""
            Me.txtCveDr.Text = ""
            Me.txtCveDr.SetFocus
        End If
    End If
    
    Me.txtCveDr.REFRESH
End Sub




'************************************************************
'*                          Ayudas                          *
'************************************************************

Private Sub cmdAyuda_Click()
Const DATOSDR = 4
Dim sCadena As String
Dim mFAyuda(DATOSDR) As String
Dim mAAyuda(DATOSDR) As Integer
Dim mCAyuda(DATOSDR) As String
Dim mEAyuda(DATOSDR) As String

    nAyuda = 1

    Set frmHMedico = New frmayuda
    
    mFAyuda(0) = "Médicos ordenados por clave"
    mFAyuda(1) = "Médicos ordenados por nombre"
    
    mAAyuda(0) = 800
    mAAyuda(1) = 1500
    mAAyuda(2) = 1500
    mAAyuda(3) = 1500
    
    mCAyuda(0) = "IdMedico"
    mCAyuda(1) = "A_Paterno"
    mCAyuda(2) = "A_Materno"
    mCAyuda(3) = "Nombre"
    
    mEAyuda(0) = "Clave"
    mEAyuda(1) = "A. paterno"
    mEAyuda(2) = "A. materno"
    mEAyuda(3) = "Nombre"
    
    With frmHMedico
        .nColActiva = 0
        .nColsAyuda = DATOSDR
        .sTabla = "Medicos"
        
        .sCondicion = ""
        .sTitAyuda = "Médicos"
        .lAgregar = True
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
'    If (Trim(Me.txtCveDr.Text) <> "") Then
'        sCadena = "(Trim(Medicos!A_Paterno) & chr(32) & Trim(Medicos!A_Materno) & chr(32) & Trim(Medicos!Nombre)) AS Medico"
'        Me.txtMedico.Text = LeeXValor(sCadena, "Medicos", "IdMedico=" & Val(Me.txtCveDr.Text), "", "s", Conn)
'    End If
    
    Me.cmdAyuda.SetFocus
End Sub
