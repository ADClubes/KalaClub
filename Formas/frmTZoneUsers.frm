VERSION 5.00
Begin VB.Form frmTZoneUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horarios de acceso de los usuarios"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "frmTZoneUsers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSec 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtCveUser 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cbNombre 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox txtReg 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   4920
      Picture         =   "frmTZoneUsers.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdGuardar 
      Height          =   615
      Left            =   4080
      Picture         =   "frmTZoneUsers.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdHZonas 
      Height          =   305
      Left            =   960
      Picture         =   "frmTZoneUsers.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   425
   End
   Begin VB.TextBox txtZona 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   4095
   End
   Begin VB.TextBox txtNoZona 
      Height          =   285
      Left            =   240
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblSec 
      Caption         =   "# Secuencial"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblCveUser 
      Caption         =   "# Usuario"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblZona 
      Caption         =   "Descripción de la zona"
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblCveZona 
      Caption         =   "# Zona"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre del usuario"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblReg 
      Caption         =   "# Reg."
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmTZoneUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de los horarios de los usuarios *
'*  Daniel Hdez                                                 *
'*  25 / Agosto / 2004                                          *
'*  Ult Act: 16 / Agosto / 2005                                 *
'****************************************************************


Dim frmHZonas As frmayuda
Dim sTextToolBar As String
Dim nRegZona As Integer
Dim sNombre As String
Dim nZona As Integer
Dim nZonaAnt As Integer
Public bNvaZona As Boolean


Private Sub cbNombre_LostFocus()
    Me.txtCveUser.Text = LeeXValor("idMember", "Usuarios_Club", "(Trim(Nombre) & Chr(32) & Trim(A_Paterno) & Chr(32) & Trim(A_Materno))='" & Trim(Me.cbNombre.Text) & "'", "IdMember", "n", Conn)
    If (Val(Me.txtCveUser.Text) < 0) Then
        MsgBox "El usuario seleccionado no existe en la base de datos.", vbExclamation, "KalaSystems"
        Me.txtCveUser.Text = ""
        Me.cbNombre.Text = ""
        Me.cbNombre.SetFocus
    Else
        Me.txtSec.Text = LeeXValor("Secuencial", "Secuencial", "IdMember=" & Val(Me.txtCveUser.Text), "Secuencial", "n", Conn)
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
    
    Me.txtNoZona.SetFocus
End Sub


Private Function ChecaDatos()
Dim sCond As String
Dim sCamp As String

    ChecaDatos = False
    
    If (bNvaZona) Then
        If (Trim(Me.cbNombre.Text) = "") Then
            MsgBox "Se debe seleccionar un usuario.", vbExclamation, "KalaSystems"
            Me.cbNombre.SetFocus
            Exit Function
        End If
    End If
    
    If (Trim(Me.txtNoZona.Text) = "") Then
        MsgBox "Se debe seleccionar una zona.", vbExclamation, "KalaSystems"
        Me.txtNoZona.SetFocus
        Exit Function
    End If
    
    ChecaDatos = True
End Function


Private Function GuardaDatos() As Boolean
Const DATOSTZONE = 4
Dim mFieldsZone(DATOSTZONE) As String
Dim mValuesZone(DATOSTZONE) As Variant
Dim nInitTrans As Long

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Usuarios_Club
    mFieldsZone(0) = "IdReg"
    mFieldsZone(1) = "NoFamilia"
    mFieldsZone(2) = "IdMember"
    mFieldsZone(3) = "IdTimeZone"

    If (bNvaZona) Then
        mValuesZone(0) = LeeUltReg("Time_Zone_Users", "IdReg") + 1
    Else
        mValuesZone(0) = Val(Me.txtReg.Text)
    End If
    
    mValuesZone(1) = Val(frmOtrosDatos.txtFamilia.Text)
    mValuesZone(2) = Val(Me.txtCveUser.Text)
    mValuesZone(3) = Val(Me.txtNoZona.Text)

    If (bNvaZona) Then
        InitTrans = Conn.BeginTrans
        
        'Registra los datos de la nueva direccion
        If (AgregaRegistro("Time_Zone_Users", mFieldsZone, DATOSTZONE, mValuesZone, Conn)) Then
            MsgBox "Los datos se dieron de alta correctamente.", vbInformation, "KalaSystems"
            
            #If SqlServer_ Then
                ActivaCredSQL 0, Val(Me.txtSec.Text), Val(Me.txtNoZona.Text), CInt(mValuesZone(2)), True, False
            #Else
                ActivaCred 0, Val(Me.txtSec.Text), Val(Me.txtNoZona.Text), CInt(mValuesZone(2)), True, False
            #End If
            
            'Escribe en disco las actualizaciones
            Conn.CommitTrans
        
            'Muestra el numero del registro
            Me.txtReg.Text = mValuesZone(0)
            Me.txtReg.REFRESH
            
            LlenaTZoneUsers

            bNvaZona = False
            GuardaDatos = True
        Else
            'En caso de algun error no baja a disco los nuevos datos
            If InitTrans > 0 Then
                Conn.RollbackTrans
            End If
            
            MsgBox "El registro no fue completado.", vbCritical, "KalaSystems"
        End If
    Else

        If (Val(Me.txtReg.Text) > 0) Then

            'Actualiza los datos de la zona
            If (CambiaReg("Time_Zone_Users", mFieldsZone, DATOSTZONE, mValuesZone, "IdReg=" & Val(Me.txtReg.Text), Conn)) Then
            
                'Si no tenia zona asignada anteriormente se salta el proceso de desactivar
                If (nZonaAnt > 0) Then
                    'Desactiva la credencial de la zona anterior
                    #If SqlServer_ Then
                        ActivaCredSQL 0, Val(Me.txtSec.Text), nZonaAnt, CInt(mValuesZone(2)), False, False
                    #Else
                        ActivaCred 0, Val(Me.txtSec.Text), nZonaAnt, CInt(mValuesZone(2)), False, False
                    #End If
                End If
            
                'Activa la credencial en la nueva zona
                #If SqlServer_ Then
                    ActivaCredSQL 0, Val(Me.txtSec.Text), Val(Me.txtNoZona.Text), CInt(mValuesZone(2)), True, False
                #Else
                    ActivaCred 0, Val(Me.txtSec.Text), Val(Me.txtNoZona.Text), CInt(mValuesZone(2)), True, False
                #End If
            
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
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Asignación de zonas")
        
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
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Asignación de zonas"
    
    If (bNvaZona) Then
        ClearCtrls
    Else
        LeeDatos
        
        Me.cbNombre.Enabled = False
    End If
    
    InitVar
End Sub


Private Sub LeeDatos()
Dim rsTZonas As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String


    sCampos = "Time_Zone_Users!idReg, Time_Zone_Users!idMember, Secuencial!Secuencial, "
    sCampos = sCampos & "Usuarios_Club!Nombre, Usuarios_Club!A_Paterno, Usuarios_Club!A_Materno, "
    sCampos = sCampos & "Time_Zone_Users!idTimeZone, Time_Zone!Descripcion "
    
    sTablas = "((Time_Zone_Users LEFT JOIN Usuarios_Club ON Time_Zone_Users.IdMember=Usuarios_Club.IdMember) "
    sTablas = sTablas & "LEFT JOIN Time_Zone ON Time_Zone_Users.IdTimeZone=Time_Zone.IdTimeZone) "
    sTablas = sTablas & "LEFT JOIN Secuencial ON Usuarios_Club.IdMember=Secuencial.IdMember "
    
    sCadena = sCadena & "WHERE Usuarios_Club.NoFamilia=" & Val(frmOtrosDatos.txtFamilia.Text) & " AND "
    sCadena = sCadena & "NOT (Secuencial.Temporal) "
    
    InitRecordSet rsTZonas, sCampos, sTablas, "Time_Zone_Users.idReg=" & frmOtrosDatos.ssdbTZone.Columns(0).Value, "", Conn
    With rsTZonas
        If (.RecordCount > 0) Then
            Me.txtReg.Text = .Fields("Time_Zone_Users!idReg")
            Me.txtCveUser.Text = .Fields("Time_Zone_Users!idMember")
            Me.txtSec.Text = .Fields("Secuencial!Secuencial")
            Me.cbNombre.Text = .Fields("Usuarios_Club!Nombre") & " " & .Fields("Usuarios_Club!A_Paterno") & " " & .Fields("Usuarios_Club!A_Materno")
            Me.txtNoZona.Text = IIf(.Fields("Time_Zone_Users!idTimeZone") > 0, .Fields("Time_Zone_Users!IdTimeZone"), "")
            Me.txtZona.Text = IIf(Trim(.Fields("Time_Zone!Descripcion")) <> "", Trim(.Fields("Time_Zone!Descripcion")), "")
        End If
        
        .Close
    End With
    Set rsTZonas = Nothing
End Sub


Private Sub ClearCtrls()
    With Me
        .txtReg.Text = ""
        .txtNoZona.Text = ""
        .txtZona.Text = ""
        
        'Llena el combo con la lista de los Estados de la Republica
        sSql = "SELECT (Trim(Nombre) & chr(32) & Trim(A_Paterno) & chr(32) & Trim(A_Materno)) AS Nombre, IdMember FROM Usuarios_Club WHERE NoFamilia=" & Val(frmOtrosDatos.txtFamilia.Text)
        LlenaCombos Me.cbNombre, sSql, "Nombre", "IdMember"
    End With
End Sub


Private Sub InitVar()
    nRegZona = Val(Me.txtReg.Text)
    sNombre = Trim(Me.cbNombre.Text)
    nZona = Val(Me.txtNoZona.Text)
    nZonaAnt = Val(Me.txtNoZona.Text)
End Sub


Public Sub LlenaTZoneUsers()
Const DATOSTZONE = 9
Dim rsTZonas As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String
Dim mAncTZone(DATOSTZONE) As Integer
Dim mEncTZone(DATOSTZONE) As String


    frmOtrosDatos.ssdbTZone.RemoveAll
    
    sCampos = "Time_Zone_Users!idReg, Time_Zone_Users!idMember, "
    sCampos = sCampos & "Secuencial!Secuencial, Secuencial!Temporal, "
    sCampos = sCampos & "Usuarios_Club!Nombre, Usuarios_Club!A_Paterno, Usuarios_Club!A_Materno, "
    sCampos = sCampos & "Time_Zone_Users!idTimeZone, Time_Zone!Descripcion "
    
    sTablas = "((Time_Zone_Users LEFT JOIN Usuarios_Club ON Time_Zone_Users.IdMember=Usuarios_Club.IdMember) "
    sTablas = sTablas & "LEFT JOIN Time_Zone ON Time_Zone_Users.IdTimeZone=Time_Zone.IdTimeZone) "
    sTablas = sTablas & "LEFT JOIN Secuencial ON Usuarios_Club.IdMember=Secuencial.IdMember "
    
    sCadena = sCadena & "WHERE Usuarios_Club.NoFamilia=" & Val(frmOtrosDatos.txtFamilia.Text) & " AND "
    sCadena = sCadena & "NOT (Secuencial.Temporal) "
    
    InitRecordSet rsTZonas, sCampos, sTablas, "Usuarios_Club.NoFamilia=" & Val(frmOtrosDatos.txtFamilia.Text), "", Conn
    With rsTZonas
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                frmOtrosDatos.ssdbTZone.AddItem .Fields("Time_Zone_Users!idReg") & vbTab & _
                .Fields("Time_Zone_Users!idMember") & vbTab & _
                .Fields("Secuencial!Secuencial") & vbTab & _
                IIf(.Fields("Secuencial!Temporal"), 1, 0) & vbTab & _
                .Fields("Usuarios_Club!Nombre") & vbTab & _
                .Fields("Usuarios_Club!A_Paterno") & vbTab & _
                .Fields("Usuarios_Club!A_Materno") & vbTab & _
                .Fields("Time_Zone_Users!idTimeZone") & vbTab & _
                .Fields("Time_Zone!Descripcion")
            
                .MoveNext
            Loop
        End If
    
        .Close
    End With
    Set rsTZonas = Nothing

    'Asigna valores a la matriz de encabezados
    mEncTZone(0) = "# Reg."
    mEncTZone(1) = "# Usuario"
    mEncTZone(2) = "# Sec."
    mEncTZone(3) = "Temporal"
    mEncTZone(4) = "Nombre"
    mEncTZone(5) = "A. paterno"
    mEncTZone(6) = "A. materno"
    mEncTZone(7) = "# Zona"
    mEncTZone(8) = "Zona"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid frmOtrosDatos.ssdbTZone, mEncTZone
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncTZone(0) = 900
    mAncTZone(1) = 1100
    mAncTZone(2) = 900
    mAncTZone(3) = 1000
    mAncTZone(4) = 2500
    mAncTZone(5) = 2500
    mAncTZone(6) = 2500
    mAncTZone(7) = 900
    mAncTZone(8) = 3500

    'Asigna el ancho de cada columna
    DefAnchossGrid frmOtrosDatos.ssdbTZone, mAncTZone
    
    frmOtrosDatos.ssdbTZone.Columns(0).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbTZone.Columns(1).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbTZone.Columns(2).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbTZone.Columns(7).Alignment = ssCaptionAlignmentRight
    
    frmOtrosDatos.ssdbTZone.Columns(3).Style = ssStyleCheckBox
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmTZoneUsers.LlenaTZoneUsers
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub


Private Sub txtNoZona_LostFocus()
    If (Trim(Me.txtNoZona.Text) <> "") Then
        If (IsNumeric(Me.txtNoZona.Text)) Then
            Me.txtZona.Text = LeeXValor("Descripcion", "Time_Zone", "IdTimeZone=" & Val(Me.txtNoZona.Text), "Descripcion", "s", Conn)
            
            If (Trim(Me.txtZona.Text) = "VACIO") Then
                MsgBox "La zona seleccionada no existe en la base de datos.", vbExclamation, "KalaSystems"
                Me.txtZona.Text = ""
                Me.txtNoZona.Text = ""
                Me.txtNoZona.SetFocus
            End If
        Else
            MsgBox "La clave de la zona es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtZona.Text = ""
            Me.txtNoZona.Text = ""
            Me.txtNoZona.SetFocus
        End If
    End If
    
    Me.txtNoZona.REFRESH
End Sub


'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    Cambios = True
    
    If (nRegZona <> Val(Me.txtReg.Text)) Then
        Exit Function
    End If
    
    If (sNombre <> Trim(Me.cbNombre.Text)) Then
        Exit Function
    End If

    If (nZona <> Val(Me.txtNoZona.Text)) Then
        Exit Function
    End If

    Cambios = False
End Function





'************************************************************
'*                          Ayudas                          *
'************************************************************


Private Sub cmdHZonas_Click()
Const DATOSZONA = 2
Dim sCadena As String
Dim mFAyuda(DATOSZONA) As String
Dim mAAyuda(DATOSZONA) As Integer
Dim mCAyuda(DATOSZONA) As String
Dim mEAyuda(DATOSZONA) As String


    nAyuda = 1

    Set frmHZonas = New frmayuda
    
    mFAyuda(0) = "Zonas horarias ordenadas por clave"
    mFAyuda(1) = "Zonas horarias ordenadss por descripción"
    
    mAAyuda(0) = 800
    mAAyuda(1) = 3800
    
    mCAyuda(0) = "IdTimeZone"
    mCAyuda(1) = "Descripcion"
    
    mEAyuda(0) = "# Zona"
    mEAyuda(1) = "Descripción"
    
    With frmHZonas
        .nColActiva = 1
        .nColsAyuda = DATOSZONA
        .sTabla = "Time_Zone"
        
        .sCondicion = ""
        .sTitAyuda = "Zonas horarias"
        .lAgregar = True
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
    If (Trim(Me.txtNoZona.Text) <> "") Then
        Me.txtZona.Text = LeeXValor("Descripcion", "Time_Zone", "IdTimeZone=" & Val(Me.txtNoZona.Text), "Descripcion", "s", Conn)
    End If
    
    Me.cmdHZonas.SetFocus
End Sub
