VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAusencias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de ausencias"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "frmAusencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPorcen 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpTermina 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   116326401
      CurrentDate     =   38223
   End
   Begin MSComCtl2.DTPicker dtpInicia 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      Format          =   116326401
      CurrentDate     =   38223
   End
   Begin VB.CommandButton cmdGuardar 
      Height          =   615
      Left            =   3960
      Picture         =   "frmAusencias.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   " Guardar registro "
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   4800
      Picture         =   "frmAusencias.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " Salir "
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtReg 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox cbNombre 
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   5175
   End
   Begin VB.TextBox txtCveUser 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblDescuen 
      Caption         =   "% descuento"
      Height          =   255
      Left            =   4320
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblTermina 
      Caption         =   "Termina"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblInicia 
      Caption         =   "Inicia"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblReg 
      Caption         =   "# Reg."
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre del usuario"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblCveUser 
      Caption         =   "# Usuario"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmAusencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de ausencias                    *
'*  Daniel Hdez                                                 *
'*  25 / Agosto / 2004                                          *
'*  Ult Act: 17 / Agosto / 2005                                 *
'****************************************************************


Dim sTextToolBar As String
Dim nRegAus As Integer
Dim sNombre As String
Dim nPorcen As Single
Dim dInicia As Date
Dim dTermina As Date
Public bNvaAus As Boolean



Private Sub cbNombre_LostFocus()
    'gpo 01/12/2005
    'Me.txtCveUser.Text = LeeXValor("idMember", "Usuarios_Club", "(Trim(Nombre) & Chr(32) & Trim(A_Paterno) & Chr(32) & Trim(A_Materno))='" & Trim(Me.cbNombre.Text) & "'", "IdMember", "n", Conn)
    If Me.cbNombre.ListIndex >= 0 Then
        Me.txtCveUser.Text = Me.cbNombre.ItemData(Me.cbNombre.ListIndex)
    End If
    If (Val(Me.txtCveUser.Text) < 0) Then
        MsgBox "El usuario seleccionado no existe en la base de datos.", vbExclamation, "KalaSystems"
        Me.txtCveUser.Text = ""
        Me.cbNombre.Text = ""
        Me.cbNombre.SetFocus
    End If
End Sub


Private Sub cmdGuardar_Click()

    If Not ChecaSeguridad(Me.Name, Me.cmdGuardar.Name) Then
        Exit Sub
    End If

    If (Cambios) Then
        If (GuardaDatos) Then
            InitVar
        Else
            MsgBox "No se registraron los datos, verifique la información.", vbCritical, "KalaSystems"
        End If
    End If
    
    Me.dtpInicia.SetFocus
End Sub


Private Function ChecaDatos()
Dim sCond As String
Dim sCamp As String

    ChecaDatos = False
    
    If (bNvaAus) Then
        If (Trim(Me.cbNombre.Text) = "") Then
            MsgBox "Se debe seleccionar un usuario.", vbExclamation, "KalaSystems"
            Me.cbNombre.SetFocus
            Exit Function
        End If
        
        If (TraslapaFechas(Val(Me.txtCveUser.Text), Me.dtpInicia.Value, Me.dtpTermina.Value)) Then
            MsgBox "Las fechas seleccionadas para este período" & Chr(13) & "de ausencia, se traslapan con otro período" & Chr(13) & "registrado anteriormente.", vbCritical, "KalaSystems"
            Me.dtpInicia.SetFocus
            Exit Function
        End If
    End If
    
'    If (Me.dtpInicia.Value < Date) Then
'        MsgBox "No se pueden registrar períodos de ausencia anteriores a la fecha actual.", vbInformation, "KalaSystems"
'        Me.dtpInicia.SetFocus
'        Exit Function
'    End If
    
    If (Me.dtpTermina.Value <= Me.dtpInicia.Value) Then
        MsgBox "La fecha final debe ser posterior a la inicial.", vbExclamation, "KalaSystems"
        Me.dtpInicia.SetFocus
        Exit Function
    End If
    
    If (Not IsNumeric(Me.txtPorcen.Text)) Then
        MsgBox "El porcentaje de descuento debe ser un valor numérico.", vbExclamation, "KalaSystems"
        Me.txtPorcen.Text = ""
        Me.txtPorcen.SetFocus
        Exit Function
    End If
    
    If ((Val(Me.txtPorcen.Text) > 100) Or (Val(Me.txtPorcen.Text) < 0)) Then
        MsgBox "El porcentaje de descuento debe ser un valor entre 0 y 100.", vbExclamation, "KalaSystems"
        Me.txtPorcen.Text = ""
        Me.txtPorcen.SetFocus
        Exit Function
    End If
    
    ChecaDatos = True
End Function


Private Function GuardaDatos() As Boolean
    Const DATOSFALTA = 5
    Dim mFieldsFalta(DATOSFALTA) As String
    Dim mValuesFalta(DATOSFALTA) As Variant

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Usuarios_Club
    mFieldsFalta(0) = "IdAusencia"
    mFieldsFalta(1) = "IdMember"
    mFieldsFalta(2) = "FechaInicial"
    mFieldsFalta(3) = "FechaFinal"
    mFieldsFalta(4) = "Porcentaje"

    If (bNvaAus) Then
        mValuesFalta(0) = LeeUltReg("Ausencias", "IdAusencia") + 1
    Else
        mValuesFalta(0) = Val(Me.txtReg.Text)
    End If
    
    #If SqlServer_ Then
        mValuesFalta(1) = Val(Me.txtCveUser.Text)
        mValuesFalta(2) = Format(Me.dtpInicia.Value, "yyyymmdd")
        mValuesFalta(3) = Format(Me.dtpTermina.Value, "yyyymmdd")
        mValuesFalta(4) = Val(Me.txtPorcen.Text)
    #Else
        mValuesFalta(1) = Val(Me.txtCveUser.Text)
        mValuesFalta(2) = Format(Me.dtpInicia.Value, "dd/mm/yyyy")
        mValuesFalta(3) = Format(Me.dtpTermina.Value, "dd/mm/yyyy")
        mValuesFalta(4) = Val(Me.txtPorcen.Text)
    #End If

    If (bNvaAus) Then
        'Registra los datos de la nueva ausencia
        If (AgregaRegistro("Ausencias", mFieldsFalta, DATOSFALTA, mValuesFalta, Conn)) Then
            MsgBox "Los datos se dieron de alta correctamente.", vbInformation, "KalaSystems"
            
            'Muestra el numero del registro de la ausencia
            Me.txtReg.Text = mValuesFalta(0)
            Me.txtReg.REFRESH

            bNvaAus = False
            GuardaDatos = True
        Else
            MsgBox "El registro no fue completado.", vbCritical, "KalaSystems"
        End If
        
    ActAccesoXUsu Val(Me.txtCveUser.Text), False
        
    Else

        If (Val(Me.txtReg.Text) > 0) Then

            'Actualiza los datos de la ausencia
            If (CambiaReg("Ausencias", mFieldsFalta, DATOSFALTA, mValuesFalta, "IdAusencia=" & Val(Me.txtReg.Text), Conn)) Then
            
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
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Ausencias")
        
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
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Registro de ausencias"
    
    If (bNvaAus) Then
        ClearCtrls
    Else
        LeeDatos
        
        Me.cbNombre.Enabled = False
    End If
    
    InitVar
End Sub


Private Sub LeeDatos()
Dim rsAusencia As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String


    sCampos = "Ausencias.IdAusencia, Ausencias.FechaInicial, "
    sCampos = sCampos & "Ausencias.FechaFinal, Ausencias.Porcentaje, "
    sCampos = sCampos & "Ausencias.IdMember, "
    sCampos = sCampos & "Usuarios_Club.Nombre, Usuarios_Club.A_Paterno, "
    sCampos = sCampos & "Usuarios_Club.A_Materno "
    
    sTablas = "Ausencias LEFT JOIN Usuarios_Club ON Ausencias.IdMember=Usuarios_Club.IdMember "
    
    InitRecordSet rsAusencia, sCampos, sTablas, "Ausencias.idAusencia=" & frmAltaSocios.ssdbAusencias.Columns(0).Value, "", Conn
    With rsAusencia
        If (.RecordCount > 0) Then
            Me.txtReg.Text = .Fields("IdAusencia")
            Me.txtCveUser.Text = .Fields("IdMember")
            Me.dtpInicia.Value = .Fields("FechaInicial")
            Me.dtpTermina.Value = .Fields("FechaFinal")
            Me.txtPorcen.Text = .Fields("Porcentaje")
            Me.cbNombre.Text = .Fields("Nombre") & " " & .Fields("A_Paterno") & " " & .Fields("A_Materno")
        End If
        
        .Close
    End With
    Set rsAusencia = Nothing
End Sub


Private Sub ClearCtrls()
    With Me
        .txtReg.Text = ""
        .dtpInicia.Value = Format(Date, "dd/mm/yyyy")
        .dtpTermina.Value = Format(Date, "dd/mm/yyyy")
        .txtPorcen.Text = "0"
        
        'Llena el combo con la lista de los Estados de la Republica
        #If SqlServer_ Then
            sSql = "SELECT (LTRIM(RTrim(Nombre)) + ' ' + LTRIM(RTrim(A_Paterno)) + ' ' + LTRIM(RTrim(A_Materno))) AS Nombre, IdMember FROM Usuarios_Club WHERE NoFamilia=" & Val(frmAltaSocios.txtFamilia.Text)
        #Else
            sSql = "SELECT (Trim(Nombre) & chr(32) & Trim(A_Paterno) & chr(32) & Trim(A_Materno)) AS Nombre, IdMember FROM Usuarios_Club WHERE NoFamilia=" & Val(frmAltaSocios.txtFamilia.Text)
        #End If
        LlenaCombos Me.cbNombre, sSql, "Nombre", "IdMember"
    End With
End Sub


Private Sub InitVar()
    With Me
        nRegAus = Val(.txtReg.Text)
        sNombre = Trim(.cbNombre.Text)
        nPorcen = Val(.txtPorcen.Text)
        dInicia = .dtpInicia.Value
        dTermina = .dtpTermina.Value
    End With
End Sub

Public Sub ActivaMiembro(pIdMember As String)

ActAccesoXUsu Val(pIdMember), True

End Sub
Public Sub LlenaAusencias()
Const DATOSFALTA = 8
Dim rsAusencia As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String
Dim mAncFalta(DATOSFALTA) As Integer
Dim mEncFalta(DATOSFALTA) As String



    frmAltaSocios.ssdbAusencias.RemoveAll

    sCampos = "Ausencias.IdAusencia, Ausencias.FechaInicial, "
    sCampos = sCampos & "Ausencias.FechaFinal, Ausencias.Porcentaje, "
    sCampos = sCampos & "Ausencias.IdMember, "
    sCampos = sCampos & "Usuarios_Club.Nombre, Usuarios_Club.A_Paterno, "
    sCampos = sCampos & "Usuarios_Club.A_Materno "
    
    sTablas = "Ausencias LEFT JOIN Usuarios_Club ON Ausencias.IdMember=Usuarios_Club.IdMember "
    
    InitRecordSet rsAusencia, sCampos, sTablas, "Usuarios_Club.NoFamilia=" & Val(frmAltaSocios.txtFamilia.Text), "", Conn
    With rsAusencia
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                frmAltaSocios.ssdbAusencias.AddItem .Fields("idAusencia") & vbTab & _
                Format(.Fields("FechaInicial"), "dd / mmm / yyyy") & vbTab & _
                Format(.Fields("FechaFinal"), "dd / mmm / yyyy") & vbTab & _
                .Fields("Porcentaje") & vbTab & _
                .Fields("idMember") & vbTab & _
                .Fields("Nombre") & vbTab & _
                .Fields("A_Paterno") & vbTab & _
                .Fields("A_Materno")
            
                .MoveNext
            Loop
        End If
        
        .Close
    End With
    Set rsAusencia = Nothing

    'Asigna valores a la matriz de encabezados
    mEncFalta(0) = "# Reg."
    mEncFalta(1) = "Inicia"
    mEncFalta(2) = "Termina"
    mEncFalta(3) = "% Desc."
    mEncFalta(4) = "# Usuario"
    mEncFalta(5) = "Nombre"
    mEncFalta(6) = "A. paterno"
    mEncFalta(7) = "A. materno"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid frmAltaSocios.ssdbAusencias, mEncFalta
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncFalta(0) = 800
    mAncFalta(1) = 1500
    mAncFalta(2) = 1500
    mAncFalta(3) = 1000
    mAncFalta(4) = 900
    mAncFalta(5) = 2500
    mAncFalta(6) = 2500
    mAncFalta(7) = 2500

    'Asigna el ancho de cada columna
    DefAnchossGrid frmAltaSocios.ssdbAusencias, mAncFalta
    
    frmAltaSocios.ssdbAusencias.Columns(0).Alignment = ssCaptionAlignmentRight
    frmAltaSocios.ssdbAusencias.Columns(1).Alignment = ssCaptionAlignmentCenter
    frmAltaSocios.ssdbAusencias.Columns(2).Alignment = ssCaptionAlignmentCenter
    frmAltaSocios.ssdbAusencias.Columns(3).Alignment = ssCaptionAlignmentRight
    frmAltaSocios.ssdbAusencias.Columns(4).Alignment = ssCaptionAlignmentRight
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmAusencias.LlenaAusencias
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub


'Revisa si hubo cambios a los datos en la forma
Private Function Cambios() As Boolean
    Cambios = True
    
    If (nRegAus <> Val(Me.txtReg.Text)) Then
        Exit Function
    End If
    
    If (sNombre <> Trim(Me.cbNombre.Text)) Then
        Exit Function
    End If
    
    If (nPorcen <> Val(Me.txtPorcen.Text)) Then
        Exit Function
    End If

    If (dInicia <> Format(Me.dtpInicia.Value, "dd/mm/yyyy")) Then
        Exit Function
    End If

    If (dTermina <> Format(Me.dtpTermina.Value, "dd/mm/yyyy")) Then
        Exit Function
    End If

    Cambios = False
End Function


Private Function TraslapaFechas(IdUser As Integer, dInicia As Date, dFinal As Date) As Boolean
Dim rsVerAusencias As ADODB.Recordset
Dim sCampos As String
Dim bTraslapo As Boolean

    bTraslapo = False
    
    sCampos = "FechaInicial, FechaFinal"
    
    InitRecordSet rsVerAusencias, sCampos, "Ausencias", "IdMember=" & IdUser, "FechaInicial", Conn
    With rsVerAusencias
        If (.RecordCount > 0) Then
        
            .MoveFirst
            Do While (Not .EOF)
                If ((dFinal < .Fields("FechaInicial")) Or (dInicia > .Fields("FechaFinal"))) Then
                    .MoveNext
                Else
                    bTraslapo = True
                    Exit Do
                End If
            Loop
        End If
        
        .Close
    End With
    
    Set rsVerAusencias = Nothing

    TraslapaFechas = bTraslapo
End Function
