VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInstructores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructores (Captura)"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9135
   Icon            =   "frmInstructores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9135
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   8145
      Picture         =   "frmInstructores.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Salir"
      Top             =   75
      Width           =   795
   End
   Begin VB.CommandButton cmdGuardar 
      Height          =   840
      Left            =   7155
      Picture         =   "frmInstructores.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Guardar"
      Top             =   75
      Width           =   795
   End
   Begin VB.TextBox txtInstructor 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3675
      TabIndex        =   0
      Top             =   825
      Width           =   1500
   End
   Begin VB.Frame frmPersonales 
      Caption         =   "Datos Personales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4035
      Left            =   120
      TabIndex        =   26
      Top             =   1320
      Width           =   8850
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   11
         Left            =   6840
         TabIndex        =   38
         Top             =   3600
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   10
         Left            =   5280
         TabIndex        =   35
         Top             =   3600
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   9
         Left            =   3840
         TabIndex        =   34
         Top             =   3600
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   8
         Left            =   2400
         TabIndex        =   32
         Top             =   3600
         Width           =   1275
      End
      Begin VB.ComboBox cmbStatus 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3600
         Width           =   1815
      End
      Begin VB.ComboBox cboDeloMuni 
         Height          =   315
         Left            =   4530
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2100
         Width           =   4065
      End
      Begin VB.ComboBox cboEntidad 
         Height          =   315
         Left            =   200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2100
         Width           =   4170
      End
      Begin MSComCtl2.DTPicker dtpAlta 
         Height          =   315
         Left            =   6705
         TabIndex        =   23
         Top             =   2895
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58654721
         CurrentDate     =   38037
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   7
         Left            =   4545
         TabIndex        =   21
         Top             =   2900
         Width           =   2000
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   0
         Left            =   200
         TabIndex        =   3
         Top             =   500
         Width           =   2685
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   1
         Left            =   3060
         TabIndex        =   5
         Top             =   500
         Width           =   2700
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   2
         Left            =   5910
         TabIndex        =   7
         Top             =   500
         Width           =   2685
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   3
         Left            =   200
         TabIndex        =   9
         Top             =   1300
         Width           =   4470
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   4
         Left            =   4830
         TabIndex        =   11
         Top             =   1300
         Width           =   3765
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   5
         Left            =   200
         TabIndex        =   17
         Top             =   2900
         Width           =   2000
      End
      Begin VB.TextBox txtDatos 
         Height          =   300
         Index           =   6
         Left            =   2370
         TabIndex        =   19
         Top             =   2900
         Width           =   2000
      End
      Begin VB.Label lblDatos 
         Caption         =   "Clave Percepcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   6840
         TabIndex        =   39
         Top             =   3360
         Width           =   1650
      End
      Begin VB.Label lblDatos 
         Caption         =   "Num. Empleado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   5280
         TabIndex        =   37
         Top             =   3360
         Width           =   1650
      End
      Begin VB.Label lblDatos 
         Caption         =   "Nómina"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3840
         TabIndex        =   36
         Top             =   3360
         Width           =   930
      End
      Begin VB.Label lblDatos 
         Caption         =   "&Patrón"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   2400
         TabIndex        =   33
         Top             =   3360
         Width           =   930
      End
      Begin VB.Label lblDatos 
         Caption         =   "&Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   31
         Top             =   3360
         Width           =   930
      End
      Begin VB.Label lblDatos 
         Caption         =   "&Alta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   6705
         TabIndex        =   22
         Top             =   2655
         Width           =   930
      End
      Begin VB.Label lblDatos 
         Caption         =   "&R. F. C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   20
         Top             =   2655
         Width           =   930
      End
      Begin VB.Label lblDatos 
         Caption         =   "Apellido &Paterno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   2
         Top             =   255
         Width           =   1560
      End
      Begin VB.Label lblDatos 
         Caption         =   "Apellido &Materno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3060
         TabIndex        =   4
         Top             =   240
         Width           =   1590
      End
      Begin VB.Label lblDatos 
         Caption         =   "&Nombre(s):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5925
         TabIndex        =   6
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblDatos 
         Caption         =   "Ca&lle y Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   195
         TabIndex        =   8
         Top             =   1050
         Width           =   1440
      End
      Begin VB.Label lblDatos 
         Caption         =   "C&olonia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4830
         TabIndex        =   10
         Top             =   1050
         Width           =   1320
      End
      Begin VB.Label lblDatos 
         Caption         =   "&Entidad Federativa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   195
         TabIndex        =   12
         Top             =   1845
         Width           =   1845
      End
      Begin VB.Label lblDatos 
         Caption         =   "&Delegación o Municipio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4530
         TabIndex        =   14
         Top             =   1845
         Width           =   2145
      End
      Begin VB.Label lblDatos 
         Caption         =   "Teléfono &1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   200
         TabIndex        =   16
         Top             =   2650
         Width           =   1305
      End
      Begin VB.Label lblDatos 
         Caption         =   "Teléfono &2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2370
         TabIndex        =   18
         Top             =   2655
         Width           =   930
      End
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INSTRUCTORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   2580
      TabIndex        =   29
      Top             =   165
      Width           =   3975
   End
   Begin VB.Label lblClave 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   3630
      TabIndex        =   28
      Top             =   810
      Width           =   1275
   End
   Begin VB.Label lblInstructor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3900
      TabIndex        =   27
      Top             =   810
      Width           =   1140
   End
   Begin VB.Label lblClaveLetrero 
      BackStyle       =   0  'Transparent
      Caption         =   "&Clave:"
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
      Left            =   3030
      TabIndex        =   1
      Top             =   870
      Width           =   660
   End
End
Attribute VB_Name = "frmInstructores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA CAPTURA DE DATOS DE LOS INSTRUCTORES
' Objetivo: CATÁLOGO DE INSTRUCTORES
' Programado por:
' Fecha: FEBRERO DE 2004
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsInstructores As ADODB.Recordset
    
Private Function VerificaDatos()
    Dim i, intInicio As Integer
        If txtInstructor.Visible = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    If (txtInstructor.Visible = True) And (txtInstructor.Text = "") Then
        MsgBox "¡ Clave del Instructor Inválida !", vbOKOnly + vbExclamation, "instructores (Captura)"
        VerificaDatos = False
        txtInstructor.SetFocus
        Exit Function
    End If
    For i = intInicio To 7
        Select Case i
            Case 0
                If (txtDatos(i).Text = "") Then
                    MsgBox "¡ Digite el APELLIDO PATERNO, Pues No es Opcional !", _
                                 vbOKOnly + vbExclamation, "instructores (Captura)"
                    VerificaDatos = False
                    txtDatos(i).SetFocus
                    Exit Function
                End If
            Case 2
                If txtDatos(i).Text = "" Then
                    MsgBox "¡ Digite el o los NOMBRES, Pues el Dato No es Opcional !", _
                                vbOKOnly + vbExclamation, "instructores (Captura)"
                    VerificaDatos = False
                    txtDatos(i).SetFocus
                    Exit Function
                End If
            Case 7
                If txtDatos(i).Text = "" Then
                    MsgBox "¡ Digite el R. F. C., Pues No es Opcional.", _
                                vbOKOnly + vbExclamation, "instructores (Captura)"
                    VerificaDatos = False
                    txtDatos(i).SetFocus
                    Exit Function
                End If
        End Select
        VerificaDatos = True
    Next

    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT idinstructor FROM instructores WHERE idinstructor = " & _
                        Val(Trim(txtInstructor.Text))
        Set AdoRcsInstructores = New ADODB.Recordset
        AdoRcsInstructores.ActiveConnection = Conn
        AdoRcsInstructores.LockType = adLockOptimistic
        AdoRcsInstructores.CursorType = adOpenKeyset
        AdoRcsInstructores.CursorLocation = adUseServer
        AdoRcsInstructores.Open strSQL
        If Not AdoRcsInstructores.EOF Then
            MsgBox "Ya Existe Un Registro Con La Clave: " & _
                            txtInstructor.Text, vbInformation + vbOKOnly, "Instructores"
            AdoRcsInstructores.Close
            VerificaDatos = False
            txtInstructor.SetFocus
            Exit Function
        Else
            AdoRcsInstructores.Close
            VerificaDatos = True
        End If
    End If
End Function

Private Sub cboDeloMuni_LostFocus()
    If (cboDeloMuni.Text <> "") And (cboDeloMuni.ListIndex < 0) Then
        MsgBox "Seleccione una DELEGACIÓN o MUNICIPIO de la Lista."
        cboDeloMuni.SetFocus
    End If
End Sub

Private Sub cboEntidad_Click()
    Dim lngCveDeloMuni As Long
    Dim strCampo1, strCampo2 As String
    If cboEntidad.ListIndex < 0 Then Exit Sub
    'Llena combo Delegación o Municipio
    lngCveDeloMuni = cboEntidad.ItemData(cboEntidad.ListIndex)
    strSQL = "SELECT cvedelomuni, nomdelomuni FROM delgamunici " & _
                    "WHERE entidadfed = " & lngCveDeloMuni & " ORDER BY nomdelomuni"
    strCampo1 = "nomdelomuni"
    strCampo2 = "cvedelomuni"
    Call LlenaCombos(cboDeloMuni, strSQL, strCampo1, strCampo2)
End Sub

Private Sub cboEntidad_LostFocus()
    If (cboEntidad.Text <> "") And (cboEntidad.ListIndex < 0) Then
        cboDeloMuni.Text = ""
        MsgBox "Seleccione una ENTIDAD FEDERATIVA de la Lista."
        cboEntidad.SetFocus
    End If
End Sub

Private Sub cmdGuardar_Click()
    If VerificaDatos = True Then
        If frmCatalogos.lblModo.Caption = "A" Then
            Call GuardaDatos
        Else
            Call RemplazaDatos
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim strCampo1, strCampo2 As String
    frmCatalogos.Enabled = False
    'Llena los combos de Entidad Federativa
    strSQL = "SELECT cveentfederativa, nomentfederativa FROM entfederativa"
    strCampo1 = "nomentfederativa"
    strCampo2 = "cveentfederativa"
    LlenaCombos cboEntidad, strSQL, strCampo1, strCampo2
    
    
    
    Me.cmbStatus.Clear
    Me.cmbStatus.AddItem "Alta"
    Me.cmbStatus.AddItem "Baja"
    
    Me.cmbStatus.Text = "Alta"
    
    dtpAlta.Value = Now()
    If frmCatalogos.lblModo.Caption = "A" Then
        txtInstructor.Visible = True
        Llena_txtInstructor
    Else
        txtInstructor.Visible = False
        LlenaDatos
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        frmCatalogos.Enabled = True
End Sub

Private Sub GuardaDatos()
    Dim AdoCmdInserta As ADODB.Command
    On Error GoTo err_Guarda
    iniTrans = Conn.BeginTrans              'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO Instructores ("
        strSQL = strSQL & " idinstructor,"
        strSQL = strSQL & " apellido_paterno,"
        strSQL = strSQL & " apellido_materno, "
        strSQL = strSQL & " nombre,"
        strSQL = strSQL & " alta,"
        strSQL = strSQL & " calle,"
        strSQL = strSQL & " colonia,"
        strSQL = strSQL & " ent_federativa,"
        strSQL = strSQL & " delegamunici,"
        strSQL = strSQL & " telefono_1,"
        strSQL = strSQL & " telefono_2,"
        strSQL = strSQL & " rfc,"
        strSQL = strSQL & " Status,"
        strSQL = strSQL & " Empresa,"
        strSQL = strSQL & " Nomina,"
        strSQL = strSQL & " NoEmpleado,"
        strSQL = strSQL & " ClavePercepcion)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & Val(Trim(txtInstructor.Text)) & ","
        strSQL = strSQL & "'" & Trim(txtDatos(0).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(1).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(2).Text) & "',"
        strSQL = strSQL & "'" & Format(dtpAlta.Value, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(3).Text) & "',"
        strSQL = strSQL & "'" & Trim(Trim(txtDatos(4).Text)) & "',"
        strSQL = strSQL & "'" & Trim(cboEntidad.Text) & "',"
        strSQL = strSQL & "'" & Trim(cboDeloMuni.Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(5).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(6).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(7).Text) & "',"
        strSQL = strSQL & "'" & Left(Me.cmbStatus.Text, 1) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(8).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(9).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(10).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(11).Text) & "')"
    #Else
        strSQL = "INSERT INTO Instructores ("
        strSQL = strSQL & " idinstructor,"
        strSQL = strSQL & " apellido_paterno,"
        strSQL = strSQL & " apellido_materno, "
        strSQL = strSQL & " nombre,"
        strSQL = strSQL & " alta,"
        strSQL = strSQL & " calle,"
        strSQL = strSQL & " colonia,"
        strSQL = strSQL & " ent_federativa,"
        strSQL = strSQL & " delegamunici,"
        strSQL = strSQL & " telefono_1,"
        strSQL = strSQL & " telefono_2,"
        strSQL = strSQL & " rfc,"
        strSQL = strSQL & " Status,"
        strSQL = strSQL & " Empresa,"
        strSQL = strSQL & " Nomina,"
        strSQL = strSQL & " NoEmpleado,"
        strSQL = strSQL & " ClavePercepcion)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & Val(Trim(txtInstructor.Text)) & ","
        strSQL = strSQL & "'" & Trim(txtDatos(0).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(1).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(2).Text) & "',"
        strSQL = strSQL & "'" & Format(dtpAlta.Value, "dd/mm/yyyy") & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(3).Text) & "',"
        strSQL = strSQL & "'" & Trim(Trim(txtDatos(4).Text)) & "',"
        strSQL = strSQL & "'" & Trim(cboEntidad.Text) & "',"
        strSQL = strSQL & "'" & Trim(cboDeloMuni.Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(5).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(6).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(7).Text) & "',"
        strSQL = strSQL & "'" & Left(Me.cmbStatus.Text, 1) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(8).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(9).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(10).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(11).Text) & "')"
    #End If
    
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    Screen.MousePointer = vbDefault
    Conn.CommitTrans                        'Termina transacción
    MsgBox "¡Registro Ingresado!"
    Call Limpia
    Exit Sub

err_Guarda:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub RemplazaDatos()
   Dim AdoCmdRemplaza As ADODB.Command
    On Error GoTo err_Guarda
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    'Eliminamos elregistro existente
    strSQL = "DELETE FROM instructores WHERE idinstructor = " & Val(lblClave.Caption)
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute

    'Ahora Insertamos el nuevo registro que sustituye al anterior
    #If SqlServer_ Then
        strSQL = "INSERT INTO Instructores ("
        strSQL = strSQL & " idinstructor,"
        strSQL = strSQL & " apellido_paterno,"
        strSQL = strSQL & " apellido_materno, "
        strSQL = strSQL & " nombre,"
        strSQL = strSQL & " alta,"
        strSQL = strSQL & " calle,"
        strSQL = strSQL & " colonia,"
        strSQL = strSQL & " ent_federativa,"
        strSQL = strSQL & " delegamunici,"
        strSQL = strSQL & " telefono_1,"
        strSQL = strSQL & " telefono_2,"
        strSQL = strSQL & " rfc,"
        strSQL = strSQL & " Status)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & Val(lblClave.Caption) & ","
        strSQL = strSQL & "'" & Trim(txtDatos(0).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(1).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(2).Text) & "',"
        strSQL = strSQL & "'" & Format(dtpAlta.Value, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(3).Text) & "',"
        strSQL = strSQL & "'" & Trim(Trim(txtDatos(4).Text)) & "',"
        strSQL = strSQL & "'" & Trim(cboEntidad.Text) & "',"
        strSQL = strSQL & "'" & Trim(cboDeloMuni.Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(5).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(6).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(7).Text) & "',"
        strSQL = strSQL & "'" & Left(Me.cmbStatus.Text, 1) & "')"
    #Else
        strSQL = "INSERT INTO Instructores ("
        strSQL = strSQL & " idinstructor,"
        strSQL = strSQL & " apellido_paterno,"
        strSQL = strSQL & " apellido_materno, "
        strSQL = strSQL & " nombre,"
        strSQL = strSQL & " alta,"
        strSQL = strSQL & " calle,"
        strSQL = strSQL & " colonia,"
        strSQL = strSQL & " ent_federativa,"
        strSQL = strSQL & " delegamunici,"
        strSQL = strSQL & " telefono_1,"
        strSQL = strSQL & " telefono_2,"
        strSQL = strSQL & " rfc,"
        strSQL = strSQL & " Status)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & Val(lblClave.Caption) & ","
        strSQL = strSQL & "'" & Trim(txtDatos(0).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(1).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(2).Text) & "',"
        strSQL = strSQL & "#" & Format(dtpAlta.Value, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "'" & Trim(txtDatos(3).Text) & "',"
        strSQL = strSQL & "'" & Trim(Trim(txtDatos(4).Text)) & "',"
        strSQL = strSQL & "'" & Trim(cboEntidad.Text) & "',"
        strSQL = strSQL & "'" & Trim(cboDeloMuni.Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(5).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(6).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtDatos(7).Text) & "',"
        strSQL = strSQL & "'" & Left(Me.cmbStatus.Text, 1) & "')"
    #End If
    
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    Conn.CommitTrans      'Termina transacción
    Screen.MousePointer = vbDefault
    MsgBox "¡Registro Actualizado!"
    Unload Me
    Exit Sub

err_Guarda:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub









Private Sub txtInstructor_GotFocus()
    txtInstructor.SelStart = 0
    txtInstructor.SelLength = Len(txtInstructor)
End Sub

Private Sub txtDatos_KeyPress(index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub Limpia()
    Dim i As Integer
    For i = 0 To 7
        txtDatos(i).Text = ""
    Next i
    cboEntidad.ListIndex = -1
    cboDeloMuni.ListIndex = -1
    Llena_txtInstructor
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM instructores WHERE idinstructor = " & _
                  Val(Trim(frmCatalogos.lblModo.Caption))
    Set AdoRcsInstructores = New ADODB.Recordset
    AdoRcsInstructores.ActiveConnection = Conn
    AdoRcsInstructores.LockType = adLockOptimistic
    AdoRcsInstructores.CursorType = adOpenKeyset
    AdoRcsInstructores.CursorLocation = adUseServer
    AdoRcsInstructores.Open strSQL
    If Not AdoRcsInstructores.EOF Then
        lblClave.Caption = AdoRcsInstructores!IdInstructor
        txtDatos(0).Text = AdoRcsInstructores!apellido_paterno
        txtDatos(1).Text = IIf(IsNull(AdoRcsInstructores!apellido_materno), "", AdoRcsInstructores!apellido_materno)
        txtDatos(2).Text = AdoRcsInstructores!Nombre
        txtDatos(3).Text = IIf(IsNull(AdoRcsInstructores!calle), "", AdoRcsInstructores!calle)
        txtDatos(4).Text = IIf(IsNull(AdoRcsInstructores!colonia), "", AdoRcsInstructores!colonia)
        If Not IsNull(AdoRcsInstructores!ent_federativa) Then
            If Len(AdoRcsInstructores!ent_federativa) > 0 Then
                cboEntidad.Text = IIf(IsNull(AdoRcsInstructores!ent_federativa), "", AdoRcsInstructores!ent_federativa)
            End If
        End If
        If Not IsNull(AdoRcsInstructores!delegamunici) Then
            If Len(AdoRcsInstructores!delegamunici) > 0 Then
                cboDeloMuni.Text = IIf(IsNull(AdoRcsInstructores!delegamunici), "", AdoRcsInstructores!delegamunici)
            End If
        End If
        txtDatos(5).Text = IIf(IsNull(AdoRcsInstructores!telefono_1), "", AdoRcsInstructores!telefono_1)
        txtDatos(6).Text = IIf(IsNull(AdoRcsInstructores!telefono_2), "", AdoRcsInstructores!telefono_2)
        txtDatos(7).Text = AdoRcsInstructores!rfc
        dtpAlta.Value = AdoRcsInstructores!alta
        Me.cmbStatus.Text = IIf(AdoRcsInstructores!Status = "A", "Alta", "Baja")
    End If
End Sub

Sub Llena_txtInstructor()
    Dim lngAnterior, lngInstructores As Long
    'Llena txtinstructores
    lngInstructores = 1
    strSQL = "SELECT idinstructor FROM instructores ORDER BY idinstructor"
    Set AdoRcsInstructores = New ADODB.Recordset
    AdoRcsInstructores.ActiveConnection = Conn
    AdoRcsInstructores.LockType = adLockOptimistic
    AdoRcsInstructores.CursorType = adOpenKeyset
    AdoRcsInstructores.CursorLocation = adUseServer
    AdoRcsInstructores.Open strSQL
    If AdoRcsInstructores.EOF Then
        lngInstructores = 1
        txtInstructor.Text = lngInstructores
        Exit Sub
    End If
    AdoRcsInstructores.MoveFirst
    Do While Not AdoRcsInstructores.EOF
        If AdoRcsInstructores.Fields!IdInstructor <> "1" Then
            If Val(AdoRcsInstructores.Fields!IdInstructor) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngInstructores
        AdoRcsInstructores.MoveNext
        If Not AdoRcsInstructores.EOF Then lngInstructores = AdoRcsInstructores.Fields!IdInstructor
    Loop
    txtInstructor.Text = lngAnterior + 1
End Sub
