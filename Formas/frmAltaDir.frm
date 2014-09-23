VERSION 5.00
Begin VB.Form frmAltaDir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de direcciones"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   Icon            =   "frmAltaDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDir 
      Height          =   8055
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6600
      Begin VB.Frame fraDatosEmpresa 
         Height          =   2175
         Left            =   120
         TabIndex        =   35
         Top             =   4920
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox txtArea 
            Height          =   285
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtEmpresa 
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   6015
         End
         Begin VB.Label lblArea 
            Caption         =   "Area"
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   1200
            Width           =   495
         End
         Begin VB.Label lblEmpresa 
            Caption         =   "Empresa"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos Fiscales"
         Height          =   2175
         Left            =   120
         TabIndex        =   32
         Top             =   4920
         Visible         =   0   'False
         Width           =   6255
         Begin VB.TextBox txtRfc 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   13
            Top             =   1680
            Width           =   2535
         End
         Begin VB.TextBox txtRazonSocial 
            Height          =   285
            Left            =   120
            MaxLength       =   70
            TabIndex        =   12
            Top             =   960
            Width           =   6015
         End
         Begin VB.OptionButton optTipoPersona 
            Caption         =   "Persona Moral"
            Height          =   375
            Index           =   1
            Left            =   3840
            TabIndex        =   11
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton optTipoPersona 
            Caption         =   "Persona Física"
            Height          =   375
            Index           =   0
            Left            =   1440
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.Label lblRfc 
            Caption         =   "RFC"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lblRazonSocial 
            Caption         =   "Razón social o persona física"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   2655
         End
      End
      Begin VB.CommandButton cmdGuardarComo 
         Height          =   615
         Left            =   3840
         Picture         =   "frmAltaDir.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "  Guardar como...  "
         Top             =   7320
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtDeloMuni 
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   3720
         Width           =   3735
      End
      Begin VB.TextBox txtCiudad 
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   6135
      End
      Begin VB.CommandButton cmdAyuda 
         Height          =   305
         Left            =   960
         Picture         =   "frmAltaDir.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   " Tipos de dirección disponible "
         Top             =   480
         Width           =   425
      End
      Begin VB.TextBox txtTipoDir 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1520
         TabIndex        =   17
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtCveDir 
         Height          =   285
         Left            =   240
         MaxLength       =   4
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtNoReg 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   7440
         Width           =   615
      End
      Begin VB.ComboBox cbEdos 
         Height          =   315
         Left            =   4080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3720
         Width           =   2295
      End
      Begin VB.CommandButton cmdGuardar 
         Height          =   615
         Left            =   4920
         Picture         =   "frmAltaDir.frx":0896
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   " Guardar registro "
         Top             =   7320
         Width           =   615
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   615
         Left            =   5760
         Picture         =   "frmAltaDir.frx":0CD8
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   " Salir "
         Top             =   7320
         Width           =   615
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   4320
         TabIndex        =   14
         Top             =   4440
         Width           =   2055
      End
      Begin VB.TextBox txtCodPos 
         Height          =   285
         Left            =   5280
         TabIndex        =   4
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtColonia 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox txtCalle 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   6135
      End
      Begin VB.TextBox txtTel1 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox txtTel2 
         Height          =   285
         Left            =   2280
         TabIndex        =   9
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label lblCiudad 
         Caption         =   "Ciudad"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lblCveDir 
         Caption         =   "Cve. Dir."
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblNoReg 
         Caption         =   "# Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   7200
         Width           =   495
      End
      Begin VB.Label lblTipoDir 
         Caption         =   "Tipo de dirección"
         Height          =   255
         Left            =   1515
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax"
         Height          =   255
         Left            =   4320
         TabIndex        =   26
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblEdo 
         Caption         =   "Estado"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4080
         TabIndex        =   25
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblCp 
         Caption         =   "Código postal"
         Height          =   255
         Left            =   5280
         TabIndex        =   24
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblCalle 
         Caption         =   "Calle"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblColonia 
         Caption         =   "Colonia"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblDeloMuni 
         Caption         =   "Delegación o municipio"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblTel1 
         Caption         =   "Teléfonos"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   4200
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAltaDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de direcciones                  *
'*  Daniel Hdez                                                 *
'*  26 / Agosto / 2004                                          *
'*  Ultima actualización: 07 / Noviembre / 2005                 *
'****************************************************************
Public bNvaDir As Boolean

Dim sTextToolBar As String
Dim sCalle As String
Dim sCol As String
Dim nCp As Long
Dim sDeloMuni As String
Dim sEdo As String
Dim sTel1 As String
Dim sTel2 As String
Dim sFax As String
Dim sTipoDir As String
Dim nTipoDir As String
Dim sRazonSocial As String
Dim sRfc As String
Dim sCiudad As String
Dim sEmpresa As String
Dim sArea As String
Dim frmHDir As frmayuda

''Busca el Edo de la Rep al cual pertenece la delegacion o el municipio
'Private Sub cbDeloMuni_LostFocus()
'Dim nCveEdo As Integer
'
'    If (Trim(Me.cbDeloMuni.Text) <> "") Then
'        nCveEdo = LeeXValor("EntidadFed", "DelgaMunici", "NomDeloMuni='" & Trim(Me.cbDeloMuni.Text) & "'", "EntidadFed", "n", Conn)
'
'        If (nCveEdo = 0) Then
'            MsgBox "La delegación o municipio seleccionado no existe, verifique los datos.", vbExclamation, "KalaSystems"
'            Me.cbDeloMuni.SetFocus
'        Else
'            Me.cbEdos.Text = LeeXValor("NomEntFederativa", "EntFederativa", "CveEntFederativa=" & nCveEdo, "NomEntFederativa", "s", Conn)
'        End If
'    Else
'        Me.cbDeloMuni.Enabled = False
'    End If
'End Sub

Private Sub cbEdos_LostFocus()
Dim nCveEdo As Integer

    If (Me.cbEdos.Text <> "") Then
        nCveEdo = LeeXValor("CveEntFederativa", "EntFederativa", "NomEntFederativa='" & Trim(Me.cbEdos.Text) & "'", "CveEntFederativa", "n", Conn)
    
        If (nCveEdo > 0) Then
            sEdo = Trim$(Me.cbEdos.Text)
        End If
    End If
End Sub

Private Sub cmdGuardarComo_Click()
    If Not ChecaSeguridad(Me.Name, Me.cmdGuardarComo.Name) Then
        Exit Sub
    End If
    
    bNvaDir = True
    
    Call cmdSalir_Click
End Sub

Private Sub cmdSalir_Click()
    Dim Respuesta As Integer

    If (Cambios) Then
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Registro de direcciones")
        
        If (Respuesta = vbYes) Then
        
            If Not ChecaSeguridad(Me.Name, Me.cmdGuardar.Name) Then
                Exit Sub
            End If

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
    Dim sSql As String

    sTextToolBar = Trim(MDIPrincipal.StatusBar1.Panels.Item(1).Text)
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Registro de direciones"
    
    'Llena el combo con la lista de los Estados de la Republica
    sSql = "SELECT CveEntFederativa, NomEntFederativa FROM Entfederativa"
    LlenaCombos Me.cbEdos, sSql, "NomEntFederativa", "CveEntFederativa"
    
    If (bNvaDir) Then
        ClearCtrls
        Me.cmdGuardarComo.Enabled = False
    Else
        LeeDatos
    End If
    
    InitVar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAltaDir.LlenaDirs
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub

Private Sub ClearCtrls()
    With Me
        .txtCalle.Text = ""
        .txtColonia.Text = ""
        .txtCodPos.Text = ""
        .txtDeloMuni.Text = ""
        .cbEdos.ListIndex = -1
        .txtTel1.Text = ""
        .txtTel2.Text = ""
        .txtFax.Text = ""
        .txtCveDir.Text = ""
        .txtTipoDir.Text = ""
        .txtDeloMuni.Text = ""
        .txtRazonSocial.Text = ""
        .txtRfc.Text = ""
        .txtCiudad.Text = ""
        .txtEmpresa.Text = ""
        .txtArea.Text = ""
    End With
End Sub

Private Sub InitVar()
    sCalle = Trim$(Me.txtCalle.Text)
    sCol = Trim$(Me.txtColonia.Text)
    nCp = Val(Me.txtCodPos.Text)
    sDeloMuni = Trim$(Me.txtDeloMuni.Text)
    sTel1 = Trim$(Me.txtTel1.Text)
    sTel2 = Trim$(Me.txtTel2.Text)
    sFax = Trim$(Me.txtFax.Text)
    nTipoDir = Val(Me.txtCveDir.Text)
    sTipoDir = Trim$(Me.txtCveDir.Text)
    sRazonSocial = Trim$(Me.txtRazonSocial.Text)
    sRfc = Trim$(Me.txtRfc.Text)
    sCiudad = Trim$(Me.txtCiudad.Text)
    sEdo = Trim$(Me.cbEdos.Text)
    sEmpresa = Trim$(Me.txtEmpresa.Text)
    sArea = Trim$(Me.txtArea.Text)
End Sub

Private Sub cmdGuardar_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdGuardar.Name) Then
        Exit Sub
    End If
    
    If (Cambios) Then
        If (GuardaDatos) Then
            'Inicializa las variables
            InitVar
        Else
            MsgBox "No se registraron los datos, verifique la información.", vbCritical, "KalaSystems"
        End If
    End If
    
    Me.txtCalle.SetFocus
End Sub

Private Function ChecaDatos()
Dim sCond As String
Dim sCamp As String

    ChecaDatos = False
    
    If Val(Me.txtCveDir.Text) <> 2 Then
        If (Trim(Me.txtCalle.Text) = "") Then
            MsgBox "La calle no puede quedar en blanco.", vbExclamation, "KalaSystems"
            Me.txtCalle.SetFocus
            Exit Function
        End If
        
        If (Trim(Me.txtCiudad.Text) = "") Then
            MsgBox "La ciudad no puede quedar en blanco.", vbExclamation, "KalaSystems"
            Me.txtCiudad.SetFocus
            Exit Function
        End If
        
        If (Trim(Me.txtDeloMuni.Text) = "") Then
            MsgBox "La delegación o municipio no puede quedar en blanco.", vbExclamation, "KalaSystems"
            Me.txtDeloMuni.SetFocus
            Exit Function
        End If
        
    '    If (Trim(Me.cbDeloMuni.Text) = "") Then
    '        MsgBox "Se debe seleccionar una delegación o municipio.", vbExclamation, "KalaSystems"
    '        Me.cbDeloMuni.SetFocus
    '        Exit Function
    '    Else
    '        nDeloMuni = LeeXValor("CveDeloMuni", "DelgaMunici", "NomDeloMuni='" & Trim(Me.cbDeloMuni.Text) & "'", "CveDeloMuni", "n", Conn)
    '
    '        If (nDeloMuni <= 0) Then
    '            MsgBox "La delegación o municipio seleccionado es incorrecto.", vbExclamation, "KalaSystems"
    '            Me.cbDeloMuni.SetFocus
    '            Exit Function
    '        End If
    '    End If
        
        If (Trim(Me.txtTipoDir.Text) = "") Then
            MsgBox "Se debe seleccionar un tipo de domicilio.", vbExclamation, "KalaSystems"
            Me.txtCveDir.SetFocus
            Exit Function
    '    Else
    '        nTipoDir = LeeXValor("IdTipoDireccion", "Tipo_Direccion", "Descripcion='" & Trim(Me.cbTipoDir.Text) & "'", "IdTipoDireccion", "n", Conn)
    '
    '        If (nTipoDir <= 0) Then
    '            MsgBox "El tipo de dirección seleccionado es incorrecto.", vbExclamation, "KalaSystems"
    '            Me.cbTipoDir.SetFocus
    '            Exit Function
    '        End If
        End If
    
        If Val(Me.txtCveDir.Text) = 3 Then
            If Me.txtRazonSocial.Text = vbNullString Then
                MsgBox "Falta razón social", vbExclamation, "Verifique"
                Me.txtRazonSocial.SetFocus
                Exit Function
            End If
            
            Me.txtRfc.Text = Trim(Me.txtRfc.Text)
            
            If Len(Me.txtRfc.Text) <> 13 And Me.optTipoPersona(0).Value Then
                MsgBox "El RFC debe ser de 13 caracteres", vbExclamation, "Verifique"
                Me.txtRfc.SetFocus
                Exit Function
            End If
            If Len(Me.txtRfc.Text) <> 12 And Not Me.optTipoPersona(0).Value Then
                MsgBox "El RFC debe ser de 12 caracteres", vbExclamation, "Verifique"
                Me.txtRfc.SetFocus
                Exit Function
            End If
            
            If InStr(Me.txtRfc.Text, " ") > 0 Then
                MsgBox "El RFC no es válido. Elimine espacios", vbExclamation, "Verifique"
                Me.txtRfc.SetFocus
                Exit Function
            End If
            
            If Not RfcValido(Me.txtRfc.Text) Then
                MsgBox "El RFC no es válido. Elimine caracteres especiales y/o signos de puntuación.", vbExclamation, "Verifique"
                Me.txtRfc.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If (Trim(Me.txtTel1.Text) = "" And Trim(Me.txtTel2.Text) = "") Then
        MsgBox "El campo teléfono no puede quedar en blanco.", vbExclamation, "KalaSystems"
        Me.txtTel1.SetFocus
        Exit Function
    End If
    
    If Val(Me.txtCveDir.Text) = 2 Then
        If Me.txtEmpresa.Text = vbNullString Then
            MsgBox "Falta empresa", vbExclamation, "Verifique"
            Me.txtEmpresa.SetFocus
            Exit Function
        End If
        If Me.txtArea.Text = vbNullString Then
            MsgBox "Falta área", vbExclamation, "Verifique"
            Me.txtArea.SetFocus
            Exit Function
        End If
    End If
    
    ChecaDatos = True
End Function

Private Function GuardaDatos() As Boolean
    Const DATOSDIRECCION = 16
    Dim bCreado As Boolean
    Dim mFieldsDir(DATOSDIRECCION) As String
    Dim mValuesDir(DATOSDIRECCION) As Variant

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Usuarios_Club
    mFieldsDir(0) = "idDireccion"
    mFieldsDir(1) = "IdMember"
    mFieldsDir(2) = "Calle"
    mFieldsDir(3) = "Colonia"
'    mFieldsDir(4) = "CveDeloMuni"
    mFieldsDir(4) = "DeloMuni"
    mFieldsDir(5) = "CodPos"
    mFieldsDir(6) = "Tel1"
    mFieldsDir(7) = "Tel2"
    mFieldsDir(8) = "Fax"
    mFieldsDir(9) = "IdTipoDireccion"
    mFieldsDir(10) = "RazonSocial"
    mFieldsDir(11) = "Rfc"
    mFieldsDir(12) = "Ciudad"
    mFieldsDir(13) = "Estado"
    mFieldsDir(14) = "TipoPersona"
    mFieldsDir(15) = "Area"

    If (bNvaDir) Then
        mValuesDir(0) = LeeUltReg("Direcciones", "idDireccion") + 1
    Else
        mValuesDir(0) = Val(Me.txtNoReg.Text)
    End If
    
    mValuesDir(1) = Val(frmAltaSocios.txtTitCve.Text)
    mValuesDir(2) = Trim$(UCase$(Me.txtCalle.Text))
    mValuesDir(3) = Trim$(UCase$(Me.txtColonia.Text))
'    mValuesDir(4) = nDeloMuni
    mValuesDir(4) = Trim$(UCase$(Me.txtDeloMuni.Text))
    mValuesDir(5) = Val(Me.txtCodPos.Text)
    mValuesDir(6) = Trim$(Me.txtTel1.Text)
    mValuesDir(7) = Trim$(Me.txtTel2.Text)
    mValuesDir(8) = Trim$(Me.txtFax.Text)
    mValuesDir(9) = Val(Me.txtCveDir.Text)
    
    If Val(Me.txtCveDir.Text) = 3 Then
        mValuesDir(10) = Trim$(UCase$(Me.txtRazonSocial.Text))
    Else
        mValuesDir(10) = Trim$(UCase$(Me.txtEmpresa.Text))
    End If
    
    mValuesDir(11) = Trim$(UCase$(Me.txtRfc.Text))
    mValuesDir(12) = Trim$(UCase$(Me.txtCiudad.Text))
    mValuesDir(13) = Trim$(UCase$(Me.cbEdos.Text))
    
    If Val(Me.txtCveDir.Text) = 3 Then
        mValuesDir(14) = IIf(Me.optTipoPersona(0).Value, "F", "M")
    Else
        mValuesDir(14) = "F"
    End If
    
    mValuesDir(15) = Trim$(UCase$(Me.txtArea.Text))
    
    If (bNvaDir) Then
        'Registra los datos de la nueva direccion
        If (AgregaRegistro("Direcciones", mFieldsDir, DATOSDIRECCION, mValuesDir, Conn)) Then
            MsgBox "Los datos se dieron de alta correctamente.", vbInformation, "KalaSystems"
            
            'Muestra el numero del registro de la direccion
            Me.txtNoReg.Text = mValuesDir(0)
            Me.txtNoReg.REFRESH

            bNvaDir = False
            GuardaDatos = True
        Else
            MsgBox "El registro no fue completado.", vbCritical, "KalaSystems"
        End If
    Else

        If (Val(Me.txtNoReg.Text) > 0) Then

            'Actualiza los datos del titular
            If (CambiaReg("Direcciones", mFieldsDir, DATOSDIRECCION, mValuesDir, "IdDireccion=" & Val(Me.txtNoReg.Text), Conn)) Then
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
    Cambios = True

    If (sCalle <> Trim$(Me.txtCalle.Text)) Then
        Exit Function
    End If

    If (sCol <> Trim$(Me.txtColonia.Text)) Then
        Exit Function
    End If

    If (nCp <> Val(Me.txtCodPos.Text)) Then
        Exit Function
    End If

    If (sDeloMuni <> Trim$(Me.txtDeloMuni.Text)) Then
        Exit Function
    End If
    
    If (sEdo <> Trim$(Me.cbEdos.Text)) Then
        Exit Function
    End If

    If (sTel1 <> Trim$(Me.txtTel1.Text)) Then
        Exit Function
    End If

    If (sTel2 <> Trim$(Me.txtTel2.Text)) Then
        Exit Function
    End If

    If (sFax <> Trim$(Me.txtFax.Text)) Then
        Exit Function
    End If

    If (sTipoDir <> Trim$(Me.txtCveDir.Text)) Then
        Exit Function
    End If
    
    If (sRazonSocial <> Trim$(Me.txtRazonSocial.Text)) Then
        Exit Function
    End If
    
    If (sRfc <> Trim$(Me.txtRfc.Text)) Then
        Exit Function
    End If
    
    If (sCiudad <> Trim$(Me.txtCiudad.Text)) Then
        Exit Function
    End If
    
    If (sEmpresa <> Trim$(Me.txtEmpresa.Text)) Then
        Exit Function
    End If
    
    If (sArea <> Trim$(Me.txtArea.Text)) Then
        Exit Function
    End If

    Cambios = False
End Function

Private Sub CambiaAMayusculas()
    With Me
        .txtCalle.Text = UCase$(.txtCalle.Text)
        .txtCalle.REFRESH

        .txtColonia.Text = UCase$(.txtColonia.Text)
        .txtColonia.REFRESH

        .txtDeloMuni.Text = UCase$(.txtDeloMuni.Text)
        .txtDeloMuni.REFRESH
        
        .txtRazonSocial.Text = UCase$(.txtRazonSocial.Text)
        .txtRazonSocial.REFRESH
        
        .txtRfc.Text = UCase$(.txtRfc.Text)
        .txtRfc.REFRESH
        
        .txtCiudad.Text = UCase$(.txtCiudad.Text)
        .txtCiudad.REFRESH
        
        .txtEmpresa.Text = UCase$(.txtEmpresa.Text)
        .txtEmpresa.REFRESH
        
        .txtArea.Text = UCase$(.txtArea.Text)
        .txtArea.REFRESH
    End With
End Sub

Public Sub LlenaDirs()
    Const DATOSDIR = 15
    Dim rsDirs As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String
    Dim mAncDir(DATOSDIR) As Integer
    Dim mEncDir(DATOSDIR) As String

    frmAltaSocios.ssdbDireccion.RemoveAll

    sCampos = "Descripcion, Calle, "
    sCampos = sCampos & "Colonia, Ciudad, CodPos, "
    sCampos = sCampos & "Tel1, Tel2, Fax, "
    sCampos = sCampos & "DeloMuni, Estado, "
    sCampos = sCampos & "idDireccion, Direcciones.IdTipoDireccion, "
    sCampos = sCampos & "RazonSocial, Rfc, Area "
    
    sTablas = "Direcciones LEFT JOIN Tipo_Direccion ON Direcciones.IdTipoDireccion=Tipo_Direccion.IdTipoDireccion "
    
    InitRecordSet rsDirs, sCampos, sTablas, "Direcciones.idMember=" & Val(frmAltaSocios.txtTitCve.Text), "Descripcion", Conn
    With rsDirs
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                frmAltaSocios.ssdbDireccion.AddItem .Fields("Descripcion") & vbTab & _
                .Fields("Calle") & vbTab & _
                .Fields("Colonia") & vbTab & _
                .Fields("Ciudad") & vbTab & _
                .Fields("CodPos") & vbTab & _
                .Fields("Tel1") & vbTab & _
                .Fields("Tel2") & vbTab & _
                .Fields("Fax") & vbTab & _
                .Fields("DeloMuni") & vbTab & _
                .Fields("Estado") & vbTab & _
                .Fields("idDireccion") & vbTab & _
                .Fields("idTipoDireccion") & vbTab & _
                .Fields("RazonSocial") & vbTab & _
                .Fields("Rfc") & vbTab & _
                .Fields("Area")
                .MoveNext
            Loop
        End If
    
        .Close
    End With
    Set rsDirs = Nothing
    
    'Asigna valores a la matriz de encabezados
    mEncDir(0) = "Tipo Dir."
    mEncDir(1) = "Calle"
    mEncDir(2) = "Colonia"
    mEncDir(3) = "Ciudad"
    mEncDir(4) = "CP"
    mEncDir(5) = "Tel 1"
    mEncDir(6) = "Tel 2"
    mEncDir(7) = "Fax"
    mEncDir(8) = "Del. o municipio"
    mEncDir(9) = "Ent. Federativa"
    mEncDir(10) = "# Reg."
    mEncDir(11) = "# Dir."
    mEncDir(12) = "Razón social o persona física"
    mEncDir(13) = "RFC"
    mEncDir(14) = "Area"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid frmAltaSocios.ssdbDireccion, mEncDir
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncDir(0) = 2500
    mAncDir(1) = 5500
    mAncDir(2) = 4000
    mAncDir(3) = 3500
    mAncDir(4) = 1100
    mAncDir(5) = 1800
    mAncDir(6) = 1800
    mAncDir(7) = 1800
    mAncDir(8) = 3500
    mAncDir(9) = 3500
    mAncDir(10) = 800
    mAncDir(11) = 800
    mAncDir(12) = 3500
    mAncDir(13) = 2500
    mAncDir(14) = 2500

    'Asigna el ancho de cada columna
    DefAnchossGrid frmAltaSocios.ssdbDireccion, mAncDir
End Sub


Private Sub LeeDatos()
    Dim rsDir As ADODB.Recordset
    Dim sCampos As String
    Dim sTablas As String
    
    sCampos = "Tipo_Direccion.Descripcion, Direcciones.Calle, "
    sCampos = sCampos & "Direcciones.Colonia, Direcciones.Ciudad, Direcciones.CodPos, "
    sCampos = sCampos & "Direcciones.Tel1, Direcciones.Tel2, Direcciones.Fax, "
    sCampos = sCampos & "Direcciones.DeloMuni, Direcciones.Estado, "
    sCampos = sCampos & "Direcciones.idDireccion, Direcciones.IdTipoDireccion, "
    sCampos = sCampos & "Direcciones.RazonSocial, Direcciones.Rfc, "
    sCampos = sCampos & "Direcciones.TipoPersona, Direcciones.Area "
    '18062011
    'HABILITAR PARA DIRECCIONADOS
'    sCampos = sCampos & "Direcciones.NumeroExterior, Direcciones.NumeroInterior, Direcciones.Referencia "
    
    sTablas = "Direcciones LEFT JOIN Tipo_Direccion ON Direcciones.IdTipoDireccion=Tipo_Direccion.IdTipoDireccion "

    InitRecordSet rsDir, sCampos, sTablas, "Direcciones.idDireccion=" & frmAltaSocios.ssdbDireccion.Columns(10).Value, "Descripcion", Conn
    With rsDir
        If (.RecordCount > 0) Then
            frmAltaDir.txtNoReg.Text = frmAltaSocios.ssdbDireccion.Columns(10).Value
    
            If (.Fields("Calle") <> "") Then
                frmAltaDir.txtCalle.Text = .Fields("Calle")
            End If
            
            If (.Fields("Colonia") <> "") Then
                frmAltaDir.txtColonia.Text = .Fields("Colonia")
            End If
            
            If (.Fields("Ciudad") <> "") Then
                frmAltaDir.txtCiudad.Text = .Fields("Ciudad")
            End If
            
            If (.Fields("CodPos") <> "") Then
                frmAltaDir.txtCodPos.Text = .Fields("CodPos")
            End If
            
            If (.Fields("DeloMuni") <> "") Then
                frmAltaDir.txtDeloMuni.Text = .Fields("DeloMuni")
            End If
    
            If (.Fields("Estado") <> "") Then
                frmAltaDir.cbEdos.Text = .Fields("Estado")
'                strLoadEstado = .Fields("Estado")
'
'                LlenaCodigosPostales strLoadEstado
'                frmAltaDir.cboCodigosPostales.SelText = GetNombreCodigoPostal(.Fields("CodPos"))
            End If
            
            If (.Fields("Tel1") <> "") Then
                frmAltaDir.txtTel1.Text = .Fields("Tel1")
            End If
            
            If (.Fields("Tel2") <> "") Then
                frmAltaDir.txtTel2.Text = .Fields("Tel2")
            End If
            
            If (.Fields("Fax") <> "") Then
                frmAltaDir.txtFax.Text = .Fields("Fax")
            End If
            
            If (.Fields("Descripcion") <> "") Then
                frmAltaDir.txtTipoDir.Text = .Fields("Descripcion")
                frmAltaDir.txtCveDir.Text = .Fields("IdTipoDireccion")
            End If
            
            If (.Fields("RazonSocial") <> "") Then
                frmAltaDir.txtRazonSocial.Text = .Fields("RazonSocial")
                frmAltaDir.txtEmpresa.Text = .Fields("RazonSocial")
            End If
            
            If (.Fields("Rfc") <> "") Then
                frmAltaDir.txtRfc.Text = .Fields("Rfc")
            End If
            
            If (.Fields("TipoPersona") <> "") Then
                If .Fields("TipoPersona") = "F" Then
                    Me.optTipoPersona(0).Value = True
                Else
                    Me.optTipoPersona(1).Value = True
                End If
                
            End If
            
            If (.Fields("Area") <> "") Then
                frmAltaDir.txtArea.Text = .Fields("Area")
            End If
            
            '18062011
            'HABILITAR PARA DIRECCIONADOS
'            If (.Fields("NumeroExterior") <> "") Then
'                frmAltaDir.txtNumExterior.Text = .Fields("NumeroExterior")
'            End If
'
'            If (.Fields("NumeroInterior") <> "") Then
'                frmAltaDir.txtNumInterior.Text = .Fields("NumeroInterior")
'            End If
'
'            If (.Fields("Referencia") <> "") Then
'                frmAltaDir.txtReferencia.Text = .Fields("Referencia")
'            End If
            
        End If
        
        .Close
    End With
    Set rsDir = Nothing
    
    If Val(Me.txtCveDir.Text) = 3 Then
        Me.Frame1.Visible = True
        Me.fraDatosEmpresa.Visible = False
    ElseIf Val(Me.txtCveDir.Text) = 2 Then
        Me.Frame1.Visible = False
        Me.fraDatosEmpresa.Visible = True
    Else
        Me.fraDatosEmpresa.Visible = False
        Me.Frame1.Visible = False
    End If
    
End Sub

Private Sub txtCveDir_LostFocus()
    If (Trim(Me.txtCveDir.Text) <> "") Then
        If (IsNumeric(Me.txtCveDir.Text)) Then
            Me.txtTipoDir.Text = LeeXValor("Descripcion", "Tipo_Direccion", "IdTipoDireccion=" & Val(Me.txtCveDir.Text), "Descripcion", "s", Conn)
            If (Trim(Me.txtTipoDir.Text) = "VACIO") Then
                MsgBox "El tipo de dirección seleccionado no existe en la base de datos.", vbExclamation, "KalaSystems"
                Me.txtTipoDir.Text = ""
                Me.txtCveDir.Text = ""
                Me.txtCveDir.SetFocus
            End If
        Else
            MsgBox "La clave del tipo de dirección es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtTipoDir.Text = ""
            Me.txtCveDir.Text = ""
            Me.txtCveDir.SetFocus
        End If
    End If
    
    If Val(Me.txtCveDir.Text) = 3 Then
        Me.Frame1.Visible = True
        Me.fraDatosEmpresa.Visible = False
    ElseIf Val(Me.txtCveDir.Text) = 2 Then
        Me.fraDatosEmpresa.Visible = True
        Me.Frame1.Visible = False
    Else
        Me.fraDatosEmpresa.Visible = False
        Me.Frame1.Visible = False
    End If
    
    Me.txtCveDir.REFRESH
End Sub

'************************************************************
'*                          Ayudas                          *
'************************************************************

Private Sub cmdAyuda_Click()
Const DATOSDIR = 2
Dim mFAyuda(DATOSDIR) As String
Dim mAAyuda(DATOSDIR) As Integer
Dim mCAyuda(DATOSDIR) As String
Dim mEAyuda(DATOSDIR) As String

    nAyuda = 1

    Set frmHDir = New frmayuda
    
    mFAyuda(0) = "Tipos de dirección ordenados por clave"
    mFAyuda(1) = "Tipos de dirección ordenados por descripción"
    
    mAAyuda(0) = 800
    mAAyuda(1) = 2500
    
    mCAyuda(0) = "IdTipoDireccion"
    mCAyuda(1) = "Descripcion"
    
    mEAyuda(0) = "Clave"
    mEAyuda(1) = "Descripción"
    
    With frmHDir
        .nColActiva = 0
        .nColsAyuda = DATOSDIR
        .sTabla = "Tipo_Direccion"
        
        .sCondicion = ""
        .sTitAyuda = "Tipos de dirección"
        .lAgregar = True
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
    If (Trim(Me.txtCveDir.Text) <> "") Then
        Me.txtTipoDir.Text = LeeXValor("Descripcion", "Tipo_Direccion", "IdTipodireccion=" & Val(Me.txtCveDir.Text), "Descripcion", "s", Conn)
    End If
    
    If Val(Me.txtCveDir.Text) = 3 Then
        Me.Frame1.Visible = True
        Me.fraDatosEmpresa.Visible = False
    ElseIf Val(Me.txtCveDir.Text) = 2 Then
        Me.Frame1.Visible = False
        Me.fraDatosEmpresa.Visible = True
    Else
        Me.Frame1.Visible = False
        Me.fraDatosEmpresa.Visible = False
    End If
    
    Me.cmdAyuda.SetFocus
End Sub

