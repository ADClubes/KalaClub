VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmTipoUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Usuario (Captura)"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6570
   Icon            =   "frmTipoUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6570
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6165
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Generales"
      TabPicture(0)   =   "frmTipoUsuario.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblUsuario(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUsuario(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUsuario(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblUsuario(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblUsuario(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frmTipo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboParentesco"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboMaxima"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cboMinima"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkFamiliar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtUsuario(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtUsuario(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Conceptos"
      TabPicture(1)   =   "frmTipoUsuario.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdElimina"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAgrega"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "sscmbConceptos"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "ssdbgConceptos"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "adodcConceptos"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "AdodcConcepCombo"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblConcepto"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdElimina 
         Caption         =   "Elimina"
         Height          =   375
         Left            =   -73320
         TabIndex        =   24
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgrega 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   -74760
         TabIndex        =   23
         Top             =   2760
         Width           =   1095
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbConceptos 
         Bindings        =   "frmTipoUsuario.frx":0342
         Height          =   375
         Left            =   -74760
         TabIndex        =   21
         Top             =   2040
         Width           =   975
         DataFieldList   =   "IdConcepto"
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgConceptos 
         Bindings        =   "frmTipoUsuario.frx":0361
         Height          =   1455
         Left            =   -74760
         TabIndex        =   20
         Top             =   480
         Width           =   4935
         _Version        =   196616
         AllowUpdate     =   0   'False
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   8705
         _ExtentY        =   2566
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
      Begin VB.TextBox txtUsuario 
         Height          =   330
         Index           =   1
         Left            =   1440
         TabIndex        =   14
         Top             =   855
         Width           =   3600
      End
      Begin VB.TextBox txtUsuario 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   855
         Width           =   1020
      End
      Begin VB.CheckBox chkFamiliar 
         Alignment       =   1  'Right Justify
         Caption         =   "&Familiar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   12
         Top             =   3030
         Width           =   1065
      End
      Begin VB.ComboBox cboMinima 
         Height          =   315
         ItemData        =   "frmTipoUsuario.frx":037E
         Left            =   240
         List            =   "frmTipoUsuario.frx":0380
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1620
         Width           =   1200
      End
      Begin VB.ComboBox cboMaxima 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1620
         Width           =   1200
      End
      Begin VB.ComboBox cboParentesco 
         Height          =   315
         ItemData        =   "frmTipoUsuario.frx":0382
         Left            =   240
         List            =   "frmTipoUsuario.frx":0392
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2520
         Width           =   2640
      End
      Begin VB.Frame frmTipo 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   3330
         TabIndex        =   5
         Top             =   1380
         Width           =   1710
         Begin VB.OptionButton optTipo 
            Caption         =   "&Propietario"
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
            Left            =   225
            TabIndex        =   8
            Top             =   300
            Width           =   1365
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "&Rentista"
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
            Left            =   225
            TabIndex        =   7
            Top             =   825
            Width           =   1365
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "&Membresia"
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
            Left            =   225
            TabIndex        =   6
            Top             =   1350
            Width           =   1365
         End
      End
      Begin MSAdodcLib.Adodc adodcConceptos 
         Height          =   330
         Left            =   -73560
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc AdodcConcepCombo 
         Height          =   375
         Left            =   -73920
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label lblConcepto 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -73680
         TabIndex        =   22
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "P&arentesco:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   2220
         Width           =   1125
      End
      Begin VB.Label lblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Edad Má&xima:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1665
         TabIndex        =   18
         Top             =   1350
         Width           =   1320
      End
      Begin VB.Label lblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "Edad Mí&nima:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   1350
         Width           =   1320
      End
      Begin VB.Label lblUsuario 
         BackStyle       =   0  'Transparent
         Caption         =   "&Descripción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1455
         TabIndex        =   16
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lblUsuario 
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
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Default         =   -1  'True
      Height          =   840
      Left            =   5580
      Picture         =   "frmTipoUsuario.frx":03BB
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Guardar"
      Top             =   1065
      Width           =   795
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   5565
      Picture         =   "frmTipoUsuario.frx":07FD
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TIPOS DE USUARIOS"
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
      Left            =   1125
      TabIndex        =   3
      Top             =   75
      Width           =   3720
   End
   Begin VB.Label lblClave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   195
      TabIndex        =   2
      Top             =   780
      Width           =   1020
   End
End
Attribute VB_Name = "frmTipoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA: TIPO USUARIO
' Objetivo: CATÁLOGO DE LOS USUARIOS DEL CLUB (no confundir con usuarios del sistema)
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim strTipo As String
    Dim AdoRcsUsuarios As ADODB.Recordset

Private Function VerificaDatos()
    Dim i, intInicio As Integer
    If txtUsuario(0).Visible = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    For i = intInicio To 1          'Que no estén vacías las casillas
        If (txtUsuario(i).Text = "") Then
            MsgBox "¡ Favor de Llenar Todas las Casillas, Pues NO son Opcionales !", _
                vbOKOnly + vbExclamation, "Tipo de Usuario (Captura)"
            VerificaDatos = False
            txtUsuario(i).SetFocus
            Exit Function
        End If
        VerificaDatos = True
    Next i
    If (txtUsuario(0).Visible = True) And (Val(txtUsuario(0).Text) < 1) Then
        MsgBox "¡ Error en la CLAVE !", vbOKOnly + vbExclamation, "Tipo de Usuario (Captura)"
        VerificaDatos = False
        txtUsuario(0).SetFocus
        Exit Function
    End If
    If (cboMinima.Text = "") Or (Val(cboMinima.Text) < 0) Then
        MsgBox "¡ Error en la EDAD MÍNIMA !", vbOKOnly + vbExclamation, "Tipo de Usuario (Captura)"
        VerificaDatos = False
        cboMinima.SetFocus
        Exit Function
    End If
    If (cboMaxima.Text = "") Or (Val(cboMaxima.Text) < 0) Then
        MsgBox "¡ Error en la EDAD MÁXIMA !", _
                        vbOKOnly + vbExclamation, "Tipo de Usuario (Captura)"
        VerificaDatos = False
        cboMaxima.SetFocus
        Exit Function
    End If
    If Val(cboMaxima.Text) < Val(cboMinima.Text) Then
        MsgBox "¡ La EDAD MÁXIMA no puede ser menor que la EDAD MÍNIMA !", _
                        vbOKOnly + vbExclamation, "Tipo de Usuario (Captura)"
        VerificaDatos = False
        cboMaxima.SetFocus
        Exit Function
    End If
    If cboParentesco.Text = "" Then
        MsgBox "¡ Favor de Seleccionar el PARENTESCO, Pues No es Opcional !", _
                        vbOKOnly + vbExclamation, "Tipo de Usuario (Captura)"
        VerificaDatos = False
        cboParentesco.SetFocus
        Exit Function
    End If
    If (frmTipo.Visible = True) And (optTipo(0).Value = False) And (optTipo(1).Value = False) And (optTipo(2).Value = False) Then
        MsgBox "¡ Si el Usuario No es un Familiar debe Seleccionar el TIPO !", _
                        vbOKOnly + vbExclamation, "Tipo de Usuario (Captura)"
        VerificaDatos = False
        Exit Function
    End If
        
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
    
        #If SqlServer_ Then
            strSQL = "SELECT idtipousuario FROM tipo_usuario WHERE descripcion = '" & _
                        Trim(txtUsuario(1).Text) & "' AND edadminima = " & cboMinima.ListIndex & _
                        " AND edadmaxima = " & cboMaxima.ListIndex & " AND parentesco = '" & _
                        Left(cboParentesco.Text, 2) & "' AND tipo = '" & strTipo & "' AND familiar = " & _
                        IIf(chkFamiliar.Value = 1, "1", "0") & " OR idtipousuario = " & _
                        Val(txtUsuario(0).Text)
        #Else
            strSQL = "SELECT idtipousuario FROM tipo_usuario WHERE (descripcion = '" & _
                        Trim(txtUsuario(1).Text) & "') AND (edadminima = " & cboMinima.ListIndex & _
                        ") AND (edadmaxima = " & cboMaxima.ListIndex & ") AND (parentesco = '" & _
                        Left(cboParentesco.Text, 2) & "') AND (tipo = '" & strTipo & "') AND (familiar = " & _
                        IIf(chkFamiliar.Value = 1, "True", "False") & ") OR (idtipousuario = " & _
                        Val(txtUsuario(0).Text) & ")"
        #End If
        
        Set AdoRcsUsuarios = New ADODB.Recordset
        AdoRcsUsuarios.ActiveConnection = Conn
        AdoRcsUsuarios.LockType = adLockOptimistic
        AdoRcsUsuarios.CursorType = adOpenKeyset
        AdoRcsUsuarios.CursorLocation = adUseServer
        AdoRcsUsuarios.Open strSQL
        If Not AdoRcsUsuarios.EOF Then
            MsgBox "¡ Ya existe un registro con esa Clave o con esa Información !", vbCritical + vbOKOnly, "Tipos usuarios"
            AdoRcsUsuarios.Close
            VerificaDatos = False
            txtUsuario(0).SetFocus
            Exit Function
        Else
            AdoRcsUsuarios.Close
            VerificaDatos = True
        End If
    End If
End Function

Private Sub cmdAgrega_Click()
    
    Dim AdoCmdInserta As ADODB.Command
    
    If Me.sscmbConceptos.Text = vbNullString Then
        Exit Sub
    End If
    
    
    
    
    strSQL = "INSERT INTO CONCEPTO_TIPO"
    strSQL = strSQL & "("
    strSQL = strSQL & "IdTipoUsuario" & ","
    strSQL = strSQL & "IdConcepto" & ","
    strSQL = strSQL & "Periodo"
    strSQL = strSQL & ")"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & Val(Trim(frmCatalogos.lblModo.Caption)) & ","
    strSQL = strSQL & Me.sscmbConceptos.Columns("IdConcepto").Value & ","
    strSQL = strSQL & 1
    strSQL = strSQL & ")"
    
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandType = adCmdText
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    
    
    Set AdoCmdInserta = Nothing
    
    
    
End Sub

Private Sub cmdElimina_Click()
    Dim AdoCmdInserta As ADODB.Command
    
    
    strSQL = "DELETE FROM CONCEPTO_TIPO"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & " IdTipoUsuario=" & Val(Trim(frmCatalogos.lblModo.Caption))
    strSQL = strSQL & " AND IdConcepto=" & Me.ssdbgConceptos.Columns("IdConcepto").Value
    strSQL = strSQL & ")"
    
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandType = adCmdText
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    
    
    Set AdoCmdInserta = Nothing

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

Private Sub chkFamiliar_Click()
    Dim i As Byte
    For i = 0 To 2
        optTipo(i).Value = False
    Next i
    If chkFamiliar.Value = 1 Then
        frmTipo.Visible = False
    Else
        frmTipo.Visible = True
    End If
    strTipo = "INDISTINTO"
End Sub

Private Sub Form_Load()
    Dim i As Integer
    frmCatalogos.Enabled = False
    For i = 0 To 150
        cboMinima.AddItem i
        cboMaxima.AddItem i
    Next i
    If frmCatalogos.lblModo.Caption = "A" Then
        txtUsuario(0).Visible = True
        Call Llena_txtUsuario
    Else
        txtUsuario(0).Visible = False
        Call LlenaDatos
    End If
    
    Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
        frmCatalogos.Enabled = True
End Sub

Private Sub GuardaDatos()
    Dim AdoCmdInserta As ADODB.Command
    On Error GoTo err_Guarda
    Screen.MousePointer = vbHourglass
    strSQL = "INSERT INTO tipo_usuario (idtipousuario, descripcion, edadminima, edadmaxima, " & _
                    "parentesco, tipo, familiar) VALUES (" & Val(Trim(txtUsuario(0))) & ", '" & _
                    Trim(txtUsuario(1)) & "', " & Val(cboMinima.Text) & ", " & Val(cboMaxima.Text) & ", '" & _
                    Left(cboParentesco.Text, 2) & "', '" & strTipo & "', " & chkFamiliar.Value & ")"
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    Screen.MousePointer = vbDefault
    Conn.CommitTrans                        'Termina transacción
    MsgBox "¡ Registro Ingresado !"
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
    
    Screen.MousePointer = vbHourglass
   
    'Remplazamos el registro sustituyendo al anterior
    strSQL = "UPDATE tipo_usuario SET descripcion = '" & Trim(txtUsuario(1)) & "', edadminima = " & _
                    Val(cboMinima.Text) & ", edadmaxima = " & Val(cboMaxima.Text) & ", parentesco = '" & _
                    Left(cboParentesco.Text, 2) & "', familiar = " & chkFamiliar.Value & ", tipo = '" & _
                    strTipo & "' WHERE idtipousuario = " & Val(lblClave.Caption)
    iniTrans = Conn.BeginTrans                              'Iniciamos transacción
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    Conn.CommitTrans                                                    'Termina transacción
    Screen.MousePointer = vbDefault
    MsgBox "¡ Registro Actualizado !"
    Unload Me
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub optTipo_Click(index As Integer)
    Select Case index
        Case 0: strTipo = "PROPIETARIO"
        Case 1: strTipo = "RENTISTA"
        Case 2: strTipo = "MEMBRESÍA"
    End Select
End Sub



Private Sub sscmbConceptos_Click()
    Me.lblConcepto.Caption = Me.sscmbConceptos.Columns("Descripcion").Value
End Sub

Private Sub txtUsuario_GotFocus(index As Integer)
    txtUsuario(index).SelStart = 0
    txtUsuario(index).SelLength = Len(txtUsuario(index))
End Sub

Private Sub txtUsuario_KeyPress(index As Integer, KeyAscii As Integer)
    Select Case index
        Case 0, 3, 4, 5
            Select Case KeyAscii
                Case 8, 22, 48 To 57  'Backspace, <Ctrl+V> y del 0 al 9
                    KeyAscii = KeyAscii
                Case Else
                    KeyAscii = 0
                End Select
        Case 1, 2
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Sub Limpia()
    Dim i As Integer
    txtUsuario(0).Text = ""
    txtUsuario(1).Text = ""
    cboMinima.ListIndex = -1
    cboMaxima.ListIndex = -1
    cboParentesco.ListIndex = -1
    chkFamiliar.Value = vbUnchecked
    For i = 0 To 2
        optTipo(i).Value = False
    Next i
    Call Llena_txtUsuario
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM tipo_usuario WHERE idtipousuario = " & _
                  Val(Trim(frmCatalogos.lblModo.Caption))
    Set AdoRcsUsuarios = New ADODB.Recordset
    AdoRcsUsuarios.ActiveConnection = Conn
    AdoRcsUsuarios.LockType = adLockOptimistic
    AdoRcsUsuarios.CursorType = adOpenKeyset
    AdoRcsUsuarios.CursorLocation = adUseServer
    AdoRcsUsuarios.Open strSQL
    If Not AdoRcsUsuarios.EOF Then
        lblClave.Caption = AdoRcsUsuarios!idtipousuario
        txtUsuario(1).Text = AdoRcsUsuarios!Descripcion
        cboMinima.ListIndex = AdoRcsUsuarios!edadminima
        cboMaxima.ListIndex = AdoRcsUsuarios!edadmaxima
        Select Case AdoRcsUsuarios!parentesco
            Case "TI"
                cboParentesco.ListIndex = 0
            Case "CO"
                cboParentesco.ListIndex = 1
            Case "HI"
                cboParentesco.ListIndex = 2
            Case "DE"
                cboParentesco.ListIndex = 3
        End Select
        Select Case AdoRcsUsuarios!tipo
            Case "PROPIETARIO"
                optTipo(0).Value = True
            Case "RENTISTA"
                optTipo(1).Value = True
            Case "MEMBRESÍA"
                optTipo(1).Value = True
        End Select
        'cboParentesco.Text = AdoRcsUsuarios!parentesco
        chkFamiliar.Value = IIf(AdoRcsUsuarios!familiar = True, vbChecked, vbUnchecked)
        
        LlenaGridConceptos Val(Trim(frmCatalogos.lblModo.Caption))
        Me.ssdbgConceptos.REFRESH
        Me.sscmbConceptos.REFRESH
    End If
    
End Sub

Sub Llena_txtUsuario()
    Dim lngAnterior, lngUsuario As Long
    'Llena txtUsuario
    lngUsuario = 1
    strSQL = "SELECT idtipousuario FROM tipo_usuario ORDER BY idtipousuario"
    Set AdoRcsUsuarios = New ADODB.Recordset
    AdoRcsUsuarios.ActiveConnection = Conn
    AdoRcsUsuarios.LockType = adLockOptimistic
    AdoRcsUsuarios.CursorType = adOpenKeyset
    AdoRcsUsuarios.CursorLocation = adUseServer
    AdoRcsUsuarios.Open strSQL
    If AdoRcsUsuarios.EOF Then
        lngUsuario = 1
        txtUsuario(0).Text = lngUsuario
        Exit Sub
    End If
    AdoRcsUsuarios.MoveFirst
    Do While Not AdoRcsUsuarios.EOF
        If AdoRcsUsuarios.Fields!idtipousuario <> "1" Then
            If Val(AdoRcsUsuarios.Fields!idtipousuario) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngUsuario
        AdoRcsUsuarios.MoveNext
        If Not AdoRcsUsuarios.EOF Then lngUsuario = AdoRcsUsuarios.Fields!idtipousuario
    Loop
    txtUsuario(0).Text = lngAnterior + 1
End Sub
Private Sub LlenaGridConceptos(lIdTipoUsuario As Long)
    
    strSQL = "SELECT CONCEPTO_TIPO.IdConcepto, CONCEPTO_INGRESOS.Descripcion"
    strSQL = strSQL & " FROM CONCEPTO_TIPO INNER JOIN CONCEPTO_INGRESOS ON CONCEPTO_TIPO.IdConcepto = CONCEPTO_INGRESOS.IdConcepto"
    strSQL = strSQL & " Where(((CONCEPTO_TIPO.idTipoUsuario) =" & lIdTipoUsuario & "))"
    
    Me.adodcConceptos.ConnectionString = Conn
    Me.adodcConceptos.CursorLocation = adUseServer
    Me.adodcConceptos.CursorType = adOpenStatic
    Me.adodcConceptos.LockType = adLockReadOnly
    Me.adodcConceptos.RecordSource = strSQL
    
    Me.adodcConceptos.REFRESH
    
    
    
    strSQL = "SELECT CONCEPTO_INGRESOS.IdConcepto, CONCEPTO_INGRESOS.Descripcion"
    strSQL = strSQL & " FROM CONCEPTO_INGRESOS"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & "(((CONCEPTO_INGRESOS.IdConcepto) Not In (SELECT Idconcepto from Concepto_Tipo Where idtipousuario=" & lIdTipoUsuario & "))"
    strSQL = strSQL & " AND ((CONCEPTO_INGRESOS.EsPeriodico)=-1))"
    strSQL = strSQL & " ORDER BY CONCEPTO_INGRESOS.IdConcepto"
    
    
    Me.AdodcConcepCombo.ConnectionString = Conn
    Me.AdodcConcepCombo.CursorLocation = adUseServer
    Me.AdodcConcepCombo.CursorType = adOpenStatic
    Me.AdodcConcepCombo.LockType = adLockReadOnly
    Me.AdodcConcepCombo.RecordSource = strSQL
    
    Me.AdodcConcepCombo.REFRESH
    
    
End Sub
