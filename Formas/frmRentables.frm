VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmRentables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rentables  (Captura)"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6045
   Icon            =   "frmRentables.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6045
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Rentable"
      TabPicture(0)   =   "frmRentables.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescripcion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblRentables(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblRentables(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblRentables(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblClave"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkPropiedad"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cboTipoRentable"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "frmSexo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtRentables(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtRentables(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Pago"
      TabPicture(1)   =   "frmRentables.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRentables(5)"
      Tab(1).Control(1)=   "dtpPagoInicio"
      Tab(1).Control(2)=   "txtRentables(4)"
      Tab(1).Control(3)=   "txtRentables(3)"
      Tab(1).Control(4)=   "txtRentables(0)"
      Tab(1).Control(5)=   "dtpFechaPago"
      Tab(1).Control(6)=   "lblRentables(8)"
      Tab(1).Control(7)=   "lblRentables(7)"
      Tab(1).Control(8)=   "lblRentables(6)"
      Tab(1).Control(9)=   "lblRentables(5)"
      Tab(1).Control(10)=   "lblRentables(4)"
      Tab(1).Control(11)=   "lblRentables(3)"
      Tab(1).ControlCount=   12
      Begin VB.Frame Frame1 
         Caption         =   "Asignado a"
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   2760
         Width           =   4575
         Begin VB.TextBox txtRentables 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   330
            Index           =   7
            Left            =   1080
            TabIndex        =   31
            Top             =   240
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox txtRentables 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   330
            Index           =   6
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.TextBox txtRentables 
         Height          =   330
         Index           =   5
         Left            =   -72840
         MaxLength       =   50
         TabIndex        =   27
         Top             =   3000
         Width           =   1600
      End
      Begin MSComCtl2.DTPicker dtpPagoInicio 
         Height          =   375
         Left            =   -72840
         TabIndex        =   25
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   39885
      End
      Begin VB.TextBox txtRentables 
         Height          =   330
         Index           =   4
         Left            =   -72840
         MaxLength       =   15
         TabIndex        =   24
         Top             =   2520
         Width           =   1600
      End
      Begin VB.TextBox txtRentables 
         Height          =   330
         Index           =   3
         Left            =   -72840
         MaxLength       =   15
         TabIndex        =   19
         Top             =   2040
         Width           =   1600
      End
      Begin VB.TextBox txtRentables 
         Height          =   330
         Index           =   0
         Left            =   -72840
         MaxLength       =   15
         TabIndex        =   18
         Top             =   1560
         Width           =   1600
      End
      Begin MSComCtl2.DTPicker dtpFechaPago 
         Height          =   375
         Left            =   -72840
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   39885
      End
      Begin VB.TextBox txtRentables 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   2
         Left            =   1170
         TabIndex        =   10
         Top             =   1665
         Width           =   1600
      End
      Begin VB.TextBox txtRentables 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   1
         Left            =   1170
         TabIndex        =   9
         Top             =   1080
         Width           =   1600
      End
      Begin VB.Frame frmSexo 
         Caption         =   "Sexo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   2925
         TabIndex        =   5
         Top             =   960
         Width           =   1770
         Begin VB.OptionButton optSexo 
            Caption         =   "&Indistinto"
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
            Left            =   315
            TabIndex        =   8
            Top             =   1335
            Width           =   1110
         End
         Begin VB.OptionButton optSexo 
            Caption         =   "&Femenino"
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
            Left            =   315
            TabIndex        =   7
            Top             =   810
            Width           =   1140
         End
         Begin VB.OptionButton optSexo 
            Caption         =   "&Masculino"
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
            Left            =   315
            TabIndex        =   6
            Top             =   285
            Width           =   1185
         End
      End
      Begin VB.ComboBox cboTipoRentable 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   510
         Width           =   3255
      End
      Begin VB.CheckBox chkPropiedad 
         Alignment       =   1  'Right Justify
         Caption         =   "Uso Diario"
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
         Left            =   135
         TabIndex        =   3
         Top             =   2160
         Width           =   1530
      End
      Begin VB.Label lblRentables 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Index           =   8
         Left            =   -74520
         TabIndex        =   28
         Top             =   3120
         Width           =   1425
      End
      Begin VB.Label lblRentables 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pagado desde:"
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
         Index           =   7
         Left            =   -74640
         TabIndex        =   26
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label lblRentables 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
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
         Index           =   6
         Left            =   -74400
         TabIndex        =   23
         Top             =   2640
         Width           =   1425
      End
      Begin VB.Label lblRentables 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Pagado:"
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
         Index           =   5
         Left            =   -74400
         TabIndex        =   22
         Top             =   2160
         Width           =   1425
      End
      Begin VB.Label lblRentables 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Area:"
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
         Index           =   4
         Left            =   -74400
         TabIndex        =   21
         Top             =   1680
         Width           =   1425
      End
      Begin VB.Label lblRentables 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pagado hasta:"
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
         Index           =   3
         Left            =   -74640
         TabIndex        =   20
         Top             =   1200
         Width           =   1665
      End
      Begin VB.Label lblClave 
         Alignment       =   1  'Right Justify
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
         Left            =   1440
         TabIndex        =   15
         Top             =   540
         Width           =   660
      End
      Begin VB.Label lblRentables 
         BackStyle       =   0  'Transparent
         Caption         =   "&Ubicación:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   1710
         Width           =   945
      End
      Begin VB.Label lblRentables 
         BackStyle       =   0  'Transparent
         Caption         =   "&Número:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1155
         Width           =   690
      End
      Begin VB.Label lblRentables 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo &Rentable: :"
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
         Left            =   120
         TabIndex        =   12
         Top             =   585
         Width           =   1365
      End
      Begin VB.Label lblDescripcion 
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
         Left            =   2235
         TabIndex        =   11
         Top             =   540
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Height          =   840
      Left            =   5055
      Picture         =   "frmRentables.frx":047A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Guardar"
      Top             =   1050
      Width           =   795
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   5070
      Picture         =   "frmRentables.frx":08BC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   2190
      Width           =   795
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RENTABLES"
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
      Left            =   960
      TabIndex        =   16
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmRentables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA: RENTABLES
' Objetivo: CATÁLOGO DE ITEMS RENTABLES
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsRentables As ADODB.Recordset
    Dim intUsuario, intTipoRentable As Integer
    Dim strSexo, strNumero As String

Private Function VerificaDatos()
    If (cboTipoRentable.Visible = True) And (cboTipoRentable.Text = "") Then
        MsgBox "¡ Favor de seleccionar el TIPO RENTABLE, No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de rentables (Captura)"
        VerificaDatos = False
        cboTipoRentable.SetFocus
        Exit Function
    End If
    If txtRentables(1).Text = "" Then
        MsgBox "¡ Favor de Llenar la Casilla NÚMERO, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "rentables (Captura)"
        VerificaDatos = False
        txtRentables(1).SetFocus
        Exit Function
    End If
    If (optSexo(0).Value = vbUnchecked) And (optSexo(1).Value = vbUnchecked) And (optSexo(2).Value = vbUnchecked) Then
        MsgBox "¡ Favor de Seleccionar el SEXO, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de rentables (Captura)"
        VerificaDatos = False
        Exit Function
    End If
    

    
    
    
    
    strNumero = Trim(txtRentables(1).Text)
    VerificaDatos = True
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        intTipoRentable = cboTipoRentable.ItemData(cboTipoRentable.ListIndex)
        strSQL = "SELECT idtiporentable FROM rentables WHERE (idtiporentable = " & _
                        intTipoRentable & ") AND (numero = '" & txtRentables(1) & "')"
        Set AdoRcsRentables = New ADODB.Recordset
        AdoRcsRentables.ActiveConnection = Conn
        AdoRcsRentables.LockType = adLockOptimistic
        AdoRcsRentables.CursorType = adOpenKeyset
        AdoRcsRentables.CursorLocation = adUseServer
        AdoRcsRentables.Open strSQL
        If Not AdoRcsRentables.EOF Then
            MsgBox "Ya Existe Un Registro con el Tipo Rentable: " & _
                        cboTipoRentable.Text & " y con el Número: " & _
                        txtRentables(1), vbInformation + vbOKOnly, "Rentables"
            AdoRcsRentables.Close
            VerificaDatos = False
            cboTipoRentable.SetFocus
            Exit Function
        Else
            AdoRcsRentables.Close
            VerificaDatos = True
        End If
    End If
End Function

Private Sub cmdGuardar_Click()
    Dim blnGuarda As Boolean
    blnGuarda = VerificaDatos
    If blnGuarda = True Then
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
    Dim i As Integer
    Dim strCampo1, strCampo2 As String
    frmCatalogos.Enabled = False
    
    'Llena los combo Id Tipo Rentable
    strSQL = "SELECT idtiporentable, descripcion FROM tipo_rentables"
    strCampo1 = "descripcion"
    strCampo2 = "idtiporentable"
    Call LlenaCombos(cboTipoRentable, strSQL, strCampo1, strCampo2)
    
    
    
    If frmCatalogos.lblModo.Caption = "A" Then
        cboTipoRentable.Visible = True
    Else
        cboTipoRentable.Visible = True
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
    Call LlenaEspacios
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO Rentables ("
        strSQL = strSQL & "idtiporentable,"
        strSQL = strSQL & " Numero,"
        strSQL = strSQL & " Sexo,"
        strSQL = strSQL & " Ubicacion,"
        strSQL = strSQL & " FechaPago,"
        strSQL = strSQL & " Propiedad,"
        strSQL = strSQL & " Observaciones,"
        strSQL = strSQL & " Area,"
        strSQL = strSQL & " FechaInicio,"
        strSQL = strSQL & " ImportePagado,"
        strSQL = strSQL & " Documento)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & cboTipoRentable.ItemData(cboTipoRentable.ListIndex) & ","
        strSQL = strSQL & "'" & strNumero & "',"
        strSQL = strSQL & "'" & strSexo & "',"
        strSQL = strSQL & "'" & Trim(txtRentables(2).Text) & "',"
        strSQL = strSQL & "'" & Format(Me.dtpFechaPago.Value, "yyyymmdd") & "',"
        strSQL = strSQL & IIf(Me.chkPropiedad.Value, -1, 0) & ","
        strSQL = strSQL & "'" & Trim(txtRentables(5).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtRentables(0).Text) & "',"
        strSQL = strSQL & "'" & Format(Me.dtpPagoInicio.Value, "yyyymmdd") & "',"
        strSQL = strSQL & Trim(txtRentables(3).Text) & ","
        strSQL = strSQL & "'" & Trim(txtRentables(4).Text) & "')"
    #Else
        strSQL = "INSERT INTO Rentables ("
        strSQL = strSQL & "idtiporentable,"
        strSQL = strSQL & " Numero,"
        strSQL = strSQL & " Sexo,"
        strSQL = strSQL & " Ubicacion,"
        strSQL = strSQL & " FechaPago,"
        strSQL = strSQL & " Propiedad,"
        strSQL = strSQL & " Observaciones,"
        strSQL = strSQL & " Area,"
        strSQL = strSQL & " FechaInicio,"
        strSQL = strSQL & " ImportePagado,"
        strSQL = strSQL & " Documento)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & cboTipoRentable.ItemData(cboTipoRentable.ListIndex) & ","
        strSQL = strSQL & "'" & strNumero & "',"
        strSQL = strSQL & "'" & strSexo & "',"
        strSQL = strSQL & "'" & Trim(txtRentables(2).Text) & "',"
        strSQL = strSQL & "#" & Format(Me.dtpFechaPago.Value, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & IIf(Me.chkPropiedad.Value, -1, 0) & ","
        strSQL = strSQL & "'" & Trim(txtRentables(5).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtRentables(0).Text) & "',"
        strSQL = strSQL & "#" & Format(Me.dtpPagoInicio.Value, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & Trim(txtRentables(3).Text) & ","
        strSQL = strSQL & "'" & Trim(txtRentables(4).Text) & "')"
    #End If
                    
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    MsgBox "¡Registro Ingresado!"
    Call Limpia
    frmCatalogos.AdoDcCatal.REFRESH
    frmCatalogos.grdCatalogos.REFRESH
    frmCatalogos.lblTotal.Caption = Format(frmCatalogos.AdoDcCatal.Recordset.RecordCount, "#######")
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
    Call LlenaEspacios
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    'Eliminamos elregistro existente
    strSQL = "DELETE"
    strSQL = strSQL & " FROM rentables"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " (idtiporentable = " & Val(lblClave.Caption) & ")"
    strSQL = strSQL & " AND (numero = '" & strNumero & "')"
    
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Ahora Insertamos el nuevo registro que sustituye al anterior
    #If SqlServer_ Then
        strSQL = "INSERT INTO Rentables ("
        strSQL = strSQL & "idtiporentable,"
        strSQL = strSQL & " Numero,"
        strSQL = strSQL & " Sexo,"
        strSQL = strSQL & " Ubicacion,"
        strSQL = strSQL & " IdUsuario,"
        strSQL = strSQL & " FechaPago,"
        strSQL = strSQL & " Propiedad,"
        strSQL = strSQL & " Observaciones,"
        strSQL = strSQL & " Area,"
        strSQL = strSQL & " FechaInicio,"
        strSQL = strSQL & " ImportePagado,"
        strSQL = strSQL & " Documento)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & cboTipoRentable.ItemData(cboTipoRentable.ListIndex) & ","
        strSQL = strSQL & "'" & strNumero & "',"
        strSQL = strSQL & "'" & strSexo & "',"
        strSQL = strSQL & "'" & Trim(txtRentables(2).Text) & "',"
        strSQL = strSQL & Me.txtRentables(7) & ","
        strSQL = strSQL & "'" & Format(Me.dtpFechaPago.Value, "yyyymmdd") & "',"
        strSQL = strSQL & IIf(Me.chkPropiedad.Value, -1, 0) & ","
        strSQL = strSQL & "'" & Trim(txtRentables(5).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtRentables(0).Text) & "',"
        strSQL = strSQL & "'" & Format(Me.dtpPagoInicio.Value, "yyyymmdd") & "',"
        strSQL = strSQL & Trim(txtRentables(3).Text) & ","
        strSQL = strSQL & "'" & Trim(txtRentables(4).Text) & "')"
    #Else
        strSQL = "INSERT INTO Rentables ("
        strSQL = strSQL & "idtiporentable,"
        strSQL = strSQL & " Numero,"
        strSQL = strSQL & " Sexo,"
        strSQL = strSQL & " Ubicacion,"
        strSQL = strSQL & " IdUsuario,"
        strSQL = strSQL & " FechaPago,"
        strSQL = strSQL & " Propiedad,"
        strSQL = strSQL & " Observaciones,"
        strSQL = strSQL & " Area,"
        strSQL = strSQL & " FechaInicio,"
        strSQL = strSQL & " ImportePagado,"
        strSQL = strSQL & " Documento)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & cboTipoRentable.ItemData(cboTipoRentable.ListIndex) & ","
        strSQL = strSQL & "'" & strNumero & "',"
        strSQL = strSQL & "'" & strSexo & "',"
        strSQL = strSQL & "'" & Trim(txtRentables(2).Text) & "',"
        strSQL = strSQL & Me.txtRentables(7) & ","
        strSQL = strSQL & "#" & Format(Me.dtpFechaPago.Value, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & IIf(Me.chkPropiedad.Value, -1, 0) & ","
        strSQL = strSQL & "'" & Trim(txtRentables(5).Text) & "',"
        strSQL = strSQL & "'" & Trim(txtRentables(0).Text) & "',"
        strSQL = strSQL & "#" & Format(Me.dtpPagoInicio.Value, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & Trim(txtRentables(3).Text) & ","
        strSQL = strSQL & "'" & Trim(txtRentables(4).Text) & "')"
    #End If
    
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    Conn.CommitTrans      'Termina transacción
    Screen.MousePointer = vbDefault
    frmCatalogos.AdoDcCatal.REFRESH
    frmCatalogos.grdCatalogos.REFRESH
    frmCatalogos.lblTotal.Caption = Format(frmCatalogos.AdoDcCatal.Recordset.RecordCount, "#######")
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










Private Sub optSexo_Click(index As Integer)
    Select Case index
        Case 0
            strSexo = "M"
        Case 1
            strSexo = "F"
        Case 2
            strSexo = "X"
    End Select
End Sub

Private Sub txtRentables_GotFocus(index As Integer)
    txtRentables(index).SelStart = 0
    txtRentables(index).SelLength = Len(txtRentables(index))
End Sub

Private Sub txtRentables_KeyPress(index As Integer, KeyAscii As Integer)
    Select Case index
        Case 0
            Select Case KeyAscii
                Case 8, 22, 48 To 57     'Backspace, <Ctrl+V> y del 0 al 9
                    KeyAscii = KeyAscii
                Case Else
                    KeyAscii = 0
            End Select
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Sub Limpia()
    Dim i As Integer
    cboTipoRentable.ListIndex = -1
    For i = 0 To 2
        If i > 0 Then txtRentables(i).Text = ""
        optSexo(i).Value = False
    Next i
End Sub

Sub LlenaEspacios()
    Dim i, intLargo As Integer
    If Left(strNumero, 1) <> " " Then
        intLargo = Len(strNumero)
        If intLargo < 6 Then
            While intLargo < 6
                strNumero = " " & strNumero
                intLargo = Len(strNumero)
            Wend
        End If
    End If
End Sub

Sub LlenaDatos()
    lblRentables(1).Enabled = False
    txtRentables(1).Enabled = False
    strSQL = "SELECT descripcion FROM tipo_rentables WHERE IdTiporentable = " & _
                  Val(Trim(frmCatalogos.lblModo.Caption))
    Set AdoRcsRentables = New ADODB.Recordset
    AdoRcsRentables.ActiveConnection = Conn
    AdoRcsRentables.LockType = adLockOptimistic
    AdoRcsRentables.CursorType = adOpenKeyset
    AdoRcsRentables.CursorLocation = adUseServer
    AdoRcsRentables.Open strSQL
    If Not AdoRcsRentables.EOF Then
        'lblDescripcion.Caption = "- " & AdoRcsRentables!descripcion
        Me.cboTipoRentable.Text = AdoRcsRentables!Descripcion
    End If
    
    #If SqlServer_ Then
        strSQL = "SELECT * FROM rentables WHERE idtiporentable = " & _
                  Val(Trim(frmCatalogos.lblModo.Caption)) & " AND LTRIM(RTrim(numero)) = '" & _
                  Trim(frmCatalogos.lblModoDescrip.Caption) & "'"
    #Else
        strSQL = "SELECT * FROM rentables WHERE (idtiporentable = " & _
                  Val(Trim(frmCatalogos.lblModo.Caption)) & ") AND (Trim(numero) = '" & _
                  Trim(frmCatalogos.lblModoDescrip.Caption) & "')"
    #End If
    
    Set AdoRcsRentables = New ADODB.Recordset
    AdoRcsRentables.ActiveConnection = Conn
    AdoRcsRentables.LockType = adLockOptimistic
    AdoRcsRentables.CursorType = adOpenKeyset
    AdoRcsRentables.CursorLocation = adUseServer
    AdoRcsRentables.Open strSQL
    If Not AdoRcsRentables.EOF Then
        lblClave.Caption = AdoRcsRentables!idtiporentable
        txtRentables(1).Text = Trim(AdoRcsRentables!Numero)
        txtRentables(2).Text = AdoRcsRentables!Ubicacion
        strSexo = AdoRcsRentables!sexo
        intUsuario = IIf(IsNull(AdoRcsRentables!idusuario), 0, AdoRcsRentables!idusuario)
        'lblFechaPago.Caption = IIf(IsNull(AdoRcsRentables!Fechapago), "", AdoRcsRentables!Fechapago)
        chkPropiedad.Value = IIf(AdoRcsRentables!Propiedad <> 0, 1, 0)
        Me.txtRentables(0).Text = IIf(IsNull(AdoRcsRentables!Area), vbNullString, AdoRcsRentables!Area)
        Me.txtRentables(3).Text = IIf(IsNull(AdoRcsRentables!ImportePagado), 0, AdoRcsRentables!ImportePagado)
        Me.txtRentables(4).Text = IIf(IsNull(AdoRcsRentables!Documento), vbNullString, AdoRcsRentables!Documento)
        Me.txtRentables(5).Text = IIf(IsNull(AdoRcsRentables!Observaciones), vbNullString, AdoRcsRentables!Observaciones)
        
        Me.dtpPagoInicio.Value = IIf(IsNull(AdoRcsRentables!FechaInicio), Date, AdoRcsRentables!FechaInicio)
        Me.dtpFechaPago.Value = IIf(IsNull(AdoRcsRentables!Fechapago), Date, AdoRcsRentables!Fechapago)
        
        
        Me.txtRentables(7).Text = IIf(IsNull(AdoRcsRentables!idusuario), 0, AdoRcsRentables!idusuario)
        
    End If
    Select Case strSexo
        Case "M"
            optSexo(0).Value = True
        Case "F"
            optSexo(1).Value = True
        Case "X"
            optSexo(2).Value = True
    End Select
End Sub
