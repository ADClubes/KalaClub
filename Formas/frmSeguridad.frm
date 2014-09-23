VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSeguridad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opciones de Seguridad"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10065
   Icon            =   "frmSeguridad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmSeguridad.frx":0442
   ScaleHeight     =   5895
   ScaleWidth      =   10065
   Begin VB.CommandButton CmdAcceso 
      Height          =   700
      Left            =   7260
      Picture         =   "frmSeguridad.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Actualizar Acceso a Modulos"
      Top             =   5100
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   700
      Left            =   8535
      Picture         =   "frmSeguridad.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5100
      Width           =   1200
   End
   Begin VB.CommandButton CmdTodas 
      Caption         =   "Seleccionar Todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5115
      TabIndex        =   1
      Top             =   4620
      Width           =   2340
   End
   Begin VB.CommandButton CmdDeseleccionar 
      Caption         =   "Desmarcar Todas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7470
      TabIndex        =   0
      Top             =   4620
      Width           =   2340
   End
   Begin MSAdodcLib.Adodc AdodcUsua 
      Height          =   330
      Left            =   1170
      Top             =   2415
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "    AdodcUsua"
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
   Begin MSDataGridLib.DataGrid DataGridUsua 
      Bindings        =   "frmSeguridad.frx":0FD0
      Height          =   4080
      Left            =   225
      TabIndex        =   4
      Top             =   900
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   7197
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "LOGIN_NAME"
         Caption         =   "LOGIN"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NOMBRE"
         Caption         =   "NOMBRE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   -1  'True
            ColumnWidth     =   3060.284
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3720
      Left            =   5100
      TabIndex        =   5
      Top             =   900
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   6562
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1725
      Top             =   5025
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":0FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":133C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":1690
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":1920
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":1C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":1F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":222C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":258C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":28E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":2C34
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":2F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":32DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":3600
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":3954
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":3C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":3FCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":4320
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":4674
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":49C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguridad.frx":4D1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "OPCIONES DE SEGURIDAD"
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
      Height          =   345
      Left            =   2647
      TabIndex        =   8
      Top             =   75
      Width           =   4770
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USUARIOS DEL SISTEMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   585
      Width           =   4725
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ACCESO A MÓDULOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   5100
      TabIndex        =   6
      Top             =   585
      Width           =   4710
   End
End
Attribute VB_Name = "frmSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim LoginOrig As String
    Dim adoRcsSeg As ADODB.Recordset

Private Sub Actualiza_Grid_Usua()
    ' Obtiene los registros de la base de datos de usuarios y los liga al data control
    strSQL = "SELECT * FROM usuarios_sistema"
    AdodcUsua.ConnectionString = Conn
    AdodcUsua.CursorLocation = adUseClient
    AdodcUsua.CursorType = adOpenDynamic
    AdodcUsua.LockType = adLockReadOnly
    AdodcUsua.RecordSource = strSQL
    AdodcUsua.Refresh
    DataGridUsua.Refresh
End Sub

'Private Sub CmdAcceso_Click()
'    Dim I, Z, TotalOpc, ClaveUsuario As Integer
'    Dim Llave, Llave2, Cadena, Posicion As String
'    Dim Arreglo()
'
'    ClaveUsuario = frmUsuarios.AdodcUsua.Recordset.Fields("Consecutivo")
'    TotalOpc = Me.TreeView1.Nodes.Count
'
'    ' LLena arreglo con las opciones seleccionadas
'    ReDim Arreglo(TotalOpc, 1)
'    For I = 1 To TotalOpc
'        Llave = Trim(Me.TreeView1.Nodes(I).Key)
'        Arreglo(I, 0) = Llave
'        Arreglo(I, 1) = Me.TreeView1.Nodes(I).Checked
'    Next
'
'    ' Activa los padres para acceso
'    For I = 1 To TotalOpc
'        Llave = Arreglo(I, 0)
'        If Arreglo(I, 1) = True Then
'            Llave2 = Left(Llave, Len(Llave) - 1)
'            While Len(Llave2) > 3
'                For Z = 1 To TotalOpc
'                    If Len(Llave2) > 3 Then
'                        If Llave2 = Arreglo(Z, 0) Then
'                            Arreglo(Z, 1) = True
'                            Me.TreeView1.Nodes(Z).Checked = True
'                            Me.TreeView1.Nodes(Z).BackColor = vbYellow
'                        End If
'                    End If
'                Next
'                Llave2 = Left(Llave2, Len(Llave2) - 1)
'            Wend
'        End If
'    Next I
'
'    ' Actualiza base de datos
'    strSQL = "SELECT nivel, usuarios FROM usuariosmenu"
'    Set adoRcsSeg = New ADODB.Recordset
'    adoRcsSeg.ActiveConnection = Conn
'    adoRcsSeg.CursorLocation = adUseClient
'    adoRcsSeg.CursorType = adOpenDynamic
'    adoRcsSeg.LockType = adLockOptimistic
'    adoRcsSeg.Open strSQL
'    If Not adoRcsSeg.EOF Then
'        adoRcsSeg.MoveFirst
'        Do While Not adoRcsSeg.EOF
'            For I = 1 To TotalOpc
'                If Arreglo(I, 0) = Trim(adoRcsSeg!Nivel) Then
'                    If IsNull(adoRcsSeg!Usuarios) Then
'                        Cadena = String(200, "0")
'                    Else
'                        Cadena = adoRcsSeg!Usuarios
'                    End If
'                    Posicion = ClaveUsuario * 2 - 1
'                    If Posicion = 1 Then
'                        If Arreglo(I, 1) = True Then
'                            Cadena = Format(ClaveUsuario, "0#") & Mid(Cadena, 3)
'                        Else
'                            Cadena = "00" & Mid(Cadena, 3)
'                        End If
'                    ElseIf Posicion > 1 Then
'                        If Arreglo(I, 1) = True Then
'                            Cadena = Left(Cadena, Posicion - 1) & Format(ClaveUsuario, "0#") & Mid(Cadena, Posicion + 2)
'                        Else
'                            Cadena = Left(Cadena, Posicion - 1) & "00" & Mid(Cadena, Posicion + 2)
'                        End If
'                    End If
'                    Conn.BeginTrans
'                    adoRcsSeg!Usuarios = Cadena
'                    adoRcsSeg.Update
'                    Conn.CommitTrans
'                End If
'            Next I
'            adoRcsSeg.MoveNext
'        Loop
'        MsgBox "Seguridad Actualizada", vbInformation, "Acceso al Sistema"
'    End If
'    adoRcsSeg.Close
'    Set adoRcsSeg = Nothing
'End Sub

Private Sub CmdDeseleccionar_Click()
    Dim I As Integer
    Dim TotalOpc As Integer

    TotalOpc = Me.TreeView1.Nodes.Count
    ' Des-selecciona todas las opciones del arbol
    For I = 1 To TotalOpc
        Me.TreeView1.Nodes(I).Checked = False
        TreeView1.Nodes(I).Bold = False
        TreeView1.Nodes(I).ForeColor = vbBlack
        Me.TreeView1.Nodes(I).BackColor = vbWhite
    Next
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdTodas_Click()
    Dim I, TotalOpc As Integer

    TotalOpc = Me.TreeView1.Nodes.Count
    ' Selecciona todas las opciones del árbol
    For I = 1 To TotalOpc
        TreeView1.Nodes(I).Checked = True
        TreeView1.Nodes(I).Bold = True
        TreeView1.Nodes(I).ForeColor = vbBlue
        TreeView1.Nodes(I).BackColor = vbCyan
    Next
End Sub

Private Sub Form_Activate()
    Actualiza_Grid_Usua
End Sub

Private Sub Form_Load()
    Dim strRelacion, strClave, strTexto  As String
    Dim bytImagen, bytImagenSel As Byte
    Dim nodX As Node    ' Declara una variable Node.
    
    Me.Top = 0
    Me.Left = 0
    Me.Width = 10155
    Me.Height = 6350
    
    strSQL = "SELECT * FROM arbol_menu"
    Set adoRcsSeg = New ADODB.Recordset
    adoRcsSeg.ActiveConnection = Conn
    adoRcsSeg.LockType = adLockReadOnly
    adoRcsSeg.CursorType = adOpenKeyset
    adoRcsSeg.CursorLocation = adUseServer
    adoRcsSeg.Open strSQL
    adoRcsSeg.MoveFirst
    Do While Not adoRcsSeg.EOF
        strClave = adoRcsSeg!Clave
        strTexto = adoRcsSeg!texto
        bytImagen = adoRcsSeg!imagen
        bytImagenSel = adoRcsSeg!imagensel
        If IsNull(adoRcsSeg!relat) Then
            Set nodX = TreeView1.Nodes.Add(Null, Null, strClave, strTexto, bytImagen, bytImagenSel)
        Else
            strRelacion = adoRcsSeg!relat
            Set nodX = TreeView1.Nodes.Add(strRelacion, tvwChild, strClave, strTexto, bytImagen, bytImagenSel)
        End If
        adoRcsSeg.MoveNext
    Loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.AdodcUsua.RecordSource <> "" Then
        Me.AdodcUsua.Recordset.Close
    End If
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim Indice As Integer
    Indice = Node.Index
    If Node.Checked Then
        Me.TreeView1.Nodes(Indice).BackColor = vbCyan
        TreeView1.Nodes(Indice).ForeColor = vbBlue
        TreeView1.Nodes(Indice).Bold = True
    Else
        Me.TreeView1.Nodes(Indice).BackColor = vbWhite
        TreeView1.Nodes(Indice).ForeColor = vbBlack
        TreeView1.Nodes(Indice).Bold = False
    End If
End Sub
