VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCotiza 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cotiza"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Totales"
      Height          =   1695
      Left            =   7440
      TabIndex        =   4
      Top             =   3960
      Width           =   3495
      Begin VB.Label lblCtl 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCtl 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCtl 
         Caption         =   "Label1"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblCtl 
         Caption         =   "Label1"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblCtl 
         Alignment       =   1  'Right Justify
         Caption         =   "Convencional"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblCtl 
         Alignment       =   1  'Right Justify
         Caption         =   "Direccionado"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCtl 
         Alignment       =   1  'Right Justify
         Caption         =   "MC Convencional"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblCtl 
         Alignment       =   1  'Right Justify
         Caption         =   "MC Direccionado"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgCalcula 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   10815
      _Version        =   196616
      DataMode        =   2
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowColumnMoving=   0
      AllowColumnSwapping=   0
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   4868
      Columns(0).Caption=   "Nombre"
      Columns(0).Name =   "Nombre"
      Columns(0).DataField=   "Column 0"
      Columns(0).FieldLen=   256
      Columns(1).Width=   3757
      Columns(1).Caption=   "Tipo"
      Columns(1).Name =   "Tipo"
      Columns(1).DataField=   "Column 1"
      Columns(1).FieldLen=   256
      Columns(2).Width=   2434
      Columns(2).Caption=   "Convencional"
      Columns(2).Name =   "Convencional"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   6
      Columns(2).NumberFormat=   "CURRENCY"
      Columns(2).FieldLen=   256
      Columns(3).Width=   2223
      Columns(3).Caption=   "Direccionado"
      Columns(3).Name =   "Direccionado"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   6
      Columns(3).NumberFormat=   "CURRENCY"
      Columns(3).FieldLen=   256
      Columns(4).Width=   2461
      Columns(4).Caption=   "MC Convencional"
      Columns(4).Name =   "MC Convencional"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   6
      Columns(4).NumberFormat=   "CURRENCY"
      Columns(4).FieldLen=   256
      Columns(5).Width=   2461
      Columns(5).Caption=   "MC Direccionado"
      Columns(5).Name =   "MC Direccionado"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   6
      Columns(5).NumberFormat=   "CURRENCY"
      Columns(5).FieldLen=   256
      _ExtentX        =   19076
      _ExtentY        =   4683
      _StockProps     =   79
      Caption         =   "Mantenimiento"
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
   Begin VB.ComboBox cmbPeriodo 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdCalcula 
      Caption         =   "Calcula"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtClave 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCotiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCalcula_Click()
    
    If Me.txtClave.Text = vbNullString Then
        Me.txtClave.SetFocus
        Exit Sub
    End If

    Actualiza
    
    
End Sub

Private Sub Form_Activate()
'    If Me.adodcCalcula.RecordSource = vbNullString Then
'
'
'        'Me.adodcCalcula.CommandType = adCmdText
'        Me.adodcCalcula.CursorLocation = adUseClient
'        Me.adodcCalcula.ConnectionString = Conn
'        'Me.adodcCalcula.LockType = adLockReadOnly
'        'Me.adodcCalcula.Mode = adModeRead
'        'Me.adodcCalcula.EOFAction = adStayEOF
'
'        Me.adodcCalcula.CursorType = adOpenKeyset
'
'    End If
    
    
    CentraForma MDIPrincipal, Me
    
End Sub

Private Sub Actualiza()

    Dim doTotalCon As Double
    Dim doTotalDir As Double
    Dim doTotalMCCon As Double
    Dim doTotalMCDir As Double
    
    Dim sStrSql As String
    
        
        #If SqlServer_ Then
            sStrSql = "SELECT USUARIOS_CLUB.Nombre + ' ' +  USUARIOS_CLUB.A_Paterno + ' ' +  USUARIOS_CLUB.A_Materno AS Nombre, TIPO_USUARIO.Descripcion, HISTORICO_CUOTAS.Monto As MontoConvencional, HISTORICO_CUOTAS.MontoDescuento AS MontoDireccionado, MC.Monto AS MC_MontoConvencional, MC.MontoDescuento AS MC_MontoDireccionado, USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.NumeroFamiliar"
            sStrSql = sStrSql & " FROM ((USUARIOS_CLUB INNER JOIN HISTORICO_CUOTAS ON USUARIOS_CLUB.IdTipoUsuario = HISTORICO_CUOTAS.IdTipoUsuario) INNER JOIN TIPO_USUARIO ON USUARIOS_CLUB.IdTipoUsuario = TIPO_USUARIO.IdTipoUsuario) LEFT JOIN (SELECT MC_EQUIVALE.IdTipoUsuario, HISTORICO_CUOTAS.Monto, HISTORICO_CUOTAS.MontoDescuento FROM MC_EQUIVALE INNER JOIN HISTORICO_CUOTAS ON MC_EQUIVALE.idTipoUsuarioMC = HISTORICO_CUOTAS.IdTipoUsuario WHERE (((HISTORICO_CUOTAS.VigenteDesde)<=GetDate()) AND ((HISTORICO_CUOTAS.VigenteHasta)>=GetDate()) AND ((HISTORICO_CUOTAS.Periodo)=" & Me.cmbPeriodo.ItemData(Me.cmbPeriodo.ListIndex) & "))) AS MC ON USUARIOS_CLUB.IdTipoUsuario = MC.IdTipoUsuario"
            sStrSql = sStrSql & " WHERE (((USUARIOS_CLUB.NoFamilia)=" & Trim(Me.txtClave.Text) & ")"
            sStrSql = sStrSql & " AND ((HISTORICO_CUOTAS.Periodo)=" & Me.cmbPeriodo.ItemData(Me.cmbPeriodo.ListIndex) & ")"
            sStrSql = sStrSql & " AND ((HISTORICO_CUOTAS.VigenteDesde) <= GetDate()) And ((HISTORICO_CUOTAS.VigenteHasta) >= GetDate()))"
            sStrSql = sStrSql & " ORDER BY USUARIOS_CLUB.NumeroFamiliar"
        #Else
            sStrSql = "SELECT USUARIOS_CLUB.Nombre & ' ' &  USUARIOS_CLUB.A_Paterno & ' ' &  USUARIOS_CLUB.A_Materno AS Nombre, TIPO_USUARIO.Descripcion, HISTORICO_CUOTAS.Monto As MontoConvencional, HISTORICO_CUOTAS.MontoDescuento AS MontoDireccionado, MC.Monto AS MC_MontoConvencional, MC.MontoDescuento AS MC_MontoDireccionado, USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.NumeroFamiliar"
            sStrSql = sStrSql & " FROM ((USUARIOS_CLUB INNER JOIN HISTORICO_CUOTAS ON USUARIOS_CLUB.IdTipoUsuario = HISTORICO_CUOTAS.IdTipoUsuario) INNER JOIN TIPO_USUARIO ON USUARIOS_CLUB.IdTipoUsuario = TIPO_USUARIO.IdTipoUsuario) LEFT JOIN (SELECT MC_EQUIVALE.IdTipoUsuario, HISTORICO_CUOTAS.Monto, HISTORICO_CUOTAS.MontoDescuento FROM MC_EQUIVALE INNER JOIN HISTORICO_CUOTAS ON MC_EQUIVALE.idTipoUsuarioMC = HISTORICO_CUOTAS.IdTipoUsuario WHERE (((HISTORICO_CUOTAS.VigenteDesde)<=Date()) AND ((HISTORICO_CUOTAS.VigenteHasta)>=Date()) AND ((HISTORICO_CUOTAS.Periodo)=" & Me.cmbPeriodo.ItemData(Me.cmbPeriodo.ListIndex) & "))) AS MC ON USUARIOS_CLUB.IdTipoUsuario = MC.IdTipoUsuario"
            sStrSql = sStrSql & " WHERE (((USUARIOS_CLUB.NoFamilia)=" & Trim(Me.txtClave.Text) & ")"
            sStrSql = sStrSql & " AND ((HISTORICO_CUOTAS.Periodo)=" & Me.cmbPeriodo.ItemData(Me.cmbPeriodo.ListIndex) & ")"
            sStrSql = sStrSql & " AND ((HISTORICO_CUOTAS.VigenteDesde) <= Date()) And ((HISTORICO_CUOTAS.VigenteHasta) >= Date()))"
            sStrSql = sStrSql & " ORDER BY USUARIOS_CLUB.NumeroFamiliar"
        #End If
        
        LlenaSsDbGrid ssdbgCalcula, Conn, sStrSql, 6
        
        Dim rsMontos As ADODB.Recordset
        Set rsMontos = New ADODB.Recordset
        rsMontos.ActiveConnection = Conn
        rsMontos.LockType = adLockReadOnly
        rsMontos.CursorType = adOpenStatic
        rsMontos.CursorLocation = adUseServer
        rsMontos.Open sStrSql
        
        If Not rsMontos.EOF Then
            rsMontos.MoveFirst
            
            Do Until rsMontos.EOF
                doTotalCon = doTotalCon + rsMontos.Fields("MontoConvencional")
                doTotalDir = doTotalDir + rsMontos.Fields("MontoDireccionado")
                doTotalMCCon = doTotalMCCon + IIf(IsNull(rsMontos.Fields("MC_MontoConvencional")), 0, rsMontos.Fields("MC_MontoConvencional"))
                doTotalMCDir = doTotalMCDir + IIf(IsNull(rsMontos.Fields("MC_MontoDireccionado")), 0, rsMontos.Fields("MC_MontoDireccionado"))
                rsMontos.MoveNext
            Loop
        End If
        
        rsMontos.Close
        Set rsMontos = Nothing
        
        Me.lblCtl(0).Caption = Format(doTotalCon, "$#,0.00")
        Me.lblCtl(1).Caption = Format(doTotalDir, "$#,0.00")
        Me.lblCtl(2).Caption = Format(doTotalMCCon, "$#,0.00")
        Me.lblCtl(3).Caption = Format(doTotalMCDir, "$#,0.00")
        
End Sub

Private Sub Form_Load()
    
    CentraForma MDIPrincipal, Me
    
    Me.cmbPeriodo.AddItem "MENSUAL"
    Me.cmbPeriodo.ItemData(Me.cmbPeriodo.NewIndex) = 1
    
    Me.cmbPeriodo.AddItem "ANUAL"
    Me.cmbPeriodo.ItemData(Me.cmbPeriodo.NewIndex) = 12
    
    
    Me.cmbPeriodo.ListIndex = 0
    
    
    Me.lblCtl(0).Caption = Format(0, "$#,0.00")
    Me.lblCtl(1).Caption = Format(0, "$#,0.00")
    Me.lblCtl(2).Caption = Format(0, "$#,0.00")
    Me.lblCtl(3).Caption = Format(0, "$#,0.00")
    
    
End Sub

