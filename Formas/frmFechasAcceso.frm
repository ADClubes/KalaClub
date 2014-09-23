VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFechasAcceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza Fechas de Acceso"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar pbProcesoAct 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker dtpFechaProceso 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   59113473
      CurrentDate     =   38672
   End
   Begin VB.Label Label3 
      Caption         =   "Proceso actual"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Proceso Total"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha de proceso:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmFechasAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim adoRcsAcceso As ADODB.Recordset
    Dim adoCmdAcceso As ADODB.Command
    
    
    Me.cmdOk.Enabled = False
    Me.cmdCancelar.Caption = "Detener"
    
    Set adoCmdAcceso = New ADODB.Command
    adoCmdAcceso.ActiveConnection = Conn
    adoCmdAcceso.CommandType = adCmdText
    
    #If SqlServer_ Then
        strSQL = "DELETE FROM ACCESO_DERECHOS"
    #Else
        strSQL = "DELETE * FROM ACCESO_DERECHOS"
    #End If
        
    adoCmdAcceso.CommandText = strSQL
    adoCmdAcceso.Execute
    
    'Todos los usuarios
    #If SqlServer_ Then
        strSQL = "INSERT INTO ACCESO_DERECHOS (IdMember, FechaAccesoPermitido)"
        strSQL = strSQL & " SELECT USUARIOS_CLUB.IdMember, DATEADD(day,0,DATEADD(month,1,Max(FECHAS_USUARIO.FechaUltimoPago))) AS FechaAcceso"
        strSQL = strSQL & " FROM USUARIOS_CLUB LEFT JOIN FECHAS_USUARIO ON USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember"
        strqsl = strSQL & " WHERE USUARIOS_CLUB.STATUS IS NULL"
        strSQL = strSQL & " GROUP BY USUARIOS_CLUB.IdMember"
    #Else
        strSQL = "INSERT INTO ACCESO_DERECHOS (IdMember, FechaAccesoPermitido)"
        strSQL = strSQL & " SELECT USUARIOS_CLUB.IdMember, dateadd('d',0,dateadd('m',1,Max(FECHAS_USUARIO.FechaUltimoPago))) AS FechaAcceso"
        strSQL = strSQL & " FROM USUARIOS_CLUB LEFT JOIN FECHAS_USUARIO ON USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember"
        strqsl = strSQL & " WHERE USUARIOS_CLUB.STATUS IS NULL"
        strSQL = strSQL & " GROUP BY USUARIOS_CLUB.IdMember"
    #End If
    
    adoCmdAcceso.CommandText = strSQL
    adoCmdAcceso.Execute
    
    'Usuarios de Staff
    #If SqlServer_ Then
        strSQL = "UPDATE ACCESO_DERECHOS SET"
        strSQL = strSQL & " FechaAccesoPermitido='21001231'"
        strSQL = strSQL & " FROM ACCESO_DERECHOS INNER JOIN USUARIOS_CLUB ON ACCESO_DERECHOS.IdMember=USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " USUARIOS_CLUB.IdTipoAcceso=2"
    #Else
        strSQL = "UPDATE ACCESO_DERECHOS INNER JOIN USUARIOS_CLUB ON ACCESO_DERECHOS.IdMember=USUARIOS_CLUB.IdMember SET"
        strSQL = strSQL & " FechaAccesoPermitido='12/31/2100'"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " USUARIOS_CLUB.IdTipoAcceso=2"
    #End If
    
    adoCmdAcceso.CommandText = strSQL
    adoCmdAcceso.Execute
    
    'Ausencias
    #If SqlServer_ Then
        strSQL = "UPDATE ACCESO_DERECHOS SET"
        strSQL = strSQL & " FechaAccesoPermitido=AUSENCIAS.FechaInicial"
        strSQL = strSQL & " FROM ACCESO_DERECHOS INNER JOIN AUSENCIAS ON ACCESO_DERECHOS.IdMember=AUSENCIAS.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " AUSENCIAS.FechaInicial <= '" & Format(Me.dtpFechaProceso.Value, "yyyymmdd") & "'"
    #Else
        strSQL = "UPDATE ACCESO_DERECHOS INNER JOIN AUSENCIAS ON ACCESO_DERECHOS.IdMember=AUSENCIAS.IdMember SET"
        strSQL = strSQL & " FechaAccesoPermitido=AUSENCIAS.FechaInicial"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " AUSENCIAS.FechaInicial <= #" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#"
        'strSQL = strSQL & " AND AUSENCIAS.FechaFinal >= #" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#"
    #End If
    
    adoCmdAcceso.CommandText = strSQL
    adoCmdAcceso.Execute
    
    Set adoCmdAcceso = Nothing
    
    Me.cmdCancelar.Caption = "Salir"
    
    MsgBox "Proceso concluido.", vbExclamation, "Mensaje"
    
End Sub

Private Sub Form_Load()

    Me.dtpFechaProceso.Value = Date
    
    CentraForma MDIPrincipal, Me
End Sub
