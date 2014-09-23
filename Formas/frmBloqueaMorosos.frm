VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBloqueaMorosos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloqueo de usuarios con adeudo"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optCtrl 
      Caption         =   "Activa"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton optCtrl 
      Caption         =   "Inactiva"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtSecs 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Text            =   "90"
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Proceder"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpFechaProc 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16842753
      CurrentDate     =   38854
   End
   Begin VB.Label lblAvance 
      Caption         =   "Label3"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Sec. en espera"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblMensaje 
      Caption         =   "Label2"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha ultimo pago"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
End
Attribute VB_Name = "frmBloqueaMorosos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 
 Dim adorcsBlock As ADODB.Recordset

Private Sub cmdCancelar_Click()
    
    
    If adorcsBlock.State > 0 Then
        adorcsBlock.Close
    End If
    
    Set adorcsBlock = Nothing
    Unload Me
    
End Sub

Private Sub cmdOk_Click()
   
    Dim iCont As Integer
    Dim lMaxRec As Long
    Dim lContRec As Long
    
    
    Me.cmdOk.Enabled = False
    
    
    #If SqlServer_ Then
        If Me.cmbTipo.ListIndex = 0 Then
            strSQL = "SELECT USUARIOS_CLUB.NoFamilia,  USUARIOS_CLUB.IdMember, FECHAS_USUARIO.FechaUltimoPago, SECUENCIAL.Secuencial"
            strSQL = strSQL & " FROM (FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember) INNER JOIN SECUENCIAL ON USUARIOS_CLUB.IdMember=SECUENCIAL.IdMember"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " FECHAS_USUARIO.FechaUltimoPago ='" & Format(Me.dtpFechaProc.Value, "yyyymmdd") & "'"
            strSQL = strSQL & " AND USUARIOS_CLUB.IdTipoAcceso = 0"
            strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NoFamilia"
        ElseIf Me.cmbTipo.ListIndex = 1 Then
            strSQL = "SELECT USUARIOS_CLUB.NoFamilia,  USUARIOS_CLUB.IdMember, FECHAS_USUARIO.FechaUltimoPago, SECUENCIAL.Secuencial"
            strSQL = strSQL & " FROM (FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember) INNER JOIN SECUENCIAL ON USUARIOS_CLUB.IdMember=SECUENCIAL.IdMember"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " FECHAS_USUARIO.FechaUltimoPago ='" & Format(Me.dtpFechaProc.Value, "yyyymmdd") & "'"
            strSQL = strSQL & " AND USUARIOS_CLUB.IdTipoAcceso = 0"
            strSQL = strSQL & " AND USUARIOS_CLUB.IdTitular Not In (SELECT IdMember From Direccionados Where Activo=-1)"
            strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NoFamilia"
        ElseIf Me.cmbTipo.ListIndex = 2 Then
            strSQL = "SELECT USUARIOS_CLUB.NoFamilia,  USUARIOS_CLUB.IdMember, FECHAS_USUARIO.FechaUltimoPago, SECUENCIAL.Secuencial"
            strSQL = strSQL & " FROM (FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember) INNER JOIN SECUENCIAL ON USUARIOS_CLUB.IdMember=SECUENCIAL.IdMember"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " FECHAS_USUARIO.FechaUltimoPago ='" & Format(Me.dtpFechaProc.Value, "yyyymmdd") & "'"
            strSQL = strSQL & " AND USUARIOS_CLUB.IdTipoAcceso = 0"
            strSQL = strSQL & " AND USUARIOS_CLUB.NoFamilia In (SELECT NoFamilia From ListaAcceso)"
            strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NoFamilia"
        ElseIf Me.cmbTipo.ListIndex = 3 Then
            strSQL = "SELECT 1 NoFamilia, 1 IdMember, DATEADD(DAY,-1,DATEADD(MONTH,1,GETDATE())) FechaUltimoPago, '00052' + Secuencial AS Secuencial"
            strSQL = strSQL & " FROM LISTA_CREDENCIALES"
        End If
    #Else
        If Me.cmbTipo.ListIndex = 0 Then
            strSQL = "SELECT USUARIOS_CLUB.NoFamilia,  USUARIOS_CLUB.IdMember, FECHAS_USUARIO.FechaUltimoPago, SECUENCIAL.Secuencial"
            strSQL = strSQL & " FROM (FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember) INNER JOIN SECUENCIAL ON USUARIOS_CLUB.IdMember=SECUENCIAL.IdMember"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " FECHAS_USUARIO.FechaUltimoPago =#" & Format(Me.dtpFechaProc.Value, "mm/dd/yyyy") & "#"
            strSQL = strSQL & " AND USUARIOS_CLUB.IdTipoAcceso = 0"
            strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NoFamilia"
        ElseIf Me.cmbTipo.ListIndex = 1 Then
            strSQL = "SELECT USUARIOS_CLUB.NoFamilia,  USUARIOS_CLUB.IdMember, FECHAS_USUARIO.FechaUltimoPago, SECUENCIAL.Secuencial"
            strSQL = strSQL & " FROM (FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember) INNER JOIN SECUENCIAL ON USUARIOS_CLUB.IdMember=SECUENCIAL.IdMember"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " FECHAS_USUARIO.FechaUltimoPago =#" & Format(Me.dtpFechaProc.Value, "mm/dd/yyyy") & "#"
            strSQL = strSQL & " AND USUARIOS_CLUB.IdTipoAcceso = 0"
            strSQL = strSQL & " AND USUARIOS_CLUB.IdTitular Not In (SELECT IdMember From Direccionados Where Activo=-1)"
            strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NoFamilia"
        ElseIf Me.cmbTipo.ListIndex = 2 Then
            strSQL = "SELECT USUARIOS_CLUB.NoFamilia,  USUARIOS_CLUB.IdMember, FECHAS_USUARIO.FechaUltimoPago, SECUENCIAL.Secuencial"
            strSQL = strSQL & " FROM (FECHAS_USUARIO INNER JOIN USUARIOS_CLUB ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember) INNER JOIN SECUENCIAL ON USUARIOS_CLUB.IdMember=SECUENCIAL.IdMember"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " FECHAS_USUARIO.FechaUltimoPago =#" & Format(Me.dtpFechaProc.Value, "mm/dd/yyyy") & "#"
            strSQL = strSQL & " AND USUARIOS_CLUB.IdTipoAcceso = 0"
            strSQL = strSQL & " AND USUARIOS_CLUB.NoFamilia In (SELECT NoFamilia From ListaAcceso)"
            strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NoFamilia"
        ElseIf Me.cmbTipo.ListIndex = 3 Then
            strSQL = "SELECT 1 NoFamilia, 1 IdMember, DATEADD(DAY,-1,DATEADD(MONTH,1,Date)) FechaUltimoPago, '00052' + Secuencial AS Secuencial"
            strSQL = strSQL & " FROM LISTA_CREDENCIALES"
        End If
    #End If
    
    Set adorcsBlock = New ADODB.Recordset
    adorcsBlock.CursorLocation = adUseServer
    
    adorcsBlock.Open strSQL, Conn, adOpenStatic, adLockReadOnly
    
    lMaxRec = adorcsBlock.RecordCount
    lContRec = 1
    Do Until adorcsBlock.EOF
    
        Me.lblMensaje.Caption = "Procesando " & lContRec & " de " & lMaxRec
        DoEvents
    
        If Me.optCtrl(0).Value Then
            #If SqlServer_ Then
                If Me.cmbTipo.ListIndex = 3 Then
                    ActivaCred2SQL 1, adorcsBlock!Secuencial, 1, adorcsBlock!Idmember, False, False
                Else
                    ActivaCredSQL 1, adorcsBlock!Secuencial, 1, adorcsBlock!Idmember, False, False
                End If
            #Else
                If Me.cmbTipo.ListIndex = 3 Then
                    ActivaCred2 1, adorcsBlock!Secuencial, 1, adorcsBlock!Idmember, False, False
                Else
                    ActivaCred 1, adorcsBlock!Secuencial, 1, adorcsBlock!Idmember, False, False
                End If
            #End If
        Else
            #If SqlServer_ Then
                If Me.cmbTipo.ListIndex = 3 Then
                    ActivaCred2SQL 1, adorcsBlock!Secuencial, 1, adorcsBlock!Idmember, True, False
                Else
                    ActivaCredSQL 1, adorcsBlock!Secuencial, 1, adorcsBlock!Idmember, True, False
                End If
            #Else
                If Me.cmbTipo.ListIndex = 3 Then
                    ActivaCred2 1, adorcsBlock!Secuencial, 1, adorcsBlock!Idmember, True, False
                Else
                    ActivaCred 1, adorcsBlock!Secuencial, 1, adorcsBlock!Idmember, True, False
                End If
            #End If
        End If
        
        adorcsBlock.MoveNext
        iCont = iCont + 1
        lContRec = lContRec + 1
        
        If iCont > 9 Then
            Me.lblMensaje.Caption = "Esperando..."
            Espera (Val(Me.txtSecs))
            Me.lblAvance.Caption = "Procesando " & lContRec & " de " & lMaxRec
            DoEvents
            iCont = 0
        End If
        
    Loop
    
    adorcsBlock.Close
    Set adorcsBlock = Nothing
    
End Sub

Private Sub Espera(lSecs)
    Dim lInicio As Long
    
    lInicio = Timer
    
    Do Until lInicio + lSecs < Timer
        DoEvents
    Loop
    
End Sub

Private Sub Form_Load()
    Me.lblMensaje.Caption = ""
    Me.lblAvance.Caption = ""
    Me.dtpFechaProc.Value = UltimoDiaDelMes(DateAdd("m", -2, Date))
    
    Me.cmbTipo.AddItem "Todos"
    Me.cmbTipo.AddItem "Convencionales"
    Me.cmbTipo.AddItem "Lista"
    Me.cmbTipo.AddItem "Lista credenciales"
    
    Me.cmbTipo.Text = "Todos"
    Me.cmbTipo.ListIndex = 0
    
    Set adorcsBlock = New ADODB.Recordset
    
End Sub

