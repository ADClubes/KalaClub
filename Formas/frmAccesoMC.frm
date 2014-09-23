VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAccesoMC 
   Caption         =   "Acceso MC"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optCtrl 
      Caption         =   "Inactiva"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.OptionButton optCtrl 
      Caption         =   "Activa"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtTEspera 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "60"
      Top             =   2640
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   62849025
      CurrentDate     =   40007
   End
   Begin VB.CommandButton cmdProcede 
      Caption         =   "Procesar"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblMensaje 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   5175
   End
End
Attribute VB_Name = "frmAccesoMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcede_Click()
    
    
    Dim adorcsMC As ADODB.Recordset
    
    Dim iCont As Integer
    Dim lMaxRec As Long
    Dim lContRec As Long
    
    
    'Me.cmdOk.Enabled = False
    
    strSQL = "SELECT USUARIOS_MC.IdClub, USUARIOS_MC.NoFamilia, USUARIOS_MC.FechaUltimoPago, USUARIOS_MC.Secuencial"
    strSQL = strSQL & " From USUARIOS_MC"
    
    #If SqlServer_ Then
        If Me.optCtrl(0).Value Then
            strSQL = strSQL & " WHERE FechaUltimoPago >= '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "'"
        Else
            strSQL = strSQL & " WHERE FechaUltimoPago < '" & Format(Me.dtpFecha.Value, "yyyymmdd") & "'"
        End If
    #Else
        If Me.optCtrl(0).Value Then
            strSQL = strSQL & " WHERE FechaUltimoPago >= #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#"
        Else
            strSQL = strSQL & " WHERE FechaUltimoPago < #" & Format(Me.dtpFecha.Value, "mm/dd/yyyy") & "#"
        End If
    #End If

    strSQL = strSQL & " ORDER BY USUARIOS_MC.IdClub, USUARIOS_MC.NoFamilia"

    
    Set adorcsMC = New ADODB.Recordset
    adorcsMC.CursorLocation = adUseServer
    
    adorcsMC.Open strSQL, Conn, adOpenStatic, adLockReadOnly
    
    lMaxRec = adorcsMC.RecordCount
    lContRec = 1
    Do Until adorcsMC.EOF
    
        Me.lblMensaje.Caption = "Procesando " & lContRec & " de " & lMaxRec
        DoEvents
    
    
        If adorcsMC!Fechaultimopago < CDate(Me.dtpFecha.Value) Then
            #If SqlServer_ Then
                ActivaCred2SQL 1, adorcsMC!Secuencial, 1, adorcsMC!NoFamilia, False, False
            #Else
                ActivaCred2 1, adorcsMC!Secuencial, 1, adorcsMC!NoFamilia, False, False
            #End If
        Else
            #If SqlServer_ Then
                ActivaCred2SQL 1, adorcsMC!Secuencial, 1, adorcsMC!NoFamilia, True, False
            #Else
                ActivaCred2 1, adorcsMC!Secuencial, 1, adorcsMC!NoFamilia, True, False
            #End If
        End If
        
        
        adorcsMC.MoveNext
        
        
        
        
        iCont = iCont + 1
        lContRec = lContRec + 1
       
       Espera (2)
       
       If lContRec Mod 10 = 0 Then
          Me.lblMensaje.Caption = "Esperando..."
          DoEvents
          Espera (Val(Me.txtTEspera.Text))
       End If
        
    Loop
    
    adorcsMC.Close
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
    Me.dtpFecha.Value = UltimoDiaDelMes(DateAdd("m", -2, Date))
End Sub
