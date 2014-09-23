VERSION 5.00
Begin VB.Form frmAccListaSQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de acceso desde lista"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSecs 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "90"
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ejecutar"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtSQL 
      Height          =   1455
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label lblAvance 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   5175
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Tiempo de espera en secs."
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Consulta SQL a ejecutar"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmAccListaSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOk_Click()
        
        
    If Me.txtSQL.Text = vbNullString Then
        Exit Sub
    End If
    
    
    On Error GoTo Error_Catch
    
    strSQL = Trim(Me.txtSQL.Text)
    Dim adorcs As ADODB.Recordset
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenStatic, adLockReadOnly
    
        
    Dim iCont As Long
    Dim nErrCode As Long
    Dim sTor As String
    
    iCont = 0
    
    Dim lContRec As Long
    lContRec = 0
    
    Do While Not adorcs.EOF
        
        #If SqlServer_ Then
            ActivaCredSQL 1, adorcs!Secuencial, 1, adorcs!IdMember, False, False
        #Else
            ActivaCred 1, adorcs!Secuencial, 1, adorcs!IdMember, False, False
        #End If
        'Registro en Torniquetes NITEGEN
        'sTor = Mid(adorcs!Secuencial, 1, 2)
        'sTor = sTor & adorcs!IdMember
             'nErrCode = EliminarAcceso(sTor)
             
             'If nErrCode <> 0 Then
             '   MsgBox "No se pudo registrar el usuario en torniquetes,Favor de hacerlo manual"
             'End If
        iCont = iCont + 1
        lContRec = lContRec + 1
        
        If iCont > 9 Then
            Me.lblAvance = "Esperando..."
            DoEvents
            Esperar (Val(Me.txtSecs))
            iCont = 0
        End If
        
        Me.lblAvance.Caption = "Procesados " & lContRec & " de " & adorcs.RecordCount & " registros..."
        DoEvents
        adorcs.MoveNext
    Loop
    
    
    adorcs.Close
    Set adorcs = Nothing
    
    Me.lblAvance.Caption = "Terminado"
    
    Exit Sub
    
Error_Catch:

    MsgError
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
