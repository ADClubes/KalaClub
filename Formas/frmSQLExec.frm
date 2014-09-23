VERSION 5.00
Begin VB.Form frmSQLExec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ejecutar Query"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "Ejecutar"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtQry 
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmSQLExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sArchivo As String
Private Sub cmdEjecutar_Click()
    
    Dim adocmd As ADODB.Command
    Dim lRecs As Long
    Dim nTrans As Integer
    
    If Me.txtQry.Text = vbNullString Then
        MsgBox "No hay query para ejecutar", vbExclamation, "Error"
        Exit Sub
    End If
    
    If MsgBox("¿Ejecutar query?", vbOKCancel + vbQuestion, "Confirme") = vbCancel Then
        Exit Sub
    End If
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = Me.txtQry.Text
    
    Err.Clear
    Conn.Errors.Clear
    
    On Error GoTo CATCH_ERROR
    nTrans = Conn.BeginTrans
    
    adocmd.Execute lRecs
    
    Conn.CommitTrans
    
    Set adocmd = Nothing
    
    MsgBox "Se afectaron " & lRecs & " registro(s)", vbInformation, "Ok"
    
    On Error GoTo 0
   
    Exit Sub
   
CATCH_ERROR:

    If nTrans > 0 Then
        Conn.RollbackTrans
    End If

    MsgError
    
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Dim fs As FileSystemObject
    Dim tsFile As TextStream
    
    CentraForma MDIPrincipal, Me
    
    sArchivo = sDB_DataSource & "\sqlqry.sql"
    
    
    Set fs = New FileSystemObject
    
    If fs.FileExists(sArchivo) Then
    
        Set tsFile = fs.OpenTextFile(sArchivo, ForReading)
    
        Me.txtQry.Text = tsFile.ReadAll
    
        tsFile.Close
    End If
    
    Set fs = Nothing
    
    
End Sub
