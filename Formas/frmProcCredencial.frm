VERSION 5.00
Begin VB.Form frmProcCredencial 
   Caption         =   "Reasigna Código"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Proceder"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "frmProcCredencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcesar_Click()
Dim adorstCod As ADODB.Recordset
    Dim adocmdCod As ADODB.Command
    
    
    Dim lCodigoBase As Long
    
    
    Err.Clear
    Conn.Errors.Clear
    
    On Error GoTo ERROR_CATCH
    
    Me.cmdProcesar.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
    
    lCodigoBase = 35000
    
    
    Set adocmdCod = New ADODB.Command
    adocmdCod.ActiveConnection = Conn
    adocmdCod.CommandType = adCmdText

    
    
    strSQL = "UPDATE SECUENCIAL SET"
    strSQL = strSQL & " IdMember=0"
    
    adocmdCod.CommandText = strSQL
    adocmdCod.Execute
    
    strSQL = "SELECT IdMember"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " ORDER BY IdMember"
    
    
    Set adorstCod = New ADODB.Recordset
    adorstCod.CursorLocation = adUseServer
    
    adorstCod.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Do Until adorstCod.EOF
        
        strSQL = "UPDATE Secuencial SET"
        strSQL = strSQL & " IdMember=" & adorstCod!IdMember
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " Secuencial=" & adorstCod!IdMember + lCodigoBase
        
        adocmdCod.CommandText = strSQL
        adocmdCod.Execute
        
        
        adorstCod.MoveNext
    Loop
    
    
    Set adocmdCod = Nothing
    
    adorstCod.Close
    Set adorstCod = Nothing
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Proceso Terminado", vbInformation, "Mensaje"
    
    Unload Me
    
    Exit Sub
    
ERROR_CATCH:

    Dim sMensaje As String
    Dim lI As Long
    
    
    If Err.Number <> 0 Then
        sMensaje = "Error: " & Err.Number & vbLf & Err.Description
    End If
    
    Screen.MousePointer = vbDefault
    
    
    MsgBox sMensaje, vbCritical, "Error"
    
End Sub

