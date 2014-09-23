VERSION 5.00
Begin VB.Form frmUtil 
   Caption         =   "Utileria"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "frmUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Dim adorcs As ADODB.Recordset
    Dim adocmd As ADODB.Command
    
    Dim dfecha As Date
    
    strSQL = "SELECT IdMember, Fecha"
    strSQL = strSQL & " FROM Consulta5"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    
        
        
        
        
    Do Until adorcs.EOF
    
        dfecha = UltimoDiaDelMes(adorcs!Fecha)
        
        strSQL = "UPDATE FECHAS_USUARIO INNER JOIN USUARIOS_CLUB"
        strSQL = strSQL & " ON FECHAS_USUARIO.IdMember=USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " SET FECHAS_USUARIO.FechaUltimoPago='" & Format(dfecha, "dd/mm/yyyy") & "'"
        strSQL = strSQL & " WHERE USUARIOS_CLUB.IdTitular=" & adorcs!IdMember
        
        adocmd.CommandText = strSQL
        adocmd.Execute
            
        adorcs.MoveNext
    Loop
    
    adorcs.Close
    
    Set adorcs = Nothing
    Set adocmd = Nothing
    
    
    
End Sub
