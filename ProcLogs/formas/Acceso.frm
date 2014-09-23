VERSION 5.00
Begin VB.Form Acceso 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblProc 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "Acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim sStrSql As String
    
    Dim adorcsinput As ADODB.Recordset
    Dim adorcsproc As ADODB.Recordset
    
    Dim adocmd As ADODB.Command
    
    Me.Command1.Enabled = False
    
    sStrSql = "SELECT * "
    sStrSql = sStrSql & " FROM ACCESO"
    sStrSql = sStrSql & " WHERE"
    sStrSql = sStrSql & " FECHA BETWEEN #6/1/2007# AND #6/30/2007#"
    sStrSql = sStrSql & " AND ENT_SAL=1"
    sStrSql = sStrSql & " AND EXCEPCION=Space(2)"
    sStrSql = sStrSql & " ORDER BY FECHA, HORA"
    
    Set adorcsinput = New ADODB.Recordset
    adorcsinput.CursorLocation = adUseServer
    
    adorcsinput.Open sStrSql, Conn, adOpenForwardOnly, adLockReadOnly
    
    Set adorcsproc = New ADODB.Recordset
    adorcsproc.CursorLocation = adUseServer
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    
    
    Do Until adorcsinput.EOF
    
        Me.lblProc.Caption = adorcsinput!Fecha & vbTab & adorcsinput!Hora
        DoEvents
        
        sStrSql = "SELECT Hora, Marca"
        sStrSql = sStrSql & " FROM ACCESO"
        sStrSql = sStrSql & " WHERE"
        sStrSql = sStrSql & " FECHA=#" & Format(adorcsinput!Fecha, "mm/dd/yyyy") & "#"
        sStrSql = sStrSql & " AND SECUENCIAL=" & adorcsinput!Secuencial
        sStrSql = sStrSql & " AND ENT_SAL=2"
        sStrSql = sStrSql & " AND EXCEPCION=Space(2)"
        sStrSql = sStrSql & " AND MARCA=0"
        sStrSql = sStrSql & " AND HORA >= '" & adorcsinput!Hora & "'"
        sStrSql = sStrSql & " ORDER BY HORA"
        
        adorcsproc.Open sStrSql, Conn, adOpenKeyset, adLockOptimistic
        
        
        
        
        sStrSql = "INSERT INTO ACCESOPROC ("
        sStrSql = sStrSql & " FECHA,"
        sStrSql = sStrSql & " HORAE,"
        sStrSql = sStrSql & " HORAS,"
        sStrSql = sStrSql & " SECUENCIAL,"
        sStrSql = sStrSql & " EXCEPCION)"
        sStrSql = sStrSql & " VALUES ("
        sStrSql = sStrSql & "'" & adorcsinput!Fecha & "',"
        sStrSql = sStrSql & "'" & Format(adorcsinput!Hora, "Hh:Nn") & "',"
        If Not adorcsproc.EOF Then
            sStrSql = sStrSql & "'" & Format(adorcsproc!Hora, "Hh:Nn") & "',"
        Else
            sStrSql = sStrSql & "'" & Format(adorcsinput!Hora, "Hh:Nn") & "',"
        End If
        sStrSql = sStrSql & adorcsinput!Secuencial & ","
        sStrSql = sStrSql & "''" & ")"
        
        adocmd.CommandText = sStrSql
        adocmd.Execute
        
        If Not adorcsproc.EOF Then
            adorcsproc!Marca = -1
            adorcsproc.Update
        End If
        
        
        adorcsproc.Close
        adorcsinput.MoveNext
    Loop
    
    adorcsinput.Close
    
    Set adocmd = Nothing
    Set adorcsproc = Nothing
    Set adorcsinput = Nothing
    
    MsgBox "Proceso Concluido", vbCritical, "Ok"
    
    Me.Command1.Enabled = True
    
End Sub

Private Sub Form_Load()
    If Not Connection_DB() Then
        End
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Close
    Set Conn = Nothing
End Sub
