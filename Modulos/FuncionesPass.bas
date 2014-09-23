Attribute VB_Name = "FuncionesPass"
Option Explicit
Public Function CambiaPassword(sUsuario As String, sPassNuevo As String) As Integer
    Dim adocmdPass As ADODB.Command
    Dim lRecords As Long
    
    CambiaPassword = 0
    
    #If SqlServer_ Then
        strSQL = "UPDATE USUARIOS_SISTEMA SET "
        strSQL = strSQL & "uPassword=" & "'" & sPassNuevo & "',"
        strSQL = strSQL & "FechaVencePass=" & "'" & Format(Date + 90, "yyyymmdd") & "'"
        strSQL = strSQL & " Where ("
        strSQL = strSQL & "(Login_Name)=" & "'" & sUsuario & "'"
        strSQL = strSQL & ")"
    #Else
        strSQL = "UPDATE USUARIOS_SISTEMA SET "
        strSQL = strSQL & "uPassword=" & "'" & sPassNuevo & "',"
        strSQL = strSQL & "FechaVencePass=" & "#" & Format(Date + 90, "mm/dd/yyyy") & "#"
        strSQL = strSQL & " Where ("
        strSQL = strSQL & "(Login_Name)=" & "'" & sUsuario & "'"
        strSQL = strSQL & ")"
    #End If
    
    Set adocmdPass = New ADODB.Command
    adocmdPass.ActiveConnection = Conn
    adocmdPass.CommandType = adCmdText
    adocmdPass.CommandText = strSQL
    
    adocmdPass.Execute lRecords
    Set adocmdPass = Nothing
    
    
    If lRecords = 0 Then
        CambiaPassword = 1
    End If
    
    
    Set adocmdPass = Nothing
    
End Function

Public Function ChecaPassword(sUsuario As String, sPassword As String) As Boolean
    Dim adorcsPass As ADODB.Recordset
    
    ChecaPassword = True
    
    strSQL = "SELECT IdUsuario"
    strSQL = strSQL & " FROM USUARIOS_SISTEMA"
    strSQL = strSQL & " Where ("
    strSQL = strSQL & "((Login_Name)=" & "'" & sUsuario & "')"
    strSQL = strSQL & " And ((uPassword)=" & "'" & sPassword & "')"
    strSQL = strSQL & ")"
    
    
    Set adorcsPass = New ADODB.Recordset
    adorcsPass.CursorLocation = adUseServer
    
    adorcsPass.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If adorcsPass.EOF Then
        ChecaPassword = False
    End If
    
    adorcsPass.Close
    Set adorcsPass = Nothing
    
    
End Function
