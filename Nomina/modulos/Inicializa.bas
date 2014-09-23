Attribute VB_Name = "Inicializa"
Option Explicit

Global sDB As String                             '  BASE DE DATOS
Global sDB_DataSource As String          ' RUTA DE LA BASE DE DATOS
Global sDB_PW As String                    ' PASSWORD


Public Conn As ADODB.Connection
Public strSQL As String
Public strConn As String



Public Function Connection_DB() As Boolean
    Dim IntError As Double
    On Error GoTo ErrorCon
    
    sDB_DataSource = App.Path
    sDB = "Nomina.mdb"
    sDB_PW = ""
    
    ' Inicializa Variables de conexion a la base de datos
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & sDB_DataSource & "\" & sDB & ";" & _
              "Persist Security Info=False;" & _
              "Jet OLEDB:Database Password=eUdomilia2006;" & _
              "User Id=Admin;" & _
              "Password=" & sDB_PW & ";"
              
    'SQL 2005
    'strConn = "Provider=SQLNCLI.1;Password=kala1228;Persist Security Info=True;User ID=sa;Initial Catalog=KALACLUB_SQL;Data Source=server"
              

    ' Crea una nueva conexion a la base de datos
    Set Conn = New ADODB.Connection
    
    Conn.Errors.Clear
    Err.Clear
    
    Conn.CursorLocation = adUseServer
    Conn.Open strConn
    
    If Conn.Errors.Count > 0 Then
        Connection_DB = False
    End If

    Connection_DB = True
    
    Exit Function
    
ErrorCon:

    IntError = Conn.Errors.Item(0).NativeError
    Select Case IntError
        Case 18456
            MsgBox "Login o Password Invalido " & Conn.Errors.Item(0).NativeError & Chr(13) & Conn.Errors.Item(0) & Chr(13) & Err.Description, vbCritical, "Conection DataBase"
        Case Else
            MsgBox "Error en Conexión DB: " & Conn.Errors.Item(0).NativeError & Chr(13) & Conn.Errors.Item(0), vbCritical, "Conection DataBase"
    End Select
    
    Connection_DB = False
    
End Function
Public Function ObtieneParametro(sNombreParametro As String) As String
    #Const sql_Access = 1
    
    Dim adorsParam As ADODB.Recordset
    
    ObtieneParametro = vbNullString
    
    #If sql_Access Then
        strSQL = "SELECT Nombre_Parametro, Valor"
        strSQL = strSQL & " FROM PARAMETROS"
        strSQL = strSQL & " WHERE UCase(Nombre_Parametro)='" & UCase(Trim(sNombreParametro)) & "'"
    #Else
        strSQL = "SELECT Nombre_Parametro, Valor"
        strSQL = strSQL & " FROM PARAMETROS"
        strSQL = strSQL & " WHERE UPPER(Nombre_Parametro)='" & UCase(Trim(sNombreParametro)) & "'"
    #End If
    Set adorsParam = New ADODB.Recordset
    
    adorsParam.CursorLocation = adUseServer
    
    adorsParam.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not adorsParam.EOF Then
        ObtieneParametro = Trim(adorsParam!Valor)
    End If
    
    adorsParam.Close
    Set adorsParam = Nothing
    
    
End Function


