Attribute VB_Name = "Parametros"
Option Explicit
Public Function ObtieneParametro(sNombreParametro As String) As String
    
    Dim adorsParam As ADODB.Recordset
    
    ObtieneParametro = vbNullString
    
    #If SQLServer_ Then
        strSQL = "SELECT Nombre_Parametro, Valor"
        strSQL = strSQL & " FROM PARAMETROS"
        strSQL = strSQL & " WHERE UPPER(Nombre_Parametro)='" & UCase(Trim(sNombreParametro)) & "'"
    #Else
        strSQL = "SELECT Nombre_Parametro, Valor"
        strSQL = strSQL & " FROM PARAMETROS"
        strSQL = strSQL & " WHERE UCase(Nombre_Parametro)='" & UCase(Trim(sNombreParametro)) & "'"
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
