Attribute VB_Name = "Preventa"
Public Function ChecaStatusPreventa(lIdProspecto As Long) As Byte
    Dim adorcs As ADODB.Recordset
    
    ChecaStatusPreventa = 128
    
    
    
    strSQL = "SELECT StatusProspecto"
    strSQL = strSQL & " FROM PROSPECTOS"
    strSQL = strSQL & " WHERE IdProspecto=" & lIdProspecto
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    On Error GoTo Error_Catch
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        ChecaStatusPreventa = adorcs!StatusProspecto
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    
    On Error GoTo 0
    
    Exit Function
    
Error_Catch:

    

    If adorcs.State Then
        adorcs.Close
        Set adorcs = Nothing
    End If
    
    MsgError
        

    
    
End Function


Public Function GetTipoMantenimiento(byTipo As Byte) As String
    
    GetTipoMantenimiento = ""
    
    Select Case byTipo
        Case 1
            GetTipoMantenimiento = "MENSUAL DIRECCIONADO"
        Case 2
            GetTipoMantenimiento = "MENSUAL CONVENCIONAL"
        Case 3
            GetTipoMantenimiento = "ANUAL"
    End Select
    
    
    
End Function
