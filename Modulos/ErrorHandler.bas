Attribute VB_Name = "ErrorHandler"
Public Sub MsgError()
    MsgBox "¡Ocurrio un error!" & vbLf _
        & "Número: " & Err.Number & vbLf _
        & "Descripción: " & Err.Description _
        , vbCritical, "Error"
        
End Sub

Public Sub MsgADOError(adoConn As ADODB.Connection)
    
    Dim eError As Error


    

End Sub
