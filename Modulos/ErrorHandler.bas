Attribute VB_Name = "ErrorHandler"
Public Sub MsgError()
    MsgBox "�Ocurrio un error!" & vbLf _
        & "N�mero: " & Err.Number & vbLf _
        & "Descripci�n: " & Err.Description _
        , vbCritical, "Error"
        
End Sub

Public Sub MsgADOError(adoConn As ADODB.Connection)
    
    Dim eError As Error


    

End Sub
