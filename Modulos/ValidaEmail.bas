Attribute VB_Name = "ValidaEmail"
Option Explicit
Public Function ValidaEmailAddress(sEmail As String) As Boolean

    Dim oRegExp As RegExp
    

    sEmail = Trim(sEmail)
    
    
    Set oRegExp = New RegExp
    
    oRegExp.Pattern = "^[\w][\w\.-]*@[\w][\w\.-]*\.[\w]+$"
    
    oRegExp.Test (sEmail)
    
    ValidaEmailAddress = oRegExp.Test(sEmail)
    
    Set oRegExp = Nothing
    


End Function

Public Function ValidaExpReg(sCadValidar As String, sExpReg As String)

    Dim oRegExp As RegExp
    
    Set oRegExp = New RegExp
    
    oRegExp.Pattern = sExpReg
    
    oRegExp.Test (sCadValidar)
    
    ValidaExpReg = oRegExp.Test(sCadValidar)
    
    Set oRegExp = Nothing


End Function
