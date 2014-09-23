Attribute VB_Name = "FuncionesGlobales"
Public Function AyudaClaveSocio() As Long
    
    Dim frmAyuSoc As frmAyudaClave
    
    Set frmAyuSoc = New frmAyudaClave
    
    frmAyuSoc.Show 1
    
    
End Function

Public Function RemoveLF(sCadena As String) As String
    Dim lI As Long
    Dim cChar As String
    Dim sCadenaFinal As String
    
    sCadenaFinal = ""
    
    For i = 1 To Len(sCadena)
        cChar = Mid$(sCadena, i, 1)
        If Asc(cChar) <> Asc(vbCr) And Asc(cChar) <> Asc(vbLf) Then
            sCadenaFinal = sCadenaFinal & cChar
        End If
    Next
    
    RemoveLF = sCadenaFinal
    
End Function

Public Function ExistWindow(sWindowName As String) As Boolean
    Dim frmForm As Form
    
    ExistWindow = False
    
    For Each frmForm In Forms
        If frmForm.Name = sWindowName Then
            ExistWindow = True
            Exit Function
        End If
    Next
    
    
End Function

Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

Public Function SelectPrinter(sNombreImpresora As String) As String
    Dim CurrentPrinter As Printer
    Dim P As Printer
    Dim iRespuesta As Long
    
    SelectPrinter = Printer.DeviceName
    
    For Each P In Printers
    'iRespuesta = MsgBox(P.DeviceName, vbQuestion + vbOKCancel, "Confirme")
        If P.DeviceName = sNombreImpresora Then
            Set Printer = P
            Exit For
        End If
    Next
    
End Function

Public Sub LlenaComboImpresoras(cmbCtrl As ComboBox, boSelectDefault As Boolean)

    cmbCtrl.Clear
    
    Dim prn As Printer
    
    For Each prn In Printers
        cmbCtrl.AddItem prn.DeviceName
    Next
    
    If boSelectDefault Then
        cmbCtrl.Text = Printer.DeviceName
    End If
    
    
End Sub

Public Function GetTipoUsuarioMulti(lidMember As Long) As String
    
    Dim adorcs As ADODB.Recordset
    
    GetTipoUsuarioMulti = ""
    
    strSQL = "SELECT Descripcion"
    strSQL = strSQL & " FROM USUARIOS_MC"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((Secuencial)=" & lidMember & ")"
    strSQL = strSQL & ")"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        GetTipoUsuarioMulti = adorcs!Descripcion
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
End Function
Public Function GetTipoUsuario(lidMember As Long) As String
    
    Dim adorcs As ADODB.Recordset
    
    GetTipoUsuario = ""
    
    strSQL = "SELECT U.IdTipoUsuario,Descripcion"
    strSQL = strSQL & " FROM USUARIOS_CLUB U INNER JOIN TIPO_USUARIO T ON U.IdTipoUsuario=T.IdTipoUsuario"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((IdMember)=" & lidMember & ")"
    strSQL = strSQL & ")"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        GetTipoUsuario = adorcs!Descripcion
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
End Function

Public Function RfcValido(ByVal pRFC As String) As Boolean
    Dim expresion As RegExp
    Set expresion = New RegExp
    
    expresion.Pattern = "^[a-z&A-Z]{3,4}(\d{6})((\D|\d){3})?$"
    
    RfcValido = expresion.Test(pRFC)
End Function
