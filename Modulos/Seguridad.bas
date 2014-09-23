Attribute VB_Name = "Seguridad"
Option Explicit
Public Sub Habilita_Seguridad(oObject As Object)
    Dim lI As Long
    Dim lJ As Long
    
    
    oObject.Controls("Toolbar1").Enabled = True
    
    
    For lI = 0 To oObject.Controls.Count - 1
        If TypeOf oObject.Controls(lI) Is Toolbar Then
            For lJ = 1 To oObject.Controls(lI).Buttons.Count
                Debug.Print oObject.Controls(lI).Buttons(lJ).Key
            Next
        End If
    Next
    
End Sub


Public Function ChecaSeguridad(sFrmName As String, sObjectName As String) As Boolean
    Dim adorcsSeg As ADODB.Recordset
    Dim bFlag As Boolean
    Dim sMensajePrev As String
    
    sMensajePrev = vbNullString
    
    ChecaSeguridad = False
    bFlag = False
    
    If Val(sDB_NivelUser) = 0 Then
        ChecaSeguridad = True
        bFlag = True
        Exit Function
    End If
    
    #If SqlServer_ Then
        strSQL = "SELECT SEGURIDAD_USUARIO.IdUsuario"
        strSQL = strSQL & " FROM (SEGURIDAD_USUARIO INNER JOIN USUARIOS_SISTEMA ON SEGURIDAD_USUARIO.IdUsuario = USUARIOS_SISTEMA.IdUsuario)"
        strSQL = strSQL & " INNER JOIN CT_Objetos_Seguridad ON SEGURIDAD_USUARIO.IdObjeto = CT_Objetos_Seguridad.IdObjeto"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((SEGURIDAD_USUARIO.IdUsuario)=" & iDB_IdUser & ")"
        strSQL = strSQL & " AND (Upper(CT_Objetos_Seguridad.FormaNombre)='" & Trim(UCase(sFrmName)) & "')"
        strSQL = strSQL & " AND (Upper(CT_Objetos_Seguridad.ObjetoNombre)='" & Trim(UCase(sObjectName)) & "')"
        strSQL = strSQL & ")"
    #Else
        strSQL = "SELECT SEGURIDAD_USUARIO.IdUsuario"
        strSQL = strSQL & " FROM (SEGURIDAD_USUARIO INNER JOIN USUARIOS_SISTEMA ON SEGURIDAD_USUARIO.IdUsuario = USUARIOS_SISTEMA.IdUsuario)"
        strSQL = strSQL & " INNER JOIN CT_Objetos_Seguridad ON SEGURIDAD_USUARIO.IdObjeto = CT_Objetos_Seguridad.IdObjeto"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & "((SEGURIDAD_USUARIO.IdUsuario)=" & iDB_IdUser & ")"
        strSQL = strSQL & " AND (UCase(CT_Objetos_Seguridad.FormaNombre)='" & Trim(UCase(sFrmName)) & "')"
        strSQL = strSQL & " AND (UCase(CT_Objetos_Seguridad.ObjetoNombre)='" & Trim(UCase(sObjectName)) & "')"
        strSQL = strSQL & ")"
    #End If
    
    Set adorcsSeg = New ADODB.Recordset
    adorcsSeg.CursorLocation = adUseServer
    
    adorcsSeg.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsSeg.EOF Then
        ChecaSeguridad = True
        bFlag = True
    End If
    
    adorcsSeg.Close
    Set adorcsSeg = Nothing
    
    If Not bFlag Then
        sMensajePrev = MDIPrincipal.StatusBar1.Panels(1).Text
        MDIPrincipal.StatusBar1.Panels(7).Text = "Seguridad: " & sFrmName & "," & sObjectName
        MsgBox "¡Nivel de seguridad insuficiente!", vbCritical + vbOKOnly, "Error"
        'MDIPrincipal.StatusBar1.Panels(1).Text = sMensajePrev
    End If
End Function
