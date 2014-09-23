Attribute VB_Name = "ModuleMain"
Option Explicit
Public adorcsSrc As ADODB.Recordset
Public adorcsTgt As ADODB.Recordset

Public ConnSrc As ADODB.Connection
Public ConnTgt As ADODB.Connection

Sub Main()


    Dim sStrConnSrc As String
    Dim sStrConnTgt As String

    Dim sStrSql As String
    
    Dim sPathDBSrc As String
    Dim sPathDBTgt As String
    Dim sDBSrc As String
    Dim sDBTgt As String
    

    sStrConnSrc = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & sPathDBSrc & "\" & sDBSrc & ";" & _
        "Persist Security Info=False;" & _
        "Jet OLEDB:Database Password=eUdomilia2006;" & _
        "User Id=Admin;" & _
        "Password=" & "" & ";"

    sStrConnTgt = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & sPathDBTgt & "\" & sDBTgt & ";" & _
        "Persist Security Info=False;" & _
        "Jet OLEDB:Database Password=eUdomilia2006;" & _
        "User Id=Admin;" & _
        "Password=" & "" & ";"
    



    sStrSql = "SELECT USUARIOS_CLUB.IdMember, USUARIOS_CLUB.NumeroFamiliar, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.FechaNacio, USUARIOS_CLUB.Sexo, USUARIOS_CLUB.IdPais, USUARIOS_CLUB.IdTipoUsuario, USUARIOS_CLUB.IdTitular, USUARIOS_CLUB.Email, USUARIOS_CLUB.Profesion, USUARIOS_CLUB.FechaIngreso, USUARIOS_CLUB.FotoFile, USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Status, USUARIOS_CLUB.Inscripcion, USUARIOS_CLUB.Direccionado, USUARIOS_CLUB.IdTipoAcceso, SECUENCIAL.Secuencial, FECHAS_USUARIO.FechaUltimoPago"
    sStrSql = sStrSql & " FROM (USUARIOS_CLUB INNER JOIN FECHAS_USUARIO ON USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember) INNER JOIN SECUENCIAL ON USUARIOS_CLUB.IdMember = SECUENCIAL.IdMember"
    sStrSql = sStrSql & " WHERE"
    sStrSql = sStrSql & " Multiclub=-1"
    sStrSql = sStrSql & " ORDER BY USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.NumeroFamiliar"


    Set adorcsSrc = New ADODB.Recordset
    
    adorcsSrc.CursorLocation = adUseServer
    
    adorcsSrc.Open , ConnSrc, adOpenForwardOnly, adLockReadOnly
    
    
    Set adorcsTgt = New ADODB.Recordset
    adorcsTgt.CursorLocation = adUseServer
    
    Do Until adorcsSrc.EOF
    
        sStrSql = "SELECT USUARIOS_CLUB.IdMember, USUARIOS_CLUB.NumeroFamiliar, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.FechaNacio, USUARIOS_CLUB.Sexo, USUARIOS_CLUB.IdPais, USUARIOS_CLUB.IdTipoUsuario, USUARIOS_CLUB.IdTitular, USUARIOS_CLUB.Email, USUARIOS_CLUB.Profesion, USUARIOS_CLUB.FechaIngreso, USUARIOS_CLUB.FotoFile, USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Status, USUARIOS_CLUB.Inscripcion, USUARIOS_CLUB.Direccionado, USUARIOS_CLUB.IdTipoAcceso, SECUENCIAL.Secuencial, FECHAS_USUARIO.FechaUltimoPago"
        sStrSql = sStrSql & " FROM (USUARIOS_CLUB INNER JOIN FECHAS_USUARIO ON USUARIOS_CLUB.IdMember = FECHAS_USUARIO.IdMember) INNER JOIN SECUENCIAL ON USUARIOS_CLUB.IdMember = SECUENCIAL.IdMember"
        sStrSql = sStrSql & " WHERE "
        sStrSql = sStrSql & " IdMember=" & adorcsSrc!IdMember
        sStrSql = sStrSql & " ORDER BY USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.NumeroFamiliar"

        adorcsTgt.Open sStrSql, ConnTgt, adOpenDynamic, adLockOptimistic
        
        If adorcsTgt.EOF Then
            Alta
        Else
            Actualiza
        End If
        
    
        adorcsSrc.MoveNext
    Loop


    adorcsSrc.Close
    Set adorcsSrc = Nothing
    
    
End Sub




Sub Alta()
    Dim sStrSql As String
    Dim adocmdAlta As ADODB.Command
    
    
    
    sStrSql = "INSERT INTO USUARIOS_CLUB ("
    sStrSql = sStrSql & " IdMember,"
    sStrSql = sStrSql & " NumeroFamiliar,"
    sStrSql = sStrSql & " Nombre,"
    sStrSql = sStrSql & " A_Paterno,"
    sStrSql = sStrSql & " A_Materno,"
    sStrSql = sStrSql & " FechaNacio,"
    sStrSql = sStrSql & " Sexo,"
    sStrSql = sStrSql & " IdPais,"
    sStrSql = sStrSql & " IdTipoUsuario,"
    sStrSql = sStrSql & " IdTitular,"
    sStrSql = sStrSql & " Email,"
    sStrSql = sStrSql & " Profesion,"
    sStrSql = sStrSql & " FechaIngreso,"
    sStrSql = sStrSql & " FotoFile,"
    sStrSql = sStrSql & " NoFamilia,"
    sStrSql = sStrSql & " Status,"
    sStrSql = sStrSql & " Inscripcion,"
    sStrSql = sStrSql & " Direccionado,"
    sStrSql = sStrSql & " IdTipoAcceso)"
    sStrSql = sStrSql & "VALUES ("
    sStrSql = sStrSql & adorcsSrc!IdMember & ","
    sStrSql = sStrSql & adorcsSrc!NumeroFamiliar & ","
    sStrSql = sStrSql & adorcsSrc!Nombre & ","
    sStrSql = sStrSql & adorcsSrc!A_Paterno & ","
    sStrSql = sStrSql & adorcsSrc!A_Materno & ","
    sStrSql = sStrSql & adorcsSrc!FechaNacio & ","
    sStrSql = sStrSql & adorcsSrc!Sexo & ","
    sStrSql = sStrSql & adorcsSrc!IdPais & ","
    sStrSql = sStrSql & adorcsSrc!IdTipoUsuario & ","
    sStrSql = sStrSql & adorcsSrc!IdTitular & ","
    sStrSql = sStrSql & adorcsSrc!Email & ","
    sStrSql = sStrSql & adorcsSrc!Profesion & ","
    sStrSql = sStrSql & adorcsSrc!FechaIngreso & ","
    sStrSql = sStrSql & adorcsSrc!FotoFile & ","
    sStrSql = sStrSql & adorcsSrc!NoFamilia & ","
    sStrSql = sStrSql & adorcsSrc!Status & ","
    sStrSql = sStrSql & adorcsSrc!Inscripcion & ","
    sStrSql = sStrSql & adorcsSrc!Direccionado & ","
    sStrSql = sStrSql & adorcsSrc!IdTipoAcceso & ")"
    
    
    Set adocmdAlta = New ADODB.Command
    adocmdAlta.ActiveConnection = ConnTgt
    adocmdAlta.CommandType = adCmdText
    adocmdAlta.CommandText = sStrSql
    adocmdAlta.Execute
    
    sStrSql = "INSERT INTO SECUENCIAL ("
    sStrSql = sStrSql & " Secuencial"
    sStrSql = sStrSql & " IdMember)"
    sStrSql = sStrSql & "VALUES ("
    sStrSql = sStrSql & adorcsSrc!Secuencial & ","
    sStrSql = sStrSql & adorcsSrc!IdMember & ")"
    
    adocmdAlta.CommandText = sStrSql
    adocmdAlta.Execute
    
    
    Set adocmdAlta = Nothing
    
End Sub

Sub Actualiza()
    
    
End Sub
