Attribute VB_Name = "Formatos"
Option Explicit

Public Function CreaFormato(iIdFormato As Integer, lidMember As Long) As Long
    Dim adorcsForm As ADODB.Recordset
    Dim adorcsUser As ADODB.Recordset
    Dim adorcsFolio As ADODB.Recordset
    Dim adocmd As ADODB.Command
    Dim sIdentifier As String
    Dim sValor As String
    Dim lIdFormato As Long
    Dim sFormaPago As String
    
    If Not DirExists(Trim(Environ("APPDATA")) & "\KALACLUB\LOG") Then MkDir Trim(Environ("APPDATA")) & "\KALACLUB\LOG"
    
    Open Trim(Environ("APPDATA")) & "\KALACLUB\LOG\CreaFormato_" & Format(Now, "yyyymmddHHnnss") & ".txt" For Output As #1
    Print #1, "Inicia CreaFormato " & CStr(Now)
    
    CreaFormato = 0
    
    strSQL = "SELECT U.NoFamilia AS UNoFamilia, U.idTitular AS UidTitular, U.Nombre AS UNombre, U.A_Paterno AS UA_Paterno, U.A_Materno AS UA_Materno, U.FechaNacio AS UFechaNacio, U.Parentesco AS UParentesco, U.Sexo AS USexo, U.FechaIngreso AS UFechaIngreso, P.Parentesco AS PParentesco, TU.Descripcion AS TUDescripcion, T.Nombre AS TNombre, T.A_Paterno AS TA_Paterno, T.A_Materno AS TA_Materno, T.FechaNacio AS TFechaNacio, T.Parentesco AS TParentesco, T.Sexo AS TSexo, T.FechaIngreso AS TFechaIngreso"
    strSQL = strSQL & " FROM ((USUARIOS_CLUB AS U LEFT JOIN USUARIOS_CLUB AS T ON U.IdTitular = T.IdMember) LEFT JOIN PARENTESCO AS P ON U.Parentesco = P.Clave) LEFT JOIN TIPO_USUARIO AS TU ON U.IdTipoUsuario = TU.IdTipoUsuario"
    strSQL = strSQL & " WHERE (((U.IdMember)=" & lidMember & "))"
    
    Print #1, "***"
    Print #1, strSQL
    
    Set adorcsUser = New ADODB.Recordset
    adorcsUser.CursorLocation = adUseServer
    adorcsUser.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    strSQL = "SELECT FD.IdItem, FD.NombreCampo, FD.TipoCampo"
    strSQL = strSQL & " FROM CT_Formatos F INNER JOIN CT_Formatos_Campos FD ON F.IdTipoFormato = FD.IdTipoFormato"
    strSQL = strSQL & " WHERE (((FD.IdTipoFormato)=" & iIdFormato & "))"
    strSQL = strSQL & " ORDER BY FD.IdItem"
    
    Print #1, "***"
    Print #1, strSQL
    
    Set adorcsForm = New ADODB.Recordset
    adorcsForm.CursorLocation = adUseServer
    adorcsForm.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    
    
    sIdentifier = Left$(sDB_User, 5) & Trim(Right$(Str(Timer()), 5))
    
    'inserta el encabezado del formato
    #If SqlServer_ Then
        strSQL = "SET NOCOUNT ON;"
        strSQL = strSQL & " INSERT INTO FORMATOS ("
        strSQL = strSQL & " IdTipoFormato,"
        strSQL = strSQL & " IdTitular,"
        strSQL = strSQL & " IdMember,"
        strSQL = strSQL & " FechaAlta,"
        strSQL = strSQL & " HoraAlta,"
        strSQL = strSQL & " UsuarioAlta,"
        strSQL = strSQL & " Identifier)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & iIdFormato & ","
        strSQL = strSQL & adorcsUser!UIdTitular & ","
        strSQL = strSQL & lidMember & ","
        strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Format(Now(), "Hh:Nn") & "',"
        strSQL = strSQL & "'" & sDB_User & "',"
        strSQL = strSQL & "'" & sIdentifier & "');"
        strSQL = strSQL & " SELECT @@IDENTITY AS IdFormato;"
        
        Set adorcsFolio = Conn.Execute(strSQL)
    #Else
        strSQL = "INSERT INTO FORMATOS ("
        strSQL = strSQL & " IdTipoFormato,"
        strSQL = strSQL & " IdTitular,"
        strSQL = strSQL & " IdMember,"
        strSQL = strSQL & " FechaAlta,"
        strSQL = strSQL & " HoraAlta,"
        strSQL = strSQL & " UsuarioAlta,"
        strSQL = strSQL & " Identifier)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & iIdFormato & ","
        strSQL = strSQL & adorcsUser!UIdTitular & ","
        strSQL = strSQL & lidMember & ","
        strSQL = strSQL & "#" & Format(Date, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "'" & Format(Now(), "Hh:Nn") & "',"
        strSQL = strSQL & "'" & sDB_User & "',"
        strSQL = strSQL & "'" & sIdentifier & "')"
        
        adocmd.CommandText = strSQL
        adocmd.Execute
    
        strSQL = "SELECT IdFormato"
        strSQL = strSQL & " FROM FORMATOS"
        strSQL = strSQL & " WHERE Identifier='" & sIdentifier & "'"
    
        Print #1, "***"
        Print #1, strSQL
    
        Set adorcsFolio = New ADODB.Recordset
        adorcsFolio.CursorLocation = adUseServer
        adorcsFolio.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    #End If
    
    Print #1, "***"
    Print #1, strSQL
    
    If Not adorcsFolio.EOF Then
        lIdFormato = adorcsFolio!IdFormato
    End If
    
    adorcsFolio.Close
    Set adorcsFolio = Nothing
    
    Do Until adorcsForm.EOF
        Select Case adorcsForm!NombreCampo
            Case "FECHA"
                #If SqlServer_ Then
                    sValor = Format(Date, "yyyymmdd")
                #Else
                    sValor = Format(Date, "dd/mm/yyyy")
                #End If
             Case "UNIDAD_NOMBRE"
                sValor = ObtieneParametro("NOMBRE DEL CLUB")
            Case "NOMBRE_INTEGRANTE"
                sValor = adorcsUser![UNombre] & " " & adorcsUser![UA_Paterno] & " " & adorcsUser![UA_Materno]
            Case "INSCRIPCION_NO"
                sValor = adorcsUser!UNoFamilia
            Case "NOMBRE_TITULAR"
                sValor = adorcsUser![TNombre] & " " & adorcsUser![TA_Paterno] & " " & adorcsUser![TA_Materno]
            Case "FECHA_NAC_INTEGRANTE"
                sValor = adorcsUser![UFechaNacio]
            Case "PARENTESCO_INTEGRANTE"
                sValor = IIf(IsNull(adorcsUser![PPARENTESCO]), "", adorcsUser![PPARENTESCO])
            Case "TIPO_USUARIO_INTEGRANTE"
                sValor = IIf(IsNull(adorcsUser!TUDescripcion), "", adorcsUser!TUDescripcion)
            Case "MANTENIMIENTO"
                If sFormaPago = vbNullString Then
                    sFormaPago = ChecaDireccionado(adorcsUser!UIdTitular)
                End If
                
                sValor = CalculaMantenimientoMes(adorcsUser!UIdTitular, IIf(sFormaPago = vbNullString, False, True), 1)
                
            Case "FORMA_PAGO_MANTENIMIENTO"
                If sFormaPago = vbNullString Then
                    sFormaPago = ChecaDireccionado(adorcsUser!UIdTitular)
                End If
                If sFormaPago = vbNullString Then
                    sValor = "C"
                Else
                    sValor = sFormaPago
                End If
            Case "USUARIO_SISTEMA"
                sValor = sDB_User
            Case "NOMBRE_REPRESENTANTE"
                sValor = ObtieneParametro("PERSONA CONTRATO")
            
            Case "FOLIO_FORMATO"
                sValor = Format(lIdFormato, "000000")
            Case Else
                sValor = ""
        End Select
        
        strSQL = "INSERT INTO FORMATOS_DETALLE ("
        strSQL = strSQL & " IdFormato,"
        strSQL = strSQL & " IdItem,"
        strSQL = strSQL & " Valor)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & lIdFormato & ","
        strSQL = strSQL & adorcsForm!IdItem & ","
        strSQL = strSQL & "'" & sValor & "')"
        
        Print #1, "***"
        Print #1, strSQL
        
        adocmd.CommandText = strSQL
        adocmd.Execute
        
        adorcsForm.MoveNext
    Loop
    
    adorcsForm.Close
    Set adorcsForm = Nothing
    
    adorcsUser.Close
    Set adorcsUser = Nothing
    
    Set adocmd = Nothing
    
    CreaFormato = lIdFormato
    
    Print #1, "***"
    Print #1, "Finalizado " & CStr(Now)
    Print #1, "***"
    Close #1
    Exit Function

Error_Catch:
    
    Print #1, "***"
    Print #1, "Error: " & Err.Description
    Print #1, "***"
    Print #1, "Finalizado " & CStr(Now)
    Print #1, "***"
    Close #1
End Function



