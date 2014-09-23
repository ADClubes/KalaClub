Attribute VB_Name = "FunsSocios"
'********************************************
'*  Funciones para el modulo de los socios  *
'*  Daniel Hdez                             *
'*  01 / Julio / 2005                       *
'*  Ult Act: 14 / Noviembre / 2005          *
'********************************************


'Devuelve la primer clave disponible para asignar
Public Function BuscaCve(sTabla As String, sCampo As String) As Long
Dim lCveNva As Long
Dim lCveSiguiente As Long
Dim rsBuscaCve As ADODB.Recordset

    lCveNva = 1
    
    InitRecordSet rsBuscaCve, sCampo, sTabla, "", sCampo, Conn
    
    If (rsBuscaCve.RecordCount > 0) Then
        rsBuscaCve.MoveFirst
        
        lCveNva = rsBuscaCve.Fields(sCampo)
        lCveSiguiente = rsBuscaCve.Fields(sCampo)
        Do While (Not rsBuscaCve.EOF)
            lCveNva = lCveSiguiente
            lCveSiguiente = rsBuscaCve.Fields(sCampo)
            
            If ((lCveSiguiente - lCveNva) > 1) Then
                lCveNva = lCveNva + 1
                Exit Do
            End If
            
            rsBuscaCve.MoveNext
        Loop
        
        If (rsBuscaCve.EOF) Then
            lCveNva = lCveSiguiente + 1
        End If
    End If
    
    'Cierra el recordset
    rsBuscaCve.Close
    Set rsBuscaCve = Nothing

    BuscaCve = lCveNva
End Function


'Busca X valor en una tabla
Public Function ExisteXValor(sCampos As String, sTablas As String, sCondicion As String, cnConnection As ADODB.Connection, sOrden As String) As Boolean
Dim rsExisteXValor As ADODB.Recordset

    ExisteXValor = False

    InitRecordSet rsExisteXValor, sCampos, sTablas, sCondicion, sOrden, cnConnection

    With rsExisteXValor
        If (.RecordCount > 0) Then
            ExisteXValor = True
        End If
        
        .Close
    End With

    Set rsExisteXValor = Nothing
End Function


' Calcula la edad en base a la fecha de nacimiento
Public Function Edad(dNacio As Date)
Dim nAnios, nMeses, nToday, nBorn As Integer
    
    On Error GoTo ErrEdad

    If (Not IsEmpty(dNacio)) Then
        'nToday = (Year(dMiFecha) * 12) + Month(dMiFecha)
        nToday = (Year(Date) * 12) + Month(Date)
        nBorn = (Year(dNacio) * 12) + Month(dNacio)
    
        nMeses = ((nToday - nBorn) Mod 12) / 100
        nAnios = Int((nToday - nBorn) / 12)
    
        Edad = Rounded(nAnios + nMeses, 2)
    Else
        Edad = 0
    End If
    Exit Function
    
ErrEdad:
    MsgBox "Revise la fecha de la edad de nacimiento.", vbExclamation, "KalaSystems"
    Edad = 0
End Function


'Funcion para redondear numeros
Function Rounded(ByVal xNumero As Variant, ByVal niDecimales As Integer) As Variant
Dim rValor As Double
Dim xRetValue As Variant
Dim lPositivo As Boolean

    xRetValue = xNumero

    ' Numero en Positivo
    rValor = xNumero * Sgn(xNumero)

    ' Dejamos las unidades en la posicion a truncar
    rValor = rValor * (10 ^ niDecimales)

    ' Comprobacion de la parte decimal
    If rValor - Fix(rValor) > 0.5 Then
        rValor = Fix(rValor) + 1
    Else
        rValor = Fix(rValor)
    End If

    'Restaurar las unidades a su posicion original
    rValor = rValor / (10 ^ niDecimales)

    ' Volvemos a colocar el signo que le corresponde
    rValor = rValor * Sgn(xNumero)

    xRetValue = rValor
    Rounded = xRetValue
End Function


'Actualiza un registro en la tabla con los datos que llegan como parametros
Public Function CambiaReg(sTabla As String, mCampos() As String, nNumDatos As Byte, mValores() As Variant, sCondicion As String, cnConnect As ADODB.Connection) As Boolean
    Dim sStrSql, sTipoValor As String
    Dim nPointer As Byte
    Dim sRecSource As String
    Dim lRegistros As Long
    Dim iniTrans As Long

    On Error GoTo ErrorGuardar

    'cnConnect.Errors.Clear
    'Err.Clear

    Screen.MousePointer = vbHourglass

    'Arma la sentencia SQL
    If (nNumDatos > 1) Then
        sStrSql = "UPDATE " & sTabla & " SET "

        For nPointer = 0 To (nNumDatos - 2)
            If (Not IsNull(mValores(nPointer))) Then
                sStrSql = sStrSql & mCampos(nPointer) & "='" & mValores(nPointer) & "',"
            Else
                sStrSql = sStrSql & mCampos(nPointer) & "=NULL,"
            End If
        Next nPointer

        If (Not IsNull(mValores(nNumDatos - 1))) Then
            sStrSql = sStrSql & mCampos(nNumDatos - 1) & "='" & mValores(nNumDatos - 1) & "'"
        Else
            sStrSql = sStrSql & mCampos(nNumDatos - 1) & "=NULL"
        End If
    Else
        If (Not IsNull(mValores(0))) Then
            sStrSql = "UPDATE " & sTabla & " SET " & mCampos(0) & "='" & mValores(0) & "'"
        Else
            sStrSql = "UPDATE " & sTabla & " SET " & mCampos(0) & "=NULL"
        End If
    End If

    sStrSql = sStrSql & " WHERE " & sCondicion

    'Agrega el registro
    'iniTrans = cnConnect.BeginTrans
    Set AdoCmdGuardar = New ADODB.Command
    AdoCmdGuardar.ActiveConnection = cnConnect
    AdoCmdGuardar.CommandText = sStrSql
    AdoCmdGuardar.Execute lRegistros
    'cnConnect.CommitTrans

    Screen.MousePointer = Default

    CambiaReg = True
    Exit Function

ErrorGuardar:

    Screen.MousePointer = Default

    'If (iniTrans > 0) Then
    '    cnConnect.RollbackTrans
    'End If
    
    CambiaReg = False

    MsgError

    'MsgBox Err.Description, vbCritical, "Alta de registro"

    
End Function


'Borra un registro en la tabla con los datos que llegan como parametros
Public Function EliminaReg(sTabla As String, sCondicion As String, sTituloError As String, Optional cnConnect As ADODB.Connection) As Boolean
    Dim sStrSql As String
    Dim nPointer As Byte
    Dim sRecSource As String
    Dim lRegistros As Long
    Dim iniTrans As Long

    On Error GoTo ErrorGuardar

    cnConnect.Errors.Clear
    Err.Clear

    Screen.MousePointer = vbHourglass

    'Arma la sentencia SQL
    If (Trim(sCondicion) <> "") Then
        sStrSql = "DELETE FROM " & sTabla & " WHERE " & sCondicion
    Else
        sStrSql = "DELETE FROM " & sTabla
    End If

    'Borra el registro
    'iniTrans = cnConnect.BeginTrans
    Set AdoCmdGuardar = New ADODB.Command
    AdoCmdGuardar.ActiveConnection = cnConnect
    AdoCmdGuardar.CommandText = sStrSql
    AdoCmdGuardar.Execute lRegistros
    'cnConnect.CommitTrans

    Screen.MousePointer = Default

    EliminaReg = True
    Exit Function

ErrorGuardar:

    Screen.MousePointer = Default

    'If (iniTrans > 0) Then
    '    cnConnect.RollbackTrans
    'End If
    
    MsgBox Err.Description, vbCritical, "Borrar registro"

    EliminaReg = False
End Function



'Agrega un registro en la tabla con los datos que llegan como parametros
Public Function AgregaRegistro(sTabla As String, mCampos() As String, nNumDatos As Byte, mValores() As Variant, cnConnect As ADODB.Connection) As Boolean
    Dim sError As String
    Dim sListaCampos As String
    Dim sListaValores As String
    Dim sStrSql As String
    Dim nPointer As Byte
    Dim sRecSource As String
    Dim lRegistros As Long
    Dim iniTrans As Long


    On Error GoTo ErrorGuardar
    
    AgregaRegistro = False
    
    'cnConnect.Errors.Clear
    'Err.Clear
    
    Screen.MousePointer = vbHourglass
    

    'Crea la lista de los campos y valores del registro
    If (nNumDatos > 1) Then
        sListaCampos = "("
        sListaValores = "("
        
        For nPointer = 0 To (nNumDatos - 2)
            sListaCampos = sListaCampos & mCampos(nPointer) & ","
            sListaValores = sListaValores & "'" & mValores(nPointer) & "',"
        Next nPointer
        
        sListaCampos = sListaCampos & mCampos(nNumDatos - 1) & ")"
        sListaValores = sListaValores & "'" & mValores(nNumDatos - 1) & "')"
    Else
        'sListaCampos = "('" & mCampos(0) & "')"
        sListaCampos = "(" & mCampos(0) & ")"
        sListaValores = "('" & mValores(0) & "')"
    End If


    'Arma la sentencia SQL
    sStrSql = "INSERT INTO " & sTabla & " " & sListaCampos & " VALUES " & sListaValores
    
    'Agrega el registro
    'iniTrans = cnConnect.BeginTrans
    Set AdoCmdGuardar = New ADODB.Command
    AdoCmdGuardar.ActiveConnection = cnConnect
    AdoCmdGuardar.CommandText = sStrSql
    AdoCmdGuardar.Execute lRegistros
    'cnConnect.CommitTrans
    
    Screen.MousePointer = Default
    
    AgregaRegistro = True
    
    Exit Function
    
ErrorGuardar:

    Screen.MousePointer = Default

    'If (iniTrans > 0) Then
    '    cnConnect.RollbackTrans
    'End If
    
    'MsgBox Err.Description, vbCritical, "Cambio en el registro"
    
    MsgError
    
End Function


'Lee el valor del campo especificado
Public Function LeeXValor(sCampos As String, sTabla As String, sCondicion As String, sCampoXRegresar As String, sTipoDato As String, cnConexion As ADODB.Connection)
Dim rsTabla As ADODB.Recordset
Dim xResultado

    xResultado = Null
    
    InitRecordSet rsTabla, sCampos, sTabla, sCondicion, sCampoXRegresar, cnConexion
    
    With rsTabla
        If (.RecordCount > 0) Then
            If (Not IsNull(.Fields(sCampoXRegresar))) Then
                xResultado = .Fields(sCampoXRegresar)
            Else
                Select Case sTipoDato
                    Case "s":
                        xResultado = ""
                        
                    Case "n":
                        xResultado = 0
                    
                    Case "d":
                        xResultado = "#01/01/1900#"
                End Select
            End If
        Else
            Select Case sTipoDato
                Case "s":
                    xResultado = "VACIO"
                    
                Case "n":
                    xResultado = -1
                
                Case "d":
                    xResultado = "#01/01/1900#"
            End Select
        End If
    End With
    
    rsTabla.Close
    Set rsTabla = Nothing
    
    LeeXValor = xResultado
End Function


'Se coloca en el ultimo registro de la tabla y lee el valor del campo indicado
Public Function LeeUltReg(sTablas As String, sCampo As String) As Long
Dim rsUltValor As ADODB.Recordset

    InitRecordSet rsUltValor, sCampo, sTablas, "", sCampo, Conn
    With rsUltValor
        If (.RecordCount > 0) Then
            .MoveLast
            LeeUltReg = .Fields(sCampo)
        Else
            LeeUltReg = 0
        End If
        
        .Close
    End With
    
    Set rsUltValor = Nothing
End Function


'Incrementa el contador mensual en las tablas de ALTAS_ANUALES y BAJAS_ANUALES
Public Function Altas_Bajas(bAlta As Boolean) As Boolean
    Const NUMDATOS = 2
    Dim mFields(NUMDATOS) As String
    Dim mValues(NUMDATOS) As Variant
    Dim sTabla As String
    Dim nValor As Integer
    Dim sMes As String
    Dim sCond As String

    Altas_Bajas = False
    
    'Nombre del mes(campo) actual
    sMes = NombreMes(Month(Date))
    
    'Nombre de la tabla que se va a utilizar
    sTabla = IIf(bAlta, "ALTAS_ANUALES", "BAJAS_ANUALES")
    
    'Condicion
    sCond = "Anio=" & Year(Date)
    
    'Lee el numero de altas o bajas del mes en cuestion
    nValor = LeeXValor(sMes, sTabla, sCond, sMes, "n", Conn)
    
    'Campos de las tablas: ALTAS_ANUALES y BAJAS_ANUALES
    mFields(0) = "Anio"
    mFields(1) = sMes
    
    'Valores de los campos
    mValues(0) = Year(Date)
    
    'Si no existe el año buscado, lo da de alta
    If (nValor = -1) Then
        mValues(1) = 1
        Altas_Bajas = AgregaRegistro(sTabla, mFields, NUMDATOS, mValues, Conn)
    Else
        'Si existe el registro correspondiente al año, solo incrementa el contador del mes
        mValues(1) = nValor + 1
        Altas_Bajas = CambiaReg(sTabla, mFields, NUMDATOS, mValues, sCond, Conn)
    End If
End Function


'Regresa el nombre del mes
Public Function NombreMes(bNoMes As Byte) As String
    Select Case bNoMes
        Case 1
            NombreMes = "ENERO"
            
        Case 2
            NombreMes = "FEBRERO"
            
        Case 3
            NombreMes = "MARZO"
            
        Case 4
            NombreMes = "ABRIL"
            
        Case 5
            NombreMes = "MAYO"
            
        Case 6
            NombreMes = "JUNIO"
            
        Case 7
            NombreMes = "JULIO"
            
        Case 8
            NombreMes = "AGOSTO"
            
        Case 9
            NombreMes = "SEPTIEMBRE"
            
        Case 10
            NombreMes = "OCTUBRE"
            
        Case 11
            NombreMes = "NOVIEMBRE"
            
        Case 12
            NombreMes = "DICIEMBRE"
    End Select
End Function


'Borra familiares de la tabla: Usuarios_Club
Public Function QuitaFamiliar(nidMember As Integer, nSystemUser As Integer) As Boolean
Const NUMDATOS = 17
Dim rsDatosFam As ADODB.Recordset
Dim mFieldsBaja(NUMDATOS) As String
Dim mValuesBaja(NUMDATOS) As Variant
Dim sCampos As String
Dim nSec As Long
Dim InitTrans As Long

    QuitaFamiliar = False

    sCampos = "idMember, Nombre, A_Paterno, A_Materno, FechaNacio, Sexo, IdPais, "
    sCampos = sCampos & "IdTipoUsuario, Email, Celular, Profesion, FechaIngreso, "
    sCampos = sCampos & "IdTitular, NoFamilia"

    'Lee los datos del familiar para registralos en la tabla de Bajas
    InitRecordSet rsDatosFam, sCampos, "Usuarios_Club", "IdMember=" & nidMember, "", Conn
    With rsDatosFam
        If (.RecordCount > 0) Then
            'Campos de la tabla Bajas
            mFieldsBaja(0) = "idMember"
            mFieldsBaja(1) = "Nombre"
            mFieldsBaja(2) = "A_Paterno"
            mFieldsBaja(3) = "A_Materno"
            mFieldsBaja(4) = "FechaNacio"
            mFieldsBaja(5) = "Sexo"
            mFieldsBaja(6) = "IdPais"
            mFieldsBaja(7) = "IdTipoUsuario"
            mFieldsBaja(8) = "Email"
            mFieldsBaja(9) = "Celular"
            mFieldsBaja(10) = "Profesion"
            mFieldsBaja(11) = "FechaIngreso"
            mFieldsBaja(12) = "IdTitular"
            mFieldsBaja(13) = "IdUsuario"
            mFieldsBaja(14) = "Usuario"
            mFieldsBaja(15) = "Fecha"
            mFieldsBaja(16) = "NoFamilia"
            
            'Valores para la tabla Bajas
            #If SqlServer_ Then
                mValuesBaja(0) = nidMember
                mValuesBaja(1) = .Fields("Nombre")
                mValuesBaja(2) = .Fields("A_Paterno")
                mValuesBaja(3) = .Fields("A_Materno")
                mValuesBaja(4) = .Fields("FechaNacio")
                mValuesBaja(5) = .Fields("Sexo")
                mValuesBaja(6) = .Fields("IdPais")
                mValuesBaja(7) = .Fields("IdTipoUsuario")
                mValuesBaja(8) = .Fields("Email")
                mValuesBaja(9) = .Fields("Celular")
                mValuesBaja(10) = .Fields("Profesion")
                mValuesBaja(11) = .Fields("FechaIngreso")
                mValuesBaja(12) = .Fields("IdTitular")
                If (nSystemUser = 0) Then
                    mValuesBaja(13) = LeeXValor("IdUsuario", "Usuarios_Sistema", "Login_Name='" & sDB_User & "'", "IdUsuario", "n", Conn)
                Else
                    mValuesBaja(13) = nSystemUser
                End If
                mValuesBaja(14) = sDB_User
                mValuesBaja(15) = Format(Date, "yyyymmdd")
                mValuesBaja(16) = .Fields("NoFamilia")
            #Else
                mValuesBaja(0) = nidMember
                mValuesBaja(1) = .Fields("Nombre")
                mValuesBaja(2) = .Fields("A_Paterno")
                mValuesBaja(3) = .Fields("A_Materno")
                mValuesBaja(4) = .Fields("FechaNacio")
                mValuesBaja(5) = .Fields("Sexo")
                mValuesBaja(6) = .Fields("IdPais")
                mValuesBaja(7) = .Fields("IdTipoUsuario")
                mValuesBaja(8) = .Fields("Email")
                mValuesBaja(9) = .Fields("Celular")
                mValuesBaja(10) = .Fields("Profesion")
                mValuesBaja(11) = .Fields("FechaIngreso")
                mValuesBaja(12) = .Fields("IdTitular")
                If (nSystemUser = 0) Then
                    mValuesBaja(13) = LeeXValor("IdUsuario", "Usuarios_Sistema", "Login_Name='" & sDB_User & "'", "IdUsuario", "n", Conn)
                Else
                    mValuesBaja(13) = nSystemUser
                End If
                mValuesBaja(14) = sDB_User
                mValuesBaja(15) = Format(Date, "dd/mm/yyyy")
                mValuesBaja(16) = .Fields("NoFamilia")
            #End If
        End If
        .Close
    End With
    Set rsDatosFam = Nothing
    
    'Busca el numero secuencial que tenia asignado
    nSec = LeeXValor("Secuencial", "Secuencial", "idMember=" & nidMember, "Secuencial", "n", Conn)
    
    'Inicia el registro de los datos en las tablas
    InitTrans = Conn.BeginTrans
    
    'Registra los datos en la tabla de Bajas
    If (AgregaRegistro("Bajas", mFieldsBaja, NUMDATOS, mValuesBaja, Conn)) Then
    
        'Elimina el registro de la tabla Time_Zone_Users
        If (EliminaReg("Time_Zone_Users", "idMember=" & nidMember, "", Conn)) Then
        
            'Elimina el registro de la tabla Certificados médicos
            If (EliminaReg("Certificados", "idMember=" & nidMember, "", Conn)) Then
            
                'Elimina el registro de la tabla Fechas_Usuario
                If (EliminaReg("Fechas_Usuario", "idMember=" & nidMember, "", Conn)) Then
    
                    'Elimina el registro de la tabla de Usuarios_Club
                    If (EliminaReg("Usuarios_Club", "idMember=" & nidMember, "", Conn)) Then
                
                        'Incrementa el contador en la tabla de Bajas
                        If (Altas_Bajas(False)) Then
                        
                            'Quita la asignacion del numero secuencial
                            If (CambiaSec(0, nSec, False)) Then
                                QuitaFamiliar = True
                    
                                'Baja a disco los cambios realizados
                                Conn.CommitTrans
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        'En caso de algun error no baja a disco los nuevos datos
        If InitTrans > 0 Then
            Conn.RollbackTrans
        End If
    End If
End Function


'Borra titulares de la tabla: Usuarios_Club
Public Function QuitaTitular(nidMember As Integer) As Boolean
Const NUMDATOS = 17
Dim rsDatosTit As ADODB.Recordset
Dim rsDatosFam As ADODB.Recordset
Dim mFieldsBaja(NUMDATOS) As String
Dim mValuesBaja(NUMDATOS) As Variant
Dim sCampos As String
Dim nSec As Long
Dim InitTrans As Long

    QuitaTitular = False

    sCampos = "idMember, Nombre, A_Paterno, A_Materno, FechaNacio, Sexo, IdPais, "
    sCampos = sCampos & "IdTipoUsuario, Email, Celular, Profesion, FechaIngreso, "
    sCampos = sCampos & "IdTitular, NoFamilia"

    'Lee los datos del titular para registralos en la tabla de Bajas
    InitRecordSet rsDatosTit, sCampos, "Usuarios_Club", "IdMember=" & nidMember, "", Conn
    With rsDatosTit
        If (.RecordCount > 0) Then
            'Campos de la tabla Bajas
            mFieldsBaja(0) = "idMember"
            mFieldsBaja(1) = "Nombre"
            mFieldsBaja(2) = "A_Paterno"
            mFieldsBaja(3) = "A_Materno"
            mFieldsBaja(4) = "FechaNacio"
            mFieldsBaja(5) = "Sexo"
            mFieldsBaja(6) = "IdPais"
            mFieldsBaja(7) = "IdTipoUsuario"
            mFieldsBaja(8) = "Email"
            mFieldsBaja(9) = "Celular"
            mFieldsBaja(10) = "Profesion"
            mFieldsBaja(11) = "FechaIngreso"
            mFieldsBaja(12) = "IdTitular"
            mFieldsBaja(13) = "IdUsuario"
            mFieldsBaja(14) = "Usuario"
            mFieldsBaja(15) = "Fecha"
            mFieldsBaja(16) = "NoFamilia"
            
            'Valores para la tabla Bajas
            #If SqlServer_ Then
                mValuesBaja(0) = nidMember
                mValuesBaja(1) = .Fields("Nombre")
                mValuesBaja(2) = .Fields("A_Paterno")
                mValuesBaja(3) = .Fields("A_Materno")
                mValuesBaja(4) = .Fields("FechaNacio")
                mValuesBaja(5) = .Fields("Sexo")
                mValuesBaja(6) = .Fields("IdPais")
                mValuesBaja(7) = .Fields("IdTipoUsuario")
                mValuesBaja(8) = .Fields("Email")
                mValuesBaja(9) = .Fields("Celular")
                mValuesBaja(10) = .Fields("Profesion")
                mValuesBaja(11) = .Fields("FechaIngreso")
                mValuesBaja(12) = .Fields("IdTitular")
                mValuesBaja(13) = LeeXValor("IdUsuario", "Usuarios_Sistema", "Login_Name='" & sDB_User & "'", "IdUsuario", "n", Conn)
                mValuesBaja(14) = sDB_User
                mValuesBaja(15) = Format(Date, "yyyymmdd")
                mValuesBaja(16) = .Fields("NoFamilia")
            #Else
                mValuesBaja(0) = nidMember
                mValuesBaja(1) = .Fields("Nombre")
                mValuesBaja(2) = .Fields("A_Paterno")
                mValuesBaja(3) = .Fields("A_Materno")
                mValuesBaja(4) = .Fields("FechaNacio")
                mValuesBaja(5) = .Fields("Sexo")
                mValuesBaja(6) = .Fields("IdPais")
                mValuesBaja(7) = .Fields("IdTipoUsuario")
                mValuesBaja(8) = .Fields("Email")
                mValuesBaja(9) = .Fields("Celular")
                mValuesBaja(10) = .Fields("Profesion")
                mValuesBaja(11) = .Fields("FechaIngreso")
                mValuesBaja(12) = .Fields("IdTitular")
                mValuesBaja(13) = LeeXValor("IdUsuario", "Usuarios_Sistema", "Login_Name='" & sDB_User & "'", "IdUsuario", "n", Conn)
                mValuesBaja(14) = sDB_User
                mValuesBaja(15) = Format(Date, "dd/mm/yyyy")
                mValuesBaja(16) = .Fields("NoFamilia")
            #End If
        End If
        .Close
    End With
    Set rsDatosTit = Nothing
    
    'Busca el numero secuencial que tenia asignado
    nSec = LeeXValor("Secuencial", "Secuencial", "idMember=" & nidMember, "Secuencial", "n", Conn)
    
    'Inicia el registro de los datos en las tablas
    InitTrans = Conn.BeginTrans
    
    'Registra los datos en la tabla de Bajas
    If (AgregaRegistro("Bajas", mFieldsBaja, NUMDATOS, mValuesBaja, Conn)) Then
    
        'Borra los datos adicionales relacionados con el titular
        If (BorraDatosAdic(nidMember)) Then
        
            'Elimina el registro de la tabla Time_Zone_Users
            If (EliminaReg("Time_Zone_Users", "idMember=" & nidMember, "", Conn)) Then
            
                'Elimina el registro de la tabla Certificados médicos
                If (EliminaReg("Certificados", "idMember=" & nidMember, "", Conn)) Then
                
                    'Elimina el registro de la tabla Fechas_Usuario
                    If (EliminaReg("Fechas_Usuario", "idMember=" & nidMember, "", Conn)) Then
    
                        'Elimina el registro de la tabla de Usuarios_Club
                        If (EliminaReg("Usuarios_Club", "idMember=" & nidMember, "", Conn)) Then
                        
                            'Incrementa el contador en la tabla de Bajas
                            If (Altas_Bajas(False)) Then
                            
                                'Quita la asignacion del numero secuencial
                                If (CambiaSec(0, nSec, False)) Then
                                    
                                    'Borra los familiares
                                    InitRecordSet rsDatosFam, "idMember", "Usuarios_Club", "idTitular=" & nidMember & " AND idMember<>idTitular", "", Conn
                                    With rsDatosFam
                                        If (.RecordCount > 0) Then
                                            .MoveFirst
                                            Do While (Not .EOF)
                                                If (Not QuitaFamiliar(.Fields("idMember"), nidMember)) Then
                                                    Exit Do
                                                End If
                                                
                                                .MoveNext
                                            Loop
                                            
                                            If (Not .EOF) Then
                                                .Close
                                                Set rsDatosFam = Nothing
                                                
                                                If InitTrans > 0 Then
                                                    Conn.RollbackTrans
                                                End If
                                                
                                                Exit Function
                                            End If
                                        End If
                                        
                                        .Close
                                    End With
                                    
                                    Set rsDatosFam = Nothing
                                
                                    QuitaTitular = True
                        
                                    'Baja a disco los cambios realizados
                                    Conn.CommitTrans
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        'En caso de algun error no baja a disco los nuevos datos
        If InitTrans > 0 Then
            Conn.RollbackTrans
        End If
    End If
End Function


'Borra los datos adicionales (calcomanias, direcciones y la relación del título) relacionados con el titular
Public Function BorraDatosAdic(nidMember As Integer) As Boolean
    BorraDatosAdic = False

    'Baja de la tabla de calcomanias
    If (EliminaReg("Calcomanias", "idMember=" & nidMember, "", Conn)) Then
        
        'Baja de direcciones
        If (EliminaReg("Direcciones", "idMember=" & nidMember, "", Conn)) Then
            
            'Baja de Usuarios_Tiutlo
            If (EliminaReg("Usuarios_Titulo", "idMember=" & nidMember, "", Conn)) Then
                BorraDatosAdic = True
            End If
        End If
    End If
End Function
'Eliminar Registrio en NITGEN
'Public Function EliminarAcceso(nUsuario As String) As Long
'Dim sIdMember As String
'Dim nErrCode As Long
'
'Set Conn = Nothing
'
''Set objAccessManager = New AccessManagerSDKLib.AccessManager
''Set objUserInfo = objAccessManager.UserInfo
''Set objTerminalUserInfo = objAccessManager.TerminalUserInfo
'
'sIdMember = nUsuario
'
'Call objAccessManager.ConnectToServer(593192, strTorniquetes, 7331)
'
'nErrCode = objAccessManager.errorCode
'
'If objAccessManager.errorCode = 0 Then
'    Call objAccessManager.CheckAdmin("0000000", "1234", 2)
'    If objAccessManager.errorCode = 0 Then
'
'        Call objUserInfo.DeleteUserInfoByUserID(nUsuario, True)
'
'            EliminarAcceso = objTerminalUserInfo.errorCode
'
'
'        Call objAccessManager.DisconnectServer
'    Else
'    EliminarAcceso = 1
'    End If
'Else
'    EliminarAcceso = 1
'    Call objAccessManager.DisconnectServer
'    'Exit Sub
'End If
'If nErrCode <> 0 Then
'EliminarAcceso = 1
'End If
'
'    'Set objUserInfo = Nothing
'     '       Set objTerminalInfo = Nothing
'      '      Set objAccessManager = Nothing
'       '     Set objTerminalUserInfo = Nothing
'        '    Set objTerminalList = Nothing
'
'Connection_DB
'
'End Function
''Valida si existe el usuario en el servidor de NITGEN
'Function ExisteUsuario(ID As String) As Boolean
'    ExisteUsuario = False
'    Dim tError As Long
'
'    'Set objUserInfo = objAccessManager.UserInfo
'    objUserInfo.GetUserInfoByUserID (ID)
'    tError = objUserInfo.errorCode
'
'    If tError = 0 Then
'        ExisteUsuario = True
'    Else
'        ExisteUsuario = False
'    End If
'
'
'End Function
''Habilita Acceso en NITGEN
'Public Function HabilitaAccesoM(IdMember As Long) As Long
'Dim sIdMember As String
'Dim nErrCode As Long
'Set Conn = Nothing
'
''Set objAccessManager = New AccessManagerSDKLib.AccessManager
''Set objUserInfo = objAccessManager.UserInfo
''Set objTerminalUserInfo = objAccessManager.TerminalUserInfo
''Set objTerminalInfo = objAccessManager.TerminalInfo
'
'sIdMember = Format(IdMember, "00000")
'Call objAccessManager.ConnectToServer(593192, strTorniquetes, 7331)
'nErrCode = objAccessManager.errorCode
'If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'    Call objAccessManager.CheckAdmin("0000000", "1234", 2)
'
'    If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'        'validar si existe usuario
'        If ExisteUsuario("68" + sIdMember) Then
'            '>>si existe usuario actualizar datos
'
'        Call objTerminalInfo.GetTerminalList
'        nErrCode = objTerminalInfo.errorCode
'
'        'Dim objTerminalList As AccessManagerSDKLib.ITerminalList
'        Dim tError As Long
'        Dim dRes As Boolean
'
'    tError = 81
'
'        For Each objTerminalList In objTerminalInfo
'            Call objTerminalUserInfo.AddUserToTerminalByUserID(objTerminalList.TerminalID, "68" + sIdMember, 2, True)
'
'              tError = objTerminalUserInfo.errorCode
'
'            If tError = 805372429 Then
'                '>El usuario ya existe en terminal, continuar.
'                HabilitaAccesoM = 0
'                Else
'                HabilitaAccesoM = objTerminalUserInfo.errorCode
'            End If
'
'        Next
'         End If
'
'
'
'        Call objAccessManager.DisconnectServer
'
'    End If
'Else
'    HabilitaAccesoM = 1
'    Call objAccessManager.DisconnectServer
'    'Exit Sub
'End If
'If nErrCode <> 0 Then
'HabilitaAccesoM = 1
'End If
'    'Set objUserInfo = Nothing
'     '       Set objTerminalInfo = Nothing
'      '      Set objAccessManager = Nothing
'       '     Set objTerminalUserInfo = Nothing
'        '    Set objTerminalList = Nothing
'Connection_DB
'End Function
'Public Function HabilitaAcceso(IdMember As Long) As Long
'Dim sIdMember As String
'Dim nErrCode As Long
''Set Conn = Nothing
'
''Set objAccessManager = New AccessManagerSDKLib.AccessManager
''Set objUserInfo = objAccessManager.UserInfo
''Set objTerminalUserInfo = objAccessManager.TerminalUserInfo
''Set objTerminalInfo = objAccessManager.TerminalInfo
'
'sIdMember = Format(IdMember, "00000")
'Call objAccessManager.ConnectToServer(593192, strTorniquetes, 7331)
'nErrCode = objAccessManager.errorCode
'If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'    Call objAccessManager.CheckAdmin("0000000", "1234", 2)
'
'    If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'        'validar si existe usuario
'        If ExisteUsuario("68" + sIdMember) Then
'            '>>si existe usuario actualizar datos
'
'        Call objTerminalInfo.GetTerminalList
'        nErrCode = objTerminalInfo.errorCode
'
'        'Dim objTerminalList As AccessManagerSDKLib.ITerminalList
'        Dim tError As Long
'        Dim dRes As Boolean
'
'    tError = 81
'
'        For Each objTerminalList In objTerminalInfo
'            Call objTerminalUserInfo.AddUserToTerminalByUserID(objTerminalList.TerminalID, "68" + sIdMember, 2, True)
'
'              tError = objTerminalUserInfo.errorCode
'
'            If tError = 805372429 Then
'                '>El usuario ya existe en terminal, continuar.
'                HabilitaAcceso = 0
'                Else
'                HabilitaAcceso = objTerminalUserInfo.errorCode
'            End If
'
'        Next
'         End If
'
'
'
'        Call objAccessManager.DisconnectServer
'
'    End If
'Else
'    HabilitaAcceso = 1
'    Call objAccessManager.DisconnectServer
'    'Exit Sub
'End If
'If nErrCode <> 0 Then
'HabilitaAcceso = 1
'End If
'    'Set objUserInfo = Nothing
'     '       Set objTerminalInfo = Nothing
'      '      Set objAccessManager = Nothing
'       '     Set objTerminalUserInfo = Nothing
'        '    Set objTerminalList = Nothing
''Connection_DB
'End Function
''Bloquea Acceso en NITGEN
'Public Function BloqueaAccesoM(IdMember As Long) As Long
'Dim sIdMember As String
'Dim nErrCode As Long
'Set Conn = Nothing
'
''Set objAccessManager = New AccessManagerSDKLib.AccessManager
''Set objUserInfo = objAccessManager.UserInfo
''Set objTerminalUserInfo = objAccessManager.TerminalUserInfo
''Set objTerminalInfo = objAccessManager.TerminalInfo
'
'sIdMember = Format(IdMember, "00000")
'Call objAccessManager.ConnectToServer(593192, strTorniquetes, 7331)
'nErrCode = objAccessManager.errorCode
'If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'    Call objAccessManager.CheckAdmin("0000000", "1234", 2)
'
'    If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'        'validar si existe usuario
'        If ExisteUsuario("68" + sIdMember) Then
'            '>>si existe usuario actualizar datos
'        Call objTerminalInfo.GetTerminalList
'
'        nErrCode = objTerminalInfo.errorCode
'
'        'Dim objTerminalList As AccessManagerSDKLib.ITerminalList
'        Dim tError As Long
'        Dim dRes As Boolean
'
'    tError = 81
'
'        For Each objTerminalList In objTerminalInfo
'            Call objTerminalUserInfo.RemoveUserFromTerminalByUserID(objTerminalList.TerminalID, "68" + sIdMember)
'
'              tError = objTerminalUserInfo.errorCode
'
'            If tError = 805372429 Then
'                '>El usuario ya existe en terminal, continuar.
'                BloqueaAccesoM = 0
'                Else
'                BloqueaAccesoM = objTerminalUserInfo.errorCode
'            End If
'
'        Next
'         End If
'
'
'
'        Call objAccessManager.DisconnectServer
'
'    End If
'Else
'    BloqueaAccesoM = 1
'    Call objAccessManager.DisconnectServer
'    'Exit Sub
'End If
'If nErrCode <> 0 Then
'BloqueaAccesoM = 1
'End If
'   ' Set objUserInfo = Nothing
'    '        Set objTerminalInfo = Nothing
'     '       Set objAccessManager = Nothing
'      '      Set objTerminalUserInfo = Nothing
'       '     Set objTerminalList = Nothing
'Connection_DB
'End Function
'Public Function BloqueaAcceso(IdMember As Long) As Long
'Dim sIdMember As String
'Dim nErrCode As Long
''Set Conn = Nothing
'
''Set objAccessManager = New AccessManagerSDKLib.AccessManager
''Set objUserInfo = objAccessManager.UserInfo
''Set objTerminalUserInfo = objAccessManager.TerminalUserInfo
''Set objTerminalInfo = objAccessManager.TerminalInfo
'
'sIdMember = Format(IdMember, "00000")
'Call objAccessManager.ConnectToServer(593192, strTorniquetes, 7331)
'nErrCode = objAccessManager.errorCode
'If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'    Call objAccessManager.CheckAdmin("0000000", "1234", 2)
'
'    If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'        'validar si existe usuario
'        If ExisteUsuario("68" + sIdMember) Then
'            '>>si existe usuario actualizar datos
'        Call objTerminalInfo.GetTerminalList
'
'        nErrCode = objTerminalInfo.errorCode
'
'        'Dim objTerminalList As AccessManagerSDKLib.ITerminalList
'        Dim tError As Long
'        Dim dRes As Boolean
'
'    tError = 81
'
'        For Each objTerminalList In objTerminalInfo
'            Call objTerminalUserInfo.RemoveUserFromTerminalByUserID(objTerminalList.TerminalID, "68" + sIdMember)
'
'              tError = objTerminalUserInfo.errorCode
'
'            If tError = 805372429 Then
'                '>El usuario ya existe en terminal, continuar.
'                BloqueaAcceso = 0
'                Else
'                BloqueaAcceso = objTerminalUserInfo.errorCode
'            End If
'
'        Next
'         End If
'
'
'
'        Call objAccessManager.DisconnectServer
'
'    End If
'Else
'    BloqueaAcceso = 1
'    Call objAccessManager.DisconnectServer
'    'Exit Sub
'End If
'If nErrCode <> 0 Then
'BloqueaAcceso = 1
'End If
'   ' Set objUserInfo = Nothing
'    '        Set objTerminalInfo = Nothing
'     '       Set objAccessManager = Nothing
'      '      Set objTerminalUserInfo = Nothing
'       '     Set objTerminalList = Nothing
''Connection_DB
'End Function
'
''Agregar Registro en NITGEN
'Public Function AgregaAcceso(IdMember As Long, Nombre As String, Codigo As Long) As Long
'Dim sIdMember As String
'Dim nErrCode As Long
'Set Conn = Nothing
''Set objAccessManager = New AccessManagerSDKLib.AccessManager
''Set objUserInfo = objAccessManager.UserInfo
''Set objTerminalUserInfo = objAccessManager.TerminalUserInfo
''Set objTerminalInfo = objAccessManager.TerminalInfo
'sIdMember = Format(IdMember, "00000")
'Call objAccessManager.ConnectToServer(593192, strTorniquetes, 7331)
'nErrCode = objAccessManager.errorCode
'If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'    Call objAccessManager.CheckAdmin("0000000", "1234", 2)
'
'    If ExisteUsuario("68" + sIdMember) Then
'    AgregaAcceso = 1
'    Call objAccessManager.DisconnectServer
'    Else
'    If objAccessManager.errorCode = ACAMAPI_ERROR_NONE Then
'        objUserInfo.szUserID = "68" + sIdMember
'        objUserInfo.wszUserName = Nombre
'        objUserInfo.GroupID = 0
'        objUserInfo.Privilege = 2
'        objUserInfo.TzCode = 0
'        objUserInfo.wszDepartment = ""
'        objUserInfo.EmployeeNum = ""
'        objUserInfo.Description = ""
'        objUserInfo.dtRegDate = Format(DateTime.Now, "dd/MM/yyyy")
'        objUserInfo.dtExpDate = "31/12/8888"
'
'        objUserInfo.RFCardID = Codigo
'
'        objUserInfo.Authtype = 4
'
'        Call objUserInfo.RegisterUserInfo
'        nErrCode = objUserInfo.errorCode
'
'
'
'        Call objTerminalInfo.GetTerminalList
'        nErrorCode = objTerminalInfo.errorCode
'
'        'Dim objTerminalList As AccessManagerSDKLib.ITerminalList
'        Dim tError As Long
'        Dim dRes As Boolean
'
'    tError = 81
'
'        For Each objTerminalList In objTerminalInfo
'            Call objTerminalUserInfo.AddUserToTerminalByUserID(objTerminalList.TerminalID, "68" + sIdMember, 2, True)
'
'              tError = objTerminalUserInfo.errorCode
'
'            If tError = 805372429 Then
'                '>El usuario ya existe en terminal, continuar.
'            End If
'
'        Next
'
'            AgregaAcceso = objTerminalUserInfo.errorCode
'
'        Call objAccessManager.DisconnectServer
'
'    End If
'    End If
'Else
'    AgregaAcceso = 1
'    Call objAccessManager.DisconnectServer
'    'Exit Sub
'End If
'If nErrCode <> 0 Then
'AgregaAcceso = 1
'End If
'
'   ' Set objUserInfo = Nothing
'    '        Set objTerminalInfo = Nothing
'     '       Set objAccessManager = Nothing
'      '      Set objTerminalUserInfo = Nothing
'       '     Set objTerminalList = Nothing
'Connection_DB
'End Function

'Asigna al socio un numero secuencial para el ctrl de acceso
Public Function AsignaSec(IdSocio As Long, bTemporal As Boolean) As Long
    Dim rsSec As ADODB.Recordset
    Dim nSecuen As Long

    AsignaSec = 0
    nSecuen = 0

    'Busca el maximo numero de secuencial asignado
    InitRecordSet rsSec, "MAX(Secuencial) as MAXIMO", "Secuencial", "IdMember>0", "", Conn
    If (rsSec.RecordCount > 0) Then
        nSecuen = IIf(IsNull(rsSec.Fields("MAXIMO")), 0, rsSec.Fields("MAXIMO"))
    End If
    rsSec.Close
    Set rsSec = Nothing
    
    'Si esta asignado el ultimo valor de la tabla de secuenciales,
    'comienza a reasignar los numeros desde el numero 1
    'If (nSecuen = 65500) Then
    '    nSecuen = 1
    'End If
    
    

    'Recorre los demas registros de la tabla de secuenciales
    'hasta que encuentre uno disponible
    InitRecordSet rsSec, "Secuencial, IdMember", "Secuencial", "IdMember=0 AND Secuencial>" & nSecuen, "", Conn
    Do While (Not rsSec.EOF)
        If (rsSec.Fields("IdMember") = 0) Then
            nSecuen = rsSec.Fields("Secuencial")
            Exit Do
        End If
        rsSec.MoveNext
    Loop
    
    If ((rsSec.EOF) And (nSecuen = 0)) Then
        nSecuen = 1
    End If
    
    rsSec.Close
    Set rsSec = Nothing
    
    If (CambiaSec(IdSocio, nSecuen, bTemporal)) Then
        AsignaSec = nSecuen
    End If
    
'    AsignaSec = CambiaSec(IdSocio, nSecuen)
End Function
'Asigna al socio un numero secuencial para la credencial nueva
Public Function AsignaSecNuevo(IdSocio As Long, bTemporal As Boolean) As Long
    Dim rsSec As ADODB.Recordset
    Dim nSecuen As Long

    AsignaSecNuevo = 0
    nSecuen = 0

    'Busca el maximo numero de secuencial asignado
    InitRecordSet rsSec, "MAX(Secuencial) as MAXIMO", "Secuencial_Nuevo", "IdMember>0", "", Conn
    If (rsSec.RecordCount > 0) Then
        nSecuen = IIf(IsNull(rsSec.Fields("MAXIMO")), 0, rsSec.Fields("MAXIMO"))
    End If
    rsSec.Close
    Set rsSec = Nothing
    
    'Si esta asignado el ultimo valor de la tabla de secuenciales,
    'comienza a reasignar los numeros desde el numero 1
    If (nSecuen = 65500) Then
        nSecuen = 1
    End If

    'Recorre los demas registros de la tabla de secuenciales
    'hasta que encuentre uno disponible
    InitRecordSet rsSec, "Secuencial, IdMember", "Secuencial_Nuevo", "IdMember=0 AND Secuencial>" & nSecuen, "Secuencial", Conn
    Do While (Not rsSec.EOF)
        If (rsSec.Fields("IdMember") = 0) Then
            nSecuen = rsSec.Fields("Secuencial")
            Exit Do
        End If
        rsSec.MoveNext
    Loop
    
    If ((rsSec.EOF) And (nSecuen = 0)) Then
        nSecuen = 1
    End If
    
    rsSec.Close
    Set rsSec = Nothing
    
    If (CambiaSecNuevo(IdSocio, nSecuen, bTemporal)) Then
        AsignaSecNuevo = nSecuen
    End If
    
'    AsignaSec = CambiaSecNuevo(IdSocio, nSecuen)
End Function


'Actualiza las asignaciones de la tabla de Secuencial
Public Function CambiaSec(IdSoc As Long, nSec As Long, bTemporal As Boolean) As Boolean
    Const DATOSSEC = 2
    Dim mFields(DATOSSEC) As String
    Dim mValues(DATOSSEC) As Variant
    Dim sCond As String

    CambiaSec = False

    'Campos
    mFields(0) = "IdMember"
    mFields(1) = "Temporal"
    
    'Valores
    mValues(0) = IdSoc
    mValues(1) = IIf(bTemporal, 1, 0)
    
    sCond = "Secuencial=" & nSec
    
    If (CambiaReg("Secuencial", mFields, DATOSSEC, mValues, sCond, Conn)) Then
        CambiaSec = True
    End If
End Function
'Actualiza las asignaciones de la tabla de Secuencial
Public Function CambiaSecNuevo(IdSoc As Long, nSec As Long, bTemporal As Boolean) As Boolean
Const DATOSSEC = 1
Dim mFields(DATOSSEC) As String
Dim mValues(DATOSSEC) As Variant
Dim sCond As String

    CambiaSecNuevo = False

    'Campos
    mFields(0) = "IdMember"
    
    'Valores
    mValues(0) = IdSoc
    
    sCond = "Secuencial=" & nSec
    
    If (CambiaReg("Secuencial_Nuevo", mFields, DATOSSEC, mValues, sCond, Conn)) Then
        CambiaSecNuevo = True
    End If
End Function


Public Function LeeUltimoFamiliar(lIdTit As Long) As Integer
    
    Dim adorcsNumFam As ADODB.Recordset
    
    strSQL = "SELECT Max( NumeroFamiliar )AS NumFam"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " WHERE IdTitular=" & lIdTit
    
    Set adorcsNumFam = New ADODB.Recordset
    
    adorcsNumFam.CursorLocation = adUseServer
    adorcsNumFam.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If adorcsNumFam.EOF Then
        LeeUltimoFamiliar = 1
    Else
        LeeUltimoFamiliar = IIf(IsNull(adorcsNumFam!numfam), 1, adorcsNumFam!numfam)
    End If
    
    adorcsNumFam.Close
    Set adorcsnumfac = Nothing
    
End Function


Public Function RegistraFechasMtto(nidMember As Long, nTipoSocio As Integer, dFechaIni As Date) As Boolean
    Const DATOSFACT = 3
    Dim rsConceptos As ADODB.Recordset
    Dim mFieldsFact(DATOSFACT) As String
    Dim mValuesFact(DATOSFACT) As Variant
    Dim nMes As Byte
    Dim nAnio As Integer
    Dim dFecha As Date
    Dim dUltdiaMes As Date

    RegistraFechasMtto = False
    
    InitRecordSet rsConceptos, "idConcepto", "Concepto_Tipo", "idTipoUsuario=" & nTipoSocio, "", Conn
    
    With rsConceptos
        If (.RecordCount > 0) Then
            'Campos de la tabla Fechas_Usuario
            mFieldsFact(0) = "idMember"
            mFieldsFact(1) = "idConcepto"
            mFieldsFact(2) = "FechaUltimoPago"
            
            'Número de socio
            mValuesFact(0) = nidMember
            
            'Ultimo día del mes anterior (fecha inicial para pagos)
            nAnio = Year(dFechaIni)
            nMes = Month(dFechaIni) - 1
            
            If (nMes = 0) Then
                nMes = 12
                nAnio = Year(dFechaIni) - 1
            End If
            
            dUltdiaMes = UltimoDiaDelMes(dFechaIni)
            
'            Select Case DateDiff("d", dUltdiaMes, dFechaIni)
'                Case Is < 0
'                    mValuesFact(2) = Format(UltimoDiaDelMes(CDate("01/" & Trim$(Str(nMes)) & "/" & Trim$(Str(nAnio)))), "dd/mm/yyyy")
'                Case Is >= 0
'            End Select
            
            
            #If SqlServer_ Then
                mValuesFact(2) = Format(UltimoDiaDelMes(CDate("01/" & Trim$(Str(nMes)) & "/" & Trim$(Str(nAnio)))), "yyyymmdd")
            #Else
                mValuesFact(2) = Format(UltimoDiaDelMes(CDate("01/" & Trim$(Str(nMes)) & "/" & Trim$(Str(nAnio)))), "dd/mm/yyyy")
            #End If
                
            .MoveFirst
            Do While (Not .EOF)
                'Concepto facturable
                mValuesFact(1) = .Fields("idConcepto")
                
                EliminaReg "FECHAS_USUARIO", "IdMember=" & nidMember, "Elimina tipos conceptos", Conn
                
                If (Not AgregaRegistro("Fechas_Usuario", mFieldsFact, DATOSFACT, mValuesFact, Conn)) Then
                    Exit Do
                End If
            
                .MoveNext
            Loop
            
            If (Not .EOF) Then
                .Close
                Set rsConceptos = Nothing
                Exit Function
            End If
        Else
            .Close
            Set rsConceptos = Nothing
            Exit Function
        End If
        
        .Close
    End With
    
    Set rsConceptos = Nothing
    
    RegistraFechasMtto = True
End Function


Public Function ActConcFact(nidMember As Integer, nTipoSocAnt As Integer, nTipoSocNvo As Integer) As Boolean
Const IGUAL = "I"                       'Permanece igual
Const BORRA = "B"                       'Se debe borrar el concepto
Const REFRESH = "R"                     'Se debe actualizar la fecha de pago
Const AGREGAR = "A"                     'Se agrega el concepto
Const DATOSFACT = 3

Dim rsAnteriores As ADODB.Recordset
Dim rsConceptosNvos As ADODB.Recordset
Dim mClaves() As Integer
Dim mFechas() As Date
Dim mStatus() As String
Dim nPos As Byte
Dim nTam As Byte
Dim nMes As Byte
Dim nAnio As Integer
Dim dFecha As Date

Dim mFields(DATOSFACT) As String
Dim mValues(DATOSFACT) As Variant


    ActConcFact = False

    'Lista de conceptos relacionados con el tipo de usuario anterior
    InitRecordSet rsAnteriores, "idConcepto, FechaUltimoPago", "Fechas_Usuario", "idMember=" & nidMember, "idConcepto", Conn
    
    'Llena la matriz de conceptos con los conceptos del tipo anterior
    With rsAnteriores
        If (.RecordCount > 0) Then
            nTam = 0
        
            .MoveFirst
            Do While (Not .EOF)
                'Agrega el concepto
                ReDim Preserve mClaves(nTam + 1)
                mClaves(nTam) = .Fields("idConcepto")
                
                'Agrega la última fecha de pago del concepto
                ReDim Preserve mFechas(nTam + 1)
                mFechas(nTam) = .Fields("FechaUltimoPago")
                
                'Agrega a la matriz de status la bandera
                ReDim Preserve mStatus(nTam + 1)
                mStatus(nTam) = BORRA
                
                nTam = nTam + 1
            
                .MoveNext
            Loop
        End If
        
        .Close
    End With
    Set rsAnteriores = Nothing
    
    
    'Recorre la lista de conceptos relacionados con el nuevo tipo de usuario
    'y los compara contra la lista de conceptos anteriores
    InitRecordSet rsConceptosNvos, "idConcepto", "Concepto_Tipo", "idTipoUsuario=" & nTipoSocNvo, "idConcepto", Conn
    
    With rsConceptosNvos
        If (.RecordCount > 0) Then
        
            .MoveFirst
            Do While (Not .EOF)
                For nPos = 0 To (nTam - 1)
                    'Si existe el concepto => conserva el concepto y la fecha
                    If (mClaves(nPos) = .Fields("idConcepto")) Then
                        mStatus(nPos) = IGUAL
                        Exit For
                    End If
                Next nPos
                
                'Si no existe el concepto entoces lo agrega
                If (nPos >= nTam) Then
                    'Ultimo día del mes anterior (fecha inicial para pagos)
                    nAnio = Year(Date)
                    nMes = Month(Date) - 1
                    
                    If (nMes = 0) Then
                        nMes = 12
                        nAnio = Year(Date) - 1
                    End If
                
                    ReDim Preserve mClaves(nTam + 1)
                    mClaves(nTam) = .Fields("idConcepto")
                    
                    'Agrega la última fecha de pago del concepto
                    ReDim Preserve mFechas(nTam + 1)
                    #If SqlServer_ Then
                        mFechas(nTam) = Format(UltimoDiaDelMes(CDate("01/" & Trim$(Str(nMes)) & "/" & Trim$(Str(nAnio)))), "dd/mm/yyyy")
                    #Else
                        mFechas(nTam) = Format(UltimoDiaDelMes(CDate("01/" & Trim$(Str(nMes)) & "/" & Trim$(Str(nAnio)))), "dd/mm/yyyy")
                    #End If
                    
                    'Agrega a la matriz de status la bandera
                    ReDim Preserve mStatus(nTam + 1)
                    mStatus(nTam) = AGREGAR
                    
                    nTam = nTam + 1
                End If
            
                .MoveNext
            Loop
        End If
        
        .Close
    End With
    Set rsConceptosNvos = Nothing
    
    'Campos de la tabla Fechas_Usuario
    mFields(0) = "idMember"
    mFields(1) = "idConcepto"
    mFields(2) = "FechaUltimoPago"
    
    'Valores de la tabla Fechas_Usuario
    mValues(0) = nidMember
    
    'Recorre el arreglo de los conceptos y actualiza
    'los registros de acuerdo al status
    For nPos = 0 To (nTam - 1)
        Select Case mStatus(nPos)
            Case "B"
                If (Not EliminaReg("Fechas_Usuario", "idMember=" & nidMember & " AND idConcepto=" & mClaves(nPos), "", Conn)) Then
                    Exit For
                End If
                
            Case "A"
                mValues(1) = mClaves(nPos)
                #If SqlServer_ Then
                    mValues(2) = Format(mFechas(nPos), "yyyymmdd")
                #Else
                    mValues(2) = mFechas(nPos)
                #End If
                
                If (Not AgregaRegistro("Fechas_Usuario", mFields, DATOSFACT, mValues, Conn)) Then
                    Exit For
                End If
            
        End Select
    Next nPos
    
    If (nPos >= nTam) Then
        ActConcFact = True
    End If
End Function
