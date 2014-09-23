Attribute VB_Name = "Inicializa"
Option Explicit
Option Compare Binary

Type xImpPara
    Parametro As Variant
End Type

Type xImpresion
    Nombre As String
    Titulo As String
    Opcion As String
    xImpPara(9) As Variant
End Type



Public Conn As ADODB.Connection
Public strSQL As String
Public strConn As String
Public Const LB_SETTABSTOPS = &H192

Global sDB_User As String                   ' USUARIO
Global sDB_PW As String                    ' PASSWORD
Global sDB_IdUser As String                ' Clave interna del Usuario mas la Hora, Minutos y Segundos
Global sDB_NivelUser As Byte              ' Clave de nivel de seguridad de usuarios
Global sDB As String                             '  BASE DE DATOS
Global sDB_DataSource As String          ' RUTA DE LA BASE DE DATOS
Global sDB_ReportSource As String       ' RUTA DE LOS REPORTES
Global LoginOk As Boolean       ' Variable de control de la forma de login para acceso

Global sG_RutaFoto As String        'Ruta para buscar fotografias

Global lNumFacIniImp As Long
Global lNumFacFinImp As Long

Global lNumRecIniImp As Long
Global lNumRecFinImp As Long

'29/06/2007
Global iNumeroCaja As Integer 'Caja
Global sSerieFactura As String  'Serie para las facturas

Global lNumFolioFacIniImp As String
Global lNumFolioFacFinImp As String

Global sReportes As xImpresion   ' Arreglo que contiene los parametros del reporte

' Inicializacion de librerias para acceso a archivos INI
Declare Function GetPrivateProfileStringA Lib "kernel32.dll" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileStringA Lib "kernel32.dll" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lplFileName As String) As Integer
Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub CentraForma(Objeto1 As Object, Objeto2 As Object)
    Dim Alto, Ancho
    Alto = Objeto1.ScaleHeight
    Ancho = Objeto1.ScaleWidth
    Alto = (Alto - Objeto2.Height) / 2
    Ancho = (Ancho - Objeto2.Width) / 2
    Objeto2.Move Ancho, Alto
End Sub

Public Function Connection_DB() As Boolean
    Dim IntError As Double
    On Error GoTo ErrorCon
    
    ' Inicializa Variables de conexion a la base de datos
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & sDB_DataSource & "\" & "mc.mdb" & ";" & _
              "Persist Security Info=False;" & _
              "Jet OLEDB:Database Password=eUdomilia2006;" & _
              "User Id=Admin;" & _
              "Password=" & sDB_PW & ";"
              
    'SQL 2005
    'strConn = "Provider=SQLNCLI.1;Password=kala1228;Persist Security Info=True;User ID=sa;Initial Catalog=KALACLUB_SQL;Data Source=server"
              

    ' Crea una nueva conexion a la base de datos
    Set Conn = New ADODB.Connection
    
    Conn.Errors.Clear
    Err.Clear
    
    Conn.CursorLocation = adUseServer
    Conn.Open strConn
    
    If Conn.Errors.Count > 0 Then
        Connection_DB = False
    End If

    Connection_DB = True
    
    Exit Function
    
ErrorCon:

    IntError = Conn.Errors.Item(0).NativeError
    Select Case IntError
        Case 18456
            MsgBox "Login o Password Invalido " & Conn.Errors.Item(0).NativeError & Chr(13) & Conn.Errors.Item(0) & Chr(13) & Err.Description, vbCritical, "Conection DataBase"
        Case Else
            MsgBox "Error en Conexión DB: " & Conn.Errors.Item(0).NativeError & Chr(13) & Conn.Errors.Item(0), vbCritical, "Conection DataBase"
    End Select
    
    Connection_DB = False
    
End Function

Public Function CREA_INI() As Boolean
    Dim Regresa As Integer
    Dim ArchivoIni As String
    Dim Mensaje As String
    Dim Rpath As String
    
    Rpath = App.Path
    
    ArchivoIni = Rpath + "\KALACLUB.INI"
        
    sDB = "kalaclub.mdb"
    sDB_DataSource = Rpath
    sDB_User = ""
    
    sG_RutaFoto = ""
    
    Regresa = WritePrivateProfileStringA("PARAMETROS", "BASE", sDB, ArchivoIni)
    If Regresa = 0 Then
        Mensaje = "Base de Datos" + Chr(10) + Chr(13)
    End If

    Regresa = WritePrivateProfileStringA("PARAMETROS", "RUTADB", sDB_DataSource, ArchivoIni)
    If Regresa = 0 Then
        Mensaje = Mensaje + "RUTA DE LA BASE DE DATOS" + Chr(10) + Chr(13)
    End If
    
    Regresa = WritePrivateProfileStringA("PARAMETROS", "USUARIO", sDB_User, ArchivoIni)
    If Regresa = 0 Then
        Mensaje = Mensaje + "Usuario" + Chr(10) + Chr(13)
    End If
    
    Regresa = WritePrivateProfileStringA("PARAMETROS", "RUTAFOTO", sG_RutaFoto, ArchivoIni)
    If Regresa = 0 Then
        Mensaje = Mensaje + "Usuario" + Chr(10) + Chr(13)
    End If

    'Si hay errores en los parametros del archivo ini
    If Len(Trim(Mensaje)) > 0 Then
        Mensaje = "A ocurrido un error mientras se creaba KALACLUB.INI" + Chr(10) + Chr(13) + Mensaje
        MsgBox Mensaje, vbCritical, "Creación KALACLUB.INI"
        CREA_INI = False
    Else
        MsgBox "KALACLUB.INI se ha creado con éxito", vbInformation, "Crea KALACLUB.INI"
        CREA_INI = True
    End If
End Function

Public Sub EndConn_DB()
    If strConn <> "" Then
        If Conn.State > 0 Then
            Conn.Close
            Set Conn = Nothing
        End If
    End If
End Sub

' FUNCIÓN GENERA MENSAJE DE ERROR
' Objetivo: GENERA UN MENSAJE DE ERROR
' Autor:
' Fecha: 24 DE OCTUBRE DE 2002

Function GeneraMensajeError(ByVal PnumNumeroError As Long, Optional PboInterface As Boolean)
    ' Esta función recibe como parámetros el número de error generado.
    ' El segundo parámetro es un opcional el cual debe activarse cuando se trate de una  interface
    If IsMissing(PboInterface) Then
        WriteLog Err.Description, True
        Err.Raise (PnumNumeroError)
    ElseIf PboInterface = True Then
        If PnumNumeroError < 0 Then
            MensajeError LoadResString(PnumNumeroError - vbObjectError)
        Else
            MensajeError Err.Description
        End If
    Else
        Conn.RollbackTrans
        WriteLog Err.Description, True
        Err.Raise (PnumNumeroError)
    End If
End Function

Public Function Lee_Ini() As Byte
    Dim IniString   As String
    Dim ArchivoIni  As String
    Dim ExistFile   As String
    Dim Regresa     As Integer
    Dim Mensaje     As String
    Dim Rpath       As String
    
    Mensaje = ""
    Rpath = App.Path
    ArchivoIni = Rpath + "\" + "KALACLUB.INI"
    
    
    
    'Checa si existe el archivo INI en %APPDATA%
    Rpath = Trim(Environ("APPDATA")) & "\KALACLUB\"
    ArchivoIni = Rpath + "\" + "KALACLUB.INI"
    ExistFile = Dir$(ArchivoIni, vbNormal)
    If ExistFile = "" Then
        Rpath = App.Path
        ArchivoIni = Rpath + "\" + "KALACLUB.INI"
    
    
        ' Checa si existe el archivo INI en archivos de programa
        ExistFile = Dir$(ArchivoIni, vbNormal)
        If ExistFile = "" Then
            MsgBox "KALACLUB.INI no existe", vbCritical, "Lee INI"
            Lee_Ini = 1
            Exit Function
        End If
    End If
    
    ' Obtiene la información del archivo INI de la base de datos
    IniString = String(255, " ")
    Regresa = GetPrivateProfileStringA("PARAMETROS", "BASE", " ", IniString, Len(IniString), ArchivoIni)
    If Regresa = 0 Then
        Mensaje = "BASE DE DATOS" + Chr(10) + Chr(13)
    End If
    sDB = Trim(UCase(IniString))
    If Asc(Right(sDB, 1)) = 32 Or Asc(Right(sDB, 1)) = 0 Then
        sDB = Left(sDB, Len(Trim(sDB)) - 1)
    End If

    IniString = String(255, " ")
    Regresa = GetPrivateProfileStringA("PARAMETROS", "RUTADB", " ", IniString, Len(IniString), ArchivoIni)
    If Regresa = 0 Then
        Mensaje = Mensaje + "RUTA DE LA BASE DE DATOS" + Chr(10) + Chr(13)
    End If
    sDB_DataSource = Trim(UCase(IniString))
    If Asc(Right(sDB_DataSource, 1)) = 32 Or Asc(Right(sDB_DataSource, 1)) = 0 Then
        sDB_DataSource = Left(sDB_DataSource, Len(Trim(sDB_DataSource)) - 1)
    End If
    
    IniString = String(255, " ")
    Regresa = GetPrivateProfileStringA("PARAMETROS", "RUTAREP", " ", IniString, Len(IniString), ArchivoIni)
    If Regresa = 0 Then
        Mensaje = "RUTA DE LOS REPORTES" + Chr(10) + Chr(13)
    End If
    sDB_ReportSource = Trim(UCase(IniString))
    If Asc(Right(sDB_ReportSource, 1)) = 32 Or Asc(Right(sDB_ReportSource, 1)) = 0 Then
        sDB_ReportSource = Left(sDB_ReportSource, Len(Trim(sDB_ReportSource)) - 1)
    End If
    
    IniString = String(255, " ")
    Regresa = GetPrivateProfileStringA("PARAMETROS", "USUARIO", " ", IniString, Len(IniString), ArchivoIni)
    If Regresa = 0 Then
        Mensaje = Mensaje + "USUARIO" + Chr(10) + Chr(13)
    End If
    sDB_User = Trim(IniString)
    If Asc(Right(sDB_User, 1)) = 32 Or Asc(Right(sDB_User, 1)) = 0 Then
        sDB_User = Left(sDB_User, Len(Trim(sDB_User)) - 1)
    End If
    
        
    sDB_User = Trim(IniString)
    If Asc(Right(sDB_User, 1)) = 32 Or Asc(Right(sDB_User, 1)) = 0 Then
        sDB_User = Left(sDB_User, Len(Trim(sDB_User)) - 1)
    End If
    
    '----------------------------------------------------------------------------
    'Path de buscar fotos
    IniString = String(255, " ")
    Regresa = GetPrivateProfileStringA("PARAMETROS", "RUTAFOTO", " ", IniString, Len(IniString), ArchivoIni)
    sG_RutaFoto = Trim(UCase(IniString))
    If Asc(Right(sG_RutaFoto, 1)) = 32 Or Asc(Right(sG_RutaFoto, 1)) = 0 Then
        sG_RutaFoto = Left(sG_RutaFoto, Len(Trim(sG_RutaFoto)) - 1)
    End If
    '--------------------------------------------------------------------------
    '29/06/07
    'Serie para las facturas
    '--------------------------------------------------------------------
    IniString = String(255, " ")
    Regresa = GetPrivateProfileStringA("PARAMETROS", "SERIEFACTURA", " ", IniString, Len(IniString), ArchivoIni)
    sSerieFactura = Trim(UCase(IniString))
    If Asc(Right(sSerieFactura, 1)) = 32 Or Asc(Right(sSerieFactura, 1)) = 0 Then
        sSerieFactura = Left(sSerieFactura, Len(Trim(sSerieFactura)) - 1)
    End If
    '--------------------------------------------------------------------
    '29/06/07
    'Numero de caja
    '--------------------------------------------------------------------
    IniString = String(255, " ")
    Regresa = GetPrivateProfileStringA("PARAMETROS", "NUMEROCAJA", " ", IniString, Len(IniString), ArchivoIni)
    IniString = Trim(IniString)
    If Asc(Right(IniString, 1)) = 32 Or Asc(Right(IniString, 1)) = 0 Then
        IniString = Left(IniString, Len(Trim(IniString)) - 1)
    End If
    
    If IniString = vbNullString Then
        iNumeroCaja = 0
    Else
        iNumeroCaja = Val(IniString)
    End If
    '--------------------------------------------------------------------
    
    If Len(Trim(Mensaje)) > 0 Then
        Mensaje = "Los siguientes valores en KALACLUB.INI están vacios, favor de verificarlos" + Chr(10) + Chr(13) + Mensaje
        MsgBox Mensaje, vbCritical, "LEE INI"
        Lee_Ini = 2
        Exit Function
    End If
    
   Lee_Ini = 0
End Function

Sub LlenaCombos(cboControl As Control, strQuery, strCampo, strCampoAlt As String)
    Dim AdoRcsCombos As ADODB.Recordset
    
    Set AdoRcsCombos = New ADODB.Recordset
    AdoRcsCombos.ActiveConnection = Conn
    AdoRcsCombos.CursorLocation = adUseClient
    AdoRcsCombos.CursorType = adOpenDynamic
    AdoRcsCombos.LockType = adLockReadOnly
    AdoRcsCombos.Open strQuery
    cboControl.Clear
    Do While Not AdoRcsCombos.EOF
        cboControl.AddItem AdoRcsCombos.Fields(strCampo)
        'Llena el campo alterno del combo
        If strCampoAlt <> "" Then
            cboControl.ItemData(cboControl.NewIndex) = AdoRcsCombos.Fields(strCampoAlt)
        End If
        AdoRcsCombos.MoveNext
   Loop
End Sub

' SUBRUTINA MENSAJE DE ERROR
' Objetivo: MUESTRA UN CUADRO DE DIÁLOGO ESPECIFICANDO EL ERROR
' Autor:
' Fecha: 24 DE OCTUBRE DE 2002

Sub MensajeError(PstrMensaje As String)
    On Error GoTo Err_MensajeError
        Screen.MousePointer = vbDefault
        ' pstrMensaje <----- Es el mensaje a desplegar !
        MsgBox PstrMensaje, vbExclamation + vbOKOnly, "¡ Error !"
        Exit Sub
Err_MensajeError:
    GeneraMensajeError Err.Number
End Sub

' SUBRUTINA CREA ARCHIVO .LOG
' Objetivo: GENERA UN ARCHIVO .LOG, CON EL(LOS) ERRORES DE APLICACIÓN
' Autor:
' Fecha: 25 DE OCTUBRE DE 2002

Sub WriteLog(ByVal pstrText As String, Optional ByVal pbolWriteADOErrors As Boolean)
    Dim intLogFile As Integer
    Dim strLogFileName As String
    On Error GoTo Err_WriteLog
    
    ' pstrText: Es la línea que se gravará
    ' pbolWriteADOErrors (Opcional): Si es verdadero, agrega en el archivo
    ' .LOG los errores ocurridos de ADO Errors Code
    
    strLogFileName = App.Path + "\" + "err" + Format(Now, "ddmmyyyy") + ".log"
    intLogFile = (FreeFile)
    Open strLogFileName For Append As #intLogFile
        Print #intLogFile, "Fecha y Hora : " & Trim(CStr(Now()))
        Print #intLogFile, "----------------------------------------------"
        Print #intLogFile, " * * * * * ERRORES DE LA APLICACIÓN * * * * *"
        Print #intLogFile, "   " + pstrText
        Print #intLogFile, "=============================================="
    Close #intLogFile
    Exit Sub
    
Err_WriteLog:
    Exit Sub
End Sub

Sub EliminaAccionista(lngExAccionista As Long)
    Dim iniTrans As Long
    Dim AdoRcsExAcc As ADODB.Recordset
    Dim AdoCmdInserta As ADODB.Command
    On Err GoTo err_EliminaAccionista
    
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    'Obtiene los datos del Accionista que dejará de serlo
    strSQL = "SELECT * FROM accionistas WHERE idproptitulo = " & lngExAccionista
    Set AdoRcsExAcc = New ADODB.Recordset
    AdoRcsExAcc.ActiveConnection = Conn
    AdoRcsExAcc.CursorLocation = adUseClient
    AdoRcsExAcc.CursorType = adOpenDynamic
    AdoRcsExAcc.LockType = adLockReadOnly
    AdoRcsExAcc.Open strSQL
    
    'Inserta la información en la tabla de EXACCIONISTAS
    strSQL = "INSERT INTO exaccionistas (fecha_alta, a_paterno, a_materno, " & _
                    "nombre, calle, colonia, ent_federativa, delegamunici, telefono_1, " & _
                    "telefono_2, empresa, e_calle, e_colonia, e_ent_federativa, " & _
                    "e_delegamunici, e_telefono_1, e_telefono_2, fecha_baja) " & _
                    "VALUES ('" & AdoRcsExAcc!fecha_alta & "', '" & _
                    AdoRcsExAcc!A_Paterno & "', '" & AdoRcsExAcc!A_Materno & "', '" & _
                    AdoRcsExAcc!Nombre & "', '" & AdoRcsExAcc!calle & "', '" & _
                    AdoRcsExAcc!colonia & "', '" & AdoRcsExAcc!ent_federativa & "', '" & _
                    AdoRcsExAcc!delegamunici & "', '" & AdoRcsExAcc!telefono_1 & "', '" & _
                    AdoRcsExAcc!telefono_2 & "', '" & AdoRcsExAcc!Empresa & "', '" & _
                    AdoRcsExAcc!e_calle & "', '" & AdoRcsExAcc!e_colonia & "', '" & _
                    AdoRcsExAcc!e_ent_federativa & "', '" & _
                    AdoRcsExAcc!e_delegamunici & "', '" & _
                    AdoRcsExAcc!e_telefono_1 & "', '" & _
                    AdoRcsExAcc!e_telefono_2 & "', '" & _
                    Format(Now, "dd/mm/yyyy") & "')"
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    
    'Elimina el Accionista de la Tabla de ACCIONISTAS
    strSQL = "DELETE FROM accionistas WHERE idproptitulo = " & lngExAccionista
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
                    
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    Exit Sub
    
err_EliminaAccionista:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

' FUNCIÓN MUESTRA ELEMENTO COMBO
' Objetivo: MOSTRAR UN ELEMENTO DE UN COMBO PRELLENADO
' Autor:
' Fecha: 11 DE JUNIO DE 2004

Public Function MuestraElementoCombo(ctlCombo As Control, varElemento As Variant, Optional strEtqMsg As String = vbNullString) As Boolean
'Muestra un elemento del un combo previamente lleno
'ctlCombo       --->    Nombre del combo
'varElemento  --->    Elemento del combo que se desea mostrar.
'                               Si es una cadena se realiza la búsqueda en la proiedad list.
'                               Si es un número se realiza la búsqueda en item data.
'strEtqMsg      -->     Es la etiqueta del combo.  Se utiliza para mostrar un mensaje o no en caso de que no se encuentre el elemento en el combo

  Dim intI As Integer
    If varElemento = "" Then Exit Function
    
    ctlCombo.ListIndex = -1
    MuestraElementoCombo = True
    
    If IsNumeric(varElemento) Then
    ' Realiza la búsqueda con la propiedad ItemData
    For intI = 0 To ctlCombo.ListCount - 1
      ctlCombo.ListIndex = intI
      If ctlCombo.ItemData(intI) = CLng(varElemento) Then Exit Function
    Next
  Else
  ' Realiza la búsqueda con la propiedad List
    For intI = 0 To ctlCombo.ListCount - 1
      If Trim(ctlCombo.List(intI)) = Trim(varElemento) Then
        ctlCombo.ListIndex = intI
        Exit Function
      End If
    Next
  End If
  If Not IsNull(strEtqMsg) Then
    ctlCombo.ListIndex = -1
    MuestraElementoCombo = False
  End If
End Function
