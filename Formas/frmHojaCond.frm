VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmHojaCond 
   Caption         =   "Hoja de Condiciones"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8445
   Icon            =   "frmHojaCond.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8445
   Begin CRVIEWERLibCtl.CRViewer crvFrente 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmHojaCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public nidTitular As Single


Dim crxApplication As New CRAXDRT.Application
Dim crxFrente As New CRAXDRT.Report
Dim crxReves As New CRAXDRT.Report

Dim rsFrente As ADODB.Recordset
Dim rsSubMembresia As ADODB.Recordset
Dim rsReves As ADODB.Recordset
Dim rsSub1Reves As ADODB.Recordset
Dim rsSub2Reves As ADODB.Recordset

'27/12/2007
Dim adorsEmerDat As ADODB.Recordset


Dim sFrente As String
Dim sSubMembresia As String
Dim sReves As String
Dim sSub1Reves As String
Dim sSub2Reves As String

Dim rsDirPar As ADODB.Recordset
Dim rsDirFis As ADODB.Recordset
Dim sDirPar As String
Dim sDirFis As String


'Variables para el subreporte
Dim crxDatabase As CRAXDRT.Database
Dim crxDatabaseTables As CRAXDRT.DatabaseTables
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section
Dim crxReportObjs As CRAXDRT.ReportObjects
Dim crxSubreportObj As CRAXDRT.SubreportObject

Dim CrxFormulaFields As CRAXDRT.FormulaFieldDefinitions
Dim CrxFormulaField As CRAXDRT.FormulaFieldDefinition


Dim crxSubFrente As CRAXDRT.Report
Dim crxSub1Reves As CRAXDRT.Report
Dim crxSub2Reves As CRAXDRT.Report

Private Sub Form_Activate()
    
    
    On Error GoTo Error_Catch
    Err.Clear
    
    
    If Me.Tag = "LOADED" Then Exit Sub
    
    Me.Tag = "LOADED"
    
    CentraForma MDIPrincipal, Me

    Screen.MousePointer = vbHourglass
    
    LlenaRecordSet
    
    If IsNull(rsFrente.Fields("MantenimientoIni")) Then
        MsgBox "No existen datos para generar el reporte!", vbCritical, "Error"
        Screen.MousePointer = vbDefault
        Unload Me
        Exit Sub
    End If
    
    
    MostrarFrente
    'MostrarReves
    Screen.MousePointer = vbDefault
    
    Me.crvFrente.Zoom 100
    
    Exit Sub

Error_Catch:

    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MsgError
        Unload Me
    End If

    
End Sub

Private Sub LlenaRecordSet()
    'Cadena que utiliza el reporte del frente
    sFrente = "SELECT DISTINCT USUARIOS.NoFamilia, USUARIOS.idTipoUsuario, TIPOMEM.Descripcion, "
    sFrente = sFrente & "USUARIOS.Nombre, USUARIOS.A_Paterno, USUARIOS.A_MATERNO, USUARIOS.FechaNacio, "
    sFrente = sFrente & "USUARIOS.Profesion, USUARIOS.Email, "
    sFrente = sFrente & "MEMBRESIA.FechaAlta, MEMBRESIA.Duracion, MEMBRESIA.Monto, MEMBRESIA.Enganche, MEMBRESIA.NumeroPagos, "
    sFrente = sFrente & "USUARIOS.idMember, MEMBRESIA.idMembresia, MEMBRESIA.idTipoMembresia, MEMBRESIA.NombrePropietario, "
    sFrente = sFrente & "MEMBRESIA.MantenimientoIni, MEMBRESIA.Observaciones "
    sFrente = sFrente & "FROM ((Usuarios_Club AS USUARIOS LEFT JOIN Direcciones AS DIRECCION ON USUARIOS.idMember=DIRECCION.idMember) "
    sFrente = sFrente & "LEFT JOIN Membresias AS MEMBRESIA ON USUARIOS.idMember=MEMBRESIA.idMember) "
    sFrente = sFrente & "LEFT JOIN Tipo_Membresia AS TIPOMEM ON MEMBRESIA.idTipoMembresia=TIPOMEM.idTipoMembresia "
    sFrente = sFrente & "WHERE USUARIOS.idMember=" & nidTitular
    
'    sFrente = sFrente & "WHERE USUARIOS.idMember=" & frmConsMembers.ssdbMembresia.Columns(5).Value
    
    'Crea la instancia del recordset
    Set rsFrente = New ADODB.Recordset
    
    'Asigna sus propiedades del recodset
    With rsFrente
        .Source = sFrente
        .ActiveConnection = Conn
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With
    
    
    
    
    'Direccion Particular
    Set rsDirPar = New ADODB.Recordset
    
    sDirPar = "SELECT TOP 1 RazonSocial, Calle, Colonia, CodPos, DelOMuni, Ciudad, Estado, Tel1, Tel2, RFC"
    sDirPar = sDirPar & " FROM DIRECCIONES"
    sDirPar = sDirPar & " WHERE"
    sDirPar = sDirPar & " IdMember=" & nidTitular
    sDirPar = sDirPar & " AND idTipoDireccion=1"
    'Asigna sus propiedades del recodset
    With rsDirPar
        .Source = sDirPar
        .ActiveConnection = Conn
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseServer
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With
    
    'Direccion Fiscal
    Set rsDirFis = New ADODB.Recordset
    
    sDirFis = "SELECT TOP 1 RazonSocial, Calle, Colonia, CodPos, DelOMuni, Ciudad, Estado, Tel1, Tel2, RFC"
    sDirFis = sDirFis & " FROM DIRECCIONES"
    sDirFis = sDirFis & " WHERE"
    sDirFis = sDirFis & " IdMember=" & nidTitular
    sDirFis = sDirFis & " AND idTipoDireccion=3"
    'Asigna sus propiedades del recodset
    With rsDirFis
        .Source = sDirFis
        .ActiveConnection = Conn
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseServer
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With
    
    
    'Cadena que utiliza el subreporte del frente
    sSubMembresia = "SELECT DETALLE.Monto, DETALLE.FechaVence, DETALLE.NoPago "
    sSubMembresia = sSubMembresia & " FROM Detalle_Mem AS DETALLE LEFT JOIN Membresias ON DETALLE.idMembresia=Membresias.idMembresia "
    sSubMembresia = sSubMembresia & " WHERE Membresias.idMember=" & nidTitular
    sSubMembresia = sSubMembresia & " ORDER BY DETALLE.NoPago"
    
'    sSubMembresia = sSubMembresia & "WHERE DETALLE.NoPago>0 AND Membresias.idMember=" & frmConsMembers.ssdbMembresia.Columns(5).Value
    
    'Crea la instancia del recordset
    Set rsSubMembresia = New ADODB.Recordset
    
    'Asigna sus propiedades del recordset
    With rsSubMembresia
        .Source = sSubMembresia
        .ActiveConnection = Conn
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With


    'Cadena que utiliza el reporte del revés
    sReves = "SELECT USUARIOS.FechaNacio, USUARIOS.Profesion, USUARIOS.idPais, "
    sReves = sReves & "Paises.Pais, USUARIOS.UFechaPago, "
    sReves = sReves & "(SELECT Direcciones.RazonSocial FROM Direcciones WHERE Direcciones.idTipoDireccion=3 AND Direcciones.idMember=USUARIOS.idMember) AS EMPRESA, "

    #If SqlServer_ Then
        sReves = sReves & "(SELECT LTRIM(RTRIM(ISNULL(Direcciones.Calle, '')))"
        sReves = sReves & " + CASE WHEN LTRIM(RTRIM(ISNULL(Direcciones.Colonia,''))) = '' THEN '' ELSE ', ' + LTRIM(RTRIM(Direcciones.Colonia)) END"
        sReves = sReves & " + CASE WHEN LTRIM(RTRIM(ISNULL(Direcciones.CodPos,''))) = '' THEN '' ELSE ', ' + LTRIM(RTRIM(CONVERT(nvarchar, Direcciones.CodPos))) END"
'        sReves = sReves & " + CASE WHEN LTRIM(RTRIM(ISNULL(Direcciones.CodPos,''))) = '' THEN '' ELSE ', ' + LTRIM(RTRIM(CatalogoCodigoPostal.CodigoPostal)) END"
        sReves = sReves & " + CASE WHEN LTRIM(RTRIM(ISNULL(Direcciones.DeloMuni,''))) = '' THEN '' ELSE ', ' + LTRIM(RTRIM(Direcciones.DeloMuni)) END"
        sReves = sReves & " + CASE WHEN LTRIM(RTRIM(ISNULL(Direcciones.Ciudad,''))) = '' THEN '' ELSE ', ' + LTRIM(RTRIM(Direcciones.Ciudad)) END"
        sReves = sReves & " + CASE WHEN LTRIM(RTRIM(ISNULL(Direcciones.Estado,''))) = '' THEN '' ELSE ', ' + LTRIM(RTRIM(Direcciones.Estado)) END"
        sReves = sReves & " FROM Direcciones"
'        sReves = sReves & " INNER JOIN CatalogoCodigoPostal"
'        sReves = sReves & " ON DIRECCIONES.CodPos = CatalogoCodigoPostal.IdCodigoPostal"
        sReves = sReves & " WHERE"
        sReves = sReves & " Direcciones.idTipoDireccion=3"
        sReves = sReves & " AND Direcciones.idMember=USUARIOS.idMember"
        sReves = sReves & ") AS DOMICILIO, "
        
        sReves = sReves & "(SELECT"
        sReves = sReves + " LTRIM(RTRIM(ISNULL(Direcciones.Tel1, '' )))"
        sReves = sReves + " + CASE WHEN LTRIM(RTRIM(ISNULL(Direcciones.Tel2, ''))) = '' THEN '' ELSE ', ' + LTRIM(RTRIM(Direcciones.Tel2)) END"
        sReves = sReves + " FROM Direcciones"
        sReves = sReves + " WHERE"
        sReves = sReves + " Direcciones.idTipoDireccion=3"
        sReves = sReves + " AND Direcciones.idMember=USUARIOS.idMember"
        sReves = sReves & ") AS TELS, "
    
        sReves = sReves & "(SELECT"
        sReves = sReves & " LTRIM(RTRIM(ISNULL(Direcciones.Fax,'')))"
        sReves = sReves & " FROM Direcciones"
        sReves = sReves & " WHERE Direcciones.idTipoDireccion=3"
        sReves = sReves & " AND Direcciones.idMember=USUARIOS.idMember"
        sReves = sReves & ") AS FAX, "
    #Else
        sReves = sReves & "(SELECT TRIM(Direcciones.Calle) & "
        sReves = sReves & "iif(NOT ISNULL(Direcciones.Colonia), ', ' & TRIM(Direcciones.Colonia), '' ) & "
        sReves = sReves & "iif(NOT ISNULL(Direcciones.CodPos), ', ' & TRIM(Direcciones.CodPos), '' ) & "
        sReves = sReves & "iif(NOT ISNULL(Direcciones.DeloMuni), ', ' & TRIM(Direcciones.DeloMuni), '' ) & "
        sReves = sReves & "iif(NOT ISNULL(Direcciones.Ciudad), ', ' & TRIM(Direcciones.Ciudad), '' ) & "
        sReves = sReves & "iif(NOT ISNULL(Direcciones.Estado), ', ' & TRIM(Direcciones.Estado), '' ) "
        sReves = sReves & "FROM Direcciones WHERE Direcciones.idTipoDireccion=3 AND Direcciones.idMember=USUARIOS.idMember) AS DOMICILIO, "

        sReves = sReves & "(SELECT iif(NOT ISNULL(Direcciones.Tel1), TRIM(Direcciones.Tel1), '' ) & "
        sReves = sReves & "iif(NOT ISNULL(Direcciones.Tel2), ', ' & TRIM(Direcciones.Tel2), '' ) "
        sReves = sReves & "FROM Direcciones WHERE Direcciones.idTipoDireccion=3 AND Direcciones.idMember=USUARIOS.idMember) AS TELS, "
    
        sReves = sReves & "(SELECT iif(NOT ISNULL(Direcciones.Fax), TRIM(Direcciones.Fax), '' ) "
        sReves = sReves & "FROM Direcciones WHERE Direcciones.idTipoDireccion=3 AND Direcciones.idMember=USUARIOS.idMember) AS FAX, "
    #End If
    
    'gpo 13/10/2005
    sReves = sReves & "(SELECT Membresias.Observaciones "
    sReves = sReves & "FROM Membresias WHERE Membresias.IdMember=USUARIOS.idMember) AS OBSER "

    sReves = sReves & "FROM Usuarios_Club AS USUARIOS LEFT JOIN Paises ON USUARIOS.idPais=Paises.idPais "
    sReves = sReves & "WHERE USUARIOS.idMember=" & nidTitular
    
'    sReves = sReves & "WHERE USUARIOS.idMember=" & frmConsMembers.ssdbMembresia.Columns(5).Value
    
    'Crea la instancia del recordset
    Set rsReves = New ADODB.Recordset
    
    'Asigna sus propiedades del recordset
    With rsReves
        .Source = sReves
        .ActiveConnection = Conn
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With
    
    
    'Cadena que utiliza el subreporte de familiares del reves
    sSub1Reves = "SELECT USUARIOS.idTipoUsuario, TIPOUSER.Descripcion, USUARIOS.Nombre, USUARIOS.A_Paterno, USUARIOS.A_Materno, "
    sSub1Reves = sSub1Reves & "USUARIOS.FechaNacio, USUARIOS.Sexo"
    sSub1Reves = sSub1Reves & " FROM Usuarios_Club AS USUARIOS LEFT JOIN Tipo_Usuario AS TIPOUSER ON USUARIOS.idTipoUsuario=TIPOUSER.idTipoUsuario"
    sSub1Reves = sSub1Reves & " WHERE"
    sSub1Reves = sSub1Reves & " USUARIOS.idTitular=" & nidTitular
    
'    sSub1Reves = sSub1Reves & "WHERE USUARIOS.idMember<>USUARIOS.idTitular AND USUARIOS.idTitular=" & frmConsMembers.ssdbMembresia.Columns(5).Value
    
    'Crea la instancia del recordset
    Set rsSub1Reves = New ADODB.Recordset
    
    'Asigna sus propiedades del recordset
    With rsSub1Reves
        .Source = sSub1Reves
        .ActiveConnection = Conn
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With
    
    
    'Cadena que utiliza el subreporte de las referencias del reves
    sSub2Reves = "SELECT Nombre, A_Paterno, A_Materno, Telefono "
    sSub2Reves = sSub2Reves & "FROM Referencias "
    sSub2Reves = sSub2Reves & "WHERE Referencias.idMember=" & nidTitular
    
'    sSub2Reves = sSub2Reves & "WHERE Referencias.idMember=" & frmConsMembers.ssdbMembresia.Columns(5).Value
    
    'Crea la instancia del recordset
    Set rsSub2Reves = New ADODB.Recordset
    
    'Asigna sus propiedades del recordset
    With rsSub2Reves
        .Source = sSub2Reves
        .ActiveConnection = Conn
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With
    
    'Para los datos de emergencia y beneficiarios
    
    strSQL = "SELECT E.NombreEmergencia, E.ParentescoEmergencia, E.TelefonosEmergencia, E.DomicilioEmergencia, E.NombreBeneficiario1, E.ParentescoBeneficiario1, E.PorcentajeBeneficiario1, E.NombreBeneficiario2, E.ParentescoBeneficiario2, E.PorcentajeBeneficiario2"
    strSQL = strSQL & " FROM EMERGENCIA_DATOS E"
    strSQL = strSQL & " WHERE (((E.IdMember)=" & nidTitular & "))"
    
    Set adorsEmerDat = New ADODB.Recordset
    adorsEmerDat.CursorLocation = adUseServer
    adorsEmerDat.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    

End Sub
Private Sub MostrarFrente()
    Dim sCantLetra As String
    
    Dim sNombreClub As String
    '27/12/2007
    Dim sEmpresa As String
    
    Dim sCadFormaPago As String
    
    Dim sCadRFC As String
    Dim sCadRazonSocial As String
    
    Dim sCadDirPar As String
    Dim sCadTelPar As String
    
    Dim sCadDirFis As String
    
    Dim dImporteMant As Double
    Dim sDirecc As String
    

    Dim X As Long
    Dim Y As Long
    
    
    Dim sEnganche As String
    Dim sPagoInsc01 As String
    Dim sPagoInsc02 As String
    Dim sPagoInsc03 As String
    Dim sPagoInsc04 As String
    Dim sPagoInsc05 As String
    Dim sPagoInsc06 As String
    Dim iPeriodo As Integer
    Dim bTipo As Boolean
    Dim sPeriodo As String
    
    sNombreClub = ObtieneParametro("NOMBRE DEL CLUB")
    '27/12/2007
    sEmpresa = ObtieneParametro("EMPRESA CONTRATO")
    
    
    'sDirecc = ChecaDireccionado(CLng(nidTitular))
    
    
    sDirecc = rsFrente.Fields("MantenimientoIni")
    sPeriodo = ""
    Select Case Left(sDirecc, 1)
                Case "M"
                    iPeriodo = 1
                    sPeriodo = "Mensual"
                Case "B"
                    iPeriodo = 2
                    sPeriodo = "Bimestral"
                Case "T"
                    iPeriodo = 3
                    sPeriodo = "Trimestral"
                Case "S"
                    iPeriodo = 6
                    sPeriodo = "Semestral"
                Case "A"
                    iPeriodo = 12
                    sPeriodo = "Anual"
            End Select
    Select Case Right(sDirecc, 1)
                Case "D"
                    bTipo = True
                    sPeriodo = sPeriodo + " Direccionado"
                Case "C"
                    bTipo = False
                    sPeriodo = sPeriodo + " Convencional"
            End Select
    dImporteMant = CalculaMantenimientoMes(CLng(nidTitular), bTipo, iPeriodo)
    
    sCadRazonSocial = Trim(rsFrente.Fields("Nombre")) & " " & Trim(rsFrente.Fields("A_Paterno")) & " " & Trim(rsFrente.Fields("A_MATERNO"))
    
    'Para la cadena de la direccion particular
    If Not rsDirPar.EOF Then
        sCadDirPar = IIf(IsNull(rsDirPar!calle), "", rsDirPar!calle)
        sCadDirPar = sCadDirPar & " COL." & IIf(IsNull(rsDirPar!colonia), "", rsDirPar!colonia)
        sCadDirPar = sCadDirPar & " C.P." & IIf(IsNull(rsDirPar!Codpos), "", Format(rsDirPar!Codpos, "00000"))
        sCadDirPar = sCadDirPar & " " & IIf(IsNull(rsDirPar!DeloMuni), "", rsDirPar!DeloMuni)
        sCadDirPar = sCadDirPar & ", " & IIf(IsNull(rsDirPar!Ciudad), "", rsDirPar!Ciudad)
        sCadDirPar = sCadDirPar & " " & IIf(IsNull(rsDirPar!Estado), "", rsDirPar!Estado)
        
        sCadTelPar = IIf(IsNull(rsDirPar!Tel1), "", rsDirPar!Tel1)
        sCadTelPar = sCadTelPar & " " & IIf(IsNull(rsDirPar!Tel2), "", rsDirPar!Tel2)
        
        sCadRFC = IIf(IsNull(rsDirPar!rfc), " ", rsDirPar!rfc)
        
    End If
    
    'Para la cadena de la direccion fiscal
    If Not rsDirFis.EOF Then
        sCadDirFis = IIf(IsNull(rsDirFis!calle), "", rsDirFis!calle)
        sCadDirFis = sCadDirFis & " COL." & IIf(IsNull(rsDirFis!colonia), "", rsDirFis!colonia)
        sCadDirFis = sCadDirFis & " C.P." & IIf(IsNull(rsDirFis!Codpos), "", Format(rsDirFis!Codpos, "00000"))
        sCadDirFis = sCadDirFis & " " & IIf(IsNull(rsDirFis!DeloMuni), "", rsDirFis!DeloMuni)
        sCadDirFis = sCadDirFis & ", " & IIf(IsNull(rsDirFis!Ciudad), "", rsDirFis!Ciudad)
        sCadDirFis = sCadDirFis & " " & IIf(IsNull(rsDirFis!Estado), "", rsDirFis!Estado)
        
        sCadRFC = IIf(IsNull(rsDirFis!rfc), " ", rsDirFis!rfc)
        sCadRazonSocial = IIf(IsNull(rsDirFis!RazonSocial), " ", rsDirFis!RazonSocial)
    Else
        sCadDirFis = sCadDirPar
    End If
    
    
    
    'Para la cadena de forma de pago
    If rsFrente.Fields("NumeroPagos") = 0 Then
        sCadFormaPago = "CONTADO"
    Else
        sCadFormaPago = "CREDITO"
        Do Until rsSubMembresia.EOF
            
            Select Case rsSubMembresia.Fields("NoPago")
                Case 0
                    sEnganche = Format(rsSubMembresia.Fields("Monto"), "$#,#0.00")
                Case 1
                    sPagoInsc01 = Format(rsSubMembresia.Fields("Monto"), "$#,#0.00")
                Case 2
                    sPagoInsc02 = Format(rsSubMembresia.Fields("Monto"), "$#,#0.00")
                Case 3
                    sPagoInsc03 = Format(rsSubMembresia.Fields("Monto"), "$#,#0.00")
                Case 4
                    sPagoInsc04 = Format(rsSubMembresia.Fields("Monto"), "$#,#0.00")
                Case 5
                    sPagoInsc01 = Format(rsSubMembresia.Fields("Monto"), "$#,#0.00")
                Case 6
                    sPagoInsc02 = Format(rsSubMembresia.Fields("Monto"), "$#,#0.00")
                    
            End Select
            
            'sCadFormaPago = sCadFormaPago & Format(rsSubMembresia.Fields("DETALLE!FechaVence"), "dd/MMM/yyyy")
            'sCadFormaPago = sCadFormaPago & ", " & Format(rsSubMembresia.Fields("DETALLE!Monto"), "$#,#0.00") & "; "
            rsSubMembresia.MoveNext
        Loop
    End If
    
    sCadFormaPago = Trim(sCadFormaPago)
    If Right(sCadFormaPago, 1) = ";" Then
        sCadFormaPago = Left(sCadFormaPago, Len(sCadFormaPago) - 1)
    End If
    
    

    'Esta linea llama al archivo .rpt que se creo con Crystal Report
    Set crxFrente = crxApplication.OpenReport(sDB_ReportSource & "\HojadeCondiciones.rpt", 1)
    Me.crvFrente.EnableGroupTree = False
    Me.crvFrente.EnableRefreshButton = False

    Set CrxFormulaFields = crxFrente.FormulaFields
    
    CrxFormulaFields.Item(2).Text = "'" & Format(rsFrente.Fields("FechaAlta"), "dd/MMM/yyyy") & "'"
    CrxFormulaFields.Item(3).Text = rsFrente.Fields("NoFamilia")
    CrxFormulaFields.Item(4).Text = "'" & sNombreClub & "'"
    
    CrxFormulaFields.Item(5).Text = "'" & rsFrente.Fields("Descripcion") & "'"
    
    If rsFrente.Fields("Duracion") <= 15 Then
        CrxFormulaFields.Item(19).Text = "'" & Format(DateAdd("yyyy", rsFrente.Fields("Duracion"), rsFrente.Fields("FechaAlta")), "dd/MMM/yyyy") & "'"
    End If
    
    CrxFormulaFields.Item(6).Text = "'" & Trim(rsFrente.Fields("Nombre")) & " " & Trim(rsFrente.Fields("A_Paterno")) & " " & Trim(rsFrente.Fields("A_MATERNO")) & "'"
    CrxFormulaFields.Item(7).Text = "'" & sCadDirPar & "'"
    CrxFormulaFields.Item(8).Text = "'" & sCadTelPar & "'"
    
    CrxFormulaFields.Item(9).Text = "'" & Format(rsFrente("FechaNacio"), "dd/MMM/yyyy") & "'"
    CrxFormulaFields.Item(10).Text = "'" & LCase(rsFrente("Email")) & "'"
    
    'Si no tiene correo electrónico
    If CrxFormulaFields.Item(10).Text = "''" Then
        CrxFormulaFields.Item(10).Text = "'" & Space(1) & "'"
    End If
    
    CrxFormulaFields.Item(11).Text = "'" & sCadRFC & "'"
    CrxFormulaFields.Item(12).Text = "'" & sCadRazonSocial & "'"
    CrxFormulaFields.Item(13).Text = "'" & sCadDirFis & "'"
    
    CrxFormulaFields.Item(14).Text = rsFrente("Monto") 'nTotalMem
    
    CrxFormulaFields.Item(15).Text = "'" & Format(rsReves.Fields("UFechaPago"), "dd/MMM/yyyy") & "'"
    CrxFormulaFields.Item(16).Text = "'" & sCadFormaPago & "'"
    
    CrxFormulaFields.Item(17).Text = dImporteMant
    
    CrxFormulaFields.Item(20).Text = "'" & sPeriodo & "'"
    CrxFormulaFields.Item(21).Text = "'" & IIf(Left(sDirecc, 1) = "A", "X", " ") & "'"
    CrxFormulaFields.Item(22).Text = "'" & IIf(sDirecc = "MD", "X", " ") & "'"
    
    CrxFormulaFields.Item(23).Text = "'" & IIf(IsNull(rsFrente.Fields("Observaciones")), " ", rsFrente.Fields("Observaciones")) & "'"

    'Si no tiene observaciones
    If CrxFormulaFields.Item(23).Text = "''" Then
        CrxFormulaFields.Item(23).Text = "'" & Space(1) & "'"
    End If
    
        
    If Not adorsEmerDat.EOF Then
        CrxFormulaFields.Item(24).Text = "'" & IIf(IsNull(adorsEmerDat!NombreEmergencia), " ", adorsEmerDat!NombreEmergencia) & "'"
        CrxFormulaFields.Item(25).Text = "'" & IIf(IsNull(adorsEmerDat!ParentescoEmergencia), " ", adorsEmerDat!ParentescoEmergencia) & "'"
        CrxFormulaFields.Item(26).Text = "'" & IIf(IsNull(adorsEmerDat!TelefonosEmergencia), " ", adorsEmerDat!TelefonosEmergencia) & "'"
        CrxFormulaFields.Item(27).Text = "'" & IIf(IsNull(adorsEmerDat!DomicilioEmergencia), " ", adorsEmerDat!DomicilioEmergencia) & "'"
    
        CrxFormulaFields.Item(28).Text = "'" & IIf(IsNull(adorsEmerDat!NombreBeneficiario1), " ", adorsEmerDat!NombreBeneficiario1) & "'"
        CrxFormulaFields.Item(29).Text = "'" & IIf(IsNull(adorsEmerDat!ParentescoBeneficiario1), " ", adorsEmerDat!ParentescoBeneficiario1) & "'"
        CrxFormulaFields.Item(30).Text = "'" & IIf(IsNull(adorsEmerDat!PorcentajeBeneficiario1), " ", Format(adorsEmerDat!PorcentajeBeneficiario1, "##0.00%")) & "'"
    
        CrxFormulaFields.Item(31).Text = "'" & IIf(IsNull(adorsEmerDat!NombreBeneficiario2), " ", adorsEmerDat!NombreBeneficiario2) & "'"
        CrxFormulaFields.Item(32).Text = "'" & IIf(IsNull(adorsEmerDat!ParentescoBeneficiario2), " ", adorsEmerDat!ParentescoBeneficiario2) & "'"
        CrxFormulaFields.Item(33).Text = "'" & IIf(IsNull(adorsEmerDat!PorcentajeBeneficiario2), " ", Format(adorsEmerDat!PorcentajeBeneficiario2, "##0.00%")) & "'"
    End If
    'Establece el nombre de la empresa
    CrxFormulaFields.Item(34).Text = "'" & sEmpresa & "'"
    
    'Datos para el pago de inscripcion
    
    
    
    CrxFormulaFields.Item(35).Text = "'" & sPagoInsc01 & "'"
    CrxFormulaFields.Item(36).Text = "'" & sPagoInsc02 & "'"
    CrxFormulaFields.Item(37).Text = "'" & sPagoInsc03 & "'"
    CrxFormulaFields.Item(38).Text = "'" & sPagoInsc04 & "'"
    CrxFormulaFields.Item(39).Text = "'" & sPagoInsc05 & "'"
    CrxFormulaFields.Item(40).Text = "'" & sPagoInsc06 & "'"
    CrxFormulaFields.Item(41).Text = "'" & sEnganche & "'"
    

    'Set crxDatabase = crxFrente.Database
    'Set crxDatabaseTables = crxDatabase.Tables
    'Set crxDatabaseTable = crxDatabaseTables.Item(1)

    'crxDatabaseTable.SetDataSource rsFrente, 3

      'Evita que Crystal intente pedir el valor de los parametros
    crxFrente.EnableParameterPrompting = False
    
    'sCantLetra = Trim$(UCase$(Num2Txt(nTotalMem)))
    'sCantLetra = sCantLetra & " PESOS " & Format((nTotalMem - Int(nTotalMem)) * 100, "#00") & "/100 M. N."
    
'    sCantLetra = Trim$(UCase$(Num2Txt(frmConsMembers.ssdbMembresia.Columns(7).Value)))
'    sCantLetra = sCantLetra & " PESOS " & Format((frmConsMembers.ssdbMembresia.Columns(7).Value - Int(frmConsMembers.ssdbMembresia.Columns(7).Value)) * 100, "#00") & "/100 M. N."
    
    'crxFrente.ParameterFields(1).AddCurrentValue sCantLetra

    'Recorre las secciones del reporte e identifica aquellas
    'que sean subreportes
    Set crxSections = crxFrente.Sections

    For X = 1 To crxSections.Count
        Set crxSection = crxSections.Item(X)

        Set crxReportObjs = crxSection.ReportObjects
        For Y = 1 To crxReportObjs.Count
            If (crxReportObjs.Item(Y).Kind = crSubreportObject) Then
                Set crxSubreportObj = crxReportObjs.Item(Y)
                Set crxSubFrente = crxSubreportObj.OpenSubreport

                Set crxDatabase = crxSubFrente.Database
                Set crxDatabaseTables = crxDatabase.Tables
                Set crxDatabaseTable = crxDatabaseTables.Item(1)

                crxDatabaseTable.SetDataSource rsSub1Reves, 3
                Exit For
            End If
        Next
    Next


    Me.crvFrente.ReportSource = crxFrente
    Me.crvFrente.ViewReport
End Sub






Private Sub Form_Resize()

    

    With Me.crvFrente
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With

'    With Me.crvReves
'        .Top = 0
'        .Left = ScaleWidth / 2
'        .Height = ScaleHeight
'        .Width = ScaleWidth / 2
'        .Zoom (50)
'    End With
    
    
End Sub





