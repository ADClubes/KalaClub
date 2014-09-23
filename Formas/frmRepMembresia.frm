VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmRepMembresia 
   Caption         =   "Impresión del contrato para membresías"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10110
   Icon            =   "frmRepMembresia.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   10110
   Begin CRVIEWERLibCtl.CRViewer crvReves 
      Height          =   6495
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   4575
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
   Begin CRVIEWERLibCtl.CRViewer crvFrente 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
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
Attribute VB_Name = "frmRepMembresia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*  Formulario para imprimir el contrato de las membresías          *
'*  Daniel Hdez                                                     *
'*  09 / Septiembre / 2005                                          *
'*  Ultima actualización: 31 / Octubre / 2005                       *
'********************************************************************


Public nidTitular As Single
Public nTotalMem As Double


Dim crxApplication As New CRAXDRT.Application
Dim crxFrente As New CRAXDRT.Report
Dim crxReves As New CRAXDRT.Report

Dim rsFrente As ADODB.Recordset
Dim rsSubFrente As ADODB.Recordset
Dim rsReves As ADODB.Recordset
Dim rsSub1Reves As ADODB.Recordset
Dim rsSub2Reves As ADODB.Recordset
Dim sFrente As String
Dim sSubFrente As String
Dim sReves As String
Dim sSub1Reves As String
Dim sSub2Reves As String

'Variables para el subreporte
Dim crxDatabase As CRAXDRT.Database
Dim crxDatabaseTables As CRAXDRT.DatabaseTables
Dim crxDatabaseTable As CRAXDRT.DatabaseTable
Dim crxSections As CRAXDRT.Sections
Dim crxSection As CRAXDRT.Section
Dim crxReportObjs As CRAXDRT.ReportObjects
Dim crxSubreportObj As CRAXDRT.SubreportObject

Dim crxSubFrente As CRAXDRT.Report
Dim crxSub1Reves As CRAXDRT.Report
Dim crxSub2Reves As CRAXDRT.Report



Private Sub Form_Activate()

    If Me.Tag = "LOADED" Then Exit Sub
    
    Me.Tag = "LOADED"

    Screen.MousePointer = vbHourglass
    LlenaRecordSet
    MostrarFrente
    MostrarReves
    Screen.MousePointer = vbDefault
    
    Me.crvFrente.Zoom (55)
    Me.crvReves.Zoom (55)
    
    Me.WindowState = 2
    
    
End Sub

Private Sub LlenaRecordSet()
    'Cadena que utiliza el reporte del frente
    sFrente = "SELECT DISTINCT USUARIOS!NoFamilia, USUARIOS!idTipoUsuario, TIPOMEM!Descripcion, "
    sFrente = sFrente & "USUARIOS!Nombre, USUARIOS!A_Paterno, USUARIOS!A_MATERNO, USUARIOS!FechaNacio, "
    sFrente = sFrente & "USUARIOS!Profesion, "
    sFrente = sFrente & "(SELECT TRIM(Direcciones!Calle) & "
    sFrente = sFrente & "iif(NOT ISNULL(Direcciones!Colonia), ', ' & TRIM(Direcciones!Colonia), '' ) & "
    sFrente = sFrente & "iif(NOT ISNULL(Direcciones!CodPos), ', ' & TRIM(Direcciones!CodPos), '' ) & "
    sFrente = sFrente & "iif(NOT ISNULL(Direcciones!DeloMuni), ', ' & TRIM(Direcciones!DeloMuni), '' ) & "
    sFrente = sFrente & "iif(NOT ISNULL(Direcciones!Ciudad), ', ' & TRIM(Direcciones!Ciudad), '' ) & "
    sFrente = sFrente & "iif(NOT ISNULL(Direcciones!Estado), ', ' & TRIM(Direcciones!Estado), '' ) "
    sFrente = sFrente & "FROM Direcciones WHERE Direcciones.idMember=USUARIOS.idMember AND Direcciones.idTipoDireccion=1) AS DOMICILIO, "
    sFrente = sFrente & "MEMBRESIA!FechaAlta, MEMBRESIA!Duracion, MEMBRESIA!Monto, MEMBRESIA!Enganche, MEMBRESIA!NumeroPagos, "
    sFrente = sFrente & "USUARIOS!idMember, MEMBRESIA!idMembresia, MEMBRESIA!idTipoMembresia, MEMBRESIA!NombrePropietario "
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
    
    
    'Cadena que utiliza el subreporte del frente
    sSubFrente = "SELECT DETALLE!Monto, DETALLE!FechaVence "
    sSubFrente = sSubFrente & "FROM Detalle_Mem AS DETALLE LEFT JOIN Membresias ON DETALLE.idMembresia=Membresias.idMembresia "
    sSubFrente = sSubFrente & "WHERE DETALLE.NoPago>0 AND Membresias.idMember=" & nidTitular
    
'    sSubFrente = sSubFrente & "WHERE DETALLE.NoPago>0 AND Membresias.idMember=" & frmConsMembers.ssdbMembresia.Columns(5).Value
    
    'Crea la instancia del recordset
    Set rsSubFrente = New ADODB.Recordset
    
    'Asigna sus propiedades del recordset
    With rsSubFrente
        .Source = sSubFrente
        .ActiveConnection = Conn
        .CursorType = adOpenForwardOnly
        .CursorLocation = adUseClient
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With


    'Cadena que utiliza el reporte del revés
    sReves = "SELECT USUARIOS!FechaNacio, USUARIOS!Profesion, USUARIOS!idPais, "
    sReves = sReves & "Paises!Pais, USUARIOS!UFechaPago, "
    sReves = sReves & "(SELECT Direcciones!RazonSocial FROM Direcciones WHERE Direcciones.idTipoDireccion=2 AND Direcciones.idMember=USUARIOS.idMember) AS EMPRESA, "

    sReves = sReves & "(SELECT TRIM(Direcciones!Calle) & "
    sReves = sReves & "iif(NOT ISNULL(Direcciones!Colonia), ', ' & TRIM(Direcciones!Colonia), '' ) & "
    sReves = sReves & "iif(NOT ISNULL(Direcciones!CodPos), ', ' & TRIM(Direcciones!CodPos), '' ) & "
    sReves = sReves & "iif(NOT ISNULL(Direcciones!DeloMuni), ', ' & TRIM(Direcciones!DeloMuni), '' ) & "
    sReves = sReves & "iif(NOT ISNULL(Direcciones!Ciudad), ', ' & TRIM(Direcciones!Ciudad), '' ) & "
    sReves = sReves & "iif(NOT ISNULL(Direcciones!Estado), ', ' & TRIM(Direcciones!Estado), '' ) "
    sReves = sReves & "FROM Direcciones WHERE Direcciones.idTipoDireccion=2 AND Direcciones.idMember=USUARIOS.idMember) AS DOMICILIO, "

    sReves = sReves & "(SELECT iif(NOT ISNULL(Direcciones!Tel1), TRIM(Direcciones!Tel1), '' ) & "
    sReves = sReves & "iif(NOT ISNULL(Direcciones!Tel2), ', ' & TRIM(Direcciones!Tel2), '' ) "
    sReves = sReves & "FROM Direcciones WHERE Direcciones.idTipoDireccion=2 AND Direcciones.idMember=USUARIOS.idMember) AS TELS, "

    sReves = sReves & "(SELECT iif(NOT ISNULL(Direcciones!Fax), TRIM(Direcciones!Fax), '' ) "
    sReves = sReves & "FROM Direcciones WHERE Direcciones.idTipoDireccion=2 AND Direcciones.idMember=USUARIOS.idMember) AS FAX, "
    
    'gpo 13/10/2005
    sReves = sReves & "(SELECT Membresias!Observaciones "
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
    sSub1Reves = "SELECT USUARIOS!idTipoUsuario, TIPOUSER!Descripcion, USUARIOS!Nombre, USUARIOS!A_Paterno, USUARIOS!A_Materno, "
    sSub1Reves = sSub1Reves & "USUARIOS!FechaNacio "
    sSub1Reves = sSub1Reves & "FROM Usuarios_Club AS USUARIOS LEFT JOIN Tipo_Usuario AS TIPOUSER ON USUARIOS.idTipoUsuario=TIPOUSER.idTipoUsuario "
    sSub1Reves = sSub1Reves & "WHERE USUARIOS.idMember<>USUARIOS.idTitular AND USUARIOS.idTitular=" & nidTitular
    
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
End Sub


Private Sub MostrarFrente()
Dim sCantLetra As String

    'Esta linea llama al archivo .rpt que se creo con Crystal Report
    Set crxFrente = crxApplication.OpenReport(sDB_ReportSource & "\RepMembresia.rpt")

    Set crxDatabase = crxFrente.Database
    Set crxDatabaseTables = crxDatabase.Tables
    Set crxDatabaseTable = crxDatabaseTables.Item(1)

    crxDatabaseTable.SetDataSource rsFrente, 3

    'Evita que Crystal intente pedir el valor de los parametros
    crxFrente.EnableParameterPrompting = False
    
    sCantLetra = Trim$(UCase$(Num2Txt(nTotalMem)))
    sCantLetra = sCantLetra & " PESOS " & Format((nTotalMem - Int(nTotalMem)) * 100, "#00") & "/100 M. N."
    
'    sCantLetra = Trim$(UCase$(Num2Txt(frmConsMembers.ssdbMembresia.Columns(7).Value)))
'    sCantLetra = sCantLetra & " PESOS " & Format((frmConsMembers.ssdbMembresia.Columns(7).Value - Int(frmConsMembers.ssdbMembresia.Columns(7).Value)) * 100, "#00") & "/100 M. N."
    
    crxFrente.ParameterFields(1).AddCurrentValue sCantLetra

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

                crxDatabaseTable.SetDataSource rsSubFrente, 3
            End If
        Next
    Next


    Me.crvFrente.ReportSource = crxFrente
    Me.crvFrente.ViewReport
End Sub


Private Sub MostrarReves()
Dim nSubRep As Byte

    'Esta linea llama al archivo .rpt que se creo con Crystal Report
    Set crxReves = crxApplication.OpenReport(sDB_ReportSource & "\DatosMembresia.rpt")
    
    Set crxDatabase = crxReves.Database
    Set crxDatabaseTables = crxDatabase.Tables
    Set crxDatabaseTable = crxDatabaseTables.Item(1)

    crxDatabaseTable.SetDataSource rsReves, 3

    'Evita que Crystal intente pedir el valor de los parametros
    crxReves.EnableParameterPrompting = False
    
    'Recorre las secciones del reporte e identifica aquellas
    'que sean subreportes
    Set crxSections = crxReves.Sections
    
    nSubRep = 1

    For X = 1 To crxSections.Count
        Set crxSection = crxSections.Item(X)

        Set crxReportObjs = crxSection.ReportObjects
        For Y = 1 To crxReportObjs.Count
            If (crxReportObjs.Item(Y).Kind = crSubreportObject) Then
                Set crxSubreportObj = crxReportObjs.Item(Y)
                
                If (nSubRep = 1) Then
                    Set crxSub1Reves = crxSubreportObj.OpenSubreport
                    Set crxDatabase = crxSub1Reves.Database
                Else
                    Set crxSub2Reves = crxSubreportObj.OpenSubreport
                    Set crxDatabase = crxSub2Reves.Database
                End If

                Set crxDatabaseTables = crxDatabase.Tables
                Set crxDatabaseTable = crxDatabaseTables.Item(1)

                If (nSubRep = 1) Then
                    crxDatabaseTable.SetDataSource rsSub1Reves, 3
                Else
                    crxDatabaseTable.SetDataSource rsSub2Reves, 3
                End If
                
                nSubRep = 2
            End If
        Next
    Next


    Me.crvReves.ReportSource = crxReves
    Me.crvReves.ViewReport
End Sub


Private Sub Form_Resize()
    With Me.crvFrente
        .Top = 20
        .Left = 10
        If (ScaleWidth / 2) > 50 Then
            .Width = (ScaleWidth / 2) - 50
        End If
        If ScaleHeight > 50 Then
            .Height = ScaleHeight - 50
        End If
        .Zoom (55)
    End With

    With Me.crvReves
        .Top = 20
        .Left = (ScaleWidth / 2) + 100
        If (ScaleWidth / 2) > 50 Then
            .Width = (ScaleWidth / 2) - 50
        End If
        
        If ScaleHeight > 50 Then
            .Height = ScaleHeight - 50
        End If
        .Zoom (55)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmRepMembresia = Nothing
End Sub
