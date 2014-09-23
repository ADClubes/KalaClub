Attribute VB_Name = "ProcsSocios"
'************************************************
'*  Procedimientos para el modulo de los socios *
'*  Daniel Hdez                                 *
'*  27 / Septiembre / 2004                      *
'************************************************

'Ult actualizacion: 29 / Junio / 2005


'Crea un control Adodc
Public Sub InitCtrlAdo(adoCtrl As Adodc, sTablas As String, sCampos As String, sCondicion As String, cnConnect As ADODB.Connection)
Dim sRecSource As String

    sRecSource = "SELECT " & sCampos & " FROM " & sTablas
    
    If (sCondicion <> "") Then
        sRecSource = sRecSource & " WHERE " & sCondicion
    End If
    
    'Propiedades de Ctrl Adodc
    With adoCtrl
        .ConnectionString = cnConnect
        .RecordSource = sRecSource
        .REFRESH
    End With
End Sub


'Crea un control Adodc con la lista de campos que recibe
Public Sub InitCtrlAdoSel(adoCtrl As Adodc, sTabla As String, mCampos() As String, nTotCols As Byte, nColumnaActiva As Integer, sCondicion As String, cnConnect As ADODB.Connection)
Dim sListaCampos As String
Dim nPointer As Integer
Dim sRecSource As String

    sListaCampos = ""

    'Crea la lista de los campos que se deben mostrar
    For nPointer = 0 To (nTotCols - 2)
        sListaCampos = sListaCampos & mCampos(nPointer) & ", "
    Next nPointer
    sListaCampos = sListaCampos & mCampos(nTotCols - 1)
    
    sRecSource = "Select " & sListaCampos & " From " & sTabla
    
    If (sCondicion <> "") Then
        sRecSource = sRecSource & " Where " & sCondicion
    End If
    
    If InStr(mCampos(nColumnaActiva), "AS") Then
        sRecSource = sRecSource & " Order by " & Left(mCampos(nColumnaActiva), InStr(mCampos(nColumnaActiva), "AS") - 1)
    Else
        sRecSource = sRecSource & " Order by " & mCampos(nColumnaActiva)
    End If
    
    'Propiedades de adoCtrl
    With adoCtrl
        .ConnectionString = cnConnect
        .RecordSource = sRecSource
        .REFRESH
    End With
End Sub


'Define el ancho de cada una de las columnas en un DataGrid
Public Sub DefAnchoDBGrid(dgDataGrid As DataGrid, mValores() As Integer)
Dim nPointer As Integer

    For nPointer = 0 To (dgDataGrid.Columns.Count - 1)
        'Asigna el ancho
        dgDataGrid.Columns(nPointer).Width = mValores(nPointer)
        
        'Evita que se modifique el valor asignado
        dgDataGrid.Columns(nPointer).AllowSizing = False
    Next nPointer
End Sub


'Escribe los encabezados de cada una de las columnas en un DataGrid
Public Sub DefHeadersDBGrid(dgDataGrid As DataGrid, mValores() As String)
Dim nPointer As Integer

    'Escribe los encabezados en negritas
    dgDataGrid.HeadFont.Bold = True
    
    For nPointer = 0 To (dgDataGrid.Columns.Count - 1)
          dgDataGrid.Columns(nPointer).Caption = mValores(nPointer)
    Next nPointer
End Sub


'Define el ancho de cada una de las columnas en un DataGrid
Public Sub DefAnchossGrid(ssGrid As SSDBGrid, mValores() As Integer)
Dim nPointer As Integer

    For nPointer = 0 To (ssGrid.Columns.Count - 1)
        'Asigna el ancho
        ssGrid.Columns(nPointer).Width = mValores(nPointer)
        
        'Evita que se modifique el valor asignado
        ssGrid.Columns(nPointer).AllowSizing = False
    Next nPointer
End Sub


'Escribe los encabezados de cada una de las columnas en un DataGrid
Public Sub DefHeaderssGrid(ssGrid As SSDBGrid, mValores() As String)
Dim nPointer As Integer

    'Escribe los encabezados en negritas
    ssGrid.HeadFont.Bold = True
    
    For nPointer = 0 To (ssGrid.Columns.Count - 1)
        ssGrid.Columns(nPointer).Caption = mValores(nPointer)
    Next nPointer
End Sub


'Agrega un registro en el Log de los usuarios del sistema
Public Sub EscribeHistorial(sUsuario As String, dFecha As Date, sSuceso As String, nNino As Integer)
'Dim mFields(4) As String
'Dim mValues(4) As Variant
'
'    mFields(0) = "CveNino"
'    mFields(1) = "Fecha"
'    mFields(2) = "Descripcion"
'    mFields(3) = "Responsable"
'
'    mValues(0) = nNino
'    mValues(1) = dFecha
'    mValues(2) = Left(sSuceso, 100)
'    mValues(3) = Left(sUsuario, 50)
'
'    If (Not AddRecord("Historial", mFields, 4, mValues, cnConexion, "Guardando historial")) Then
'        MsgBox "No se guardó registro en el historial", vbInformation, NOMBREEMPRESA
'    End If
End Sub


'Crea un RecordSet con los datos que llegan como parametros
Public Sub InitRecordSet(rsRecordSet As ADODB.Recordset, sCadenaCampos As String, sCadenaTablas As String, Optional sCadenaCondicion As String, Optional sCadenaOrden As String, Optional cnConecta As ADODB.Connection)
Dim sSql As String


    'Instancia el RecordSet
    Set rsRecordSet = New ADODB.Recordset
    
    sSql = "SELECT " & sCadenaCampos & " FROM " & sCadenaTablas
    
    If (sCadenaCondicion <> "") Then
        sSql = sSql & " WHERE " & sCadenaCondicion
    End If
    
    If (sCadenaOrden <> "") Then
        sSql = sSql & " ORDER BY " & sCadenaOrden
    End If
    
    'Asigna sus propiedades
    With rsRecordSet
        .Source = sSql
        .ActiveConnection = cnConecta
        .CursorType = adOpenStatic
        .CursorLocation = adUseServer
        .LockType = adLockReadOnly
        .Open Options:=adCmdText
    End With
End Sub


'Busqueda secuencial en un DataGrid
Public Sub BuscarEnDG(adoDataCtrl As Adodc, bDesdeInicio As Boolean, sCriterio As String)
Dim nRen As Variant

    nRen = adoDataCtrl.Recordset.Bookmark
    
    'Si se solicita la busqueda desde el primer registro
    If (bDesdeInicio) Then
        adoDataCtrl.Recordset.MoveFirst
    Else
        'Se pasa al siguiente registro y a partir de ahi inicia la busqueda
        If (Not adoDataCtrl.Recordset.EOF) Then
            adoDataCtrl.Recordset.MoveNext
        End If
        
        'Si llega al ultimo registro, se regresa al primero
        If (adoDataCtrl.Recordset.EOF) Then
            adoDataCtrl.Recordset.MoveFirst
        End If
    End If
    
    adoDataCtrl.Recordset.Find (sCriterio)
    
    If (adoDataCtrl.Recordset.EOF) Then
        adoDataCtrl.Recordset.Bookmark = nRen
'        MsgBox "Registro no localizado.", vbExclamation, "Buscar datos"
        Exit Sub
    Else
        nRen = adoDataCtrl.Recordset.Bookmark
    End If
    
    adoDataCtrl.Recordset.Bookmark = nRen
End Sub


'Habilita o deshabilita las pestañas de la forma
Public Sub HabilitaTabs(bValor As Boolean)
Dim i As Byte

    'Pestañas del ctrl ssTab
    For i = 1 To (frmAltaSocios.sstabSocios.Tabs - 1)
        frmAltaSocios.sstabSocios.TabEnabled(i) = bValor
    Next
    
'    'En caso de ser propietario o rentista deshabilita la pestaña de membresias
'    If (Not frmAltaSocios.optMembresia.Value) Then
'        frmAltaSocios.sstabSocios.TabEnabled(frmAltaSocios.sstabSocios.Tabs - 2) = False
'    End If
End Sub


'Habilita o deshabilita los OptionButton de la forma frmAltaSocios
Public Sub HabilitaOpciones(bValor As Boolean)
    'frmAltaSocios.optRentista.Enabled = bValor
    'frmAltaSocios.optMembresia.Enabled = bValor
    'frmAltaSocios.optPropietario.Enabled = bValor
End Sub


'Llena los datos de la pestaña Generales en la forma frmAltaSocios
Public Sub LlenaGrales()
Dim rsTitular As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String
Dim sCond As String

    If (frmAltaSocios.nTitCve > 0) Then
        #If SqlServer_ Then
            sCampos = "Usuarios_Club.A_Paterno AS [Usuarios_Club.A_Paterno], Usuarios_Club.A_Materno AS [Usuarios_Club.A_Materno], Usuarios_Club.Nombre AS [Usuarios_Club.Nombre], "
            sCampos = sCampos & "Usuarios_Club.IdMember, Usuarios_Club.Profesion, "
            sCampos = sCampos & "Usuarios_Club.FechaNacio, Usuarios_Club.FechaIngreso, "
            sCampos = sCampos & "Usuarios_Club.Email, Usuarios_Club.Celular, "
            sCampos = sCampos & "Usuarios_Club.Sexo, Usuarios_Club.FotoFile, Usuarios_Club.NoFamilia, "
            sCampos = sCampos & "Usuarios_Club.IdPais, Usuarios_Club.IdTipoUsuario, "
            sCampos = sCampos & "Tipo_Pago.Descripcion AS Tipo_Pago_Descripcion, Tipo_Uso_Accion.Descripcion AS Tipo_Uso_Accion_Descripcion, "
            sCampos = sCampos & "Titulos.Serie, Titulos.Tipo, Titulos.Numero, "
            sCampos = sCampos & "Accionistas.A_Paterno AS [Accionistas.A_Paterno], Accionistas.A_Materno AS [Accionistas.A_Materno], "
            sCampos = sCampos & "Accionistas.Nombre AS [Accionistas.Nombre], Accionistas.IdPropTitulo, "
            sCampos = sCampos & "Accionistas.Telefono_1, Accionistas.Telefono_2, "
            sCampos = sCampos & "Paises.Pais, Secuencial.Secuencial, "
            sCampos = sCampos & "Tipo_Usuario.Descripcion AS Tipo_Usuario_Descripcion, "
            sCampos = sCampos & "Usuarios_Club.UFechaPago, "
            '11/10/2005
            sCampos = sCampos & "Usuarios_Club.Inscripcion "
            '18/06/2011
            'HABILITAR PARA DIRECCIONADOS
'            sCampos = sCampos & "Usuarios_Club.ISOPais, "
'            sCampos = sCampos & "Usuarios_Club.ISOEstado, "
'            sCampos = sCampos & "Usuarios_Club.CURP "
            
            sTablas = "(((((((Usuarios_Club LEFT JOIN Usuarios_Titulo ON Usuarios_Club.IdMember=Usuarios_Titulo.IdMember) "
            sTablas = sTablas & "LEFT JOIN Tipo_Pago ON Usuarios_Titulo.IdTipoPago=Tipo_Pago.IdTipoPago) "
            sTablas = sTablas & "LEFT JOIN Tipo_Uso_Accion ON Usuarios_Titulo.IdTipoUsoAccion=Tipo_Uso_Accion.IdTipoUsoAccion) "
            sTablas = sTablas & "LEFT JOIN Titulos ON Usuarios_Titulo.Serie=Titulos.Serie AND Usuarios_Titulo.Tipo=Titulos.Tipo AND Usuarios_Titulo.Numero=Titulos.Numero) "
            sTablas = sTablas & "LEFT JOIN Accionistas ON Titulos.IdPropietario=Accionistas.IdPropTitulo) "
            sTablas = sTablas & "LEFT JOIN Paises ON Usuarios_Club.IdPais=Paises.IdPais) "
            sTablas = sTablas & "LEFT JOIN Secuencial ON Usuarios_Club.IdMember=Secuencial.IdMember) "
            sTablas = sTablas & "LEFT JOIN Tipo_Usuario ON Usuarios_Club.IdTipoUsuario=Tipo_Usuario.IdTipoUsuario "
        #Else
    
            sCampos = "Usuarios_Club.A_Paterno, Usuarios_Club.A_Materno, Usuarios_Club.Nombre, "
            sCampos = sCampos & "Usuarios_Club.IdMember, Usuarios_Club.Profesion, "
            sCampos = sCampos & "Usuarios_Club.FechaNacio, Usuarios_Club.FechaIngreso, "
            sCampos = sCampos & "Usuarios_Club.Email, Usuarios_Club.Celular, "
            sCampos = sCampos & "Usuarios_Club.Sexo, Usuarios_Club.FotoFile, Usuarios_Club.NoFamilia, "
            sCampos = sCampos & "Usuarios_Club.IdPais, Usuarios_Club.IdTipoUsuario, "
            sCampos = sCampos & "Tipo_Pago.Descripcion AS Tipo_Pago_Descripcion, Tipo_Uso_Accion.Descripcion AS Tipo_Uso_Accion_Descripcion, "
            sCampos = sCampos & "Titulos.Serie, Titulos.Tipo, Titulos.Numero, "
            sCampos = sCampos & "Accionistas.A_Paterno, Accionistas.A_Materno, "
            sCampos = sCampos & "Accionistas.Nombre, Accionistas.IdPropTitulo, "
            sCampos = sCampos & "Accionistas.Telefono_1, Accionistas.Telefono_2, "
            sCampos = sCampos & "Paises.Pais, Secuencial.Secuencial, "
            sCampos = sCampos & "Tipo_Usuario.Descripcion AS Tipo_Usuario_Descripcion, "
            sCampos = sCampos & "Usuarios_Club.UFechaPago, "
            '11/10/2005
            sCampos = sCampos & "Usuarios_Club.Inscripcion "
            
            sTablas = "(((((((Usuarios_Club LEFT JOIN Usuarios_Titulo ON Usuarios_Club.IdMember=Usuarios_Titulo.IdMember) "
            sTablas = sTablas & "LEFT JOIN Tipo_Pago ON Usuarios_Titulo.IdTipoPago=Tipo_Pago.IdTipoPago) "
            sTablas = sTablas & "LEFT JOIN Tipo_Uso_Accion ON Usuarios_Titulo.IdTipoUsoAccion=Tipo_Uso_Accion.IdTipoUsoAccion) "
            sTablas = sTablas & "LEFT JOIN Titulos ON Usuarios_Titulo.Serie=Titulos.Serie AND Usuarios_Titulo.Tipo=Titulos.Tipo AND Usuarios_Titulo.Numero=Titulos.Numero) "
            sTablas = sTablas & "LEFT JOIN Accionistas ON Titulos.IdPropietario=Accionistas.IdPropTitulo) "
            sTablas = sTablas & "LEFT JOIN Paises ON Usuarios_Club.IdPais=Paises.IdPais) "
            sTablas = sTablas & "LEFT JOIN Secuencial ON Usuarios_Club.IdMember=Secuencial.IdMember) "
            sTablas = sTablas & "LEFT JOIN Tipo_Usuario ON Usuarios_Club.IdTipoUsuario=Tipo_Usuario.IdTipoUsuario "
        #End If
        sCond = " Usuarios_Club.IdMember = " & frmAltaSocios.nTitCve
    
        InitRecordSet rsTitular, sCampos, sTablas, sCond, "", Conn
        
        With rsTitular
            If (.RecordCount > 0) Then
                Select Case .Fields("Tipo_Uso_Accion_Descripcion")
                    Case "PROPIETARIO"
                        frmAltaSocios.bProp = True
                        frmAltaSocios.bRent = False
                        frmAltaSocios.bMemb = False
                        frmAltaSocios.optPropietario.Value = True
                        
                    Case "RENTISTA"
                        frmAltaSocios.bRent = True
                        frmAltaSocios.bProp = False
                        frmAltaSocios.bMemb = False
                        frmAltaSocios.optRentista.Value = True
                        
                    Case "MEMBRESIA"
                        frmAltaSocios.bMemb = True
                        frmAltaSocios.bProp = False
                        frmAltaSocios.bRent = False
                        frmAltaSocios.optMembresia.Value = True
                End Select
                
                If (Trim(.Fields("Serie")) <> "") Then
                    frmAltaSocios.txtSerie.Text = Trim(.Fields("Titulos!Serie"))
                End If
                
                If (Trim(.Fields("Tipo")) <> "") Then
                    frmAltaSocios.txtTipo.Text = Trim(.Fields("Titulos!Tipo"))
                End If
                
                If (Trim(.Fields("Numero")) > 0) Then
                    frmAltaSocios.txtNumero.Text = .Fields("Titulos!Numero")
                End If
                
                If (Trim(.Fields("Accionistas.A_Paterno")) <> "") Then
                    frmAltaSocios.txtNombre.Text = Trim(.Fields("Accionistas!A_Paterno")) & " "
                End If
                
                If (Trim(.Fields("Accionistas.A_Materno")) <> "") Then
                    frmAltaSocios.txtNombre.Text = frmAltaSocios.txtNombre.Text & Trim(.Fields("Accionistas!A_Materno")) & " "
                End If
                
                If (Trim(.Fields("Accionistas.Nombre")) <> "") Then
                    frmAltaSocios.txtNombre.Text = frmAltaSocios.txtNombre.Text & Trim(.Fields("Accionistas!Nombre"))
                End If
                
                If (Trim(.Fields("Telefono_1")) <> "") Then
                    frmAltaSocios.txtTel1.Text = .Fields("Accionistas!Telefono_1")
                End If
                
                If (Trim(.Fields("Telefono_2")) <> "") Then
                    frmAltaSocios.txtTel2.Text = .Fields("Accionistas!Telefono_2")
                End If
                
                If (Trim(.Fields("IdPropTitulo")) > 0) Then
                    frmAltaSocios.txtCveAccionista.Text = .Fields("Accionistas!IdPropTitulo")
                    frmAltaSocios.nAccCve = .Fields("Accionistas!IdPropTitulo")
                End If
                
                If (Trim(.Fields("Usuarios_Club.A_Paterno")) <> "") Then
                    frmAltaSocios.txtTitPaterno = .Fields("Usuarios_Club.A_Paterno")
                End If
                
                If (Trim(.Fields("Usuarios_Club.A_Materno")) <> "") Then
                    frmAltaSocios.txtTitMaterno = .Fields("Usuarios_Club.A_Materno")
                End If
                
                If (Trim(.Fields("Usuarios_Club.Nombre")) <> "") Then
                    frmAltaSocios.txtTitNombre.Text = .Fields("Usuarios_Club.Nombre")
                End If
                
                If (Trim(.Fields("IdMember")) > 0) Then
                    frmAltaSocios.txtTitCve.Text = .Fields("IdMember")
                End If
                
                If (Trim(.Fields("NoFamilia")) > 0) Then
                    frmAltaSocios.txtFamilia.Text = .Fields("NoFamilia")
                End If
                
                If (Trim(.Fields("Profesion")) <> "") Then
                    frmAltaSocios.txtTitProf.Text = .Fields("Profesion")
                End If
                
                frmAltaSocios.dtpTitNacio.Value = .Fields("FechaNacio")
                frmAltaSocios.txtTitEdad.Text = Format(Edad(.Fields("FechaNacio")), "#0.00")
                frmAltaSocios.dtpTitRegistro.Value = .Fields("FechaIngreso")
                
                If (Trim(.Fields("Celular")) <> "") Then
                    frmAltaSocios.txtTitCel.Text = .Fields("Celular")
                End If
                
                If (Trim(.Fields("Email")) <> "") Then
                    frmAltaSocios.txtTitEmail.Text = .Fields("Email")
                End If
        
                If (.Fields("Sexo") = "F") Then
                    frmAltaSocios.optFemenino.Value = True
                Else
                    frmAltaSocios.optMasculino.Value = True
                End If
        
                If (Trim(.Fields("Pais")) <> "") Then
                    frmAltaSocios.txtPaisTit.Text = .Fields("Pais")
                    frmAltaSocios.txtCvePais.Text = .Fields("IdPais")
                End If
                
                If (Trim(.Fields("Tipo_Usuario_Descripcion")) <> "") Then
                    frmAltaSocios.txtTipoTit.Text = .Fields("Tipo_Usuario_Descripcion")
                    frmAltaSocios.txtCveTipo.Text = .Fields("IdTipoUsuario")
                End If
                
'                If (Trim(.Fields("Tipo_Pago!Descripcion")) <> "") Then
'                    frmAltaSocios.cbComoPaga.Text = .Fields("Tipo_Pago!Descripcion")
'                End If
                
                If (.Fields("Secuencial") > 0) Then
                    frmAltaSocios.txtSecuencial.Text = .Fields("Secuencial")
                End If
                
                If (Not IsNull(.Fields("FotoFile"))) Then
                    frmAltaSocios.txtImagen.Text = .Fields("FotoFile")
                End If
                
                If (Dir(sG_RutaFoto & "\" & Trim(.Fields("FotoFile")) & ".jpg") <> "") Then
                    frmAltaSocios.imgFoto.Picture = LoadPicture(sG_RutaFoto & "\" & Trim(.Fields("FotoFile")) & ".jpg")
                Else
                    frmAltaSocios.imgFoto.Picture = LoadPicture("")
                End If
                
                frmAltaSocios.dtpFechaUPago.Value = .Fields("UFechaPago")
                
                '11/10/2005 gpo
                If (Trim(.Fields("Inscripcion"))) <> "" Then
                    frmAltaSocios.txtNoIns.Text = Trim(.Fields("Inscripcion"))
                End If
                
                '18/06/2011
'                If (Trim(.Fields("ISOPais"))) <> "" Then
'                    frmAltaSocios.cboPaises.Text = Trim(.Fields("ISOPais"))
'                End If
'
'                If (Trim(.Fields("ISOEstado"))) <> "" Then
'                    frmAltaSocios.cboEstados.Text = Trim(.Fields("ISOEstado"))
'                End If
'
'                If (Trim(.Fields("CURP"))) <> "" Then
'                    frmAltaSocios.txtCURP.Text = Trim(.Fields("CURP"))
'                End If
                
            End If
        End With
        
        rsTitular.Close
        
        Set rsTitular = Nothing
    
        frmAltaSocios.txtTitCve.Enabled = False
        frmAltaSocios.dtpTitRegistro.Enabled = False
    End If
End Sub


Public Sub ActivaCred(nPanel As Byte, nCred As Long, nTZone As Integer, nSocio As Long, bActiva As Boolean, bMensaje As Boolean)
    Const DATOSCRED = 6
    Dim mCampos(DATOSCRED) As String
    Dim mValor(DATOSCRED) As Variant
    Dim sComando As String
    
    Dim sSiteCode As String
    Dim sSiteCodeOld As String
    Dim sSiteCodeNew As String
    
    Dim lLimiteSecSup As Long
    Dim lLimiteSecInf As Long
    
    '_C=pn_cn_tz_d1_d2
    
    '_  -> espacio en blanco
    'C  -> letra c siempre en mayuscula
    'pn -> Numero de panel
    'cn -> Numero de credencial (en nuestro caso es el secuencial)
    'tz -> Time zone (se lee de la tabla time_Zone)
    'd1 -> Dispositivo de lectura 1 (en nuestro caso siempre es 1)
    'd2 -> Dispositivo de lectura 2 (en nuestro caso siempre es 2)

    
    If (ObtieneParametro("ACTIVA RENOVAR CREDENCIAL") = "1") Then
        sSiteCodeOld = ObtieneParametro("SiteCode")
        sSiteCodeNew = ObtieneParametro("SiteCodeNew")
        lLimiteSecSup = CLng(ObtieneParametro("LIMITESECUENCIALSUP"))
        lLimiteSecInf = CLng(ObtieneParametro("LIMITESECUENCIALINF"))
        
        If nCred > lLimiteSecInf And nCred < lLimiteSecSup Then
            sSiteCode = sSiteCodeNew
        Else
            sSiteCode = sSiteCodeOld
        End If
    Else
        sSiteCode = ObtieneParametro("SiteCode")
    End If
    
    
    If ObtieneParametro("PART_TIME") = "1" Then
        If GetTipoUsuario(nSocio) = 59 Then
            nTZone = 2
        End If
    End If


    'Arma el comando
    sComando = " C=" & nPanel & " " & sSiteCode & Format(nCred, "00000")
    
    'Activar credenciales
    If (bActiva) Then
         sComando = sComando & " " & nTZone & " 1 2 3 4"
    End If

    mCampos(0) = "Id"
    mCampos(1) = "Comando"
    mCampos(2) = "Enviado"
    mCampos(3) = "IdMember"
    mCampos(4) = "Fecha"
    mCampos(5) = "Activar"
    
    #If SqlServer_ Then
        mValor(0) = LeeUltReg("Poll", "Id") + 1
        mValor(1) = sComando
        mValor(2) = 0                              '0 = False, 1 = True
        mValor(3) = nSocio
        mValor(4) = Format(Date, "yyyymmdd")
        mValor(5) = IIf(bActiva, 1, 0)
    #Else
        mValor(0) = LeeUltReg("Poll", "Id") + 1
        mValor(1) = sComando
        mValor(2) = 0                              '0 = False, 1 = True
        mValor(3) = nSocio
        mValor(4) = Format(Date, "dd/mm/yyyy")
        mValor(5) = IIf(bActiva, 1, 0)
    #End If
    
    If (Not AgregaRegistro("Poll", mCampos, DATOSCRED, mValor, Conn)) Then
        MsgBox "No se enviaron los datos de la credencial: " & nCred & " del socio: " & nSocio & " correctamente, intente de nuevo.", vbExclamation, "KalaSystems"
    Else
        If (bActiva) Then
            sComando = "La credencial se activó correctamente."
        Else
            sComando = "La credencial se desactivó correctamente."
        End If
        
        If (bMensaje) Then
            MsgBox sComando, vbInformation, "KalaSystems"
        End If
    End If
End Sub

''
'' Realiza la activación de credenciales en SQL Server.
''
Public Sub ActivaCredSQLMulti(nPanel As Byte, nCred As Long, nTZone As Integer, nSocio As Long, bActiva As Boolean, bMensaje As Boolean)
    Const DATOSCRED = 5
    Dim mCampos(DATOSCRED) As String
    Dim mValor(DATOSCRED) As Variant
    Dim sComando As String
    Dim sSiteCode As String
    Dim sSiteCodeOld As String
    Dim sSiteCodeNew As String
    Dim lLimiteSecSup As Long
    Dim lLimiteSecInf As Long
    Dim sTipoUsuario As String
    
    
    
    'If ObtieneParametro("PART_TIME") = "1" Then
    sTipoUsuario = GetTipoUsuarioMulti(nCred)
        If sTipoUsuario Like "*MCF*" Then
            nTZone = 3
        ElseIf sTipoUsuario Like "*FIN*" Then
        nTZone = 3
        ElseIf sTipoUsuario Like "*PART*" Then
        nTZone = 2
        End If
    'End If


    'Arma el comando
    sComando = " C=" & nPanel & " " & Format(nCred, "0000000000")
    
    'Activar credenciales
    If (bActiva) Then
         sComando = sComando & " " & nTZone & " 1 2 3 4"
    End If

    mCampos(0) = "Comando"
    mCampos(1) = "Enviado"
    mCampos(2) = "IdMember"
    mCampos(3) = "Fecha"
    mCampos(4) = "Activar"
    
    mValor(0) = sComando
    mValor(1) = 0                              '0 = False, 1 = True
    mValor(2) = nSocio
    #If SqlServer_ Then
        mValor(3) = Format(Date, "yyyymmdd")
    #Else
        mValor(3) = Format(Date, "dd/mm/yyyy")
    #End If
    mValor(4) = IIf(bActiva, 1, 0)
    
    If (Not AgregaRegistro("Poll", mCampos, DATOSCRED, mValor, Conn)) Then
        MsgBox "No se enviaron los datos de la credencial: " & nCred & " del socio: " & nSocio & " correctamente, intente de nuevo.", vbExclamation, "KalaSystems"
    Else
        If (bActiva) Then
            sComando = "La credencial se activó correctamente."
        Else
            sComando = "La credencial se desactivó correctamente."
        End If
        
        If (bMensaje) Then
            MsgBox sComando, vbInformation, "KalaSystems"
        End If
    End If
End Sub
Public Sub ActivaCredSQL(nPanel As Byte, nCred As Long, nTZone As Integer, nSocio As Long, bActiva As Boolean, bMensaje As Boolean)
    Const DATOSCRED = 5
    Dim mCampos(DATOSCRED) As String
    Dim mValor(DATOSCRED) As Variant
    Dim sComando As String
    Dim sSiteCode As String
    Dim sSiteCodeOld As String
    Dim sSiteCodeNew As String
    Dim lLimiteSecSup As Long
    Dim lLimiteSecInf As Long
    Dim sTipoUsuario As String
    
    '_C=pn_cn_tz_d1_d2
    
    '_  -> espacio en blanco
    'C  -> letra c siempre en mayuscula
    'pn -> Numero de panel
    'cn -> Numero de credencial (en nuestro caso es el secuencial)
    'tz -> Time zone (se lee de la tabla time_Zone)
    'd1 -> Dispositivo de lectura 1 (en nuestro caso siempre es 1)
    'd2 -> Dispositivo de lectura 2 (en nuestro caso siempre es 2)

    
    If (ObtieneParametro("ACTIVA RENOVAR CREDENCIAL") = "1") Then
        sSiteCodeOld = ObtieneParametro("SiteCode")
        sSiteCodeNew = ObtieneParametro("SiteCodeNew")
        lLimiteSecSup = CLng(ObtieneParametro("LIMITESECUENCIALSUP"))
        lLimiteSecInf = CLng(ObtieneParametro("LIMITESECUENCIALINF"))
        
        If nCred > lLimiteSecInf And nCred < lLimiteSecSup Then
            sSiteCode = sSiteCodeNew
        Else
            sSiteCode = sSiteCodeOld
        End If
    Else
        sSiteCode = ObtieneParametro("SiteCode")
    End If
    
    
    'If ObtieneParametro("PART_TIME") = "1" Then
    sTipoUsuario = GetTipoUsuario(nSocio)
        If sTipoUsuario Like "*FIN*" Then
        nTZone = 3
        ElseIf sTipoUsuario Like "*PART*" Then
        nTZone = 2
        End If
    'End If


    'Arma el comando
    sComando = " C=" & nPanel & " " & sSiteCode & Format(nCred, "00000")
    
    'Activar credenciales
    If (bActiva) Then
         sComando = sComando & " " & nTZone & " 1 2 3 4"
    End If

    mCampos(0) = "Comando"
    mCampos(1) = "Enviado"
    mCampos(2) = "IdMember"
    mCampos(3) = "Fecha"
    mCampos(4) = "Activar"
    
    mValor(0) = sComando
    mValor(1) = 0                              '0 = False, 1 = True
    mValor(2) = nSocio
    #If SqlServer_ Then
        mValor(3) = Format(Date, "yyyymmdd")
    #Else
        mValor(3) = Format(Date, "dd/mm/yyyy")
    #End If
    mValor(4) = IIf(bActiva, 1, 0)
    
    If (Not AgregaRegistro("Poll", mCampos, DATOSCRED, mValor, Conn)) Then
        MsgBox "No se enviaron los datos de la credencial: " & nCred & " del socio: " & nSocio & " correctamente, intente de nuevo.", vbExclamation, "KalaSystems"
    Else
        If (bActiva) Then
            sComando = "La credencial se activó correctamente."
        Else
            sComando = "La credencial se desactivó correctamente."
        End If
        
        If (bMensaje) Then
            MsgBox sComando, vbInformation, "KalaSystems"
        End If
    End If
End Sub

'--------------------------
Public Sub ActivaCred2(nPanel As Byte, sCred As String, nTZone As Integer, nSocio As Long, bActiva As Boolean, bMensaje As Boolean)
    Const DATOSCRED = 6
    Dim mCampos(DATOSCRED) As String
    Dim mValor(DATOSCRED) As Variant
    Dim sComando As String
    
    Dim sSiteCode As String

    '_C=pn_cn_tz_d1_d2
    
    '_  -> espacio en blanco
    'C  -> letra c siempre en mayuscula
    'pn -> Numero de panel
    'cn -> Numero de credencial (en nuestro caso es el secuencial)
    'tz -> Time zone (se lee de la tabla time_Zone)
    'd1 -> Dispositivo de lectura 1 (en nuestro caso siempre es 1)
    'd2 -> Dispositivo de lectura 2 (en nuestro caso siempre es 2)

    
    
    
    


    'Arma el comando
    sComando = " C=" & nPanel & " " & sCred
    
    'Activar credenciales
    If (bActiva) Then
         sComando = sComando & " " & nTZone & " 1 2 3 4"
    End If

    mCampos(0) = "Id"
    mCampos(1) = "Comando"
    mCampos(2) = "Enviado"
    mCampos(3) = "IdMember"
    mCampos(4) = "Fecha"
    mCampos(5) = "Activar"
    
    #If SqlServer_ Then
        mValor(0) = LeeUltReg("Poll", "Id") + 1
        mValor(1) = sComando
        mValor(2) = 0                              '0 = False, 1 = True
        mValor(3) = nSocio
        mValor(4) = Format(Date, "yyyymmdd")
        mValor(5) = IIf(bActiva, 1, 0)
    #Else
        mValor(0) = LeeUltReg("Poll", "Id") + 1
        mValor(1) = sComando
        mValor(2) = 0                              '0 = False, 1 = True
        mValor(3) = nSocio
        mValor(4) = Format(Date, "dd/mm/yyyy")
        mValor(5) = IIf(bActiva, 1, 0)
    #End If
        
    If (Not AgregaRegistro("Poll", mCampos, DATOSCRED, mValor, Conn)) Then
        MsgBox "No se enviaron los datos de la credencial: " & nCred & " del socio: " & nSocio & " correctamente, intente de nuevo.", vbExclamation, "KalaSystems"
    Else
        If (bActiva) Then
            sComando = "La credencial se activó correctamente."
        Else
            sComando = "La credencial se desactivó correctamente."
        End If
        
        If (bMensaje) Then
            MsgBox sComando, vbInformation, "KalaSystems"
        End If
    End If
End Sub

''
'' Realiza la activación de credenciales en SQL Server.
''
Public Sub ActivaCred2SQL(nPanel As Byte, sCred As String, nTZone As Integer, nSocio As Long, bActiva As Boolean, bMensaje As Boolean)
    Const DATOSCRED = 5
    Dim mCampos(DATOSCRED) As String
    Dim mValor(DATOSCRED) As Variant
    Dim sComando As String
    Dim sSiteCode As String

    '_C=pn_cn_tz_d1_d2
    
    '_  -> espacio en blanco
    'C  -> letra c siempre en mayuscula
    'pn -> Numero de panel
    'cn -> Numero de credencial (en nuestro caso es el secuencial)
    'tz -> Time zone (se lee de la tabla time_Zone)
    'd1 -> Dispositivo de lectura 1 (en nuestro caso siempre es 1)
    'd2 -> Dispositivo de lectura 2 (en nuestro caso siempre es 2)
    
    'Arma el comando
    sComando = " C=" & nPanel & " " & sCred
    
    'Activar credenciales
    If (bActiva) Then
         sComando = sComando & " " & nTZone & " 1 2 3 4"
    End If
    
    mCampos(0) = "Comando"
    mCampos(1) = "Enviado"
    mCampos(2) = "IdMember"
    mCampos(3) = "Fecha"
    mCampos(4) = "Activar"
    
    #If SqlServer_ Then
        mValor(0) = sComando
        mValor(1) = 0                              '0 = False, 1 = True
        mValor(2) = nSocio
        mValor(3) = Format(Date, "yyyymmdd")
        mValor(4) = IIf(bActiva, 1, 0)
    #Else
        mValor(0) = sComando
        mValor(1) = 0                              '0 = False, 1 = True
        mValor(2) = nSocio
        mValor(3) = Format(Date, "dd/mm/yyyy")
        mValor(4) = IIf(bActiva, 1, 0)
    #End If
    
    If (Not AgregaRegistro("Poll", mCampos, DATOSCRED, mValor, Conn)) Then
        MsgBox "No se enviaron los datos de la credencial: " & nCred & " del socio: " & nSocio & " correctamente, intente de nuevo.", vbExclamation, "KalaSystems"
    Else
        If (bActiva) Then
            sComando = "La credencial se activó correctamente."
        Else
            sComando = "La credencial se desactivó correctamente."
        End If
        
        If (bMensaje) Then
            MsgBox sComando, vbInformation, "KalaSystems"
        End If
    End If
End Sub

'---------------------------
Public Function IncFolio(sTable As String, sField As String, nInc As Integer) As Boolean
    Const NDATOS = 1
    Dim mCampos(NDATOS) As String
    Dim mValores(NDATOS) As Variant

    'Campo que se va a actualizar
    mCampos(0) = sField

    'Valor que se debe guardar
    mValores(0) = LeeUltReg(sTable, sField) + 1
    
    If (CambiaReg(sTable, mCampos, NDATOS, mValores, sField & "=" & (mValores(0) - 1), Conn)) Then
        IncFolio = True
    Else
        IncFolio = False
        MsgBox "Error al incrementar la clave.", vbCritical, DEVELOPER
    End If
    
End Function


'-------------------------------------------------------------------------------
'ChecaTipoUsuario
'Germán Pérez Ortíz
'1 agosto 2005
'Parametros
'   lIdTitular  long    Numero del titular cuya familia será checada.

Public Sub ChecaTipoUsuario(lIdTitular As Long)
    Dim adorcsChecaTipos As ADODB.Recordset
    
    Dim dFechaComp As Date
    
    Set adorcsChecaTipos = New ADODB.Recordset
    adorcsChecaTipos.CursorLocation = adUseServer
    
    
    
    Do While True
    
    
        #If SqlServer_ Then
            strSQL = "SELECT USU.IdMember, USU.IdTipoUsuario, USU.FechaNacio, DateDiff(""yyyy"",USU.FechaNacio, GetDate()) AS Edad ,TIPO.EdadMaxima, REG.Accion, REG.IdTipoNuevo"
            strSQL = strSQL & " FROM (USUARIOS_CLUB USU INNER JOIN TIPO_USUARIO TIPO ON USU.IdTipoUsuario = TIPO.IdTipoUsuario) INNER JOIN REGLAS_TIPO REG ON TIPO.IdTipoUsuario = REG.IdTipoActual"
            strSQL = strSQL & " WHERE USU.IdTitular=" & lIdTitular
            'strSQL = strSQL & " AND (Int(DateDiff(""d"", dateadd('m',1,USU.FechaNacio), date())/365)) > EdadMaxima"
            strSQL = strSQL & " AND (Round((DateDiff(""d"",DateAdd(""d"",-1*(Day(FechaNacio)-1),DateAdd(""m"",1,FechaNacio)),GetDate()))/365, 0, 1)) > EdadMaxima"
        #Else
            strSQL = "SELECT USU.IdMember, USU.IdTipoUsuario, USU.FechaNacio, DateDiff(""yyyy"",USU.FechaNacio, Date()) AS Edad ,TIPO.EdadMaxima, REG.Accion, REG.IdTipoNuevo"
            strSQL = strSQL & " FROM (USUARIOS_CLUB USU INNER JOIN TIPO_USUARIO TIPO ON USU.IdTipoUsuario = TIPO.IdTipoUsuario) INNER JOIN REGLAS_TIPO REG ON TIPO.IdTipoUsuario = REG.IdTipoActual"
            strSQL = strSQL & " WHERE USU.IdTitular=" & lIdTitular
            'strSQL = strSQL & " AND (Int(DateDiff(""d"", dateadd('m',1,USU.FechaNacio), date())/365)) > EdadMaxima"
            strSQL = strSQL & " AND (Int((DateDiff(""d"",DateAdd(""d"",-1*(Day(FechaNacio)-1),DateAdd(""m"",1,FechaNacio)),date()))/365) ) > EdadMaxima"
        #End If
        
        adorcsChecaTipos.Open strSQL, Conn, adOpenDynamic, adLockOptimistic
        
        If adorcsChecaTipos.EOF Then Exit Do
    
        Do Until adorcsChecaTipos.EOF
        
            'dFechaComp = CDate("01" & "/" & Month(adorcsChecaTipos!Fechanacio) & "/" & Year(Date))
            
            'dFechaComp = DateAdd("m", 1, dFechaComp)
        
            'If dFechaComp <= Date Then
                adorcsChecaTipos!idtipousuario = adorcsChecaTipos!idtiponuevo
                adorcsChecaTipos.Update
            'End If
            adorcsChecaTipos.MoveNext
        Loop
    
        adorcsChecaTipos.Close
    Loop
    
    Set adorcsChecaTipos = Nothing
    
End Sub
'------------------------------------------------------------------------------


Public Sub VerificaStatus(lidMember As Long)
    
    Dim adorcsVerStat As ADODB.Recordset
    Dim adocmdVerStat As ADODB.Command
    
    Dim sValor As String
    
    #If SqlServer_ Then
        strSQL = "SELECT IdMember"
        strSQL = strSQL & " FROM Ausencias"
        strSQL = strSQL & " WHERE IdMember=" & lidMember
        strSQL = strSQL & " AND FechaInicial < " & "'" & Format(Date, "yyyymmdd") & "'"
    #Else
        strSQL = "SELECT IdMember"
        strSQL = strSQL & " FROM Ausencias"
        strSQL = strSQL & " WHERE IdMember=" & lidMember
        strSQL = strSQL & " AND FechaInicial <" & "#" & Format(Date, "mm/dd/yyyy") & "#"
    #End If
    
    Set adorcsVerStat = New ADODB.Recordset
    adorcsVerStat.CursorLocation = adUseServer
    adorcsVerStat.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If adorcsVerStat.EOF Then
        sValor = "INACTIVO"
    Else
        sValor = vbNullString
    End If
    
    adorcsVerStat.Close
    Set adorcsVerStat = Nothing
    
    
    
    strSQL = "UPDATE USUARIOS_CLUB SET"
    strSQL = strSQL & " Status=" & "'" & sValor & "'"
    
    Set adocmdVerStat = New ADODB.Command
    adocmdVerStat.ActiveConnection = Conn
    adocmdVerStat.CommandType = adCmdText
    adocmdVerStat.CommandText = strSQL
    adocmdVerStat.Execute
    
    Set adocmdVerStat = Nothing
    
End Sub

Public Sub MsjError()
    
    Dim sCadError As String
    Dim lI As Long
    
    sCadError = ""
    
    For lI = 0 To Conn.Errors.Count - 1
        sCadError = sCadError & Conn.Errors.Item(lI).Description & "(" & Conn.Errors.Item(lI).Number & ")" & vbLf
    Next
    
    If sCadError = "" Then
        If Err.Number <> 0 Then
            sCadError = Err.Description
        End If
    End If
    
    MsgBox sCadError, vbCritical, "Error"
    
    
End Sub
