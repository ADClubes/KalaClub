VERSION 5.00
Begin VB.Form frmAltaRenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de artículos rentables"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmAltaRenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmRenta 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdAyuda 
         Height          =   305
         Left            =   960
         Picture         =   "frmAltaRenta.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " Lista de equipos disponibles  "
         Top             =   1200
         Width           =   425
      End
      Begin VB.TextBox txtNoRenta 
         Height          =   285
         Left            =   120
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cbSexo 
         Height          =   315
         ItemData        =   "frmAltaRenta.frx":058C
         Left            =   2520
         List            =   "frmAltaRenta.frx":0599
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox cbTipoRenta 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   615
         Left            =   3600
         Picture         =   "frmAltaRenta.frx":05BE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Salir "
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdGuardar 
         Height          =   615
         Left            =   2760
         Picture         =   "frmAltaRenta.frx":08C8
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Guardar registro "
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblNoRenta 
         Caption         =   "Número"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblSexo 
         Caption         =   "Sexo"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTipoRenta 
         Caption         =   "Rentable tipo:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmAltaRenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de articulos rentables          *
'*  Daniel Hdez                                                 *
'*  11 / Febrero / 2004                                         *
'*  Ult Act: 15 / Agosto / 2005                                 *
'****************************************************************


'Tamaño del campo: NUMERO en la tabla de Rentables
Const TAMCAMPO = 6

Public frmHRenta As frmayuda
Public bNewRent As Boolean

Dim sTextToolBar As String
Dim sTipo As String
Dim sSexo As String
Dim sNumero As String
Dim nTipo As Integer


Private Sub cbTipoRenta_DropDown()
Dim sSql As String

    'Llena el combo con la lista de las delegaciones o municipios
    sSql = "SELECT IdTipoRentable, Descripcion FROM Tipo_Rentables"
    LlenaCombos Me.cbTipoRenta, sSql, "Descripcion", "IdTipoRentable"
End Sub


Private Sub cmdSalir_Click()
Dim Respuesta As Integer

    If (Cambios) Then
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Registro de Arts. rentables")

        If (Respuesta = vbYes) Then
            If (GuardaDatos) Then
                Unload Me
            Else
                Exit Sub
            End If
        End If
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()
    CentraForma MDIPrincipal, frmAltaRenta
    sTextToolBar = Trim(MDIPrincipal.StatusBar1.Panels.Item(1).Text)
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Asignación de artículos rentables"

    ClearCtrls

    'Clave para el nuevo registro
    nCveRent = 0
    
    InitVar
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmAltaRenta.LlenaRenta
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub


Private Sub ClearCtrls()
    With Me
        .cbTipoRenta.Text = ""
        .cbSexo.Text = ""
        .txtNoRenta.Text = ""
    End With
End Sub


Private Sub InitVar()
    sTipo = Trim(Me.cbTipoRenta.Text)
    sSexo = Trim(Me.cbSexo.Text)
    sNumero = Me.txtNoRenta.Text
End Sub


Private Sub cmdGuardar_Click()
    Dim i As Byte
    If Not ChecaSeguridad(Me.Name, Me.cmdGuardar.Name) Then
        Exit Sub
    End If

    If (Cambios) Then
        If (GuardaDatos) Then
            Unload Me
        Else
            MsgBox "No se registraron los datos, verifique la información.", vbCritical, "KalaSystems"
        End If
    End If
End Sub


Private Function ChecaDatos()
    Dim sCond As String
    Dim sTablas As String
    Dim nTitular As Integer
    
    Dim bPropiedad As Boolean
    Dim lNoFamilia As Long
    
    Dim adorcsRenta As ADODB.Recordset

    ChecaDatos = False

    If (Trim(Me.cbTipoRenta.Text) = "") Then
        MsgBox "Se debe seleccionar un tipo de rentable.", vbExclamation, "KalaSystems"
        Me.cbTipoRenta.SetFocus
        Exit Function
    Else
        nTipo = LeeXValor("IdTipoRentable", "Tipo_Rentables", "Descripcion='" & Trim(Me.cbTipoRenta.Text) & "'", "IdTipoRentable", "n", Conn)

        If (nTipo <= 0) Then
            MsgBox "El tipo seleccionado es incorrecto.", vbExclamation, "KalaSystems"
            Me.cbTipoRenta.SetFocus
            Exit Function
        End If
    End If

    If (Trim(Me.cbSexo.Text) = "") Then
        MsgBox "Se debe seleccionar una opción del sexo.", vbExclamation, "KalaSystems"
        Me.cbSexo.SetFocus
        Exit Function
    Else
        If (Trim(UCase(Me.cbSexo.Text)) = "FEMENINO") Then
            sSexo = "F"
        ElseIf (Trim(UCase(Me.cbSexo.Text)) = "MASCULINO") Then
            sSexo = "M"
        ElseIf (Trim(UCase(Me.cbSexo.Text)) = "INDISTINTO") Then
            sSexo = "X"
        Else
            MsgBox "La opción seleccionada en sexo es incorrecta.", vbExclamation, "KalaSystems"
            Me.cbSexo.SetFocus
            Exit Function
        End If
    End If
    
    If (Trim(Me.txtNoRenta.Text) = "") Then
'        If (Not IsNumeric(Me.txtNoRenta.Text)) Then
'            MsgBox "El número del artículo rentable es incorrecto.", vbExclamation, "KalaSystems"
'            Me.txtNoRenta.SetFocus
'            Exit Function
'        End If
'    Else
        MsgBox "El número del rentable rentable no puede quedar vacío.", vbExclamation, "KalaSystems"
        Me.txtNoRenta.SetFocus
        Exit Function
    End If
    
    'Verifica que exista el rentable
    'sCond = "Tipo_Rentables.Descripcion='" & Trim(UCase(Me.cbTipoRenta.Text)) & "' AND "
    'sCond = sCond & "Rentables.Sexo='" & sSexo & "' AND "
    'sCond = sCond & "Rentables.Numero='" & Space(TAMCAMPO - Len(Trim(Me.txtNoRenta.Text))) & Trim(Me.txtNoRenta.Text) & "'"
    'sCond = sCond & "NOT (Rentables.IdUsuario IS NULL)"
    
    'sTablas = "Rentables LEFT JOIN Tipo_Rentables ON Rentables.IdTipoRentable=Tipo_Rentables.IdTipoRentable"
    
    'nTitular = LeeXValor("Rentables!IdUsuario", sTablas, sCond, "Rentables!IdUsuario", "n", Conn)
    
    strSQL = "SELECT USUARIOS_CLUB.NoFamilia, Numero, IdUsuario, Propiedad"
    strSQL = strSQL & " FROM RENTABLES"
    strSQL = strSQL & " LEFT JOIN USUARIOS_CLUB ON RENTABLES.IdUsuario=USUARIOS_CLUB.IdMember"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdTipoRentable=" & Me.cbTipoRenta.ItemData(Me.cbTipoRenta.ListIndex)
    strSQL = strSQL & " AND RENTABLES.Sexo=" & "'" & sSexo & "'"
    strSQL = strSQL & " AND Numero=" & "'" & Space(TAMCAMPO - Len(Trim(Me.txtNoRenta.Text))) & Trim(Me.txtNoRenta.Text) & "'"
    
    
    
    Set adorcsRenta = New ADODB.Recordset
    adorcsRenta.CursorLocation = adUseServer
    
    adorcsRenta.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsRenta.EOF Then
        nTitular = adorcsRenta!idusuario
        lNoFamilia = IIf(IsNull(adorcsRenta!NoFamilia), 0, adorcsRenta!NoFamilia)
        bPropiedad = adorcsRenta!Propiedad
    Else
        nTitular = -1
    End If
    
    adorcsRenta.Close
    Set adorcsRenta = Nothing
    
    'Revisa que exista el articulo selecionado
    If (nTitular = -1) Then
        MsgBox "El rentable especificado no existe, verifique los datos.", vbExclamation, "KalaSystems"
        Me.cbTipoRenta.SetFocus
        Exit Function
        
    'Revisa que no este asignado
    ElseIf (nTitular > 0) Then
        MsgBox "El rentable seleccionado ya está asignado" & vbLf & "al Usuario " & lNoFamilia, vbExclamation, "Verifique"
        Me.cbTipoRenta.SetFocus
        Exit Function
    End If
    
    If bPropiedad Then
        MsgBox "El rentable seleccionado es para uso!", vbExclamation, "Verifique"
        Me.cbTipoRenta.SetFocus
        Exit Function
    End If
    
    
'    If (BuscaXValor("Numero", sTablas, sCond, Conn, "")) Then
'        MsgBox "El artículo especificado no existe o ya está asignado, verifique los datos.", vbExclamation, "KalaSystems"
'        Me.cbTipoRenta.SetFocus
'        Exit Function
'    End If

    ChecaDatos = True
End Function


Private Function Cambios() As Boolean
    Cambios = True
    
    If (sTipo <> Trim(Me.cbTipoRenta.Text)) Then
        Exit Function
    End If

    If (sSexo <> Trim(Me.cbSexo.Text)) Then
        Exit Function
    End If
    
    If (Trim(sNumero) <> Trim(Me.txtNoRenta.Text)) Then
        Exit Function
    End If
    
    Cambios = False
End Function


Private Function GuardaDatos() As Boolean
    Const DATOSRENTA = 2
    Dim bCreado As Boolean
    Dim mFieldsRent(DATOSRENTA) As String
    Dim mValuesRent(DATOSRENTA) As Variant
    Dim sCond As String
    Dim nMes As Byte
    Dim nAnio As Integer
    

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Rentables
    mFieldsRent(0) = "IdUsuario"
    mFieldsRent(1) = "FechaPago"

    'Valores de la tabla Rentables
    mValuesRent(0) = Val(frmAltaSocios.txtTitCve.Text)
    
    'Ultimo día del mes anterior (fecha inicial para pagos)
    nAnio = Year(Date)
    nMes = Month(Date) - 1
    
    If (nMes = 0) Then
        nMes = 12
        nAnio = Year(Date) - 1
    End If
    
    'mValuesRent(1) = Format(UltimoDiaDelMes(CDate("01/" & Trim$(Str(nMes)) & "/" & Trim$(Str(nAnio)))), "dd/mm/yyyy")
    #If SqlServer_ Then
        mValuesRent(1) = Format(Date - 1, "yyyymmdd")
    #Else
        mValuesRent(1) = Date - 1
    #End If
    
    'Condicion para remplazar la clave del titular
    sCond = "IdTipoRentable=" & nTipo & " AND "
    sCond = sCond & "Sexo='" & sSexo & "' AND "
    sCond = sCond & "Numero='" & Space(TAMCAMPO - Len(Trim(Me.txtNoRenta.Text))) & Trim(Me.txtNoRenta.Text) & "'"

    'Asigna la clave del titular al art rentable
    If (CambiaReg("Rentables", mFieldsRent, DATOSRENTA, mValuesRent, sCond, Conn)) Then
        MsgBox "El artículo fué asignado.", vbInformation, "KalaSystems"

        GuardaDatos = True
    Else
        MsgBox "El registro no fue completado.", vbCritical, "KalaSystems"
    End If
End Function


Public Sub LlenaRenta()
Const DATOSRENTA = 5
Dim rsRenta As ADODB.Recordset
Dim sCampos As String
Dim sTablas  As String
Dim mAncRent(DATOSRENTA) As Integer
Dim mEncRent(DATOSRENTA) As String


    frmAltaSocios.ssdbRenta.RemoveAll

    sCampos = "Tipo_Rentables.Descripcion, Rentables.Sexo, "
    sCampos = sCampos & "Rentables.Numero, Rentables.Ubicacion, Rentables.FechaPago "
    
    sTablas = "Rentables LEFT JOIN Tipo_Rentables ON Rentables.IdTipoRentable=Tipo_Rentables.IdTipoRentable "
    
    InitRecordSet rsRenta, sCampos, sTablas, "Rentables.idUsuario=" & Val(frmAltaSocios.txtTitCve.Text), "", Conn
    With rsRenta
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                frmAltaSocios.ssdbRenta.AddItem .Fields("Descripcion") & vbTab & _
                .Fields("Sexo") & vbTab & _
                .Fields("Numero") & vbTab & _
                .Fields("Ubicacion") & vbTab & _
                .Fields("FechaPago")
                
                .MoveNext
            Loop
        End If
    
        .Close
    End With
    Set rsRenta = Nothing

    'Asigna valores a la matriz de encabezados
    mEncRent(0) = "Tipo"
    mEncRent(1) = "Sexo"
    mEncRent(2) = "Número"
    mEncRent(3) = "Ubicación"
    mEncRent(4) = "Pagado hasta"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid frmAltaSocios.ssdbRenta, mEncRent

    'Asigna valores a la matriz que define el ancho de cada columna
    mAncRent(0) = 2000
    mAncRent(1) = 1100
    mAncRent(2) = 1100
    mAncRent(3) = 1100
    mAncRent(4) = 1500

    'Asigna el ancho de cada columna
    DefAnchossGrid frmAltaSocios.ssdbRenta, mAncRent
End Sub


Public Sub QuitaRenta()
    Const DATOSRENTA = 1
    Dim mFieldsRent(DATOSRENTA) As String
    Dim mValuesRent(DATOSRENTA) As Variant
    Dim sCond As String
    
    
    '08/04/2008 gpo
    Dim dFechaVigencia As Date

    'Campos de la tabla Rentables
    mFieldsRent(0) = "idUsuario"

    'Valores de la tabla Rentables
    mValuesRent(0) = 0
    
    'Condicion para quitar la clave del titular
'    sCond = "IdTipoRentable=" & LeeXValor("IdTipoRentable", "Tipo_Rentables", "Descripcion='" & Trim(frmAltaSocios.adoRenta.Recordset.Fields("Tipo_Rentables!Descripcion")) & "'", "IdTipoRentable", "n", Conn) & " AND "
'    sCond = sCond & "Sexo='" & frmAltaSocios.adoRenta.Recordset.Fields("Rentables!Sexo") & "' AND "
'    sCond = sCond & "Numero='" & frmAltaSocios.adoRenta.Recordset.Fields("Rentables!Numero") & "' AND "
'    sCond = sCond & "IdUsuario = " & Val(frmAltaSocios.txtTitCve.Text)
    
    sCond = "idTipoRentable=" & LeeXValor("idTipoRentable", "Tipo_Rentables", "Descripcion='" & Trim$(frmAltaSocios.ssdbRenta.Columns(0).Text) & "'", "IdTipoRentable", "n", Conn) & " AND "
    sCond = sCond & "Sexo='" & frmAltaSocios.ssdbRenta.Columns(1).Text & "' AND "
    sCond = sCond & "Numero='" & frmAltaSocios.ssdbRenta.Columns(2).Text & "' AND "
    sCond = sCond & "IdUsuario = " & Val(frmAltaSocios.txtTitCve.Text)
    
    
    If VigenciaRentable(frmAltaSocios.ssdbRenta.Columns(2).Text, dFechaVigencia) Then
        If dFechaVigencia >= Date Then
            MsgBox "El rentable seleccionado aun está vigente!", vbExclamation, "Verifique"
            Exit Sub
        End If
    
        
    End If

    'Asigna la clave del titular al art rentable
    If (CambiaReg("Rentables", mFieldsRent, DATOSRENTA, mValuesRent, sCond, Conn)) Then
        MsgBox "El artículo fué retirado.", vbInformation, "KalaSystems"
    Else
        MsgBox "El cambio no se realizó.", vbCritical, "KalaSystems"
    End If
End Sub




'************************************************************
'*                          Ayudas                          *
'************************************************************

Private Sub cmdAyuda_Click()
Const DATOSRENTA = 4
Dim sCadena As String
Dim sChar As String
Dim mFAyuda(DATOSRENTA) As String
Dim mAAyuda(DATOSRENTA) As Integer
Dim mCAyuda(DATOSRENTA) As String
Dim mEAyuda(DATOSRENTA) As String

    nAyuda = 1

    Set frmHRenta = New frmayuda
    
    mFAyuda(0) = "Rentables ordenados por tipo"
    mFAyuda(1) = "Rentables ordenados por número"
    mFAyuda(2) = "Rentables ordenados por sexo"
    mFAyuda(3) = "Rentables ordenados por ubicación"
    
    mAAyuda(0) = 1800
    mAAyuda(1) = 800
    mAAyuda(2) = 800
    mAAyuda(3) = 2000
    
    mCAyuda(0) = "Tipo_Rentables.Descripcion"
    mCAyuda(1) = "Rentables.Numero"
    mCAyuda(2) = "Rentables.Sexo"
    mCAyuda(3) = "Rentables.Ubicacion"
    
    mEAyuda(0) = "Tipo"
    mEAyuda(1) = "Número"
    mEAyuda(2) = "Sexo"
    mEAyuda(3) = "Ubicación"
    
    With frmHRenta
        .nColActiva = 0
        .nColsAyuda = DATOSRENTA
        .sTabla = "Rentables LEFT JOIN Tipo_Rentables ON Rentables.IdTipoRentable=Tipo_Rentables.IdTipoRentable"
        
        .sCondicion = "IdUsuario=0" & " AND Propiedad=0"
        If (Trim(Me.cbTipoRenta.Text) <> "") Then
            .sCondicion = .sCondicion & " AND Tipo_Rentables.Descripcion='" & Trim(Me.cbTipoRenta.Text) & "'"
        End If
        
        If (Trim(Me.cbSexo.Text) <> "") Then
            Select Case Trim(Me.cbSexo.Text)
                Case "FEMENINO"
                    sChar = "F"
                    
                Case "MASCULINO"
                    sChar = "M"
                    
                Case "INDISTINTO"
                    sChar = "X"
            End Select
        
            .sCondicion = .sCondicion & " AND Rentables.Sexo='" & sChar & "'"
        End If
        
        If (Trim(Me.txtNoRenta.Text <> vbNullString)) Then
            #If SqlServer_ Then
                .sCondicion = .sCondicion & " AND RTRIM(LTRIM(Rentables.Numero)) LIKE '" & Trim(UCase(Me.txtNoRenta.Text)) & "%'"
            #Else
                .sCondicion = .sCondicion & " AND TRIM(Rentables.Numero) LIKE '" & Trim(UCase(Me.txtNoRenta.Text)) & "%'"
            #End If
        End If
        
        .sTitAyuda = "Art. rentables aún no asignados"
        .lAgregar = False
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
    If (Trim(Me.txtNoRenta.Text) <> Trim(sNumero)) Then
        sNumero = Me.txtNoRenta.Text
    End If
End Sub

