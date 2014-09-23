VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCredyPases 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Credenciales y pases de olvido"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   Icon            =   "frmCredyPases.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbTipoPase 
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   4215
      DataFieldList   =   "Column 0"
      AllowInput      =   0   'False
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   5874
      Columns(0).Caption=   "Credencial"
      Columns(0).Name =   "Credencial"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Visible=   0   'False
      Columns(1).Caption=   "IdTipoCredencial"
      Columns(1).Name =   "IdTipoCredencial"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "TieneCredencial"
      Columns(2).Name =   "TieneCredencial"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "DiasMaximo"
      Columns(3).Name =   "DiasMaximo"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "RequiereNombre"
      Columns(4).Name =   "RequiereNombre"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "VecesSinCosto"
      Columns(5).Name =   "VecesSinCosto"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "IdConcepto"
      Columns(6).Name =   "IdConcepto"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      _ExtentX        =   7435
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.CommandButton cmdRenovar 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   615
      Left            =   7920
      Picture         =   "frmCredyPases.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdElimina 
      Height          =   615
      Left            =   6240
      Picture         =   "frmCredyPases.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdBloquea 
      Height          =   615
      Left            =   7080
      Picture         =   "frmCredyPases.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Bloquea"
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton optTodos 
      Caption         =   "Todos"
      Height          =   255
      Left            =   7440
      TabIndex        =   19
      Top             =   1080
      Width           =   1215
   End
   Begin VB.OptionButton optSoloMes 
      Caption         =   "Mes Actual"
      Height          =   255
      Left            =   5640
      TabIndex        =   18
      Top             =   1080
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalir 
      Height          =   615
      Left            =   8760
      Picture         =   "frmCredyPases.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Salir"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdAgrega 
      Height          =   615
      Left            =   5400
      Picture         =   "frmCredyPases.frx":12EA
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Nuevo"
      Top             =   120
      Width           =   615
   End
   Begin VB.ComboBox cmbUsuario 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8400
      TabIndex        =   12
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   6960
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtNombre 
      Height          =   375
      Left            =   240
      MaxLength       =   60
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSComCtl2.DTPicker dtpFechaFin 
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102236161
      CurrentDate     =   38686
   End
   Begin MSComCtl2.DTPicker dtpFechaIni 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   102236161
      CurrentDate     =   38686
   End
   Begin VB.TextBox txtNoPase 
      Height          =   375
      Left            =   240
      MaxLength       =   5
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgCreden 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   9375
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   3200
      Columns(0).Visible=   0   'False
      Columns(0).Caption=   "IdReg"
      Columns(0).Name =   "IdReg"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2249
      Columns(1).Caption=   "FechaAlta"
      Columns(1).Name =   "FechaAlta"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2619
      Columns(2).Caption=   "IniciaVigencia"
      Columns(2).Name =   "IniciaVigencia"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "TerminaVigencia"
      Columns(3).Name =   "TerminaVigencia"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   5874
      Columns(4).Caption=   "Nombre"
      Columns(4).Name =   "Nombre"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2196
      Columns(5).Caption=   "Numero"
      Columns(5).Name =   "Numero"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   16536
      _ExtentY        =   3413
      _StockProps     =   79
      BeginProperty PageFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PageHeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ver"
      Height          =   495
      Left            =   5040
      TabIndex        =   20
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label lblNoReg 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6960
      TabIndex        =   21
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Credencial o pase"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label6 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblVigenteHasta 
      Caption         =   "Vigente Hasta"
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblVigenteDesde 
      Caption         =   "Vigente Desde"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblTipoPaseNom 
      Caption         =   "Tipo de pase"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblTipoPase 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblNoPase 
      Caption         =   "# Pase"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmCredyPases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lIdTitular As Long
Public lIdMemberSel As Long
Dim lSecuencialPase As Long
Dim iIdTipoPase As Integer
Dim boActivaBotonRenovar As Boolean



Private Sub cmbUsuario_Click()
    LlenaGridCred
End Sub

Private Sub cmdAgrega_Click()
    
    If Not ChecaSeguridad(Me.Name, Me.cmdAgrega.Name) Then
        Exit Sub
    End If
    
    
        
    
    If Not AccesoValido(Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)) Then
        MsgBox "¡No se puede asignar a este usuario!", vbCritical, "Verifique"
        Exit Sub
    End If
    
    If Me.optTodos.Value = True Then
        MsgBox "Seleccione la opción de vista mensual" & vbLf & "para poder agregar", vbCritical, "Verifique"
        Exit Sub
    End If
    
    
    LimpiaControles
    ActivaBotones False
    ActivaControles True
    
    
    
    
    
    
    Me.txtNombre.Text = Trim(Me.cmbUsuario.Text)
    Me.txtNoPase.SetFocus
    lSecuencialPase = 0
    iIdTipoPase = 0
    
    'Para reposición de credencial
    If Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value = 0 Then
    
        If ObtieneParametro("CAMBIOCREDENCIAL") = "1" Then
            MsgBox "Si el usuario no ha renovado su credencial," & vbCrLf & _
                "por favor imprimala con el nuevo formato" & vbCrLf & _
                "y al entregarla realice el proceso de renovación." & vbCrLf & _
                "Si ya renovó, continue con el proceso de reposición," & vbCrLf & _
                "y al entregarla realice el proceso de renovación", vbExclamation, "Verifique"
        End If
        
        Me.txtNoPase.Text = "0"
        Me.txtNoPase.Enabled = False
        Me.txtNombre.Enabled = False
        Me.dtpFechaFin.Value = CDate("31/12/" & Year(Date))
    End If
    
End Sub

Private Sub cmdBloquea_Click()
    Dim iResp As Integer
    Dim sTor As String
    Dim nErrCode As Long
    
     If Not ChecaSeguridad(Me.Name, Me.cmdBloquea.Name) Then
        Exit Sub
    End If
    
    If Me.ssdbgCreden.Rows = 0 Then
        Exit Sub
    End If
    
    iResp = MsgBox("¿Bloquear el pase " & Me.ssdbgCreden.Columns("Numero").Value & "?", vbQuestion + vbYesNo, "Confirme")
    
    If iResp <> vbYes Then
        Exit Sub
    End If
    
    #If SqlServer_ Then
        ActivaCredSQL 1, Val(Me.ssdbgCreden.Columns("Numero").Value), 1, 0, False, True
    #Else
        ActivaCred 1, Val(Me.ssdbgCreden.Columns("Numero").Value), 1, 0, False, True
    #End If
    sTor = Mid(Val(Me.ssdbgCreden.Columns("Numero").Value), 4, 7)
    
             'nErrCode = EliminarAcceso(sTor)
             
    'If nErrCode <> 0 Then
    '            MsgBox "No se pudo registrar el usuario en torniquetes,Favor de hacerlo manual"
    'End If
End Sub

Private Sub cmdCancelar_Click()
    If Not Me.txtNoPase.Enabled Then
        Me.txtNoPase.Enabled = True
    End If
    
    ActivaControles False
    ActivaBotones True
    
End Sub



Private Sub cmdElimina_Click()
     If Not ChecaSeguridad(Me.Name, Me.cmdElimina.Name) Then
        Exit Sub
    End If
End Sub

Private Sub cmdOk_Click()

    Dim iRespuesta As Integer
    Dim lDiasTrans As Long
    
    

    If Me.txtNoPase.Enabled And Me.txtNoPase.Text = vbNullString Then
        MsgBox "Indicar un número de pase!", vbExclamation, "Verifique"
        Me.txtNoPase.SetFocus
        Exit Sub
    End If
    
    
    If Me.dtpFechaFin.Value < Me.dtpFechaIni.Value Then
        MsgBox "La fecha de terminación debe ser mayor que la de inicio!", vbExclamation, "Verifique"
        Me.dtpFechaFin.SetFocus
        Exit Sub
    End If
    
    
    If DateDiff("d", Me.dtpFechaIni.Value, Me.dtpFechaFin.Value) > Val(Me.ssCmbTipoPase.Columns("DiasMaximo").Value) Then
        MsgBox "El número máximo de días es " & Me.ssCmbTipoPase.Columns("DiasMaximo").Value & "dia(s)", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    
    If Me.txtNombre.Enabled And Me.txtNombre.Text = vbNullString Then
        MsgBox "Indicar un Nombre!", vbExclamation, "Verifique"
        Me.txtNombre.SetFocus
        Exit Sub
    End If
    
    
    If Me.ssCmbTipoPase.Columns("TieneCredencial").Value = 1 Then
        If Not ValidaNumeroPase(CSng(Me.txtNoPase.Text)) Then
            Me.txtNoPase.SetFocus
            Exit Sub
        End If
    
    
    
        If iIdTipoPase <> Me.ssCmbTipoPase.Columns("IdtipoCredencial").Value Then
            MsgBox "Este pase no coincide con el tipo indicado", vbExclamation, "Verifique"
            Me.txtNoPase.SetFocus
            Exit Sub
        End If
    
    
    
      
    
        If lSecuencialPase = 0 And Not Me.txtNoPase.Enabled Then
            If MsgBox("¿Está seguro que desea proceder con la reposición de credencial?", vbQuestion + vbYesNo, "Confirmar") = vbNo Then
                Exit Sub
            End If
        End If
    Else
        iIdTipoPase = Val(Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value)
    End If
    
    InsertaPase
    LlenaGridCred
    
    
    
    If Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value = 0 Then
        iRespuesta = MsgBox("Se generará un cargo por reposición?", vbQuestion + vbYesNo, "Confirme")
        If iRespuesta = vbYes Then
            If Not InsertaCargoVario(lIdTitular, 102, "REPOSICION DE CREDENCIAL " & Me.txtNoPase.Text, 0, Date, Date) Then
                MsgBox "Ocurrio un error al insertar el cargo!", vbCritical, "Error"
            End If
        End If
    End If
    
    
    ''Aqui
    If Me.ssdbgCreden.Rows > Val(Me.ssCmbTipoPase.Columns("VecesSinCosto").Value) Then
        If Not InsertaCargoVario(lIdTitular, Val(Me.ssCmbTipoPase.Columns("IdConcepto").Value), Me.ssCmbTipoPase.Columns("Credencial").Value & " " & Me.txtNoPase.Text, -1, Date, Date) Then
            MsgBox "Ocurrio un error al insertar el cargo!", vbCritical, "Error"
        End If
    Else
        If Not InsertaCargoVario(lIdTitular, Val(Me.ssCmbTipoPase.Columns("IdConcepto").Value), Me.ssCmbTipoPase.Columns("Credencial").Value & " " & Me.txtNoPase.Text, 0, Date, Date) Then
            MsgBox "Ocurrio un error al insertar el cargo!", vbCritical, "Error"
        End If
    End If
    
    
'    If Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value = 1 Then
'        Dim iPasesSinCosto As Integer
'
'        iPasesSinCosto = Val(ObtieneParametro("PASES SIN COSTO AL MES"))
'
'        If Me.ssdbgCreden.Rows > iPasesSinCosto Then
'            MsgBox "Se generará un cargo por este pase", vbInformation, "Aviso"
'            If Not InsertaCargoVario(lIdTitular, 100, "OLVIDO DE CREDENCIAL " & Me.txtNoPase.Text, 0, Date, Date) Then
'                MsgBox "Ocurrio un error al insertar el cargo!", vbCritical, "Error"
'            End If
'        End If
'    End If
'
'    If Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value = 2 Then
'        iRespuesta = MsgBox("Se generará un cargo por este invitado?", vbQuestion + vbYesNo, "Confirme")
'        If iRespuesta = vbYes Then
'            If Not InsertaCargoVario(lIdTitular, 101, "INVITADO POR DIA " & Me.txtNoPase.Text, 0, Date, Date) Then
'                MsgBox "Ocurrio un error al insertar el cargo!", vbCritical, "Error"
'            End If
'        End If
'    End If
    
    
    
    
    MsgBox Me.ssCmbTipoPase.Text & vbLf & "agregado y activado", vbInformation, "Correcto"
    
    
    If Not Me.txtNoPase.Enabled Then
        Me.txtNoPase.Enabled = True
    End If
    
    ActivaControles False
    ActivaBotones True
    
    
    

End Sub

Private Sub cmdRenovar_Click()
    Dim iResp As Integer
    Dim adorcsRenovar As ADODB.Recordset
    Dim adocmdRenovar As ADODB.Command
    Dim lSecAnt As Long
    Dim lSecNuevo As Long
    Dim lNumeroFoto As Long
    Dim nErrCode As Long
    Dim sTor As String
    Dim nTor As Long
    
    
    Open Trim(Environ("APPDATA")) & "\KALACLUB\" & "kalalog.txt" For Output As #1
    Print #1, "Inicia Renovar " & CStr(Now)
    
    Print #1, "ChecaSeguridad"
    If Not ChecaSeguridad(Me.Name, Me.cmdRenovar.Name) Then
        Print #1, "Termina Renovar " & CStr(Now)
        Close #1
        Exit Sub
    End If
    Print #1, "OK"
    
    On Error GoTo Error_Catch
    
    iResp = MsgBox("Si procede, la credencial actual quedará inactiva y se activara la nueva", vbInformation + vbOKCancel, "Confirme")
    
    If iResp = vbCancel Then
        Print #1, "Termina Renovar " & CStr(Now)
        Close #1
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    
    'Verifica que no se hay realizado la renovacion de credencial
    strSQL = "SELECT Secuencial, SecuencialAnterior"
    strSQL = strSQL & " FROM Secuencial_Nuevo"
    strSQL = strSQL & " WHERE IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
    
    Set adorcsRenovar = New ADODB.Recordset
    adorcsRenovar.CursorLocation = adUseServer
    
    Print #1, "Verificar que no se hay realizado la renovacion de credencial"
    adorcsRenovar.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    Print #1, "OK"
    
    If adorcsRenovar.EOF Then
        Screen.MousePointer = vbDefault
        adorcsRenovar.Close
        Set adorcsRenovar = Nothing
        MsgBox "Usuario sin código asignado", vbCritical, "Error"
        
        Print #1, "Termina Renovar " & CStr(Now)
        Close #1
        Exit Sub
    End If
    
    lSecAnt = IIf(IsNull(adorcsRenovar!SecuencialAnterior), 0, adorcsRenovar!SecuencialAnterior)
    lSecNuevo = adorcsRenovar!Secuencial
    
    adorcsRenovar.Close
    
    If lSecAnt > 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Este Usuario YA renovo su credencial", vbCritical, "Verifique"
        
        Print #1, "Termina Renovar " & CStr(Now)
        Close #1
        Exit Sub
    End If
    
    
    'Obtiene el secuencial actual
    strSQL = "SELECT Secuencial"
    strSQL = strSQL & " FROM SECUENCIAL"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
    
    Print #1, "Obtener el secuencial actual"
    adorcsRenovar.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    Print #1, "OK"
    
    lSecAnt = adorcsRenovar!Secuencial
    
    adorcsRenovar.Close
    
    
    'Obtiene el número de Foto del usuario
    strSQL = "SELECT FotoFile"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
    
    Print #1, "Obtener el número de Foto del usuario"
    adorcsRenovar.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    Print #1, "OK"
    
    
    If Not adorcsRenovar.EOF Then
        lNumeroFoto = adorcsRenovar!FotoFile
    End If
    
    adorcsRenovar.Close
    
    'Desactiva la credencial actual
    #If SqlServer_ Then
        ActivaCredSQL 1, lSecAnt, 1, Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), False, False
    #Else
        ActivaCred 1, lSecAnt, 1, Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), False, False
    #End If
    
    'sTor = Mid(lSecAnt, 1, 2)
    'sTor = sTor & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
    
    'nErrCode = EliminarAcceso(sTor)
  
    'Desasigna el secuencial actual en Secuencial
    strSQL = "UPDATE SECUENCIAL SET"
    strSQL = strSQL & " IdMember=0"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " Secuencial=" & lSecAnt
    
    Print #1, "Desasignar el secuencial actual en Secuencial"
    Set adocmdRenovar = New ADODB.Command
    adocmdRenovar.ActiveConnection = Conn
    adocmdRenovar.CommandType = adCmdText
    adocmdRenovar.CommandText = strSQL
    adocmdRenovar.Execute
    Print #1, "OK"
    
    
    'Asigna el secuencial nuevo en Secuencial
    strSQL = "UPDATE SECUENCIAL SET"
    strSQL = strSQL & " IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " Secuencial=" & lSecNuevo
    
    Print #1, "Asignar el secuencial nuevo en Secuencial"
    adocmdRenovar.CommandText = strSQL
    adocmdRenovar.Execute
    Print #1, "OK"
    
    'Marca en secuencial nuevo que se hizo el cambio de credencial
    strSQL = "UPDATE SECUENCIAL_NUEVO SET"
    strSQL = strSQL & " SecuencialAnterior=" & lSecAnt
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
    
    Print #1, "Marcar en secuencial nuevo que se hizo el cambio de credencial"
    adocmdRenovar.CommandText = strSQL
    adocmdRenovar.Execute
    Print #1, "OK"
    
    
    'Inserta la nueva credencial
    #If SqlServer_ Then
        strSQL = "INSERT INTO CREDENCIALES ("
        strSQL = strSQL & " IdMember,"
        strSQL = strSQL & " IdTipoCredencial,"
        strSQL = strSQL & " FechaAlta,"
        strSQL = strSQL & " IniciaVigencia,"
        strSQL = strSQL & " TerminaVigencia,"
        strSQL = strSQL & " Secuencial,"
        strSQL = strSQL & " FotoFile,"
        strSQL = strSQL & " Nombre,"
        strSQL = strSQL & " Numero,"
        strSQL = strSQL & " SecuencialAnterior)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex) & ","
        strSQL = strSQL & Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value & ","
        strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Format(CDate("31/12/2012"), "yyyymmdd") & "',"
        strSQL = strSQL & lSecNuevo & ","
        strSQL = strSQL & lNumeroFoto & ","
        strSQL = strSQL & "'" & Trim(UCase(Me.cmbUsuario.Text)) & "',"
        strSQL = strSQL & "0" & ","
        strSQL = strSQL & lSecAnt & ")"
    #Else
        strSQL = "INSERT INTO CREDENCIALES ("
        strSQL = strSQL & " IdMember,"
        strSQL = strSQL & " IdTipoCredencial,"
        strSQL = strSQL & " FechaAlta,"
        strSQL = strSQL & " IniciaVigencia,"
        strSQL = strSQL & " TerminaVigencia,"
        strSQL = strSQL & " Secuencial,"
        strSQL = strSQL & " FotoFile,"
        strSQL = strSQL & " Nombre,"
        strSQL = strSQL & " Numero,"
        strSQL = strSQL & " SecuencialAnterior)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex) & ","
        strSQL = strSQL & Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value & ","
        strSQL = strSQL & "'" & Format(Date, "dd/mm/yyyy") & "',"
        strSQL = strSQL & "'" & Format(Date, "dd/mm/yyyy") & "',"
        strSQL = strSQL & "'" & Format(CDate("31/12/2012"), "dd/mm/yyyy") & "',"
        strSQL = strSQL & lSecNuevo & ","
        strSQL = strSQL & lNumeroFoto & ","
        strSQL = strSQL & "'" & Trim(UCase(Me.cmbUsuario.Text)) & "',"
        strSQL = strSQL & "0" & ","
        strSQL = strSQL & lSecAnt & ")"
    #End If
    
    Print #1, "Insertar la nueva credencial"
    adocmdRenovar.CommandText = strSQL
    adocmdRenovar.Execute
    Print #1, "OK"
    
    'Activa la credencial nueva
    Print #1, "Activar la nueva credencial"
    #If SqlServer_ Then
        ActivaCredSQL 1, lSecNuevo, 1, Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), True, False
    #Else
        ActivaCred 1, lSecNuevo, 1, Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), True, False
    #End If
    Print #1, "OK"
    
    'nTor = ("68" * 16777216) + lSecNuevo
    
    'nErrCode = AgregaAcceso(Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), Trim(UCase(Me.cmbUsuario.Text)), nTor)
   
    Set adorcsRenovar = Nothing
    Set adocmdRenovar = Nothing
    
    Screen.MousePointer = vbDefault
    
    MsgBox "Credencial renovada", vbInformation, "Correcto"
    
    Print #1, "LlenaGridCred"
    LlenaGridCred
    Print #1, "OK"
    
    On Error GoTo 0
    
    Print #1, "Termina Renovar " & CStr(Now)
    Close #1
    Exit Sub

Error_Catch:
    
    Screen.MousePointer = vbDefault
    
    Print #1, "Error en Renovar " & CStr(Now) & Err.Description
    Close #1
    
    MsgError
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    
    Dim sValor As String
    
    LlenaComboTipoCred
    LlenaComboUsuarios
    
    sValor = ObtieneParametro("ACTIVA RENOVAR CREDENCIAL")
    
    If sValor = vbNullString Then
        sValor = 0
    End If
    
    boActivaBotonRenovar = CBool(sValor)
    
End Sub

Private Sub LlenaComboUsuarios()
    Dim adorcs As ADODB.Recordset
    Dim lIndex As Long
    
    strSQL = "SELECT IdMember, NOMBRE, A_PATERNO, A_MATERNO"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdTitular=" & lIdTitular
    strSQL = strSQL & " ORDER BY NumeroFamiliar"
    
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.cmbUsuario.Clear
    
    Do Until adorcs.EOF
        Me.cmbUsuario.AddItem Trim(adorcs!Nombre) & " " & Trim(adorcs!A_Paterno) & " " & Trim(adorcs!A_Materno)
        Me.cmbUsuario.ItemData(Me.cmbUsuario.NewIndex) = adorcs!IdMember
        If adorcs!IdMember = lIdMemberSel Then
            lIndex = Me.cmbUsuario.NewIndex
        End If
        adorcs.MoveNext
    Loop
    
    
    Me.cmbUsuario.ListIndex = lIndex
    
    adorcs.Close
    
    Set adorcs = Nothing

End Sub

Private Sub LlenaGridCred()
    Dim adorcs As ADODB.Recordset
    
    
    'if Me.ssCmbTipoPase.Bookmark < 0 Then Exit Sub
    If Me.cmbUsuario.ListIndex < 0 Then Exit Sub
    
    strSQL = "SELECT IdReg, FechaAlta, IniciaVigencia, TerminaVigencia, Nombre, Numero"
    strSQL = strSQL & " FROM CREDENCIALES"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
    strSQL = strSQL & " AND IdTipoCredencial=" & Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value
    If Me.optSoloMes Then
        #If SqlServer_ Then
            strSQL = strSQL & " AND Month(FechaAlta)=Month(getDate()) AND Year(FechaAlta)=Year(getDate())"
        #Else
            strSQL = strSQL & " AND Month(FechaAlta)=Month(Date()) AND Year(FechaAlta)=Year(Date())"
        #End If
    End If
    strSQL = strSQL & " ORDER BY IdReg, TerminaVigencia"
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Me.ssdbgCreden.RemoveAll
    
    Do Until adorcs.EOF
        Me.ssdbgCreden.AddItem adorcs!IdReg & vbTab & adorcs!FechaAlta & vbTab & adorcs!IniciaVigencia & vbTab & adorcs!TerminaVigencia & vbTab & adorcs!Nombre & vbTab & adorcs!Numero
        adorcs.MoveNext
    Loop
    
    
    adorcs.Close
    
    Set adorcs = Nothing
    
    Me.lblNoReg.Caption = Me.ssdbgCreden.Rows & " Registro(s)"
    
    If Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value = 0 Then
        If boActivaBotonRenovar Then
            Me.cmdRenovar.Enabled = True
        End If
    Else
        Me.cmdRenovar.Enabled = False
    End If
    
    
End Sub

Private Sub ActivaControles(bValor As Boolean)

    Me.lblNoPase.Visible = bValor
    Me.lblTipoPaseNom.Visible = bValor
    Me.lblTipoPase.Visible = bValor
    Me.lblVigenteDesde.Visible = bValor
    Me.lblVigenteHasta.Visible = bValor
    Me.lblTipoPase.Visible = bValor
    
    
    
    Me.lblTipoPaseNom.Visible = bValor
    


    Me.txtNoPase.Visible = bValor
    Me.dtpFechaIni.Visible = bValor
    Me.dtpFechaFin.Visible = bValor
    Me.txtNombre.Visible = bValor
    Me.cmdOk.Visible = bValor
    Me.cmdCancelar.Visible = bValor
    
    
    
    If Me.ssCmbTipoPase.Columns("RequiereNombre").Value = 1 And bValor Then
        Me.lblNombre.Visible = True
        Me.txtNombre.Visible = True
    Else
        Me.lblNombre.Visible = False
        Me.txtNombre.Visible = False
    End If
    

End Sub


Private Function ValidaNumeroPase(lNumeroPase As Long) As Boolean

    Dim adorcs As ADODB.Recordset
    Dim lDiasTrans As Long
    
    Dim dFechaVigencia As Date
    
    Dim lDiasMinimo As Long
    Dim sInter As String

    ValidaNumeroPase = True
    
    lDiasMinimo = 45
    
    If Me.txtNoPase.Enabled = False Then
        lSecuencialPase = 0
        Exit Function
    End If
    
    
    sInter = ObtieneParametro("DIAS ROTACION PASES")
    
    If sInter <> vbNullString Then
        lDiasMinimo = CSng(sInter)
    End If
    
    
    
    strSQL = "SELECT IdTipoPase, Secuencial"
    strSQL = strSQL & " FROM PASES"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " NumeroPase=" & Trim(Me.txtNoPase)

    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly


    If adorcs.EOF Then
        adorcs.Close
        Set adorcs = Nothing
        MsgBox "# de pase inexistente!", vbExclamation, "Verifique"
        ValidaNumeroPase = False
        Exit Function
    End If


    lSecuencialPase = adorcs!Secuencial
    iIdTipoPase = adorcs!IdTipoPase

    adorcs.Close


    strSQL = "SELECT TOP 1 TerminaVigencia"
    strSQL = strSQL & " FROM CREDENCIALES"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " Numero=" & Trim(Me.txtNoPase)
    strSQL = strSQL & " ORDER BY TerminaVigencia DESC"


    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly


    If Not adorcs.EOF Then
    
        dFechaVigencia = adorcs!TerminaVigencia
    
        lDiasTrans = DateDiff("d", dFechaVigencia, Date)
    
        Select Case lDiasTrans
            Case Is <= 0
                adorcs.Close
                Set adorcs = Nothing
                MsgBox "Este pase está asignado y vigente!", vbExclamation, "Verificar"
                ValidaNumeroPase = False
                Exit Function
            Case 1 To lDiasMinimo
                adorcs.Close
                Set adorcs = Nothing
                MsgBox "No ha pasado el tiempo mínimo" & vbLf & "para volver a usar este pase!" & vbLf & "(" & dFechaVigencia & ") " & lDiasTrans & " dias", vbExclamation, "Verificar"
                ValidaNumeroPase = False
                Exit Function
        End Select
    End If
    
    adorcs.Close
    Set adorcs = Nothing

    
    
    

End Function


Private Sub InsertaPase()

    Dim adorcs As ADODB.Recordset
    Dim adocmd As ADODB.Command
    
    Dim lNumeroFoto As Long
    Dim lSecuencialAnterior As Long
    
    Dim lSecuencialNuevo As Long
    
    
    Dim lIdTipoPase As Long
    
    
    Dim sCambioCredencial As String
    Dim lCambioYaHecho As Boolean
    
    lNumeroFoto = 0
    
    
    lSecuencialAnterior = 0
    lSecuencialNuevo = 0
    
    lIdTipoPase = Me.ssCmbTipoPase.Columns("IdTipoCredencial").Value
    
    
    
    'Obtiene el número de Foto del usuario
    strSQL = "SELECT FotoFile"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not adorcs.EOF Then
        lNumeroFoto = adorcs!FotoFile
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    
    'Si es un pase por olvido, o una reposición de credencial obtiene
    'el número de Secuencial actual
    If lIdTipoPase < 2 Then
        strSQL = "SELECT Secuencial"
        strSQL = strSQL & " FROM SECUENCIAL"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
        
        Set adorcs = New ADODB.Recordset
        adorcs.CursorLocation = adUseServer
        adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        
        If Not adorcs.EOF Then
            lSecuencialAnterior = adorcs!Secuencial
        End If
    
        adorcs.Close
        Set adorcs = Nothing
    
    End If
    
    
    sCambioCredencial = ObtieneParametro("CAMBIOCREDENCIAL")
    
    'Si lSecuencialPase es 0, se trata de una reposicion de
    'credencial
    If lSecuencialPase = 0 And lIdTipoPase = 0 Then
    
        
        'Si está activo el cambio de credencial
        If sCambioCredencial = "1" Then
            'Verifica si ya hizo el cambio de credencial
            strSQL = "SELECT Secuencial, SecuencialAnterior"
            strSQL = strSQL & " FROM SECUENCIAL_NUEVO"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
        
            Set adorcs = New ADODB.Recordset
            adorcs.CursorLocation = adUseServer
            adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        
            'Ya hizo el cambio de credencial
            If adorcs!SecuencialAnterior <> 0 Then
                lCambioYaHecho = True
            End If
    
            adorcs.Close
            Set adorcs = Nothing
        End If
        
        
        'En tabla secuencial
        lSecuencialPase = AsignaSec(Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), False)
        
        If lCambioYaHecho Then
            
            'Desasigna el secuencial anterior
            If lSecuencialAnterior > 0 Then
                
                strSQL = "UPDATE SECUENCIAL_NUEVO SET"
                strSQL = strSQL & " IdMember = 0" & ","
                strSQL = strSQL & " SecuencialAnterior = 0"
                strSQL = strSQL & " WHERE"
                strSQL = strSQL & " IdMember=" & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex)
            
                Set adocmd = New ADODB.Command
                adocmd.ActiveConnection = Conn
                adocmd.CommandType = adCmdText
                adocmd.CommandText = strSQL
                adocmd.Execute
    
                Set adocmd = Nothing
                
                'En tabla secuencial nuevo
                lSecuencialNuevo = AsignaSecNuevo(Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), False)
                
            End If
        
        End If
        
        'Desasigna el secuencial anterior
        If lSecuencialAnterior > 0 Then
            strSQL = "UPDATE SECUENCIAL SET"
            strSQL = strSQL & " IdMember=0"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " Secuencial=" & lSecuencialAnterior
            
            Set adocmd = New ADODB.Command
            adocmd.ActiveConnection = Conn
            adocmd.CommandType = adCmdText
            adocmd.CommandText = strSQL
            adocmd.Execute
    
            Set adocmd = Nothing
            
        End If
        
    End If
    
    
    If Not (lSecuencialPase = 0 And lIdTipoPase = 0 And sCambioCredencial = "1") Then
        'Inserta
        #If SqlServer_ Then
            strSQL = "INSERT INTO CREDENCIALES ("
            strSQL = strSQL & " IdMember,"
            strSQL = strSQL & " IdTipoCredencial,"
            strSQL = strSQL & " FechaAlta,"
            strSQL = strSQL & " IniciaVigencia,"
            strSQL = strSQL & " TerminaVigencia,"
            strSQL = strSQL & " Secuencial,"
            strSQL = strSQL & " FotoFile,"
            strSQL = strSQL & " Nombre,"
            strSQL = strSQL & " Numero,"
            strSQL = strSQL & " SecuencialAnterior)"
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex) & ","
            strSQL = strSQL & iIdTipoPase & ","
            strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "',"
            strSQL = strSQL & "'" & Format(Me.dtpFechaIni.Value, "yyyymmdd") & "',"
            strSQL = strSQL & "'" & Format(Me.dtpFechaFin.Value, "yyyymmdd") & "',"
            strSQL = strSQL & lSecuencialPase & ","
            strSQL = strSQL & lNumeroFoto & ","
            strSQL = strSQL & "'" & Trim(UCase(Me.txtNombre.Text)) & "',"
            strSQL = strSQL & Trim(Me.txtNoPase.Text) & ","
            strSQL = strSQL & lSecuencialAnterior & ")"
        #Else
            strSQL = "INSERT INTO CREDENCIALES ("
            strSQL = strSQL & " IdMember,"
            strSQL = strSQL & " IdTipoCredencial,"
            strSQL = strSQL & " FechaAlta,"
            strSQL = strSQL & " IniciaVigencia,"
            strSQL = strSQL & " TerminaVigencia,"
            strSQL = strSQL & " Secuencial,"
            strSQL = strSQL & " FotoFile,"
            strSQL = strSQL & " Nombre,"
            strSQL = strSQL & " Numero,"
            strSQL = strSQL & " SecuencialAnterior)"
            strSQL = strSQL & " VALUES ("
            strSQL = strSQL & Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex) & ","
            strSQL = strSQL & iIdTipoPase & ","
            strSQL = strSQL & "#" & Format(Date, "mm/dd/yyyy") & "#,"
            strSQL = strSQL & "#" & Format(Me.dtpFechaIni.Value, "mm/dd/yyyy") & "#,"
            strSQL = strSQL & "#" & Format(Me.dtpFechaFin.Value, "mm/dd/yyyy") & "#,"
            strSQL = strSQL & lSecuencialPase & ","
            strSQL = strSQL & lNumeroFoto & ","
            strSQL = strSQL & "'" & Trim(UCase(Me.txtNombre.Text)) & "',"
            strSQL = strSQL & Trim(Me.txtNoPase.Text) & ","
            strSQL = strSQL & lSecuencialAnterior & ")"
        #End If

        Set adocmd = New ADODB.Command
        adocmd.ActiveConnection = Conn
        adocmd.CommandType = adCmdText
        adocmd.CommandText = strSQL
        adocmd.Execute
    
        Set adocmd = Nothing
    
    
        'Si tiene vigencia, lo activa
        If (Me.ssCmbTipoPase.Columns("TieneCredencial").Value = 1) And Me.dtpFechaIni.Value <= Date Then
            #If SqlServer_ Then
                ActivaCredSQL 1, lSecuencialPase, 1, Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), True, False
            #Else
                ActivaCred 1, lSecuencialPase, 1, Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), True, False
            #End If
        End If
    
    
    
        'En el caso de credenciales o reposiciones desactiva la anterior
        'o desactiva la credencial meintras el pase este vigente
        If lSecuencialAnterior > 0 Then
            #If SqlServer_ Then
                ActivaCredSQL 1, lSecuencialAnterior, 1, Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), False, False
            #Else
                ActivaCred 1, lSecuencialAnterior, 1, Me.cmbUsuario.ItemData(Me.cmbUsuario.ListIndex), False, False
            #End If
        End If
    End If
    
End Sub

Private Sub LimpiaControles()

    
    
    


    Me.txtNoPase.Text = ""
    Me.dtpFechaIni.Value = Date
    Me.dtpFechaFin.Value = Date
    Me.txtNombre.Text = ""

End Sub

Private Sub LlenaComboTipoCred()
    
    #If SqlServer_ Then
        strSQL = "SELECT Tipo_Credencial_Descripcion, IdTipoCredencial, CONVERT(int, ISNULL(TieneCredencial,0)) AS TieneCredencial, DiasMaximo, CONVERT(int, ISNULL(RequiereNombre,0)) AS RequiereNombre, VecesSinCosto, idConcepto"
        strSQL = strSQL & " FROM TIPO_CREDENCIAL"
        strSQL = strSQL & " ORDER BY IdTipoCredencial"
    #Else
        strSQL = "SELECT Tipo_Credencial_Descripcion, IdTipoCredencial, iif(TieneCredencial, 1, 0), DiasMaximo, iif(RequiereNombre, 1, 0), VecesSinCosto, idConcepto"
        strSQL = strSQL & " FROM TIPO_CREDENCIAL"
        strSQL = strSQL & " ORDER BY IdTipoCredencial"
    #End If
    
    LlenaSsCombo Me.ssCmbTipoPase, Conn, strSQL, 7
    
    
    If Me.ssCmbTipoPase.Rows > 0 Then
        Me.ssCmbTipoPase.Bookmark = Me.ssCmbTipoPase.AddItemBookmark(1)
        Me.ssCmbTipoPase.Text = Me.ssCmbTipoPase.Columns("Credencial").Value
    End If
    
    
    
End Sub



Private Sub optSoloMes_Click()
    LlenaGridCred
End Sub

Private Sub optTodos_Click()
    LlenaGridCred
End Sub





Private Sub ssCmbTipoPase_Click()
    LlenaGridCred
End Sub



Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNoPase_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = 0
            SendKeys vbTab
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub ActivaBotones(bValor As Boolean)
        
    Me.cmbUsuario.Enabled = bValor
    Me.ssCmbTipoPase.Enabled = bValor
    Me.optSoloMes.Enabled = bValor
    Me.optTodos.Enabled = bValor
    
    Me.cmdAgrega.Enabled = bValor
    Me.cmdElimina.Enabled = bValor
    Me.cmdBloquea.Enabled = bValor
    
    
End Sub
