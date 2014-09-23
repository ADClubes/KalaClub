VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmOperLockers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lockers"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   4560
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Operacion"
      Height          =   975
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdOperacion 
         Caption         =   "Proceder"
         Height          =   495
         Left            =   4200
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Index           =   8
         Left            =   2760
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbOPeracion 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2415
         DataFieldList   =   "Column 0"
         AllowInput      =   0   'False
         AllowNull       =   0   'False
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
         Columns(0).Width=   4366
         Columns(0).Caption=   "Operacion"
         Columns(0).Name =   "Operacion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos"
      Enabled         =   0   'False
      Height          =   3855
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   10
         Left            =   5040
         TabIndex        =   25
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   9
         Left            =   5040
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkUsoDiario 
         Caption         =   "Uso diario"
         Height          =   255
         Left            =   4200
         TabIndex        =   23
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   7
         Left            =   1560
         TabIndex        =   19
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   6
         Left            =   1560
         TabIndex        =   18
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   1560
         TabIndex        =   17
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   16
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   15
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label lblCtrl 
         Alignment       =   1  'Right Justify
         Caption         =   "Documento"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Alignment       =   1  'Right Justify
         Caption         =   "Importe pagado"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Alignment       =   1  'Right Justify
         Caption         =   "Pagado hasta"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Alignment       =   1  'Right Justify
         Caption         =   "Pagado desde"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblCtrl 
         Alignment       =   1  'Right Justify
         Caption         =   "Ubicacion:"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   6
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblCtrl 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo:"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblCtrl 
         Alignment       =   1  'Right Justify
         Caption         =   "Asignado a:"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Locker"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton cmdBusca 
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmOperLockers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBusca_Click()
    Dim adorcsLocker As ADODB.Recordset
    
    
    
    
    If Me.txtNombre(0).Text = vbNullString Then
        Me.txtNombre(0).SetFocus
        Exit Sub
    End If
    
    Me.txtNombre(0).Text = UCase(Me.txtNombre(0).Text)
    
    #If SqlServer_ Then
        strSQL = "SELECT RENTABLES.IdTipoRentable, RENTABLES.Numero, RENTABLES.Sexo, RENTABLES.Ubicacion, RENTABLES.IdUsuario, RENTABLES.FechaPago, RENTABLES.Propiedad, RENTABLES.Observaciones, RENTABLES.Area, RENTABLES.FechaInicio, RENTABLES.ImportePagado, RENTABLES.Documento, TIPO_RENTABLES.Descripcion, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.NoFamilia"
        strSQL = strSQL & " FROM RENTABLES INNER JOIN TIPO_RENTABLES ON RENTABLES.IdTipoRentable = TIPO_RENTABLES.IdTipoRentable LEFT JOIN USUARIOS_CLUB ON RENTABLES.IdUsuario = USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " LTRIM(RTRIM(RENTABLES.Numero))='" & Trim(Me.txtNombre(0).Text) & "'"
    #Else
        strSQL = "SELECT RENTABLES.IdTipoRentable, RENTABLES.Numero, RENTABLES.Sexo, RENTABLES.Ubicacion, RENTABLES.IdUsuario, RENTABLES.FechaPago, RENTABLES.Propiedad, RENTABLES.Observaciones, RENTABLES.Area, RENTABLES.FechaInicio, RENTABLES.ImportePagado, RENTABLES.Documento, TIPO_RENTABLES.Descripcion, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.NoFamilia"
        strSQL = strSQL & " FROM (RENTABLES INNER JOIN TIPO_RENTABLES ON RENTABLES.IdTipoRentable = TIPO_RENTABLES.IdTipoRentable) LEFT JOIN USUARIOS_CLUB ON RENTABLES.IdUsuario = USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " (Trim(RENTABLES.Numero)='" & Trim(Me.txtNombre(0).Text) & "')"
        strSQL = strSQL & ")"
    #End If
    
    Set adorcsLocker = New ADODB.Recordset
    adorcsLocker.CursorLocation = adUseServer
    
    adorcsLocker.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsLocker.EOF Then
        If IIf(IsNull(adorcsLocker!idusuario), 0, adorcsLocker!idusuario) <> 0 Then
            Me.txtNombre(1).Text = "(" & IIf(IsNull(adorcsLocker!NoFamilia), 0, adorcsLocker!NoFamilia) & ") " & IIf(IsNull(adorcsLocker!A_Paterno), vbNullString, adorcsLocker!A_Paterno) & " " & IIf(IsNull(adorcsLocker!A_Materno), vbNullString, adorcsLocker!A_Materno) & " " & IIf(IsNull(adorcsLocker!Nombre), vbNullString, adorcsLocker!Nombre)
        Else
            Me.txtNombre(1).Text = "SIN ASIGNAR"
        End If
        Me.txtNombre(2).Text = IIf(IsNull(adorcsLocker!Descripcion), vbNullString, adorcsLocker!Descripcion)
        Me.txtNombre(3).Text = IIf(IsNull(adorcsLocker!Ubicacion), vbNullString, adorcsLocker!Ubicacion)
        Me.txtNombre(4).Text = IIf(IsNull(adorcsLocker!FechaInicio), vbNullString, adorcsLocker!FechaInicio)
        Me.txtNombre(5).Text = IIf(IsNull(adorcsLocker!Fechapago), vbNullString, adorcsLocker!Fechapago)
        Me.txtNombre(6).Text = IIf(IsNull(adorcsLocker!ImportePagado), 0, adorcsLocker!ImportePagado)
        Me.txtNombre(7).Text = IIf(IsNull(adorcsLocker!Documento), vbNullString, adorcsLocker!Documento)
        Me.txtNombre(9).Text = IIf(IsNull(adorcsLocker!idusuario), 0, adorcsLocker!idusuario)
        Me.txtNombre(10).Text = IIf(IsNull(adorcsLocker!idtiporentable), 0, adorcsLocker!idtiporentable)
        
        Me.chkUsoDiario.Value = IIf(adorcsLocker!Propiedad, 1, 0)
        
        Me.Frame2.Visible = True
        
        'Si no es de uso diario y no está asignado
        If Me.chkUsoDiario.Value = 0 And adorcsLocker!idusuario > 0 Then
            Me.Frame3.Visible = True
        End If
        
        Me.txtNombre(0).Enabled = False
        
        Me.cmdCancelar.Visible = True
        Me.cmdBusca.Enabled = False
        
    Else
        Me.txtNombre(0).SelStart = 0
        Me.txtNombre(0).SelLength = Len(Me.txtNombre(0).Text)
        Me.txtNombre(0).SetFocus
        MsgBox "¡Este Locker no existe!", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    adorcsLocker.Close
    Set adorcsLocker = Nothing

    
    
    
    
End Sub

Private Sub cmdCancelar_Click()
    
    Me.Frame2.Visible = False
    Me.Frame3.Visible = False
    
    Me.txtNombre(0).Enabled = True
    
    Me.txtNombre(0).SelStart = 0
    Me.txtNombre(0).SelLength = Len(Me.txtNombre(0).Text)
    
    Me.txtNombre(0).SetFocus
    
    Me.cmdBusca.Enabled = Enabled
    Me.cmdCancelar.Visible = False
    
    Me.sscmbOPeracion.Bookmark = 0
    
    Me.sscmbOPeracion.Text = vbNullString
    
    Me.txtNombre(8).Text = vbNullString
    
    
End Sub

Private Sub cmdOperacion_Click()
    
    Dim lCommit As Long
    Dim bSinError As Boolean
    
    Dim lTipoRentable_Nuevo As Long
    
    Dim sMensaje As String
    
    
    
    bSinError = True
    
    Me.txtNombre(8).Text = UCase(Me.txtNombre(8).Text)
    
    If Me.sscmbOPeracion.Text = vbNullString Then
        MsgBox "Seleccione una operación", vbExclamation, "Verifique"
        Me.sscmbOPeracion.SetFocus
        Exit Sub
    End If
    
    
    If Me.sscmbOPeracion.row = 1 And Me.txtNombre(8).Text = vbNullString Then
        MsgBox "Indique el nuevo # de locker", vbExclamation, "Verifique"
        Me.txtNombre(8).SetFocus
        Exit Sub
    End If
    
    'Si se va a reemplazar el locker, valida que exista
    sMensaje = vbNullString
    If Me.sscmbOPeracion.row = 1 Then
    
        
        
        Select Case ValidaLocker(Me.txtNombre(8).Text, lTipoRentable_Nuevo)
            Case 0 'Existe el locker y esta disponible
            Case 1 'Existe esta asignado a otro usuario y tiene vigencia
                sMensaje = "El locker está asignado a otro usuario y con pago vigente"
            Case 2 'Existe esta asignado a otro usuario, pero ya esta vencido
                sMensaje = "El locker está asignado a otro usuario y pago pendiente"
            Case 3 'Existe esta asignado a otro usuario, pero ya esta vencido
                sMensaje = "El locker es de propiedad!"
                MsgBox sMensaje, vbExclamation, "Verifique"
                Me.txtNombre(8).SetFocus
                Exit Sub
            Case 4 'No existe
                MsgBox "¡El locker capturado no existe!", vbExclamation, "Verifique"
                Me.txtNombre(8).SetFocus
                Exit Sub
        End Select
        
        If lTipoRentable_Nuevo <> Val(Me.txtNombre(10)) Then
            MsgBox "Los Lockers no son del mismo tipo", vbExclamation, "verifique"
            Exit Sub
        End If
        
        If sMensaje <> vbNullString Then
            sMensaje = sMensaje + vbCrLf + "¿Continuar?"
            If MsgBox(sMensaje, vbExclamation + vbOKCancel, "Confirme") = vbCancel Then
                Exit Sub
            End If
        End If
        
    End If
    
    
    
    
    Conn.Errors.Clear
    lCommit = Conn.BeginTrans
    
    Select Case Me.sscmbOPeracion.row
        Case 0
            'Para dejar libre un locker
            If Ejecuta(Me.txtNombre(0).Text, 0) Then
                Conn.CommitTrans
            Else
                If lCommit > 0 Then
                    Conn.RollbackTrans
                    bSinError = False
                End If
            End If
           
           
        Case 1
            'Para cambiar
            If Ejecuta(Me.txtNombre(8).Text, 1) Then
                If Ejecuta(Me.txtNombre(0).Text, 0) Then
                    Conn.CommitTrans
                Else
                    If lCommit > 0 Then
                        Conn.RollbackTrans
                        bSinError = False
                    End If
                End If
            Else
                If lCommit > 0 Then
                    Conn.RollbackTrans
                    bSinError = False
                End If
            End If
        
    End Select
    
    
    If bSinError Then
        MsgBox "Operación efectuada", vbInformation, "Ok"
        Me.cmdCancelar.Value = True
    Else
        MsgBox "Ocurrio un error", vbExclamation, "Error"
    End If
    
    
End Sub

Private Sub Form_Load()
    
    
    CentraForma MDIPrincipal, Me
    
    
    Me.sscmbOPeracion.AddItem "Dejar disponible"
    Me.sscmbOPeracion.AddItem "Cambiar por otro locker"
    
    
    
    
    
End Sub

Private Sub sscmbOPeracion_Click()
    Debug.Print "Hola"
End Sub


Private Function Ejecuta(sNumeroLocker As String, iModo As Integer) As Boolean

    Dim adocmd As ADODB.Command
    
    Ejecuta = True
    
    #If SqlServer_ Then
        If iModo = 0 Then
            strSQL = "UPDATE RENTABLES SET"
            strSQL = strSQL & " IdUsuario = 0" & ","
            strSQL = strSQL & " FechaPago = Null" & ","
            strSQL = strSQL & " FechaInicio = Null" & ","
            strSQL = strSQL & " ImportePagado = 0" & ","
            strSQL = strSQL & " Documento =  Null"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " LTRIM(RTrim(Numero))='" & Trim(sNumeroLocker) & "'"
        Else
            strSQL = "UPDATE RENTABLES SET"
            strSQL = strSQL & " IdUsuario = " & Me.txtNombre(9).Text & ","
            strSQL = strSQL & " FechaPago ='" & Format(Me.txtNombre(5).Text, "yyyymmdd") & "',"
            strSQL = strSQL & " FechaInicio = " & IIf(Me.txtNombre(4).Text = vbNullString, "Null", "'" & Format(Me.txtNombre(4).Text, "yyyymmdd") & "'") & ","
            strSQL = strSQL & " ImportePagado = " & Me.txtNombre(6).Text & ","
            strSQL = strSQL & " Documento = '" & Me.txtNombre(7).Text & "'"
            strSQL = strSQL & " WHERE"
            strSQL = strSQL & " RTRIM(LTrim(Numero))='" & Trim(sNumeroLocker) & "'"
        End If
    #Else
        If iModo = 0 Then
            strSQL = "UPDATE RENTABLES SET"
            strSQL = strSQL & " IdUsuario = 0" & ","
            strSQL = strSQL & " FechaPago = Null" & ","
            strSQL = strSQL & " FechaInicio = Null" & ","
            strSQL = strSQL & " ImportePagado = 0" & ","
            strSQL = strSQL & " Documento =  Null"
            strSQL = strSQL & " Where ("
            strSQL = strSQL & " (Trim(Numero)='" & Trim(sNumeroLocker) & "')"
            strSQL = strSQL & ")"
        Else
            strSQL = "UPDATE RENTABLES SET"
            strSQL = strSQL & " IdUsuario = " & Me.txtNombre(9).Text & ","
            strSQL = strSQL & " FechaPago ='" & Me.txtNombre(5).Text & "',"
            strSQL = strSQL & " FechaInicio = " & IIf(Me.txtNombre(4).Text = vbNullString, "Null", "'" & Me.txtNombre(4).Text & "'") & ","
            strSQL = strSQL & " ImportePagado = " & Me.txtNombre(6).Text & ","
            strSQL = strSQL & " Documento = '" & Me.txtNombre(7).Text & "'"
            strSQL = strSQL & " Where ("
            strSQL = strSQL & " (Trim(Numero)='" & Trim(sNumeroLocker) & "')"
            strSQL = strSQL & ")"
        End If
    #End If
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    
    
    On Error GoTo CATCH_ERROR
    adocmd.Execute
    
    Set adocmd = Nothing
    
    
    
    Exit Function
    
CATCH_ERROR:
    
    Ejecuta = False
    
    Set adocmd = Nothing
    

End Function

 Private Function ValidaLocker(sNumeroLocker As String, ByRef lTipoLockerNuevo As Long) As Integer
    Dim adorcs As ADODB.Recordset
    
    #If SqlServer_ Then
        strSQL = "SELECT RENTABLES.IdTipoRentable, RENTABLES.Numero, RENTABLES.Sexo, RENTABLES.Ubicacion, RENTABLES.IdUsuario, RENTABLES.FechaPago, RENTABLES.Propiedad, RENTABLES.Observaciones, RENTABLES.Area, RENTABLES.FechaInicio, RENTABLES.ImportePagado, RENTABLES.Documento, TIPO_RENTABLES.Descripcion, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.NoFamilia"
        strSQL = strSQL & " FROM RENTABLES INNER JOIN TIPO_RENTABLES ON RENTABLES.IdTipoRentable = TIPO_RENTABLES.IdTipoRentable LEFT JOIN USUARIOS_CLUB ON RENTABLES.IdUsuario = USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " RTRIM(LTRIM(RENTABLES.Numero))='" & Trim(sNumeroLocker) & "'"
    #Else
        strSQL = "SELECT RENTABLES.IdTipoRentable, RENTABLES.Numero, RENTABLES.Sexo, RENTABLES.Ubicacion, RENTABLES.IdUsuario, RENTABLES.FechaPago, RENTABLES.Propiedad, RENTABLES.Observaciones, RENTABLES.Area, RENTABLES.FechaInicio, RENTABLES.ImportePagado, RENTABLES.Documento, TIPO_RENTABLES.Descripcion, USUARIOS_CLUB.Nombre, USUARIOS_CLUB.A_Paterno, USUARIOS_CLUB.A_Materno, USUARIOS_CLUB.NoFamilia"
        strSQL = strSQL & " FROM (RENTABLES INNER JOIN TIPO_RENTABLES ON RENTABLES.IdTipoRentable = TIPO_RENTABLES.IdTipoRentable) LEFT JOIN USUARIOS_CLUB ON RENTABLES.IdUsuario = USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " (Trim(RENTABLES.Numero)='" & Trim(sNumeroLocker) & "')"
        strSQL = strSQL & ")"
    #End If
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    ValidaLocker = 0
    
    If adorcs.EOF Then 'No existe
        ValidaLocker = 4
    Else 'Si existe
    
        lTipoLockerNuevo = adorcs!idtiporentable
        
        If adorcs!Propiedad = -1 Then
            ValidaLocker = 3
        Else
            If adorcs!idusuario <> 0 Then 'Esta asignado
                If adorcs!Fechapago >= Date Then 'Esta vigente
                    ValidaLocker = 1
                Else 'No esta vigente
                    ValidaLocker = 2
                End If
            End If
        End If
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
 End Function
