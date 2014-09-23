VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCupones 
   Caption         =   "Administración de cupones"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   7920
      TabIndex        =   23
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancela"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   22
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   6600
      TabIndex        =   20
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   5280
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   3720
      TabIndex        =   14
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   1920
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdAsignar 
      Caption         =   "Asignar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   7920
      TabIndex        =   9
      Top             =   1200
      Width           =   1090
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   6600
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   5
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdBuscaCupon 
      Caption         =   "Busca"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   480
      Width           =   7095
   End
   Begin VB.TextBox txtControl 
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdbcmbInstructor 
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   2880
      Width           =   4455
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
      Columns.Count   =   2
      Columns(0).Width=   8678
      Columns(0).Caption=   "Nombre"
      Columns(0).Name =   "Nombre"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "IdInstructor"
      Columns(1).Name =   "IdInstructor"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   7858
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label lblControl 
      Caption         =   "Status"
      Height          =   255
      Index           =   9
      Left            =   6600
      TabIndex        =   21
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblControl 
      Caption         =   "# Recibo"
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   19
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblControl 
      Caption         =   "Instructor Asignado"
      Height          =   255
      Index           =   7
      Left            =   1920
      TabIndex        =   17
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Label lblControl 
      Caption         =   "Fecha Pago"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   15
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblControl 
      Caption         =   "Fecha Asignación"
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblControl 
      Caption         =   "Instructor Solicitado"
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblControl 
      Caption         =   "#Cupon"
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblControl 
      Caption         =   "Fecha Venta"
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblControl 
      Caption         =   "Concepto"
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblControl 
      Caption         =   "# Folio Cupón"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmCupones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAsignar_Click()
    Dim adocmdCupones As ADODB.Command
    
    If Not ChecaSeguridad(Me.Name, Me.cmdAsignar.Name) Then
        Exit Sub
    End If
    
    If Me.ssdbcmbInstructor.Text = vbNullString Then
        MsgBox "Seleccionar un instructor", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    If Val(Me.txtControl(9).Text) = 0 Then
        MsgBox "El importe a pagar no puede ser cero.", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    If Me.txtControl(7).Text = CDate(0) Then
        MsgBox "Fecha de pago inválida", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    
    
    
    
    strSQL = "UPDATE"
    strSQL = strSQL & " CUPONES SET"
    #If SqlServer_ Then
        strSQL = strSQL & " FechaAplicacion=" & "'" & Format(Date, "yyyymmdd") & "',"
        strSQL = strSQL & " FechaPago=" & "'" & Format(CDate(Me.txtControl(7).Text), "yyyymmdd") & "',"
    #Else
        strSQL = strSQL & " FechaAplicacion=" & "#" & Format(Date, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & " FechaPago=" & "#" & Format(CDate(Me.txtControl(7).Text), "mm/dd/yyyy") & "#,"
    #End If
    strSQL = strSQL & " IdInstructorAplicacion=" & Me.ssdbcmbInstructor.Columns("IdInstructor").Value & ","
    strSQL = strSQL & " UsuarioAsigna=" & "'" & sDB_User & "',"
    strSQL = strSQL & " ImporteAPagar=" & Trim(Me.txtControl(9).Text)
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " FolioCupon=" & Trim(Me.txtControl(0).Text)
    
    Set adocmdCupones = New ADODB.Command
    adocmdCupones.ActiveConnection = Conn
    adocmdCupones.CommandType = adCmdText
    adocmdCupones.CommandText = strSQL
    adocmdCupones.Execute
    
    
    Set adocmdCupones = Nothing
    
    
    MsgBox "Cupón " & Me.txtControl(0).Text & " asignado", vbInformation, "Ok"
    
    Me.cmdBuscaCupon.Enabled = True
    Me.cmdAsignar.Enabled = False
    Me.cmdAsignar.Enabled = False
    
    
    LimpiaControles
    
    Me.txtControl(0).Enabled = True
    Me.txtControl(0).SetFocus
    
End Sub

Private Sub cmdBuscaCupon_Click()
    
    Dim sStatus As String
    Dim lI As Long
    Dim vBm As Variant
    Dim sTipoDoc As String
    
    If Me.txtControl(0).Text = "" Then
        MsgBox "Se requiere # de cupón!", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    Me.txtControl(0).Enabled = False
    Me.cmdBuscaCupon.Enabled = False
    
    
    sTipoDoc = ObtieneTipoDocCupon(CLng(Me.txtControl(0).Text))
    
    If sTipoDoc = vbNullString Then
        MsgBox "Número de cupón inexistente!", vbExclamation, "Verifique"
        Me.txtControl(0).SelStart = 0
        Me.txtControl(0).SelLength = Len(Me.txtControl(0).Text)
        Me.txtControl(0).Enabled = True
        Me.cmdBuscaCupon.Enabled = True
        Me.txtControl(0).SetFocus
        Exit Sub
    End If
    
    Dim adoRcsCupon As ADODB.Recordset
    
    #If SqlServer_ Then
        strSQL = "SELECT"
        strSQL = strSQL & " CUPONES.FolioCupon,"
        strSQL = strSQL & " CUPONES.NumeroDocumento,"
        strSQL = strSQL & " CUPONES.IdConcepto,"
        strSQL = strSQL & " CUPONES.FechaAlta,"
        strSQL = strSQL & " CUPONES.FechaVigencia,"
        strSQL = strSQL & " CUPONES.FechaAplicacion,"
        strSQL = strSQL & " CUPONES.FechaPago,"
        strSQL = strSQL & " CUPONES.IdInstructor,"
        strSQL = strSQL & " CUPONES.IdInstructorAplicacion,"
        strSQL = strSQL & " CUPONES.NumeroCupon,"
        strSQL = strSQL & " CUPONES.TotalCupones,"
        strSQL = strSQL & " CONCEPTO_INGRESOS.Descripcion,"
        strSQL = strSQL & " CONCEPTO_INGRESOS.ImporteaPagar,"
        strSQL = strSQL & " INSTRUCTORES.Nombre + ' ' + INSTRUCTORES.Apellido_Paterno + ' ' +  INSTRUCTORES.Apellido_Materno As Instructor,"
        strSQL = strSQL & " INS.Nombre + ' ' + INS.Apellido_Paterno + ' ' +  INS.Apellido_Materno As InstructorAplica,"
        If sTipoDoc = "R" Then
            strSQL = strSQL & " RECIBOS.Cancelada"
        Else
            strSQL = strSQL & " FACTURAS.Cancelada"
        End If
        strSQL = strSQL & " FROM CUPONES"
        If sTipoDoc = "R" Then
            strSQL = strSQL & " INNER JOIN RECIBOS ON CUPONES.NumeroDocumento=RECIBOS.NumeroRecibo"
        Else
            strSQL = strSQL & " INNER JOIN FACTURAS ON CUPONES.NumeroDocumento=FACTURAS.NumeroFactura"
        End If
        strSQL = strSQL & " LEFT JOIN CONCEPTO_INGRESOS ON CUPONES.IdConcepto=CONCEPTO_INGRESOS.IdConcepto"
        strSQL = strSQL & " LEFT JOIN INSTRUCTORES ON CUPONES.IdInstructor=INSTRUCTORES.IdInstructor"
        strSQL = strSQL & " LEFT JOIN INSTRUCTORES INS ON CUPONES.IdInstructorAplicacion=INS.IdInstructor"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " FolioCupon=" & Trim(Me.txtControl(0).Text)
    #Else
        strSQL = "SELECT"
        strSQL = strSQL & " CUPONES.FolioCupon,"
        strSQL = strSQL & " CUPONES.NumeroDocumento,"
        strSQL = strSQL & " CUPONES.IdConcepto,"
        strSQL = strSQL & " CUPONES.FechaAlta,"
        strSQL = strSQL & " CUPONES.FechaVigencia,"
        strSQL = strSQL & " CUPONES.FechaAplicacion,"
        strSQL = strSQL & " CUPONES.FechaPago,"
        strSQL = strSQL & " CUPONES.IdInstructor,"
        strSQL = strSQL & " CUPONES.IdInstructorAplicacion,"
        strSQL = strSQL & " CUPONES.NumeroCupon,"
        strSQL = strSQL & " CUPONES.TotalCupones,"
        strSQL = strSQL & " CONCEPTO_INGRESOS.Descripcion,"
        strSQL = strSQL & " CONCEPTO_INGRESOS.ImporteaPagar,"
        strSQL = strSQL & " INSTRUCTORES.Nombre & ' ' & INSTRUCTORES.Apellido_Paterno & ' ' &  INSTRUCTORES.Apellido_Materno As Instructor,"
        strSQL = strSQL & " INS.Nombre & ' ' & INS.Apellido_Paterno & ' ' &  INS.Apellido_Materno As InstructorAplica,"
        If sTipoDoc = "R" Then
            strSQL = strSQL & " RECIBOS.Cancelada"
        Else
            strSQL = strSQL & " FACTURAS.Cancelada"
        End If
        strSQL = strSQL & " FROM (((CUPONES"
        If sTipoDoc = "R" Then
            strSQL = strSQL & " INNER JOIN RECIBOS ON CUPONES.NumeroDocumento=RECIBOS.NumeroRecibo)"
        Else
            strSQL = strSQL & " INNER JOIN FACTURAS ON CUPONES.NumeroDocumento=FACTURAS.NumeroFactura)"
        End If
        strSQL = strSQL & " LEFT JOIN CONCEPTO_INGRESOS ON CUPONES.IdConcepto=CONCEPTO_INGRESOS.IdConcepto)"
        strSQL = strSQL & " LEFT JOIN INSTRUCTORES ON CUPONES.IdInstructor=INSTRUCTORES.IdInstructor)"
        strSQL = strSQL & " LEFT JOIN INSTRUCTORES INS ON CUPONES.IdInstructorAplicacion=INS.IdInstructor"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " FolioCupon=" & Trim(Me.txtControl(0).Text)
    #End If
    
    Set adoRcsCupon = New ADODB.Recordset
    adoRcsCupon.ActiveConnection = Conn
    adoRcsCupon.CursorLocation = adUseServer
    adoRcsCupon.CursorType = adOpenForwardOnly
    adoRcsCupon.LockType = adLockReadOnly
    adoRcsCupon.Open strSQL
    
    
    If adoRcsCupon.EOF Then
        adoRcsCupon.Close
        Set adoRcsCupon = Nothing
        MsgBox "Número de cupón inexistente!", vbExclamation, "Verifique"
        Me.txtControl(0).SelStart = 0
        Me.txtControl(0).SelLength = Len(Me.txtControl(0).Text)
        Me.txtControl(0).Enabled = True
        Me.cmdBuscaCupon.Enabled = True
        Me.txtControl(0).SetFocus
        Exit Sub
    End If
    
    sStatus = "SIN ASIGNAR"
    
    If adoRcsCupon!Cancelada = -1 Then
        sStatus = "CANCELADO POR RECIBO/FACTURA"
    ElseIf adoRcsCupon!FechaAplicacion > CDate("31/12/1999") Then
        sStatus = "ASIGNADO"
    ElseIf adoRcsCupon!FechaVigencia < Date Then
        sStatus = "VENCIDO"
    End If
    
    
    
    
    
    
    Me.txtControl(1).Text = "(" & adoRcsCupon!IdConcepto & ") " & IIf(IsNull(adoRcsCupon!Descripcion), "", adoRcsCupon!Descripcion)
    Me.txtControl(2).Text = "(" & adoRcsCupon!IdInstructor & ") " & IIf(IsNull(adoRcsCupon!Instructor), "", adoRcsCupon!Instructor)
    Me.txtControl(3).Text = Format(adoRcsCupon!FechaAlta, "dd/mmm/yy")
    Me.txtControl(4).Text = adoRcsCupon!NumeroCupon & "/" & adoRcsCupon!TotalCupones
    Me.txtControl(5).Text = adoRcsCupon!NumeroDocumento
    Me.txtControl(6).Text = IIf(IsNull(adoRcsCupon!FechaAplicacion), Format(Date, "dd/mmm/yy"), Format(adoRcsCupon!FechaAplicacion, "dd/mmm/yy"))
    'Me.txtControl(7).Text = IIf(IsNull(adoRcsCupon!Fechapago), Format(DateSerial(Year(Date), Month(Date), IIf(Day(Date) <= 15, 15, 30)), "dd/mmm/yy"), Format(adoRcsCupon!Fechapago, "dd/mmm/yy"))
    '15/09/2007
    Me.txtControl(7).Text = IIf(IsNull(adoRcsCupon!Fechapago), ObtieneFechaPago(), Format(adoRcsCupon!Fechapago, "dd/mmm/yy"))
    Me.txtControl(8).Text = sStatus
    Me.txtControl(9).Text = IIf(IsNull(adoRcsCupon!ImporteaPagar), 0, adoRcsCupon!ImporteaPagar)
    
    
    If adoRcsCupon!IdInstructor > 0 Then
        For lI = 0 To Me.ssdbcmbInstructor.Rows - 1
            vBm = Me.ssdbcmbInstructor.AddItemBookmark(lI)
            If Val(Me.ssdbcmbInstructor.Columns(1).CellValue(vBm)) = adoRcsCupon!IdInstructor Then
                Me.ssdbcmbInstructor.Bookmark = Me.ssdbcmbInstructor.AddItemBookmark(lI)
                Me.ssdbcmbInstructor.Text = Me.ssdbcmbInstructor.Columns(0).CellText(vBm)
                Exit For
            End If
        Next
    End If
    
    
    Me.cmdCancela.Enabled = True
    Me.cmdAsignar.Enabled = True
    
    
    If sStatus = "CANCELADO POR RECIBO/FACTURA" Then
        'Me.ssdbcmbInstructor.Text = adoRcsCupon!InstructorAplica
        Me.ssdbcmbInstructor.Enabled = False
        Me.cmdAsignar.Enabled = False
    End If
    
    
    If sStatus = "ASIGNADO" Then
        Me.ssdbcmbInstructor.Text = adoRcsCupon!InstructorAplica
        Me.ssdbcmbInstructor.Enabled = False
        Me.cmdAsignar.Enabled = False
    End If

    adoRcsCupon.Close
    Set adoRcsCupon = Nothing
    
    
    
End Sub

Private Sub cmdCancela_Click()
    
    
    LimpiaControles
    
    Me.cmdAsignar.Enabled = False
    Me.cmdCancela.Enabled = False
    Me.cmdBuscaCupon.Enabled = True
    
    
    Me.txtControl(0).Enabled = True
    Me.txtControl(0).SetFocus
    
End Sub

Private Sub Form_Activate()
    Me.txtControl(0).SetFocus
End Sub

Private Sub Form_Load()
    'LlenaGrid
    Carga_Instructores
    CentraForma MDIPrincipal, Me
    
End Sub

Private Sub Carga_Instructores()
    Dim adorcsInstructor As ADODB.Recordset
    
    #If SqlServer_ Then
        strSQL = "SELECT IdInstructor, Nombre, Apellido_Paterno, Apellido_Materno"
        strSQL = strSQL & " FROM INSTRUCTORES"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " Status='A'"
        strSQL = strSQL & " ORDER BY Nombre + ' ' + Apellido_Paterno + ' ' + Apellido_Materno"
    #Else
        strSQL = "SELECT IdInstructor, Nombre, Apellido_Paterno, Apellido_Materno"
        strSQL = strSQL & " FROM INSTRUCTORES"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " Status='A'"
        strSQL = strSQL & " ORDER BY Nombre & ' ' & Apellido_Paterno & ' ' & Apellido_Materno"
    #End If
    
    Set adorcsInstructor = New ADODB.Recordset
    adorcsInstructor.CursorLocation = adUseServer
    
    adorcsInstructor.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Me.ssdbcmbInstructor.RemoveAll
    
    Do While Not adorcsInstructor.EOF
        Me.ssdbcmbInstructor.AddItem adorcsInstructor!Nombre & " " & adorcsInstructor!apellido_paterno & " " & adorcsInstructor!apellido_materno & vbTab & adorcsInstructor!IdInstructor
        adorcsInstructor.MoveNext
    Loop
    
    adorcsInstructor.Close
    
    Set adorcsInstructor = Nothing
    
End Sub

Private Sub LimpiaControles()
    Dim lI As Long
    
    For lI = 0 To Me.txtControl.Count - 1
        Me.txtControl(lI).Text = vbNullString
    Next
    
    Me.ssdbcmbInstructor.Text = vbNullString
    Me.ssdbcmbInstructor.Enabled = True
End Sub

Private Function ObtieneFechaPago() As Date
    
    Dim adorcsFechaPago As ADODB.Recordset
    
    ObtieneFechaPago = CDate(0)
    
    
    On Error GoTo Error_Catch
    
    #If SqlServer_ Then
        strSQL = "SELECT FechaPago"
        strSQL = strSQL & " FROM NOMINA_CALENDARIO"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " FechaInicial <= getdate()"
        strSQL = strSQL & " AND FechaFinal >= getdate()"
    #Else
        strSQL = "SELECT FechaPago"
        strSQL = strSQL & " FROM NOMINA_CALENDARIO"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " FechaInicial <= Now()"
        strSQL = strSQL & " AND FechaFinal >= Now()"
    #End If
    
    Set adorcsFechaPago = New ADODB.Recordset
    
    
    adorcsFechaPago.CursorLocation = adUseServer
    
    adorcsFechaPago.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcsFechaPago.EOF Then
        ObtieneFechaPago = adorcsFechaPago!Fechapago
    End If
    
    On Error GoTo 0
    
    Exit Function
    
Error_Catch:

    MsgError
    

End Function

Private Function ObtieneTipoDocCupon(lFolioCupon As Long) As String
    
    Dim adorcs As ADODB.Recordset
    
    
    ObtieneTipoDocCupon = vbNullString
    
    strSQL = "SELECT TipoDocumento"
    strSQL = strSQL & " FROM CUPONES"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & "((FolioCupon)=" & lFolioCupon & ")"
    strSQL = strSQL & ")"
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        ObtieneTipoDocCupon = adorcs!TipoDocumento
    End If
    
    adorcs.Close
    
    Set adorcs = Nothing
    
    
End Function
