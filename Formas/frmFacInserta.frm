VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmFacInserta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Usuario"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdbcmbConcepto 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   360
      Width           =   1695
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
      Columns.Count   =   10
      Columns(0).Width=   3200
      Columns(0).Caption=   "Clave"
      Columns(0).Name =   "Clave"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   7197
      Columns(1).Caption=   "Descripcion"
      Columns(1).Name =   "Descripcion"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "Importe"
      Columns(2).Name =   "Importe"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "Impuesto1"
      Columns(3).Name =   "Impuesto1"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Visible=   0   'False
      Columns(4).Caption=   "Impuesto2"
      Columns(4).Name =   "Impuesto2"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1376
      Columns(5).Caption=   "FacORec"
      Columns(5).Name =   "FacORec"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Visible=   0   'False
      Columns(6).Caption=   "RequiereUsuario"
      Columns(6).Name =   "RequiereUsuario"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   3200
      Columns(7).Visible=   0   'False
      Columns(7).Caption=   "RequiereInstructor"
      Columns(7).Name =   "RequiereInstructor"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Visible=   0   'False
      Columns(8).Caption=   "EsModificable"
      Columns(8).Name =   "EsModificable"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Unidad"
      Columns(9).Name =   "Unidad"
      Columns(9).DataField=   "Column 9"
      Columns(9).DataType=   8
      Columns(9).FieldLen=   256
      _ExtentX        =   2990
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdbcmbInstructor 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   6255
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
      _ExtentX        =   11033
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssdbcmbUsuario 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   6255
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
      Columns.Count   =   4
      Columns(0).Width=   8414
      Columns(0).Caption=   "Usuario"
      Columns(0).Name =   "Usuario"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   3200
      Columns(1).Caption=   "IdMember"
      Columns(1).Name =   "IdMember"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Visible=   0   'False
      Columns(2).Caption=   "NumeroFamiliar"
      Columns(2).Name =   "NumeroFamiliar"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "IdTipoUsuario"
      Columns(3).Name =   "IdTipoUsuario"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   11033
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Insertar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblFactura 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lblInstructor 
      Caption         =   "Instructor"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Importe"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblConceptoNombre 
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmFacInserta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim dIvaRen As Double
    Dim iRequiereUsuario As Integer
    Dim lidMember As Long
    Dim lNumeroFamiliar As Long
    Dim lIdTipoUsuario As Long
    Dim sNombre As String
    
    Dim iRequiereInstructor As Integer
    Dim lIdInstructor As Long
    
    
    If Me.ssdbcmbConcepto.Text = "" Then
        MsgBox "Seleccione un concepto!", vbExclamation, "Facturacion"
        Exit Sub
    End If
    
    
    iRequiereUsuario = Val(Me.ssdbcmbConcepto.Columns(6).Value)
    
    If iRequiereUsuario = 1 And Me.ssdbcmbUsuario.Text = vbNullString Then
        MsgBox "Seleccione un usuario!", vbExclamation, "Facturacion"
        Exit Sub
    End If
    
    iRequiereInstructor = Val(Me.ssdbcmbConcepto.Columns(7).Value)
    
    If iRequiereInstructor = 1 And Me.ssdbcmbInstructor.Text = vbNullString Then
        MsgBox "Seleccione un instructor!", vbExclamation, "Facturacion"
        Exit Sub
    End If
    
    
    If Val(Me.txtCantidad.Text) <= 0 Then
        MsgBox "La cantidad debe ser mayor a cero!", vbExclamation, "Facturacion"
        Exit Sub
    End If
    
'    If Val(Me.txtImporte.Text) <= 0 Then
'        MsgBox "El importe debe ser mayor a cero!", vbExclamation, "Facturacion"
'        Exit Sub
'    End If
    
    
    dIvaRen = Val(Me.txtCantidad.Text) * Val(Me.txtImporte.Text) - Round((Val(Me.txtCantidad.Text) * Val(Me.txtImporte.Text)) / (1 + (Me.ssdbcmbConcepto.Columns(3).Value / 100)), 2)
    
    If iRequiereUsuario = 1 Then
        sNombre = Me.ssdbcmbUsuario.Columns(0).Value
        lidMember = Val(Me.ssdbcmbUsuario.Columns(1).Value)
        lNumeroFamiliar = Val(Me.ssdbcmbUsuario.Columns(2).Value)
        lIdTipoUsuario = Val(Me.ssdbcmbUsuario.Columns(3).Value)
    Else
        sNombre = ""
        lidMember = Val(Me.Tag)
        lNumeroFamiliar = 1
        lIdTipoUsuario = 0
    End If
    
    If iRequiereInstructor = 1 Then
        lIdInstructor = Me.ssdbcmbInstructor.Columns(1).Value
    Else
        lIdInstructor = 0
    End If
    
    
    '                                   Descripcion concepto                          Nombre            Periodo                              Cantidad                      Importe                                      Int         Desc        Total
    frmFacturacion.ssdbgFactura.AddItem Me.ssdbcmbConcepto.Columns(1).Value & vbTab & sNombre & vbTab & Format(Date, "dd/mm/yyyy") & vbTab & Me.txtCantidad.Text & vbTab & Round(CDbl(Me.txtImporte.Text), 2) & vbTab & 0 & vbTab & 0 & vbTab & Val(Me.txtCantidad.Text) * Val(Me.txtImporte.Text) & vbTab & Me.ssdbcmbConcepto.Columns(0).Value & vbTab & Me.ssdbcmbConcepto.Columns(3).Value / 100 & vbTab & dIvaRen & vbTab & 0 & vbTab & 0 & vbTab & 0 & vbTab & lidMember & vbTab & lNumeroFamiliar & vbTab & 1 & vbTab & lIdTipoUsuario & vbTab & 5 & vbTab & vbNullString & vbTab & Me.ssdbcmbConcepto.Columns("FacORec").Value & vbTab & lIdInstructor & vbTab & Me.ssdbcmbConcepto.Columns("Unidad").Value
    
    Unload Me
    
End Sub

Private Sub Form_Activate()
    If Me.ssdbcmbUsuario.Rows = 0 Then
        Carga_Usuarios
    End If
    Me.ssdbcmbConcepto.SetFocus
End Sub

Private Sub Form_Load()
    Me.Height = 4785
    Me.Width = 6795
    
    CentraForma MDIPrincipal, Me
    
    
    strSQL = "SELECT IdConcepto, Descripcion, Monto, Impuesto1, Impuesto2, FacORec,"
    
    #If SqlServer_ Then
        strSQL = strSQL & " CONVERT(int, ISNULL(RequiereUsuario,0)) AS RequiereUsuario, CONVERT(int, ISNULL(RequiereInstructor,0)) AS RequiereInstructor, CONVERT(int, ISNULL(EsModificable,0)) EsModificable, Unidad"
    #Else
        strSQL = strSQL & " IIf(RequiereUsuario, 1, 0) AS RequiereUsuario, IIf(RequiereInstructor, 1, 0) AS RequiereInstructor, EsModificable, Unidad"
    #End If
    
    strSQL = strSQL & " FROM CONCEPTO_INGRESOS"
    strSQL = strSQL & " WHERE EsPeriodico=0"
    strSQL = strSQL & " ORDER BY IdConcepto"
    
    
    
    LlenaSsCombo Me.ssdbcmbConcepto, Conn, strSQL, 10
    
    'Carga_Conceptos
    
    Carga_Instructores
    
    Me.txtImporte.Text = 0
    Me.txtCantidad.Text = 1
    
End Sub

'Private Sub Carga_Conceptos()
'    Dim adorcsConcepto As ADODB.Recordset
'
'    strSQL = "SELECT IdConcepto, Descripcion, Monto, Impuesto1, Impuesto2, FacORec,"
'    strSQL = strSQL & " RequiereUsuario, RequiereInstructor, EsModificable"
'    strSQL = strSQL & " FROM CONCEPTO_INGRESOS"
'    strSQL = strSQL & " WHERE EsPeriodico=0"
'    strSQL = strSQL & " ORDER BY IdConcepto"
'
'    Set adorcsConcepto = New ADODB.Recordset
'    adorcsConcepto.CursorLocation = adUseServer
'
'    adorcsConcepto.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
'
'
'    Me.ssdbcmbConcepto.RemoveAll
'
'    Do While Not adorcsConcepto.EOF
'
'        Me.ssdbcmbConcepto.AddItem adorcsConcepto!IdConcepto & vbTab & adorcsConcepto!Descripcion & vbTab & adorcsConcepto!Monto & vbTab & adorcsConcepto!Impuesto1 & vbTab & adorcsConcepto!impuesto2 & vbTab & adorcsConcepto!FacORec & vbTab & IIf(adorcsConcepto!RequiereUsuario, 1, 0) & vbTab & IIf(adorcsConcepto!RequiereInstructor, 1, 0)
'
'        adorcsConcepto.MoveNext
'    Loop
'
'    adorcsConcepto.Close
'
'    Set adorcsConcepto = Nothing
'
'End Sub



'Private Sub ssdbcmbConcepto_Change()
'
'    If Not Me.ssdbcmbConcepto.DroppedDown Then
'        Me.ssdbcmbConcepto.DroppedDown = True
'    End If
'
''    If Me.ssdbcmbConcepto.IsItemInList Then
''        Me.lblConceptoNombre = Me.ssdbcmbConcepto.Columns(1).Value
''    Else
''        Me.lblConceptoNombre = ""
''    End If
'End Sub

Private Sub ssdbcmbConcepto_Click()
    Me.lblConceptoNombre = Me.ssdbcmbConcepto.Columns(1).Value
    Me.txtImporte.Text = Me.ssdbcmbConcepto.Columns(2).Value
    
    '27/02/20'06
    If Me.ssdbcmbConcepto.Columns(5).Value = "F" Then
        Me.lblFactura = "GENERA UNA FACTURA"
    Else
        Me.lblFactura = "GENERA UN RECIBO"
    End If
    
    If Val(Me.ssdbcmbConcepto.Columns(6).Value) = 1 Then
        Me.lblUsuario.Visible = True
        Me.ssdbcmbUsuario.Visible = True
    Else
        Me.lblUsuario.Visible = False
        Me.ssdbcmbUsuario.Visible = False
    End If
    
    If Val(Me.ssdbcmbConcepto.Columns(7).Value) = 1 Then
        Me.lblInstructor.Visible = True
        Me.ssdbcmbInstructor.Visible = True
    Else
        Me.lblInstructor.Visible = False
        Me.ssdbcmbInstructor.Visible = False
    End If
    
    'Si el concepto se le puede modificar el precio
    If sDB_NivelUser <> 0 And Me.ssdbcmbConcepto.Columns(8).Value = 0 Then
        Me.txtImporte.Enabled = False
    Else
        Me.txtImporte.Enabled = True
    End If
    
    
End Sub








Private Sub ssdbcmbConcepto_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            KeyAscii = vbTab
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub ssdbcmbConcepto_LostFocus()
    If Not Me.ssdbcmbConcepto.IsItemInList Then
        Me.ssdbcmbConcepto.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 46 'punto decimal
            If InStr(Me.txtCantidad.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8 ' Tecla backspace
            KeyAscii = KeyAscii
        Case 45 'guion signo menos
            If InStr(Me.txtCantidad.Text, "-") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 46 'punto decimal
            If InStr(Me.txtCantidad.Text, ".") Then
                KeyAscii = 0
            Else
                KeyAscii = KeyAscii
            End If
        Case 48 To 57 ' Numeros del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
            MsgBox "Solo admite números", vbInformation, "Productos"
    End Select

End Sub

Private Sub Carga_Usuarios()
    Dim adorcsUsuario As ADODB.Recordset
    
    strSQL = "SELECT IdMember, Nombre, A_Paterno, A_Materno, NumeroFamiliar, IdTipoUsuario"
    strSQL = strSQL & " FROM USUARIOS_CLUB"
    strSQL = strSQL & " WHERE IdTitular=" & Me.Tag
    strSQL = strSQL & " ORDER BY NumeroFamiliar"
    
    Set adorcsUsuario = New ADODB.Recordset
    adorcsUsuario.CursorLocation = adUseServer
    
    adorcsUsuario.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Me.ssdbcmbUsuario.RemoveAll
    
    Do While Not adorcsUsuario.EOF
        Me.ssdbcmbUsuario.AddItem adorcsUsuario!Nombre & " " & adorcsUsuario!A_Paterno & " " & adorcsUsuario!A_Materno & vbTab & adorcsUsuario!Idmember & vbTab & adorcsUsuario!NumeroFamiliar & vbTab & adorcsUsuario!idtipousuario
        adorcsUsuario.MoveNext
    Loop
    
    adorcsUsuario.Close
    
    Set adorcsUsuario = Nothing
    
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
