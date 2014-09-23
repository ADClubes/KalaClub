VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmProspectos 
   Caption         =   "Datos de prospectos"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   8460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   615
      Left            =   6960
      TabIndex        =   14
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   615
      Left            =   5520
      TabIndex        =   13
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Frame frmDatosVentas 
      Caption         =   "Datos capturados por ventas"
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   3120
      Width           =   7695
      Begin VB.TextBox txtControl 
         Height          =   375
         Index           =   7
         Left            =   240
         MaxLength       =   50
         TabIndex        =   12
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtControl 
         Height          =   375
         Index           =   6
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   9
         Top             =   600
         Width           =   3375
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbAsesor2 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   2535
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
         Columns(0).Width=   4551
         Columns(0).Caption=   "Nombre"
         Columns(0).Name =   "Nombre"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2514
         Columns(1).Caption=   "IdVendedor"
         Columns(1).Name =   "IdVendedor"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
         Enabled         =   0   'False
         DataFieldToDisplay=   "Column 0"
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbNivel3 
         Height          =   375
         Left            =   3360
         TabIndex        =   15
         Top             =   2880
         Visible         =   0   'False
         Width           =   1935
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
         Columns(0).Width=   3200
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "IdMedio"
         Columns(1).Name =   "IdMedio"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbNivel2 
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2535
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
         Columns(0).Width=   3200
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "IdMedio"
         Columns(1).Name =   "IdMedio"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbVendedor 
         Height          =   375
         Left            =   3120
         TabIndex        =   11
         Top             =   1680
         Width           =   2535
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
         Columns(0).Width=   4551
         Columns(0).Caption=   "Nombre"
         Columns(0).Name =   "Nombre"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   2514
         Columns(1).Caption=   "IdVendedor"
         Columns(1).Name =   "IdVendedor"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
         DataFieldToDisplay=   "Column 0"
      End
      Begin VB.Label lblControl 
         Caption         =   "Vendedor"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   11
         Left            =   3120
         TabIndex        =   28
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblControl 
         Caption         =   "# Folio pase"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   27
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lblControl 
         Caption         =   "Cerrador"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblControl 
         Caption         =   "Comentarios medio"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   25
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblControl 
         Caption         =   "Medio por el que se entero"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame frmDatosIniciales 
      Caption         =   "Datos iniciales"
      Height          =   2775
      Left            =   360
      TabIndex        =   16
      Top             =   240
      Width           =   7695
      Begin VB.TextBox txtControl 
         Height          =   375
         Index           =   0
         Left            =   240
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtControl 
         Height          =   375
         Index           =   1
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtControl 
         Height          =   375
         Index           =   2
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtControl 
         Height          =   375
         Index           =   3
         Left            =   240
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtControl 
         Height          =   375
         Index           =   4
         Left            =   3120
         MaxLength       =   20
         TabIndex        =   5
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtControl 
         Height          =   375
         Index           =   5
         Left            =   240
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2160
         Width           =   2535
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbNivel1 
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   2160
         Width           =   3375
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
         Columns(0).Width=   5001
         Columns(0).Caption=   "Descripcion"
         Columns(0).Name =   "Descripcion"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3200
         Columns(1).Visible=   0   'False
         Columns(1).Caption=   "IdMedio"
         Columns(1).Name =   "IdMedio"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   3
         Columns(1).FieldLen=   256
         _ExtentX        =   5953
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.Label lblControl 
         Caption         =   "Medio por el que se entero"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   23
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label lblControl 
         Caption         =   "Nombre(s)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblControl 
         Caption         =   "Apellido Paterno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   21
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblControl 
         Caption         =   "Apellido Materno"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   20
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblControl 
         Caption         =   "Colonia"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblControl 
         Caption         =   "Telefono"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   3120
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblControl 
         Caption         =   "Pregunto por"
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmProspectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lRecNum As Long
Public bReadOnly As Boolean

Private Sub cmdGuardar_Click()
    Dim adocmdProspectos As ADODB.Command
    
    'se restringe desde base de datos
'    If sDB_NivelUser <> 6 Then
'        MsgBox "Función disponible sólo para ejecutivos de ventas"
'        Exit Sub
'    End If
    
    '23/12/2011 UCM
    If Not ChecaSeguridad(Me.Name, cmdGuardar.Name) Then
        Exit Sub
    End If
    
    If Not ValidaDatos() Then
        Exit Sub
    End If
    
    On Error GoTo Error_Catch
    
    Screen.MousePointer = vbHourglass
    
    If lRecNum = 0 Then
        strSQL = "INSERT INTO PROSPECTOS ("
        strSQL = strSQL & " FechaRegistro" & ","
        strSQL = strSQL & " HoraRegistro" & ","
        strSQL = strSQL & " NOMBRE" & ","
        strSQL = strSQL & " A_PATERNO" & ","
        strSQL = strSQL & " A_MATERNO" & ","
        strSQL = strSQL & " COLONIA" & ","
        strSQL = strSQL & " TELEFONO" & ","
        strSQL = strSQL & " ASESOR" & ","
        strSQL = strSQL & " MEDIO1" & ","
        strSQL = strSQL & " MEDIO2" & ","
        strSQL = strSQL & " MEDIO3" & ","
        strSQL = strSQL & " MEDIO4" & ","
        strSQL = strSQL & " USUARIO" & ","
        strSQL = strSQL & " MEDIO5" & ")"
        strSQL = strSQL & " VALUES ("
        #If SqlServer_ Then
            strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "',"
        #Else
            strSQL = strSQL & "#" & Format(Date, "mm/dd/yyyy") & "#,"
        #End If
        strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "'" & ","
        strSQL = strSQL & "'" & Trim(Me.txtControl(0).Text) & "'" & ","
        strSQL = strSQL & "'" & Trim(Me.txtControl(1).Text) & "'" & ","
        strSQL = strSQL & "'" & Trim(Me.txtControl(2).Text) & "'" & ","
        strSQL = strSQL & "'" & Trim(Me.txtControl(3).Text) & "'" & ","
        strSQL = strSQL & "'" & Trim(Me.txtControl(4).Text) & "'" & ","
        strSQL = strSQL & "'" & Trim(Me.txtControl(5).Text) & "'" & ","
        strSQL = strSQL & "'" & UCase(Trim(Me.sscmbNivel1.Text)) & "'" & ","
        strSQL = strSQL & "'" & UCase(Trim(Me.sscmbNivel2.Text)) & "'" & ","
        strSQL = strSQL & "'" & UCase(Trim(Me.sscmbNivel3.Text)) & "'" & ","
        strSQL = strSQL & "'" & "'" & ","
        strSQL = strSQL & "'" & sDB_User & "'" & ","
        strSQL = strSQL & "'" & Trim(Me.txtControl(6).Text) & "'" & ")"
    Else
        strSQL = "UPDATE PROSPECTOS SET"
        strSQL = strSQL & " NOMBRE=" & "'" & Trim(Me.txtControl(0).Text) & "',"
        strSQL = strSQL & " A_PATERNO=" & "'" & Trim(Me.txtControl(1).Text) & "',"
        strSQL = strSQL & " A_MATERNO=" & "'" & Trim(Me.txtControl(2).Text) & "',"
        strSQL = strSQL & " COLONIA=" & "'" & Trim(Me.txtControl(3).Text) & "',"
        strSQL = strSQL & " TELEFONO=" & "'" & Trim(Me.txtControl(4).Text) & "',"
        strSQL = strSQL & " MEDIO2=" & "'" & UCase(Trim(Me.sscmbNivel2.Text)) & "'" & ","
        strSQL = strSQL & " MEDIO5=" & "'" & Trim(Me.txtControl(6).Text) & "',"
        strSQL = strSQL & " IdVendedor=" & Me.sscmbVendedor.Columns("IdVendedor").Value & ","
        #If SqlServer_ Then
            strSQL = strSQL & " FechaPreventa=" & "'" & Format(Date, "yyyymmdd") & "'" & ","
        #Else
            strSQL = strSQL & " FechaPreventa=" & "#" & Format(Date, "mm/dd/yyyy") & "#" & ","
        #End If
        strSQL = strSQL & " HoraPreventa=" & "'" & Format(Now, "Hh:Nn:Ss") & "'" & ","
        strSQL = strSQL & " StatusProspecto=" & 1 & ","
        strSQL = strSQL & " FolioPase=" & Trim(Me.txtControl(7).Text) & ","
        strSQL = strSQL & " IdVendedor2=" & Me.sscmbAsesor2.Columns("IdVendedor").Value
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdProspecto=" & lRecNum
        strSQL = strSQL & " And StatusProspecto <= 1"
    End If
    
    Set adocmdProspectos = New ADODB.Command
    adocmdProspectos.ActiveConnection = Conn
    adocmdProspectos.CommandType = adCmdText
    adocmdProspectos.CommandText = strSQL
    adocmdProspectos.Execute
    
    Set adocmdProspectos = Nothing
    
    On Error GoTo 0
    Screen.MousePointer = vbDefault
    
    If lRecNum = 0 Then
        MsgBox "Prospecto ingresado", vbInformation + vbOKOnly, "Correcto"
        LimpiaControles
    
        Me.txtControl(0).SetFocus
        Me.cmdGuardar.Enabled = False
        Me.sscmbNivel2.Enabled = False
        Me.sscmbNivel3.Enabled = False
        
    Else
        MsgBox "Prospecto actualizado", vbInformation + vbOKOnly, "Correcto"
    End If
    
    Exit Sub
    
Error_Catch:
    Screen.MousePointer = vbDefault
    MsgError
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.txtControl(0).SetFocus
End Sub

Private Sub Form_Load()
    LlenaCombo Me.sscmbNivel1, 0, 0
    LlenaCombo Me.sscmbNivel2, 0, 0
    
'    If lRecNum > CLng(ObtieneParametro("MAX_ULTIMO_PROSPECTO")) Then
'        strSQL = ""
'        strSQL = "SELECT Nombre, IdUsuario"
'        strSQL = strSQL & " FROM USUARIOS_SISTEMA"
'        strSQL = strSQL & " WHERE IdPerfil = 6 AND Status = 'A'"
'        strSQL = strSQL & " ORDER BY Nombre"
'    Else
        strSQL = ""
        strSQL = "SELECT Nombre, IdVendedor"
        strSQL = strSQL & " FROM Vendedores WHERE Activo=1"
        strSQL = strSQL & " ORDER BY Nombre"
'    End If
    
    LlenaSsCombo Me.sscmbAsesor2, Conn, strSQL, 2
    sscmbAsesor2.Text = NombreVendedor(iDB_IdUser)
    
    LlenaSsCombo Me.sscmbVendedor, Conn, strSQL, 2
    
    If lRecNum = 0 Then
        Me.frmDatosVentas.Visible = False
    Else
        CargaDatos lRecNum
        Me.txtControl(5).Enabled = False
        Me.sscmbNivel1.Enabled = False
        Me.frmDatosVentas = True
    End If
    
    If bReadOnly Then
        Me.cmdGuardar.Enabled = False
    End If
    
    CentraForma MDIPrincipal, Me
    
    Exit Sub
End Sub

Private Sub sscmbNivel1_Click()
    Dim lRen As Long
    
    'Me.sscmbNivel2.Visible = False
    'Me.sscmbNivel3.Visible = False
    'Me.txtControl(6).Visible = False
    
    'lRen = Me.sscmbNivel1.AddItemRowIndex(Me.sscmbNivel1.Bookmark)
End Sub

Private Sub sscmbNivel2_Click()
    Dim lRen As Long
    Dim lRen2 As Long
    
    'Me.sscmbNivel3.Visible = False
    'Me.txtControl(6).Visible = False
    
    'lRen = Me.sscmbNivel1.AddItemRowIndex(Me.sscmbNivel1.Bookmark)
    'lRen2 = Me.sscmbNivel2.AddItemRowIndex(Me.sscmbNivel2.Bookmark)
End Sub

Private Sub txtControl_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 7 Then
        Select Case KeyAscii
            Case 8 ' Tecla backspace
                KeyAscii = KeyAscii
            Case 48 To 57 ' Numeros del 0 al 9
                KeyAscii = KeyAscii
            Case Else
                KeyAscii = 0
        End Select
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Function ValidaDatos() As Boolean
    Dim lI As Long
    Dim sCampo As String

    sCampo = ""

    ValidaDatos = True
    
    For lI = 0 To 4
        If Me.txtControl(lI).Text = "" Then
            ValidaDatos = False
            Select Case lI
                Case 0
                    sCampo = "Nombre"
                Case 1
                    sCampo = "Apellido paterno"
                Case 2
                    sCampo = "Apellido materno"
                Case 3
                    sCampo = "Colonia"
                Case 4
                    sCampo = "Telefono"
            End Select
            
            MsgBox "El campo " & sCampo & " no puede quedar vacio!", vbCritical, "Verifique"
            Me.txtControl(lI).SetFocus
            Exit Function
                            
        End If
    Next
    
    If Me.sscmbNivel1.Text = "" Then
        ValidaDatos = False
        MsgBox "El campo medio no puede quedar vacio!", vbCritical, "Verifique"
        Me.sscmbNivel1.SetFocus
        Exit Function
    End If
    
    If lRecNum Then
    
        If Me.sscmbNivel2.Rows <> 0 And Me.sscmbNivel2.Text = "" Then
            ValidaDatos = False
            MsgBox "El campo medio 2 no puede quedar vacio!", vbCritical, "Verifique"
            Me.sscmbNivel2.SetFocus
            Exit Function
        End If
    
'        If Me.sscmbNivel3.Rows <> 0 And Me.sscmbNivel3.Text = "" Then
'            ValidaDatos = False
'            MsgBox "El campo medio 3 no puede quedar vacio!", vbCritical, "Verifique"
'            Me.sscmbNivel3.SetFocus
'            Exit Function
'        End If
    End If
End Function

Private Sub LimpiaControles()
    Dim lI As Long
    
    For lI = 0 To 6
        Me.txtControl(lI).Text = ""
    Next
    
    Me.sscmbNivel1.Text = ""
End Sub

Private Sub LlenaCombo(sscmbControl As SSOleDBCombo, lIdMedio As Long, lNivel As Long)
    Dim adorcsMedio As ADODB.Recordset

    strSQL = ""

    strSQL = "SELECT IdMedio, Nivel, Padre, Descripcion, TextoLibre"
    strSQL = strSQL & " FROM MEDIOS"
    strSQL = strSQL & " WHERE Nivel=" & lNivel
    If lNivel > 0 Then
        strSQL = strSQL & " AND Padre=" & lIdMedio
    End If
    strSQL = strSQL & " ORDER BY IdMedio"
    
    Set adorcsMedio = New ADODB.Recordset
    adorcsMedio.CursorLocation = adUseServer
    
    sscmbControl.RemoveAll
    sscmbControl.Text = vbNullString
    
    adorcsMedio.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do Until adorcsMedio.EOF
        sscmbControl.AddItem adorcsMedio!Descripcion & vbTab & adorcsMedio!IdMedio
        adorcsMedio.MoveNext
    Loop
    
    adorcsMedio.Close

    Set adorcsMedio = Nothing
    
    If sscmbControl.Rows > 0 Then
        sscmbControl.Visible = True
    End If
End Sub

Private Sub CargaDatos(lRec As Long)
    Dim adorcs As ADODB.Recordset
    
    If lRec > CLng(ObtieneParametro("MAX_ULTIMO_PROSPECTO")) Then
        strSQL = ""
        strSQL = "SELECT Prospectos.Nombre, A_Paterno, A_Materno, Colonia, Telefono, Asesor, Medio1, Medio2, Medio5, IdProspecto, FolioPase, USUARIOS_SISTEMA.Nombre AS VendedoresNom, V2.Nombre AS VendedorNom"
        strSQL = strSQL & " FROM Prospectos LEFT JOIN USUARIOS_SISTEMA ON Prospectos.IdVendedor2=USUARIOS_SISTEMA.IdUsuario LEFT JOIN USUARIOS_SISTEMA V2 ON Prospectos.IdVendedor=V2.IdUsuario"
        strSQL = strSQL & " WHERE IdProspecto=" & lRec
    Else
        strSQL = ""
        strSQL = "SELECT Prospectos.Nombre, A_Paterno , A_Materno, Colonia, Telefono, Asesor, Medio1, Medio2, Medio5, IdProspecto, FolioPase, Vendedores.Nombre AS VendedoresNom, V2.Nombre AS VendedorNom"
        strSQL = strSQL & " FROM Prospectos LEFT JOIN VENDEDORES ON Prospectos.IdVendedor2=VENDEDORES.IdVendedor LEFT JOIN VENDEDORES V2 ON Prospectos.IdVendedor=V2.IdVendedor"
        strSQL = strSQL & " WHERE IdProspecto=" & lRec
    End If
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not adorcs.EOF Then
        
        Me.txtControl(0).Text = adorcs!Nombre
        Me.txtControl(1).Text = adorcs!A_Paterno
        Me.txtControl(2).Text = adorcs!A_Materno
        Me.txtControl(3).Text = adorcs!colonia
        Me.txtControl(4).Text = adorcs!Telefono
        Me.txtControl(5).Text = adorcs!Asesor
        Me.sscmbNivel1.Text = adorcs!Medio1
        
        Me.sscmbNivel2.Text = IIf(IsNull(adorcs!Medio2), "", adorcs!Medio2)
        Me.txtControl(6).Text = IIf(IsNull(adorcs!Medio5), "", adorcs!Medio5)
        If Not IsNull(adorcs!VendedoresNom) Then
            Me.sscmbAsesor2.Text = adorcs![VendedoresNom]
        End If
        
        If Not IsNull(adorcs!VendedorNom) Then
            Me.sscmbVendedor.Text = adorcs!VendedorNom
        End If
        
        Me.txtControl(7).Text = IIf(IsNull(adorcs!FolioPase), 0, adorcs!FolioPase)
        
    End If
    
    adorcs.Close
    Set adorcs = Nothing
End Sub
