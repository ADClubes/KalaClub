VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmTraspaso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Títulos"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7905
   Icon            =   "frmTraspaso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmTraspaso.frx":030A
   ScaleHeight     =   5445
   ScaleWidth      =   7905
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   3690
      Picture         =   "frmTraspaso.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   4125
      Width           =   1000
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   2595
      Picture         =   "frmTraspaso.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guardar"
      Top             =   4125
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      Caption         =   "PARA:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   150
      TabIndex        =   3
      Top             =   1935
      Width           =   4500
      Begin VB.ComboBox cboAccionista 
         Height          =   315
         Index           =   1
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1275
         Visible         =   0   'False
         Width           =   4050
      End
      Begin VB.OptionButton optTraspaso 
         Caption         =   "Traspasar Acciones a otro &Accionista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   210
         TabIndex        =   5
         Top             =   540
         Width           =   3630
      End
      Begin VB.OptionButton optTraspaso 
         Caption         =   "Traspasar Acciones al &Club"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   225
         Width           =   2880
      End
      Begin VB.Label lblCombo2 
         Caption         =   "Accionista A &Quien Le Hará el Traspaso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   6
         Top             =   1020
         Visible         =   0   'False
         Width           =   3945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   150
      TabIndex        =   1
      Top             =   720
      Width           =   4500
      Begin VB.ComboBox cboAccionista 
         Height          =   315
         Index           =   0
         Left            =   210
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   4050
      End
      Begin VB.Label lblCombo1 
         Caption         =   "Accionista &Que Traspasa:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   2
         Top             =   345
         Width           =   2745
      End
   End
   Begin SSDataWidgets_B.SSDBGrid SSGrdDispon 
      Height          =   4350
      Left            =   4830
      TabIndex        =   0
      Top             =   780
      Width           =   2865
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   3
      HeadFont3D      =   3
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   1270
      Columns(0).Caption=   "SERIE"
      Columns(0).Name =   "SERIE"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1667
      Columns(1).Caption=   "NUMERO"
      Columns(1).Name =   "NUMERO"
      Columns(1).Alignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1005
      Columns(2).Caption=   "TIPO"
      Columns(2).Name =   "TIPO"
      Columns(2).Alignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      _ExtentX        =   5054
      _ExtentY        =   7673
      _StockProps     =   79
      Caption         =   "DISPONIBLES"
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
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   180
      TabIndex        =   10
      Top             =   4320
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59047937
      CurrentDate     =   37957
   End
   Begin SSDataWidgets_B.SSDBGrid SSGrdSelec 
      Height          =   3960
      Left            =   8955
      TabIndex        =   12
      Top             =   870
      Width           =   2865
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Col.Count       =   3
      HeadFont3D      =   3
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   1270
      Columns(0).Caption=   "SERIE"
      Columns(0).Name =   "SERIE"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(1).Width=   1667
      Columns(1).Caption=   "NUMERO"
      Columns(1).Name =   "NUMERO"
      Columns(1).Alignment=   2
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   2
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   1005
      Columns(2).Caption=   "TIPO"
      Columns(2).Name =   "TIPO"
      Columns(2).Alignment=   2
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      _ExtentX        =   5054
      _ExtentY        =   6985
      _StockProps     =   79
      Caption         =   "SELECCIONADOS"
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
   Begin VB.Label lblFecha 
      Caption         =   "&Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   11
      Top             =   4050
      Width           =   690
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TRASPASO DE TÍTULOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1470
      TabIndex        =   7
      Top             =   105
      Width           =   4965
   End
End
Attribute VB_Name = "frmTraspaso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA: TRASPASO
' Objetivo: PERMITE EL TRASPASO DE ACCIONES
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim blnExAccionista As Boolean
    Dim AdoRcsAccionistas As ADODB.Recordset

Private Sub cboAccionista_Click(index As Integer)
    Dim strCampo1 As String, strCampo2 As String
    If (index = 0) And (cboAccionista(0).ListIndex <> -1) Then
        Call ActualizaGrid
        strSQL = "SELECT idproptitulo, a_paterno & ' ' & a_materno & ' ' & nombre as accionista" & _
                        " FROM accionistas WHERE idproptitulo <> " & _
                        cboAccionista(0).ItemData(cboAccionista(0).ListIndex) & _
                        " ORDER BY a_paterno, a_materno, nombre"
        strCampo1 = "accionista"
        strCampo2 = "idproptitulo"
        Call LlenaCombos(cboAccionista(1), strSQL, strCampo1, strCampo2)
        Frame2.Enabled = True
    End If
End Sub

Private Sub cmdGuardar_Click()
    If VerificaDatos = True Then
        Call GuardaDatos
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Unload frmSelecReportes
    Me.Top = 0
    Me.Left = 0
    Me.Height = 5900
    Me.Width = 8000
End Sub

Private Sub Form_Load()
    Dim strCampo1 As String, strCampo2 As String
    blnExAccionista = False
    dtpFecha.Value = Now
    optTraspaso(0).Value = False
    optTraspaso(1).Value = False
    strSQL = "SELECT idproptitulo, a_paterno & ' ' & a_materno & ' ' & nombre as accionista" & _
                    " FROM accionistas ORDER BY a_paterno, a_materno, nombre"
    strCampo1 = "accionista"
    strCampo2 = "idproptitulo"
    Call LlenaCombos(cboAccionista(0), strSQL, strCampo1, strCampo2)
End Sub

Private Sub optTraspaso_Click(index As Integer)
    Select Case index
        Case 0
            lblCombo2.Visible = False
            cboAccionista(1).Visible = False
        Case 1
            lblCombo2.Visible = True
            cboAccionista(1).Visible = True
    End Select
End Sub

Private Function VerificaDatos()
    If cboAccionista(0).Text = "" Then
        MsgBox "¡ Nada que Guardar !", _
                    vbOKOnly + vbExclamation, "Propietarios de Títulos (Captura)"
        VerificaDatos = False
        cboAccionista(0).SetFocus
        Exit Function
    End If
    If (optTraspaso(0).Value = False) And (optTraspaso(1).Value = False) Then
        MsgBox "¿ A Quién le Hará el Traspaso ?", _
                    vbOKOnly + vbQuestion, "Propietarios de Títulos (Captura)"
        VerificaDatos = False
        Exit Function
    End If
    If (optTraspaso(1).Value = True) And (cboAccionista(1).Text = "") Then
     MsgBox "¿ A Quién le Hará el Traspaso ?", _
                    vbOKOnly + vbQuestion, "Propietarios de Títulos (Captura)"
        VerificaDatos = False
        cboAccionista(0).SetFocus
        Exit Function
    End If
    If SSGrdDispon.SelBookmarks.Count = 0 Then
        MsgBox "¡ Seleccione Los Títulos que Desea Traspasar !", _
                    vbOKOnly + vbExclamation, "Propietarios de Títulos (Captura)"
        VerificaDatos = False
        cboAccionista(0).SetFocus
        Exit Function
    End If
    If cboAccionista(0).Text = cboAccionista(1).Text Then
        MsgBox "¡ No se puede hacer un traspaso a la misma persona !", _
                    vbOKOnly + vbExclamation, "Propietarios de Títulos (Captura)"
        VerificaDatos = False
        cboAccionista(1).SetFocus
        Exit Function
    End If
    
    VerificaDatos = True
End Function

Sub ActualizaGrid()
    SSGrdDispon.RemoveAll
    strSQL = "SELECT serie, numero, tipo FROM titulos WHERE idpropietario = " & _
                    IIf(cboAccionista(0).Text = "", -1, cboAccionista(0).ItemData(cboAccionista(0).ListIndex)) & _
                    " ORDER BY numero"
    Set AdoRcsAccionistas = New ADODB.Recordset
    AdoRcsAccionistas.ActiveConnection = Conn
    AdoRcsAccionistas.LockType = adLockOptimistic
    AdoRcsAccionistas.CursorType = adOpenKeyset
    AdoRcsAccionistas.CursorLocation = adUseServer
    AdoRcsAccionistas.Open strSQL
    If Not AdoRcsAccionistas.EOF Then
        Do While Not AdoRcsAccionistas.EOF
            SSGrdDispon.AddItem AdoRcsAccionistas!Serie + _
            Chr$(9) + Str(AdoRcsAccionistas!Numero) + _
            Chr$(9) + AdoRcsAccionistas!tipo
            AdoRcsAccionistas.MoveNext
        Loop
    End If
End Sub

Sub GuardaDatos()
    Dim i As Integer
    Dim lngNumero As Long
    Dim strTipo, strSerie As String
    Dim AdoCmdInserta As ADODB.Command
    On Error GoTo err_Guarda
    cmdGuardar.Enabled = False
    cmdSalir.Enabled = False
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    'Mueve los registros seleccionados al grid oculto
    Call Mueve_A_La_Derecha
    
    'Guarda la información de los títulos seleccionados
    SSGrdSelec.MoveFirst
    For i = 0 To SSGrdSelec.Rows - 1
        strSerie = SSGrdSelec.Columns("serie").CellValue(SSGrdSelec.Bookmark)
        strTipo = SSGrdSelec.Columns("tipo").CellValue(SSGrdSelec.Bookmark)
        lngNumero = Val(SSGrdSelec.Columns("numero").CellValue(SSGrdSelec.Bookmark))
        
        'Actualiza los propietarios en la tabla TÍTULOS
        strSQL = "UPDATE titulos SET idpropietario = "
        If optTraspaso(0).Value = True Then
            strSQL = strSQL & "0"
        Else
            strSQL = strSQL & cboAccionista(1).ItemData(cboAccionista(1).ListIndex)
        End If
            #If SqlServer_ Then
                strSQL = strSQL & ",  fecha_asignacion = '" & _
                            Format(dtpFecha.Value, "yyyymmdd") & "' WHERE tipo = '" & _
                            strTipo & "' AND numero = " & lngNumero & " AND serie = '" & _
                            strSerie & "'"
            #Else
                strSQL = strSQL & ",  fecha_asignacion = '" & _
                            Format(dtpFecha.Value, "mm/dd/yyyy") & "' WHERE tipo = '" & _
                            strTipo & "' AND numero = " & lngNumero & " AND serie = '" & _
                            strSerie & "'"
            #End If
                            
        Set AdoCmdInserta = New ADODB.Command
        AdoCmdInserta.ActiveConnection = Conn
        AdoCmdInserta.CommandText = strSQL
        AdoCmdInserta.Execute
        
        'Crea el histórico
        #If SqlServer_ Then
            strSQL = "INSERT INTO histoacciones (tipo, numero, serie, " & _
                       "tipomovimiento, fechamovimiento, propanterior, propactual) " & _
                       "VALUES ('" & strTipo & "', " & lngNumero & ", '" & strSerie & _
                       "', 'TRASPASO', '" & Format(dtpFecha.Value, "yyyymmdd") & _
                       "', '" & cboAccionista(0).Text & "', '"
        #Else
            strSQL = "INSERT INTO histoacciones (tipo, numero, serie, " & _
                       "tipomovimiento, fechamovimiento, propanterior, propactual) " & _
                       "VALUES ('" & strTipo & "', " & lngNumero & ", '" & strSerie & _
                       "', 'TRASPASO', '" & Format(dtpFecha.Value, "mm/dd/yyyy") & _
                       "', '" & cboAccionista(0).Text & "', '"
        #End If
        
        If optTraspaso(0).Value = True Then
            strSQL = strSQL & "CLUB')"
        Else
            strSQL = strSQL & cboAccionista(1).Text & "')"
        End If
        
        Set AdoCmdInserta = New ADODB.Command
        AdoCmdInserta.ActiveConnection = Conn
        AdoCmdInserta.CommandText = strSQL
        AdoCmdInserta.Execute
        
        SSGrdSelec.MoveNext
    Next i
    
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    If blnExAccionista = True Then
        Call EliminaAccionista(cboAccionista(0).ItemData(cboAccionista(0).ListIndex))
    End If
    MsgBox "¡ Traspaso Realizado !", vbOKOnly + vbInformation, "Propietarios de Títulos"
    cmdGuardar.Enabled = True
    cmdSalir.Enabled = True
    Call Limpia
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    cmdGuardar.Enabled = True
    cmdSalir.Enabled = True
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Sub Limpia()
    Dim i As Integer
    blnExAccionista = False
    cboAccionista(0).ListIndex = -1
    cboAccionista(1).ListIndex = -1
    optTraspaso(0).Value = False
    optTraspaso(1).Value = False
    lblCombo2.Visible = False
    cboAccionista(1).Visible = False
    Frame2.Enabled = False
    SSGrdDispon.RemoveAll
    SSGrdSelec.RemoveAll
    dtpFecha.Value = Now
End Sub

Sub Mueve_A_La_Derecha()
    Dim i As Integer
    'Copia los registros seleccionados del primer grid al grid oculto
    For i = 0 To SSGrdDispon.SelBookmarks.Count - 1
        SSGrdSelec.AddItem SSGrdDispon.Columns("serie").CellValue(SSGrdDispon.SelBookmarks(i)) + Chr$(9) + _
                                         Str(SSGrdDispon.Columns("numero").CellValue(SSGrdDispon.SelBookmarks(i))) + Chr$(9) + _
                                         SSGrdDispon.Columns("tipo").CellValue(SSGrdDispon.SelBookmarks(i))
    Next i
    If SSGrdDispon.SelBookmarks.Count = SSGrdDispon.Rows Then
        blnExAccionista = True
    End If
End Sub
