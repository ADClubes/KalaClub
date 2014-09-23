VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMultiplesTitulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Títulos (Captura Múltiple)"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6285
   Icon            =   "frmTitulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmTitulos.frx":030A
   ScaleHeight     =   3030
   ScaleWidth      =   6285
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   315
      Left            =   1110
      TabIndex        =   5
      Top             =   1965
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      Format          =   59179011
      CurrentDate     =   37953
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Números a Generar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   2985
      TabIndex        =   6
      Top             =   855
      Width           =   2190
      Begin VB.TextBox txtHasta 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   1380
         TabIndex        =   10
         Top             =   765
         Width           =   700
      End
      Begin VB.TextBox txtDesde 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   700
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000005&
         X1              =   1005
         X2              =   1275
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         X1              =   990
         X2              =   1260
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Label lblRentables 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fi&nal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1380
         TabIndex        =   9
         Top             =   495
         Width           =   465
      End
      Begin VB.Label lblRentables 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Inicial:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   7
         Top             =   495
         Width           =   585
      End
   End
   Begin VB.TextBox txtSerie 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   200
      MaxLength       =   2
      TabIndex        =   3
      Top             =   1965
      Width           =   510
   End
   Begin VB.ComboBox cboTipoTitulo 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmTitulos.frx":B386
      Left            =   720
      List            =   "frmTitulos.frx":B390
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1050
      Width           =   2040
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      Height          =   840
      Left            =   5310
      Picture         =   "frmTitulos.frx":B3AB
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Guardar"
      Top             =   960
      Width           =   795
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   5325
      Picture         =   "frmTitulos.frx":B7ED
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   1905
      Width           =   795
   End
   Begin MSComctlLib.ProgressBar pgrAvance 
      Height          =   285
      Left            =   90
      TabIndex        =   13
      Top             =   2460
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TÍTULOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   3
      Left            =   1215
      TabIndex        =   14
      Top             =   105
      Width           =   3975
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "&Fecha Creación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   1140
      TabIndex        =   4
      Top             =   1545
      Width           =   885
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "&Serie Título:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   195
      TabIndex        =   2
      Top             =   1545
      Width           =   615
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo &Título:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   780
      Width           =   1095
   End
End
Attribute VB_Name = "frmMultiplesTitulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 ' ************************************************************************
' PANTALLA MÚLTIPLES TÍTULOS
' Objetivo: GENERA TÍTULOS DE MANERA MASIVA
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsTitulos As ADODB.Recordset
    Dim strSexo, strTipoTitulo As String
    
Private Function VerificaDatos()
    If (cboTipoTitulo.Text = "") Then
        MsgBox "¡ Favor de seleccionar el TIPO DELTÍTULO, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Titulos (Captura)"
        VerificaDatos = False
        cboTipoTitulo.SetFocus
        Exit Function
    End If
    If (txtSerie.Text = "") Then
        MsgBox "¡ Favor de Digitar la SERIE, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Titulos (Captura)"
        VerificaDatos = False
        txtSerie.SetFocus
        Exit Function
    End If
    If (txtDesde.Text = "") Then
        MsgBox "¡ Favor de Digitar el Valor INICIAL, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Titulos (Captura)"
        VerificaDatos = False
        txtDesde.SetFocus
        Exit Function
    End If
    If (txtHasta.Text = "") Then
        MsgBox "¡ Favor de Digitar el Valor FINAL, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Titulos (Captura)"
        VerificaDatos = False
        txtHasta.SetFocus
        Exit Function
    End If
    If (Val(Trim(txtDesde.Text)) > Val(Trim(txtHasta.Text))) Then
        MsgBox "¡ No son Correctos los Datos INICIAL y/o FINAL !", _
                    vbOKOnly + vbExclamation, "Titulos (Captura)"
        VerificaDatos = False
        txtDesde.SetFocus
        Exit Function
    End If
    VerificaDatos = True
End Function

Private Sub cboTipoTitulo_Click()
    Select Case cboTipoTitulo.Text
        Case "ACCIONISTA"
            strTipoTitulo = "AC"
        Case "MEMBRECÍA"
            strTipoTitulo = "ME"
        End Select
End Sub

Private Sub cmdGuardar_Click()
    If VerificaDatos = True Then
        Call GuardaDatos
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmCatalogos.Enabled = False
    dtpFecha.Value = Now
    cboTipoTitulo.ListIndex = 0
    strTipoTitulo = "AC"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCatalogos.Enabled = True
End Sub

Private Sub GuardaDatos()
    Dim AdoCmdInserta As ADODB.Command
    Dim i, intNoGuardados, intRespuesta, intTotal As Integer
    On Error GoTo err_Guarda
    intRespuesta = MsgBox("Se Generarán los Títulos Indicados. " & _
                                        "¿Está Seguro?", vbOKCancel + vbQuestion, "Generación Múltiple")
    If intRespuesta = 2 Then
        Exit Sub
    End If
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    cmdGuardar.Enabled = False
    cmdSalir.Enabled = False
    intTotal = (Val(txtHasta.Text) - Val(txtDesde.Text)) + 1
    pgrAvance.Visible = True
    pgrAvance.Max = intTotal
    pgrAvance.Value = 0
    intNoGuardados = 0
    For i = Val(txtDesde.Text) To Val(txtHasta.Text)
        'Consultamos para ver si no exixte previamente es registro
        strSQL = "SELECT tipo FROM titulos WHERE (tipo = '" & strTipoTitulo & _
                        "') AND (numero = " & i & ") AND (serie = '" & txtSerie.Text & "')"
        Set AdoRcsTitulos = New ADODB.Recordset
        AdoRcsTitulos.ActiveConnection = Conn
        AdoRcsTitulos.LockType = adLockOptimistic
        AdoRcsTitulos.CursorType = adOpenKeyset
        AdoRcsTitulos.CursorLocation = adUseServer
        AdoRcsTitulos.Open strSQL
        If AdoRcsTitulos.EOF Then
            'Crea los Títulos en la tabla TÍTULOS
            #If SqlServer_ Then
                strSQL = "INSERT INTO titulos (tipo, numero, serie, fecha_creacion) " & _
                            "VALUES ('" & strTipoTitulo & "', " & i & ", '" & txtSerie.Text & _
                            "', '" & Format(dtpFecha.Value, "yyyymmdd") & "')"
            #Else
                strSQL = "INSERT INTO titulos (tipo, numero, serie, fecha_creacion) " & _
                            "VALUES ('" & strTipoTitulo & "', " & i & ", '" & txtSerie.Text & _
                            "', #" & Format(dtpFecha.Value, "mm/dd/yyyy") & "#)"
            #End If
            
            Set AdoCmdInserta = New ADODB.Command
            AdoCmdInserta.ActiveConnection = Conn
            AdoCmdInserta.CommandText = strSQL
            AdoCmdInserta.Execute
            
            'Ahora los crea en el histórico
            #If SqlServer_ Then
                strSQL = "INSERT INTO histoacciones (tipo, numero, serie, " & _
                            "tipomovimiento, fechamovimiento, propanterior, propactual) " & _
                            "VALUES ('" & strTipoTitulo & "', " & i & ", '" & txtSerie.Text & _
                            "', 'ALTA', '" & Format(dtpFecha.Value, "yyyymmdd") & _
                            "','', 'CLUB')"
            #Else
                strSQL = "INSERT INTO histoacciones (tipo, numero, serie, " & _
                            "tipomovimiento, fechamovimiento, propanterior, propactual) " & _
                            "VALUES ('" & strTipoTitulo & "', " & i & ", '" & txtSerie.Text & _
                            "', 'ALTA', #" & Format(dtpFecha.Value, "mm/dd/yyyy") & _
                            "#,'', 'CLUB')"
            #End If
            
            Set AdoCmdInserta = New ADODB.Command
            AdoCmdInserta.ActiveConnection = Conn
            AdoCmdInserta.CommandText = strSQL
            AdoCmdInserta.Execute
        Else
            intNoGuardados = intNoGuardados + 1
        End If
        pgrAvance.Value = pgrAvance.Value + 1
    Next i
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    MsgBox "¡ Proceso Terminado ! Títulos Creados: " & intTotal - intNoGuardados & ", No Creados: " & intNoGuardados
    cmdGuardar.Enabled = True
    cmdSalir.Enabled = True
    Call Limpia
    frmCatalogos.AdoDcCatal.REFRESH
    frmCatalogos.grdCatalogos.REFRESH
    frmCatalogos.lblTotal.Caption = Format(frmCatalogos.AdoDcCatal.Recordset.RecordCount, "#######")
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

Private Sub txtDesde_GotFocus()
    txtDesde.SelStart = 0
    txtDesde.SelLength = Len(txtDesde)
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 22, 48 To 57     'Backspace, <Ctrl+V> y del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtHasta_GotFocus()
    txtHasta.SelStart = 0
    txtHasta.SelLength = Len(txtHasta)
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 22, 48 To 57     'Backspace, <Ctrl+V> y del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
    End Select
End Sub

Sub Limpia()
'    cboTipoTitulo.ListIndex = -1
    txtSerie.Text = ""
    txtDesde.Text = ""
    txtHasta.Text = ""
    pgrAvance.Value = 0
    pgrAvance.Visible = False
End Sub

Private Sub txtSerie_GotFocus()
    txtSerie.SelStart = 0
    txtSerie.SelLength = Len(txtSerie)
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
