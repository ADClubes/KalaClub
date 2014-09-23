VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConceptoIngresos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Concepto de Ingresos (Captura)"
   ClientHeight    =   5895
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   8190
   Icon            =   "frmIngresos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8190
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmIngresos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblIngresos(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblIngresos(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblIngresos(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblIngresos(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblIngresos(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblIngresos(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblIngresos(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtIngresos(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtIngresos(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtIngresos(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbFacORec"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "chkPeriodico"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtIngresos(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtIngresos(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtIngresos(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtIngresos(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkReqIns"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkReqUsu"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Cupones"
      TabPicture(1)   =   "frmIngresos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtIngresos(11)"
      Tab(1).Control(1)=   "txtIngresos(10)"
      Tab(1).Control(2)=   "txtIngresos(9)"
      Tab(1).Control(3)=   "txtIngresos(8)"
      Tab(1).Control(4)=   "txtIngresos(7)"
      Tab(1).Control(5)=   "lblIngresos(11)"
      Tab(1).Control(6)=   "lblIngresos(10)"
      Tab(1).Control(7)=   "lblIngresos(9)"
      Tab(1).Control(8)=   "lblIngresos(8)"
      Tab(1).Control(9)=   "lblIngresos(7)"
      Tab(1).ControlCount=   10
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   11
         Left            =   -72720
         MaxLength       =   9
         TabIndex        =   33
         Top             =   4320
         Width           =   1080
      End
      Begin VB.TextBox txtIngresos 
         Height          =   1890
         Index           =   10
         Left            =   -72720
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   2280
         Width           =   4080
      End
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   9
         Left            =   -72720
         MaxLength       =   30
         TabIndex        =   18
         Top             =   1800
         Width           =   4080
      End
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   8
         Left            =   -72720
         MaxLength       =   3
         TabIndex        =   17
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   7
         Left            =   -72720
         MaxLength       =   2
         TabIndex        =   16
         Top             =   600
         Width           =   840
      End
      Begin VB.CheckBox chkReqUsu 
         Caption         =   "Requiere Usuario"
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
         Left            =   1800
         TabIndex        =   13
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CheckBox chkReqIns 
         Caption         =   "Requiere Instructor"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   6
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   15
         Top             =   4320
         Width           =   2400
      End
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   3
         Left            =   1815
         TabIndex        =   8
         Top             =   2160
         Width           =   1230
      End
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   5
         Left            =   4500
         TabIndex        =   11
         Top             =   2610
         Width           =   480
      End
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   4
         Left            =   1815
         TabIndex        =   10
         Top             =   2610
         Width           =   450
      End
      Begin VB.CheckBox chkPeriodico 
         Alignment       =   1  'Right Justify
         Caption         =   "&Es Periodico:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3750
         TabIndex        =   9
         Top             =   2205
         Width           =   1455
      End
      Begin VB.ComboBox cmbFacORec 
         Height          =   315
         Left            =   1815
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3225
         Width           =   2415
      End
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   1
         Left            =   1815
         TabIndex        =   6
         Top             =   1200
         Width           =   4365
      End
      Begin VB.TextBox txtIngresos 
         Height          =   330
         Index           =   2
         Left            =   1815
         TabIndex        =   7
         Top             =   1650
         Width           =   4365
      End
      Begin VB.TextBox txtIngresos 
         Alignment       =   2  'Center
         Height          =   330
         Index           =   0
         Left            =   1815
         TabIndex        =   5
         Top             =   720
         Width           =   1200
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe a pagar al instructor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   -74760
         TabIndex        =   32
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones"
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
         Index           =   10
         Left            =   -74760
         TabIndex        =   31
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción a imprimir"
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
         Index           =   9
         Left            =   -74760
         TabIndex        =   30
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "Dias de vigencia"
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
         Index           =   8
         Left            =   -74760
         TabIndex        =   29
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "# de Cupones a emitir"
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
         Index           =   7
         Left            =   -74760
         TabIndex        =   28
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
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
         Index           =   6
         Left            =   840
         TabIndex        =   27
         Top             =   4320
         Width           =   720
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "&Monto: $"
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
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   2250
         Width           =   1575
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "&I. V. A. % :"
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
         Index           =   4
         Left            =   240
         TabIndex        =   25
         Top             =   2685
         Width           =   1575
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "Im&puesto 2: (0-100) %"
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
         Index           =   5
         Left            =   2445
         TabIndex        =   24
         Top             =   2715
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Factura o Recibo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   495
         TabIndex        =   23
         Top             =   3225
         Width           =   1095
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "&Descripción:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   1305
         Width           =   1575
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "C&uenta Contable:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Label lblIngresos 
         BackStyle       =   0  'Transparent
         Caption         =   "&Clave:"
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
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   840
      Left            =   7035
      Picture         =   "frmIngresos.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   2265
      Width           =   795
   End
   Begin VB.CommandButton cmdGuardar 
      Height          =   840
      Left            =   7020
      Picture         =   "frmIngresos.frx":064C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar"
      Top             =   1125
      Width           =   795
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONCEPTOS DE INGRESOS"
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
      Left            =   2085
      TabIndex        =   1
      Top             =   90
      Width           =   3975
   End
   Begin VB.Label lblClave 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1890
      TabIndex        =   0
      Top             =   4320
      Width           =   1110
   End
End
Attribute VB_Name = "frmConceptoIngresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA CONCEPTO DE INGRESOS
' Objetivo: CATÁLOGO DE CONCEPTO DE INGRESOS
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsIngresos As ADODB.Recordset
    
Private Function VerificaDatos()
    Dim i, intInicio As Integer
    If txtIngresos(0).Enabled = True Then
        intInicio = 0
    Else
        intInicio = 1
    End If
    For i = intInicio To 5
        Select Case i
            Case 0
                If (txtIngresos(i).Text = "") Then
                    MsgBox "¡ Favor de Llenar la Casilla CLAVE, Pues No es Opcional !", _
                                 vbOKOnly + vbExclamation, "Ingresos (Captura)"
                    VerificaDatos = False
                    txtIngresos(i).SetFocus
                    Exit Function
                End If
            Case 1
                If txtIngresos(i).Text = "" Then
                    MsgBox "¡ Favor de Llenar la Casilla DESCRIPCIÓN, Pues No es Opcional.", _
                                vbOKOnly + vbExclamation, "Ingresos (Captura)"
                    VerificaDatos = False
                    txtIngresos(i).SetFocus
                    Exit Function
                End If
            Case 4, 5
                If (Val(Trim(txtIngresos(i))) < 0) Or (Val(Trim(txtIngresos(i))) > 100) Then
                    MsgBox "¡ Impuesto Incorrecto !", vbOKOnly + vbExclamation, _
                                "Ingresos (Captura)"
                    VerificaDatos = False
                    txtIngresos(i).SetFocus
                    Exit Function
                End If
        End Select
        VerificaDatos = True
    Next
    
    'Se verifica el combo de Factura o Recibo
    If Me.cmbFacORec.Text = "" Then
        MsgBox "Seleccionar el tipo de documento!", vbExclamation + vbOKOnly, "Ingresos"
        VerificaDatos = False
        Exit Function
    End If
    
    
    If Val(Me.txtIngresos(7).Text) And Me.txtIngresos(9).Text = vbNullString Then
        MsgBox "Falta la descripción del cupón!", vbExclamation + vbOKOnly, "Ingresos"
        VerificaDatos = False
        Exit Function
    End If
    
    
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT idconcepto FROM concepto_ingresos WHERE idconcepto = " & _
                    Val(Trim(txtIngresos(0).Text))
        Set AdoRcsIngresos = New ADODB.Recordset
        AdoRcsIngresos.ActiveConnection = Conn
        AdoRcsIngresos.LockType = adLockOptimistic
        AdoRcsIngresos.CursorType = adOpenKeyset
        AdoRcsIngresos.CursorLocation = adUseServer
        AdoRcsIngresos.Open strSQL
        If Not AdoRcsIngresos.EOF Then
            MsgBox "Ya Existe Un Registro Con La Clave: " & _
                        txtIngresos(0).Text, vbInformation + vbOKOnly, "Ingresos"
            AdoRcsIngresos.Close
            VerificaDatos = False
            txtIngresos(0).SetFocus
            Exit Function
        Else
            AdoRcsIngresos.Close
            VerificaDatos = True
        End If
    End If
    
    
    
    
End Function

Private Sub cmdGuardar_Click()
    If VerificaDatos = True Then
        If frmCatalogos.lblModo.Caption = "A" Then
            Call GuardaDatos
        Else
            Call RemplazaDatos
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    frmCatalogos.Enabled = False
    
    
    LlenaComboFacORec
    
    If frmCatalogos.lblModo.Caption = "A" Then
        txtIngresos(0).Enabled = True
        Call Llena_txtIngresos
    Else
        txtIngresos(0).Enabled = False
        Call LlenaDatos
    End If
    
    Me.SSTab1.Tab = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
        frmCatalogos.Enabled = True
End Sub

Private Sub GuardaDatos()
    Dim AdoCmdInserta As ADODB.Command
    On Error GoTo err_Guarda
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    
    strSQL = "INSERT INTO concepto_ingresos ("
    strSQL = strSQL & " idconcepto,"
    strSQL = strSQL & " descripcion,"
    strSQL = strSQL & " cuentacontable,"
    strSQL = strSQL & " monto,"
    strSQL = strSQL & " impuesto1,"
    strSQL = strSQL & " impuesto2,"
    strSQL = strSQL & " esperiodico,"
    strSQL = strSQL & " FacORec,"
    '23/03/2006
    strSQL = strSQL & " RequiereUsuario,"
    strSQL = strSQL & " RequiereInstructor,"
    strSQL = strSQL & " Grupo,"
    '01/12/2006
    strSQL = strSQL & " NumeroCupones,"
    strSQL = strSQL & " DiasVigenciaCupones,"
    strSQL = strSQL & " DescripcionCupon,"
    strSQL = strSQL & " ObservacionesCupon,"
    '16/06/2008
    strSQL = strSQL & " ImporteAPagar)"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & Val(Trim(txtIngresos(0))) & ","
    strSQL = strSQL & "'" & Trim(txtIngresos(1)) & "',"
    strSQL = strSQL & "'" & Trim(txtIngresos(2)) & "',"
    strSQL = strSQL & Val(Trim(txtIngresos(3))) & ","
    strSQL = strSQL & Val(Trim(txtIngresos(4))) & ","
    strSQL = strSQL & Val(Trim(txtIngresos(5))) & ","
    strSQL = strSQL & chkPeriodico.Value & ","
    strSQL = strSQL & "'" & Left$(Me.cmbFacORec.Text, 1) & "',"
    '23/03/2006
    strSQL = strSQL & Me.chkReqUsu.Value & ","
    strSQL = strSQL & Me.chkReqIns.Value & ","
    strSQL = strSQL & "'" & UCase(Trim(Me.txtIngresos(6).Text)) & "',"
    '1/12/2006
    strSQL = strSQL & Val(Trim(Me.txtIngresos(7).Text)) & ","
    strSQL = strSQL & Val(Trim(Me.txtIngresos(8).Text)) & ","
    strSQL = strSQL & "'" & Trim(Me.txtIngresos(9).Text) & "',"
    strSQL = strSQL & "'" & Trim(Me.txtIngresos(10).Text) & "',"
    '16/06/2008
    strSQL = strSQL & Val(Trim(Me.txtIngresos(11).Text)) & ")"
    
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    MsgBox "¡Registro Ingresado!"
    Call Limpia
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub RemplazaDatos()
   Dim AdoCmdRemplaza As ADODB.Command
    On Error GoTo err_Guarda
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    'Eliminamos elregistro existente
    strSQL = "DELETE FROM concepto_ingresos WHERE idconcepto = " & _
                    Val(lblClave.Caption)
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Ahora Insertamos el nuevo registro que sustituye al anterior
    
    strSQL = "INSERT INTO concepto_ingresos ("
    strSQL = strSQL & " idconcepto,"
    strSQL = strSQL & " descripcion,"
    strSQL = strSQL & " cuentacontable,"
    strSQL = strSQL & " monto,"
    strSQL = strSQL & " impuesto1,"
    strSQL = strSQL & " impuesto2,"
    strSQL = strSQL & " esperiodico,"
    strSQL = strSQL & " FacORec,"
    '23/03/2006
    strSQL = strSQL & " RequiereUsuario,"
    strSQL = strSQL & " RequiereInstructor,"
    strSQL = strSQL & " Grupo,"
    '01/12/2006
    strSQL = strSQL & " NumeroCupones,"
    strSQL = strSQL & " DiasVigenciaCupones,"
    strSQL = strSQL & " DescripcionCupon,"
    strSQL = strSQL & " ObservacionesCupon,"
    '16/06/2008
    strSQL = strSQL & " ImporteAPagar)"
    strSQL = strSQL & " VALUES ("
    strSQL = strSQL & Val(lblClave.Caption) & ","
    strSQL = strSQL & "'" & Trim(txtIngresos(1)) & "',"
    strSQL = strSQL & "'" & Trim(txtIngresos(2)) & "',"
    strSQL = strSQL & Val(Trim(txtIngresos(3))) & ","
    strSQL = strSQL & Val(Trim(txtIngresos(4))) & ","
    strSQL = strSQL & Val(Trim(txtIngresos(5))) & ","
    strSQL = strSQL & chkPeriodico.Value & ","
    strSQL = strSQL & "'" & Left$(Me.cmbFacORec.Text, 1) & "',"
    strSQL = strSQL & Me.chkReqUsu.Value & ","
    strSQL = strSQL & Me.chkReqIns.Value & ","
    strSQL = strSQL & "'" & UCase(Trim(Me.txtIngresos(6).Text)) & "',"
    '1/12/2006
    strSQL = strSQL & Val(Trim(Me.txtIngresos(7).Text)) & ","
    strSQL = strSQL & Val(Trim(Me.txtIngresos(8).Text)) & ","
    strSQL = strSQL & "'" & Trim(Me.txtIngresos(9).Text) & "',"
    strSQL = strSQL & "'" & Trim(Me.txtIngresos(10).Text) & "',"
    '16/06/2008
    strSQL = strSQL & Val(Trim(Me.txtIngresos(11).Text)) & ")"
    
    
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    Conn.CommitTrans      'Termina transacción
    Screen.MousePointer = vbDefault
    MsgBox "¡Registro Actualizado!"
    Unload Me
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub









Private Sub txtIngresos_GotFocus(Index As Integer)
    txtIngresos(Index).SelStart = 0
    txtIngresos(Index).SelLength = Len(txtIngresos(Index))
End Sub

Private Sub txtIngresos_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 3, 4, 5, 7, 8, 11
            Select Case KeyAscii
                Case 8, 22, 46, 48 To 57    'Backspace, <Ctrl+V>, punto y del 0 al 9
                    KeyAscii = KeyAscii
                Case Else
                    KeyAscii = 0
                End Select
        Case 1, 2, 6, 9, 10
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Sub Limpia()
    Dim i As Integer
    For i = 0 To 5
        txtIngresos(i).Text = ""
    Next i
    chkPeriodico.Value = False
    Call Llena_txtIngresos
    
    
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM concepto_ingresos WHERE idconcepto = " & _
                  Val(Trim(frmCatalogos.lblModo.Caption))
    Set AdoRcsIngresos = New ADODB.Recordset
    AdoRcsIngresos.ActiveConnection = Conn
    AdoRcsIngresos.LockType = adLockOptimistic
    AdoRcsIngresos.CursorType = adOpenKeyset
    AdoRcsIngresos.CursorLocation = adUseServer
    AdoRcsIngresos.Open strSQL
    If Not AdoRcsIngresos.EOF Then
        lblClave.Caption = AdoRcsIngresos!IdConcepto
        Me.txtIngresos(0) = AdoRcsIngresos!IdConcepto
        txtIngresos(1).Text = AdoRcsIngresos!Descripcion
        txtIngresos(2).Text = IIf(IsNull(AdoRcsIngresos!cuentacontable), vbNullString, AdoRcsIngresos!cuentacontable)
        txtIngresos(3).Text = AdoRcsIngresos!Monto
        txtIngresos(4).Text = AdoRcsIngresos!Impuesto1
        txtIngresos(5).Text = AdoRcsIngresos!impuesto2
        chkPeriodico.Value = AdoRcsIngresos!esperiodico * -1
        If AdoRcsIngresos!FacORec = "F" Then
            Me.cmbFacORec.Text = "FACTURA"
        Else
            Me.cmbFacORec.Text = "RECIBO"
        End If
        
        '23/03/2006
        Me.chkReqUsu.Value = AdoRcsIngresos!RequiereUsuario * -1
        Me.chkReqIns.Value = AdoRcsIngresos!RequiereInstructor * -1
        
        Me.txtIngresos(6).Text = Trim(IIf(IsNull(AdoRcsIngresos!Grupo), "", AdoRcsIngresos!Grupo))
        
        '1/12/2006
        Me.txtIngresos(7).Text = IIf(IsNull(AdoRcsIngresos!NumeroCupones), 0, AdoRcsIngresos!NumeroCupones)
        Me.txtIngresos(8).Text = IIf(IsNull(AdoRcsIngresos!DiasVigenciaCupones), 0, AdoRcsIngresos!DiasVigenciaCupones)
        Me.txtIngresos(9).Text = IIf(IsNull(AdoRcsIngresos!DescripcionCupon), "", AdoRcsIngresos!DescripcionCupon)
        Me.txtIngresos(10).Text = IIf(IsNull(AdoRcsIngresos!ObservacionesCupon), "", AdoRcsIngresos!ObservacionesCupon)
        
        '16/6/2008
        Me.txtIngresos(11).Text = IIf(IsNull(AdoRcsIngresos!ImporteaPagar), 0, AdoRcsIngresos!ImporteaPagar)
        
    End If
End Sub

Sub Llena_txtIngresos()
    Dim lngAnterior, lngIngresos As Long
    'Llena txtingresos
    lngIngresos = 1
    strSQL = "SELECT idconcepto FROM concepto_ingresos ORDER BY idconcepto"
    Set AdoRcsIngresos = New ADODB.Recordset
    AdoRcsIngresos.ActiveConnection = Conn
    AdoRcsIngresos.LockType = adLockOptimistic
    AdoRcsIngresos.CursorType = adOpenKeyset
    AdoRcsIngresos.CursorLocation = adUseServer
    AdoRcsIngresos.Open strSQL
    If AdoRcsIngresos.EOF Then
        lngIngresos = 1
        txtIngresos(0).Text = lngIngresos
        Exit Sub
    End If
    AdoRcsIngresos.MoveFirst
    Do While Not AdoRcsIngresos.EOF
        If AdoRcsIngresos.Fields!IdConcepto <> "1" Then
            If Val(AdoRcsIngresos.Fields!IdConcepto) - lngAnterior > 1 Then
                Exit Do
            End If
        End If
        lngAnterior = lngIngresos
        AdoRcsIngresos.MoveNext
        If Not AdoRcsIngresos.EOF Then lngIngresos = AdoRcsIngresos.Fields!IdConcepto
    Loop
    txtIngresos(0).Text = lngAnterior + 1
    
    
    Me.txtIngresos(7).Text = 0
    Me.txtIngresos(8).Text = 1
    
    
End Sub


Private Sub LlenaComboFacORec()
    Me.cmbFacORec.Clear
    
    Me.cmbFacORec.AddItem "FACTURA"
    Me.cmbFacORec.AddItem "RECIBO"
    
End Sub
