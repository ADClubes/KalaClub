VERSION 5.00
Begin VB.Form frmReglasTipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reglas Tipo de Usuario"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6000
   Icon            =   "frmReglasTipo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6000
   Begin VB.ComboBox cboNuevo 
      Height          =   315
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2820
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.Frame frmSexo 
      Caption         =   "Sexo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   225
      TabIndex        =   2
      Top             =   1200
      Width           =   4545
      Begin VB.OptionButton optSexo 
         Caption         =   "Mi&xto"
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
         Left            =   3345
         TabIndex        =   5
         Top             =   255
         Width           =   1065
      End
      Begin VB.OptionButton optSexo 
         Caption         =   "&Femenino"
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
         Left            =   1845
         TabIndex        =   4
         Top             =   255
         Width           =   1140
      End
      Begin VB.OptionButton optSexo 
         Caption         =   "&Masculino"
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
         Left            =   315
         TabIndex        =   3
         Top             =   240
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   5010
      Picture         =   "frmReglasTipo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   2055
      Width           =   795
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      Height          =   840
      Left            =   5010
      Picture         =   "frmReglasTipo.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Guardar"
      Top             =   825
      Width           =   795
   End
   Begin VB.ComboBox cboActual 
      Height          =   315
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   705
      Width           =   3390
   End
   Begin VB.ComboBox cboAccion 
      Height          =   315
      ItemData        =   "frmReglasTipo.frx":0A56
      Left            =   1365
      List            =   "frmReglasTipo.frx":0A66
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2175
      Width           =   2460
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REGLAS"
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
      Left            =   825
      TabIndex        =   13
      Top             =   150
      Width           =   3975
   End
   Begin VB.Label lblReglas 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo &Nuevo:"
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
      Left            =   210
      TabIndex        =   8
      Top             =   2910
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblClave 
      Alignment       =   2  'Center
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
      Left            =   7000
      TabIndex        =   12
      Top             =   585
      Width           =   270
   End
   Begin VB.Label lblReglas 
      BackStyle       =   0  'Transparent
      Caption         =   "A&cción:"
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
      Left            =   210
      TabIndex        =   6
      Top             =   2265
      Width           =   825
   End
   Begin VB.Label lblReglas 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo &Actual:"
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
      Left            =   210
      TabIndex        =   0
      Top             =   810
      Width           =   1125
   End
End
Attribute VB_Name = "frmReglasTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA: REGLAS DEL TIPO DE USUARIO
' Objetivo: CATÁLOGO DE REGLAS PARA LOS USUARIOS
' Programado por:
' Fecha: NOVIEMBRE DE 2003
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim intActual, intNuevo As Integer, intNuevoAnt As Integer
    Dim AdoRcsReglas As ADODB.Recordset
    Dim strSexo As String, strSexoAnt As String, strAccion As String, strAccionAnt As String

Private Function VerificaDatos()
    If cboActual.Text = "" Then
        MsgBox "¡ Favor de Llenar el TIPO ACTUAL, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de Reglas (Captura)"
        VerificaDatos = False
        cboActual.SetFocus
        Exit Function
    End If
    If (optSexo(0).Value = vbUnchecked) And (optSexo(1).Value = vbUnchecked) And (optSexo(2).Value = vbUnchecked) Then
        MsgBox "¡ Favor de Seleccionar el SEXO, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de Reglas (Captura)"
        VerificaDatos = False
        Exit Function
    End If
    If (cboNuevo.Visible = True) And (cboNuevo.Text = "") Then
        MsgBox "¡ Favor de Llenar el TIPO NUEVO, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de Reglas (Captura)"
        VerificaDatos = False
        cboNuevo.SetFocus
        Exit Function
    End If
    If cboAccion.Text = "" Then
        MsgBox "¡ Favor de Indicar la ACCIÓN, Pues No es Opcional !", _
                    vbOKOnly + vbExclamation, "Tipo de Reglas (Captura)"
        VerificaDatos = False
        cboAccion.SetFocus
        Exit Function
    End If
    If cboActual.Text = cboNuevo.Text Then
        MsgBox "¡ El Tipo NUEVO, NO Puede Ser Igual al ACTUAL !", _
                    vbOKOnly + vbExclamation, "Tipo de Reglas (Captura)"
        VerificaDatos = False
        cboNuevo.SetFocus
        Exit Function
    End If
    VerificaDatos = True
    'Ahora verificamos que el registro no exista (Si estamos insertando)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT * FROM reglas_tipo WHERE (sexo = '" & strSexo & "') AND (accion = '" & cboAccion.Text & "')"
        If cboNuevo.Visible = True Then
            strSQL = strSQL & " AND (idtiponuevo = " & cboNuevo.ItemData(cboNuevo.ListIndex) & ")"
        End If
        strSQL = strSQL & " AND (idtipoactual = " & cboActual.ItemData(cboActual.ListIndex) & ")"
        Set AdoRcsReglas = New ADODB.Recordset
        AdoRcsReglas.ActiveConnection = Conn
        AdoRcsReglas.LockType = adLockOptimistic
        AdoRcsReglas.CursorType = adOpenKeyset
        AdoRcsReglas.CursorLocation = adUseServer
        AdoRcsReglas.Open strSQL
        If Not AdoRcsReglas.EOF Then
            MsgBox "Ya Existe Un Registro Con Esos Datos ", vbInformation & _
                        vbOKOnly, "Reglas Tipo de Usuarios"
            AdoRcsReglas.Close
            VerificaDatos = False
            cboActual.SetFocus
            Exit Function
        Else
            AdoRcsReglas.Close
            VerificaDatos = True
        End If
    End If
End Function

Private Sub cboAccion_Click()
    If cboAccion.Text = "CAMBIAR" Then
        lblReglas(2).Visible = True
        cboNuevo.Visible = True
    Else
        cboNuevo.ListIndex = -1
        lblReglas(2).Visible = False
        cboNuevo.Visible = False
    End If
End Sub

Private Sub cmdGuardar_Click()
    Dim blnGuarda As Boolean
    blnGuarda = VerificaDatos
    If blnGuarda = True Then
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
    Dim i As Integer
    Dim strCampo1, strCampo2 As String
    frmCatalogos.Enabled = False
    
    'Llena los combos Tipo Actual y Tipo Nuevo
    strSQL = "SELECT idtipousuario, descripcion FROM tipo_usuario"
    strCampo1 = "descripcion"
    strCampo2 = "idtipousuario"
    Call LlenaCombos(cboActual, strSQL, strCampo1, strCampo2)
    Call LlenaCombos(cboNuevo, strSQL, strCampo1, strCampo2)
    
    If frmCatalogos.lblModo.Caption <> "A" Then
        Call LlenaDatos
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        frmCatalogos.Enabled = True
End Sub

Private Sub GuardaDatos()
    Dim AdoCmdInserta As ADODB.Command
    Dim intGuardaActual, intGuardaNuevo As Integer
    On Error GoTo err_Guarda
    intGuardaActual = cboActual.ItemData(cboActual.ListIndex)
    If cboNuevo.Visible = True Then
        intGuardaNuevo = cboNuevo.ItemData(cboNuevo.ListIndex)
    Else
        intGuardaNuevo = 0
    End If
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    strSQL = "INSERT INTO reglas_tipo (idtipoactual, idtiponuevo, sexo, accion) " & "VALUES (" & _
                    intGuardaActual & ", " & intGuardaNuevo & ", '" & strSexo & "', '" & cboAccion.Text & "')"
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    
    MsgBox "¡ Registro Ingresado !"
    Call Limpia
    frmCatalogos.AdoDcCatal.REFRESH
    frmCatalogos.grdCatalogos.REFRESH
    frmCatalogos.lblTotal.Caption = Format(frmCatalogos.AdoDcCatal.Recordset.RecordCount, "#######")
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
   Dim intGuardaActual, intGuardaNuevo As Integer
    On Error GoTo err_Guarda
    intGuardaActual = cboActual.ItemData(cboActual.ListIndex)
    If cboNuevo.Visible = True Then
        intGuardaNuevo = cboNuevo.ItemData(cboNuevo.ListIndex)
    Else
        intGuardaNuevo = 0
    End If
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    'Eliminamos el registro existente
    strSQL = "DELETE FROM reglas_tipo WHERE (idtipoactual = " & intGuardaActual & _
                    ") AND (idtiponuevo = " & intNuevoAnt & ") AND (sexo = '" & strSexoAnt & _
                    "') AND (accion = '" & strAccionAnt & "')"
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    
    'Ahora Insertamos el nuevo registro que sustituye al anterior
    strSQL = "INSERT INTO reglas_tipo (idtipoactual, idtiponuevo, sexo, accion) " & "VALUES (" & _
                    intGuardaActual & ", " & intGuardaNuevo & ", '" & strSexo & "', '" & cboAccion.Text & "')"
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute
    Conn.CommitTrans      'Termina transacción
    Screen.MousePointer = vbDefault
    MsgBox "¡ Registro Actualizado !"
    frmCatalogos.AdoDcCatal.REFRESH
    frmCatalogos.grdCatalogos.REFRESH
    frmCatalogos.lblTotal.Caption = Format(frmCatalogos.AdoDcCatal.Recordset.RecordCount, "#######")
    Unload Me
    Exit Sub
    
err_Guarda:
    Screen.MousePointer = Default
    If iniTrans > 0 Then
        Conn.RollbackTrans
    End If
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub optSexo_Click(Index As Integer)
    Select Case Index
        Case 0
            strSexo = "M"
        Case 1
            strSexo = "F"
        Case 2
            strSexo = "X"
    End Select
End Sub

Private Sub txtReglas_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 22, 48 To 57 'Backspace, <Ctrl+V> y del 0 al 9
            KeyAscii = KeyAscii
        Case Else
            KeyAscii = 0
        End Select
End Sub

Sub Limpia()
    Dim i As Integer
    cboActual.ListIndex = -1
    For i = 0 To 2
        optSexo(i).Value = False
    Next i
    cboNuevo.ListIndex = -1
    lblReglas(2).Visible = False
    cboNuevo.Visible = False
    cboAccion.ListIndex = -1
End Sub

Sub LlenaDatos()
    strSQL = "SELECT * FROM reglas_tipo WHERE (idtipoactual = " & Val(frmCatalogos.lblModo.Caption) & _
                    ") AND (accion = '" & frmCatalogos.lblModoDescrip.Caption & "')"
    Set AdoRcsReglas = New ADODB.Recordset
    AdoRcsReglas.ActiveConnection = Conn
    AdoRcsReglas.LockType = adLockOptimistic
    AdoRcsReglas.CursorType = adOpenKeyset
    AdoRcsReglas.CursorLocation = adUseServer
    AdoRcsReglas.Open strSQL
    If Not AdoRcsReglas.EOF Then
        intActual = AdoRcsReglas!idtipoactual
        intNuevo = AdoRcsReglas!idtiponuevo
        intNuevoAnt = AdoRcsReglas!idtiponuevo
        If intNuevo <> 0 Then
            lblReglas(2).Visible = True
            cboNuevo.Visible = True
        End If
        strSexo = AdoRcsReglas!sexo
        strSexoAnt = AdoRcsReglas!sexo
        If Not MuestraElementoCombo(cboAccion, AdoRcsReglas!accion) Then
        End If
        strAccionAnt = AdoRcsReglas!accion
    End If
    cboActual.ListIndex = intActual - 1
    cboNuevo.ListIndex = intNuevo - 1
    Select Case strSexo
        Case "M"
            optSexo(0).Value = True
        Case "F"
            optSexo(1).Value = True
        Case "X"
            optSexo(2).Value = True
    End Select
End Sub
