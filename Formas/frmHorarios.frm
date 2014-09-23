VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHorarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horarios (Captura)"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4845
   Icon            =   "frmHorarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4845
   Begin VB.ComboBox cboClase 
      Height          =   315
      Left            =   200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   930
      Width           =   4500
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Height          =   840
      Left            =   3540
      Picture         =   "frmHorarios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Salir"
      Top             =   4380
      Width           =   795
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFC0&
      Height          =   840
      Left            =   2370
      Picture         =   "frmHorarios.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Guardar"
      Top             =   4380
      Width           =   795
   End
   Begin VB.Frame frmHoras 
      Caption         =   "Horario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   2010
      TabIndex        =   18
      Top             =   2160
      Width           =   2655
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   330
         Left            =   885
         TabIndex        =   12
         Top             =   555
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         Format          =   57081858
         CurrentDate     =   38047
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   330
         Left            =   885
         TabIndex        =   14
         Top             =   1215
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "24:59"
         Format          =   57081858
         CurrentDate     =   38047
      End
      Begin VB.Label lblHasta 
         Alignment       =   1  'Right Justify
         Caption         =   "&A:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   300
         TabIndex        =   13
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label lblDesde 
         Alignment       =   1  'Right Justify
         Caption         =   "D&e:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   330
         TabIndex        =   11
         Top             =   600
         Width           =   390
      End
   End
   Begin VB.Frame frmDias 
      Caption         =   "Días"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3105
      Left            =   150
      TabIndex        =   17
      Top             =   2160
      Width           =   1635
      Begin VB.CheckBox chkDias 
         Caption         =   "&Sabado"
         Height          =   420
         Index           =   6
         Left            =   330
         TabIndex        =   10
         Top             =   2580
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         Caption         =   "&Viernes"
         Height          =   420
         Index           =   5
         Left            =   330
         TabIndex        =   9
         Top             =   2190
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         Caption         =   "&Jueves"
         Height          =   420
         Index           =   4
         Left            =   330
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         Caption         =   "Mie&rcoles"
         Height          =   420
         Index           =   3
         Left            =   330
         TabIndex        =   7
         Top             =   1425
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         Caption         =   "&Martes"
         Height          =   420
         Index           =   2
         Left            =   330
         TabIndex        =   6
         Top             =   1035
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         Caption         =   "&Lunes"
         Height          =   420
         Index           =   1
         Left            =   330
         TabIndex        =   5
         Top             =   645
         Width           =   975
      End
      Begin VB.CheckBox chkDias 
         Caption         =   "&Domingo"
         Height          =   420
         Index           =   0
         Left            =   330
         TabIndex        =   4
         Top             =   255
         Width           =   975
      End
   End
   Begin VB.ComboBox cboInstructor 
      Height          =   315
      Left            =   200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   4500
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HORARIOS"
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
      Left            =   435
      TabIndex        =   19
      Top             =   180
      Width           =   3975
   End
   Begin VB.Label lblInstructor 
      BackStyle       =   0  'Transparent
      Caption         =   "&Instructor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   200
      TabIndex        =   2
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label lblTipoClase 
      BackStyle       =   0  'Transparent
      Caption         =   "&Clase"
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
      Left            =   200
      TabIndex        =   0
      Top             =   690
      Width           =   630
   End
End
Attribute VB_Name = "frmHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ************************************************************************
' PANTALLA HORARIOS
' Objetivo: CATÁLOGO DE HORARIOS DE CLASES
' Programado por:
' Fecha: FEBRERO DE 2004
' ************************************************************************

Option Explicit
    Dim iniTrans As Long
    Dim AdoRcsHorarios As ADODB.Recordset
    Dim intCveTipoClase, intCveInstructor As Integer
    Dim strCadenaDias As String
    
Private Function VerificaDatos()
    If cboClase.Text = "" Then
        MsgBox "¡ Seleccione el TIPO DE CLASE, Pues No es Opcional !", _
                     vbOKOnly + vbExclamation, "Horarios (Captura)"
        VerificaDatos = False
        cboClase.SetFocus
        Exit Function
    End If
    If cboInstructor.Text = "" Then
        MsgBox "¡ Seleccione al INSTRUCTOR, Pues No es Opcional.", _
                    vbOKOnly + vbExclamation, "Horarios (Captura)"
        VerificaDatos = False
        cboInstructor.SetFocus
        Exit Function
    End If
    If dtpHasta.Value <= dtpDesde.Value Then
        MsgBox "¡ La HORA DE TÉRMINO de la clase, NO puede ser menor o igual a la HORA DE INICIO !", vbOKOnly + vbExclamation, "Horarios (Captura)"
        VerificaDatos = False
        dtpDesde.SetFocus
        Exit Function
    End If
    
    VerificaDatos = True

    'Ahora verificamos que el registro no exista (Si estamos insertando)
    intCveTipoClase = cboClase.ItemData(cboClase.ListIndex)
    intCveInstructor = cboInstructor.ItemData(cboInstructor.ListIndex)
    If frmCatalogos.lblModo.Caption = "A" Then
        strSQL = "SELECT * FROM horarios_clases WHERE (idinstructor = " & intCveInstructor & _
                    ") AND (dias = '" & strCadenaDias & "') AND ((hora_inicio = format('" & _
                    Format(dtpDesde.Value, "hh:mm") & "', 'hh:mm')) OR (hora_fin = format('" & _
                    Format(dtpHasta.Value, "hh:mm") & "','hh:mm')))"
        Set AdoRcsHorarios = New ADODB.Recordset
        AdoRcsHorarios.ActiveConnection = Conn
        AdoRcsHorarios.LockType = adLockOptimistic
        AdoRcsHorarios.CursorType = adOpenKeyset
        AdoRcsHorarios.CursorLocation = adUseServer
        AdoRcsHorarios.Open strSQL
        If Not AdoRcsHorarios.EOF Then
            MsgBox "¡ Por Verifique sus Datos, Pues el Instructor(a): " & cboInstructor.Text & _
                        " ya tiene registrado los Días o un Horario que causaría conflictos con el usted propone !", vbCritical + vbOKOnly, "Horarios"
            AdoRcsHorarios.Close
            VerificaDatos = False
            cboClase.SetFocus
            Exit Function
        Else
            AdoRcsHorarios.Close
            VerificaDatos = True
        End If
    End If
End Function

Private Sub cboClase_LostFocus()
    If (cboClase.Text <> "") And (cboClase.ListIndex < 0) Then
        MsgBox "Seleccione una de las CLASES de la Lista."
        cboClase.SetFocus
    End If
End Sub

Private Sub cboInstructor_LostFocus()
    If (cboInstructor.Text <> "") And (cboInstructor.ListIndex < 0) Then
        MsgBox "Seleccione un INSTRUCTOR de la Lista."
        cboInstructor.SetFocus
    End If
End Sub

Private Sub cmdGuardar_Click()
    Call FormaCadenaDias
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
    Dim strCampo1, strCampo2 As String
    
    frmCatalogos.Enabled = False
    
    'Llena combo de Tipo de Clase
    strSQL = "SELECT idtipoclase, descripcion FROM tipo_clase ORDER BY descripcion"
    strCampo1 = "descripcion"
    strCampo2 = "idtipoclase"
    Call LlenaCombos(cboClase, strSQL, strCampo1, strCampo2)
    
    'Llena combo de Instructor
    strSQL = "SELECT idinstructor, " & _
                    "(apellido_paterno & ' ' & apellido_materno & ' ' & nombre) as nombres " & _
                    "FROM instructores ORDER BY idinstructor"
    strCampo1 = "nombres"
    strCampo2 = "idinstructor"
    Call LlenaCombos(cboInstructor, strSQL, strCampo1, strCampo2)
    
    If frmCatalogos.lblModo.Caption = "A" Then
        cboClase.Enabled = True
    Else
        Call LlenaDatos
        cboClase.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
        frmCatalogos.Enabled = True
End Sub

Private Sub GuardaDatos()
    Dim AdoCmdInserta As ADODB.Command
    On Error GoTo err_Guarda
    iniTrans = Conn.BeginTrans          'Iniciamos transacción
    Screen.MousePointer = vbHourglass
    strSQL = "INSERT INTO horarios_clases (idtipoclase, idinstructor, dias, hora_inicio, " & _
                    "hora_fin) VALUES (" & intCveTipoClase & ", " & intCveInstructor & ", '" & _
                    strCadenaDias & "', '" & Format(dtpDesde.Value, "hh:mm") & "', '" & _
                    Format(dtpHasta.Value, "hh:mm") & "')"
    Set AdoCmdInserta = New ADODB.Command
    AdoCmdInserta.ActiveConnection = Conn
    AdoCmdInserta.CommandText = strSQL
    AdoCmdInserta.Execute
    Screen.MousePointer = vbDefault
    Conn.CommitTrans      'Termina transacción
    MsgBox "¡ Registro Ingresado !"
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
    strSQL = "DELETE FROM horarios_clases WHERE (idtipoclase = " & _
                    Val(frmCatalogos.lblModo.Caption) & ") AND (idinstructor = " & _
                    Val(frmCatalogos.lblModoDescrip.Caption) & ") AND (dias = '" & _
                    frmCatalogos.AdoDcCatal.Recordset.Fields("dias") & _
                    "') AND (hora_inicio = format('" & _
                    Format(frmCatalogos.AdoDcCatal.Recordset.Fields("hora_inicio"), "hh:mm") & _
                    "', 'hh:mm')) AND (hora_fin = format('" & _
                    Format(frmCatalogos.AdoDcCatal.Recordset.Fields("hora_fin"), "hh:mm") & "', 'hh:mm'))"
    Set AdoCmdRemplaza = New ADODB.Command
    AdoCmdRemplaza.ActiveConnection = Conn
    AdoCmdRemplaza.CommandText = strSQL
    AdoCmdRemplaza.Execute

    'Ahora Insertamos el nuevo registro que sustituye al anterior
    strSQL = "INSERT INTO horarios_clases (idtipoclase, idinstructor, dias, hora_inicio, " & _
                    "hora_fin) VALUES (" & intCveTipoClase & ", " & intCveInstructor & ", '" & _
                    strCadenaDias & "', '" & Format(dtpDesde.Value, "hh:mm") & "', '" & _
                    Format(dtpHasta.Value, "hh:mm") & "')"
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

Sub Limpia()
    Dim i As Integer
        
    cboClase.ListIndex = -1
    cboInstructor.ListIndex = -1

    For i = 0 To 6
        chkDias(i).Value = vbUnchecked
    Next i
'    dtpDesde.Value = Now()
'    dtpHasta.Value = Now()
End Sub

Sub FormaCadenaDias()
    Dim i As Integer
    strCadenaDias = ""
    For i = 0 To 6
        strCadenaDias = strCadenaDias & IIf(chkDias(i).Value = vbChecked, "1", "0")
    Next i
End Sub

Sub LlenaDatos()
    Dim i As Integer
    Dim blnDia As Boolean
    strSQL = "SELECT * FROM horarios_clases WHERE (idtipoclase = " & _
                  Val(frmCatalogos.lblModo.Caption) & ") AND (idinstructor = " & _
                  Val(frmCatalogos.lblModoDescrip.Caption) & ") AND (dias = '" & _
                  frmCatalogos.AdoDcCatal.Recordset.Fields("dias") & _
                  "') AND (hora_inicio = format('" & _
                  Format(frmCatalogos.AdoDcCatal.Recordset.Fields("hora_inicio"), "hh:mm") & _
                  "', 'hh:mm')) AND (hora_fin = format('" & _
                  Format(frmCatalogos.AdoDcCatal.Recordset.Fields("hora_fin"), "hh:mm") & "','hh:mm'))"
    Set AdoRcsHorarios = New ADODB.Recordset
    AdoRcsHorarios.ActiveConnection = Conn
    AdoRcsHorarios.LockType = adLockOptimistic
    AdoRcsHorarios.CursorType = adOpenKeyset
    AdoRcsHorarios.CursorLocation = adUseServer
    AdoRcsHorarios.Open strSQL
    If Not AdoRcsHorarios.EOF Then
        If MuestraElementoCombo(cboClase, AdoRcsHorarios!idtipoclase) Then
        End If
        If MuestraElementoCombo(cboInstructor, AdoRcsHorarios!IdInstructor) Then
        End If
        For i = 0 To 6
            If Mid(AdoRcsHorarios!dias, i + 1, 1) = 1 Then
                chkDias(i).Value = vbChecked
            Else
                chkDias(i).Value = vbUnchecked
            End If
        Next i
        dtpDesde.Value = AdoRcsHorarios!hora_inicio
        dtpHasta.Value = AdoRcsHorarios!hora_fin
    End If
End Sub
