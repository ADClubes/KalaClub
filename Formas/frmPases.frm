VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPases 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de pases"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   Icon            =   "frmPases.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmDir 
      Height          =   3135
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   6720
      Begin VB.CommandButton cmdHZonas 
         Height          =   305
         Left            =   960
         Picture         =   "frmPases.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   " Lista de zonas hábiles en las instalaciones "
         Top             =   2640
         Width           =   425
      End
      Begin VB.TextBox txtZona 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   2640
         Width           =   4935
      End
      Begin VB.TextBox txtCveZona 
         Height          =   285
         Left            =   240
         MaxLength       =   4
         TabIndex        =   7
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtReg 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   19
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtSec 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtMotivo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1920
         Width           =   4935
      End
      Begin VB.CommandButton cmdHCausas 
         Height          =   305
         Left            =   960
         Picture         =   "frmPases.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Lista de opciones "
         Top             =   1920
         Width           =   425
      End
      Begin VB.TextBox txtCve 
         Height          =   285
         Left            =   240
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1920
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpTermina 
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   58654721
         CurrentDate     =   38190
      End
      Begin MSComCtl2.DTPicker dtpInicia 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   58654721
         CurrentDate     =   38190
      End
      Begin VB.TextBox txtAQuien 
         Height          =   285
         Left            =   240
         MaxLength       =   60
         TabIndex        =   0
         Top             =   480
         Width           =   6255
      End
      Begin VB.CommandButton cmdSalir 
         Height          =   615
         Left            =   5880
         Picture         =   "frmPases.frx":06D6
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   " Salir "
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cmdGuardar 
         Height          =   615
         Left            =   5040
         Picture         =   "frmPases.frx":09E0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   " Guardar registro "
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblZona 
         Caption         =   "Zona"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblCveZona 
         Caption         =   "Cve"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label lblReg 
         Caption         =   "# Reg."
         Height          =   255
         Left            =   4080
         TabIndex        =   20
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblSec 
         Caption         =   "# Sec."
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblMotivo 
         Caption         =   "Motivo"
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label lblCve 
         Caption         =   "Cve"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label lblTermina 
         Caption         =   "Termina"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1920
         TabIndex        =   15
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblAQuien 
         Caption         =   "¿A quien se otorga?"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblInicia 
         Caption         =   "Inicia"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmPases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************
'*  Formulario para el registro de pases temporales             *
'*  Daniel Hdez                                                 *
'*  27 / Septiembre / 2004                                      *
'*  Ultima actualización: 16 / Agosto / 2005                        *
'****************************************************************


Public bNvoPase As Boolean

Dim sAQuien As String
Dim dInicia As Date
Dim dTermina As Date
Dim nMotivo As Integer
Dim nZona As Integer
Dim frmHCausas As frmayuda
Dim frmHZonas As frmayuda
Public nAyuda As Byte


Private Sub cmdSalir_Click()
Dim Respuesta As Integer

    If (Cambios) Then
        Respuesta = MsgBox("¿Desea guardar los datos antes de salir?", vbYesNo, "Registro de pases")

        If (Respuesta = vbYes) Then
            If (GuardaDatos) Then
                Unload Me
            Else
                Exit Sub
            End If
        End If
    End If
    
    Unload Me
End Sub


Private Sub Form_Load()
    CentraForma MDIPrincipal, frmPases
    sTextToolBar = Trim(MDIPrincipal.StatusBar1.Panels.Item(1).Text)
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = "Pases temporales"

    If (bNvoPase) Then
        ClearCtrls
    Else
        LeePase
    End If
    
    InitVar
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmPases.LlenaPases
    MDIPrincipal.StatusBar1.Panels.Item(1).Text = sTextToolBar
End Sub


Private Sub ClearCtrls()
    With Me
        .txtAQuien.Text = ""
        .dtpInicia.Value = Format(Date, "dd/mm/yyyy")
        .dtpTermina.Value = Format(Date, "dd/mm/yyyy")
        .txtSec.Text = ""
        .txtCve.Text = ""
        .txtMotivo.Text = ""
        .txtCveZona.Text = ""
        .txtZona.Text = ""
    End With
End Sub


Private Sub InitVar()
    sAQuien = Trim(Me.txtAQuien.Text)
    dInicia = Me.dtpInicia.Value
    dTermina = Me.dtpTermina.Value
    nMotivo = Val(Me.txtCve.Text)
    nZona = Val(Me.txtCveZona.Text)
End Sub


Private Sub LeePase()
Dim rsPase As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String


    sCampos = "Pases_Temporales!IdPase, Pases_Temporales!Secuencial, "
    sCampos = sCampos & "Pases_Temporales!FechaInicio, Pases_Temporales!FechaFinal, "
    sCampos = sCampos & "Pases_Temporales!QuienRecibe, Causas_Pase!Descripcion, "
    sCampos = sCampos & "Pases_Temporales!IdCausa, Pases_Temporales!IdTimeZone, "
    sCampos = sCampos & "Time_Zone!Descripcion "
    
    sTablas = "((Pases_Temporales LEFT JOIN Secuencial ON Pases_Temporales.Secuencial=Secuencial.Secuencial) "
    sTablas = sTablas & "LEFT JOIN Causas_Pase ON Pases_Temporales.IdCausa=Causas_Pase.IdCausa) "
    sTablas = sTablas & "LEFT JOIN Time_Zone ON Pases_Temporales.IdTimeZone=Time_Zone.IdTimeZone "
    
    InitRecordSet rsPase, sCampos, sTablas, "Pases_Temporales.idPase=" & frmOtrosDatos.ssdbPases.Columns(0).Value, "", Conn
    With rsPase
        If (.RecordCount > 0) Then
            .MoveFirst
            Do While (Not .EOF)
                Me.txtReg.Text = .Fields("Pases_Temporales!IdPase")
                frmPases.txtSec.Text = .Fields("Pases_Temporales!Secuencial")
                
                If (.Fields("Pases_Temporales!QuienRecibe") <> "") Then
                    frmPases.txtAQuien.Text = .Fields("Pases_Temporales!QuienRecibe")
                End If
                
                frmPases.dtpInicia.Value = .Fields("Pases_Temporales!FechaInicio")
                frmPases.dtpTermina.Value = .Fields("Pases_Temporales!FechaFinal")
                
                If (.Fields("Pases_Temporales!IdCausa") <> 0) Then
                    frmPases.txtCve.Text = .Fields("Pases_Temporales!IdCausa")
                End If
        
                If (.Fields("Causas_Pase!Descripcion") <> "") Then
                    frmPases.txtMotivo.Text = .Fields("Causas_Pase!Descripcion")
                End If
                
                If (.Fields("Pases_Temporales!IdTimeZone") <> "") Then
                    frmPases.txtCveZona.Text = .Fields("Pases_Temporales!IdTimeZone")
                End If
                
                If (.Fields("Time_Zone!Descripcion") <> "") Then
                    frmPases.txtZona.Text = .Fields("Time_Zone!Descripcion")
                End If
                
                .MoveNext
            Loop
        End If
        
        .Close
    End With
    Set rsPase = Nothing
End Sub


Private Sub cmdGuardar_Click()
Dim i As Byte

    If (Cambios) Then
        If (GuardaDatos) Then
            InitVar
        Else
            MsgBox "No se registraron los datos, verifique la información.", vbCritical, "KalaSystems"
        End If
    End If
    
    Me.txtAQuien.SetFocus
End Sub


Private Function ChecaDatos()
Dim sCond As String
Dim sTablas As String
Dim nTitular As Integer

    ChecaDatos = False
    
    If (Trim(Me.txtAQuien.Text) = "") Then
        MsgBox "Se debe escribir el nombre de la persona a quien se otorga el pase", vbExclamation, "KalaSystems"
        Me.txtAQuien.SetFocus
        Exit Function
    End If
    
    If (Trim(Me.txtCve.Text) = "") Then
        MsgBox "La causa por la cual se otorga el pase no puede quedar vacía.", vbExclamation, "KalaSystems"
        Me.txtCve.SetFocus
        Exit Function
    End If

    If (Trim(Me.txtMotivo.Text) = "") Then
        MsgBox "Es necesario especificar la causa por la cual se otorga el pase.", vbExclamation, "KalaSystems"
        Me.txtCve.SetFocus
        Exit Function
    Else
        nMotivo = LeeXValor("IdCausa", "Causas_Pase", "Descripcion='" & Trim(Me.txtMotivo.Text) & "'", "IdCausa", "n", Conn)

        If (nMotivo <= 0) Then
            MsgBox "El motivo especificado no está registrado.", vbExclamation, "KalaSystems"
            Me.txtCve.SetFocus
            Exit Function
        End If
    End If
    
    If (Trim(Me.txtCveZona.Text) = "") Then
        MsgBox "La zona para la cual se otorga el pase no puede quedar vacía.", vbExclamation, "KalaSystems"
        Me.txtCveZona.SetFocus
        Exit Function
    End If

    If (Trim(Me.txtZona.Text) = "") Then
        MsgBox "Es necesario especificar la zona para la cual se otorga el pase.", vbExclamation, "KalaSystems"
        Me.txtCveZona.SetFocus
        Exit Function
    Else
        nZona = LeeXValor("IdTimeZone", "Time_Zone", "Descripcion='" & Trim(Me.txtZona.Text) & "'", "IdTimeZone", "n", Conn)

        If (nZona <= 0) Then
            MsgBox "La zona especificada no está registrada.", vbExclamation, "KalaSystems"
            Me.txtCveZona.SetFocus
            Exit Function
        End If
    End If
    
    If (Me.dtpTermina.Value < Me.dtpInicia.Value) Then
        MsgBox "La fecha final debe ser mayor o igual a la inicial.", vbExclamation, "KalaSystems"
        Me.dtpInicia.SetFocus
        Exit Function
    End If

    ChecaDatos = True
End Function


Private Function Cambios() As Boolean
    Cambios = True
    
    If (sAQuien <> Trim(Me.txtAQuien.Text)) Then
        Exit Function
    End If

    If (dInicia <> Me.dtpInicia.Value) Then
        Exit Function
    End If
    
    If (dTermina <> Me.dtpTermina.Value) Then
        Exit Function
    End If
    
    If (nMotivo <> Val(Me.txtCve.Text)) Then
        Exit Function
    End If
    
    If (nZona <> Val(Me.txtCveZona.Text)) Then
        Exit Function
    End If
    
    Cambios = False
End Function


Private Function GuardaDatos() As Boolean
Const DATOSPASE = 7
Dim bCreado As Boolean
Dim mFieldsPase(DATOSPASE) As String
Dim mValuesPase(DATOSPASE) As Variant
Dim sCond As String
Dim nInitTrans As Long

    If (Not ChecaDatos) Then
        GuardaDatos = False
        Exit Function
    End If

    'Campos de la tabla Pases_Temporales
    mFieldsPase(0) = "IdPase"
    mFieldsPase(1) = "Secuencial"
    mFieldsPase(2) = "QuienRecibe"
    mFieldsPase(3) = "IdCausa"
    mFieldsPase(4) = "FechaInicio"
    mFieldsPase(5) = "FechaFinal"
    mFieldsPase(6) = "IdTimeZone"

    'Valores de la tabla Rentables
    #If SqlServer_ Then
        mValuesPase(2) = Trim(UCase(Me.txtAQuien.Text))
        mValuesPase(3) = Val(Me.txtCve.Text)
        mValuesPase(4) = Format(Me.dtpInicia.Value, "yyyymmdd")
        mValuesPase(5) = Format(Me.dtpTermina.Value, "yyyymmdd")
        mValuesPase(6) = Val(Me.txtCveZona.Text)
    #Else
        mValuesPase(2) = Trim(UCase(Me.txtAQuien.Text))
        mValuesPase(3) = Val(Me.txtCve.Text)
        mValuesPase(4) = Format(Me.dtpInicia.Value, "dd/mm/yyyy")
        mValuesPase(5) = Format(Me.dtpTermina.Value, "dd/mm/yyyy")
        mValuesPase(6) = Val(Me.txtCveZona.Text)
    #End If
    
    If (bNvoPase) Then
        nInitTrans = Conn.BeginTrans
        
        'Asigna el Id del siguiente registro
        mValuesPase(0) = LeeUltReg("Pases_Temporales", "IdPase") + 1
        
        'Asigna el # secuencial para el pase temporal
        mValuesPase(1) = AsignaSec(Val(frmOtrosDatos.txtTitCve.Text), True)
        
        If (mValuesPase(1) > 0) Then
            If (AgregaRegistro("Pases_Temporales", mFieldsPase, DATOSPASE, mValuesPase, Conn)) Then
                'Activa el pase
                #If SqlServer_ Then
                    ActivaCredSQL 0, CLng(mValuesPase(1)), 0, Val(frmOtrosDatos.txtTitCve.Text), True, False
                #Else
                    ActivaCred 0, CLng(mValuesPase(1)), 0, Val(frmOtrosDatos.txtTitCve.Text), True, False
                #End If
                
                MsgBox "Los datos se dieron de alta correctamente.", vbInformation, "KalaSystems"
                        
                'Actualiza el numero de secuencial
                Me.txtSec.Text = mValuesPase(1)
                Me.txtSec.REFRESH
                
                'Actualiza el numero de registro
                Me.txtReg.Text = mValuesPase(0)
                Me.txtReg.REFRESH
    
                bNvoPase = False
                GuardaDatos = True
                
                'Baja a disco los nuevos datos
                Conn.CommitTrans
            Else
                'En caso de algun error no baja a disco los nuevos datos
                
                If InitTrans > 0 Then
                    Conn.RollbackTrans
                End If
                
                MsgBox "No se reservó el pase, intentelo nuevamente.", vbCritical, "KalaSystems"
            End If
        Else
            'En caso de algun error no baja a disco los nuevos datos
            If InitTrans > 0 Then
                Conn.RollbackTrans
            End If
            
            MsgBox "No se asignó el pase, intentelo nuevamente.", vbCritical, "KalaSystems"
        End If
    Else
        'Lee el Id del registro que se va a modificar
        mValuesPase(0) = Val(Me.txtReg.Text)
        
        'Lee el # secuencial ya asignado
        mValuesPase(1) = Val(Me.txtSec.Text)
        
        sCond = "IdPase=" & mValuesPase(0)
    
        'Asigna la clave del titular al art rentable
        If (CambiaReg("Pases_Temporales", mFieldsPase, DATOSPASE, mValuesPase, sCond, Conn)) Then
            MsgBox "El pase fué actualizado.", vbInformation, "KalaSystems"

            GuardaDatos = True
        Else
            MsgBox "El pase no se registró.", vbCritical, "KalaSystems"
        End If
    End If
End Function


Private Sub txtCve_LostFocus()
    If (Me.txtCve.Text <> "") Then
        If (IsNumeric(Me.txtCve.Text)) Then
            Me.txtMotivo.Text = LeeXValor("Descripcion", "Causas_Pase", "IdCausa=" & Val(Me.txtCve.Text), "Descripcion", "s", Conn)
            If (Trim(Me.txtMotivo.Text) = "VACIO") Then
                MsgBox "La causa seleccionada no existe.", vbCritical, "KalaSystems"
                Me.txtMotivo.Text = ""
                Me.txtCve.Text = ""
                Me.txtCve.SetFocus
            End If
        Else
            MsgBox "La clave del motivo para el pase es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtMotivo.Text = ""
            Me.txtCve.Text = ""
            Me.txtCve.SetFocus
        End If
    End If
    
    Me.txtCve.REFRESH
End Sub


Private Sub txtCveZona_LostFocus()
    If (Me.txtCveZona.Text <> "") Then
        If (IsNumeric(Me.txtCveZona.Text)) Then
            Me.txtZona.Text = LeeXValor("Descripcion", "Time_Zone", "IdTimeZone=" & Val(Me.txtCveZona.Text), "Descripcion", "s", Conn)
            If (Trim(Me.txtZona.Text) = "VACIO") Then
                MsgBox "La zona seleccionada no existe.", vbCritical, "KalaSystems"
                Me.txtZona.Text = ""
                Me.txtCveZona.Text = ""
                Me.txtCveZona.SetFocus
            End If
        Else
            MsgBox "La clave de la zona para el pase es incorrecta.", vbExclamation, "KalaSystems"
            Me.txtZona.Text = ""
            Me.txtCveZona.Text = ""
            Me.txtCveZona.SetFocus
        End If
    End If
    
    Me.txtCveZona.REFRESH
End Sub


Public Sub QuitaPase()
Const DATOSPASE = 2
Dim mFieldsPase(DATOSPASE) As String
Dim mValuesPase(DATOSPASE) As Variant
Dim sCond As String
Dim InitTrans As Long

    'Campos de la tabla Pases temporales
    mFieldsPase(0) = "IdPase"

    'Valores de la tabla Rentables
    mValuesPase(0) = 0

    With frmOtrosDatos.ssdbPases
        InitTrans = Conn.BeginTrans
    
        'Elimina el registro del pase
        If (EliminaReg("Pases_Temporales", "IdPase=" & .Columns(0).Value, "", Conn)) Then
            'Condicion para cambiar los valores del registro
            sCond = "Secuencial=" & .Columns(1).Value
            
            'Campos de la tabla secuencial
            mFieldsPase(0) = "IdMember"
            mFieldsPase(1) = "Temporal"
            
            'Valores para la tabla Secuencial
            mValuesPase(0) = 0
            mValuesPase(1) = 0
            
            If (CambiaReg("Secuencial", mFieldsPase, DATOSPASE, mValuesPase, sCond, Conn)) Then
                'Desactiva el pase
                #If SqlServer_ Then
                    ActivaCredSQL 0, .Columns(1).Value, 0, Val(frmOtrosDatos.txtTitCve.Text), False, False
                #Else
                    ActivaCred 0, .Columns(1).Value, 0, Val(frmOtrosDatos.txtTitCve.Text), False, False
                #End If
                
                Conn.CommitTrans
        
                MsgBox "El pase fué dado de baja.", vbInformation, "KalaSystems"
            End If
        Else
            'En caso de algun error no baja a disco los nuevos datos
            If InitTrans > 0 Then
                Conn.RollbackTrans
            End If
            
            MsgBox "No se dió de baja el pase.", vbCritical, "KalaSystems"
        End If
    End With
End Sub


Public Sub LlenaPases()
Const DATOSPASE = 9
Dim rsPase As ADODB.Recordset
Dim sCampos As String
Dim sTablas As String
Dim mAncPase(DATOSPASE) As Integer
Dim mEncPase(DATOSPASE) As String


    frmOtrosDatos.ssdbPases.RemoveAll

    sCampos = "Pases_Temporales!IdPase, Pases_Temporales!Secuencial, "
    sCampos = sCampos & "Pases_Temporales!FechaInicio, Pases_Temporales!FechaFinal, "
    sCampos = sCampos & "Pases_Temporales!QuienRecibe, Causas_Pase!Descripcion, "
    sCampos = sCampos & "Pases_Temporales!IdCausa, Pases_Temporales!IdTimeZone, "
    sCampos = sCampos & "Time_Zone!Descripcion "
    
    sTablas = "((Pases_Temporales LEFT JOIN Secuencial ON Pases_Temporales.Secuencial=Secuencial.Secuencial) "
    sTablas = sTablas & "LEFT JOIN Causas_Pase ON Pases_Temporales.IdCausa=Causas_Pase.IdCausa) "
    sTablas = sTablas & "LEFT JOIN Time_Zone ON Pases_Temporales.IdTimeZone=Time_Zone.IdTimeZone "
    
    InitRecordSet rsPase, sCampos, sTablas, "Secuencial.idMember=" & Val(frmOtrosDatos.txtTitCve.Text), "", Conn
    With rsPase
        If (.RecordCount > 0) Then
            .MoveFirst
            
            Do While (Not .EOF)
                frmOtrosDatos.ssdbPases.AddItem .Fields("Pases_Temporales!idPase") & vbTab & _
                .Fields("Pases_Temporales!Secuencial") & vbTab & _
                Format(.Fields("Pases_Temporales!FechaInicio"), "dd / mmm / yyyy") & vbTab & _
                Format(.Fields("Pases_Temporales!FechaFinal"), "dd / mmm / yyyy") & vbTab & _
                .Fields("Pases_Temporales!QuienRecibe") & vbTab & _
                .Fields("Causas_Pase!Descripcion") & vbTab & _
                .Fields("Pases_Temporales!idCausa") & vbTab & _
                .Fields("Pases_Temporales!idTimeZone") & vbTab & _
                .Fields("Time_Zone!Descripcion")
            
                .MoveNext
            Loop
        End If
        
        .Close
    End With
    Set rsPase = Nothing
    
    'Asigna valores a la matriz de encabezados
    mEncPase(0) = "# Reg."
    mEncPase(1) = "# Sec."
    mEncPase(2) = "Inició"
    mEncPase(3) = "Vence"
    mEncPase(4) = "Se otorgó a"
    mEncPase(5) = "Causa o motivo"
    mEncPase(6) = "# Causa"
    mEncPase(7) = "# Zona"
    mEncPase(8) = "Zona"

    'Asigna los encabezados de las columnas
    DefHeaderssGrid frmOtrosDatos.ssdbPases, mEncPase
    
    'Asigna valores a la matriz que define el ancho de cada columna
    mAncPase(0) = 800
    mAncPase(1) = 900
    mAncPase(2) = 1500
    mAncPase(3) = 1500
    mAncPase(4) = 4300
    mAncPase(5) = 2500
    mAncPase(6) = 900
    mAncPase(7) = 900
    mAncPase(8) = 2500

    'Asigna el ancho de cada columna
    DefAnchossGrid frmOtrosDatos.ssdbPases, mAncPase
    
    frmOtrosDatos.ssdbPases.Columns(0).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbPases.Columns(1).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbPases.Columns(2).Alignment = ssCaptionAlignmentCenter
    frmOtrosDatos.ssdbPases.Columns(3).Alignment = ssCaptionAlignmentCenter
    frmOtrosDatos.ssdbPases.Columns(6).Alignment = ssCaptionAlignmentRight
    frmOtrosDatos.ssdbPases.Columns(7).Alignment = ssCaptionAlignmentRight
End Sub



'************************************************************
'*                          Ayudas                          *
'************************************************************

Private Sub cmdHCausas_Click()
Const DATOSCAUSA = 2
Dim sCadena As String
Dim mFAyuda(DATOSCAUSA) As String
Dim mAAyuda(DATOSCAUSA) As Integer
Dim mCAyuda(DATOSCAUSA) As String
Dim mEAyuda(DATOSCAUSA) As String

    nAyuda = 1

    Set frmHCausas = New frmayuda
    
    mFAyuda(0) = "Causas ordenadas por clave"
    mFAyuda(1) = "Causas ordenadas por descripción"
    
    mAAyuda(0) = 800
    mAAyuda(1) = 3500
    
    mCAyuda(0) = "IdCausa"
    mCAyuda(1) = "Descripcion"
    
    mEAyuda(0) = "Clave"
    mEAyuda(1) = "Descripción"
    
    With frmHCausas
        .nColActiva = 0
        .nColsAyuda = DATOSCAUSA
        .sTabla = "Causas_Pase"
        
        .sCondicion = ""
        .sTitAyuda = "Causas para pases temporales"
        .lAgregar = True
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
    If (Trim(Me.txtCve.Text) <> "") Then
        Me.txtMotivo.Text = LeeXValor("Descripcion", "Causas_Pase", "IdCausa=" & Val(Me.txtCve.Text), "Descripcion", "s", Conn)
    End If
    
    Me.cmdHCausas.SetFocus
End Sub


Private Sub cmdHZonas_Click()
Const DATOSZONAS = 2
Dim sCadena As String
Dim mFAyuda(DATOSZONAS) As String
Dim mAAyuda(DATOSZONAS) As Integer
Dim mCAyuda(DATOSZONAS) As String
Dim mEAyuda(DATOSZONAS) As String

    nAyuda = 2

    Set frmHZonas = New frmayuda
    
    mFAyuda(0) = "Zonas ordenadas por clave"
    mFAyuda(1) = "Zonas ordenadas por descripción"
    
    mAAyuda(0) = 800
    mAAyuda(1) = 3500
    
    mCAyuda(0) = "IdTimeZone"
    mCAyuda(1) = "Descripcion"
    
    mEAyuda(0) = "Clave"
    mEAyuda(1) = "Descripción"
    
    With frmHZonas
        .nColActiva = 0
        .nColsAyuda = DATOSZONAS
        .sTabla = "Time_Zone"
        
        .sCondicion = ""
        .sTitAyuda = "Zonas para pases temporales"
        .lAgregar = True
        
        .ConfigAyuda mFAyuda, mAAyuda, mCAyuda, mEAyuda
        
        .Show (1)
    End With
    
    If (Trim(Me.txtCveZona.Text) <> "") Then
        Me.txtZona.Text = LeeXValor("Descripcion", "Time_Zone", "IdTimeZone=" & Val(Me.txtCveZona.Text), "Descripcion", "s", Conn)
    End If
    
    Me.cmdHZonas.SetFocus
End Sub

