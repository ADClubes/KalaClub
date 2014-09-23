VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmTurnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Turnos"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   Icon            =   "frmTurnos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8640
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Salir"
      Height          =   615
      Left            =   7320
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos consulta"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8175
      Begin VB.CheckBox chkSoloAbiertos 
         Caption         =   "Solo Turnos Abiertos"
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   6600
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFechaFin 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   63307777
         CurrentDate     =   38920
      End
      Begin MSComCtl2.DTPicker dtpFechaIni 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   63307777
         CurrentDate     =   38920
      End
      Begin VB.Label Label2 
         Caption         =   "Al"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Del"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir Turno"
      Height          =   615
      Left            =   5760
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgTurnos 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8295
      _Version        =   196616
      DataMode        =   2
      AllowUpdate     =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   9
      Columns(0).Width=   1244
      Columns(0).Caption=   "IdTurno"
      Columns(0).Name =   "IdTurno"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2170
      Columns(1).Caption=   "FechaApertura"
      Columns(1).Name =   "FechaApertura"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1984
      Columns(2).Caption=   "HoraApertura"
      Columns(2).Name =   "HoraApertura"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1535
      Columns(3).Caption=   "Caja"
      Columns(3).Name =   "Caja"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1111
      Columns(4).Caption=   "Turno"
      Columns(4).Name =   "NoTurno"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2487
      Columns(5).Caption=   "UsuarioApertura"
      Columns(5).Name =   "UsuarioApertura"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   2143
      Columns(6).Caption=   "FechaCierre"
      Columns(6).Name =   "FechaCierre"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(7).Width=   1826
      Columns(7).Caption=   "HoraCierre"
      Columns(7).Name =   "HoraCierre"
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(8).Width=   3200
      Columns(8).Caption=   "UsuarioCierre"
      Columns(8).Name =   "UsuarioCierre"
      Columns(8).DataField=   "Column 8"
      Columns(8).DataType=   8
      Columns(8).FieldLen=   256
      _ExtentX        =   14631
      _ExtentY        =   4683
      _StockProps     =   79
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
End
Attribute VB_Name = "frmTurnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkSoloAbiertos_Click()
    ActualizaGrid
End Sub

Private Sub cmdAbrir_Click()
    Dim frmNuevoTurno As frmTurnoAbre
    
    Set frmNuevoTurno = New frmTurnoAbre
    
    frmNuevoTurno.Show vbModal
    
    ActualizaGrid
    
End Sub

Private Sub cmdActualiza_Click()
    ActualizaGrid
End Sub


Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForma MDIPrincipal, Me
    
    Me.dtpFechaIni.Value = Date
    Me.dtpFechaFin.Value = Date
    
    ActualizaGrid
End Sub
Private Sub ActualizaGrid()
    Dim adorcsTurnos As ADODB.Recordset
    
    #If SqlServer_ Then
        strSQL = "SELECT T.Idturno, T.FechaApertura, T.HoraApertura,"
        strSQL = strSQL & " T.UsuarioApertura, T.Numerocaja, T.TurnoNo, T.FechaCierre, T.HoraCierre, T.UsuarioCierre"
        strSQL = strSQL & " FROM TURNOS T"
        strSQL = strSQL & " WHERE"
        '29/06/07
        strSQL = strSQL & " Numerocaja=" & iNumeroCaja
        strSQL = strSQL & " AND T.FechaApertura BETWEEN " & "'" & Format(Me.dtpFechaIni.Value, "yyyymmdd") & "'" & " AND " & "'" & Format(Me.dtpFechaFin.Value, "yyyymmdd") & "'"
        If Me.chkSoloAbiertos.Value Then
            strSQL = strSQL & " AND T.Cerrado=0"
        End If
        strSQL = strSQL & " ORDER BY T.IdTurno, T.FechaApertura, T.TurnoNo"
    #Else
        strSQL = "SELECT T.Idturno, T.FechaApertura, T.HoraApertura,"
        strSQL = strSQL & " T.UsuarioApertura, T.Numerocaja, T.TurnoNo, T.FechaCierre, T.HoraCierre, T.UsuarioCierre"
        strSQL = strSQL & " FROM TURNOS T"
        strSQL = strSQL & " WHERE"
        '29/06/07
        strSQL = strSQL & " Numerocaja=" & iNumeroCaja
        strSQL = strSQL & " AND T.FechaApertura BETWEEN " & "#" & Format(Me.dtpFechaIni.Value, "mm/dd/yyyy") & "#" & " AND " & "#" & Format(Me.dtpFechaFin.Value, "mm/dd/yyyy") & "#"
        If Me.chkSoloAbiertos.Value Then
            strSQL = strSQL & " AND T.Cerrado=0"
        End If
        strSQL = strSQL & " ORDER BY T.IdTurno, T.FechaApertura, T.TurnoNo"
    #End If
    Set adorcsTurnos = New ADODB.Recordset
    adorcsTurnos.CursorLocation = adUseServer
    adorcsTurnos.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Me.ssdbgTurnos.RemoveAll
    Do Until adorcsTurnos.EOF
        Me.ssdbgTurnos.AddItem adorcsTurnos!IdTurno & vbTab & adorcsTurnos!FechaApertura & vbTab & adorcsTurnos!HoraApertura & vbTab & adorcsTurnos!NumeroCaja & vbTab & adorcsTurnos!TurnoNo & vbTab & adorcsTurnos!UsuarioApertura & vbTab & adorcsTurnos!FechaCierre & vbTab & adorcsTurnos!HoraCierre & vbTab & adorcsTurnos!UsuarioCierre
        adorcsTurnos.MoveNext
    Loop
    
    adorcsTurnos.Close
    Set adorcsTurnos = Nothing
    
End Sub







