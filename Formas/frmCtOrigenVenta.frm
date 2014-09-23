VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCtOrigenVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de origen de venta"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Modificación"
      Height          =   495
      Left            =   1920
      TabIndex        =   16
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   5640
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame FrameEdita 
      Caption         =   "Datos"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   6855
      Begin MSComCtl2.DTPicker dtpFechaInicial 
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   40125
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   5040
         TabIndex        =   7
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdGuarda 
         Caption         =   "Guardar"
         Height          =   495
         Left            =   3360
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chkSinCosto 
         Caption         =   "¿Es sin costo?"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CheckBox chkEsVenta 
         Caption         =   "¿Es venta?"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   1935
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbGrupo 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   480
         Width           =   2895
         DataFieldList   =   "Column 0"
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
         Row.Count       =   6
         Row(0)          =   "VENTA"
         Row(1)          =   "TRASPASO"
         Row(2)          =   "CORTESIA"
         Row(3)          =   "INTERCAMBIO"
         Row(4)          =   "STAFF"
         Row(5)          =   "SISTEMA"
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).Caption=   "Grupo"
         Columns(0).Name =   "Grupo"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         _ExtentX        =   5106
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin VB.TextBox txtCtrl 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtpFechaFinal 
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   40125
      End
      Begin VB.Label Label4 
         Caption         =   "Vigente hasta"
         Height          =   255
         Left            =   3840
         TabIndex        =   15
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Vigente desde"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgOrigenesVenta 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6855
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   7
      AllowUpdate     =   0   'False
      AllowGroupSwapping=   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   609
      Columns.Count   =   7
      Columns(0).Width=   1799
      Columns(0).Caption=   "IdOrigenVta"
      Columns(0).Name =   "IdOrigenVta"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4445
      Columns(1).Caption=   "Descripcion"
      Columns(1).Name =   "Descripcion"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2275
      Columns(2).Caption=   "Grupo"
      Columns(2).Name =   "Grupo"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1191
      Columns(3).Caption=   "Venta"
      Columns(3).Name =   "Venta"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1588
      Columns(4).Caption=   "Sin Costo"
      Columns(4).Name =   "Sin Costo"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   1588
      Columns(5).Caption=   "FecIni"
      Columns(5).Name =   "FecIni"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "FecFin"
      Columns(6).Name =   "FecFin"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      _ExtentX        =   12091
      _ExtentY        =   2990
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
Attribute VB_Name = "frmCtOrigenVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lIdOrigenVenta As Long

Private Sub cmdCancelar_Click()
    Me.FrameEdita.Visible = False
    Me.cmdNuevo.Enabled = True
    Me.cmdSalir.Enabled = True
    
End Sub

Private Sub cmdEditar_Click()
    lIdOrigenVenta = Me.ssdbgOrigenesVenta.Columns("IdOrigenVta").Value
    Me.txtCtrl.Text = Me.ssdbgOrigenesVenta.Columns("Descripcion").Value
    Me.sscmbGrupo.Text = Me.ssdbgOrigenesVenta.Columns("Grupo").Value
    Me.chkEsVenta.Value = IIf(Me.ssdbgOrigenesVenta.Columns("Venta").Value = "S", 1, 0)
    Me.chkSinCosto.Value = IIf(Me.ssdbgOrigenesVenta.Columns("Sin Costo").Value = "S", 1, 0)
    Me.dtpFechaInicial = Me.ssdbgOrigenesVenta.Columns("FecIni").Value
    Me.dtpFechaFinal = Me.ssdbgOrigenesVenta.Columns("FecFin").Value
    
    Me.cmdNuevo.Enabled = False
    Me.cmdSalir.Enabled = False
    
    
    Me.FrameEdita.Visible = True
    
    
    
End Sub

Private Sub cmdGuarda_Click()
    
    Dim adocmd As ADODB.Command
    
    
    If Me.txtCtrl.Text = vbNullString Then
        MsgBox "La descripción no puede quedar vacía", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    If Me.sscmbGrupo.Text = vbNullString Then
        MsgBox "Seleccione un grupo", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    If Me.dtpFechaInicial.Value < Date Then
        MsgBox "La fecha inicial no puede ser menor a la fecha actual", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    If Me.dtpFechaFinal.Value < Date Then
        MsgBox "La fecha final no puede ser menor a la fecha actual", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    If Me.dtpFechaInicial.Value > Me.dtpFechaFinal Then
        MsgBox "La fecha inicial no puede ser mayor a la fecha final", vbExclamation, "Verifique"
        Exit Sub
    End If
    
    
    If lIdOrigenVenta = 0 Then
        strSQL = "INSERT INTO CT_ORIGEN_VENTA("
        strSQL = strSQL & "DescripcionOrigenVenta,"
        strSQL = strSQL & "Grupo,"
        strSQL = strSQL & "EsVenta,"
        strSQL = strSQL & "SinCosto,"
        strSQL = strSQL & "FechaInicial,"
        strSQL = strSQL & "FechaFinal)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & "'" & UCase(Me.txtCtrl.Text) & "',"
        strSQL = strSQL & "'" & Me.sscmbGrupo.Text & "',"
        strSQL = strSQL & "'" & IIf(Me.chkEsVenta.Value, "S", "N") & "',"
        strSQL = strSQL & "'" & IIf(Me.chkSinCosto.Value, "S", "N") & "',"
        strSQL = strSQL & "'" & Format(Me.dtpFechaInicial.Value, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Format(Me.dtpFechaFinal.Value, "yyyymmdd") & "'"
        strSQL = strSQL & ")"
    Else
        strSQL = "UPDATE CT_ORIGEN_VENTA SET"
        strSQL = strSQL & " DescripcionOrigenVenta = " & "'" & UCase(Me.txtCtrl.Text) & "',"
        strSQL = strSQL & " Grupo = " & "'" & Me.sscmbGrupo.Text & "',"
        strSQL = strSQL & " EsVenta = " & "'" & IIf(Me.chkEsVenta.Value, "S", "N") & "',"
        strSQL = strSQL & " SinCosto = " & "'" & IIf(Me.chkSinCosto.Value, "S", "N") & "',"
        strSQL = strSQL & " FechaInicial = " & "#" & Format(Me.dtpFechaInicial.Value, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & " FechaFinal = " & "#" & Format(Me.dtpFechaFinal.Value, "mm/dd/yyyy") & "#"
        strSQL = strSQL & " Where ("
        strSQL = strSQL & "IdOrigenVenta=" & lIdOrigenVenta
        strSQL = strSQL & ")"
    End If
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    
    adocmd.Execute
    
    Set adocmd = Nothing
    
       
    Me.FrameEdita.Visible = False
    
    ActGrid
    
End Sub

Private Sub cmdNuevo_Click()
    Me.cmdNuevo.Enabled = False
    Me.cmdSalir.Enabled = False
    
    Me.txtCtrl.Text = vbNullString
    Me.dtpFechaInicial = Date
    Me.dtpFechaFinal = Date
    Me.chkEsVenta.Value = 0
    Me.chkSinCosto.Value = 0
    
    lIdOrigenVenta = 0
    
    Me.FrameEdita.Visible = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub ActGrid()


strSQL = "SELECT CT_ORIGEN_VENTA.idOrigenVenta, CT_ORIGEN_VENTA.DescripcionOrigenVenta, CT_ORIGEN_VENTA.Grupo, CT_ORIGEN_VENTA.EsVenta, CT_ORIGEN_VENTA.SinCosto, CT_ORIGEN_VENTA.FechaInicial, CT_ORIGEN_VENTA.FechaFinal"
strSQL = strSQL & " From CT_ORIGEN_VENTA"
strSQL = strSQL & " ORDER BY CT_ORIGEN_VENTA.idOrigenVenta"

LlenaSsDbGrid Me.ssdbgOrigenesVenta, Conn, strSQL, 7



End Sub


Private Sub Form_Load()
    CentraForma MDIPrincipal, Me
    ActGrid
End Sub
