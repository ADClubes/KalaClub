VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmPreVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preventa"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   9315
   Begin VB.Frame Frame1 
      Caption         =   "Filtro"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton cmdFiltro 
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   7440
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpFecFin 
         Height          =   375
         Left            =   5520
         TabIndex        =   7
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125566977
         CurrentDate     =   40550
      End
      Begin MSComCtl2.DTPicker dtpFecIni 
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   125566977
         CurrentDate     =   40550
      End
      Begin VB.ComboBox cmbTipoProspecto 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Fec. Final"
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Fec. Inicial"
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Status"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdRegPago 
      Caption         =   "Registrar Pago"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdDatosEnc 
      Caption         =   "Datos Encuesta"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdDatosPreVta 
      Caption         =   "Datos Preventa"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgProspectos 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   8895
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   7
      AllowUpdate     =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      BackColorOdd    =   12632256
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   5424
      Columns(0).Caption=   "Nombre"
      Columns(0).Name =   "Nombre"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4366
      Columns(1).Caption=   "Colonia"
      Columns(1).Name =   "Colonia"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Telefono"
      Columns(2).Name =   "Telefono"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   1852
      Columns(3).Caption=   "IdProspecto"
      Columns(3).Name =   "IdProspecto"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   2223
      Columns(4).Caption=   "FecRegistro"
      Columns(4).Name =   "FecRegistro"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   2117
      Columns(5).Caption=   "HoraRegistro"
      Columns(5).Name =   "HoraRegistro"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3200
      Columns(6).Caption=   "Status"
      Columns(6).Name =   "Status"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      _ExtentX        =   15690
      _ExtentY        =   3413
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
   Begin VB.Label lblCtrl 
      Height          =   255
      Index           =   3
      Left            =   7320
      TabIndex        =   12
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmPreVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbTipoProspecto_Change()
    ActGrid
End Sub



Private Sub cmdDatosEnc_Click()
    Dim frmPros As frmProspectos
    Dim bSoloLec As Boolean
    
    If Me.ssdbgProspectos.Rows = 0 Then
        Exit Sub
    End If
           
    If ChecaStatusPreventa(Me.ssdbgProspectos.Columns("IdProspecto").Value) > 1 Then
        MsgBox "Ya no es posible registar datos de" & vbCrLf & "encuesta para este prospecto", vbExclamation, "Error"
        bSoloLec = True
    End If
    
    
    Set frmPros = New frmProspectos
    frmPros.lRecNum = Me.ssdbgProspectos.Columns("IdProspecto").Value
    If bSoloLec Then frmPros.bReadOnly = True
    
    frmPros.Show vbModal
End Sub

Private Sub cmdDatosPreVta_Click()
    Dim frmPreVen As frmPreVentaDatos
    Dim bSoloLec As Boolean
    
    If Me.ssdbgProspectos.Rows = 0 Then
        Exit Sub
    End If
    
    
    
    If ChecaStatusPreventa(Me.ssdbgProspectos.Columns("IdProspecto").Value) < 1 Then
        MsgBox "Registre primero los datos" & vbCrLf & "de encuesta para este prospecto", vbExclamation, "Error"
        Exit Sub
    End If
    
    
    If ChecaStatusPreventa(Me.ssdbgProspectos.Columns("IdProspecto").Value) > 2 Then
        MsgBox "Ya no es posible registar datos de" & vbCrLf & "preventa para este prospecto", vbExclamation, "Error"
        bSoloLec = True
    End If
    
    
    
    Set frmPreVen = New frmPreVentaDatos
    frmPreVen.lRecNum = Me.ssdbgProspectos.Columns("IdProspecto").Value
    If bSoloLec Then frmPreVen.bReadOnly = True
    
    frmPreVen.Show vbModal
    
End Sub

Private Sub cmdFiltro_Click()
    ActGrid
End Sub

Private Sub cmdRegPago_Click()

    Dim frmPreVenPago As frmPreVentaPago
    
    
    If Me.ssdbgProspectos.Rows = 0 Then
        Exit Sub
    End If
    
    
    If ChecaStatusPreventa(Me.ssdbgProspectos.Columns("IdProspecto").Value) > 2 Then
        MsgBox "Ya no es posible registar datos de" & vbCrLf & "pago para este prospecto", vbExclamation, "Error"
        Exit Sub
    End If
    
    
    
    Set frmPreVenPago = New frmPreVentaPago
    frmPreVenPago.lRecNum = Me.ssdbgProspectos.Columns("IdProspecto").Value
    
    
    frmPreVenPago.Show vbModal

End Sub

Private Sub Form_Load()
        
        
    CentraForma MDIPrincipal, Me
    
    Me.dtpFecIni.Value = DateSerial(Year(Date), Month(Date), 1)
    Me.dtpFecFin.Value = DateSerial(Year(Date), Month(Date) + 1, 1) - 1
    
    
    
    Me.cmbTipoProspecto.Clear
    Me.cmbTipoProspecto.AddItem "Registrado"
    Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.NewIndex) = 0
    
    Me.cmbTipoProspecto.AddItem "Con Encuesta"
    Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.NewIndex) = 1
    
    Me.cmbTipoProspecto.AddItem "Con Datos Preventa"
    Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.NewIndex) = 2
    
    Me.cmbTipoProspecto.AddItem "Con Pago"
    Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.NewIndex) = 3
    
    Me.cmbTipoProspecto.AddItem "Cerrados"
    Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.NewIndex) = 3
    
    Me.cmbTipoProspecto.AddItem "Todos"
    Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.NewIndex) = 99
    
    Me.cmbTipoProspecto.Text = "Registrado"
    
    ActGrid
    
End Sub

Private Sub ActGrid()
    
    Dim adorcs As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    
    #If SqlServer_ Then
        strSQL = "SELECT Nombre + ' ' + A_Paterno + ' ' + A_Materno As Nombre, Colonia, Telefono, IdProspecto, FechaRegistro, HoraRegistro, PS.DescripcionStatus"
        strSQL = strSQL & " FROM Prospectos P INNER JOIN Prospectos_Status PS ON P.StatusProspecto = PS.StatusProspecto"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " FechaRegistro Between '" & Format(Me.dtpFecIni.Value, "yyyymmdd") & "' AND '" & Format(Me.dtpFecFin.Value, "yyyymmdd") & "'"
        
        If Me.cmbTipoProspecto.ListIndex <> -1 Then
            If Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.ListIndex) < 99 Then
                strSQL = strSQL & " And P.StatusProspecto=" & Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.ListIndex)
            End If
        End If
    #Else
        strSQL = "SELECT Nombre & ' ' & A_Paterno & ' ' & A_Materno As Nombre, Colonia, Telefono, IdProspecto, FechaRegistro, HoraRegistro, PS.DescripcionStatus"
        strSQL = strSQL & " FROM Prospectos P INNER JOIN Prospectos_Status PS ON P.StatusProspecto = PS.StatusProspecto"
        strSQL = strSQL & " Where ("
        strSQL = strSQL & "((FechaRegistro) Between #" & Format(Me.dtpFecIni.Value, "mm/dd/yyyy") & "# And #" & Format(Me.dtpFecFin.Value, "mm/dd/yyyy") & "#)"
        
        If Me.cmbTipoProspecto.ListIndex <> -1 Then
            If Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.ListIndex) < 99 Then
                strSQL = strSQL & " And ((P.StatusProspecto)=" & Me.cmbTipoProspecto.ItemData(Me.cmbTipoProspecto.ListIndex) & ")"
            End If
        End If
    
        strSQL = strSQL & ")"
    #End If
    
    
    Set adorcs = New ADODB.Recordset
    
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Me.ssdbgProspectos.RemoveAll
    Do Until adorcs.EOF
        Me.ssdbgProspectos.AddItem adorcs!Nombre & vbTab & adorcs!colonia & vbTab & adorcs!Telefono & vbTab & adorcs!IdProspecto & vbTab & adorcs!FechaRegistro & vbTab & adorcs!HoraRegistro & vbTab & adorcs!DescripcionStatus
        adorcs.MoveNext
    Loop
    
    adorcs.Close
    Set adorcs = Nothing
    
    Me.lblCtrl(3).Caption = Me.ssdbgProspectos.Rows & " Renglon(es)"
    
    Screen.MousePointer = vbDefault
    
End Sub

