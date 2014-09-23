VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmStatus 
   Caption         =   "Status"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Status Nuevo"
      Height          =   975
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   6735
      Begin VB.CommandButton cmdCambia 
         Caption         =   "Cambiar"
         Height          =   495
         Left            =   5160
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   3135
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo sscmbStatus 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
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
         Columns(0).Width=   1535
         Columns(0).Caption=   "IdStatus"
         Columns(0).Name =   "IdStatus"
         Columns(0).DataField=   "Column 0"
         Columns(0).DataType=   8
         Columns(0).FieldLen=   256
         Columns(1).Width=   3493
         Columns(1).Caption=   "Descripcion"
         Columns(1).Name =   "Descripcion"
         Columns(1).DataField=   "Column 1"
         Columns(1).DataType=   8
         Columns(1).FieldLen=   256
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   93
         BackColor       =   -2147483643
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status Actual"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Fecha aplicación"
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Status"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lIdTitular As Long

Private Sub cmdCambia_Click()

    Dim adocmd As ADODB.Command
    
    
    On Error GoTo Error_Catch
    
    #If SqlServer_ Then
        strSQL = ""
        strSQL = "UPDATE USUARIOS_CLUB SET"
        strSQL = strSQL & " IdStatusAnterior = IdStatus" & ","
        strSQL = strSQL & " FechaStatusAnterior = FechaStatus" & ","
        strSQL = strSQL & " IdStatus = " & Me.sscmbStatus.Text & ","
        strSQL = strSQL & " FechaStatus =" & "'" & Format(Date, "yyyymmdd") & "'"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdMember=" & lIdTitular
    #Else
        strSQL = ""
        strSQL = "UPDATE USUARIOS_CLUB SET"
        strSQL = strSQL & " IdStatusAnterior = IdStatus" & ","
        strSQL = strSQL & " FechaStatusAnterior = FechaStatus" & ","
        strSQL = strSQL & " IdStatus = " & Me.sscmbStatus.Text & ","
        strSQL = strSQL & " FechaStatus =" & "#" & Format(Date, "mm/dd/yyyy") & "#"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IdMember=" & lIdTitular
    #End If
    
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandText = strSQL
    adocmd.CommandType = adCmdText
    adocmd.Execute
    
    Set adocmd = Nothing
    
    MsgBox "Status actualizado", vbInformation, "Ok"
    
    Unload Me
    
    On Error GoTo 0
    
    Exit Sub
    
Error_Catch:

    MsgError
    
    
End Sub

Private Sub Form_Load()
    
    Dim adorcs As ADODB.Recordset
    
    
    
    strSQL = ""
    strSQL = "SELECT CT_STATUS.IdStatus, CT_STATUS.DescripcionStatus"
    strSQL = strSQL & " FROM CT_STATUS"
    strSQL = strSQL & " ORDER BY CT_STATUS.IdStatus"
    
    
    LlenaSsCombo Me.sscmbStatus, Conn, strSQL, 2
    
    
    
    
    
    strSQL = ""
    strSQL = "SELECT USUARIOS_CLUB.IdStatus, USUARIOS_CLUB.FechaStatus, CT_STATUS.DescripcionStatus"
    strSQL = strSQL & " FROM USUARIOS_CLUB INNER JOIN CT_STATUS ON USUARIOS_CLUB.IdStatus = CT_STATUS.IdStatus"
    strSQL = strSQL & " WHERE (((USUARIOS_CLUB.IdMember)=" & lIdTitular & "))"
    
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not adorcs.EOF Then
        If IsNull(adorcs!idstatus) Then
            Me.txtCtrl(0).Text = "ACTIVO A"
            Me.txtCtrl(1).Text = vbNullString
        Else
            Me.txtCtrl(0).Text = adorcs!DescripcionStatus & " (" & adorcs!idstatus & ")"
            Me.txtCtrl(1).Text = Format(adorcs!FechaStatus, "dd/MMM/yyyy")
        End If
    End If

    CentraForma MDIPrincipal, Me
    
End Sub

Private Sub sscmbStatus_Click()
    Me.txtCtrl(2).Text = Me.sscmbStatus.Columns("Descripcion").Value
End Sub

