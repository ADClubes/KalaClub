VERSION 5.00
Begin VB.Form frmPreVentaPago 
   Caption         =   "Registro de pago inicial"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCtrl 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   17
      Top             =   240
      Width           =   4695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mantenimiento"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   6495
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   15
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox txtCtrl 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Forma de pago"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Importe a pagar"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Inscripcion"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6495
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Top             =   960
         Width           =   4695
      End
      Begin VB.TextBox txtCtrl 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtCtrl 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Forma de pago"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Importe a pagar"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblCtrl 
         Caption         =   "Tipo Inscripcion"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancela 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Pago OK"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmPreVentaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lRecNum As Long

Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim adocmd As ADODB.Command
    
    #If SqlServer_ Then
        strSQL = ""
        strSQL = "UPDATE PROSPECTOS SET"
        strSQL = strSQL & " FechaCobro=" & "'" & Format(Date, "yyyymmdd") & "'" & ","
        strSQL = strSQL & " HoraCobro=" & "'" & Format(Now, "Hh:Nn:Ss") & "'" & ","
        strSQL = strSQL & " StatusProspecto=" & 3
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IDProspecto = " & lRecNum
        strSQL = strSQL & " AND StatusProspecto = 2"
    #Else
        strSQL = ""
        strSQL = "UPDATE PROSPECTOS SET"
        strSQL = strSQL & " FechaCobro=" & "#" & Format(Date, "mm/dd/yyyy") & "#" & ","
        strSQL = strSQL & " HoraCobro=" & "'" & Format(Now, "Hh:Nn:Ss") & "'" & ","
        strSQL = strSQL & " StatusProspecto=" & 3
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " IDProspecto = " & lRecNum
        strSQL = strSQL & " AND StatusProspecto = 2"
    #End If
    
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    
    adocmd.Execute
    
    
    Set adocmd = Nothing
    
    MsgBox "Pago registrado", vbInformation, "Ok"
    
    
    
    
End Sub

Private Sub Form_Load()
    LlenaDatos
    
    CentraForma MDIPrincipal, Me
End Sub




Private Sub LlenaDatos()
    
    Dim adorcs As ADODB.Recordset
    
    strSQL = "SELECT PROSPECTOS.Nombre, PROSPECTOS.A_Paterno, PROSPECTOS.A_Materno, PROSPECTOS.MontoEnganche, PROSPECTOS.ImporteMantenimiento, PROSPECTOS.TipoMantenimiento, PROSPECTOS.IdOpcionPagoMant, TIPO_MEMBRESIA.Descripcion AS TM_Descripcion, FORMA_PAGO.Descripcion AS FP_Descripcion, PROSPECTOS.IdOpcionPagoInsc, FORMA_PAGO_OPCION.Descripcion AS FPO_Descripcion, FPM.Descripcion AS FPM_Descripcion, FPOM.Descripcion AS FPOM_Descripcion"
    strSQL = strSQL & " FROM PROSPECTOS INNER JOIN TIPO_MEMBRESIA ON PROSPECTOS.IdTipoInscripcion = TIPO_MEMBRESIA.idTipoMembresia INNER JOIN FORMA_PAGO ON PROSPECTOS.IdFormaPagoInsc = FORMA_PAGO.IdFormaPago LEFT JOIN FORMA_PAGO_OPCION ON PROSPECTOS.IdFormaPagoInsc = FORMA_PAGO_OPCION.IdFormaPago AND PROSPECTOS.IdOpcionPagoInsc = FORMA_PAGO_OPCION.IdFormadePagoOpcion INNER JOIN FORMA_PAGO AS FPM ON PROSPECTOS.IdFormaPagoMant = FPM.IdFormaPago LEFT JOIN FORMA_PAGO_OPCION AS FPOM ON PROSPECTOS.IdFormaPagoMant = FPOM.IdFormaPago AND PROSPECTOS.IdOpcionPagoMant = FPOM.IdFormadePagoOpcion"
    strSQL = strSQL & " WHERE PROSPECTOS.idProspecto = " & lRecNum
    
    Set adorcs = New ADODB.Recordset
    adorcs.CursorLocation = adUseServer
    
    adorcs.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    If Not adorcs.EOF Then
    
        Me.txtCtrl(6).Text = adorcs!Nombre & " " & adorcs!A_Paterno & " " & adorcs!A_Materno & " (" & lRecNum & ")"
    
        Me.txtCtrl(0).Text = adorcs![TM_Descripcion]
        Me.txtCtrl(1).Text = Format(adorcs!MontoEnganche, "$#,##0.00")
        Me.txtCtrl(2).Text = adorcs![FP_Descripcion] & " " & adorcs![FPO_Descripcion]
        
        Me.txtCtrl(3).Text = GetTipoMantenimiento(adorcs!Tipomantenimiento)
        Me.txtCtrl(4).Text = Format(adorcs!ImporteMantenimiento, "$#,##0.00")
        Me.txtCtrl(5).Text = adorcs![FPM_Descripcion] & " " & adorcs![FPOM_Descripcion]
    End If
    
    adorcs.Close
    Set adorcs = Nothing
    

End Sub
