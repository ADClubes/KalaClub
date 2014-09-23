VERSION 5.00
Begin VB.Form frmMensajes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensajes"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9180
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTitulo 
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      ToolTipText     =   "Ultimo"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Inicio"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      TabIndex        =   10
      ToolTipText     =   "Anterior"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtUsuario 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      ToolTipText     =   "Siguiente"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   7680
      TabIndex        =   5
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtMensaje 
      Height          =   2055
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1560
      Width           =   8415
   End
   Begin VB.TextBox txtHora 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblPos 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Hora"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lIdMember As Long

Dim adorcsMensajes As ADODB.Recordset
Dim lCurrentMessage As Long

Private Sub cmdFirst_Click()
    
    adorcsMensajes.MoveFirst
    lCurrentMessage = 1
    
    Me.cmdPrev.Enabled = False
    
    If adorcsMensajes.RecordCount > 1 Then
        Me.cmdNext.Enabled = True
        Me.cmdLast.Enabled = True
    End If
    
    DisplayMessage
    
End Sub

Private Sub cmdLast_Click()
        
    adorcsMensajes.MoveLast
    lCurrentMessage = adorcsMensajes.RecordCount
    
    Me.cmdNext.Enabled = False
    
    Me.cmdPrev.Enabled = True
    Me.cmdFirst.Enabled = True
    
    
    
    DisplayMessage
End Sub

Private Sub cmdNext_Click()
    
    adorcsMensajes.MoveNext
    
    lCurrentMessage = lCurrentMessage + 1
    
    If lCurrentMessage = adorcsMensajes.RecordCount Then
        Me.cmdNext.Enabled = False
        Me.cmdLast.Enabled = False
    End If
    
    If lCurrentMessage > 1 Then
        Me.cmdPrev.Enabled = True
        Me.cmdFirst.Enabled = True
    End If
    
    If Not adorcsMensajes.EOF Then
        DisplayMessage
    End If
    
    
    
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdPrev_Click()
    adorcsMensajes.MovePrevious
    
    lCurrentMessage = lCurrentMessage - 1
    
    If lCurrentMessage = 1 Then
        Me.cmdPrev.Enabled = False
        Me.cmdFirst.Enabled = False
    End If
    
    If lCurrentMessage < adorcsMensajes.RecordCount Then
        Me.cmdNext.Enabled = True
        Me.cmdLast.Enabled = True
    End If
    
    If Not adorcsMensajes.BOF Then
        DisplayMessage
    End If
    
End Sub

Private Sub Form_Activate()
    
    strSQL = "SELECT FechaAlta, HoraAlta, IdUsuarioAlta, Titulo,TextoMensaje"
    strSQL = strSQL & " FROM MENSAJES"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " IdMember=" & lIdMember
    strSQL = strSQL & " AND Leido=0"
    strSQL = strSQL & " ORDER BY FechaAlta DESC, HoraAlta DESC"
    
    
    Set adorcsMensajes = New ADODB.Recordset
    With adorcsMensajes
        .ActiveConnection = Conn
        .CursorLocation = adUseServer
        .CursorType = adOpenStatic
        .LockType = adLockReadOnly
        .Open strSQL
    End With
    
    
    If Not adorcsMensajes.EOF Then
        lCurrentMessage = 1
        DisplayMessage
    End If
    
    If adorcsMensajes.RecordCount > 1 Then
        Me.cmdNext.Enabled = True
        Me.cmdLast.Enabled = True
    End If
    
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()
    CENTRAFORMA MDIPrincipal, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If adorcsMensajes.State Then
        adorcsMensajes.Close
    End If
    
    Set adorcsMensajes = Nothing
End Sub
Private Sub DisplayMessage()
    
    Me.txtFecha.Text = Format(adorcsMensajes!FechaAlta, "dd/MMM/yyyy")
    Me.txtHora.Text = adorcsMensajes!HoraAlta
    Me.txtUsuario.Text = adorcsMensajes!IdUsuarioAlta
    Me.txtTitulo.Text = adorcsMensajes!Titulo
    Me.txtMensaje = adorcsMensajes!TextoMensaje
    
    Me.lblPos.Caption = "Mensaje " & lCurrentMessage & " de " & adorcsMensajes.RecordCount
    
End Sub
