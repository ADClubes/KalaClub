VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmFormatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formatos"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   Icon            =   "frmFormatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Expediente 
      Caption         =   "Expediente"
      Height          =   1455
      Left            =   4920
      TabIndex        =   6
      Top             =   5040
      Width           =   6015
      Begin VB.Label Label3 
         Caption         =   "Credencial:"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Expediente:"
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Expediente:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdQuita 
      Caption         =   ">"
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      ToolTipText     =   "Quitar"
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdImprimeFormato 
      Caption         =   "Imprimir"
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdAgrega 
      Caption         =   "<"
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      ToolTipText     =   "Agregar"
      Top             =   2880
      Width           =   615
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgUsuarios 
      Height          =   2055
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   4
      AllowUpdate     =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   6350
      Columns(0).Caption=   "Nombre"
      Columns(0).Name =   "Nombre"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4815
      Columns(1).Caption=   "Tipo"
      Columns(1).Name =   "Tipo"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   1270
      Columns(2).Caption=   "Edad"
      Columns(2).Name =   "Edad"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Visible=   0   'False
      Columns(3).Caption=   "IdMember"
      Columns(3).Name =   "IdMember"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   18653
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Usuario"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgFormatosUsuario 
      Height          =   2055
      Left            =   360
      TabIndex        =   1
      Top             =   2640
      Width           =   5295
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   4
      AllowUpdate     =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   4
      Columns(0).Width=   4498
      Columns(0).Caption=   "Formato"
      Columns(0).Name =   "Formato"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1693
      Columns(1).Caption=   "Alta"
      Columns(1).Name =   "Alta"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Usuario"
      Columns(2).Name =   "Usuario"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3201
      Columns(3).Caption=   "IdFormato"
      Columns(3).Name =   "IdFormato"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      _ExtentX        =   9340
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Formatos existentes"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgFormatos 
      Height          =   2055
      Left            =   6720
      TabIndex        =   2
      Top             =   2640
      Width           =   4215
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   2
      AllowUpdate     =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   2
      Columns(0).Width=   3200
      Columns(0).Caption=   "Formato"
      Columns(0).Name =   "Formato"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1455
      Columns(1).Caption=   "IdFormato"
      Columns(1).Name =   "IdFormato"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   7435
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Seleccionar formato"
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
Attribute VB_Name = "frmFormatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lIdTitular As Long

Private Sub cmdAgrega_Click()
    Dim lFormato As Long
    
    If Not ChecaSeguridad(Me.Name, Me.cmdAgrega.Name) Then
        Exit Sub
    End If
    
    If MsgBox("Esta seguro de que desea asignar este formato", vbQuestion + vbOKCancel, "Confirme") = vbCancel Then
        Exit Sub
    End If
    
    
    lFormato = CreaFormato(Me.ssdbgFormatos.Columns("IdFormato").Value, Me.ssdbgUsuarios.Columns("IdMember").Value)
    
     ActualGridFormatosUsuario
     ActualGridFormatos
    
End Sub

Private Sub cmdImprimeFormato_Click()
    Dim frmimfmt As frmImprimeFormato
    
    If Not ChecaSeguridad(Me.Name, Me.cmdImprimeFormato.Name) Then
        Exit Sub
    End If
    
    If Me.ssdbgFormatos.Rows = 0 Then
        Exit Sub
    End If
    
    If Me.ssdbgFormatosUsuario.Rows = 0 Then
        Exit Sub
    End If
    
    Set frmimfmt = New frmImprimeFormato
    
    frmimfmt.lIdFormato = Me.ssdbgFormatosUsuario.Columns("IdFormato").Value
    frmimfmt.Show vbModal
    
End Sub

Private Sub cmdQuita_Click()
    
    Dim adocmd As ADODB.Command
    
    If Not ChecaSeguridad(Me.Name, Me.cmdQuita.Name) Then
        Exit Sub
    End If
    
    
    If Me.ssdbgFormatosUsuario.Rows = 0 Then
        Exit Sub
    End If
    
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    
    #If SqlServer_ Then
        strSQL = "DELETE FROM FORMATOS_DETALLE"
        strSQL = strSQL & " where IdFormato = " & Me.ssdbgFormatosUsuario.Columns("IdFormato").Value
    #Else
        strSQL = "DELETE * FROM FORMATOS_DETALLE"
        strSQL = strSQL & " where IdFormato = " & Me.ssdbgFormatosUsuario.Columns("IdFormato").Value
    #End If
    
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    #If SqlServer_ Then
        strSQL = "DELETE FROM FORMATOS"
        strSQL = strSQL & " where IdFormato = " & Me.ssdbgFormatosUsuario.Columns("IdFormato").Value
    #Else
        strSQL = "DELETE * FROM FORMATOS"
        strSQL = strSQL & " where IdFormato = " & Me.ssdbgFormatosUsuario.Columns("IdFormato").Value
    #End If
    
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    Set adocmd = Nothing
    
    ActualGridFormatos
    ActualGridFormatosUsuario
    
End Sub

Private Sub Form_Load()
    CentraForma MDIPrincipal, Me
    
    #If SqlServer_ Then
        strSQL = "SELECT U.Nombre + ' ' + U.A_Paterno + ' ' + A_Materno AS Nombre, T.Descripcion, CONVERT(int, DateDiff(day, U.Fechanacio, getDate())/365.25) As Edad, U.IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB U INNER JOIN TIPO_USUARIO T"
        strSQL = strSQL & " ON U.IdTipoUsuario=T.IdTipoUsuario"
        strSQL = strSQL & " WHERE U.IdTitular = " & lIdTitular
        strSQL = strSQL & " ORDER BY NumeroFamiliar"
    #Else
        strSQL = "SELECT U.Nombre & ' ' & U.A_Paterno & ' ' & A_Materno AS Nombre, T.Descripcion, Int(DateDiff('d', U.Fechanacio, Date())/365.25) As Edad, U.IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB U INNER JOIN TIPO_USUARIO T"
        strSQL = strSQL & " ON U.IdTipoUsuario=T.IdTipoUsuario"
        strSQL = strSQL & " WHERE U.IdTitular = " & lIdTitular
        strSQL = strSQL & " ORDER BY NumeroFamiliar"
    #End If
    
    LlenaSsDbGrid Me.ssdbgUsuarios, Conn, strSQL, 4
    
    ActualGridFormatos
    ActualGridFormatosUsuario
    
End Sub
Private Sub ActualGridFormatosUsuario()
    
        
    strSQL = "SELECT CT_Formatos.NombreFormato, FORMATOS.FechaAlta, FORMATOS.UsuarioAlta, FORMATOS.IdFormato"
    strSQL = strSQL & " FROM FORMATOS INNER JOIN CT_Formatos ON FORMATOS.IdTipoFormato = CT_Formatos.IdTipoFormato"
    strSQL = strSQL & " WHERE Formatos.IdMember = " & Me.ssdbgUsuarios.Columns("IdMember").Value
    
    
    LlenaSsDbGrid Me.ssdbgFormatosUsuario, Conn, strSQL, 4
    
End Sub

Private Sub ActualGridFormatos()
    
    
    
    strSQL = "SELECT CT_Formatos.NombreFormato, CT_Formatos.IdTipoFormato"
    strSQL = strSQL & " From CT_Formatos"
    strSQL = strSQL & " Where ("
    strSQL = strSQL & "((Status)='A')"
    strSQL = strSQL & " AND ((IdTipoFormato) Not In (SELECT FORMATOS.IdTipoFormato FROM FORMATOS Where Formatos.IdMember = " & Me.ssdbgUsuarios.Columns("IdMember").Value & ")))"
    strSQL = strSQL & " ORDER BY CT_Formatos.IdTipoFormato"
    
    
    LlenaSsDbGrid Me.ssdbgFormatos, Conn, strSQL, 2
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub ssdbgUsuarios_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    
    
    ActualGridFormatosUsuario
    ActualGridFormatos
    
    
    
End Sub
