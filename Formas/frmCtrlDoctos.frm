VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCtrlDoctos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de documentos"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10035
   Icon            =   "frmCtrlDoctos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDesmarcaEntrega 
      Caption         =   "Marcar como NO entregado"
      Height          =   615
      Left            =   1920
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrintChkLst 
      Caption         =   "Imprimir Lista Documentos"
      Height          =   615
      Left            =   6960
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheckList 
      Caption         =   "Imprimir Lista Pendientes"
      Height          =   615
      Left            =   8640
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuita 
      Caption         =   ">"
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdMarcaEntrega 
      Caption         =   "Marcar como entregado"
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgrega 
      Caption         =   "<"
      Height          =   615
      Left            =   4800
      TabIndex        =   3
      Top             =   2880
      Width           =   615
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgDocs 
      Height          =   2055
      Left            =   5520
      TabIndex        =   2
      Top             =   2640
      Width           =   4215
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   3
      AllowUpdate     =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   3
      Columns(0).Width=   873
      Columns(0).Caption=   "IdDocumento"
      Columns(0).Name =   "IdDocumento"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   4101
      Columns(1).Caption=   "Documento"
      Columns(1).Name =   "Documento"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   6694
      Columns(2).Caption=   "Ejemplos"
      Columns(2).Name =   "Ejemplos"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      _ExtentX        =   7435
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Posibles documentos"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgDocsUsuario 
      Height          =   2055
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   4365
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   1032
      Columns(0).Caption=   "Marca"
      Columns(0).Name =   "Marca"
      Columns(0).Alignment=   2
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Style=   2
      Columns(1).Width=   5371
      Columns(1).Caption=   "Documento"
      Columns(1).Name =   "Documento"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   3200
      Columns(2).Caption=   "Fecha"
      Columns(2).Name =   "Fecha"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   3200
      Columns(3).Caption=   "Hora"
      Columns(3).Name =   "Hora"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   3200
      Columns(4).Caption=   "Usuario"
      Columns(4).Name =   "Usuario"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Visible=   0   'False
      Columns(5).Caption=   "IdDocumento"
      Columns(5).Name =   "IdDocumento"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   7699
      _ExtentY        =   3625
      _StockProps     =   79
      Caption         =   "Documentos requeridos"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgUsuarios 
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9510
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   5
      AllowUpdate     =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   5
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
      Columns(4).Width=   1667
      Columns(4).Caption=   "Pendientes"
      Columns(4).Name =   "Pendientes"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      _ExtentX        =   16775
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
End
Attribute VB_Name = "frmCtrlDoctos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lIdTitular As Long

Private Sub cmdAgrega_Click()
        
    If Not ChecaSeguridad(Me.Name, Me.cmdAgrega.Name) Then
        Exit Sub
    End If
    
    If Me.ssdbgDocs.Rows = 0 Then
        Exit Sub
    End If
    
    Dim adocmd As ADODB.Command
    
    strSQL = "INSERT INTO DOCS_USUARIO ("
    strSQL = strSQL & " IdMember,"
    strSQL = strSQL & " IdDocumento,"
    strSQL = strSQL & " FechaAsigna,"
    strSQL = strSQL & " HoraAsigna,"
    strSQL = strSQL & " UsuarioAsigna"
    strSQL = strSQL & ") VALUES ("
    strSQL = strSQL & Me.ssdbgUsuarios.Columns("IdMember").Value & ","
    strSQL = strSQL & Me.ssdbgDocs.Columns("IdDocumento").Value & ","
    #If SqlServer_ Then
        strSQL = strSQL & "'" & Format(Date, "yyyymmdd") & "',"
    #Else
        strSQL = strSQL & "#" & Format(Date, "mm/dd/yyyy") & "#,"
    #End If
    strSQL = strSQL & "'" & Format(Now, "Hh:Nn") & "',"
    strSQL = strSQL & "'" & sDB_User & "')"
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    
    Set adocmd = Nothing
    
    ActualGridDocsUsuario
    ActualGridDocs
    
    
    
    
End Sub

Private Sub cmdCheckList_Click()
    Dim frmReport As frmReportViewer
    
    #If SqlServer_ Then
        strSQL = "SELECT USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Nombre + ' ' +  USUARIOS_CLUB.A_Paterno + ' ' +  USUARIOS_CLUB.A_Materno AS Nombre, CT_Documentos.NombreDocumento, USUARIOS_CLUB.NumeroFamiliar, CT_Documentos.EjemploDocumento"
        strSQL = strSQL & " FROM DOCS_USUARIO INNER JOIN CT_Documentos ON DOCS_USUARIO.IdDocumento = CT_Documentos.IdDocumento INNER JOIN USUARIOS_CLUB ON DOCS_USUARIO.IdMember = USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE DOCS_USUARIO.FechaEntrega Is Null And USUARIOS_CLUB.IdTitular = " & lIdTitular
        strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NumeroFamiliar"
    #Else
        strSQL = "SELECT USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Nombre & ' ' &  USUARIOS_CLUB.A_Paterno & ' ' &  USUARIOS_CLUB.A_Materno AS Nombre, CT_Documentos.NombreDocumento, USUARIOS_CLUB.NumeroFamiliar, CT_Documentos.EjemploDocumento"
        strSQL = strSQL & " FROM (DOCS_USUARIO INNER JOIN CT_Documentos ON DOCS_USUARIO.IdDocumento = CT_Documentos.IdDocumento) INNER JOIN USUARIOS_CLUB ON DOCS_USUARIO.IdMember = USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " Where (((DOCS_USUARIO.FechaEntrega) Is Null) And ((USUARIOS_CLUB.IdTitular) = " & lIdTitular & "))"
        strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NumeroFamiliar"
    #End If
    
    Set frmReport = New frmReportViewer
    frmReport.sNombreReporte = sDB_ReportSource & "\chklstDocs.rpt"
    frmReport.sQuery = strSQL
    
    frmReport.Show vbModal
    
    
End Sub

Private Sub cmdDesmarcaEntrega_Click()
     Dim lI As Long
     
     
      If Not ChecaSeguridad(Me.Name, Me.cmdDesmarcaEntrega.Name) Then
        Exit Sub
    End If
     
    
    If Me.ssdbgDocsUsuario.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdDesmarcaEntrega.Name) Then
        Exit Sub
    End If
    
    
    If Me.ssdbgDocsUsuario.SelBookmarks.Count = 0 Then
        If Me.ssdbgDocsUsuario.Columns("Marca").Value = 1 Then
            QuitaMarca Me.ssdbgUsuarios.Columns("IdMember").Value, Me.ssdbgDocsUsuario.Columns("IdDocumento").Value
        End If
    End If
    
    
    For lI = 0 To Me.ssdbgDocsUsuario.SelBookmarks.Count - 1
    
        Me.ssdbgDocsUsuario.Bookmark = Me.ssdbgDocsUsuario.SelBookmarks(lI)
    
        If Me.ssdbgDocsUsuario.Columns("Marca").Value = 1 Then
            QuitaMarca Me.ssdbgUsuarios.Columns("IdMember").Value, Me.ssdbgDocsUsuario.Columns("IdDocumento").Value
        End If
        
    Next
    
    ActualGridDocsUsuario
End Sub

Private Sub cmdMarcaEntrega_Click()
    
    Dim lI As Long
    
    
    If Not ChecaSeguridad(Me.Name, Me.cmdMarcaEntrega.Name) Then
        Exit Sub
    End If
    
    
    If Me.ssdbgDocsUsuario.Rows = 0 Then
        Exit Sub
    End If
    
    If Not ChecaSeguridad(Me.Name, Me.cmdMarcaEntrega.Name) Then
        Exit Sub
    End If
    
    
    If Me.ssdbgDocsUsuario.SelBookmarks.Count = 0 Then
        If Me.ssdbgDocsUsuario.Columns("Marca").Value = 0 Then
            ActualMarca Me.ssdbgUsuarios.Columns("IdMember").Value, Me.ssdbgDocsUsuario.Columns("IdDocumento").Value
        End If
    End If
    
    
    For lI = 0 To Me.ssdbgDocsUsuario.SelBookmarks.Count - 1
    
        Me.ssdbgDocsUsuario.Bookmark = Me.ssdbgDocsUsuario.SelBookmarks(lI)
    
        If Me.ssdbgDocsUsuario.Columns("Marca").Value = 0 Then
            ActualMarca Me.ssdbgUsuarios.Columns("IdMember").Value, Me.ssdbgDocsUsuario.Columns("IdDocumento").Value
        End If
        
    Next
    
    ActualGridDocsUsuario
    
End Sub

Private Sub cmdPrintChkLst_Click()
    Dim frmReport As frmReportViewer
    
    #If SqlServer_ Then
        strSQL = "SELECT USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Nombre + ' ' +  USUARIOS_CLUB.A_Paterno + ' ' +  USUARIOS_CLUB.A_Materno AS Nombre, CT_Documentos.NombreDocumento, USUARIOS_CLUB.NumeroFamiliar, CT_Documentos.EjemploDocumento"
        strSQL = strSQL & " FROM DOCS_USUARIO INNER JOIN CT_Documentos ON DOCS_USUARIO.IdDocumento = CT_Documentos.IdDocumento INNER JOIN USUARIOS_CLUB ON DOCS_USUARIO.IdMember = USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " WHERE USUARIOS_CLUB.IdTitular = " & lIdTitular
        strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NumeroFamiliar"
    #Else
        strSQL = "SELECT USUARIOS_CLUB.NoFamilia, USUARIOS_CLUB.Nombre & ' ' &  USUARIOS_CLUB.A_Paterno & ' ' &  USUARIOS_CLUB.A_Materno AS Nombre, CT_Documentos.NombreDocumento, USUARIOS_CLUB.NumeroFamiliar, CT_Documentos.EjemploDocumento"
        strSQL = strSQL & " FROM (DOCS_USUARIO INNER JOIN CT_Documentos ON DOCS_USUARIO.IdDocumento = CT_Documentos.IdDocumento) INNER JOIN USUARIOS_CLUB ON DOCS_USUARIO.IdMember = USUARIOS_CLUB.IdMember"
        strSQL = strSQL & " Where ( ((USUARIOS_CLUB.IdTitular) = " & lIdTitular & "))"
        strSQL = strSQL & " ORDER BY USUARIOS_CLUB.NumeroFamiliar"
    #End If
        
    Set frmReport = New frmReportViewer
    frmReport.sNombreReporte = sDB_ReportSource & "\chklstDocs.rpt"
    frmReport.sQuery = strSQL
    
    frmReport.Show vbModal
End Sub

Private Sub cmdQuita_Click()
    
    Dim adocmd As ADODB.Command
    Dim lRecords As Long
    
    
    
    If Me.ssdbgDocsUsuario.Rows = 0 Then
        Exit Sub
    End If
    
    
    If Not ChecaSeguridad(Me.Name, Me.cmdQuita.Name) Then
        Exit Sub
    End If
    
        
    #If SqlServer_ Then
        strSQL = "DELETE FROM DOCS_USUARIO"
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " IdMember = " & Me.ssdbgUsuarios.Columns("IdMember").Value
        strSQL = strSQL & " And IdDocumento = " & Me.ssdbgDocsUsuario.Columns("IdDocumento").Value
        strSQL = strSQL & " And FechaEntrega Is Null"
    #Else
        strSQL = "DELETE * FROM DOCS_USUARIO"
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " IdMember = " & Me.ssdbgUsuarios.Columns("IdMember").Value
        strSQL = strSQL & " And IdDocumento = " & Me.ssdbgDocsUsuario.Columns("IdDocumento").Value
        strSQL = strSQL & " And FechaEntrega Is Null"
    #End If
    
    Set adocmd = New ADODB.Command
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    adocmd.Execute lRecords
    
    
    If lRecords = 0 Then
        MsgBox "¡No fue posible remover este documento!", vbExclamation, "Error"
        Exit Sub
    End If
    
    
    Set adocmd = Nothing
    
    ActualGridDocsUsuario
    ActualGridDocs
    
End Sub

Private Sub Form_Load()


    CentraForma MDIPrincipal, Me
    
    #If SqlServer_ Then
        strSQL = "SELECT U.Nombre + ' ' + U.A_Paterno + ' ' + A_Materno AS Nombre, T.Descripcion, CONVERT(int, DateDiff(day, U.Fechanacio, getDate())/365.25) As Edad, U.IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB U INNER JOIN TIPO_USUARIO T"
        strSQL = strSQL & " ON U.IdTipoUsuario=T.IdTipoUsuario"
        strSQL = strSQL & " WHERE U.IdTitular = " & lIdTitular
    #Else
        strSQL = "SELECT U.Nombre & ' ' & U.A_Paterno & ' ' & A_Materno AS Nombre, T.Descripcion, Int(DateDiff('d', U.Fechanacio, Date())/365.25) As Edad, U.IdMember"
        strSQL = strSQL & " FROM USUARIOS_CLUB U INNER JOIN TIPO_USUARIO T"
        strSQL = strSQL & " ON U.IdTipoUsuario=T.IdTipoUsuario"
        strSQL = strSQL & " WHERE U.IdTitular = " & lIdTitular
    #End If
    
    LlenaSsDbGrid Me.ssdbgUsuarios, Conn, strSQL, 4
    
    
    
    Me.ssdbgDocsUsuario.StyleSets.Add ("NoEntregado")
    Me.ssdbgDocsUsuario.StyleSets("NoEntregado").BackColor = RGB(255, 255, 0)
    Me.ssdbgDocsUsuario.StyleSets("NoEntregado").ForeColor = RGB(0, 0, 0)
    
    
    Me.ssdbgDocsUsuario.StyleSets.Add ("Entregado")
    Me.ssdbgDocsUsuario.StyleSets("Entregado").BackColor = RGB(0, 255, 0)
    Me.ssdbgDocsUsuario.StyleSets("Entregado").ForeColor = RGB(0, 0, 0)
    
    
    
    ActualGridDocs
    
    
End Sub



Private Sub ssdbgDocsUsuario_RowLoaded(ByVal Bookmark As Variant)
    If Me.ssdbgDocsUsuario.Columns("Marca").CellValue(Bookmark) = 0 Then
        Me.ssdbgDocsUsuario.Columns(0).CellStyleSet "NoEntregado"
    Else
        Me.ssdbgDocsUsuario.Columns(0).CellStyleSet "Entregado"
    End If
End Sub

Private Sub ssdbgUsuarios_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    
    ActualGridDocsUsuario
    ActualGridDocs
    
    
End Sub


Private Sub ActualGridDocsUsuario()
    
    #If SqlServer_ Then
        strSQL = "SELECT CASE WHEN D.FechaEntrega IS NULL THEN 0 ELSE 1 END AS Marca, C.NombreDocumento, D.FechaEntrega, D.HoraEntrega, D.UsuarioEntrega, D.IdDocumento"
        strSQL = strSQL & " FROM DOCS_USUARIO D INNER JOIN CT_Documentos C ON D.IdDocumento = C.IdDocumento"
        strSQL = strSQL & " Where D.IdMember = " & Me.ssdbgUsuarios.Columns("IdMember").Value
    #Else
        strSQL = "SELECT iif( IsNull(D.FechaEntrega), 0, 1) AS Marca, C.NombreDocumento, D.FechaEntrega, D.HoraEntrega, D.UsuarioEntrega, D.IdDocumento"
        strSQL = strSQL & " FROM DOCS_USUARIO D INNER JOIN CT_Documentos C ON D.IdDocumento = C.IdDocumento"
        strSQL = strSQL & " Where D.IdMember = " & Me.ssdbgUsuarios.Columns("IdMember").Value
    #End If
    
    LlenaSsDbGrid Me.ssdbgDocsUsuario, Conn, strSQL, 6
End Sub


Private Sub ActualGridDocs()
    strSQL = "SELECT D.IdDocumento, D.NombreDocumento, D.EjemploDocumento"
    strSQL = strSQL & " FROM CT_Documentos D"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " D.IdDocumento Not In (SELECT U.IdDocumento FROM DOCS_USUARIO U WHERE U.IdMember = " & Me.ssdbgUsuarios.Columns("IdMember").Value & ")"
    
    LlenaSsDbGrid Me.ssdbgDocs, Conn, strSQL, 3
End Sub


Private Sub ActualMarca(lidMember As Long, lIdDocumento As Long)

    Dim adocmd As ADODB.Command

    strSQL = "UPDATE DOCS_USUARIO SET"
    #If SqlServer_ Then
        strSQL = strSQL & " FechaEntrega = " & "'" & Format(Date, "yyyymmdd") & "',"
    #Else
        strSQL = strSQL & " FechaEntrega = " & "#" & Format(Date, "mm/dd/yyyy") & "#,"
    #End If
    strSQL = strSQL & " HoraEntrega = " & "'" & Format(Now, "Hh:Nn") & "',"
    strSQL = strSQL & " UsuarioEntrega = " & "'" & sDB_User & "'"
    strSQL = strSQL & " Where"
    strSQL = strSQL & " IdMember = " & lidMember
    strSQL = strSQL & " And IdDocumento =" & lIdDocumento
    
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    
    
    

End Sub

Private Sub QuitaMarca(lidMember As Long, lIdDocumento As Long)

    Dim adocmd As ADODB.Command

    strSQL = "UPDATE DOCS_USUARIO SET"
    strSQL = strSQL & " FechaEntrega = Null" & ","
    strSQL = strSQL & " HoraEntrega = Null" & ","
    strSQL = strSQL & " UsuarioEntrega = " & "''"
    strSQL = strSQL & " Where"
    strSQL = strSQL & " IdMember = " & lidMember
    strSQL = strSQL & " And IdDocumento =" & lIdDocumento
    
    
    
    Set adocmd = New ADODB.Command
    
    adocmd.ActiveConnection = Conn
    adocmd.CommandType = adCmdText
    adocmd.CommandText = strSQL
    adocmd.Execute
    
    
    
    

End Sub
