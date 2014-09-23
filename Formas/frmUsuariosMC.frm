VERSION 5.00
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmUsuariosMC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios MultiClub"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optNombre 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   960
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton optClave 
      Caption         =   "No. familia"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDesActiva 
      Caption         =   "Desactivar"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdActiva 
      Caption         =   "Activar"
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   5640
      Width           =   1215
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgUsuarios 
      Height          =   3855
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   8175
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   979
      Columns(0).Caption=   "IdUnidad"
      Columns(0).Name =   "IdUnidad"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   2011
      Columns(1).Caption=   "NoInscripcion"
      Columns(1).Name =   "NoInscripcion"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   4577
      Columns(2).Caption=   "Nombre"
      Columns(2).Name =   "Nombre"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   2117
      Columns(3).Caption=   "FechaUPago"
      Columns(3).Name =   "FechaUPago"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1773
      Columns(4).Caption=   "Secuencial"
      Columns(4).Name =   "Secuencial"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "Fotofile"
      Columns(5).Name =   "Fotofile"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      _ExtentX        =   14420
      _ExtentY        =   6800
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
   Begin VB.CommandButton cmdBusca 
      Caption         =   "Buscar"
      Default         =   -1  'True
      Height          =   495
      Left            =   6720
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtBuscar 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo ssCmbUnidad 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      DataFieldList   =   "Column 0"
      AllowInput      =   0   'False
      AllowNull       =   0   'False
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
      Columns(0).Width=   3863
      Columns(0).Caption=   "Unidad"
      Columns(0).Name =   "Unidad"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1429
      Columns(1).Caption=   "IdUnidad"
      Columns(1).Name =   "IdUnidad"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      _ExtentX        =   4260
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Unidad"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmUsuariosMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActiva_Click()
Dim lPer As Long
Const nPerMax = 1
Const nDiasMax = 15

    If Me.ssdbgUsuarios.Rows = 0 Then
        Exit Sub
    End If
    
    lPer = CalculaPeriodos(Me.ssdbgUsuarios.Columns("FechaUPago").Value, Date, 1)
    
    
    'Si el número de periodos es mayor que el maximo permitido
'    If lPer > nPerMax Then
'    MsgBox "No se puede activar ya que tiene adeudo, favor de validar con su club sede.", vbExclamation, "Error"
'        Exit Sub
'    End If
    
        If lPer < nPerMax Then
            
            ActivaCredSQLMulti 1, Me.ssdbgUsuarios.Columns("Secuencial").Value, 1, Me.ssdbgUsuarios.Columns("NoInscripcion").Value, True, True
        
        
       ElseIf lPer = nPerMax Then
            If Day(Date) <= nDiasMax Then

                    ActivaCredSQLMulti 1, Me.ssdbgUsuarios.Columns("Secuencial").Value, 1, Me.ssdbgUsuarios.Columns("NoInscripcion").Value, True, True
               End If
         Else
            MsgBox "No se puede activar ya que tiene adeudo, favor de validar con su club.", vbExclamation, "Error"
            Exit Sub
        End If
    
End Sub

Private Sub cmdBusca_Click()
    If (Trim$(Me.txtBuscar.Text) <> vbNullString) Then
    
    'Checa la concordancia entre el tipo de dato a buscar
    'y la opcion de por nombre o por clave
    If Me.optNombre.Value And IsNumeric(Trim(Me.txtBuscar.Text)) Then
        Me.optClave.Value = True
    End If
    
    If Me.optClave.Value And Not IsNumeric(Trim(Me.txtBuscar.Text)) Then
        Me.optNombre.Value = True
    End If
    
    
    
    'Genera la porcion del query dependiendo si se busca por clave o por nombre
    If (Me.optNombre.Value) Then
        #If SqlServer_ Then
            sQryIni = "((Nombre + ' ' + A_Paterno + ' ' + A_Materno) LIKE '%" & Trim$(UCase$(Me.txtBuscar.Text)) & "%') "
        #Else
            sQryIni = "((Nombre & ' ' & A_Paterno & ' ' & A_Materno) LIKE '%" & Trim$(UCase$(Me.txtBuscar.Text)) & "%') "
        #End If
    Else
        sQryIni = "NoFamilia=" & Int(CDbl(Me.txtBuscar.Text))
    End If
    
    #If SqlServer_ Then
        strSQL = "SELECT USUARIOS_MC.IdClub, USUARIOS_MC.NoFamilia, USUARIOS_MC.A_Paterno + ' ' + USUARIOS_MC.A_Materno + ' ' + USUARIOS_MC.Nombre As Nombre, USUARIOS_MC.FechaUltimoPago,USUARIOS_MC.Secuencial,USUARIOS_MC.Descripcion As TipoUsuario , USUARIOS_MC.FotoFile"
        strSQL = strSQL & " From USUARIOS_MC "
        strSQL = strSQL & " Where"
        strSQL = strSQL & " USUARIOS_MC.IdClub = " & Me.ssCmbUnidad.Columns("IdUnidad").Value
        strSQL = strSQL & " And " & sQryIni & " ORDER BY "
        
   
        If (Me.optNombre.Value) Then
            strSQL = strSQL & "Nombre"
        Else
            strSQL = strSQL & "NumeroFamiliar"
        End If
        
    #End If
    
    LlenaSsDbGrid Me.ssdbgUsuarios, Conn, strSQL, 6
    
    If Me.ssdbgUsuarios.Rows = 0 Then
        MsgBox "No se encontró información", vbExclamation, "Error"
        Me.txtBuscar.SetFocus
        Exit Sub
    End If
    
    'MuestraFoto
End If
End Sub

Private Sub cmdDesActiva_Click()

    If Me.ssdbgUsuarios.Rows = 0 Then
        Exit Sub
    End If

     ActivaCredSQLMulti 1, Me.ssdbgUsuarios.Columns("Secuencial").Value, 1, Me.ssdbgUsuarios.Columns("NoInscripcion").Value, False, True
End Sub

Private Sub Form_Load()
    
    CentraForma MDIPrincipal, Me
    
    strSQL = "SELECT CT_UNIDAD.NombreUnidad, CT_UNIDAD.IdUnidad"
    strSQL = strSQL & " From CT_UNIDAD"
    strSQL = strSQL & " ORDER BY CT_UNIDAD.IdUnidad"
    
    LlenaSsCombo Me.ssCmbUnidad, Conn, strSQL, 2
    
End Sub




Private Sub ssdbgUsuarios_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    
    'MuestraFoto
    
End Sub

'Private Sub MuestraFoto()
'    Dim sNombreArc As String
'
'    sNombreArc = sG_RutaFoto & "\" & Me.ssCmbUnidad.Columns("IdUnidad").Value & "\" & Me.ssdbgUsuarios.Columns("Fotofile").Value & ".jpg"
'
'    If (Dir(sNombreArc) <> "") Then
'         Me.imgUsuario.Picture = LoadPicture(sNombreArc)
'    Else
'        Me.imgUsuario.Picture = LoadPicture("")
'    End If
'End Sub



Private Sub txtBuscar_GotFocus()
    txtBuscar.SelStart = 0
    txtBuscar.SelLength = Len(txtBuscar.Text)
End Sub
