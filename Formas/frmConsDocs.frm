VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmConsDocs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Documentos"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11910
   Begin VB.Frame frmTmp 
      Caption         =   "Ver"
      Height          =   735
      Left            =   2520
      TabIndex        =   11
      Top             =   1080
      Width           =   3135
      Begin VB.OptionButton optMes 
         Caption         =   "Mes"
         Height          =   195
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optTodas 
         Caption         =   "Todas"
         Height          =   195
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optDia 
         Caption         =   "Día"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.TextBox txtObserva 
      Enabled         =   0   'False
      Height          =   405
      Left            =   120
      TabIndex        =   9
      Top             =   6840
      Width           =   11655
   End
   Begin VB.ComboBox cmbOrden 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   2535
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imglst"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ImpFac"
            Description     =   "Impresión de Documento"
            Object.ToolTipText     =   "Impresión de Documentos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Description     =   "Actualizar Consulta"
            Object.ToolTipText     =   "Actualizar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CanFac"
            Description     =   "Cancelación de Documentos"
            Object.ToolTipText     =   "Cancelación de Documentos"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ImpCupon"
            Description     =   "Impresión de Cupones"
            Object.ToolTipText     =   "Impresion de Cupones"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "MuestraCFD"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgDocsDet 
      Bindings        =   "frmConsDocs.frx":0000
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   11655
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   1402
      Columns(0).Caption=   "Renglon"
      Columns(0).Name =   "Renglon"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   2
      Columns(0).FieldLen=   256
      Columns(1).Width=   1323
      Columns(1).Caption=   "Clave"
      Columns(1).Name =   "IdConcepto"
      Columns(1).DataField=   "Column 1"
      Columns(1).FieldLen=   256
      Columns(2).Width=   4710
      Columns(2).Caption=   "Concepto"
      Columns(2).Name =   "Concepto"
      Columns(2).DataField=   "Column 2"
      Columns(2).FieldLen=   256
      Columns(3).Width=   2355
      Columns(3).Caption=   "Periodo"
      Columns(3).Name =   "Periodo"
      Columns(3).DataField=   "Column 3"
      Columns(3).FieldLen=   256
      Columns(4).Width=   1561
      Columns(4).Caption=   "Cantidad"
      Columns(4).Name =   "Cantidad"
      Columns(4).DataField=   "Column 4"
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "Total"
      Columns(5).Name =   "Total"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   6
      Columns(5).NumberFormat=   "$##,##0.00"
      Columns(5).FieldLen=   256
      _ExtentX        =   20558
      _ExtentY        =   4260
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgDocs 
      Bindings        =   "frmConsDocs.frx":001B
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   11655
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   11
      AllowUpdate     =   0   'False
      AllowRowSizing  =   0   'False
      AllowColumnMoving=   0
      AllowColumnSwapping=   0
      SelectTypeCol   =   0
      SelectTypeRow   =   1
      RowHeight       =   423
      Columns.Count   =   11
      Columns(0).Width=   1931
      Columns(0).Caption=   "Folio"
      Columns(0).Name =   "Folio"
      Columns(0).DataField=   "Column 0"
      Columns(0).FieldLen=   256
      Columns(1).Width=   1773
      Columns(1).Caption=   "Numero"
      Columns(1).Name =   "Numero"
      Columns(1).DataField=   "Column 1"
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(2).Width=   2514
      Columns(2).Caption=   "Fecha"
      Columns(2).Name =   "Fecha"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   7
      Columns(2).FieldLen=   256
      Columns(3).Width=   1429
      Columns(3).Caption=   "Hora"
      Columns(3).Name =   "Hora"
      Columns(3).DataField=   "Column 3"
      Columns(3).FieldLen=   256
      Columns(4).Width=   1746
      Columns(4).Caption=   "Clave"
      Columns(4).Name =   "NoFamilia"
      Columns(4).DataField=   "Column 4"
      Columns(4).FieldLen=   256
      Columns(5).Width=   4763
      Columns(5).Caption=   "Nombre"
      Columns(5).Name =   "NombreFactura"
      Columns(5).DataField=   "Column 5"
      Columns(5).FieldLen=   256
      Columns(6).Width=   2064
      Columns(6).Caption=   "Total"
      Columns(6).Name =   "Total"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   6
      Columns(6).NumberFormat=   "$##,##0.00"
      Columns(6).FieldLen=   256
      Columns(7).Width=   1746
      Columns(7).Caption=   "Cancelada"
      Columns(7).Name =   "Cancelada"
      Columns(7).CaptionAlignment=   2
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   17
      Columns(7).FieldLen=   256
      Columns(7).Style=   2
      Columns(8).Width=   1402
      Columns(8).Caption=   "Turno"
      Columns(8).Name =   "Turno"
      Columns(8).DataField=   "Column 8"
      Columns(8).FieldLen=   256
      Columns(9).Width=   3200
      Columns(9).Caption=   "Usuario"
      Columns(9).Name =   "Usuario"
      Columns(9).DataField=   "Column 9"
      Columns(9).FieldLen=   256
      Columns(10).Width=   3200
      Columns(10).Visible=   0   'False
      Columns(10).Caption=   "Observaciones"
      Columns(10).Name=   "Observaciones"
      Columns(10).DataField=   "Column 10"
      Columns(10).DataType=   8
      Columns(10).FieldLen=   256
      _ExtentX        =   20558
      _ExtentY        =   3625
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
   Begin MSComctlLib.ImageList imglst 
      Left            =   6840
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsDocs.frx":0033
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsDocs.frx":034D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsDocs.frx":0667
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsDocs.frx":0981
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsDocs.frx":0DD3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmTipoDoc 
      Caption         =   "Tipo de Documento"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2295
      Begin VB.OptionButton optRec 
         Caption         =   "Recibos"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optFac 
         Caption         =   "Facturas"
         CausesValidation=   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton optNCred 
         Caption         =   "Notas Credito"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame frmOrden 
      Caption         =   "Ordenar por"
      Height          =   735
      Left            =   5760
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Frame frmBuscar 
      Caption         =   "Buscar"
      Height          =   735
      Left            =   8760
      TabIndex        =   6
      Top             =   1080
      Width           =   3015
      Begin VB.CommandButton cmdBusca 
         Caption         =   "Buscar"
         Default         =   -1  'True
         Height          =   375
         Left            =   2160
         Picture         =   "frmConsDocs.frx":1225
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtBusca 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Observaciones"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6600
      Width           =   3735
   End
End
Attribute VB_Name = "frmConsDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbOrden_Click()
    optSelect
End Sub

Private Sub cmdBusca_Click()
    Dim sColNombre As String
    Dim vOldBM As Variant


    If ssdbgDocs.Rows = 0 Then
        MsgBox "No existen datos para buscar!", vbInformation, "Documentos"
        Exit Sub
    End If

    If Me.txtBusca.Text = "" Then
        MsgBox "No se ha capturado que buscar!", vbInformation
        Exit Sub
    End If


    Err.Clear

    On Error GoTo Error_Catch
    
    
    Select Case Me.cmbOrden.Text
        Case "Clave"
            sColNombre = "NoFamilia"
        Case "Fecha"
            sColNombre = "Fecha"
        Case "Nombre"
            sColNombre = "NombreFactura"
        Case "Número"
            sColNombre = "Numero"
        Case "Folio"
            sColNombre = "Folio"
    End Select
    
    'This code is designed to search for the contents of the
    'text box in column 1 of the DataGrid.
    Dim bm As Variant
    ssdbgDocs.Redraw = False
    ssdbgDocs.MoveFirst
    For i = 0 To ssdbgDocs.Rows - 1
        bm = ssdbgDocs.GetBookmark(i)
        If txtBusca.Text = ssdbgDocs.Columns(sColNombre).CellText(bm) Then
            ssdbgDocs.Bookmark = ssdbgDocs.GetBookmark(i)
            Exit For
        End If
    Next i
    ssdbgDocs.Redraw = True
    

Error_Catch:

    Dim sCadErr As String

    sCadErr = ""

    If Err.Number <> 0 Then
        sCadErr = "Origen: " & Err.Source & vbLf
        sCadErr = sCadErr & "# Error: " & Err.Number & vbLf
        sCadErr = sCadErr & "Descripción: " & Err.Description
        MsgBox sCadErr, vbCritical, "Error"
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 7785
    Me.Width = 12000
    
    
    Me.cmbOrden.AddItem "Clave"
    Me.cmbOrden.AddItem "Fecha"
    Me.cmbOrden.AddItem "Nombre"
    Me.cmbOrden.AddItem "Número"
    Me.cmbOrden.AddItem "Folio"
    
    Me.cmbOrden.Text = "Número"
    
    '05/Dic/2011
    If sDB_NivelUser = 0 Then
        Toolbar1.Buttons("CanFac").Enabled = True
        Toolbar1.Buttons("ImpCupon").Enabled = True
        Toolbar1.Buttons("MuestraCFD").Enabled = True
    End If
    
    CentraForma MDIPrincipal, Me
End Sub

Private Sub optDia_Click()
    optSelect
End Sub

Private Sub optFac_Click()
    optSelect
End Sub

Private Sub optMes_Click()
    optSelect
End Sub

Private Sub optNCred_Click()
    optSelect
End Sub

Private Sub optRec_Click()
    optSelect
End Sub

Private Sub optRecientes_Click()
    optSelect
End Sub

Private Sub optTodas_Click()
    optSelect
End Sub

Private Sub ssdbgDocs_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)
    ActDetalle
End Sub

Private Sub ssdbgDocs_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)
    ActDetalle
End Sub

Private Sub optSelect()
    '03/01/2012 UCM
    Dim strMesesConsulta As String
    strMesesConsulta = ObtieneParametro("MESES_CONSULTA")
    
    Screen.MousePointer = vbHourglass
    
    #If SqlServer_ Then
        If Me.optFac.Value Then
            strSQL = "SET NOCOUNT ON ; SELECT SerieCFD + CONVERT(varchar, FolioCFD) AS Folio, NumeroFactura AS Numero, LOWER(REPLACE(CONVERT(varchar, FechaFactura, 106), ' ', '/')) AS Fecha, CONVERT(varchar, HoraFactura, 108) AS HoraFactura, NoFamilia, NombreFactura, "
        ElseIf Me.optRec.Value Then
            strSQL = "SET NOCOUNT ON ; SELECT Serie + CONVERT(varchar, Folio) AS Folio, NumeroRecibo AS Numero, LOWER(REPLACE(CONVERT(varchar, FechaFactura, 106), ' ', '/')) AS Fecha, CONVERT(varchar, HoraFactura, 108) AS HoraFactura, NoFamilia, NombreFactura, "
        Else
            strSQL = "SET NOCOUNT ON ; SELECT SerieCFD + CONVERT(varchar, FolioCFD) AS Folio, NumeroNota AS Numero, LOWER(REPLACE(CONVERT(varchar, FechaNota, 106), ' ', '/')) AS Fecha, CONVERT(varchar, HoraNota, 108) AS HoraNota, NoFamilia, NombreNota, "
        End If
    
        strSQL = strSQL & " Total, CASE WHEN ISNULL(Cancelada,0) = 0 THEN 0 ELSE -1 END AS Cancelada, Turno, Usuario, Observaciones"
    #Else
        If Me.optFac.Value Then
            strSQL = "SELECT SerieCFD & FolioCFD AS Folio, NumeroFactura AS Numero, Format(FechaFactura,'dd/MMM/yyyy') AS Fecha, HoraFactura, NoFamilia, NombreFactura, "
        ElseIf Me.optRec.Value Then
            strSQL = "SELECT Serie & Folio AS Folio, NumeroRecibo AS Numero, Format(FechaFactura,'dd/MMM/yyyy') AS Fecha, HoraFactura, NoFamilia, NombreFactura, "
        Else
            strSQL = "SELECT SerieCFD & FolioCFD AS Folio, NumeroNota AS Numero, Format(FechaNota,'dd/MMM/yyyy') AS Fecha, HoraNota, NoFamilia, NombreNota, "
        End If
    
        strSQL = strSQL & " Total, Cancelada, Turno, Usuario, Observaciones"
    #End If
    
    If Me.optFac.Value Then
        strSQL = strSQL & " FROM FACTURAS"
    ElseIf Me.optRec.Value Then
        strSQL = strSQL & " FROM RECIBOS"
    Else
        strSQL = strSQL & " FROM NOTAS_CRED"
    End If
    
    If Me.optFac.Value Or Me.optRec.Value Then
        If optDia.Value Then
            strSQL = strSQL & " WHERE FechaFactura >= CONVERT(varchar, GETDATE(), 112)"
        ElseIf optMes.Value Then
            strSQL = strSQL & " WHERE YEAR(FechaFactura) = YEAR(GETDATE()) AND MONTH(FechaFactura) = MONTH(GETDATE())"
        Else
            If strMesesConsulta = "" Then
                strSQL = strSQL & " WHERE (YEAR(FechaFactura) * 12) + MONTH(FechaFactura) >= (YEAR(GETDATE()) * 12) + MONTH(getdate()) - 6"
            Else
                strSQL = strSQL & " WHERE (YEAR(FechaFactura) * 12) + MONTH(FechaFactura) >= (YEAR(GETDATE()) * 12) + MONTH(getdate()) - " + strMesesConsulta
            End If
        End If
    Else
        If optDia.Value Then
            strSQL = strSQL & " WHERE FechaNota >= CONVERT(varchar, GETDATE(), 112)"
        ElseIf optMes.Value Then
            strSQL = strSQL & " WHERE YEAR(FechaNota) = YEAR(GETDATE()) AND MONTH(FechaNota) = MONTH(GETDATE())"
        Else
            strSQL = strSQL & " WHERE YEAR(FechaNota) = YEAR(GETDATE())"
        End If
    End If
    
    ''05/Dic/2011
    ''NO SE MOSTRARÁN FACTURAS CON TURNO ABIERTO
    If sDB_NivelUser <> 0 Then
        If Me.optFac.Value Then
            strSQL = strSQL & " AND ISNULL(IdCorteCaja,0) <> 0"
        End If
    End If
    
    strSQL = strSQL & " ORDER BY "
    
    If Me.optFac.Value Then
        Select Case Me.cmbOrden.Text
            Case "Clave"
                strSQL = strSQL & "NoFamilia" & ", " & "FechaFactura"
            Case "Fecha"
                strSQL = strSQL & "FechaFactura"
            Case "Nombre"
                strSQL = strSQL & "NombreFactura"
            Case "Número"
                strSQL = strSQL & "NumeroFactura"
            Case "Folio"
                strSQL = strSQL & "Folio"
            Case Else
                strSQL = strSQL & "NumeroFactura"
        End Select
    ElseIf Me.optRec.Value Then
        Select Case Me.cmbOrden.Text
            Case "Clave"
                strSQL = strSQL & "NoFamilia"
            Case "Fecha"
                strSQL = strSQL & "FechaFactura"
            Case "Nombre"
                strSQL = strSQL & "NombreFactura"
            Case "Número"
                strSQL = strSQL & "NumeroRecibo"
            Case "Folio"
                strSQL = strSQL & "Folio"
            Case Else
                strSQL = strSQL & "NumeroRecibo"
        End Select
    Else
        Select Case Me.cmbOrden.Text
            Case "Clave"
                strSQL = strSQL & "NoFamilia"
            Case "Fecha"
                strSQL = strSQL & "FechaNota"
            Case "Nombre"
                strSQL = strSQL & "NombreNota"
            Case "Número"
                strSQL = strSQL & "NumeroNota"
            Case "Folio"
                strSQL = strSQL & "Folio"
            Case Else
                strSQL = strSQL & "NumeroNota"
        End Select
    
    End If
    
    LlenaSsDbGrid ssdbgDocs, Conn, strSQL, 11

    ssdbgDocsDet.RemoveAll

    Screen.MousePointer = vbDefault

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim byResp As Byte
    Dim sMensaje As String
    Dim bmCurRow As Variant
    
    Select Case UCase(Button.Key)
        Case "IMPFAC"
            
            '05/Dic/2011 UCM
            If ssdbgDocs.Columns("Cancelada").Text = "" Then
                MsgBox "No hay documento seleccionado!", vbExclamation, "Documentos"
                Exit Sub
            End If
            
            If CBool(ssdbgDocs.Columns("Cancelada").Text) Then
                MsgBox "No se puede Imprimir este documento por que está cancelado!", vbExclamation, "Documentos"
                Exit Sub
            End If

            If IsNull(lNumFolioFacIniImp = ssdbgDocs.Columns("Folio").Text) Then
                MsgBox "No se puede Imprimir este documento por que no tiene asignado Folio!", vbExclamation, "Documentos"
                Exit Sub
            End If

            If Me.optFac.Value Then

                Dim frmImpRec As frmImpFac

                lNumFacIniImp = ssdbgDocs.Columns("Numero").Text
                lNumFacFinImp = ssdbgDocs.Columns("Numero").Text

                lNumFolioFacIniImp = ssdbgDocs.Columns("Folio").Text
                lNumFolioFacFinImp = ssdbgDocs.Columns("Folio").Text
                
                Set frmImpRec = New frmImpFac
                
                frmImpRec.cModo = "R"
                frmImpRec.Tag = "R"
                frmImpRec.lNumeroInicial = lNumFacIniImp
                frmImpRec.lNumeroFinal = lNumFacFinImp
                frmImpRec.Caption = "Imprime Recibos"
                 frmImpRec.Show 1

                Set frmImpRec = Nothing

            ElseIf Me.optRec.Value Then

                Dim frmImpR As New frmImpFac

                lNumRecIniImp = ssdbgDocs.Columns("Numero").Text
                lNumRecFinImp = ssdbgDocs.Columns("Numero").Text

                lNumFacIniImp = ssdbgDocs.Columns("Numero").Text
                lNumFacFinImp = ssdbgDocs.Columns("Numero").Text

                lNumFolioFacIniImp = ssdbgDocs.Columns("Folio").Text
                lNumFolioFacFinImp = ssdbgDocs.Columns("Folio").Text

                frmImpR.Tag = "R"
                frmImpR.lNumeroInicial = lNumRecIniImp
                frmImpR.lNumeroFinal = lNumRecFinImp
                frmImpR.Caption = "Imprime Recibos"
                frmImpR.Show 1

                Set frmImpR = Nothing
            Else
                frmImpNCred.lNumNotaIni = ssdbgDocs.Columns("Numero").Text
                frmImpNCred.Show 1
            End If



        Case "REFRESH"
            optSelect

        Case "CANFAC"

            If (Year(CDate(Me.ssdbgDocs.Columns("Fecha").Value)) * 100 + Month(CDate(Me.ssdbgDocs.Columns("Fecha").Value))) < (Year(Date) * 100 + Month(Date)) Then
                MsgBox "Sólo se pueden cancelar facturas del mes actual", vbCritical, "Error"
                Exit Sub
            End If
            
            If (Me.ssdbgDocs.Columns("Folio").Text <> "") Then
                MsgBox "Sólo se pueden cancelar facturas si no han generado Folio.", vbCritical, "Error"
                Exit Sub
            End If

            If Not ChecaSeguridad(Me.Name, "CANFAC") Then
                Exit Sub
            End If

            If CDate(Me.ssdbgDocs.Columns("Fecha").Value) < Date And (sDB_NivelUser > 0) Then
                MsgBox "No tiene atributos para cancelar" & vbLf & "con fecha anterior a hoy.", vbCritical, "Error"
                Exit Sub
            End If

            If ssdbgDocs.Columns("Cancelada").Value Then '-1
                If Me.optFac.Value Then
                    MsgBox "Este Recibo YA está Cancelada!", vbInformation, "Cancelar"
                Else
                    MsgBox "Este Recibo YA está Cancelado!", vbInformation, "Cancelar"
                End If
                Exit Sub
            End If

            sMensaje = "¿Desea cancelar "

            If Me.optFac.Value Then
                sMensaje = sMensaje & "el Recibo #"
            Else
                sMensaje = sMensaje & "el Recibo #"
            End If

            sMensaje = sMensaje & ssdbgDocs.Columns("Numero").Text & "?"

            byResp = MsgBox(sMensaje, vbYesNo Or vbQuestion, "Cancelar")

            If byResp = vbYes Then

                bmCurRow = ssdbgDocs.Bookmark
                
                If Me.optFac.Value Then
                    CancelaDoc ssdbgDocs.Columns("Numero").Text, 0
                Else
                    CancelaDoc ssdbgDocs.Columns("Numero").Text, 1
                End If
                
                optSelect
                
                ssdbgDocs.Bookmark = bmCurRow

            

                Dim sFolioCFD As String
                Dim sSerieCFD As String
                Dim sRespuesta As String

                'If FolioCFD(ssdbgDocs.Columns("Numero").Text, sFolioCFD, sSerieCFD) = 0 Then
                '    sRespuesta = CancelaCFD(sFolioCFD, sSerieCFD)
                'End If


                'If sRespuesta = "true" Then
                    MsgBox "Recibo Cancelado", vbInformation, "Ok"
                'End If
            End If
        Case "IMPCUPON"

            If Me.optNCred.Value Then
                Exit Sub
            End If

            If ssdbgDocs.Columns("Cancelada").Value Then
                MsgBox "No se puede Imprimir este documento por que está cancelado!", vbExclamation, "Documentos"
                Exit Sub
            End If

            lNumRecIniImp = ssdbgDocs.Columns("Numero").Text
            lNumRecFinImp = ssdbgDocs.Columns("Numero").Text

            lNumFacIniImp = ssdbgDocs.Columns("Numero").Text
            lNumFacFinImp = ssdbgDocs.Columns("Numero").Text


            If Me.optFac.Value Then
                frmImpCupones.Tag = "F"
                frmImpCupones.Caption = "Imprime Cupones"
                frmImpCupones.Show 1
            Else
                frmImpCupones.Tag = "R"
                frmImpCupones.Caption = "Imprime Cupones"
                frmImpCupones.Show 1
            End If

        Case "MUESTRACFD"
            'If Not Me.optFac.Value Then
            '    Exit Sub
            'End If
            
            If Me.ssdbgDocs.Rows = 0 Then
                Exit Sub
            End If

            If IsNull(lNumFolioFacIniImp = ssdbgDocs.Columns("Folio").Text) Then
                MsgBox "No se puede Imprimir este documento por que no tiene asignado Folio!", vbExclamation, "Documentos"
                Exit Sub
            End If

            If Me.optFac.Value Then
                MuestraCFD ssdbgDocs.Columns("Numero").Text, "F"
            Else
                MuestraCFD ssdbgDocs.Columns("Numero").Text, "N"
            End If
            
    End Select
End Sub

Private Sub CancelaDoc(lNumDoc As Long, iModo As Integer)
    Dim adorcsCan As ADODB.Recordset
    Dim adocmdCan As ADODB.Command
    Dim sCancelar As String
    Dim lPointer As Long
    Dim sCadProceso As String
    
    Set adocmdCan = New ADODB.Command
    adocmdCan.ActiveConnection = Conn
    adocmdCan.CommandType = adCmdText
    
    strSQL = "SELECT CadenaCancela1, CadenaCancela2"
    If iModo = 0 Then
        strSQL = strSQL & " FROM FACTURAS_CANCELA INNER JOIN FACTURAS ON FACTURAS_CANCELA.NumeroFactura=FACTURAS.NumeroFactura"
        strSQL = strSQL & " WHERE FACTURAS.NumeroFactura=" & lNumDoc
        strSQL = strSQL & " AND FACTURAS.Cancelada=0"
    Else
        strSQL = strSQL & " FROM RECIBOS_CANCELA INNER JOIN RECIBOS ON RECIBOS_CANCELA.NumeroRecibo=RECIBOS.NumeroRecibo"
        strSQL = strSQL & " WHERE RECIBOS.NumeroRecibo=" & lNumDoc
        strSQL = strSQL & " AND RECIBOS.Cancelada=0"
    End If
    
    Set adorcsCan = New ADODB.Recordset
    adorcsCan.CursorLocation = adUseServer
    
    adorcsCan.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do Until adorcsCan.EOF
        sCancelar = Trim(adorcsCan!CadenaCancela1 & adorcsCan!CadenaCancela2)
        adorcsCan.MoveNext
    Loop
    
    adorcsCan.Close
    Set adorcsCan = Nothing
    
    lPointer = 1
    
    Do Until lPointer >= Len(sCancelar)
        
        Select Case Mid$(sCancelar, lPointer, 1)
            Case "0" 'Cuotas de mantenimiento
                sCadProceso = Mid$(sCancelar, lPointer, 27)
                lPointer = lPointer + 27
                
                #If SqlServer_ Then
                    strSQL = "UPDATE FECHAS_USUARIO SET"
                    strSQL = strSQL & " FechaUltimoPago='" & Format(Mid$(sCadProceso, 18, 10), "yyyymmdd") & "'"
                    strSQL = strSQL & " WHERE IdMember=" & Val(Mid$(sCadProceso, 2, 10))
                    strSQL = strSQL & " AND IdConcepto=" & Val(Mid$(sCadProceso, 12, 6))
                #Else
                    strSQL = "UPDATE FECHAS_USUARIO SET"
                    strSQL = strSQL & " FechaUltimoPago='" & Mid$(sCadProceso, 18, 10) & "'"
                    strSQL = strSQL & " WHERE IdMember=" & Val(Mid$(sCadProceso, 2, 10))
                    strSQL = strSQL & " AND IdConcepto=" & Val(Mid$(sCadProceso, 12, 6))
                #End If
                                
                adocmdCan.CommandText = strSQL
                adocmdCan.Execute
                
            Case "1" 'Rentables
                sCadProceso = Mid$(sCancelar, lPointer, 25)
                lPointer = lPointer + 27
                
                #If SqlServer_ Then
                    strSQL = "UPDATE RENTABLES SET"
                    strSQL = strSQL & " FechaPago='" & Mid$(sCadProceso, 18, 8) & "'"
                    strSQL = strSQL & " WHERE IdUsuario=" & Val(Mid$(sCadProceso, 2, 10))
                    strSQL = strSQL & " AND LTRIM(RTRIM(Numero))='" & Trim(Mid$(sCadProceso, 12, 6)) & "'"
                #Else
                    strSQL = "UPDATE RENTABLES SET"
                    strSQL = strSQL & " FechaPago='" & Mid$(sCadProceso, 18, 10) & "'"
                    strSQL = strSQL & " WHERE IdUsuario=" & Val(Mid$(sCadProceso, 2, 10))
                    strSQL = strSQL & " AND Trim(Numero)='" & Trim(Mid$(sCadProceso, 12, 6)) & "'"
                #End If
                                
                adocmdCan.CommandText = strSQL
                adocmdCan.Execute
                
                
            Case "2" 'Cargos varios
                sCadProceso = Mid$(sCancelar, lPointer, 27)
                lPointer = lPointer + 21
                
                strSQL = "UPDATE CARGOS_VARIOS SET"
                strSQL = strSQL & " Pagado=0"
                strSQL = strSQL & " WHERE IdCargoVario=" & Val(Mid$(sCadProceso, 12, 10))
                strSQL = strSQL & " AND IdMember=" & Val(Mid$(sCadProceso, 2, 10))
                
                adocmdCan.CommandText = strSQL
                adocmdCan.Execute
                
            Case "3" 'Cargos por membresia
                sCadProceso = Mid$(sCancelar, lPointer, 27)
                lPointer = lPointer + 21
                
                #If SqlServer_ Then
                    strSQL = "UPDATE DETALLE_MEM SET"
                    strSQL = strSQL & " DETALLE_MEM.FechaPago=Null,"
                    strSQL = strSQL & " DETALLE_MEM.Observaciones=''"
                    strSQL = strSQL & " FROM DETALLE_MEM INNER JOIN MEMBRESIAS"
                    strSQL = strSQL & " ON DETALLE_MEM.IdMembresia = MEMBRESIAS.IdMembresia"
                    strSQL = strSQL & " WHERE DETALLE_MEM.IdReg=" & Val(Mid$(sCadProceso, 12, 10))
                    strSQL = strSQL & " AND MEMBRESIAS.IdMember=" & Val(Mid$(sCadProceso, 2, 10))
                #Else
                    strSQL = "UPDATE DETALLE_MEM INNER JOIN MEMBRESIAS"
                    strSQL = strSQL & " ON DETALLE_MEM.IdMembresia = MEMBRESIAS.IdMembresia"
                    strSQL = strSQL & " SET"
                    strSQL = strSQL & " DETALLE_MEM.FechaPago=Null,"
                    strSQL = strSQL & " DETALLE_MEM.Observaciones=''"
                    strSQL = strSQL & " WHERE DETALLE_MEM.IdReg=" & Val(Mid$(sCadProceso, 12, 10))
                    strSQL = strSQL & " AND MEMBRESIAS.IdMember=" & Val(Mid$(sCadProceso, 2, 10))
                #End If
                
                adocmdCan.CommandText = strSQL
                adocmdCan.Execute
            Case "9" 'Facturas de varios generadas por el sistema
            
                lPointer = lPointer + 5
                
                strSQL = "UPDATE RECIBOS"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " FACTURA=0"
                strSQL = strSQL & " WHERE"
                strSQL = strSQL & " FACTURA = " & lNumDoc
                
                adocmdCan.CommandText = strSQL
                adocmdCan.Execute
                
        End Select
    Loop
    
    If iModo = 0 Then
        strSQL = "UPDATE FACTURAS"
    Else
        strSQL = "UPDATE RECIBOS"
    End If
    
    #If SqlServer_ Then
        strSQL = strSQL & " SET Cancelada=-1,"
        strSQL = strSQL & " FechaCancelacion='" & Format(Now, "yyyymmdd") & "',"
        strSQL = strSQL & " HoraCancelacion='" & Format(Now, "Hh:Nn:Ss") & "'"
    #Else
        strSQL = strSQL & " SET Cancelada=-1,"
        strSQL = strSQL & " FechaCancelacion='" & Format(Now, "mm/dd/yyyy") & "',"
        strSQL = strSQL & " HoraCancelacion='" & Format(Now, "Hh:Nn:Ss") & "'"
    #End If
    
    If iModo = 0 Then
        strSQL = strSQL & " WHERE NumeroFactura=" & lNumDoc
    Else
        strSQL = strSQL & " WHERE NumeroRecibo=" & lNumDoc
    End If
    
    adocmdCan.CommandText = strSQL
    adocmdCan.Execute
    
    Set adocmdCan = Nothing
    
    If iModo = 0 Then
        MsgBox "La Factura " & lNumDoc & " fue Cancelada", vbExclamation, "Cancela"
    Else
        MsgBox "El Recibo " & lNumDoc & " fue Cancelado", vbExclamation, "Cancela"
    End If
End Sub

Private Sub ActDetalle()

    strSQL = "Select Renglon, IdConcepto, Concepto, Periodo, Cantidad, Total"

    If Me.optFac Then
        strSQL = strSQL & " FROM FACTURAS_DETALLE"
        strSQL = strSQL & " WHERE NumeroFactura=" & ssdbgDocs.Columns("Numero").Text
    ElseIf Me.optRec Then
        strSQL = strSQL & " FROM RECIBOS_DETALLE"
        strSQL = strSQL & " WHERE NumeroRecibo=" & ssdbgDocs.Columns("Numero").Text
    Else
        strSQL = strSQL & " FROM NOTAS_CRED_DETALLE"
        strSQL = strSQL & " WHERE NumeroNota=" & ssdbgDocs.Columns("Numero").Text
    End If

    strSQL = strSQL & " ORDER BY Renglon"

    LlenaSsDbGrid ssdbgDocsDet, Conn, strSQL, 6

    Me.txtObserva.Text = ssdbgDocs.Columns("Observaciones").Text
    
End Sub


