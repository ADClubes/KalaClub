VERSION 5.00
Begin VB.Form frmTurnoAbre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apertura de turno"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   3
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2880
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   2
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   1
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox txtCtrl 
      Height          =   375
      Index           =   0
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Fondo en efectivo (Confirmación)"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Serie Factura Inicial"
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Folio Factura Inicial"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblCtrl 
      Caption         =   "Fondo en efectivo"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmTurnoAbre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim iResp As Integer
    
    Dim adocmdTurnos As ADODB.Command
    Dim lTurno As Long
    
    
    
    If Me.txtCtrl(0).Text = vbNullString Then
        MsgBox "Indicar el fondo inicial", vbInformation, "Verifique"
        Exit Sub
        Me.txtCtrl(0).SetFocus
    End If
    
    If Me.txtCtrl(1).Text = vbNullString Then
        MsgBox "Indicar la confirmación del fondo inicial", vbInformation, "Verifique"
        Exit Sub
        Me.txtCtrl(0).SetFocus
    End If
    
    If Me.txtCtrl(0).Text <> Me.txtCtrl(1).Text Then
        MsgBox "El fondo inicial y su confirmación no son iguales", vbInformation, "Verifique"
        Exit Sub
        Me.txtCtrl(0).SetFocus
    End If
    
    If Me.txtCtrl(2).Text = vbNullString Then
        MsgBox "Indicar el folio inicial", vbInformation, "Verifique"
        Exit Sub
        Me.txtCtrl(2).SetFocus
    End If
    
    If Me.txtCtrl(3).Text = vbNullString Then
        MsgBox "Indicar la serie", vbInformation, "Verifique"
        Exit Sub
        Me.txtCtrl(3).SetFocus
    End If
    
    
    lTurno = OpenShift()
    
    If lTurno > 0 Then
        MsgBox "No se ha cerrado el turno " & lTurno, vbCritical, "Error"
        Exit Sub
    End If
    
    iResp = MsgBox("¿Abrir siguiente turno?", vbQuestion + vbOKCancel, "Confirme")
    
    If iResp = vbCancel Then
        Exit Sub
    End If
    
    lTurno = NextShift()
    
    #If SqlServer_ Then
        strSQL = "INSERT INTO TURNOS ("
        strSQL = strSQL & " FechaApertura,"
        strSQL = strSQL & " HoraApertura,"
        strSQL = strSQL & " UsuarioApertura,"
        '29/06/2007
        strSQL = strSQL & " NumeroCaja,"
        strSQL = strSQL & " TurnoNo,"
        strSQL = strSQL & " FondoApertura,"
        strSQL = strSQL & " FolioApertura,"
        strSQL = strSQL & " SerieApertura)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & "'" & Format(Now, "yyyymmdd") & "',"
        strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "',"
        strSQL = strSQL & "'" & sDB_User & "',"
        '29/06/2007
        strSQL = strSQL & iNumeroCaja & ","
        strSQL = strSQL & lTurno & ","
        '29/06/2010
        strSQL = strSQL & Trim(Me.txtCtrl(0).Text) & ","
        strSQL = strSQL & Trim(Me.txtCtrl(2).Text) & ","
        strSQL = strSQL & "'" & Trim(Me.txtCtrl(3).Text) & "'" & ")"
    #Else
        strSQL = "INSERT INTO TURNOS ("
        strSQL = strSQL & " FechaApertura,"
        strSQL = strSQL & " HoraApertura,"
        strSQL = strSQL & " UsuarioApertura,"
        '29/06/2007
        strSQL = strSQL & " NumeroCaja,"
        strSQL = strSQL & " TurnoNo,"
        strSQL = strSQL & " FondoApertura,"
        strSQL = strSQL & " FolioApertura,"
        strSQL = strSQL & " SerieApertura)"
        strSQL = strSQL & " VALUES ("
        strSQL = strSQL & "#" & Format(Now, "mm/dd/yyyy") & "#,"
        strSQL = strSQL & "'" & Format(Now, "Hh:Nn:Ss") & "',"
        strSQL = strSQL & "'" & sDB_User & "',"
        '29/06/2007
        strSQL = strSQL & iNumeroCaja & ","
        strSQL = strSQL & lTurno & ","
        '29/06/2010
        strSQL = strSQL & Trim(Me.txtCtrl(0).Text) & ","
        strSQL = strSQL & Trim(Me.txtCtrl(2).Text) & ","
        strSQL = strSQL & "'" & Trim(Me.txtCtrl(3).Text) & "'" & ")"
    #End If
    
    Set adocmdTurnos = New ADODB.Command
    adocmdTurnos.ActiveConnection = Conn
    adocmdTurnos.CommandType = adCmdText
    adocmdTurnos.CommandText = strSQL
    adocmdTurnos.Execute
    
    Set adocmdTurnos = Nothing
    
    
    
    MsgBox "Turno " & lTurno & " abierto exitosamente", vbInformation, "Ok"
    
    Unload Me
    
End Sub


Private Sub Form_Load()
    
    CentraForma MDIPrincipal, Me
    
    CargaInformacion
    
End Sub

Private Sub txtCtrl_KeyPress(index As Integer, KeyAscii As Integer)
   
   If index = 0 Or index = 1 Then
        
        Select Case KeyAscii
             Case 8 ' Tecla backspace
                 KeyAscii = KeyAscii
             Case 46 'punto decimal
                 If InStr(Me.txtCtrl(index).Text, ".") Then
                     KeyAscii = 0
                 Else
                     KeyAscii = KeyAscii
                 End If
             Case 48 To 57 ' Numeros del 0 al 9
                 KeyAscii = KeyAscii
             Case Else
                 KeyAscii = 0
         End Select
         
    ElseIf index = 2 Then
        Select Case KeyAscii
             Case 8 ' Tecla backspace
                 KeyAscii = KeyAscii
             Case 48 To 57 ' Numeros del 0 al 9
                 KeyAscii = KeyAscii
             Case Else
                 KeyAscii = 0
         End Select
    Else
        Select Case KeyAscii
             Case 8 ' Tecla backspace
                 KeyAscii = KeyAscii
             Case 65 To 90 ' Letras de la A a la Z
                 KeyAscii = KeyAscii
             Case 97 To 122 ' Letras de la a a la z
                 KeyAscii = Asc(UCase(Chr(KeyAscii)))
             Case Else
                 KeyAscii = 0
         End Select
    End If
    
    
End Sub

Private Sub CargaInformacion()
    Dim strSQL As String
    Dim rsFacturacion As ADODB.Recordset
    Dim sSerie As String
    
    If iNumeroCaja <> 2 Then
        sSerie = ObtieneParametro("SERIE_CFD_FACTURA_CAJA")
    ElseIf iNumeroCaja = 2 Then
        sSerie = ObtieneParametro("SERIE_CFD_FACTURA_DIRE")
    Else
        Exit Sub
    End If
        
    strSQL = "SELECT MAX(CONVERT(int,ISNULL(FolioCFD,0))) + 1 AS FolioCFD, SerieCFD FROM FACTURAS WHERE SerieCFD = '" & sSerie & "' GROUP BY SerieCFD"
    
    Set rsFacturacion = New ADODB.Recordset
    
    rsFacturacion.CursorLocation = adUseServer
    
    rsFacturacion.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    If Not rsFacturacion.EOF Then
        txtCtrl(2).Text = rsFacturacion!FolioCFD
        txtCtrl(3).Text = rsFacturacion!SerieCFD
    End If
    
    rsFacturacion.Close
    Set rsFacturacion = Nothing
End Sub
