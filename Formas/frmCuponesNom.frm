VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmCuponesNom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genera Archivo para Nómina"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   7200
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox txtControl 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   5280
      TabIndex        =   6
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdBusca 
      Caption         =   "Busca"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBGrid ssdbgNomina 
      Height          =   3015
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   8895
      _Version        =   196616
      DataMode        =   2
      Col.Count       =   6
      AllowUpdate     =   0   'False
      RowHeight       =   423
      Columns.Count   =   6
      Columns(0).Width=   1508
      Columns(0).Caption=   "Empresa"
      Columns(0).Name =   "Empresa"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(1).Width=   1217
      Columns(1).Caption=   "Nomina"
      Columns(1).Name =   "Nomina"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(2).Width=   2143
      Columns(2).Caption=   "NoEmpleado"
      Columns(2).Name =   "NoEmpleado"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(3).Width=   6138
      Columns(3).Caption=   "Nombre"
      Columns(3).Name =   "Nombre"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(4).Width=   1773
      Columns(4).Caption=   "# Cupones"
      Columns(4).Name =   "Cupones"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   3
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "Importe"
      Columns(5).Name =   "Importe"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   6
      Columns(5).NumberFormat=   "CURRENCY"
      Columns(5).FieldLen=   256
      _ExtentX        =   15690
      _ExtentY        =   5318
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
   Begin MSComCtl2.DTPicker dtpFechaProceso 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   59310081
      CurrentDate     =   39106
   End
   Begin VB.CommandButton cmdGenera 
      Caption         =   "Genera Nómina"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      Caption         =   "Total Instructores"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      Caption         =   "Total Importe"
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblControl 
      Alignment       =   2  'Center
      Caption         =   "Total Cupones"
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmCuponesNom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBusca_Click()

    Dim adorcsNomina As ADODB.Recordset
    
    Dim lTotalCupones As Long
    Dim dTotalImporte As Double
    
    
    Me.cmdBusca.Enabled = False
    
    #If SqlServer_ Then
        strSQL = "SELECT INSTRUCTORES.Empresa, INSTRUCTORES.Nomina, INSTRUCTORES.NoEmpleado, Instructores.Nombre + ' ' + Instructores.Apellido_Paterno + ' ' + Instructores.Apellido_Materno AS Nombre, Count(CUPONES.FolioCupon) AS Cupones, Sum(CUPONES.ImporteAPagar) AS Importe"
        strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " Cupones.FechaPago = " & "'" & Format(Me.dtpFechaProceso.Value, "yyyymmdd") & "'"
        strSQL = strSQL & " AND Cupones.Pagado = 0"
        strSQL = strSQL & " AND Cupones.FechaAplicacion Is Not Null"
        strSQL = strSQL & " GROUP BY INSTRUCTORES.Empresa, INSTRUCTORES.Nomina, INSTRUCTORES.NoEmpleado, Instructores.Nombre + ' ' + Instructores.Apellido_Paterno + ' ' + Instructores.Apellido_Materno"
    #Else
        strSQL = "SELECT INSTRUCTORES.Empresa, INSTRUCTORES.Nomina, INSTRUCTORES.NoEmpleado, Instructores.Nombre & ' ' & Instructores.Apellido_Paterno & ' ' & Instructores.Apellido_Materno AS Nombre, Count(CUPONES.FolioCupon) AS Cupones, Sum(CUPONES.ImporteAPagar) AS Importe"
        strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((Cupones.FechaPago) = " & "#" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
        strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
        strSQL = strSQL & " GROUP BY INSTRUCTORES.Empresa, INSTRUCTORES.Nomina, INSTRUCTORES.NoEmpleado, Instructores.Nombre & ' ' & Instructores.Apellido_Paterno & ' ' & Instructores.Apellido_Materno"
    #End If
    
    Set adorcsNomina = New ADODB.Recordset
    
    adorcsNomina.CursorLocation = adUseServer
    
    adorcsNomina.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    Me.ssdbgNomina.RemoveAll
    lTotalCupones = 0
    dTotalImporte = 0
    
    If adorcsNomina.EOF Then
        MsgBox "No hay datos con la fecha especificada", vbExclamation, "Verifique"
        Me.cmdBusca.Enabled = True
        Exit Sub
    End If
    
    Do Until adorcsNomina.EOF
        Me.ssdbgNomina.AddItem adorcsNomina!Empresa & vbTab & adorcsNomina!Nomina & vbTab & adorcsNomina!NoEmpleado & vbTab & adorcsNomina!Nombre & vbTab & adorcsNomina!Cupones & vbTab & adorcsNomina!Importe
        
        lTotalCupones = lTotalCupones + adorcsNomina!Cupones
        dTotalImporte = dTotalImporte + adorcsNomina!Importe
    
        
        adorcsNomina.MoveNext
    Loop
    
    
    adorcsNomina.Close
    
    Set adorcsNomina = Nothing
    
    Me.txtControl(0).Text = lTotalCupones
    Me.txtControl(1).Text = Format(dTotalImporte, "$#,#0.00")
    Me.txtControl(2).Text = Me.ssdbgNomina.Rows
    
    If Me.ssdbgNomina.Rows > 0 Then
        HabilitaControles True
    End If
    
    
    
    
End Sub

Private Sub cmdGenera_Click()
    
    Dim adorcsPatron As ADODB.Recordset
    Dim adorcsNumero As ADODB.Recordset
    Dim adorcsNomina As ADODB.Recordset
    Dim fs As Object
    Dim outputFile As Object
    
    
    Dim sDir As String
    Dim sFileName As String
    Dim sRenglon As String
    
    Dim sNumeroClub As String
    Dim sNumeroPatron As String
    
    Dim sClavePercepcion As String
    
    Dim sMensaje As String
    
    
    Dim adocmdNomina As ADODB.Command
    Dim lResult As Long
    
    #If SqlServer_ Then
        strSQL = "SELECT DISTINCT INSTRUCTORES.Empresa"
        strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((Cupones.FechaPago) = " & "'" & Format(Me.dtpFechaProceso.Value, "yyyymmdd") & "')"
        strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
        strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
    #Else
        strSQL = "SELECT DISTINCT INSTRUCTORES.Empresa"
        strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((Cupones.FechaPago) = " & "#" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
        strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
    #End If
    
    Set adorcsPatron = New ADODB.Recordset
    adorcsPatron.CursorLocation = adUseServer
    adorcsPatron.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do Until adorcsPatron.EOF
    
        #If SqlServer_ Then
            strSQL = "SELECT DISTINCT INSTRUCTORES.Nomina"
            strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((Instructores.Empresa) =" & adorcsPatron!Empresa & ")"
            strSQL = strSQL & " AND ((Cupones.FechaPago) = '" & Format(Me.dtpFechaProceso.Value, "yyyymmdd") & "')"
            strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
            strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
        #Else
            strSQL = "SELECT DISTINCT INSTRUCTORES.Nomina"
            strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((Instructores.Empresa) =" & adorcsPatron!Empresa & ")"
            strSQL = strSQL & " AND ((Cupones.FechaPago) = " & "#" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
            strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
        #End If
    
        Set adorcsNumero = New ADODB.Recordset
        adorcsNumero.CursorLocation = adUseServer
        adorcsNumero.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
        Set adorcsNomina = New ADODB.Recordset
        adorcsNomina.CursorLocation = adUseServer

    
        Do Until adorcsNumero.EOF
    
            #If SqlServer_ Then
                strSQL = "SELECT INSTRUCTORES.NoEmpleado, INSTRUCTORES.ClavePercepcion, Sum(CUPONES.ImporteAPagar) AS Importe"
                strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
                strSQL = strSQL & " WHERE ("
                strSQL = strSQL & " ((Cupones.FechaPago) = '" & Format(Me.dtpFechaProceso.Value, "yyyymmdd") & "')"
                strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
                strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
                strSQL = strSQL & " AND INSTRUCTORES.Empresa=" & adorcsPatron!Empresa
                strSQL = strSQL & " AND INSTRUCTORES.Nomina=" & adorcsNumero!Nomina
                strSQL = strSQL & " GROUP BY INSTRUCTORES.NoEmpleado, INSTRUCTORES.ClavePercepcion"
            #Else
                strSQL = "SELECT INSTRUCTORES.NoEmpleado, INSTRUCTORES.ClavePercepcion, Sum(CUPONES.ImporteAPagar) AS Importe"
                strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
                strSQL = strSQL & " WHERE ("
                strSQL = strSQL & " ((Cupones.FechaPago) = " & "#" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#)"
                strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
                strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
                strSQL = strSQL & " AND INSTRUCTORES.Empresa=" & adorcsPatron!Empresa
                strSQL = strSQL & " AND INSTRUCTORES.Nomina=" & adorcsNumero!Nomina
                strSQL = strSQL & " GROUP BY INSTRUCTORES.NoEmpleado, INSTRUCTORES.ClavePercepcion"
            #End If
            
            adorcsNomina.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
            sDir = ObtieneParametro("RUTA_ARCHIVO_NOMINA") '"d:\kalaclub"
            sNumeroClub = ObtieneParametro("NUMERO_CLUB")  '"1"
            sNumeroPatron = adorcsPatron!Empresa
    
            sFileName = sDir & "\" & sNumeroClub & sNumeroPatron & adorcsNumero!Nomina & Format(Me.dtpFechaProceso.Value, "ddmmyy") & ".dat"
    
            sMensaje = sMensaje & vbLf & sFileName
    
            Set fs = CreateObject("Scripting.FileSystemObject")
    
            'Si el archivo existe
            If Dir(sFileName) = sFileName Then
        
            End If
    
            Set outputFile = fs.CreateTextFile(sFileName)
                    
    
            sClavePercepcion = ObtieneParametro("CLAVE_PERCEP_ENT_PER") '"511"
    
            Do Until adorcsNomina.EOF
                sRenglon = vbNullString
                sRenglon = sRenglon & Format(adorcsNomina!NoEmpleado, "@@@@@")
                sRenglon = sRenglon & "P"
                sRenglon = sRenglon & adorcsNomina!ClavePercepcion
                sRenglon = sRenglon & "4"
                sRenglon = sRenglon & Format(Me.dtpFechaProceso.Value, "ddmmyy")
                sRenglon = sRenglon & Format(Me.dtpFechaProceso.Value, "ddmmyy")
                sRenglon = sRenglon & "N"
                sRenglon = sRenglon & Space(11)
                sRenglon = sRenglon & Space(11)
                sRenglon = sRenglon & Format(Format(adorcsNomina!Importe, "#0.00"), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
    
                outputFile.WriteLine sRenglon
        
                adorcsNomina.MoveNext
            Loop
            adorcsNomina.Close
            adorcsNumero.MoveNext
        Loop
    
        adorcsNumero.Close
        adorcsPatron.MoveNext
    Loop
    Set adorcsPatron = Nothing
    Set adorcsNumero = Nothing
    Set adorcsNomina = Nothing
    
    
    MsgBox "Se creo el archivo " & vbLf & sMensaje, vbInformation, "Correcto"
    
    
    
    lResult = MsgBox("¿Actualizar la marca de pagado?", vbOKCancel + vbQuestion, "Confirme")
    
    If lResult = vbOK Then
    
        #If SqlServer_ Then
            strSQL = "UPDATE CUPONES SET"
            strSQL = strSQL & " CUPONES.Pagado = -1"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & "((CUPONES.Pagado)=0)"
            strSQL = strSQL & " And ((CUPONES.FechaPago)= '" & Format(Me.dtpFechaProceso.Value, "yyyymmdd") & "')"
            strSQL = strSQL & ")"
        #Else
            strSQL = "UPDATE CUPONES SET"
            strSQL = strSQL & " CUPONES.Pagado = -1"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & "((CUPONES.Pagado)=0)"
            strSQL = strSQL & " And ((CUPONES.FechaPago)=#" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & ")"
        #End If
    
        Set adocmdNomina = New ADODB.Command
        adocmdNomina.ActiveConnection = Conn
        adocmdNomina.CommandType = adCmdText
        adocmdNomina.CommandText = strSQL
    
        adocmdNomina.Execute lResult
    
        If lResult Then
            MsgBox "Se afectarón " & lResult & " registro(s)", vbInformation, "Correcto"
        Else
            MsgBox "Ocurrio un error al actualizar la marca de pagado!", vbExclamation, "Error"
        End If
    
        Set adocmdNomina = Nothing
        
    End If
    
    HabilitaControles False
    Me.ssdbgNomina.RemoveAll
    Me.cmdBusca.Enabled = True
    
    Exit Sub
    
    
    
    
  
End Sub

Private Sub Form_Load()
    
    CentraForma MDIPrincipal, Me
    Me.dtpFechaProceso.Value = Date
    HabilitaControles False
    
End Sub

Private Sub HabilitaControles(bValor As Boolean)
    
    
    Me.lblControl(0).Visible = bValor
    Me.lblControl(1).Visible = bValor
    Me.lblControl(2).Visible = bValor
    
    Me.txtControl(0).Visible = bValor
    Me.txtControl(1).Visible = bValor
    Me.txtControl(2).Visible = bValor
    
    Me.cmdGenera.Visible = bValor
    
    
End Sub
