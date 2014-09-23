VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNomina 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker dtpFechaProceso 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   21561345
      CurrentDate     =   39252
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
End
Attribute VB_Name = "frmNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
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
    
    
    strSQL = "SELECT DISTINCT INSTRUCTORES.Empresa"
    strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
    strSQL = strSQL & " WHERE ("
    strSQL = strSQL & " ((Cupones.FechaPago) = " & "#" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#)"
    strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
    strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
    
    Set adorcsPatron = New ADODB.Recordset
    adorcsPatron.CursorLocation = adUseServer
    adorcsPatron.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do Until adorcsPatron.EOF
    
    
        strSQL = "SELECT DISTINCT INSTRUCTORES.Nomina"
        strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
        strSQL = strSQL & " WHERE ("
        strSQL = strSQL & " ((Instructores.Empresa) =" & adorcsPatron!Empresa & ")"
        strSQL = strSQL & " AND ((Cupones.FechaPago) = " & "#" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#)"
        strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
        strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
    
    
        Set adorcsNumero = New ADODB.Recordset
        adorcsNumero.CursorLocation = adUseServer
        adorcsNumero.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    
        Set adorcsNomina = New ADODB.Recordset
        adorcsNomina.CursorLocation = adUseServer

    
        Do Until adorcsNumero.EOF
    
            strSQL = "SELECT INSTRUCTORES.NoEmpleado, INSTRUCTORES.ClavePercepcion, Sum(CUPONES.ImporteAPagar) AS Importe"
            strSQL = strSQL & " FROM CUPONES INNER JOIN INSTRUCTORES ON CUPONES.IdInstructorAplicacion = INSTRUCTORES.IdInstructor"
            strSQL = strSQL & " WHERE ("
            strSQL = strSQL & " ((Cupones.FechaPago) = " & "#" & Format(Me.dtpFechaProceso.Value, "mm/dd/yyyy") & "#)"
            strSQL = strSQL & " AND ((Cupones.Pagado) = 0)"
            strSQL = strSQL & " AND ((Cupones.FechaAplicacion) Is Not Null))"
            strSQL = strSQL & " AND INSTRUCTORES.Empresa=" & adorcsPatron!Empresa
            strSQL = strSQL & " AND INSTRUCTORES.Nomina=" & adorcsNumero!Nomina
            strSQL = strSQL & " GROUP BY INSTRUCTORES.NoEmpleado, INSTRUCTORES.ClavePercepcion"
    
            
            adorcsNomina.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
            'sDir = ObtieneParametro("RUTA_ARCHIVO_NOMINA") '"d:\kalaclub"
            sDir = App.Path
            
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
    
                outputFile.writeline sRenglon
        
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
    
    
    
    Exit Sub
    
    
    
    
  

End Sub

Private Sub Form_Load()
    If Not Connection_DB Then
        End
    End If
    
    Me.dtpFechaProceso.Value = Date
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Conn.Close
    Set Conn = Nothing
End Sub
