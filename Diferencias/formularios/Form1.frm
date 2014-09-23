VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Enabled         =   0   'False
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label lblAvance 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Dim sConn As String
    
    
    Dim Conn As ADODB.Connection
    
    Dim adorcsRecibo As ADODB.Recordset
    Dim adorcsFactura As ADODB.Recordset
    
    Dim adocmdDifer As ADODB.Command
    
    Dim sDB_DataSource  As String
    Dim sDB As String
    Dim sDB_PW As String
    
    Dim strSql As String
    
    
    sDB_DataSource = "c:\datos sportium\base san angel\2008 12"
    sDB = "kalaclub.mdb"
    
    
     sConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
              "Data Source=" & sDB_DataSource & "\" & sDB & ";" & _
              "Persist Security Info=False;" & _
              "Jet OLEDB:Database Password=eUdomilia2006;" & _
              "User Id=Admin;" & _
              "Password=" & sDB_PW & ";"
              
              
    Set Conn = New ADODB.Connection
              
              
    Conn.Errors.Clear
    Err.Clear
    
    Conn.CursorLocation = adUseServer
    Conn.Open sConn
    
    strSql = "SELECT RECIBOS_DETALLE.NumeroRecibo, RECIBOS.FechaFactura, RECIBOS.NoFamilia, RECIBOS_DETALLE.IdConcepto, RECIBOS_DETALLE.Concepto, RECIBOS_DETALLE.Total, RECIBOS.Cancelada"
    strSql = strSql & " FROM RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo = RECIBOS.NumeroRecibo"
    strSql = strSql & " WHERE (((RECIBOS.FechaFactura) Between #1/1/2008# And #12/31/2008#)"
    strSql = strSql & " AND ((RECIBOS.NoFamilia) Not IN (472) ))"
    strSql = strSql & " ORDER BY RECIBOS_DETALLE.NumeroRecibo;"

    
    
    
    Set adorcsRecibo = New ADODB.Recordset
    
    adorcsRecibo.CursorLocation = adUseServer
    
    adorcsRecibo.Open strSql, Conn, adOpenForwardOnly, adLockReadOnly
    
    
    
    Set adorcsFactura = New ADODB.Recordset
    
    adorcsFactura.CursorLocation = adUseServer
    
    
    Set adocmdDifer = New ADODB.Command
    adocmdDifer.ActiveConnection = Conn
    adocmdDifer.CommandType = adCmdText
    
    strSql = "DELETE * FROM DIFERENCIAS"
    
    adocmdDifer.CommandText = strSql
    adocmdDifer.Execute
    
    Do While Not adorcsRecibo.EOF
    
        Me.lblAvance.Caption = adorcsRecibo!FechaFactura
        DoEvents
        
        strSql = "SELECT FACTURAS.NumeroFactura"
        strSql = strSql & " FROM FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura"
        strSql = strSql & " WHERE (((FACTURAS.FechaFactura)=#" & Format(adorcsRecibo!FechaFactura, "mm/dd/yyyy") & "#)"
        strSql = strSql & " AND ((FACTURAS.NoFamilia)=" & adorcsRecibo!NoFamilia & ")"
        strSql = strSql & " AND ((FACTURAS_DETALLE.IdConcepto) In (799,899)))"
    
        adorcsFactura.Open strSql, Conn, adOpenForwardOnly, adLockReadOnly
        
        If Not adorcsFactura.EOF Then
        
        
            'Inserta el recibo
            strSql = "INSERT INTO DIFERENCIAS (Fecha, Inscripcion, Documento, NumeroDocumento, IdConcepto, Descripcion, Importe, Cancelado)"
            strSql = strSql & " SELECT RECIBOS.FechaFactura,RECIBOS.NoFamilia, 'R' AS Documento, RECIBOS_DETALLE.NumeroRecibo, RECIBOS_DETALLE.IdConcepto, RECIBOS_DETALLE.Concepto, RECIBOS_DETALLE.Total, RECIBOS.Cancelada"
            strSql = strSql & " FROM RECIBOS_DETALLE INNER JOIN RECIBOS ON RECIBOS_DETALLE.NumeroRecibo = RECIBOS.NumeroRecibo"
            strSql = strSql & " WHERE (((RECIBOS.FechaFactura)=#" & Format(adorcsRecibo!FechaFactura, "mm/dd/yyyy") & "#)"
            strSql = strSql & " AND ((RECIBOS.NumeroRecibo)=" & adorcsRecibo!NumeroRecibo & ")"
            strSql = strSql & " AND ((RECIBOS_DETALLE.IdConcepto)=" & adorcsRecibo!IdConcepto & "))"
            
            adocmdDifer.CommandText = strSql
            adocmdDifer.Execute
        
            strSql = "INSERT INTO DIFERENCIAS (Fecha, Inscripcion, Documento, NumeroDocumento, IdConcepto, Descripcion, Importe, Cancelado)"
            strSql = strSql & " SELECT  FACTURAS.FechaFactura,  FACTURAS.NoFamilia, 'F' as Documento,  FACTURAS.NumeroFactura,  FACTURAS_DETALLE.IdConcepto,  FACTURAS_DETALLE.Concepto, FACTURAS_DETALLE.Total,FACTURAS.Cancelada"
            strSql = strSql & " FROM FACTURAS INNER JOIN FACTURAS_DETALLE ON FACTURAS.NumeroFactura = FACTURAS_DETALLE.NumeroFactura"
            strSql = strSql & " WHERE (((FACTURAS.FechaFactura)=#" & Format(adorcsRecibo!FechaFactura, "mm/dd/yyyy") & "#)"
            strSql = strSql & " AND ((FACTURAS.NoFamilia)=" & adorcsRecibo!NoFamilia & ")"
            strSql = strSql & " AND ((FACTURAS_DETALLE.IdConcepto)  In (799,899)))"
            
            adocmdDifer.CommandText = strSql
            adocmdDifer.Execute
        
        End If
        
        
        
        adorcsFactura.Close
    
        adorcsRecibo.MoveNext
    Loop
    
    
    Set adocmdDifer = Nothing
    

    Set adorcsFactura = Nothing
    
    
    adorcsRecibo.Close
    Set adorcsRecibo = Nothing
    
    
    Conn.Close
    Set Conn = Nothing
    
    
    Me.cmdsalir.Enabled = True
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
