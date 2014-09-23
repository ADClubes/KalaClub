VERSION 5.00
Begin VB.Form frmProcMem 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblCount 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "frmProcMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Dim adorcsMem As ADODB.Recordset
    Dim adorcsProc As ADODB.Recordset
    
    Dim adocmdProc As ADODB.Command
    
    
    Dim doMonto As Double
    Dim doEnganche As Double
    Dim doPago1 As Double
    Dim doPago2 As Double
    Dim doPago3 As Double
    Dim doPago4 As Double
    Dim dFechaAlta As Date
    Dim dFecha1 As Date
    Dim dFecha2 As Date
    Dim dFecha3 As Date
    Dim dFecha4 As Date
    Dim lNoPagos As Long
    Dim lDuracion As Long
    Dim lIdTipoMem As Long
    Dim sNomProp As String
    Dim dFechaVig As Date
    
    
    Dim lI As Long
    
    Dim doMontoP As Double
    Dim dfecvence As Date
    Dim lIdReg As Long
    
    Dim sProceso As String
    
    lIdReg = 1
    
    Set adocmdProc = New ADODB.Command
    adocmdProc.ActiveConnection = Conn
    adocmdProc.CommandType = adCmdText
    
    
    Set adorcsProc = New ADODB.Recordset
    adorcsProc.CursorLocation = adUseServer
    
    
    strSQL = "SELECT *"
    strSQL = strSQL & " FROM Titulos_Membresia_dsi"
    strSQL = strSQL & " ORDER BY Mid(Titulo_clave,3)"
    
    
    Set adorcsMem = New ADODB.Recordset
    adorcsMem.CursorLocation = adUseServer
    
    adorcsMem.Open strSQL, Conn, adOpenDynamic, adLockOptimistic
    
    Do Until adorcsMem.EOF
        
        sProceso = "NO SE ENCONTRO INSCRIPCION"
        
        'Busca si ya esta dada de alta la membresia.
        strSQL = "SELECT USU.NoFamilia, USU.IdMember, MEM.IdMembresia"
        strSQL = strSQL & " FROM USUARIOS_CLUB USU LEFT JOIN MEMBRESIAS MEM"
        strSQL = strSQL & " ON USU.IdMember=MEM.IdMember"
        strSQL = strSQL & " WHERE"
        strSQL = strSQL & " Trim(Usu.Inscripcion)='" & Trim(adorcsMem!Titulo_clave) & "'"
        
        adorcsProc.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
        
        
        
        
        'si encuentra el número de inscripcion
        If Not adorcsProc.EOF Then
             
            sProceso = "YA ESTA DADA DE ALTA"
            'Si no esta dada de alta
            If IsNull(adorcsProc!idmembresia) Then
                
                sProceso = "SE DIO DE ALTA"
            
                doMonto = IIf(IsNull(adorcsMem!Precio), 0, adorcsMem!Precio)
                doEnganche = IIf(IsNull(adorcsMem!Enganche), 0, adorcsMem!Enganche)
                doPago1 = IIf(IsNull(adorcsMem!P1), 0, adorcsMem!P1)
                doPago2 = IIf(IsNull(adorcsMem!P2), 0, adorcsMem!P2)
                doPago3 = IIf(IsNull(adorcsMem!P3), 0, adorcsMem!P3)
                doPago4 = IIf(IsNull(adorcsMem!P4), 0, adorcsMem!P4)
                dFecha1 = IIf(IsNull(adorcsMem!F1), CDate(0), adorcsMem!F1)
                dFecha2 = IIf(IsNull(adorcsMem!F2), CDate(0), adorcsMem!F2)
                dFecha3 = IIf(IsNull(adorcsMem!F3), CDate(0), adorcsMem!F3)
                dFecha4 = IIf(IsNull(adorcsMem!F4), CDate(0), adorcsMem!F4)
                dFechaAlta = IIf(IsNull(adorcsMem!Fecha_adquisicion), CDate(0), adorcsMem!Fecha_adquisicion)
                sNomProp = IIf(IsNull(adorcsMem!Propietario), "", adorcsMem!Propietario)
                dFechaVig = IIf(IsNull(adorcsMem!Vigencia), Null, adorcsMem!Vigencia)
                
                
                
                'dFecha1 = Format("mm/dd/yyyy", dFecha1)
                'dFecha2 = Format("mm/dd/yyyy", dFecha2)
                'dFecha3 = Format("mm/dd/yyyy", dFecha3)
                'dFecha4 = Format("mm/dd/yyyy", dFecha4)
                
                'dFechaAlta = Format("mm/dd/yyyy", dFechaAlta)
                'dFechaVig = Format("mm/dd/yyyy", dFechaVig)
                
                lNoPagos = 0
                
                If doMonto > 0 Then
                         
                    If doEnganche = doMonto Then
                        doPago1 = 0
                        doPago2 = 0
                        doPago3 = 0
                        doPago4 = 0
                        
                        lNoPagos = 0
                        
                    Else
                        If doPago1 > 0 Then
                            lNoPagos = 1
                            If doPago2 > 0 Then
                                lNoPagos = 2
                                If doPago3 > 0 Then
                                    lNoPagos = 3
                                    If doPago4 > 0 Then
                                        lNoPagos = 4
                                    Else
                                        doPago4 = 0
                                    End If
                                Else
                                    doPago3 = 0
                                    doPago4 = 0
                                End If
                            Else
                                doPago2 = 0
                                doPago3 = 0
                                doPago4 = 0
                            End If
                        Else
                            doPago1 = 0
                            doPago2 = 0
                            doPago3 = 0
                            doPago4 = 0
                        End If
                    End If
                    
                    
                Else
                    doEnganche = 0
                    doPago1 = 0
                    doPago2 = 0
                    doPago3 = 0
                    doPago4 = 0
                End If
                
                Select Case adorcsMem!tipo_titulo_clave
                    Case "II"
                        lIdTipoMem = 1
                    Case "IM"
                        lIdTipoMem = 2
                    Case "IF"
                        lIdTipoMem = 3
                    Case "IE"
                        lIdTipoMem = 4
                    Case Else
                        lIdTipoMem = 0
                End Select
                
                If IsNull(dFechaVig) Then
                    lDuracion = 0
                Else
                    lDuracion = DateDiff("yyyy", dFechaAlta, dFechaVig)
                End If
                
                
                
                strSQL = "INSERT INTO MEMBRESIAS_T ("
                strSQL = strSQL & " IdMembresia,"
                strSQL = strSQL & " IdMember,"
                strSQL = strSQL & " Monto,"
                strSQL = strSQL & " Enganche,"
                strSQL = strSQL & " FechaAlta,"
                strSQL = strSQL & " NumeroPagos,"
                strSQL = strSQL & " Duracion,"
                strSQL = strSQL & " IdTipoMembresia,"
                strSQL = strSQL & " NombrePropietario,"
                strSQL = strSQL & " IdVendedor,"
                strSQL = strSQL & " MantenimientoIni,"
                strSQL = strSQL & " FechaVigencia)"
                strSQL = strSQL & " Values ("
                strSQL = strSQL & adorcsProc!NoFamilia & ","
                strSQL = strSQL & adorcsProc!Idmember & ","
                strSQL = strSQL & doMonto & ","
                strSQL = strSQL & doEnganche & ","
                strSQL = strSQL & "'" & dFechaAlta & "',"
                strSQL = strSQL & lNoPagos & ","
                strSQL = strSQL & lDuracion & ","
                strSQL = strSQL & lIdTipoMem & ","
                strSQL = strSQL & "'" & sNomProp & "',"
                strSQL = strSQL & 0 & ","
                strSQL = strSQL & "'" & "MC" & "',"
                strSQL = strSQL & "'" & dFechaVig & "')"
                
                
                adocmdProc.CommandText = strSQL
                adocmdProc.Execute
                
                
                For lI = 0 To lNoPagos
                     
                    Select Case lI
                        Case 0
                            doMontoP = doEnganche
                            dfecvence = dFechaAlta
                        Case 1
                            doMontoP = doPago1
                            dfecvence = dFecha1
                        Case 2
                            doMontoP = doPago2
                            dfecvence = dFecha2
                        Case 3
                            doMontoP = doPago3
                            dfecvence = dFecha3
                        Case 4
                            doMontoP = doPago4
                            dfecvence = dFecha4
                    End Select
                    
                    strSQL = "INSERT INTO DETALLE_MEM_T ("
                    strSQL = strSQL & " IdReg,"
                    strSQL = strSQL & " IdMembresia,"
                    strSQL = strSQL & " NoPago,"
                    strSQL = strSQL & " Monto,"
                    strSQL = strSQL & " FechaVence)"
                    strSQL = strSQL & " VALUES ("
                    strSQL = strSQL & lIdReg & ","
                    strSQL = strSQL & adorcsProc!NoFamilia & ","
                    strSQL = strSQL & lI & ","
                    strSQL = strSQL & doMontoP & ","
                    strSQL = strSQL & "'" & dfecvence & "')"
                    
                    adocmdProc.CommandText = strSQL
                    adocmdProc.Execute
                    
                    lIdReg = lIdReg + 1
                Next
            Else
            
            End If
        
        End If
        
        adorcsProc.Close
        
        adorcsMem!pROCESO = sProceso
        adorcsMem.Update
        
        adorcsMem.MoveNext
    Loop
    
    
    Set adorcsProc = Nothing
    
    adorcsMem.Close
    Set adorcsMem = Nothing
    
    Set adocmdProc = Nothing
    
    MsgBox "Terminado!", vbInformation, "Proceso"
    
End Sub

Private Sub Command2_Click()
    Dim adorcsSoc As ADODB.Recordset
    Dim adocmdMem As ADODB.Command
    
    Dim strSQL1 As String
    Dim strSQL2 As String
    
    Dim lIdTipoMembresia As Long
    Dim lIdMembresia As Long
    Dim lIdReg As Long
    
    lIdMembresia = 1
    lIdReg = 1338
    
    
    Set adocmdMem = New ADODB.Command
    adocmdMem.ActiveConnection = Conn
    adocmdMem.CommandType = adCmdText
    
    
    strSQL = "SELECT USUARIOS_CLUB.IdMember, USUARIOS_CLUB.Inscripcion, MEMBRESIAS.IdMember, Usuarios_Club.FechaIngreso, "
    strSQL = strSQL & " USUARIOS_CLUB.A_PATERNO & ' ' & USUARIOS_CLUB.A_MATERNO & ' ' & USUARIOS_CLUB.NOMBRE AS Nombre"
    strSQL = strSQL & " FROM USUARIOS_CLUB LEFT JOIN MEMBRESIAS"
    strSQL = strSQL & " ON USUARIOS_CLUB.IdMember=MEMBRESIAS.IdMember"
    strSQL = strSQL & " WHERE USUARIOS_CLUB.IdMember=USUARIOS_CLUB.IdTitular"
    strSQL = strSQL & " AND MEMBRESIAS.IdMember is Null"
    strSQL = strSQL & " AND USUARIOS_CLUB.Inscripcion is not Null"
    strSQL = strSQL & " ORDER BY USUARIOS_CLUB.IdMember"
    
    
    Set adorcsSoc = New ADODB.Recordset
    adorcsSoc.CursorLocation = adUseServer
    
    adorcsSoc.Open strSQL, Conn, adOpenForwardOnly, adLockReadOnly
    
    Do Until adorcsSoc.EOF
    
        Me.lblCount.Caption = lIdMembresia
        DoEvents
    
        lIdTipoMembresia = 0
    
        If Len(Trim(adorcsSoc!inscripcion)) > 0 Then
            If Left$(adorcsSoc!inscripcion, 2) = "IF" Then
                lIdTipoMembresia = 3
            ElseIf Left$(adorcsSoc!inscripcion, 2) = "IM" Then
                lIdTipoMembresia = 2
            ElseIf Left$(adorcsSoc!inscripcion, 2) = "II" Then
                lIdTipoMembresia = 1
            ElseIf Left$(adorcsSoc!inscripcion, 2) = "IE" Then
                lIdTipoMembresia = 4
            End If
        End If
        
    
        
        If lIdTipoMembresia > 0 Then
    
            #If SqlServer_ Then
                strSQL1 = "INSERT INTO MEMBRESIAS ("
                strSQL1 = strSQL1 & " IdMembresia,"
                strSQL1 = strSQL1 & " IdMember,"
                strSQL1 = strSQL1 & " Monto,"
                strSQL1 = strSQL1 & " Enganche,"
                strSQL1 = strSQL1 & " FechaAlta,"
                strSQL1 = strSQL1 & " NumeroPagos,"
                strSQL1 = strSQL1 & " Duracion,"
                strSQL1 = strSQL1 & " IdTipoMembresia,"
                strSQL1 = strSQL1 & " NombrePropietario,"
                strSQL1 = strSQL1 & " Observaciones,"
                strSQL1 = strSQL1 & " MantenimientoIni)"
                strSQL1 = strSQL1 & " VALUES ("
                strSQL1 = strSQL1 & lIdMembresia & ","
                strSQL1 = strSQL1 & adorcsSoc.Fields("USUARIOS_CLUB.IdMember") & ","
                strSQL1 = strSQL1 & 0.01 & ","
                strSQL1 = strSQL1 & 0.01 & ","
                strSQL1 = strSQL1 & "'" & Format(adorcsSoc!FechaIngreso, "yyyymmdd") & "',"
                strSQL1 = strSQL1 & 0 & ","
                strSQL1 = strSQL1 & 95 & ","
                strSQL1 = strSQL1 & lIdTipoMembresia & ","
                strSQL1 = strSQL1 & "'" & adorcsSoc!Nombre & "',"
                strSQL1 = strSQL1 & "'" & "GENERADO POR SISTEMA" & "',"
                strSQL1 = strSQL1 & "'" & "MC" & "')"
                
                strSQL2 = "INSERT INTO DETALLE_MEM ("
                strSQL2 = strSQL2 & " IdReg,"
                strSQL2 = strSQL2 & " IdMembresia,"
                strSQL2 = strSQL2 & " NoPago,"
                strSQL2 = strSQL2 & " Monto,"
                strSQL2 = strSQL2 & " FechaVence,"
                strSQL2 = strSQL2 & " FechaPago,"
                strSQL2 = strSQL2 & " Observaciones)"
                strSQL2 = strSQL2 & " VALUES ("
                strSQL2 = strSQL2 & lIdReg & ","
                strSQL2 = strSQL2 & lIdMembresia & ","
                strSQL2 = strSQL2 & 0 & ","
                strSQL2 = strSQL2 & 0.01 & ","
                strSQL2 = strSQL2 & "'" & Format(adorcsSoc!FechaIngreso, "yyyymmdd") & "',"
                strSQL2 = strSQL2 & "'" & Format(adorcsSoc!FechaIngreso, "yyyymmdd") & "',"
                strSQL2 = strSQL2 & "'" & "SISTEMA" & "')"
            #Else
                strSQL1 = "INSERT INTO MEMBRESIAS ("
                strSQL1 = strSQL1 & " IdMembresia,"
                strSQL1 = strSQL1 & " IdMember,"
                strSQL1 = strSQL1 & " Monto,"
                strSQL1 = strSQL1 & " Enganche,"
                strSQL1 = strSQL1 & " FechaAlta,"
                strSQL1 = strSQL1 & " NumeroPagos,"
                strSQL1 = strSQL1 & " Duracion,"
                strSQL1 = strSQL1 & " IdTipoMembresia,"
                strSQL1 = strSQL1 & " NombrePropietario,"
                strSQL1 = strSQL1 & " Observaciones,"
                strSQL1 = strSQL1 & " MantenimientoIni)"
                strSQL1 = strSQL1 & " VALUES ("
                strSQL1 = strSQL1 & lIdMembresia & ","
                strSQL1 = strSQL1 & adorcsSoc.Fields("USUARIOS_CLUB.IdMember") & ","
                strSQL1 = strSQL1 & 0.01 & ","
                strSQL1 = strSQL1 & 0.01 & ","
                strSQL1 = strSQL1 & "#" & Format(adorcsSoc!FechaIngreso, "mm/dd/yyyy") & "#,"
                strSQL1 = strSQL1 & 0 & ","
                strSQL1 = strSQL1 & 95 & ","
                strSQL1 = strSQL1 & lIdTipoMembresia & ","
                strSQL1 = strSQL1 & "'" & adorcsSoc!Nombre & "',"
                strSQL1 = strSQL1 & "'" & "GENERADO POR SISTEMA" & "',"
                strSQL1 = strSQL1 & "'" & "MC" & "')"
                
                
                strSQL2 = "INSERT INTO DETALLE_MEM ("
                strSQL2 = strSQL2 & " IdReg,"
                strSQL2 = strSQL2 & " IdMembresia,"
                strSQL2 = strSQL2 & " NoPago,"
                strSQL2 = strSQL2 & " Monto,"
                strSQL2 = strSQL2 & " FechaVence,"
                strSQL2 = strSQL2 & " FechaPago,"
                strSQL2 = strSQL2 & " Observaciones)"
                strSQL2 = strSQL2 & " VALUES ("
                strSQL2 = strSQL2 & lIdReg & ","
                strSQL2 = strSQL2 & lIdMembresia & ","
                strSQL2 = strSQL2 & 0 & ","
                strSQL2 = strSQL2 & 0.01 & ","
                strSQL2 = strSQL2 & "#" & Format(adorcsSoc!FechaIngreso, "mm/dd/yyyy") & "#,"
                strSQL2 = strSQL2 & "#" & Format(adorcsSoc!FechaIngreso, "mm/dd/yyyy") & "#,"
                strSQL2 = strSQL2 & "'" & "SISTEMA" & "')"
            #End If
            
            adocmdMem.CommandText = strSQL1
            adocmdMem.Execute
            
            adocmdMem.CommandText = strSQL2
            adocmdMem.Execute
        
        End If
        
        adorcsSoc.MoveNext
        
        lIdMembresia = lIdMembresia + 1
        lIdReg = lIdReg + 1
        
    Loop
    
    adorcsSoc.Close
    Set adorcsSoc = Nothing
    
    Set adocmdMem = Nothing
    
    MsgBox "Proceso concluido!", vbInformation, "Mensaje"
    
    
End Sub
