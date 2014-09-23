Attribute VB_Name = "License"
Option Explicit

Public Function CheckLicense() As Boolean
    
    Dim HL_Coded_Data(8) As Byte
    
    Dim sCad As String
    Dim lI As Long
    
    Dim lValor As Long
    
    Dim dFechaMax As Date
    
    
    HL_Coded_Data(0) = &H0
    HL_Coded_Data(1) = &H0
    HL_Coded_Data(2) = &H0
    HL_Coded_Data(3) = &H0
    HL_Coded_Data(4) = &H0
    HL_Coded_Data(5) = &H0
    HL_Coded_Data(6) = &H0
    HL_Coded_Data(7) = &H5
    
    sCad = ""
    
    
    'Checa si existe el archivo
    If Dir(sDB_DataSource & "\" & "license.kls") = "" Then
       CheckLicense = False
       Exit Function
    End If
    
    
    
    Open sDB_DataSource & "\" & "license.kls" For Input As #1
    
    Do While Not EOF(1)
        Input #1, sCad
    Loop
    
    Close #1
    
    If Len(sCad) <> 24 Then
        CheckLicense = False
        Exit Function
    End If
    
    
    Err.Clear
    On Error Resume Next
    For lI = 0 To 7
        HL_Coded_Data(lI) = Val("&H" & Mid$(sCad, lI * 3 + 1, 3))
        If Err.Number <> 0 Then
            HL_Coded_Data(lI) = 0
            Err.Clear
        End If
    Next
    On Error GoTo 0
    
    hlresult = HL_CODE(HL_Coded_Data(0), 1)
    
    If hlresult <> 0 Then
        CheckLicense = False
        Exit Function
    End If
    
    dFechaMax = LicenseDate(HL_Coded_Data)
    
    If dFechaMax < Date Then
        CheckLicense = False
        Exit Function
    End If
    
    
    
    CheckLicense = True
    
End Function


Public Function LicenseDate(aDatos() As Byte) As Date
    Dim lI As Integer
    Dim sCadena As String
    
    sCadena = ""
    
    For lI = 0 To 7
        sCadena = sCadena & Chr$(aDatos(lI))
    Next
    
    Err.Clear
    On Error Resume Next
    LicenseDate = CDate(Left$(sCadena, 2) & "/" & Mid$(sCadena, 3, 2) & "/" & Right$(sCadena, 4))
    
    If Err.Number <> 0 Then
        LicenseDate = CDate("01/01/1990")
    End If
    
    On Error GoTo 0
    
End Function
