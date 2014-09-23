Attribute VB_Name = "PrinterFunctions"
Option Explicit

'Public Enum ModoTexto
'    prnTextLeft = 0
'    prnTextCenter = 1
'    prnTextRight = 2
'End Enum



Public Sub ImprimeTexto(sCadTxt As String, iModo As Integer, lPosIni As Long, lPosFin As Long)

    Dim lX As Long
    Dim lIni As Long
    Dim lFin As Long
    Dim sCad As String
    Dim sRenglon As String
    
    



    'El texto no cabe en una sola línea
    If Printer.TextWidth(sCadTxt) > lPosFin - lPosIni Then
        sCad = sCadTxt
        lIni = 1
        
        Do While lIni <= Len(sCadTxt)
        
        
            lFin = BuscaEspacioTextoPrinter(sCad, lPosFin - lPosIni)
            sRenglon = Left$(sCad, lFin)
        
            If iModo = 0 Then
                lX = 0
            ElseIf iModo = 1 Then
                lX = (lPosFin - Printer.TextWidth(sCadTxt)) / 2
            Else
                lX = lPosFin - Printer.TextWidth(sCadTxt)
            End If
        
            Printer.CurrentX = lPosIni + lX
        
            Printer.Print sRenglon
        
        
            sCad = Mid$(sCad, lFin)
            lIni = lIni + lFin
        Loop
    Else
        If iModo = 0 Then
            lX = 0
        ElseIf iModo = 1 Then
            lX = (lPosFin - Printer.TextWidth(sCadTxt)) / 2
        Else
            lX = lPosFin - Printer.TextWidth(sCadTxt)
        End If
        
        Printer.CurrentX = lPosIni + lX
        
        Printer.Print sCadTxt
    End If

    

End Sub

Public Function BuscaEspacioTextoPrinter(sCadTxt As String, lLongEspacio As Long) As Long

    Dim sChar As String
    Dim lI As Long
    Dim lTamCadena As String
    
    BuscaEspacioTextoPrinter = 0
    
    lTamCadena = Len(sCadTxt)
    
    If Printer.TextWidth(sCadTxt) <= lLongEspacio Then
        BuscaEspacioTextoPrinter = lTamCadena
        Exit Function
    End If
    
    
    For lI = lTamCadena To 1 Step -1
        sChar = Mid$(sCadTxt, lI, 1)
        If sChar = " " Then
            If Printer.TextWidth(Left$(sCadTxt, lI)) <= lLongEspacio Then
                BuscaEspacioTextoPrinter = lI
                Exit For
            End If
        End If
        
        
    Next
    
    

End Function
