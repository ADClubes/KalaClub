Attribute VB_Name = "CreditCard"
Public Function ValidateCardNumber(ByVal card_number As String, ByRef sError As String) As Boolean
    Dim num_digits As Integer
    Dim ch As String
    Dim i As Integer
    Dim odd_digit As Boolean
    Dim doubled_digit As String
    Dim checksum As Integer

    
    ValidateCardNumber = False

    'se quitan los espacios
    card_number = Replace(card_number, " ", "")

    'Verifica que sean 16 digitos
    If Len(card_number) < 15 Then
        sError = "Longitud incorrecta"
        Exit Function
    End If
    ' Examinamos los digitos
    odd_digit = False
    For i = Len(card_number) To 1 Step -1
        ch = Mid$(card_number, i, 1)
        If ch < "0" Or ch > "9" Then
            'No es un número, sale
            sError = "Contiene caracteres no válidos - " & card_number
            Exit Function
        Else
            ' Procesa el digito
            odd_digit = Not odd_digit
            If odd_digit Then
                ' Digito impar, lo agrega al checksum
                checksum = checksum + CInt(ch)
            Else
                ' Digito par, lo duplica y la agrega
                ' al resultado del checksum
                doubled_digit = Format$(2 * CInt(ch))
                checksum = checksum + CInt(Left$(doubled_digit, 1))
                If Len(doubled_digit) = 2 Then checksum = checksum + CInt(Mid$(doubled_digit, 2, 1))
            End If
        End If
    Next i

    ' Check the checksum.
    If (checksum Mod 10) <> 0 Then
        sError = "Checksum incorrecto - " & card_number
        Exit Function
    End If
    ValidateCardNumber = True

End Function

Public Function ValidateCardName(sNombre As String, ByRef sError As String) As Boolean
    Dim lSize As Long
    Dim sNoValidos As String
    Dim lI As Long
    
    
    
    ValidateCardName = False
    
    sNoValidos = ".,:;Ñ-_"
    
    sNombre = Trim(sNombre)
    sNombre = UCase(sNombre)
    
    
    lSize = Len(sNombre)
    
    For lI = 1 To lSize
        If InStr(1, sNoValidos, Mid$(sNombre, lI, 1)) > 0 Then
            sError = "Caracter no válido " & sNombre
            Exit Function
        End If
    Next
    
    ValidateCardName = True
    
End Function
