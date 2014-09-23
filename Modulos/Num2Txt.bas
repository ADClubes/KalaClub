Attribute VB_Name = "NumLetras"
Global Caract(15) As Integer
Global Unidad(9) As String
Global Teens(10) As String
Global Decenas(1 To 9) As String
Global Centenas(1 To 9) As String

Function Num2Txt(Numero As Double) As String
  'Unidades
  Unidad(0) = "Cero"
  Unidad(1) = "Un"
  Unidad(2) = "Dos"
  Unidad(3) = "Tres"
  Unidad(4) = "Cuatro"
  Unidad(5) = "Cinco"
  Unidad(6) = "Seis"
  Unidad(7) = "Siete"
  Unidad(8) = "Ocho"
  Unidad(9) = "Nueve"
  '10 al 15
  Teens(0) = "Diez"
  Teens(1) = "Once"
  Teens(2) = "Doce"
  Teens(3) = "Trece"
  Teens(4) = "Catorce"
  Teens(5) = "Quince"
  Teens(6) = "Dieciseis"
  Teens(7) = "Diecisiete"
  Teens(8) = "Dieciocho"
  Teens(9) = "Diecinueve"

  
  '20 al 90
  Decenas(1) = "Diez"
  Decenas(2) = "Veinte"
  Decenas(3) = "Treinta"
  Decenas(4) = "Cuarenta"
  Decenas(5) = "Cincuenta"
  Decenas(6) = "Sesenta"
  Decenas(7) = "Setenta"
  Decenas(8) = "Ochenta"
  Decenas(9) = "Noventa"
  '100 al 900
  Centenas(1) = "Cien"
  Centenas(2) = "Doscientos"
  Centenas(3) = "Trescientos"
  Centenas(4) = "Cuatrocientos"
  Centenas(5) = "Quinientos"
  Centenas(6) = "Seiscientos"
  Centenas(7) = "Setecientos"
  Centenas(8) = "Ochocientos"
  Centenas(9) = "Novecientos"
  
  Dim NumPPP As String
  Dim NumStr As String
  Dim Largo As Integer
  
  NumStr = ""
  NumPPP = ""
  For i = 1 To (15 - Len(CStr(Numero)))
    NumStr = NumStr + "0"
  Next i
  NumStr = NumStr + CStr(Numero)
  For i = 1 To 15
    If Mid(NumStr, i, 1) <> "." Then
        Caract(i) = CInt(Mid(NumStr, i, 1))
    End If
  Next i
  Largo = Len(CStr(Numero))
  
  Select Case Largo
    Case 1 'Unidad
      NumPPP = Unidad(Numero) + " "
    Case 2 'Decena
      NumPPP = DecenaTxt(14) + Unidadtxt(15)
    Case 3 'Centena
      NumPPP = CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 4 'Mil
      NumPPP = MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 5 'Decena Mil
      NumPPP = DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 6 'Centena Mil
      NumPPP = CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 7 'Millon
      If Caract(9) = 1 Then
        NumPPP = "Un Millón " + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
      Else
        NumPPP = MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
      End If
    Case 8 'Decena Mill
      NumPPP = DecenaTxt(8) + MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 9 'Centena Mill
      NumPPP = CentenaTxt(7) + DecenaTxt(8) + MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 10 'Mil Mill
      NumPPP = MilTxt(6) + CentenaTxt(7) + DecenaTxt(8) + MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 11 'Decena Mill
      NumPPP = DecenaTxt(5) + MilTxt(6) + CentenaTxt(7) + DecenaTxt(8) + MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 12 'Centena Mill
      NumPPP = CentenaTxt(4) + DecenaTxt(5) + MilTxt(6) + CentenaTxt(7) + DecenaTxt(8) + MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 13 'Billon
      If Caract(3) = 1 Then
        NumPPP = "Un Billón " + CentenaTxt(4) + DecenaTxt(5) + MilTxt(6) + CentenaTxt(7) + DecenaTxt(8) + MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
      Else
        NumPPP = MilTxt(3) + "Billones " + CentenaTxt(4) + DecenaTxt(5) + MilTxt(6) + CentenaTxt(7) + DecenaTxt(8) + MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
      End If
    Case 14 'Decena Bill
      NumPPP = DecenaTxt(2) + MilTxt(3) + "Billones " + CentenaTxt(4) + DecenaTxt(5) + MilTxt(6) + CentenaTxt(7) + DecenaTxt(8) + MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
    Case 15 'Centena Bill
      NumPPP = CentenaTxt(1) + DecenaTxt(2) + MilTxt(3) + "Billones " + CentenaTxt(4) + DecenaTxt(5) + MilTxt(6) + CentenaTxt(7) + DecenaTxt(8) + MilTxt(9) + CentenaTxt(10) + DecenaTxt(11) + MilTxt(12) + CentenaTxt(13) + DecenaTxt(14) + Unidadtxt(15)
  End Select
  Num2Txt = NumPPP
End Function

Function Unidadtxt(pos As Integer) As String
  If Caract(pos) > 0 And Caract(pos - 1) = 0 Then
    Unidadtxt = Unidad(Caract(pos)) + " "
  End If
End Function

Function DecenaTxt(pos As Integer) As String
  Select Case Caract(pos)
    Case 1
      Select Case Caract(pos + 1)
        Case Is < 10
          Select Case pos
            Case 14
              DecenaTxt = Teens(Caract(pos + 1)) + " "
            Case 8
              DecenaTxt = Teens(Caract(pos + 1)) + " Millones "
            Case 2
              DecenaTxt = Teens(Caract(pos + 1)) + " "
            Case 5
              DecenaTxt = Teens(Caract(pos + 1)) + " Mil Millones "
            Case Else
              DecenaTxt = Teens(Caract(pos + 1)) + " Mil "
          End Select
        
'        Case Is >= 6
'          Select Case pos
'            Case 14
'              DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " "
'            Case 2
'              DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " "
'            Case 8
'              DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " Millones "
'            Case 5
'              DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " Mil Millones "
'            Case Else
'              DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " Mil "
'          End Select
          
      End Select
    Case 2
        If Caract(pos + 1) > 0 Then
            Select Case pos
                Case 14
                    DecenaTxt = "Veinti" + Unidad(Caract(pos + 1)) + " "
                Case 8
                    DecenaTxt = "Veinti" + Unidad(Caract(pos + 1)) + " Millones "
                Case 5
                    DecenaTxt = "Veinti" + Unidad(Caract(pos + 1)) + " Mil Millones "
                Case 2
                    DecenaTxt = "Veinti" + Unidad(Caract(pos + 1)) + " "
                Case Else
                    DecenaTxt = "Veinti" + Unidad(Caract(pos + 1)) + " Mil "
                End Select
        Else
            Select Case pos
                Case 14
                    DecenaTxt = Decenas(Caract(pos)) + " "
                Case 8
                    DecenaTxt = Decenas(Caract(pos)) + " Millones "
                Case 5
                    DecenaTxt = Decenas(Caract(pos)) + " Mil Millones "
                Case 2
                    DecenaTxt = Decenas(Caract(pos)) + " "
                Case Else
                    DecenaTxt = Decenas(Caract(pos)) + " Mil "
        End Select
        End If
    Case Is > 2
      If Caract(pos + 1) > 0 Then
        Select Case pos
          Case 14
            DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " "
          Case 8
            DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " Millones "
          Case 5
            DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " Mil Millones "
          Case 2
            DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " "
          Case Else
            DecenaTxt = Decenas(Caract(pos)) + " y " + Unidad(Caract(pos + 1)) + " Mil "
        End Select
      Else
        Select Case pos
          Case 14
            DecenaTxt = Decenas(Caract(pos)) + " "
          Case 8
            DecenaTxt = Decenas(Caract(pos)) + " Millones "
          Case 5
            DecenaTxt = Decenas(Caract(pos)) + " Mil Millones "
          Case 2
            DecenaTxt = Decenas(Caract(pos)) + " "
          Case Else
            DecenaTxt = Decenas(Caract(pos)) + " Mil "
        End Select
      End If
  End Select
End Function

Function CentenaTxt(pos As Integer) As String
  Select Case Caract(pos)
    Case 1
      If Caract(pos + 1) = 0 And Caract(pos + 2) = 0 Then
        Select Case pos
          Case 4
            CentenaTxt = "Cien Mil Millones "
          Case 7
            CentenaTxt = "Cien Millones "
          Case 10
            CentenaTxt = "Cien Mil "
          Case Else
            CentenaTxt = "Cien "
        End Select
      Else
        CentenaTxt = "Ciento "
      End If
    Case Is > 1
      CentenaTxt = Centenas(Caract(pos)) + " "
  End Select
End Function

Function MilTxt(pos As Integer) As String
  If Caract(pos - 1) = 0 Then
    Select Case Caract(pos)
      Case 1
        Select Case pos
          Case 6
            MilTxt = "Mil Millones "
          Case 12
            MilTxt = "Un Mil "
        End Select
      Case Is > 1
        Select Case pos
          Case 6
          Case 12
            MilTxt = Unidad(Caract(pos)) + " Mil "
          Case 9
            MilTxt = Unidad(Caract(pos)) + " Millones "
          Case Else
            MilTxt = Unidad(Caract(pos)) + " "
        End Select
    Case Else
        MilTxt = " Mil "
    End Select
  End If
End Function
