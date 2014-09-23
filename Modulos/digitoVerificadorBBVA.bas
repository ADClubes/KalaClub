Attribute VB_Name = "digitoVerificadorBBVA"
Option Explicit
Const sCadenaValida As String = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Function dvAlgoritmo35(sCadena As String) As Integer
    Dim lLongMax As Integer
    Dim aValores() As Integer
    Dim lLongCadena As Integer
    
    Dim iSumaValores As Integer
    
    lLongMax = 19
    
    
    
    dvAlgoritmo35 = -1
    
    lLongCadena = Len(sCadena)
    
    If (lLongCadena > lLongMax) Then
        Return
    End If
    
    dvAlgoritmo35 = -2
    
    
    If (Not CaracteresValidos(sCadena)) Then
        Return
    End If
    
    
    aValores = LlenaArray(sCadena)
    aValores = multiplicaArray(aValores)
    
    iSumaValores = SumaArray(aValores)
    
    dvAlgoritmo35 = ValorSuperior(iSumaValores) - iSumaValores
    
End Function

Private Function CaracteresValidos(sCadena As String) As Boolean
    
    
    Dim lI As Integer
    
    CaracteresValidos = False
    
    
    For lI = 1 To Len(sCadena)
        If InStr(sCadenaValida, Mid$(sCadena, lI, 1)) = 0 Then
            Exit Function
        End If
    Next lI
    
    CaracteresValidos = True
    
End Function

Private Function LlenaArray(sCadena As String) As Integer()
    Dim aArray() As Integer
    Dim iValor As Integer
    Dim lI As Integer
    
    ReDim aArray(Len(sCadena) - 1)
    
    
    For lI = 0 To Len(sCadena) - 1
        iValor = InStr(sCadenaValida, Mid$(sCadena, lI + 1, 1))
        If iValor > 9 Then
            iValor = iValor - 10
        End If
        aArray(lI) = iValor
    Next lI
    
    LlenaArray = aArray
    
End Function

Private Function multiplicaArray(aArray() As Integer) As Integer()
    Dim lI As Integer
    Dim aFactores(3) As Integer
    Dim aResultado() As Integer
    
    ReDim aResultado(UBound(aArray))
    
    aFactores(0) = 4
    aFactores(1) = 3
    aFactores(2) = 8
    
    For lI = 0 To UBound(aArray)
        aResultado(lI) = SumaDigito(aArray(lI) * aFactores(IIf(lI Mod UBound(aFactores) = 0, 0, lI Mod UBound(aFactores))))
    Next lI
    
    multiplicaArray = aResultado

End Function

Private Function SumaDigito(iDigito As Integer) As Integer
    Dim iResto As Integer
    
    SumaDigito = 0
    
    iResto = iDigito
    
    If iDigito > 9 Then
        Do While iResto > 0
            SumaDigito = SumaDigito + (iResto - Int(iResto / 10) * 10)
            iResto = Int(iResto / 10)
        Loop
        If SumaDigito > 9 Then
            SumaDigito = SumaDigito(SumaDigito)
        End If
    Else
        SumaDigito = iDigito
    End If
    
End Function

Private Function SumaArray(aArray() As Integer) As Integer
    
    Dim iResultado As Integer
    
    Dim lI As Integer
    
    For lI = 0 To UBound(aArray)
        iResultado = iResultado + aArray(lI)
    Next lI

    SumaArray = iResultado

End Function

Private Function ValorSuperior(iValor As Integer) As Integer
    Dim iUnidad As Integer
    
    iUnidad = 10 - iValor Mod 10
    
    If iUnidad = 10 Then iUnidad = 0
    
    ValorSuperior = iValor + iUnidad
    
End Function

