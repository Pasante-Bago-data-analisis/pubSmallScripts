# pubSmallScripts
VBA Adaptation attempt
```
Option Explicit
  'Main Function
  Function SpellNumber(ByVal MyNumber)
      Dim Dollars, Cents, Temp
      Dim DecimalPlace, Count
      ReDim Place(9) As String
      Place(2) = " mil "
      Place(3) = " millon "
      Place(4) = " billon "
      Place(5) = " trillon "

      MyNumber = Trim(Str(MyNumber))
      DecimalPlace = InStr(MyNumber, ".")
      If DecimalPlace > 0 Then
          Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & _
                    "00", 2))
          MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
      End If
      Count = 1
      Do While MyNumber <> ""
          Temp = GetHundreds(Right(MyNumber, 3))
          If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars
          If Len(MyNumber) > 3 Then
              MyNumber = Left(MyNumber, Len(MyNumber) - 3)
          Else
              MyNumber = ""
          End If
          Count = Count + 1
      Loop
      Select Case Dollars
          Case ""
              Dollars = "cero bolivianos"
          Case "One"
              Dollars = "un boliviano"
           Case Else
              Dollars = Dollars & " bolivianos"
      End Select
      Select Case Cents
          Case ""
              Cents = " y cero centavos"
          Case "One"
              Cents = " y un centavo"
                Case Else
              Cents = " y " & Cents & " centavos"
      End Select
      SpellNumber = Dollars & Cents
  End Function

  Function GetHundreds(ByVal MyNumber)
      Dim Result As String
      If Val(MyNumber) = 0 Then Exit Function
      MyNumber = Right("000" & MyNumber, 3)
      ' Convert the hundreds place.
      If Mid(MyNumber, 1, 1) <> "0" Then
          Result = GetDigit(Mid(MyNumber, 1, 1)) & "cientos "
      End If
      ' Convert the tens and ones place.
      If Mid(MyNumber, 2, 1) <> "0" Then
          Result = Result & GetTens(Mid(MyNumber, 2))
      Else
          Result = Result & GetDigit(Mid(MyNumber, 3))
      End If
      GetHundreds = Result
  End Function

  Function GetTens(TensText)
      Dim Result As String
      Result = "" ' Null out the temporary function value.
      If Val(Left(TensText, 1)) = 1 Then   ' If value between 10-19&hellip;
          Select Case Val(TensText)
              Case 10: Result = "diez"
              Case 11: Result = "once"
              Case 12: Result = "doce"
              Case 13: Result = "trece"
              Case 14: Result = "catorce"
              Case 15: Result = "quince"
              Case 16: Result = "dieciseis"
              Case 17: Result = "diecisiete"
              Case 18: Result = "dieciocho"
              Case 19: Result = "diecinueve"
              Case Else
          End Select
      Else ' If value between 20-99&hellip;
          Select Case Val(Left(TensText, 1))
              Case 2: Result = "veinti"
              Case 3: Result = "treinta y "
              Case 4: Result = "cuarenta y "
              Case 5: Result = "cincuenta y "
              Case 6: Result = "sesenta y "
              Case 7: Result = "setenta y "
              Case 8: Result = "ochenta y "
              Case 9: Result = "noventa y "
              Case Else
          End Select
          Result = Result & GetDigit _
              (Right(TensText, 1))  ' Retrieve ones place.
      End If
      GetTens = Result
  End Function

  Function GetDigit(Digit)
      Select Case Val(Digit)
          Case 1: GetDigit = "un"
          Case 2: GetDigit = "dos"
          Case 3: GetDigit = "tres"
          Case 4: GetDigit = "cuatro"
          Case 5: GetDigit = "cinco"
          Case 6: GetDigit = "seis"
          Case 7: GetDigit = "siete"
          Case 8: GetDigit = "ocho"
          Case 9: GetDigit = "nueve"
          Case Else: GetDigit = ""
      End Select
  End Function

Sub NumToLettersEnglish()

End Sub

```
Functional
```
Option Explicit

Public Function NumeLetras(ByVal numero As Double, conector As String, moneda As String, ByVal Estilo As Integer) As String
  Dim NumTmp As String
  Dim c01 As Integer
  Dim c02 As Integer
  Dim pos As Integer
  Dim dig As Integer
  Dim cen As Integer
  Dim dec As Integer
  Dim uni As Integer
  Dim letra1 As String
  Dim letra2 As String
  Dim letra3 As String
  Dim Leyenda As String
  Dim Leyenda1 As String
  Dim TFNumero As String
        
  If numero < 0 Then numero = Abs(numero)

  NumTmp = Format(numero, "000000000000000.00")        'Le da un formato fijo
  c01 = 1
  pos = 1
  TFNumero = ""
  'Para extraer tres digitos cada vez
  Do While c01 <= 5
    c02 = 1
    Do While c02 <= 3
      'Extrae un digito cada vez de izquierda a derecha
      dig = Val(Mid(NumTmp, pos, 1))
      Select Case c02
        Case 1: cen = dig
        Case 2: dec = dig
        Case 3: uni = dig
      End Select
      c02 = c02 + 1
      pos = pos + 1
    Loop
    letra3 = Centena(uni, dec, cen)
    letra2 = Decena(uni, dec)
    letra1 = Unidad(uni, dec)
            
    Select Case c01
      Case 1
        If cen + dec + uni = 1 Then
          Leyenda = "Billon "
        ElseIf cen + dec + uni > 1 Then
          Leyenda = "Billones "
        End If
      Case 2
        If cen + dec + uni >= 1 And Val(Mid(NumTmp, 7, 3)) = 0 Then
          Leyenda = "Mil Millones "
        ElseIf cen + dec + uni >= 1 Then
          Leyenda = "Mil "
        End If
      Case 3
        If cen + dec = 0 And uni = 1 Then
          Leyenda = "Millon "
        ElseIf cen > 0 Or dec > 0 Or uni > 1 Then
          Leyenda = "Millones "
        End If
      Case 4
        If cen + dec + uni >= 1 Then
          Leyenda = "Mil "
        End If
      Case 5
        If cen + dec + uni >= 1 Then
          Leyenda = ""
        End If
      End Select
            
      c01 = c01 + 1
      TFNumero = TFNumero + letra3 + letra2 + letra1 + Leyenda
      
      Leyenda = ""
      letra1 = ""
      letra2 = ""
      letra3 = ""
  Loop
  TFNumero = TFNumero & conector
  
  Select Case Estilo
    Case 1
      TFNumero = StrConv(TFNumero, vbUpperCase)
      moneda = StrConv(moneda, vbUpperCase)
    Case 2
      TFNumero = StrConv(TFNumero, vbLowerCase)
       moneda = StrConv(moneda, vbLowerCase)
    Case Else
      TFNumero = StrConv(TFNumero, vbProperCase)
            moneda = StrConv(moneda, vbProperCase)
  End Select
  
  TFNumero = TFNumero & " " & Mid(NumTmp, 17) & "/100 "
            
  NumeLetras = TFNumero & moneda
    
End Function

Private Function Centena(ByVal uni As Integer, ByVal dec As Integer, _
                         ByVal cen As Integer) As String
Dim cTexto As String

  Select Case cen
    Case 1
      If dec + uni = 0 Then
        cTexto = "cien "
      Else
        cTexto = "ciento "
      End If
    Case 2: cTexto = "doscientos "
    Case 3: cTexto = "trescientos "
    Case 4: cTexto = "cuatrocientos "
    Case 5: cTexto = "quinientos "
    Case 6: cTexto = "seiscientos "
    Case 7: cTexto = "setecientos "
    Case 8: cTexto = "ochocientos "
    Case 9: cTexto = "novecientos "
    Case Else: cTexto = ""
  End Select
  Centena = cTexto
    
End Function

Private Function Decena(ByVal uni As Integer, ByVal dec As Integer) As String
Dim cTexto As String
  
  Select Case dec
    Case 1:
      Select Case uni
        Case 0: cTexto = "diez "
        Case 1: cTexto = "once "
        Case 2: cTexto = "doce "
        Case 3: cTexto = "trece "
        Case 4: cTexto = "catorce "
        Case 5: cTexto = "quince "
        Case 6 To 9: cTexto = "dieci"
      End Select
    Case 2:
      If uni = 0 Then
        cTexto = "veinte "
      ElseIf uni > 0 Then
        cTexto = "veinti"
      End If
    Case 3: cTexto = "treinta "
    Case 4: cTexto = "cuarenta "
    Case 5: cTexto = "cincuenta "
    Case 6: cTexto = "sesenta "
    Case 7: cTexto = "setenta "
    Case 8: cTexto = "ochenta "
    Case 9: cTexto = "noventa "
    Case Else: cTexto = ""
  End Select
  
  If uni > 0 And dec > 2 Then cTexto = cTexto + "y "
    
  Decena = cTexto
  
End Function

Private Function Unidad(ByVal uni As Integer, ByVal dec As Integer) As String
Dim cTexto As String
  
  If dec <> 1 Then
    Select Case uni
      Case 1: cTexto = "un "
      Case 2: cTexto = "dos "
      Case 3: cTexto = "tres "
      Case 4: cTexto = "cuatro "
      Case 5: cTexto = "cinco "
    End Select
  End If
  Select Case uni
    Case 6: cTexto = "seis "
    Case 7: cTexto = "siete "
    Case 8: cTexto = "ocho "
    Case 9: cTexto = "nueve "
  End Select
  
  Unidad = cTexto

End Function

```
