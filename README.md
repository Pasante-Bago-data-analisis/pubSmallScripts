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
