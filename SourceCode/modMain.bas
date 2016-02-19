Attribute VB_Name = "modMain"
Function BaseConvert(NumIn As String, BaseIn As Byte, BaseOut As Byte) As String
   'Binär       = Basis 2
   'Oktal       = Basis 8
   'Dezimal     = Basis 10
   'Hexadezimal = Basis 16

   Dim i As Integer, CurrentCharacter As String, CharacterValue As Integer
   Dim PlaceValue As Integer, RunningTotal As Double, Remainder As Double
   Dim BaseOutDouble As Double, NumInCaps As String

   If NumIn = "" Or BaseIn < 2 Or BaseIn > 36 Or BaseOut < 1 Or BaseOut > 36 Then
      'Keine Angabe oder ungültiges Zahlensystem
      BaseConvert = "Error"
      Exit Function
   End If

   NumInCaps = UCase(NumIn)

   PlaceValue = Len(NumInCaps)

   For i = 1 To Len(NumInCaps)
      PlaceValue = PlaceValue - 1
      CurrentCharacter = Mid$(NumInCaps, i, 1)
      CharacterValue = 0
      If Asc(CurrentCharacter) > 64 And Asc(CurrentCharacter) < 91 Then
         CharacterValue = Asc(CurrentCharacter) - 55
      End If

      If CharacterValue = 0 Then
         If Asc(CurrentCharacter) < 48 Or Asc(CurrentCharacter) > 57 Then
            BaseConvert = "Error"
            Exit Function
         Else
            CharacterValue = Val(CurrentCharacter)
         End If
      End If

      If CharacterValue < 0 Or CharacterValue > BaseIn - 1 Then
         BaseConvert = "Error"
         Exit Function
      End If
      RunningTotal = RunningTotal + CharacterValue * (BaseIn ^ PlaceValue)
   Next i

   Do
      BaseOutDouble = CDbl(BaseOut)
      Remainder = RunningTotal - (Int(RunningTotal / BaseOutDouble) * BaseOutDouble)
      RunningTotal = (RunningTotal - Remainder) / BaseOut

      If Remainder >= 10 Then
         CurrentCharacter = Chr$(Remainder + 55)
      Else
         CurrentCharacter = Right$(Str$(Remainder), Len(Str$(Remainder)) - 1)
      End If
      BaseConvert = CurrentCharacter & BaseConvert
   Loop While RunningTotal > 0
End Function
