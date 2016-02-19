Attribute VB_Name = "mdl_array"
'_________________________________________________________________________________'
'|                               MODUL mdl_array                                 |'
'| Dieses Modul beinhaltet Routinen für das Array, wie zum Beispiel das          |'
'| Errechnen des Feldes .                                                        |'
'---------------------------------------------------------------------------------'
Option Explicit

'VARIABLEN..
'Arrays für die Speicherung der Daten
'Werden später mit ReDim zweidimensional gemacht
'RGB-Arrays
Public Array_Red() As Byte
Public Array_Green() As Byte
Public Array_Blue() As Byte
'SW-Array
Public Array_SW() As Byte
'Array für Felder unter dem Cursor
Public Array_CursorRGB() As Long
'Array für Felder des Cursors (Schwarz oder Weiß)
Public Array_CursorSW() As Long
'Array für die Positionen des Cursors
Public Array_CursorPos() As POINTAPI
'Array für das Brush-Objekt
Public Array_RGBBrush(0 To 1, 0 To 1, 0 To 1) As Long
'Variablen für Schleifen
Private X As Long
Private Y As Long

Public Sub Array_Anpassen(ByVal X As Single, ByVal Y As Single, ByVal tmpCol As Long)
'--------------------------------------------------------------------'
'| Prozedur zum Zeichnen des Modells                                |'
'--------------------------------------------------------------------'
   'Bei einem Fehler weitermachen
   'On Error Resume Next
   
   'Koordinaten in Zellangaben umrechnen
   Maths_Long2Cell X, Y
   
   'Long-Farbwert in RGB umrechnen
   ARRAYRGB = Draw_Color2RGB_Bool(tmpCol)
   
   'Werte in die Farb-Arrays eintragen
   Array_Red(BreiteFeld + 1, HöheFeld + 1) = ARRAYRGB.R / 255
   Array_Green(BreiteFeld + 1, HöheFeld + 1) = ARRAYRGB.G / 255
   Array_Blue(BreiteFeld + 1, HöheFeld + 1) = ARRAYRGB.B / 255
   'SW-Wert ins SW-Array eintragen
   Array_SW(BreiteFeld + 1, HöheFeld + 1) = Draw_Color2SW(tmpCol)

   'Änderungen sind nicht gespeichert
   FileSystem_SavedChanges False
End Sub

Public Sub Array_ClearCursorArray()
'--------------------------------------------------------------------'
'| Prozedur zum Löschen des Cursor-Arrays                           |'
'--------------------------------------------------------------------'
   'Schleife durch alle Pixel des Cursors
   For X = 1 To Letter.FontHeight
      'Aktuelles Pixel mit weiß füllen (löschen)
      Array_CursorRGB(X - 1) = vbWhite
   Next X
End Sub

Public Function Array_GetByte(ByVal tmpSpalte As Integer, ByVal tmpLED As Integer) As Integer
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen eines Bytes aus dem SW-Array               |'
'--------------------------------------------------------------------'
   'Schleife durch 8 LEDs in einer Spalte
   For Y = tmpLED To tmpLED + 7
      'LEDs in Array_GetByte speichern
      Array_GetByte = Array_GetByte + Array_SW(tmpSpalte, Y) * 2 ^ (Y - tmpLED)
   Next Y
End Function

Public Function Array_GetRGBByte(ByVal tmpSpalte As Integer, ByVal tmpLED As Integer) As String
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen je eines Bytes aus den RGB-Arrays          |'
'--------------------------------------------------------------------'
   'Variable für Speicherung eines Bytes
   Dim RGBByte As Integer
   'Schleife durch 8 LEDs in einer Spalte des Red-Arrays
   For Y = tmpLED To tmpLED + 7
      'LEDs in Array_GetByte speichern
      RGBByte = RGBByte + Array_Red(tmpSpalte, Y) * 2 ^ (Y - tmpLED)
   Next Y
   'Byte in Buchstaben umwandeln
   Array_GetRGBByte = Array_GetRGBByte & Chr(RGBByte)
   'Variable löschen
   RGBByte = 0
   
   'Schleife durch 8 LEDs in einer Spalte des Green-Arrays
   For Y = tmpLED To tmpLED + 7
      'LEDs in Array_GetByte speichern
      RGBByte = RGBByte + Array_Green(tmpSpalte, Y) * 2 ^ (Y - tmpLED)
   Next Y
   'Byte in Buchstaben umwandeln
   Array_GetRGBByte = Array_GetRGBByte & Chr(RGBByte)
   'Variable löschen
   RGBByte = 0
   
   'Schleife durch 8 LEDs in einer Spalte des Blue-Arrays
   For Y = tmpLED To tmpLED + 7
      'LEDs in Array_GetByte speichern
      RGBByte = RGBByte + Array_Blue(tmpSpalte, Y) * 2 ^ (Y - tmpLED)
   Next Y
   'Byte in Buchstaben umwandeln
   Array_GetRGBByte = Array_GetRGBByte & Chr(RGBByte)
End Function

Public Sub Array_SetByte(ByVal tmpSpalte As Integer, ByVal tmpLED As Integer, ByVal tmpByte As Byte, ByRef tmpPicSource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen eines Bytes in dem SW-Array                  |'
'--------------------------------------------------------------------'
   'String zum Speichern des Binärwertes
   Dim BinärString As String
   'Dezimalwert in Binärwert umwandeln
   BinärString = Maths_Dez2Bin(tmpByte)
   'Binärstring vorne mit 0-en füllen
   BinärString = Format(BinärString, "00000000")
   'Daten umdrehen
   BinärString = StrReverse(BinärString)
   
   'Schleife durch 8 LEDs in einer Spalte
   For Y = tmpLED To tmpLED + 7
      
      'LEDs in Array_GetByte speichern
      Array_SW(tmpSpalte, Y + 1) = CInt(Mid(BinärString, Y - tmpLED + 1, 1))
      
      'RGB-Arrays füllen
      Array_Red(tmpSpalte, Y + 1) = IIf(Array_SW(tmpSpalte, tmpLED + 7 - Y + 1) = 1, Draw_Color2RGB_Bool(LEDColor).R, 0)
      Array_Green(tmpSpalte, Y + 1) = IIf(Array_SW(tmpSpalte, tmpLED + 7 - Y + 1) = 1, Draw_Color2RGB_Bool(LEDColor).G, 0)
      Array_Blue(tmpSpalte, Y + 1) = IIf(Array_SW(tmpSpalte, tmpLED + 7 - Y + 1) = 1, Draw_Color2RGB_Bool(LEDColor).B, 0)
      'Wenn eine "1" im Array steht
      If Array_SW(tmpSpalte, Y + 1) = 1 Then
         'Zelle im Modell füllen
         Draw_FillCell tmpSpalte, Y + 1, tmpPicSource, LEDColor, False
      'Wenn eine "0" im Array steht
      Else
         'Zelle im Modell löschen
         Draw_FillCell tmpSpalte, Y + 1, tmpPicSource, vbWhite, False
      End If
   Next Y
End Sub

Public Sub Array_SetRGBByte(ByVal tmpSpalte As Integer, ByVal tmpLED As Integer, ByVal tmpByteRed As Byte, ByVal tmpByteGreen As Byte, ByVal tmpByteBlue As Byte, ByRef tmpPicSource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen eines Bytes in dem SW-Array                  |'
'--------------------------------------------------------------------'
   'String zum Speichern des Binärwertes
   Dim BinärString As String
   'Dezimalwert in Binärwert umwandeln, mit 0en füllen und String umdrehen: Rot-Wert
   BinärString = StrReverse(Format(Maths_Dez2Bin(tmpByteRed), "00000000"))
   'Dezimalwert in Binärwert umwandeln, mit 0en füllen und String umdrehen: Grün-Wert
   BinärString = BinärString & StrReverse(Format(Maths_Dez2Bin(tmpByteGreen), "00000000"))
   'Dezimalwert in Binärwert umwandeln, mit 0en füllen und String umdrehen: Blau-Wert
   BinärString = BinärString & StrReverse(Format(Maths_Dez2Bin(tmpByteBlue), "00000000"))
   BinärString = StrReverse(BinärString)
   
   'Schleife durch 8 LEDs in einer Spalte
   For Y = tmpLED To tmpLED + 7
'
'      'LEDs in Array_GetByte speichern
'      Array_SW(tmpSpalte, Y + 1) = CInt(Mid(BinärString, Y - tmpLED + 1, 1))
'
      'RGB-Arrays füllen
      Array_Red(tmpSpalte, Y + 1) = CInt(Mid(BinärString, Y - tmpLED + 1, 1))
      Array_Green(tmpSpalte, Y + 1) = CInt(Mid(BinärString, Y - tmpLED + 9, 1))
      Array_Blue(tmpSpalte, Y + 1) = CInt(Mid(BinärString, Y - tmpLED + 17, 1))
      
      'Zelle im Modell füllen
      Draw_FillCell tmpSpalte, Y + 1, tmpPicSource, RGB(Array_Red(tmpSpalte, Y + 1) * 255, Array_Green(tmpSpalte, Y + 1) * 255, Array_Blue(tmpSpalte, Y + 1) * 255), False
   Next Y
End Sub

Public Function Array_GetRGBCell(ByVal tmpSpalte As Integer, ByVal tmpLED As Integer, Optional ByVal Invert As Boolean = False) As Long
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen einer Zelle aus den RGB-Arrays             |'
'--------------------------------------------------------------------'
   'Integer-Variable fürs Invertieren
   Dim intInvert As Integer
   
   'Wenn invertiert werden soll, dann Variable auf 255 setzen
   If Invert Then intInvert = 255
   
   'Wert auslesen, bei Bedarf gleich invertieren
   Array_GetRGBCell = RGB(Abs(intInvert - Array_Red(tmpSpalte, tmpLED) * 255), Abs(intInvert - Array_Green(tmpSpalte, tmpLED) * 255), Abs(intInvert - Array_Blue(tmpSpalte, tmpLED) * 255))

   'Wenn Farbe ganz schwarz, auf dunkelgrau setzen
   If Array_GetRGBCell = 0 Then Array_GetRGBCell = RGB(60, 60, 60)
End Function


