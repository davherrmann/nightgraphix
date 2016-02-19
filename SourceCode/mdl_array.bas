Attribute VB_Name = "mdl_array"
'_________________________________________________________________________________'
'|                               MODUL mdl_array                                 |'
'| Dieses Modul beinhaltet Routinen f�r das Array, wie zum Beispiel das          |'
'| Errechnen des Feldes .                                                        |'
'---------------------------------------------------------------------------------'
Option Explicit

'VARIABLEN..
'Arrays f�r die Speicherung der Daten
'Werden sp�ter mit ReDim zweidimensional gemacht
'RGB-Arrays
Public Array_Red() As Byte
Public Array_Green() As Byte
Public Array_Blue() As Byte
'SW-Array
Public Array_SW() As Byte
'Array f�r Felder unter dem Cursor
Public Array_CursorRGB() As Long
'Array f�r Felder des Cursors (Schwarz oder Wei�)
Public Array_CursorSW() As Long
'Array f�r die Positionen des Cursors
Public Array_CursorPos() As POINTAPI
'Array f�r das Brush-Objekt
Public Array_RGBBrush(0 To 1, 0 To 1, 0 To 1) As Long
'Variablen f�r Schleifen
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
   Array_Red(BreiteFeld + 1, H�heFeld + 1) = ARRAYRGB.R / 255
   Array_Green(BreiteFeld + 1, H�heFeld + 1) = ARRAYRGB.G / 255
   Array_Blue(BreiteFeld + 1, H�heFeld + 1) = ARRAYRGB.B / 255
   'SW-Wert ins SW-Array eintragen
   Array_SW(BreiteFeld + 1, H�heFeld + 1) = Draw_Color2SW(tmpCol)

   '�nderungen sind nicht gespeichert
   FileSystem_SavedChanges False
End Sub

Public Sub Array_ClearCursorArray()
'--------------------------------------------------------------------'
'| Prozedur zum L�schen des Cursor-Arrays                           |'
'--------------------------------------------------------------------'
   'Schleife durch alle Pixel des Cursors
   For X = 1 To Letter.FontHeight
      'Aktuelles Pixel mit wei� f�llen (l�schen)
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
   'Variable f�r Speicherung eines Bytes
   Dim RGBByte As Integer
   'Schleife durch 8 LEDs in einer Spalte des Red-Arrays
   For Y = tmpLED To tmpLED + 7
      'LEDs in Array_GetByte speichern
      RGBByte = RGBByte + Array_Red(tmpSpalte, Y) * 2 ^ (Y - tmpLED)
   Next Y
   'Byte in Buchstaben umwandeln
   Array_GetRGBByte = Array_GetRGBByte & Chr(RGBByte)
   'Variable l�schen
   RGBByte = 0
   
   'Schleife durch 8 LEDs in einer Spalte des Green-Arrays
   For Y = tmpLED To tmpLED + 7
      'LEDs in Array_GetByte speichern
      RGBByte = RGBByte + Array_Green(tmpSpalte, Y) * 2 ^ (Y - tmpLED)
   Next Y
   'Byte in Buchstaben umwandeln
   Array_GetRGBByte = Array_GetRGBByte & Chr(RGBByte)
   'Variable l�schen
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
   'String zum Speichern des Bin�rwertes
   Dim Bin�rString As String
   'Dezimalwert in Bin�rwert umwandeln
   Bin�rString = Maths_Dez2Bin(tmpByte)
   'Bin�rstring vorne mit 0-en f�llen
   Bin�rString = Format(Bin�rString, "00000000")
   'Daten umdrehen
   Bin�rString = StrReverse(Bin�rString)
   
   'Schleife durch 8 LEDs in einer Spalte
   For Y = tmpLED To tmpLED + 7
      
      'LEDs in Array_GetByte speichern
      Array_SW(tmpSpalte, Y + 1) = CInt(Mid(Bin�rString, Y - tmpLED + 1, 1))
      
      'RGB-Arrays f�llen
      Array_Red(tmpSpalte, Y + 1) = IIf(Array_SW(tmpSpalte, tmpLED + 7 - Y + 1) = 1, Draw_Color2RGB_Bool(LEDColor).R, 0)
      Array_Green(tmpSpalte, Y + 1) = IIf(Array_SW(tmpSpalte, tmpLED + 7 - Y + 1) = 1, Draw_Color2RGB_Bool(LEDColor).G, 0)
      Array_Blue(tmpSpalte, Y + 1) = IIf(Array_SW(tmpSpalte, tmpLED + 7 - Y + 1) = 1, Draw_Color2RGB_Bool(LEDColor).B, 0)
      'Wenn eine "1" im Array steht
      If Array_SW(tmpSpalte, Y + 1) = 1 Then
         'Zelle im Modell f�llen
         Draw_FillCell tmpSpalte, Y + 1, tmpPicSource, LEDColor, False
      'Wenn eine "0" im Array steht
      Else
         'Zelle im Modell l�schen
         Draw_FillCell tmpSpalte, Y + 1, tmpPicSource, vbWhite, False
      End If
   Next Y
End Sub

Public Sub Array_SetRGBByte(ByVal tmpSpalte As Integer, ByVal tmpLED As Integer, ByVal tmpByteRed As Byte, ByVal tmpByteGreen As Byte, ByVal tmpByteBlue As Byte, ByRef tmpPicSource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen eines Bytes in dem SW-Array                  |'
'--------------------------------------------------------------------'
   'String zum Speichern des Bin�rwertes
   Dim Bin�rString As String
   'Dezimalwert in Bin�rwert umwandeln, mit 0en f�llen und String umdrehen: Rot-Wert
   Bin�rString = StrReverse(Format(Maths_Dez2Bin(tmpByteRed), "00000000"))
   'Dezimalwert in Bin�rwert umwandeln, mit 0en f�llen und String umdrehen: Gr�n-Wert
   Bin�rString = Bin�rString & StrReverse(Format(Maths_Dez2Bin(tmpByteGreen), "00000000"))
   'Dezimalwert in Bin�rwert umwandeln, mit 0en f�llen und String umdrehen: Blau-Wert
   Bin�rString = Bin�rString & StrReverse(Format(Maths_Dez2Bin(tmpByteBlue), "00000000"))
   Bin�rString = StrReverse(Bin�rString)
   
   'Schleife durch 8 LEDs in einer Spalte
   For Y = tmpLED To tmpLED + 7
'
'      'LEDs in Array_GetByte speichern
'      Array_SW(tmpSpalte, Y + 1) = CInt(Mid(Bin�rString, Y - tmpLED + 1, 1))
'
      'RGB-Arrays f�llen
      Array_Red(tmpSpalte, Y + 1) = CInt(Mid(Bin�rString, Y - tmpLED + 1, 1))
      Array_Green(tmpSpalte, Y + 1) = CInt(Mid(Bin�rString, Y - tmpLED + 9, 1))
      Array_Blue(tmpSpalte, Y + 1) = CInt(Mid(Bin�rString, Y - tmpLED + 17, 1))
      
      'Zelle im Modell f�llen
      Draw_FillCell tmpSpalte, Y + 1, tmpPicSource, RGB(Array_Red(tmpSpalte, Y + 1) * 255, Array_Green(tmpSpalte, Y + 1) * 255, Array_Blue(tmpSpalte, Y + 1) * 255), False
   Next Y
End Sub

Public Function Array_GetRGBCell(ByVal tmpSpalte As Integer, ByVal tmpLED As Integer, Optional ByVal Invert As Boolean = False) As Long
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen einer Zelle aus den RGB-Arrays             |'
'--------------------------------------------------------------------'
   'Integer-Variable f�rs Invertieren
   Dim intInvert As Integer
   
   'Wenn invertiert werden soll, dann Variable auf 255 setzen
   If Invert Then intInvert = 255
   
   'Wert auslesen, bei Bedarf gleich invertieren
   Array_GetRGBCell = RGB(Abs(intInvert - Array_Red(tmpSpalte, tmpLED) * 255), Abs(intInvert - Array_Green(tmpSpalte, tmpLED) * 255), Abs(intInvert - Array_Blue(tmpSpalte, tmpLED) * 255))

   'Wenn Farbe ganz schwarz, auf dunkelgrau setzen
   If Array_GetRGBCell = 0 Then Array_GetRGBCell = RGB(60, 60, 60)
End Function


