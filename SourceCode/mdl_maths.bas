Attribute VB_Name = "mdl_maths"
'_________________________________________________________________________________'
'|                               MODUL mdl_maths                                 |'
'| Dieses Modul beinhaltet Berechnungs-Routinen, wie z.B. das Berechnen          |'
'| eines Winkels.                                                                |'
'---------------------------------------------------------------------------------'
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN
'Zählvariable
Private i As Long
'Variablen für 2-dimensionale Schleifen
Private X As Integer
Private Y As Integer

'KONSTANTEN..
'PI (statt Zahl könnte man auch mit 4*Atn(1) nehmen)
Public Const PI = 3.14159265358979

Private Function Maths_ATan(ByVal X As Double, ByVal Y As Double) As Double
'--------------------------------------------------------------------'
'| Prozedur zum Berechnen des Arkustangens                          |'
'--------------------------------------------------------------------'
   'Den Arkustangens berechnen
   'Wenn X = 0, X sehr klein machen (ca. 0), da nicht durch
   '0 geteilt werden darf
   If X = 0 Then
      X = 1E-300
   End If
   
   'Arkustangens mit VB-Funktion ausrechnen
   'VB-Atn kann keine negativen Vorzeichen berechnen
   Maths_ATan = Atn(Y / X)
   
   'Wenn X = negativ
   If X < 0 Then
      'PI zu Ergebnis addieren
      Maths_ATan = Maths_ATan + PI
   'Wenn Y = negativ
   ElseIf Y < 0 Then
      '2 * PI zu Ergebnis addieren
      Maths_ATan = Maths_ATan + 2 * PI
   End If
End Function

Public Function Maths_GetWinkel(ByVal X As Double, ByVal Y As Double) As Double
'--------------------------------------------------------------------'
'| Prozedur zum Berechnen des Winkels "zwischen zwei Punkten"       |'
'--------------------------------------------------------------------'
   'Den Winkel berechnen
   Maths_GetWinkel = (270 - Maths_ATan((Y - AbstandY - KreisRadius), (X - AbstandX - KreisRadius)) * 360 / (2 * PI) + 270)
   
   'Wenn Winkel kleiner als 0 ist, positiv machen
   If Maths_GetWinkel < 0 Then Maths_GetWinkel = 360 + Maths_GetWinkel
   'Wenn Winkel größer als 360 ist, 360 abziehen
   If Maths_GetWinkel >= 360 Then Maths_GetWinkel = Maths_GetWinkel - 360
End Function

Public Sub Maths_Long2Cell(ByVal X As Double, ByVal Y As Double)
'--------------------------------------------------------------------'
'| Prozedur zum Umrechnen von Koordinaten in eine Zellenangabe      |'
'--------------------------------------------------------------------'
   'Die Zellangaben werden in BreiteFeld und HöheFeld geschrieben
   'Bei einem Fehler einfach weitermachen
   On Error Resume Next
   
   'Kreisviertel herausfinden
   'Linke oder rechte Seite ?
   'Wenn X auf der rechten Seite liegt..
   If X - AbstandX > KreisRadius Then
      'Wenn Y in der unteren Hälfte liegt..
      If Y - AbstandY > KreisRadius Then
         'Kreisviertel auf 2 setzten
         KreisViertel = 2
      'Wenn Y in der oberen Hälfte liegt
      Else
         'Kreisviertel auf 1 setzten
         KreisViertel = 1
      End If
   'Wenn X auf der linken Seite liegt..
   Else
      'Wenn Y in der unteren Hälfte liegt..
      If Y - AbstandY > KreisRadius Then
         'Kreisviertel auf 3 setzten
         KreisViertel = 3
      Else
         'Kreisviertel auf 0 setzten
         KreisViertel = 0
      End If
   End If
   
   'Winkel ausrechnen
   Winkel = Maths_GetWinkel(X, Y)
   
   'Spalte herausfinden
   For i = 1 To Spalten
      'Wenn Winkel erreicht worden ist..
      If Winkel / 360 * Spalten <= i Then
         'BreiteFeld setzen (welche Spalte wurde gewählt ?)
         BreiteFeld = i - 1
         Exit For
      Else
         'Sonst BreiteFeld auf i setzen
         BreiteFeld = 0
      End If
   Next i
      
   'Led herausfinden
   For i = 0 To Leds - 1
      'Wenn LED erreicht worden ist..
      If (Sqr((Y - AbstandY - KreisRadius) ^ 2 + (X - AbstandX - KreisRadius) ^ 2) - InnenKreisRadius) / (KreisRadius - InnenKreisRadius) * Leds < i Then
         'HöheFeld setzen (welche LED wurde gewählt ?)
         HöheFeld = i - 1
         Exit For
      Else
         'Sonst HöheFeld auf i setzen
         HöheFeld = i
      End If
   Next i
End Sub

Public Function Maths_GetCellPosition(ByVal X As Long, ByVal Y As Long, ByRef picsource As PictureBox) As POINTAPI
'--------------------------------------------------------------------'
'| Prozedur zum Umrechnen von Zellenangaben in Koordinaten          |'
'--------------------------------------------------------------------'
   'Hub addieren
   Y = Y + Hub
   
   'Winkel berechnen
   Winkel = 2 * PI - (((X / Spalten) * 360 - 0.5) / 180 * PI)
   
   'X und Y umrechnen in geometrische Daten
   Maths_GetCellPosition.X = picsource.Height / 2 - SpaltenAbstand * (Y - 0.5) * Sin(Winkel)
   Maths_GetCellPosition.Y = picsource.Width / 2 - SpaltenAbstand * (Y - 0.5) * Cos(Winkel)
End Function

Public Function Maths_Dez2Bin(ByVal tmpDezimal As Long) As String
'--------------------------------------------------------------------'
'| Prozedur zum Umrechnen von Dezimalzahlen in Binärstrings         |'
'--------------------------------------------------------------------'
   'Zählervariable i auf 0 setzen
   i = 0
   'Schleife mit Abbruchbedingung
   Do
      'Eine "0" oder "1" zum Binärstring hinzufügen
      Maths_Dez2Bin = CStr(IIf(tmpDezimal And 2 ^ i, "1", "0")) & Maths_Dez2Bin
      'Zählervariable um 1 erhöhen
      i = i + 1
   'Abbruchbedingung
   Loop Until 2 ^ i > tmpDezimal
End Function

Public Sub Maths_EEPROMData()
'--------------------------------------------------------------------'
'| Prozedur zum Errechnen der EEPROM-Daten                          |'
'--------------------------------------------------------------------'
   'Adresse im EEPROM ausrechnen
   'Das geht hardwarebedingt nicht, da die Hardware immer 8 Byte verlangt:
   'EEPROM_Address = (NumberOfPicture - 1) * Spalten * (Leds \ 8) + (NGRequestSpalte - 1) * (Leds \ 8)
   'Deshalb immer "mal 8 Byte"
   EEPROM_Address = (NumberOfPicture - 1 + Abs(CInt(BottomSide)) * 3) * Spalten * 8 + (NGRequestSpalte - 1) * 8
   
   'High-Adresse ausrechnen
   EEPROM_Address_High = EEPROM_Address \ 256
   'Low-Adresse ausrechnen
   EEPROM_Address_Low = EEPROM_Address Mod 256
End Sub

Public Function Maths_CountActiveLEDs() As Long
'--------------------------------------------------------------------'
'| Prozedur zum Zählen der aktiven LEDs                             |'
'--------------------------------------------------------------------'
   'Schleife durch alle Spalten
   For X = 1 To Spalten
      'Schleife durch alle LEDs
      For Y = 1 To Leds
         'Wenn LED aktiv ist..
         If Array_SW(X, Y) Then
            'Aktive LEDs um eins erhöhen
            Maths_CountActiveLEDs = Maths_CountActiveLEDs + 1
         End If
      Next Y
   Next X
End Function
