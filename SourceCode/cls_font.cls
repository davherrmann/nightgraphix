VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_font"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'EVENTs..
Public Event Error(ByVal ErrorDescription As String)

'APIs..
'API für das Lesen eines Pixels
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'API zum Setzen eines Pixels
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'TYPEN..
'Type für das Speichern von RGB-Werten
Private Type tRGB
   'Variable für den Rot-Wert
   R As Integer
   'Variable für den Grün-Wert
   G As Integer
   'Variable für den Blau-Wert
   B As Integer
End Type

'VARIABLEN..
'Variablen für Schleifen etc.
Private X As Integer
Private Y As Integer

'Variablen für das Suchen des Buchstabens in der PictureBox:
'Rechter Rand
Private Spalte_1 As Integer
'Linker Rand
Private Spalte_2
'Obere Kante
Private Zeile_1 As Integer
'Untere Kante
Private Zeile_2 As Integer

'RGB-Type für SW-Umwandlung
Private SWRGB As tRGB
'PictureBox für das Erzeugen eines Buchstabens
Private picLetter As PictureBox
'Variablen für Font:
'Schriftart
Private mFontName As String
'Schriftgröße
Private mFontSize As Integer
'Fett ?
Private mFontBold As Boolean
'Kursiv ?
Private mFontItalic As Boolean
'Unterstrichen ?
Private mFontUnderline As Boolean
'Durchgestrichen
Private mFontStrikethru As Boolean
'Höhe des Buchstabens
Private mHeight As Integer
'Breite des Buchstabens
Private mWidth As Integer
'Position des Buchstabens (Pixel von Unterkante "T" zu Unterkante Buchstabe)
Private mPosition As Integer
'Höhe des Fonts ("T")
Private mFontHeight As Integer
'Position des Fonts
Private mFontPosition As Integer
'Pixel-Array
Private arrPixel() As Boolean
'Buchstabe
Private mLetter As String

'KONSTANTEN..
'Schwellwert für das Umwandeln von Farbe in SW
Private Const Schwellwert = 145
'Masken für die Umrechnung von Long-Farbwerten nach RGB
Private Const Mask_R = &HFF&
Private Const Mask_G = &HFF00&
Private Const Mask_B = &HFF0000

Public Sub Initialize(ByRef tmpPictureBox As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur beim Initialisieren der Klasse                          |'
'--------------------------------------------------------------------'
   'Referenz auf eine PictureBox erstellen
   Set picLetter = tmpPictureBox
   'AutoRedraw der PictureBox auf True setzen
   picLetter.AutoRedraw = True
   'ScaleMode der PictureBox auf "3 - Pixel" setzen
   picLetter.ScaleMode = 3
   
   'Font-Grundeinstellungen ändern
   'FontName setzen
   FontName = "Arial"
   'FontSize setzen
   FontSize = 12
   'Font-Werte aktualisieren
   Refresh_FontValues
End Sub

Public Function Create_Letter(ByVal tmpLetter As String) As Boolean()
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben eines tLetter-Types aus einem Buchstaben|'
'--------------------------------------------------------------------'
   'Buchstabe in Variable schreiben
   mLetter = tmpLetter
   'Buchstaben in die PictureBox schreiben
   Write_Letter tmpLetter
   'Größe des Buchstabens herausfinden
   Get_LetterSize
   'Buchstabe überarbeiten
   Revise_Letter
   'Größe nochmals ermitteln, da Kontrast sich geändert hat
   Get_LetterSize
   'Ausgelesenen Buchstaben zurückgeben
   Create_Letter = Read_Letter()
End Function

Private Sub Write_Letter(ByVal tmpLetter As String)
'--------------------------------------------------------------------'
'| Prozedur zum Schreiben des Buchstabens in picLetter              |'
'--------------------------------------------------------------------'
   If tmpLetter = " " Then tmpLetter = "T"
   'Inhalt der PictureBox löschen
   picLetter.Cls
   'Print-Methode benutzen
   picLetter.Print tmpLetter
   'picLetter refreshen
   picLetter.Refresh
End Sub

Private Sub Get_LetterSize()
'--------------------------------------------------------------------'
'| Prozedur zum Ermitteln der Größe des Buchstabens in picLetter    |'
'--------------------------------------------------------------------'
   'Ränder des Buchstabens initialisieren
   'Linken Rand nach rechts setzen
   Zeile_1 = 50
   'Rechten Rand nach links setzen
   Zeile_2 = 0
   'Oberen Rand nach unten setzen
   Spalte_1 = 50
   'Unteren Rand nach oben setzen
   Spalte_2 = 0

   'Bereich von 50x50 Pixeln durchsuchen
   For X = 0 To 50
      For Y = 0 To 50
         'Wenn Pixel schwarz ist
         If GetPixel(picLetter.hdc, X, Y) <> vbWhite Then
            'Wenn X kleiner als rechter Rand, dann rechten Rand auf X setzen
            If X < Spalte_1 Then Spalte_1 = X
            'Wenn X größer als linker Rand, dann linken Rand auf X setzen
            If X > Spalte_2 Then Spalte_2 = X
            'Wenn Y kleiner als oberer Rand, dann oberen Rand auf Y setzen
            If Y < Zeile_1 Then Zeile_1 = Y
            'Wenn Y größer als unterer Rand, dann unteren Rand auf Y setzen
            If Y > Zeile_2 Then Zeile_2 = Y
         End If
      Next Y
   Next X
   
   'Array anpassen
   ReDim arrPixel(Spalte_2 - Spalte_1, Zeile_2 - Zeile_1)
   
   'Höhe und Breite anpassen
   mHeight = Zeile_2 - Zeile_1 + 1
   mWidth = Spalte_2 - Spalte_1 + 1
   
   'Position festlegen
   mPosition = Zeile_1 - mFontPosition
End Sub

Private Sub Revise_Letter()
'--------------------------------------------------------------------'
'| Prozedur zum Überarbeiten des Buchstabens in picLetter           |'
'--------------------------------------------------------------------'
   'Bereich des Buchstabens überarbeiten
   For X = Spalte_1 To Spalte_2
      For Y = Zeile_1 To Zeile_2
         'Wenn Pixel dunkel ist..
         If Not Draw_Color2SW(GetPixel(picLetter.hdc, X, Y)) Then
            'Ein schwarzes Pixel darübersetzen
            SetPixel picLetter.hdc, X, Y, vbBlack
         'Wenn Pixel hell ist..
         Else
            'Ein weißes Pixel darübersetzen
            SetPixel picLetter.hdc, X, Y, vbWhite
         End If
      Next Y
   Next X
   'Picturebox refreshen
   picLetter.Refresh
End Sub

Private Function Read_Letter() As Boolean()
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen des Buchstabens in picLetter               |'
'--------------------------------------------------------------------'
   'Bereich des Buchstabens auslesen
   For X = Spalte_1 To Spalte_2
      For Y = Zeile_1 To Zeile_2
         'Wenn Pixel schwarz ist und der Buchstabe kein " " ist..
         If (GetPixel(picLetter.hdc, X, Y) = 0) And (mLetter <> " ") Then
            'Eine "1" ins Array setzen
            arrPixel(X - Spalte_1, Y - Zeile_1) = True
         'Wenn Pixel weiß ist..
         Else
            'Eine "0" ins Array setzen
            arrPixel(X - Spalte_1, Y - Zeile_1) = False
         End If
      Next Y
   Next X
   
   'Array zurückgeben
   Read_Letter = arrPixel
End Function

Public Sub Refresh_FontValues()
'--------------------------------------------------------------------'
'| Prozedur zum Aktualisieren der Font-Höhe                         |'
'--------------------------------------------------------------------'
   'Buchstabe "T" in PictureBox schreiben
   Write_Letter "T"
   'Größe des Buchstabens herausfinden
   Get_LetterSize
   'Buchstabe überarbeiten
   Revise_Letter
   'Buchstabe auslesen
   Read_Letter
   'Font-Höhe aktualisieren
   mFontHeight = mHeight
   'Font-Position aktualisieren
   mFontPosition = Zeile_1
End Sub

Private Function Draw_Color2RGB(tmpCol As Long) As tRGB
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben von RGB-Werten einer Long-Farbe         |'
'--------------------------------------------------------------------'
   'RGB-Werte errechnen und zurückgeben
   Draw_Color2RGB.R = (tmpCol And Mask_R) / 1
   Draw_Color2RGB.G = (tmpCol And Mask_G) / &H100&
   Draw_Color2RGB.B = (tmpCol And Mask_B) / &H10000
End Function

Private Function Draw_Color2SW(tmpCol As Long) As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben von Schwarz oder Weiß aus Farbe         |'
'--------------------------------------------------------------------'

   'Wennn Schwarz zurückgegeben werden soll, wird Draw_Color2SW
   'auf False gesetzt, bei weiß auf True
   
   'RGB-Wert der übergebenen Farbe errechnen
   SWRGB = Draw_Color2RGB(tmpCol)
   
   'Wenn Mittelwert größer als Schwellwert ist, dann Weiß zurückgeben, sonst schwarz
   If (SWRGB.R + SWRGB.G + SWRGB.B) / 3.5 > Schwellwert Then
      Draw_Color2SW = True
   Else
      Draw_Color2SW = False
   End If
End Function

Public Property Get FontName() As String
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mFontName                       |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   FontName = mFontName
End Property

Public Property Let FontName(ByVal tmpFontName As String)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen der Variablen mFontName                      |'
'--------------------------------------------------------------------'
   'Bei einem Fehler ("Schrifart nicht vorhanden") zu ErrHandler gehen
   On Error GoTo ErrHandler
   'Alte Schriftart aus PictureBox in mFontName sichern
   mFontName = picLetter.FontName
   'Neue Schriftart der PictureBox zuweisen
   picLetter.FontName = tmpFontName
   '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   'Hier würde der Fehler ausgelöst werden,
   'wenn die Schriftart nicht vorhanden ist
   
   'Neue Schriftart der Variable mFontName zuweisen
   mFontName = picLetter.FontName
   'Prozedur beenden (Nicht in Fehlerbehandlung springen)
   Exit Property
   
ErrHandler:
   'Fehler-Event auslösen
   RaiseEvent Error("Schriftart """ & tmpFontName & """ nicht vorhanden !")
End Property

Public Property Get FontSize() As Integer
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mFontSize                       |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   FontSize = mFontSize
End Property

Public Property Let FontSize(ByVal tmpFontSize As Integer)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen der Variablen mFontSize                      |'
'--------------------------------------------------------------------'
   'Bei einem Fehler ("Schrifgröße ungültig") zu ErrHandler gehen
   On Error GoTo ErrHandler
   'Alte Schriftgröße aus PictureBox in mFontSize sichern
   mFontSize = picLetter.FontSize
   'Neue Schriftgröße der PictureBox zuweisen
   picLetter.FontSize = tmpFontSize
   '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
   'Hier würde der Fehler ausgelöst werden,
   'wenn die Schriftgröße ungültig ist
   
   'Neue Schriftgröße der Variable mFontSize zuweisen
   mFontSize = picLetter.FontSize
   'Prozedur beenden (Nicht in Fehlerbehandlung springen)
   Exit Property
   
ErrHandler:
   'Fehler-Event auslösen
   RaiseEvent Error("Schriftgröße """ & CStr(tmpFontSize) & """ nicht gültig !")
End Property

Public Property Get FontBold() As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mFontBold                       |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   FontBold = mFontBold
End Property

Public Property Let FontBold(ByVal tmpFontBold As Boolean)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen der Variablen mFontBold                      |'
'--------------------------------------------------------------------'
   '"Fett/Nicht Fett" der PictureBox zuweisen
   picLetter.FontBold = tmpFontBold
   '"Fett/Nicht Fett" der Variable mFontBold zuweisen
   mFontBold = picLetter.FontBold
End Property

Public Property Get FontItalic() As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mFontItalic                     |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   FontItalic = mFontItalic
End Property

Public Property Let FontItalic(ByVal tmpFontItalic As Boolean)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen der Variablen mFontItalic                    |'
'--------------------------------------------------------------------'
   '"Kursiv/Nicht Kursiv" der PictureBox zuweisen
   picLetter.FontItalic = tmpFontItalic
   '"Kursiv/Nicht Kursiv" der Variable mFontItalic zuweisen
   mFontItalic = picLetter.FontItalic
End Property

Public Property Get FontUnderline() As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mFontUnderline                  |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   FontUnderline = mFontUnderline
End Property

Public Property Let FontUnderline(ByVal tmpFontUnderline As Boolean)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen der Variablen mFontUnderline                 |'
'--------------------------------------------------------------------'
   '"Unterstrichen/Nicht Unterstrichen" der PictureBox zuweisen
   picLetter.FontUnderline = tmpFontUnderline
   '"Unterstrichen/Nicht Unterstrichen" der Variable mFontUnderline zuweisen
   mFontUnderline = picLetter.FontUnderline
End Property

Public Property Get FontStrikethru() As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mFontStrikethru                 |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   FontStrikethru = mFontStrikethru
End Property

Public Property Let FontStrikethru(ByVal tmpFontStrikethru As Boolean)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen der Variablen mFontStrikethru                |'
'--------------------------------------------------------------------'
   '"Durchgestrichen/Nicht Durchgestrichen" der PictureBox zuweisen
   picLetter.FontStrikethru = tmpFontStrikethru
   '"Durchgestrichen/Nicht Durchgestrichen" der Variable mFontStrikethru zuweisen
   mFontStrikethru = picLetter.FontStrikethru
End Property

Public Property Get Height() As Integer
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mHeight                         |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   Height = mHeight
End Property

Public Property Get Width() As Integer
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mWidth                          |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   Width = mWidth
End Property

Public Property Get FontHeight() As Integer
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mFontHeight                     |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   FontHeight = mFontHeight
End Property

Public Property Get Position()
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Variablen mPosition                       |'
'--------------------------------------------------------------------'
   'Variable zurückgeben
   Position = mPosition
End Property
