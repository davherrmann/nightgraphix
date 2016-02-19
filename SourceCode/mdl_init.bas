Attribute VB_Name = "mdl_init"
'_________________________________________________________________________________'
'|                               MODUL mdl_init                                  |'
'| Dieses Modul beinhaltet Initalisierungs-Routinen, wie z.B. das Einstellen     |'
'| von Variablen.                                                                |'
'---------------------------------------------------------------------------------'

Option Explicit

'VARIABLEN..
'Farbe der LED
Public LEDColor As Long
'Index der SW-Farbe (Für Option-Buttons)
Public SWColorIndex As Integer
'Index der RGB-Farbe (Für Option-Buttons)
Public RGBColorIndex As Integer
'Standard Textzeile
Public StandardTextZeile As Integer
'InnenkreisRadius in Spalten
Public Hub As Long
'Abstand der Spalten
Public SpaltenAbstand As Double
'Das ist der Außenkreisradius, also LED-Höhe + Innenkreisradius
Public KreisRadius As Single
'Radius des ungefüllten Kreises
Public InnenKreisRadius As Single
'Variable zur Bestimmung in welchem Viertel der zu füllende
'Punkt liegt
Public KreisViertel As Integer
'Höhe des Arrays, also die Anzahl der LEDs
Public Leds As Integer
'Breite des Arrays, also Schritte pro Umdrehung
Public Spalten As Integer
'Variablen in denen die Position des aktuellen Array-Feldes
'gespeichert wird
Public HöheFeld As Integer
Public BreiteFeld As Integer
'Winkel-Variable
Public Winkel As Double
'Abstände des Außenkreises von Links und Oben
Public AbstandX As Integer
Public AbstandY As Integer
'Wieviele Zeichen wurden geschrieben ?
Public ZeichenAnzahl As Integer
'Zwischenspeicher für Buchstaben
Public Zeichen As String
'Zwischenspeicher für ZeichenArray
Public ZeichenArray() As Boolean
'Abstand zwischen Zeichen
Public ZeichenAbstand As Integer
'Array für das Benutzen der Backspace-Taste (Schwarz-Weiß)
Public OldLetterArraySW() As tLetter
'Array für das Benutzen der Backspace-Taste (RGB)
Public OldLetterArrayRGB() As tLetter
'Variablen für Schleifen
Private x As Integer
Private y As Integer
'Variable für Zähler, Schleifen etc.
Private i As Integer

'KONSTANTEN..
'Korrekturwert fürs Zeichnen
Public Const KorrekturWert = 0.01

'APIs..
'API zum Initialisieren des XP-Styles
Public Declare Sub InitCommonControls Lib "comctl32" ()

'TYPEs..
'Type zum Speichern der Pixel unter einem Buchstaben (Backspace)
Public Type tLetter
   'Daten (Pixelarray)
   Data() As Byte
   'Position: Left
   Left As Integer
   'Position: Top
   Top As Integer
End Type

'REFERENZEN..
'Referenz auf cls_font
Public Letter As cls_font

Public Sub Init_Variables(picsource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren der Variablen                        |'
'--------------------------------------------------------------------'
   'Es soll im Uhrzeigersinn geschrieben werden
   TextClockWise = True
   'Entfernung des Außenkreises zu den Rändern
   AbstandX = 38
   AbstandY = 38
   'Bild-Nummer
   NumberOfPicture = 1
   'Abstand der Spalten
   SpaltenAbstand = (picsource.Width / 2 - AbstandX) / (Leds + Hub)
   'Innenkreisradius
   InnenKreisRadius = Hub * SpaltenAbstand
   'Außenkreisradius
   KreisRadius = SpaltenAbstand * (Leds + Hub)
   'Radius des Offset-Pfeils
   RadiusOffsetArrow = Hub * SpaltenAbstand
   'Cursor soll am Anfang gezeichnet werden
   CursorVisible = False
   'Zeichenabstand setzen
   ZeichenAbstand = 3
   'AnimationFrames auf 2 setzen, wenn es kleiner als 2 ist
   If (NGAnimationFrames < 2) Or (NGAnimationFrames > 5) Then NGAnimationFrames = 2
   'Wenn die obere LED-Farbe noch nicht gesetzt ist..
   If NGTopLEDColor = 0 Then NGTopLEDColor = vbYellow
   'Wenn die untere LED-Farbe noch nicht gesetzt ist..
   If NGBottomLEDColor = 0 Then NGBottomLEDColor = vbBlue
   
   'Wenn SW-Version..
   If Not RGBVersion Then
      'LedColor auf Ober/Unterseite setzen
      LEDColor = IIf((Not NGRotationSystemLeft) = (NGRotationMCRight), NGTopLEDColor, NGBottomLEDColor)
   End If
   
   'Wenn ein Wert des NGRotorSizeArray 0 ist..
   If NGRotorSizeArray(6) = 0 Then
      'Alle Werte initialisieren
      NGRotorSizeArray(0) = 650
      NGRotorSizeArray(1) = 700
      NGRotorSizeArray(2) = 1000
      NGRotorSizeArray(3) = 1100
      NGRotorSizeArray(4) = 1350
      NGRotorSizeArray(5) = 1500
      NGRotorSizeArray(6) = 1650
   End If
   
   'Font-Referenz setzen
   Set Letter = New cls_font
   
   'Standard: Textzeile = 1
   StandardTextZeile = 1
   
   'Masken für Draw_Color2RGB_Bool initialisieren
   'Maske: Rot
   Mask_R = &HFF&
   'Maske: Grün
   Mask_G = &HFF00&
   'Maske: Blau
   Mask_B = &HFF0000
   
   'Am Anfang sind alle Änderungen gespeichert
   SavedChanges = True
   'DateiName am Anfang "Untitled"
   DateiName = "Untitled"
   'Am Anfang muss der Dialog beim Speichern erscheinen
   DateiSpeichernUnter = False
   
   'DateiHeader ändern
   'StandardFarbe: RGB(1*255, 0*255 ,0*255)
   NGHeader.m_lColor.Data = 100
   'Spalten
   NGHeader.m_lColumns.Data = CStr(Format(Spalten, "0000"))
   'LEDs
   NGHeader.m_lRows.Data = CStr(Format(Leds, "0000"))
   '32.LED
   NGHeader.m_bLastLED.Data = 0
   'RGB/SW
   NGHeader.m_bRGB.Data = IIf(RGBVersion = True, 1, 0)
   'Zusätzliches
   NGHeader.m_sAdditional.Data = String(10, ".")
   'Name: NG
   NGHeader.m_sName.Data = "NG"
   'Version
   NGHeader.m_sVersion.Data = "01.5"
End Sub

Public Sub Init_XPStyle()
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren des XP-Styles                        |'
'--------------------------------------------------------------------'
   'XP-Style initialisieren
   InitCommonControls
End Sub

Public Sub Init_Form(frmForm As Form)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren der Form                             |'
'--------------------------------------------------------------------'
   'FormNamen ändern
   frmForm.Caption = DateiName & " - NightGraphix V1.0"
End Sub

Public Sub Init_Arrays()
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren der Arrays                           |'
'--------------------------------------------------------------------'
   'Arrays den vorgegebenen Größen anpassen
   'RGB-Arrays
   ReDim Array_Red(1 To Spalten, 1 To Leds)
   ReDim Array_Green(1 To Spalten, 1 To Leds)
   ReDim Array_Blue(1 To Spalten, 1 To Leds)
   'SW-Arrays
   ReDim Array_SW(1 To Spalten, 1 To Leds)
   
   'Alle Arrays mit 1sen füllen
   For x = 1 To Spalten
      For y = 1 To Leds
         'RGB-Arrays
         Array_Red(x, y) = 0
         Array_Green(x, y) = 0
         Array_Blue(x, y) = 0
         'SW-Array
         Array_SW(x, y) = 0
      Next y
   Next x
End Sub

Public Sub Init_PictureBoxes(picsource As PictureBox, pictarget As PictureBox, picrahmen As PictureBox, picbild As PictureBox, picbildgroß As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Einstellen der PictureBoxes                         |'
'--------------------------------------------------------------------'
   'ScaleMode auf Pixel setzen
   picrahmen.ScaleMode = vbPixels
   picsource.ScaleMode = vbPixels
   pictarget.ScaleMode = vbPixels
   picbild.ScaleMode = vbPixels
   picbildgroß.ScaleMode = vbPixels
      
   'AutoRedraw bei Source und Target auf True setzen
   picsource.AutoRedraw = True
   pictarget.AutoRedraw = True
   picbild.AutoRedraw = True
   picbildgroß.AutoRedraw = True
   
   'BackColor einstellen
   picsource.BackColor = vbWhite
   
   'Position und Größe der PictureBoxen einstellen
   pictarget.Height = 266.99
   pictarget.Width = 266.99
   
   'pic_source einstellen
   picsource.Height = pictarget.Height * 15
   picsource.Width = pictarget.Width * 15
   picsource.Top = 1000
   picsource.Left = 100
   
   'pic_bild einstellen
   picbild.AutoSize = True
   picbild.Top = 1000
   picbild.Left = 100
   
   'pic_target einstellen
   pictarget.Height = picsource.Height * 3 / 4
   pictarget.Width = picsource.Width * 3 / 4
   pictarget.Left = 0
   pictarget.Top = 0
   
   'pic_bildgroß einstellen
   picbildgroß.Top = 1000
   picbildgroß.Left = 100
   picbildgroß.Width = picsource.Width
   picbildgroß.Height = picsource.Width
   
   'Verhältnis zwischen Source und Target ausrechnen
   ZoomX = picsource.Width / pictarget.Width
   ZoomY = picsource.Height / pictarget.Height
End Sub

Public Sub Init_Scrollbars(vscrollbar As Object, hscrollbar As Object, picrahmen As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Einstellen der Scrollbars                           |'
'--------------------------------------------------------------------'
   'Large- und Smallchange an picrahmen anpassen
   vscrollbar.LargeChange = picrahmen.Height / 100
   vscrollbar.SmallChange = 100
   hscrollbar.LargeChange = picrahmen.Width / 100
   hscrollbar.SmallChange = 100
   
   'Modell beim Start in die Mitte scrollen
   '(Hälfte von den Maximalwerten)
   vscrollbar.Value = vscrollbar.Max / 2
   hscrollbar.Value = hscrollbar.Max / 2
End Sub

Public Sub Init_Menu(mnubild As Object)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren des Bilder-Menüs                     |'
'--------------------------------------------------------------------'
   'Bilder-Menü initialisieren
   mnubild(0).Checked = True
   mnubild(1).Checked = False
   mnubild(2).Checked = False
   mnubild(3).Checked = False
End Sub

Public Sub Init_ComboBox(tmpComboBox As ComboBox, tmpItem As String)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren der ComboBox                         |'
'--------------------------------------------------------------------'
   'Eintrag hinzufügen
   tmpComboBox.AddItem tmpItem
   'ListIndex setzen
   tmpComboBox.ListIndex = 0
End Sub

Public Sub Init_ComboBoxPorts(ByRef tmpComboBox As ComboBox, ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Anzeigen der vorhandenen COM-Ports                  |'
'--------------------------------------------------------------------'
   'Schleife von 1-16 (mögliche COM-Ports zum Öffnen)
   For i = 1 To 16
      'Wenn der COM-Port sich öffnen lässt..
      If ComPort_Available(tmpMSComm, i) = True Then
         'COM-Port in ComboBox eintragen
         tmpComboBox.AddItem "COM-Port " & CStr(i)
      End If
   Next i
End Sub

Public Sub Init_ImageList(ByRef imagelist As imagelist)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren der Imagelist                        |'
'--------------------------------------------------------------------'
   'Imagelist löschen
   imagelist.ListImages.Clear
   
   For x = 1 To 45
      'Bilder der ImageList hinzufügen
      imagelist.ListImages.Add , , LoadResPicture(x + 100, vbResIcon)
   Next x
End Sub

Public Sub Init_Toolbar(ByRef tmpToolBar As Toolbar)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren der Toolbar                          |'
'--------------------------------------------------------------------'
   'Jedem einzelnen Button ein Bild zuweisen
   'Neue Datei
   tmpToolBar.Buttons(1).Image = 22
   'Datei öffnen
   tmpToolBar.Buttons(2).Image = 23
   'Bild importieren
   tmpToolBar.Buttons(3).Image = 19
   'Datei speichern
   tmpToolBar.Buttons(4).Image = 34
   'Datei speichern unter
   tmpToolBar.Buttons(5).Image = 35
   '------------------
   'Löschen
   tmpToolBar.Buttons(7).Image = 8
   '------------------
   'Zoom+
   tmpToolBar.Buttons(9).Image = 39
   'Zoom-
   tmpToolBar.Buttons(10).Image = 40
   '------------------
   'EEPROM beschreiben
   tmpToolBar.Buttons(12).Image = 38
   'EEPROM auslesen
   tmpToolBar.Buttons(13).Image = 31
   '------------------
   'Verbinden
   tmpToolBar.Buttons(15).Image = 43
   'Trennen
   tmpToolBar.Buttons(16).Image = 42
   'Demo starten
   tmpToolBar.Buttons(17).Image = 45
   'Demo beenden
   tmpToolBar.Buttons(18).Image = 44
   '------------------
   'Optionen Software
   tmpToolBar.Buttons(20).Image = 26
   'Optionen Hardware
   tmpToolBar.Buttons(21).Image = 25
   'Hilfe
   tmpToolBar.Buttons(22).Image = 18
   '------------------
   'NG1.0 beenden
   tmpToolBar.Buttons(24).Image = 14
End Sub

Public Sub Init_FontPictureBox(ByRef tmpPictureBox As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren von pic_letter                       |'
'--------------------------------------------------------------------'
   'pic_letter an clsFont übergeben
   Letter.Initialize tmpPictureBox
   'FontName einstellen
   Letter.FontName = "Arial"
   'FontSize einstellen
   Letter.FontSize = 12
   'Font-Werte aktualisieren
   Letter.Refresh_FontValues
End Sub

Public Sub Init_FileHeader()
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren des Datei-Headers                    |'
'--------------------------------------------------------------------'
   'NG-Header initialisieren
   'Name
   NGHeader.m_sName.DataTyp = tString
   NGHeader.m_sName.DataLength = 2
   NGHeader.m_sName.DataStart = 1
   
   'Version
   NGHeader.m_sVersion.DataTyp = tString
   NGHeader.m_sVersion.DataLength = 4
   NGHeader.m_sVersion.DataStart = NGHeader.m_sName.DataStart + NGHeader.m_sName.DataLength
   
   'Spalten
   NGHeader.m_lColumns.DataTyp = tLong
   NGHeader.m_lColumns.DataLength = 4
   NGHeader.m_lColumns.DataStart = NGHeader.m_sVersion.DataStart + NGHeader.m_sVersion.DataLength
   
   'Zeilen
   NGHeader.m_lRows.DataTyp = tLong
   NGHeader.m_lRows.DataLength = 4
   NGHeader.m_lRows.DataStart = NGHeader.m_lColumns.DataStart + NGHeader.m_lColumns.DataLength
   
   'Gewählte Farbe
   NGHeader.m_lColor.DataTyp = tLong
   NGHeader.m_lColor.DataLength = 3
   NGHeader.m_lColor.DataStart = NGHeader.m_lRows.DataStart + NGHeader.m_lRows.DataLength
   
   '32. LED
   NGHeader.m_bLastLED.DataTyp = tBoolean
   NGHeader.m_bLastLED.DataLength = 1
   NGHeader.m_bLastLED.DataStart = NGHeader.m_lColor.DataStart + NGHeader.m_lColor.DataLength
   
   'RGB/SW
   '1 bei RGB, 0 bei SW
   NGHeader.m_bRGB.DataTyp = tBoolean
   NGHeader.m_bRGB.DataLength = 1
   NGHeader.m_bRGB.DataStart = NGHeader.m_bLastLED.DataStart + NGHeader.m_bLastLED.DataLength
   
   'Zusätzliches
   NGHeader.m_sAdditional.DataTyp = tString
   NGHeader.m_sAdditional.DataLength = 10
   NGHeader.m_sAdditional.DataStart = NGHeader.m_bRGB.DataStart + NGHeader.m_bRGB.DataLength
End Sub
