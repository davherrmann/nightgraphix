Attribute VB_Name = "mdl_text"
'_________________________________________________________________________________'
'|                               MODUL mdl_text                                  |'
'| Dieses Modul beinhaltet Routinen zum Schreiben von Text in das Modell.        |'
'|                                                                               |'
'---------------------------------------------------------------------------------'
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Variablen für Schleifen
Private x As Long
Private y As Long
Private i As Long
'In welche Richtung wird der Text geschrieben (Anti-/Clockwise)
'Im Uhrzeigersinn:         True
'Gegen den Uhrzeigersinn:  False
Public TextClockWise As Boolean

Public Sub Text_Write(ByRef picsource As PictureBox, ByRef pictarget As PictureBox, progbar As Object, tmptext As String, tmpSpalte As Long, tmpZeile As Long, tmpcolor As Long)
'--------------------------------------------------------------------'
'| Prozedur zum Schreiben des Textes in das Modell                  |'
'--------------------------------------------------------------------'
   'Bei einem Fehler weitermachen
   On Error Resume Next
   
   'Progressbar einstellen (Val = 0, Max = Länge des Textes)
   progbar.Value = 0
   progbar.Max = Len(tmptext)
   
   'i auf 0 stellen
   i = 0
   
   'Schleife zum Schreiben des Textes
   For i = 1 To Len(tmptext)
      'Zeichen in ein Array schreiben
      ZeichenArray = Letter.Create_Letter(tmptext)
      
      'Array mit Hintergründen vergrößern (RGB)
      ReDim Preserve OldLetterArrayRGB(ZeichenAnzahl)
      'Array mit Hintergründen vergrößern (SW)
      ReDim Preserve OldLetterArraySW(ZeichenAnzahl)
      'Zeichenanzahl um 1 erhöhen
      ZeichenAnzahl = ZeichenAnzahl + 1
      
      'Wenn der Text im Uhrzeigersinn geschrieben werden soll..
      If TextClockWise = True Then
         'Variable für alten Hintergrund anpassen (RGB)
         ReDim OldLetterArrayRGB(ZeichenAnzahl - 1).Data(UBound(ZeichenArray, 1), UBound(ZeichenArray, 2))
         'Variable für alten Hintergrund anpassen (SW)
         ReDim OldLetterArraySW(ZeichenAnzahl - 1).Data(Letter.Width - 1, Letter.Height - 1)
         
         'Position des Buchstabens speichern (RGB)
         OldLetterArrayRGB(ZeichenAnzahl - 1).Left = tmpSpalte
         OldLetterArrayRGB(ZeichenAnzahl - 1).Top = tmpZeile - Letter.Position - UBound(ZeichenArray, 2) + Letter.FontHeight - 1
         'Position des Buchstabens speichern (SW)
         OldLetterArraySW(ZeichenAnzahl - 1).Left = tmpSpalte
         OldLetterArraySW(ZeichenAnzahl - 1).Top = tmpZeile - Letter.Position - UBound(ZeichenArray, 2) + Letter.FontHeight - 1
         
         'Schleife zum Eintragen des Textes
         For x = tmpSpalte To tmpSpalte + UBound(ZeichenArray, 1)
            For y = tmpZeile - Letter.Position + Letter.FontHeight - 1 To tmpZeile - Letter.Position - UBound(ZeichenArray, 2) + Letter.FontHeight - 1 Step -1
               'Array für alten Hintergrund füllen (RGB)
               OldLetterArrayRGB(ZeichenAnzahl - 1).Data(x - tmpSpalte, y - tmpZeile + Letter.Position + UBound(ZeichenArray, 2) - Letter.FontHeight + 1) = Array_Red((x - 1) Mod Spalten + 1, y - 1) * 100 + Array_Green((x - 1) Mod Spalten + 1, y - 1) * 10 + Array_Blue((x - 1) Mod Spalten + 1, y - 1)
               'Array für alten Hintergrund füllen (SW)
               OldLetterArraySW(ZeichenAnzahl - 1).Data(x - tmpSpalte, y - tmpZeile + Letter.Position + UBound(ZeichenArray, 2) - Letter.FontHeight + 1) = CBool(Array_SW((x - 1) Mod Spalten + 1, y - 1))
               If ZeichenArray(x - tmpSpalte, tmpZeile - Letter.Position + Letter.FontHeight - 1 - y) Then
                  'Zelle füllen, ohne Array anzupassen
                  Draw_FillCell (x - 1) Mod Spalten + 1, y - 1, picsource, tmpcolor, False
                  
                  'Array anpassen, geht schneller als in Draw_FillCell
                  'RGB-Arrays anpassen
                  Array_Red((x - 1) Mod Spalten + 1, y - 1) = Draw_Color2RGB_Bool(tmpcolor).R / 255
                  Array_Green((x - 1) Mod Spalten + 1, y - 1) = Draw_Color2RGB_Bool(tmpcolor).G / 255
                  Array_Blue((x - 1) Mod Spalten + 1, y - 1) = Draw_Color2RGB_Bool(tmpcolor).B / 255
                  
                  'S/W-Array anpassen
                  Array_SW((x - 1) Mod Spalten + 1, y - 1) = IIf(Draw_Color2SW(tmpcolor), 1, 0)
               End If
            Next y
         Next x
      'Wenn gegen den Uhrzeigersinn geschrieben werden soll..
      Else
         'Variable für alten Hintergrund anpassen (RGB)
         ReDim OldLetterArrayRGB(ZeichenAnzahl - 1).Data(UBound(ZeichenArray, 1), UBound(ZeichenArray, 2))
         'Variable für alten Hintergrund anpassen (SW)
         ReDim OldLetterArraySW(ZeichenAnzahl - 1).Data(Letter.Width - 1, Letter.Height - 1)
         
         'Position des Buchstabens speichern (RGB)
         OldLetterArrayRGB(ZeichenAnzahl - 1).Left = tmpSpalte
         OldLetterArrayRGB(ZeichenAnzahl - 1).Top = tmpZeile + Letter.Position
         'Position des Buchstabens speichern (SW)
         OldLetterArraySW(ZeichenAnzahl - 1).Left = tmpSpalte
         OldLetterArraySW(ZeichenAnzahl - 1).Top = tmpZeile + Letter.Position
         
         'Schleife zum Eintragen des Textes
         For x = tmpSpalte To tmpSpalte - UBound(ZeichenArray, 1) Step -1
            For y = tmpZeile + Letter.Position To tmpZeile + Letter.Position + UBound(ZeichenArray, 2)
               'Array für alten Hintergrund füllen (RGB)
               OldLetterArrayRGB(ZeichenAnzahl - 1).Data(x - tmpSpalte, y - tmpZeile + Letter.Position) = Array_Red((x - 1) Mod Spalten + 1, y - 1) * 100 + Array_Green((x - 1) Mod Spalten + 1, y - 1) * 10 + Array_Blue((x - 1) Mod Spalten + 1, y - 1)
               'Array für alten Hintergrund füllen (SW)
               OldLetterArraySW(ZeichenAnzahl - 1).Data(x - tmpSpalte, y - tmpZeile - Letter.Position) = CBool(Array_SW((x - 1) Mod Spalten + 1, y - 1))
               If ZeichenArray(tmpSpalte - x, y - tmpZeile - Letter.Position) Then
                  'Zelle füllen, ohne Array anzupassen
                  Draw_FillCell (x - 1) Mod Spalten + 1, y - 1, picsource, tmpcolor, False
                  'Array anpassen, geht schneller als in Draw_FillCell
                  'RGB-Arrays anpassen
                  Array_Red((x - 1) Mod Spalten + 1, y - 1) = Draw_Color2RGB_Bool(tmpcolor).R / 255
                  Array_Green((x - 1) Mod Spalten + 1, y - 1) = Draw_Color2RGB_Bool(tmpcolor).G / 255
                  Array_Blue((x - 1) Mod Spalten + 1, y - 1) = Draw_Color2RGB_Bool(tmpcolor).B / 255
                  'S/W-Array anpassen
                  Array_SW((x - 1) Mod Spalten + 1, y - 1) = IIf(Draw_Color2SW(tmpcolor), 1, 0)
               End If
            Next y
         Next x
      End If

      'Progressbar erhöhen
      progbar.Value = i
   Next i
   
   'Progressbar auf 0 setzen
   progbar.Value = 0
      
   'Modell neu zeichnen
   Draw_Zoom picsource, pictarget
End Sub

Public Function Text_OpenFontDialog(ByRef tmpCommonDialog As CommonDialog, ByRef tmpPicSource As PictureBox, ByRef tmpPicTarget As PictureBox, ByRef tmpTmrCursor As Timer)
'--------------------------------------------------------------------'
'| Prozedur zum Öffnen des Font-Dialoges                            |'
'--------------------------------------------------------------------'
   'Bei einem Fehler (Abbruch) zu ErrHandler springen
   On Error GoTo ErrHandler
   
   'Wenn das Text-Tool nicht gewählt ist, Prozedur beenden
   If Tool <> Text Then Exit Function
   
   'CommonDialog selektieren
   With tmpCommonDialog
      'Bei Abbruch einen Fehler auslösen
      .CancelError = True
      
      'Die Flags-Eigenschaft muss auf cdlCFScreenFonts,
      'cdlCFPrinterFonts oder cdlCFBoth gesetzt werden,
      'bevor der Font-Dialog geöffnet wird,
      'sonst tritt der Fehler "Keine Schriftarten vorhanden" auf.
      .flags = cdlCFEffects Or cdlCFScreenFonts
      
      'Font-Eigenschaften einstellen
      .FontBold = Letter.FontBold
      .FontItalic = Letter.FontItalic
      .FontName = Letter.FontName
      .FontSize = Letter.FontSize
      .FontStrikethru = Letter.FontStrikethru
      .FontUnderline = Letter.FontUnderline
      
      'Font-Dialog anzeigen
      .ShowFont
      
      'Wenn kein Fehler auftrat und der Timer aktiviert ist..
      If (Err = 0) And tmpTmrCursor.Enabled Then
         'Timer für den Cursor ausschalten
         tmpTmrCursor.Enabled = False
         'Cursor ausschalten, wenn noch sichtbar
         If CursorVisible = True Then
            'Cursor ausschalten
            Draw_Cursor CursorPosition.x, CursorPosition.y, tmpPicSource, False
            'Bild refreshen
            Draw_Zoom tmpPicSource, tmpPicTarget
            'Cursor soll am Anfang wieder gezeichnet werden
            CursorVisible = False
         End If
         'Font-Eigenschaften an die Klasse cls_font übergeben
         Letter.FontName = .FontName
         Letter.FontBold = .FontBold
         Letter.FontItalic = .FontItalic
         Letter.FontSize = .FontSize
         Letter.FontStrikethru = .FontStrikethru
         Letter.FontUnderline = .FontUnderline
         'Font-Werte aktualisieren
         Letter.Refresh_FontValues
         'Cursor-Höhe speichern
         CursorHeight = Letter.Height
         'Neue Cursor-Position bestimmen
         CursorPosition.y = CursorPosition.y - CursorHeight + Letter.Height
         'Wenn Cursor-Position über LED-Anzahl - 1 ist, dann auf LED-Anzahl - 1 setzen
         If CursorPosition.y > Leds - 1 Then CursorPosition.y = Leds - 1
         'Cursor wieder initialisieren
         Draw_InitCursor CursorPosition.x, CursorPosition.y, tmpPicSource
         'Timer für Cursor anschalten
         tmpTmrCursor.Enabled = True
      'Wenn ein Fehler auftrat oder der Timer nicht aktiviert ist..
      Else
         'Font-Eigenschaften an die Klasse cls_font übergeben
         Letter.FontName = .FontName
         Letter.FontBold = .FontBold
         Letter.FontItalic = .FontItalic
         Letter.FontSize = .FontSize
         Letter.FontStrikethru = .FontStrikethru
         Letter.FontUnderline = .FontUnderline
         'Font-Werte aktualisieren
         Letter.Refresh_FontValues
      End If
   End With
   
'Fehlerbehandlung
ErrHandler:
   'Abbruch wurde gewählt..
   'Keine Fehlerbehandlung notwendig
End Function
