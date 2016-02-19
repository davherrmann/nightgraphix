Attribute VB_Name = "mdl_filesystem"
'_________________________________________________________________________________'
'|                               MODUL mdl_filesystem                            |'
'| Dieses Modul beinhaltet Routinen für das Öffnen und Speichern von Dateien     |'
'|                                                                               |'
'---------------------------------------------------------------------------------'

Option Explicit

'VARIABLEN..
'Type für DatenTypen
Public Enum eNGHeaderPartTyp
   'Boolean
   tBoolean = 1
   'Integer
   tInteger = 2
   'Long
   tLong = 3
   'String
   tString = 4
   'Variant
   tVariant = 5
End Enum

'Type für Header-Teile
Private Type tNGHeaderPart
   'Feste Länge
   DataLength As Long
   'Daten
   Data As Variant
   'Typ
   DataTyp As eNGHeaderPartTyp
   'Start der Daten
   DataStart As Long
End Type

'Type für NG-Header
Private Type tNGHeader
   'Name: NG
   m_sName As tNGHeaderPart
   'Version: xx.x
   m_sVersion As tNGHeaderPart
   'Spalten: xxxx
   m_lColumns As tNGHeaderPart
   'Zeilen: xxxx
   m_lRows As tNGHeaderPart
   'Farbe: xxx (R,G,B)
   m_lColor As tNGHeaderPart
   '32 LED: x
   m_bLastLED As tNGHeaderPart
   'Farbe oder SW: x
   m_bRGB As tNGHeaderPart
   'Zusätzlich: xxxxxxxxxx
   m_sAdditional As tNGHeaderPart
End Type
'Representation für Type
Public NGHeader As tNGHeader

'Gespeichert oder nicht ?
Public SavedChanges As Boolean
'Name der Datei
Public DateiName As String
'Datei
Public Datei As String
'DateiPfad
Public DateiPfad As String
'Datei Speichern oder Speichern unter ?
Public DateiSpeichernUnter As Boolean
'Pfad des importierten Bildes
Public BildPfad As String
'Referenz auf Wscript
Public WScript As Object
'Pfad für die Registry
Public RegistryPath As String
'Soll beim Starten eine Datei geöffnet werden ?
Public File2Open As String

'Soll die "Ungültige Datei"-Fehlermeldung unterdrückt werden ?
Private HideFileError As Boolean
'FileNumber
Private FreeFileNumber As Long

'Virtuelle LogDatei
Private LogFile As String

'Variable für Path2File
Private Pos As Long

'Variablen für Schleifen
Private x As Long, y As Long
Private i As Long

'KONSTANTEN..
'Konstanten für Filter
Public Const Filter_Graphik = "Graphikdateien (*.bmp, *.jpg, *.jpeg, *.gif)|*.bmp;*.jpg;*.jpeg;*.gif"
Public Const Filter_Text = "TextDateien (*.txt)|*.txt"
Public Const Filter_NightGraphix = "NightGraphix-Dateien (*.ng)|*.ng"

'APIs..
Private Declare Function FindFirstFile Lib "kernel32" _
        Alias "FindFirstFileA" (ByVal lpFileName As String, _
        lpFindFileData As WIN32_FIND_DATA) As Long
        
Private Declare Function FindNextFile Lib "kernel32" _
        Alias "FindNextFileA" (ByVal hFindFile As Long, _
        lpFindFileData As WIN32_FIND_DATA) As Long
        
Private Declare Function FindClose Lib "kernel32" (ByVal _
        hFindFile As Long) As Long

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Const MAX_PATH = 259

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_COMPRESSED = &H800
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

Public Function FileSystem_OpenDialog(cdialog As Object, tmpFilter As String, Typ As String) As String
'--------------------------------------------------------------------'
'| Prozedur zum Öffnen des Datei-Dialoges                           |'
'--------------------------------------------------------------------'
   'Typ: Entweder "Save" oder "Open"

   'Bei einem Fehler Fehlerbehandlung ausführen
   On Error GoTo ErrHandler
   
   'CommonDialog einstellen und öffnen
   With cdialog
      'Bei Abbruch einen Laufzeitfehler auslösen
      .CancelError = True
      
      'Startpfad
      .InitDir = App.Path
      
      'Filter: Was für Dateitypen dürfen genutzt werden ?
      .Filter = tmpFilter
      
      'Open- oder Save-Dialog
      If UCase(Typ) = UCase("Open") Then
         'Flags setzen
         .flags = cdlOFNFileMustExist
         'Dialog öffnen ("Open")
         .ShowOpen
      ElseIf UCase(Typ) = UCase("Save") Then
         'Flags einstellen
         .flags = cdlOFNOverwritePrompt
         'Dialog öffnen ("Save")
         .ShowSave
      Else
         GoTo ErrHandler
      End If
      
      'DoEvents
      DoEvents
      
      'Dateiname zurückgeben
      FileSystem_OpenDialog = .FileName
   End With
   
   'Prozedur beenden
   Exit Function
   
'Fehlerbehandlungsroutine
ErrHandler:
   'Wenn im Dialog Abbrechen gewählt wurde
   'Braucht man eigentlich nicht, nur zu Übersicht
   If Err.Number = cdlCancel Then
      FileSystem_OpenDialog = ""
      Exit Function
   End If
End Function

Public Sub FileSystem_CreateDatei(cdialog As Object, progbar As Object, SpeichernUnter As Boolean)
'--------------------------------------------------------------------'
'| Prozedur zum Abspeichern einer Datei                             |'
'--------------------------------------------------------------------'

   'Variable auf "" setzen
   Datei = ""
   
   'DateiHeader einfügen
   FileSystem_CreateDateiHeader
   
   'ProgressBar einstellen
   progbar.Max = CLng(Spalten) * Leds
   
   'Array auslesen und in Datei eintragen
   For x = 1 To Spalten
      For y = 1 To Leds
         'Arrayinhalt zu Datei hinzufügen
         'Wenn farbig, ..
         If NGHeader.m_bRGB.Data = 1 Then
            '..dann drei Werte hintereinander schreiben
            'Rot-Wert
            Datei = Datei & Array_Red(x, y)
            'Grün-Wert
            Datei = Datei & Array_Green(x, y)
            'Blau-Wert
            Datei = Datei & Array_Blue(x, y)
         'Wenn S/W, ..
         Else
            '..dann nur S/W-Wert in Datei schreiben
            Datei = Datei & Array_SW(x, y)
         End If
         
         'Progressbar einstellen
         progbar.Value = (x - 1) * Leds + y
      Next y
   Next x
   
   'Progressbar auf 0 stellen
   progbar.Value = 0
   
   'Wenn SpeichernUnter True ist, Dialog auf jeden Fall öffnen
   If SpeichernUnter = True Then
      'Dialog öffnen
      DateiPfad = FileSystem_OpenDialog(cdialog, Filter_NightGraphix, "Save")
   Else
      'Wenn Speicherpfad leer ist
      If DateiPfad = "" Then
         'Dialog öffnen
         DateiPfad = FileSystem_OpenDialog(cdialog, Filter_NightGraphix, "Save")
      End If
   End If
   
   'Prüfen ob Abbrechen gewählt wurde
   If DateiPfad = "" Then Exit Sub

   'DoEvents
   DoEvents
   
   'Freie DateiNummer holen
   FreeFileNumber = FreeFile
   'Datei öffnen
   Open DateiPfad For Output As #FreeFileNumber
   'Datei schreiben
   Print #FreeFileNumber, Datei
   'Datei schließen
   Close #FreeFileNumber
   
   'Dateinamen ändern
   DateiName = FileSystem_Path2File(DateiPfad)
   
   'Den Caption der Form ändern
   frm_nightgraphix.Caption = DateiName & "* - NightGraphix V1.0"
   
   'Den * aus dem Formtitel entfernen
   FileSystem_SavedChanges True
End Sub

Public Sub FileSystem_SavedChanges(tmpWert As Boolean)
'--------------------------------------------------------------------'
'| Prozedur zum Ändern von Variable SavedChanges                    |'
'--------------------------------------------------------------------'

   'Variable ändern
   SavedChanges = tmpWert
   
   'Wenn tmpWert = False ist, einen * an Titel der Form anhängen
   If tmpWert = False Then
      'Schlecht für Wiederverwendbarkeit:
      'Von einem Modul ohne Referenz auf die Form zugreifen
      frm_nightgraphix.Caption = DateiName & " *" & " - NightGraphix V1.0"
   'Wenn tmpWert = True ist, den * aus dem Titel entfernen
   Else
      frm_nightgraphix.Caption = Replace(frm_nightgraphix.Caption, "*", "")
   End If
End Sub

Private Sub FileSystem_DateiAdd(tmpVar As Variant)
'--------------------------------------------------------------------'
'| Prozedur zum Schreiben der Variable Datei                        |'
'--------------------------------------------------------------------'
   'tmpVar in String umwandeln
   tmpVar = CStr(tmpVar)
   'Variable ändern
   Datei = Datei & tmpVar
End Sub

Private Sub FileSystem_CreateDateiHeader()
'--------------------------------------------------------------------'
'| Prozedur zum Erstellen des Datei-Headers                         |'
'--------------------------------------------------------------------'
   'DateiHeader einfügen
   'Name: NG
   FileSystem_DateiAdd NGHeader.m_sName.Data
   'Version
   FileSystem_DateiAdd NGHeader.m_sVersion.Data
   'Spalten
   FileSystem_DateiAdd NGHeader.m_lColumns.Data
   'Zeilen
   FileSystem_DateiAdd NGHeader.m_lRows.Data
   'StandardFarbe
   FileSystem_DateiAdd NGHeader.m_lColor.Data
   '32. LED
   FileSystem_DateiAdd NGHeader.m_bLastLED.Data
   'RGB/SW
   FileSystem_DateiAdd NGHeader.m_bRGB.Data
   'Zusätzliches
   FileSystem_DateiAdd NGHeader.m_sAdditional.Data
End Sub

Public Sub FileSystem_OpenDatei(cdialog As Object, picsource As PictureBox, progbar As Object, Optional ByRef tmpFile As String = "")
'--------------------------------------------------------------------'
'| Prozedur zum Öffnen einer Datei                                  |'
'--------------------------------------------------------------------'
   'Bei einem Fehler zu ErrHandler springen
   'Fehler kann nur bei ungültiger Datei auftreten.
   On Error GoTo ErrHandler
      
   'Wenn keine Datei zu öffnen ist..
   If Len(tmpFile) = 0 Then
      'Dialog öffnen
      DateiPfad = FileSystem_OpenDialog(cdialog, Filter_NightGraphix, "Open")
      'Prüfen ob Abbrechen gewählt wurde
      If DateiPfad = "" Then Exit Sub
   'Wenn eine Datei zu öffnen ist..
   Else
      'DateiPfad ändern
      DateiPfad = tmpFile
   End If
      
   'Freie DateiNummer holen
   FreeFileNumber = FreeFile
   'Datei öffnen
   Open DateiPfad For Input As #FreeFileNumber
   'Datei einlesen
   Datei = Input(LOF(FreeFileNumber), #FreeFileNumber)
   'Datei schließen
   Close #FreeFileNumber
   
   'Wenn Demomodus nicht gestartet ist und Hardware nicht verbunden ist..
   If (Not NGDemoModus) And (Not Connected2Hardware) Then
      'Fileheader initialisieren, um die Headerdaten auszulesen
      Init_FileHeader
      'Headerdaten auslesen/bei leerer Datei abfragen
      FileSystem_InitOpenedFile Datei
      'Demomodus starten
      frm_nightgraphix.NightGraphix_StartDemo
      'Dateinamen anzeigen
      FileSystem_ShowDateiNameSaved
      'Wenn die Datei leer ist, dann Prozedur beenden, da keine LEDs gefüllt werden
      If Len(Datei) = 0 Then Exit Sub
   'Wenn Demomodus gestartet ist oder Hardware verbunden ist..
   Else
      'Headerdaten auslesen/bei leerer Datei abfragen
      FileSystem_InitOpenedFile Datei
      'Wenn Dateiversion nicht mit Hardware oder Demo übereinstimmt,
      'Fehlermeldung anzeigen und Prozedur beenden
      If ((NGHeader.m_bRGB.Data = Not RGBVersion) Or (NGHeader.m_lColumns.Data <> Spalten) Or (NGHeader.m_lRows.Data <> Leds)) And NGDemoModus Then
         'MessageBox anzeigen
         MsgBox "Datei wurde nicht für diese Demoeinstellungen erstellt!", vbCritical & vbOKOnly, "Dateikonflikt"
         'Prozedur beenden
         Exit Sub
      End If
   End If
      
   'ProgressBar einstellen
   progbar.Max = Leds * CLng(Spalten)
   
   'DateiHeader abschneiden
   Datei = Mid(Datei, 30)
   
   'Datei S/W oder RGB ?
   If NGHeader.m_bRGB.Data = 1 Then
      For x = 1 To Spalten
         For y = 1 To Leds * 3 Step 3
            'DoEvents
            DoEvents
            'Arrays anpassen
            'RGB-Arrays
            Array_Red(x, (y + 2) / 3) = Mid(Datei, (x - 1) * Leds * 3 + y, 1)
            Array_Green(x, (y + 2) / 3) = Mid(Datei, (x - 1) * Leds * 3 + y + 1, 1)
            Array_Blue(x, (y + 2) / 3) = Mid(Datei, (x - 1) * Leds * 3 + y + 2, 1)
            'Zellen füllen
            Draw_FillCell x, (y + 2) / 3, picsource, RGB(Array_Red(x, (y + 2) / 3) * 255, Array_Green(x, (y + 2) / 3) * 255, Array_Blue(x, (y + 2) / 3) * 255), False
            
            'ProgressBar einstellen
            progbar.Value = (x - 1) * Leds + y / 3
         Next y
      Next x
   Else
      For x = 1 To Spalten
         For y = 1 To Leds
            'DoEvents
            DoEvents
            'Arrays anpassen
            'SW-Array
            Array_SW(x, y) = Mid(Datei, (x - 1) * Leds + y, 1)
            'Zellen füllen
            Draw_FillCell x, y, picsource, IIf(Array_SW(x, y) = 1, LEDColor, vbWhite), False
            
            'ProgressBar einstellen
            progbar.Value = (x - 1) * Leds + y
         Next y
      Next x
   End If
   
   'Progressbar auf 0 setzen
   progbar.Value = 0
   
   'Dateinamen anzeigen
   FileSystem_ShowDateiNameSaved
   
   'Anzahl der aktiven LEDs und die Laufzeit berechnen
   frm_nightgraphix.NightGraphix_ShowPanelProperties
   
   'Fehlerbehandlung nicht ausführen
   Exit Sub
   
'Fehlerbehandlung
ErrHandler:
   'Modell und Arrays löschen
   Draw_ClearModell
   'MessageBox mit Fehlermeldung anzeigen
   MsgBox DateiPfad & vbCrLf & "kann nicht geöffnet werden.", vbOKOnly & vbCritical, "Achtung - NightGraphix V1.0"
End Sub

Public Function FileSystem_Path2File(tmpPath As String) As String
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben des Dateinamens aus einem Pfad          |'
'--------------------------------------------------------------------'
   'Wird z.B. "C:\Dokumente und Einstellungen\User\Eigene Dateien\test.txt"
   'übergeben, gibt dies Prozedur "test.txt" zurück
   
   'Position auf 1 setzen
   Pos = 1
   
   'Position des letzten Backslashes herausfinden
   Do While InStr(Pos + 1, tmpPath, "\") <> 0
      'Position auf nächsten Backslash setzen
      Pos = InStr(Pos + 1, tmpPath, "\")
      'DoEvents
      DoEvents
   Loop

   'Dateiname zurückgeben
   FileSystem_Path2File = Mid(tmpPath, Pos + 1)
End Function

Public Sub FileSystem_SaveSettings()
'--------------------------------------------------------------------'
'| Prozedur zum Speichern der Einstellungen in der Registry         |'
'--------------------------------------------------------------------'
   'Aktuelle Farbe speichern
   SaveSetting "NG1.0", "Einstellungen", "Color", CStr(LEDColor)
   'Offset speichern
   SaveSetting "NG1.0", "Einstellungen", "Offset", CStr(Offset)
   'NGRotorSize speichern
   SaveSetting "NG1.0", "Einstellungen", "Rotorgröße", CStr(NGRotorSize)
   'SWColorIndex speichern
   SaveSetting "NG1.0", "Einstellungen", "SWColorIndex", CStr(SWColorIndex)
   'RGBColorIndex speichern
   SaveSetting "NG1.0", "Einstellungen", "RGBColorIndex", CStr(RGBColorIndex)
   'Links/Rechtsdrehend speichern
   SaveSetting "NG1.0", "Einstellungen", "NGRotationSystemLeft", CStr(NGRotationSystemLeft)
   'LiPo-Schwellwert speichern
   SaveSetting "NG1.0", "Einstellungen", "LiPoTreshold", CStr(LiPoTreshold)
   'Animationsrate speichern
   SaveSetting "NG1.0", "Einstellungen", "NGAnimationRate", CStr(NGAnimationRate)
   'Animationframes speichern
   SaveSetting "NG1.0", "Einstellungen", "NGAnimationFrames", CStr(NGAnimationFrames)
   'Demo-Größe speichern
   SaveSetting "NG1.0", "Einstellungen", "NGDemoSize", CStr(NGDemoSize)
   'LiPo-Kapazität speichern
   SaveSetting "NG1.0", "Einstellungen", "NGLiPoCapacity", CStr(NGLiPoCapacity)
   'Schleife durch alle Rotorgrößen im Rotorgrößenarray
   For i = 0 To UBound(NGRotorSizeArray)
      'Eintrag des Rotorgrößenarray speichern
      SaveSetting "NG1.0", "Einstellungen", "NGRotorSizeArray" & CStr(i), CStr(NGRotorSizeArray(i))
   Next i
   
   'Schleife durch alle möglichen Comports
   For i = 1 To 16
      'Speichern, ob Comport durchsucht werden soll
      SaveSetting "NG1.0", "Einstellungen", "NGComportSearch" & CStr(i), CStr(NGComportSearch(i))
   Next i
   
   'Obere LED-Farbe speichern
   SaveSetting "NG1.0", "Einstellungen", "NGTopLEDColor", CStr(NGTopLEDColor)
   'Untere LED-Farbe speichern
   SaveSetting "NG1.0", "Einstellungen", "NGBottomLEDColor", CStr(NGBottomLEDColor)
   
   'Demoversion speichern
   SaveSetting "NG1.0", "Einstellungen", "NGDemoVersion", CStr(NGDemoVersion)
   'Transparenz beim Import speichern
   SaveSetting "NG1.0", "Einstellungen", "NGTransparentImport", CStr(NGTransparentImport)
   'Transparente Farbe beim Import speichern
   SaveSetting "NG1.0", "Einstellungen", "NGTransparentColor", CStr(NGTransparentColor)
   'Aktuelle Sprache
   SaveSetting "NG1.0", "Einstellungen", "CurrentLanguage", CurrentLanguage
End Sub

Public Sub FileSystem_GetSettings()
'--------------------------------------------------------------------'
'| Prozedur zum Lesen der Einstellungen aus der Registry            |'
'--------------------------------------------------------------------'
   'Bei einem Fehler weitermachen, da die Einstellungen möglicherweise noch nicht gesetzt sind
   On Error Resume Next
   
   'Aktuelle Farbe auslesen
   LEDColor = CLng(GetSetting("NG1.0", "Einstellungen", "Color"))
   'Offset auslesen
   Offset = CInt(GetSetting("NG1.0", "Einstellungen", "Offset"))
   'NGRotorSize auslesen
   NGRotorSize = CInt(GetSetting("NG1.0", "Einstellungen", "Rotorgröße"))
   'SWColorIndex auslesen
   SWColorIndex = GetSetting("NG1.0", "Einstellungen", "SWColorIndex")
   'RGBColorIndex auslesen
   RGBColorIndex = GetSetting("NG1.0", "Einstellungen", "RGBColorIndex")
   'Links/Rechtsdrehend auslesen
   NGRotationSystemLeft = CBool(GetSetting("NG1.0", "Einstellungen", "NGRotationSystemLeft"))
   'LiPo-Schwellwert auslesen
   LiPoTreshold = CInt(GetSetting("NG1.0", "Einstellungen", "LiPoTreshold"))
   'Animationsrate auslesen
   NGAnimationRate = CInt(GetSetting("NG1.0", "Einstellungen", "NGAnimationRate"))
   'AnimationFrames auslesen
   NGAnimationFrames = CInt(GetSetting("NG1.0", "Einstellungen", "NGAnimationFrames"))
   'Demo-Größe auslesen
   NGDemoSize = CInt(GetSetting("NG1.0", "Einstellungen", "NGDemoSize"))
   'LiPo-Kapazität auslesen
   NGLiPoCapacity = CInt(GetSetting("NG1.0", "Einstellungen", "NGLiPoCapacity"))
   'Schleife durch alle Rotorgrößen im Rotorgrößenarray
   For i = 0 To UBound(NGRotorSizeArray)
      'Eintrag des Rotorgrößenarray auslesen
      NGRotorSizeArray(i) = GetSetting("NG1.0", "Einstellungen", "NGRotorSizeArray" & CStr(i))
   Next i
   
   'Schleife durch alle möglichen Comports
   For i = 0 To UBound(NGRotorSizeArray)
      'Auslesen, ob Comport durchsucht werden soll
      NGComportSearch(i) = GetSetting("NG1.0", "Einstellungen", "NGComportSearch" & CStr(i))
   Next i

   
   'Obere LED-Farbe auslesen
   NGTopLEDColor = CInt(GetSetting("NG1.0", "Einstellungen", "NGTopLEDColor"))
   'Untere LED-Farbe auslesen
   NGBottomLEDColor = CInt(GetSetting("NG1.0", "Einstellungen", "NGBottomLEDColor"))

   'DemoVersion auslesen
   NGDemoVersion = CInt(GetSetting("NG1.0", "Einstellungen", "NGDemoVersion"))
   'Wenn RGB nicht verfügbar sein soll, dann NGDemoVersion auf 0 setzen..
   If Not NGRGBAvailable Then NGDemoVersion = 0
   
   'Transparenz beim Import auslesen
   NGTransparentImport = CInt(GetSetting("NG1.0", "Einstellungen", "NGTransparentImport"))
   'Transparente Farbe beim Import auslesen
   NGTransparentColor = CLng(GetSetting("NG1.0", "Einstellungen", "NGTransparentColor"))
   'Aktuelle Sprache auslesen
   CurrentLanguage = GetSetting("NG1.0", "Einstellungen", "CurrentLanguage")
End Sub

Public Sub FileSystem_CreateWScriptReference()
'--------------------------------------------------------------------'
'| Prozedur zum Erzeugen einer Referenz auf die WScript-Shell       |'
'--------------------------------------------------------------------'
   'Referenz erstellen
   Set WScript = CreateObject("WScript.Shell")
End Sub

Public Sub FileSystem_DeleteWScriptReference()
'--------------------------------------------------------------------'
'| Prozedur zum Entfernen der Referenz auf die WScript-Shell        |'
'--------------------------------------------------------------------'
   'Referenz auf Nothing setzen
   Set WScript = Nothing
End Sub

Public Sub FileSystem_CheckRegistry4Extension()
'--------------------------------------------------------------------'
'| Prozedur zum Prüfen der Extensions-Einträge in der Registry      |'
'--------------------------------------------------------------------'
   'Programm-Pfad in Variable speichern (mit %1 für die Datei zum Öffnen)
   RegistryPath = """" & App.Path & "\" & App.Title & ".exe" & """" & " %1"
   'Wenn Eintrag schon existiert, also Dateiendung schon mit Programm verknüpft ist..
   If (CheckIfRunningInIDE = True) Then  '(FileSystem_ReadRegistryKey("HKEY_CLASSES_ROOT\NGfile\shell\open\command\" & RegistryPath) <> "") Or
      'Prozedur beenden
      Exit Sub
   Else
      'Beschreibung
      FileSystem_WriteRegistryKey "HKEY_CLASSES_ROOT\.ng\", "ngfile"
      
      'Eintrag ins Kontextmenü des Explorers
      FileSystem_WriteRegistryKey "HKEY_CLASSES_ROOT\.ng\ShellNew\NullFile", ""
      
      'Dateibezeichnung erstellen
      FileSystem_WriteRegistryKey "HKEY_CLASSES_ROOT\ngfile\", "NG-Datei"
      
      'Icon zuweisen
      FileSystem_WriteRegistryKey "HKEY_CLASSES_ROOT\ngfile\DefaultIcon\", App.Path & "\" & App.Title & ".exe" & ",1"
      
      'Datei soll mit NightGraphix geöffnet werden
      FileSystem_WriteRegistryKey "HKEY_CLASSES_ROOT\ngfile\shell\open\command\", RegistryPath
   End If
End Sub

Public Function FileSystem_ReadRegistryKey(ByRef tmpPath As String) As String
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen eines Keys in der Registry                 |'
'--------------------------------------------------------------------'
   'Bei einem Fehler zu ErrHandler springen
   On Error GoTo ErrHandler
   
   'Key aus Registry auslesen
   FileSystem_ReadRegistryKey = WScript.RegRead(tmpPath)
   
   'Keine Fehlerbehandlung durchführen
   Exit Function

'Fehlerbehandlung
ErrHandler:
   'Leeren String zurückgeben
   FileSystem_ReadRegistryKey = ""
End Function

Public Function FileSystem_WriteRegistryKey(ByRef tmpPath As String, ByRef tmpValue As String, Optional ByRef tmpTyp As String = "REG_SZ") As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Setzen eines Keys in der Registry                   |'
'--------------------------------------------------------------------'
   'Bei einem Fehler zu ErrHandler springen
   On Error GoTo ErrHandler
   
   'Key in Registry schreiben
   WScript.RegWrite tmpPath, tmpValue, tmpTyp
   
   'Key wurde erfolgreich geschrieben
   FileSystem_WriteRegistryKey = True

   'Keine Fehlerbehandlung durchführen
   Exit Function

'Fehlerbehandlung
ErrHandler:
   'Key konnte nicht geschrieben werden
   FileSystem_WriteRegistryKey = False
End Function

Public Function FileSystem_DeleteRegistryKey(ByRef tmpPath As String) As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Löschen eines Keys in der Registry                  |'
'--------------------------------------------------------------------'
   'Bei einem Fehler zu ErrHandler springen
   On Error GoTo ErrHandler
   
   'Key in Registry löschen
   WScript.RegDelete tmpPath
   
   'Key wurde erfolgreich gelöscht
   FileSystem_DeleteRegistryKey = True
   
   'Fehlerbehandlung nicht durchführen
   Exit Function

'Fehlerbehandlung
ErrHandler:
   'Key konnte nicht gelöscht werden
   FileSystem_DeleteRegistryKey = False
End Function

Public Function FileSystem_CreateFileTemplate() As String
'--------------------------------------------------------------------'
'| Prozedur zum Erstellen einer Dateiheader-Vorlage                 |'
'--------------------------------------------------------------------'
   'DateiHeader einfügen
   'Name: NG
   FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & CStr(NGHeader.m_sName.Data)
   'Version
   FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & CStr(NGHeader.m_sVersion.Data)
   'Spalten
   FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & CStr(NGHeader.m_lColumns.Data)
   'Zeilen
   FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & CStr(NGHeader.m_lRows.Data)
   'StandardFarbe
   FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & CStr(NGHeader.m_lColor.Data)
   '32. LED
   FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & CStr(NGHeader.m_bLastLED.Data)
   'RGB/SW
   FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & CStr(NGHeader.m_bRGB.Data)
   'Zusätzliches
   FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & CStr(NGHeader.m_sAdditional.Data)
   
   'Die Größe des Arrays in die Datei schreiben und mit Nullen füllen
   For x = 1 To Spalten
      For y = 1 To Leds
         'Arrayinhalt zu Datei hinzufügen
         'Wenn farbig, ..
         If NGHeader.m_bRGB.Data = 1 Then
            '..dann drei Werte hintereinander schreiben
            'RGB-Werte alle zusammen schreiben
            FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & "000"
         'Wenn S/W, ..
         Else
            '..dann nur S/W-Wert in Datei schreiben
            FileSystem_CreateFileTemplate = FileSystem_CreateFileTemplate & "0"
         End If
      Next y
   Next x
End Function

Public Function FileSystem_ClearPath(ByRef tmpPath As String) As String
'--------------------------------------------------------------------'
'| Prozedur zum Bereinigen eines Pfades (ein "/" am Ende)           |'
'--------------------------------------------------------------------'
   'Ein "\" ans Ende anhängen
   FileSystem_ClearPath = tmpPath & "\"
   'Alle doppelten "\\" durch einfache "\" ersetzen
   FileSystem_ClearPath = Replace(FileSystem_ClearPath, "\\", "\")
End Function

Public Sub FileSystem_GetHeaderData(ByRef tmpDatei As String)
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen der Headerdaten                            |'
'--------------------------------------------------------------------'
   'Name
   NGHeader.m_sName.Data = Mid(tmpDatei, NGHeader.m_sName.DataStart, NGHeader.m_sName.DataLength)
   'Version
   NGHeader.m_sVersion.Data = Mid(tmpDatei, NGHeader.m_sVersion.DataStart, NGHeader.m_sVersion.DataLength)
   'Zeilen (LEDs)
   NGHeader.m_lColumns.Data = Mid(tmpDatei, NGHeader.m_lColumns.DataStart, NGHeader.m_lColumns.DataLength)
   'Spalten
   NGHeader.m_lRows.Data = Mid(tmpDatei, NGHeader.m_lRows.DataStart, NGHeader.m_lRows.DataLength)
   'Farbe
   NGHeader.m_lColor.Data = Mid(tmpDatei, NGHeader.m_lColor.DataStart, NGHeader.m_lColor.DataLength)
   'Letzte LED an ?
   NGHeader.m_bLastLED.Data = Mid(tmpDatei, NGHeader.m_bLastLED.DataStart, NGHeader.m_bLastLED.DataLength)
   'RGB oder SW ?
   NGHeader.m_bRGB.Data = Mid(tmpDatei, NGHeader.m_bRGB.DataStart, NGHeader.m_bRGB.DataLength)
   'Zusätzliches
   NGHeader.m_sAdditional.Data = Mid(tmpDatei, NGHeader.m_sAdditional.DataStart, NGHeader.m_sAdditional.DataLength)
End Sub

Public Sub FileSystem_InitOpenedFile(ByRef tmpDatei As String)
   'Wenn Datei leer ist, dann die Daten für die zu erstellende Datei abfragen
   'Datei ist leer, nachdem sie im Explorer neu erstellt wurde
   If (tmpDatei = "") Then
      'Fenster zum Abfragen der Version laden
      Load frm_choosefileversion
      'Fenster zum Abfragen der Version anzeigen
      frm_choosefileversion.Show vbModal, frm_nightgraphix
      'LED-Anzahl setzen
      Leds = NGNewFileVersion * 8 + 16
      'Spalten-Anzahl setzen
      Spalten = 512
      'RGB-Version setzen
      RGBVersion = False
   'Wenn Datei nicht leer ist..
   Else
      'Headerdaten auslesen
      FileSystem_GetHeaderData Datei
      'LED-Anzahl in Variable speichern
      Leds = CInt(NGHeader.m_lRows.Data)
      'Spaltenanzahl setzen
      Spalten = CInt(NGHeader.m_lColumns.Data)
      'RGB-Version in Variable speichern
      RGBVersion = CBool(NGHeader.m_bRGB.Data)
   End If
End Sub

Public Sub FileSystem_ShowDateiNameSaved()
   'Datei ist gespeichert
   SavedChanges = True
   'Dateiname bestimmen
   DateiName = FileSystem_Path2File(DateiPfad)
   'Beschriftung der Form ändern
   frm_nightgraphix.Caption = DateiName & " - NightGraphix V1.0"
End Sub

Public Function FileSystem_GetFilesInFolder(ByRef tmpPath As String, ByRef tmpFilter As String) As String()
'--------------------------------------------------------------------'
'| Prozedur zum Ermitteln aller Dateien in einem Ordner             |'
'--------------------------------------------------------------------'
   Dim File$, hFile&, FD As WIN32_FIND_DATA, dats&
   'Array mit Dateinamen
   Dim FileNames() As String
   
   'Erste Datei suchen
   hFile = FindFirstFile(FileSystem_ClearPath(tmpPath) & tmpFilter, FD)
   If hFile = 0 Then Exit Function
   Do
      'Dateinamen parsen
      File = Left(FD.cFileName, InStr(FD.cFileName, Chr(0)) - 1)
      'Wenn der Dateiname kein Verzeichnis ist..
      If Not ((FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) Then
         'Größe des FileNames-Arrays anpassen
         ReDim Preserve FileNames(dats)
         'Dateiname in Array schreiben
         FileNames(dats) = File
         'Variable um eins erhöhen
         dats = dats + 1
      End If
   Loop While FindNextFile(hFile, FD)
   'Dateihandle schließen
   Call FindClose(hFile)
   'Dateinamen zurückgeben
   FileSystem_GetFilesInFolder = FileNames
End Function

Public Sub FileSystem_LogEvent(ByRef tmpEvent As String)
'--------------------------------------------------------------------'
'| Prozedur zum Loggen eines Events in eine Textdatei               |'
'--------------------------------------------------------------------'
   'Event an Logfile anhängen
   LogFile = LogFile & tmpEvent & vbCrLf
End Sub

Public Sub FileSystem_SaveLogFile(ByRef tmpFileName As String)
'--------------------------------------------------------------------'
'| Prozedur zum Speichern der Logdatei                              |'
'--------------------------------------------------------------------'
   'Freie DateiNummer holen
   FreeFileNumber = FreeFile
   'Datei öffnen
   Open FileSystem_ClearPath(App.Path) & tmpFileName & ".txt" For Output As #FreeFileNumber
   'Datei schreiben
   Print #FreeFileNumber, LogFile
   'Datei schließen
   Close #FreeFileNumber
End Sub

Public Sub LogEvent(ByRef tmpEvent As String)
   FileSystem_LogEvent tmpEvent
End Sub

