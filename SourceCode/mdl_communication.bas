Attribute VB_Name = "mdl_communication"
'_________________________________________________________________________________'
'|                          MODUL mdl_communication                              |'
'| Dieses Modul beinhaltet Prozeduren zum Kommunizieren mit der dem MC           |'
'| über die serielle Schnittstelle.                                              |'
'---------------------------------------------------------------------------------'

Option Explicit

'VARIABLEN..
'###################RGB-Verfügbarkeit umstellen######################'
'Ist die RGB-Demo verfügbar ?
Public Const NGRGBAvailable As Boolean = False
'###################RGB-Verfügbarkeit umstellen######################'

'Welches Bild soll übertragen werden ?
Public NumberOfPicture As Integer
'Soll das Bild auf Ober- oder Unterseite übertragen werden ?
'Unterseite = True, Oberseite = False
Public BottomSide As Boolean
'Rotor rechts- oder linksdrehend ?
'Rechtsdrehen = True, Linksdrehend = False
Public RightRotation As Boolean
'Welcher LiPo-Schwellwert ist gesetzt ?
Public LiPoTreshold As Integer
'Offset (Um wie viele Spalten ist der erste Eintrag verschoben ?)
Public Offset As Integer
'An welchem COM-Port ist die Hardware
Public NGCOMPort As Byte
'Welche Daten wurden von der Hardware empfangen ?
Public NGReceivedData As String
'Wird auf Spaltendaten gewartet ?
Public NGWaitForData As Boolean
'Welche Spalte wird gerade abgefragt ?
Public NGRequestSpalte As Integer
'Wird auf LiPoTreshold-Daten gewartet ?
Public NGWaitForTresholdData As Boolean
'Wird auf Hardwaredaten gewartet ?
Public NGWaitForHardwareData As Boolean
'Wird nach dem Senden der HW-Daten auf ein Acknowledge gewartet
Public NGWaitForHWDataAck As Boolean
'Systemdrehrichtung von NG (Rotordrehrichtung)
Public NGRotationSystemLeft As Boolean
'Drehrichtung des angeschlossenen Blattes
Public NGRotationMCRight As Boolean
'Ist das angeschlossene Blatt oben oder unten ?
Public NGTopSideMC As Boolean
'Animationsrate der Hardware
Public NGAnimationRate As Integer
'Wieviel Frames hat die Animation ?
Public NGAnimationFrames As Integer
'Wird mit Demo gearbeitet ?
Public NGDemoModus As Boolean
'Demogröße
Public NGDemoSize As Integer
'Rotorgröße in mm
Public NGRotorSize As Integer
'Array der Rotorgrößen in mm
Public NGRotorSizeArray(0 To 6) As Long
'Welche Größe soll bei einer neuen Datei benutzt werden ?
Public NGNewFileVersion As Integer
'Welche Kapazität hat der Akku ?
Public NGLiPoCapacity As Integer
'LED-Farbe für die obere Seite
Public NGTopLEDColor
'LED-Farbe für die untere Seite
Public NGBottomLEDColor
'Welche Demoversion - RGB oder SW ?
Public NGDemoVersion As Integer
'Soll der Import transparent sein ?
Public NGTransparentImport As Integer
'Welche Farbe soll transparent sein ?
Public NGTransparentColor As Long
'Welche Comports sollen durchsucht werden ?
Public NGComportSearch(1 To 16) As Boolean
'Variablen für 2-dimensionale Schleifen
Private x As Integer
Private y As Integer
'Ober/Untergrenze für X-Schleife
Private minX As Integer
Private maxX As Integer
'X-Schleife: In welcher Schrittgröße ?
Private stepX As Integer
'Variable für 1-dimensionale Schleife
Private i As Integer

'String zum Senden
Private SendString As String
'Adresse im EEPROM
Public EEPROM_Address As Integer
'High-Adresse im EEPROM
Public EEPROM_Address_High
'Low-Adresse im EEPROm
Public EEPROM_Address_Low

'Ist NG mit dem MicroController verbunden ?
Public Connected2Hardware As Boolean
'Soll das Verbinden mit dem MC abgebrochen werden (Demo - falsche LED-Anzahl)
Public CancelConnect As Boolean

'Aktion: wird der EEPROM beschrieben oder ausgelesen ?
Public ReadOrWrite As tReadOrWrite
'Soll das Ausgelesene gespiegelt werden ?
Public ReadMirror As Integer

'ENUMs..
'Enum für die Aktion: entweder Auslesen oder Beschreiben
Public Enum tReadOrWrite
   'Auslesen
   ReadEEPROM = 0
   'Beschreiben
   WriteEEPROM = 1
End Enum

Public Sub Communication_HandleData(ByVal tmpData As String, ByRef tmpMSComm As MSComm, ByRef tmpPicSource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Auswerten der empfangenen Daten                     |'
'--------------------------------------------------------------------'
   'Wird auf Daten gewartet ?
   If NGWaitForData Then
      'Daten zu vorher empfangenen hinzufügen
      NGReceivedData = NGReceivedData & tmpData
      'Empfangene Daten prüfen
      HandleData_Check4FullData tmpPicSource, tmpMSComm
   'Wird auf ein HW-Daten-Ack vom MC gewartet ?
   ElseIf NGWaitForHWDataAck Then
      'Wenn ein "O" in den Daten enthalten ist..
      If InStr(1, tmpData, "O") > 0 Then
         'Es wird nicht mehr auf ein Ack vom MC gewartet
         NGWaitForHWDataAck = False
      End If
   'Wird auf Hardwaredaten gewartet ?
   ElseIf NGWaitForHardwareData And (tmpData <> "") Then
      'Es wird nicht mehr auf Hardwaredaten gewartet
      NGWaitForHardwareData = False
      
      'Daten kommen im Format "Sxyz"
      ' -x: LiPo-Treshold (0 bis 4,2 Volt, entspricht Chr(0) bis Chr(84))
      ' -y: Animationsrate (0 bis ca 250, pro wieviel Umdrehungen soll das nächste
      '                     Bild gesetzt werden ?, 0 = keine Animation)
      ' -z: Animationframes (2 bis 6, bei wieviel Bilder wieder gestartet werden soll)
      
      'Treshold-Daten in Variable speichern (*5, da Schritte von 0,05)
      'LiPo-Treshold in Volt = 0 bis 255 * 0,05
      LiPoTreshold = Asc(Mid(tmpData, 2, 1)) * 5
            
      'Animationsrate wird ausgelesen und in Variable gespeichert
      NGAnimationRate = Asc(Mid(tmpData, 3, 1))
      
      'Animationsframes wird ausgelesen und in Variable gespeichert
      NGAnimationFrames = Asc(Mid(tmpData, 4, 1))
   
      'Wenn Offsetwert in HW schon gesetzt ist..
      If Asc(Mid(tmpData, 5, 1)) * CLng(256) + Asc(Mid(tmpData, 6, 1)) <= 512 Then 'Mid(tmpData, 5, 2) <> Chr(255) & Chr(255) Then
         'Offset wird ausgelesen und in Variable gespeichert
         Offset = Asc(Mid(tmpData, 5, 1)) * CInt(256) + Asc(Mid(tmpData, 6, 1))
      End If
      
      'Event wieder nach jedem Zeichen auslösen
      tmpMSComm.RThreshold = 1
   'Wird auf Treshold-Daten gewartet ?
   ElseIf NGWaitForTresholdData Then
      'Es wird nicht mehr auf Treshold-Daten gewartet
      NGWaitForTresholdData = False
      
      'Treshold-Daten in Variable speichern
      'LiPoTreshold = Asc(tmpData) * 3
      LiPoTreshold = Asc(Mid(tmpData, 2, 1)) * 3
      'LiPoTreshold = LiPoTreshold * 3

      'Treshold-Daten in Optionen anzeigen
      'frm_optionshardware.txt_liposchwellwert.Text = CStr(LiPoTreshold) & " Volt"
      frm_optionshardware.txt_liposchwellwert.Text = CStr(Mid(LiPoTreshold, 1, 1)) & "." & CStr(Mid(LiPoTreshold, 2, 2)) & " Volt"
  'Wurden die Hardwaredaten gesendet ?
   ElseIf HandleData_Check4HardwareData(tmpData) Then
      'Wenn das Verbinden abgebrochen werden soll, dann Prozedur beenden
      If CancelConnect Then Exit Sub
      'Acknowledge senden
      Communication_SendAcknowledge tmpMSComm
      'COM-Port speichern
      NGCOMPort = tmpMSComm.CommPort
   'Wurden EEPROM-Daten gesendet ?
   ElseIf HandleData_Check4RequestData(tmpData) Then
      'Daten in NGReceivedData schreiben
      NGReceivedData = Mid(tmpData, 2)
      'Es wird auf Daten gewartet
      NGWaitForData = True
      'Empfangene Daten prüfen
      HandleData_Check4FullData tmpPicSource, tmpMSComm
   End If
End Sub

Public Sub Communication_RequestData(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Anfordern der Daten vom MC                          |'
'--------------------------------------------------------------------'
   'Hier werden die Daten ins Array_SW geschrieben
   'Benötigte Variablen:
   'Offset:          Offset         (Integer, Wert von 0 bis Spalten - 1)
   'Drehrichtung:    RightRotation  (Boolean, True bei Rechts-, False bei Linksdrehung)
   'Spalten-Anzahl:  Spalten        (Integer)
   'LED-Anzahl:      Leds           (Integer, Wert modulo 8 = 0)
   'Daten-Array:     Array_SW       (Boolean, Array(1 to Spalten, 1 to Leds))
   
   'Progressbar einstellen
   frm_nightgraphix.prg_fortschritt.Max = Spalten
   
   'Wenn die Hardware rechtsrum dreht
   If NGRotationMCRight Then
      'Das Ausgelesene soll nicht gespiegelt werden
      ReadMirror = 0
   'Wenn die Hardware linksrum dreht
   Else
      'Das Ausgelesene soll gespiegelt werden
      ReadMirror = 513
   End If

   'Request-Spalte auf 1 setzen
   NGRequestSpalte = 1
   'Erstes Request schicken
   Communication_SendRequest tmpMSComm
End Sub

Public Sub Communication_SendData(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Senden der Daten an den MC                          |'
'--------------------------------------------------------------------'
   'Hier werden die Daten aus dem Array_SW gesendet
   'Benötigte Variablen:
   'Offset:          Offset         (Integer, Wert von 0 bis Spalten - 1)
   'Drehrichtung:    RightRotation  (Boolean, True bei Rechts-, False bei Linksdrehung)
   'Spalten-Anzahl:  Spalten        (Integer)
   'LED-Anzahl:      Leds           (Integer, Wert modulo 8 = 0)
   'Daten-Array:     Array_SW       (Boolean, Array(1 to Spalten, 1 to Leds))
   
   'Progressbar einrichten
   frm_nightgraphix.prg_fortschritt.Max = Spalten
   
   'Wenn rechts+oben oder links+unten gedreht wird
   If NGRotationMCRight Then
      'Ober/Untergrenze der Schleife festlegen
      minX = Offset + 1
      maxX = Spalten + Offset
      'Schrittweite: 1
      stepX = 1
   'Wenn rechts+unten oder links+oben gedreht wird
   Else
      'Ober/Untergrenze der Schleife festlegen
      minX = Spalten + Offset
      maxX = Offset + 1
      'Schrittweite: -1
      stepX = -1
   End If
   
   'Hochzähler: i auf 0 stellen
   i = 0
   
   'Schleife durch alle Spalten
   For x = minX To maxX Step stepX
      'Hochzähler: i um 1 erhöhen
      i = i + 1
      
      'Spalte in Variable schreiben
      NGRequestSpalte = (i - 1) Mod Spalten + 1
      'EEPROM-Adressen berechnen
      Maths_EEPROMData
      
      'SendString löschen
      SendString = ""
      
      'Wenn RGB-Hardware angeschlossen ist..
      If RGBVersion Then
         'Schleife durch alle x Bytes
         For y = 1 To Leds \ 8
            'SendString erzeugen
            SendString = SendString & Array_GetRGBByte(((x - 1) Mod Spalten) + 1, (y - 1) * 8 + 1)
         Next y
      'Wenn SW-Hardware angeschlossen ist..
      Else
         'Schleife durch alle x Bytes
         For y = 1 To Leds \ 8
            'SendString erzeugen
            SendString = SendString & Chr(Array_GetByte(((x - 1) Mod Spalten) + 1, (y - 1) * 8 + 1))
         Next y
      End If
      
      'Sendstring umdrehen
      SendString = StrReverse(SendString)
      
      On Error Resume Next
      'Daten an MC senden
      tmpMSComm.Output = "W:" & Chr(EEPROM_Address_High) & Chr(EEPROM_Address_Low) & SendString & ";"
      On Error GoTo 0
      
      'Progressbar einstellen
      frm_nightgraphix.prg_fortschritt.Value = i
   Next x
      
   'Progressbar zurücksetzen
   frm_nightgraphix.prg_fortschritt.Value = 0
End Sub

Public Sub Communication_SendAcknowledge(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Senden eines "P" an den MC                          |'
'--------------------------------------------------------------------'
   'Ein "P" über die serielle Schnittstelle senden
   tmpMSComm.Output = "P"
   'NG ist mit Hardware verbunden
   Connected2Hardware = True
   'Hardwaredaten anfordern
   Communication_RequestHardwareData tmpMSComm
End Sub

Public Sub Communication_SendQuit(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Senden eines "X;" an den MC                         |'
'--------------------------------------------------------------------'
   'Hardwaredaten schreiben
   Communication_SendHardwareData tmpMSComm
   'Warten, bis die FW den Befehl bearbeitet hat und die HW-Daten gesendet hat..
   Do While NGWaitForHWDataAck
      'Anderen Events Zeit lassen
      DoEvents
   Loop
   'Ein "X:;" über die serielle Schnittstelle senden
   tmpMSComm.Output = "X:;"
   'NG ist nicht mit Hardware verbunden
   Connected2Hardware = False
End Sub

Public Sub Communication_SendRequest(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Anfordern der Daten vom MC                          |'
'--------------------------------------------------------------------'
   'Datenspeicher löschen
   NGReceivedData = ""
   
   'EEPROM-Daten errechnen
   Maths_EEPROMData
   
   On Error Resume Next
   'Daten von Hardware anfordern
   tmpMSComm.Output = "R:" & Chr(EEPROM_Address_High) & Chr(EEPROM_Address_Low) & ";"
End Sub

Public Sub Communication_SendLiPoTreshold(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Senden des LiPo-Schwellwertes an den µC             |'
'--------------------------------------------------------------------'
'On Error Resume Next
'   'Ein "F:x;" über die serielle Schnittstelle senden
'   tmpMSComm.Output = "F:" & Chr(LiPoTreshold) & ";"

   'Aufruf an Communication_SendHardwareData weiterleiten
   Communication_SendHardwareData tmpMSComm
End Sub

Public Sub Communication_RequestLiPoTreshold(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Anfordern des LiPo-Schwellwertes vom µC             |'
'--------------------------------------------------------------------'
'On Error Resume Next
'   'Es wird auf Treshold-Daten gewartet
'   NGWaitForTresholdData = True
'   'Ein "E:;" über die serielle Schnittstelle senden
'   tmpMSComm.Output = "E:;"

   'Aufruf an Communication_RequestHardwareData weiterleiten
   Communication_RequestHardwareData tmpMSComm
End Sub

Public Sub Communication_SendHardwareData(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Senden der Hardwaredaten an den MC                  |'
'--------------------------------------------------------------------'
On Error Resume Next
   'Ein "F:xyz;" über die serielle Schnittstelle senden
   '  - x: LiPo Treshold (Chr(0) bis Chr(255))
   '  - y: Animationsrate (Chr(0) bis Chr(255))
   '  - z: Animationsframes (Chr(0) bis Chr(255))
   
   'Daten senden
   tmpMSComm.Output = "F:" & Chr(LiPoTreshold / 5) & Chr(NGAnimationRate) & Chr(NGAnimationFrames) & Chr(Offset \ 256) & Chr(Offset Mod 256) & ";"
   'Es wird auf ein Acknowledge vom MC gewartet
   NGWaitForHWDataAck = True
   'Ereignis nach jedem Zeichen auslösen
   tmpMSComm.RThreshold = 1
End Sub

Public Sub Communication_RequestHardwareData(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Anfordern der Hardwaredaten vom µC                  |'
'--------------------------------------------------------------------'
On Error Resume Next
   'Erst nach vier Zeichen ein Event auslösen
   tmpMSComm.RThreshold = 6
   'Es wird auf Hardwaredaten gewartet
   NGWaitForHardwareData = True
   'Ein "E:;" über die serielle Schnittstelle senden
   tmpMSComm.Output = "E:;"
End Sub
