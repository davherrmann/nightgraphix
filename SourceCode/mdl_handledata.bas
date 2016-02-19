Attribute VB_Name = "mdl_handledata"
'_________________________________________________________________________________'
'|                          MODUL mdl_handledata                                 |'
'| Dieses Modul beinhaltet Prozeduren zum Kommunizieren zum Auswerten der        |'
'| Daten, die vom Microcontroller empfangen wurden.                              |'
'---------------------------------------------------------------------------------'
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Variable für Schleifen etc.
Private y As Integer

Public Function HandleData_Check4HardwareData(ByVal tmpData As String) As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Prüfen, ob Hardwaredaten empfangen wurden           |'
'--------------------------------------------------------------------'
   'Wenn die Daten mit "NG" anfangen und NG noch nicht mit dem MC verbunden ist..
   If (InStr(1, tmpData, "NG") <> 0) And (Not Connected2Hardware) Then
      'Hardwaredaten auswerten
      HandleData_GetHardwareData tmpData
      'Es wurden Hardwaredaten empfangen
      HandleData_Check4HardwareData = True
   'Wenn es andere Daten sind..
   Else
      'Es wurden keine Hardwaredaten empfangen
      HandleData_Check4HardwareData = False
   End If
End Function

Public Function HandleData_Check4RequestData(ByVal tmpData As String) As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Prüfen, ob Hardwaredaten empfangen wurden           |'
'--------------------------------------------------------------------'
   'Wenn die Daten mit "D" anfangen
   If (InStr(1, tmpData, "D") = 1) Then
      'Es wurden Hardwaredaten empfangen
      HandleData_Check4RequestData = True
   'Wenn es andere Daten sind..
   Else
      'Es wurden keine Hardwaredaten empfangen
      HandleData_Check4RequestData = False
   End If
End Function

Private Sub HandleData_GetHardwareData(ByVal tmpData As String)
'--------------------------------------------------------------------'
'| Prozedur zum Auswerten der Hardwaredaten                         |'
'--------------------------------------------------------------------'
   'Daten werden in der Form "NGxxYY" übertragen:
   '  - xx ist die Anzahl der LEDs, von 16 bis 64
   '  - YY ist die Drehrichtung des angeschlossenen Blattes (YY, LI, RE)
   
   'Daten bis zum ersten NG abschneiden
   tmpData = Mid(tmpData, InStr(1, tmpData, "NG"))
   
   'Wenn NG im Demomodus ist..
   If NGDemoModus Then
      'Wenn die LED-Anzahl der Demo nicht mit der Hardware übereinstimmt..
      If Leds <> CInt(Mid(tmpData, 3, 2)) Then
         'Verbinden abbrechen
         CancelConnect = True
         'Prozedur beenden
         Exit Sub
      End If
   End If
   
   'LED-Anzahl auslesen
   Leds = CInt(Mid(tmpData, 3, 2))
   'Spalten-Anzahl setzen
   Spalten = 512
   'Drehrichtung des angeschlossenen Blattes auswerten
   Select Case Mid(tmpData, 5, 2)
      'Wenn das Blatt linksdrehend ist..
      Case "LI"
         NGRotationMCRight = False
      'Wenn das Blatt rechtsdrehend ist..
      Case "RE"
         NGRotationMCRight = True
      'Wenn die Drehrichtung noch nicht gesezt wurde..
      Case "YY"
         'Das Fenster zur Abfrage von TOP/BOT laden
         Load frm_choosetopbottom
         'Fenster anzeigen
         frm_choosetopbottom.Show
      'Wenn das Blatt linksdrehend ist, und RGB-Hardware..
      Case "LC"
         NGRotationMCRight = False
         'RGB-Hardware
         RGBVersion = True
      'Wenn das Blatt rechtsdrehend ist, und RGB-Hardware..
      Case "RC"
         NGRotationMCRight = True
         'RGB-Hardware
         RGBVersion = True
   End Select
   'Farbe einstellen (Standard:Gelb für Oberseite, Blau für Unterseite)
   LEDColor = IIf((Not NGRotationSystemLeft) = (NGRotationMCRight), NGTopLEDColor, NGBottomLEDColor)

'Da noch keine RGB-Hardware vorhanden ist, voerst noch nicht implementiert
'   'Wenn die Hardware RGB ist..
'   If UCase(Mid(tmpData, 5, 3)) = "RGB" Then
'      'Farbversions-Variable setzen
'      RGBVersion = True
'   'Wenn Hardware S/W ist..
'   Else
'      'Farbversions-Variable setzen
'      RGBVersion = False
'   End If
End Sub

Public Sub HandleData_Check4FullData(ByRef tmpPicSource As PictureBox, ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Prüfen der empfangenen Daten                        |'
'--------------------------------------------------------------------'
   'Wenn die Daten vollständig sind..
   If Len(NGReceivedData) >= (Leds \ 8 * (IIf(RGBVersion, 3, 1))) Then
      'Es wird nicht mehr auf Daten gewartet
      NGWaitForData = False
      'Empfangene Daten umdrehen
      NGReceivedData = StrReverse(NGReceivedData)
      
      'Fortschritt anzeigen
      frm_nightgraphix.prg_fortschritt.Value = NGRequestSpalte
      
      'Wenn RGB-Version angeschlossen ist..
      If RGBVersion Then
         'Schleife durch alle Bytes
         For y = 1 To (Leds \ 8) * 3 - 2 Step 3
            'Im Array ein Byte setzen
            Array_SetRGBByte Abs(ReadMirror - NGRequestSpalte), ((y + 2) / 3 - 1) * 8, Asc(Mid(NGReceivedData, y, 1)), Asc(Mid(NGReceivedData, y + 1, 1)), Asc(Mid(NGReceivedData, y + 2, 1)), tmpPicSource
         Next y
      'Wenn SW-Hardware angeschlossen ist..
      Else
         'Schleife durch alle Bytes
         For y = 1 To (Leds \ 8)
            'Im Array ein Byte setzen
            Array_SetByte Abs(ReadMirror - NGRequestSpalte), (y - 1) * 8, Asc(Mid(NGReceivedData, y, 1)), tmpPicSource
         Next y
      End If
      
      'Wenn schon alle Spalten empfangen wurden..
      If NGRequestSpalte >= Spalten Then
         'Progressbar auf 0 stellen
         frm_nightgraphix.prg_fortschritt.Value = 0
         'Bild neu zoomen
         Draw_Zoom tmpPicSource, frm_nightgraphix.pic_target
         'Request-Spalte auf 1 setzen
         NGRequestSpalte = 1
         'Prozedur beenden
         Exit Sub
      End If
      'Request-Spalte um 1 erhöhen
      NGRequestSpalte = NGRequestSpalte + 1
      'Neue Daten anfordern
      Communication_SendRequest tmpMSComm
   End If
End Sub
