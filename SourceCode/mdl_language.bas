Attribute VB_Name = "mdl_language"
'_________________________________________________________________________________'
'|                               MODUL mdl_language                              |'
'| Dieses Modul beinhaltet Prozeduren vom Benutzen mehrere Sprachen              |'
'|                                                                               |'
'---------------------------------------------------------------------------------'
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Variable für Schleifen etc.
Private i As Long
Private n As Long
'Welches ist die aktuelle Sprache ?
Public CurrentLanguage As String
'Läuft die Verarbeitung von Language-Dateien gerade ?
Public LngInProcess As Boolean
'Freie FileNumber
Private FreeFileNumber As Long
'Wo sind die Language-Dateien gespeichert ?
Public LngFilePath As String
'Erste eingelesene Zeile
Private FirstLine As String
'Aktuell eingelesene Zeile
Private CurrentLine As String
'Aktuelle Zeilennummer
Private CurrentLineNumber As Long
'Aktuell eingelesene Zeile als Split-Array
Private CurrentSplitLine() As String
'Aktuell eingelesenes Control (Teil vor dem "=")
Private CurrentControl As String
'Aktuell eingelesener Value (Teil hinter dem "=")
Private CurrentValue As String
'Aktuell eingelesener Value als Split-Array
Private CurrentSplitControl() As String
'Referenz auf aktuelles Fenster
Private CurrentWindow As Form
'Referenz auf aktuelles Control
Private CurrentWindowControl As Variant
'Referenz auf den Typ tAllControls
Private AllControls() As tAllControls
'Referenz auf den Typ tLngSpecStrings
Public LngSpecStrings As New cls_lngspecstrings

'TYPEs..
'Type für alle Controls aus der Lng-Datei
Private Type tAllControls
   'Nur für Control-Array: Index
   Index As Integer
   'Name des Controls
   Name As String
   'Container des Controls
   Container As String
   'Eigenschaft, die gesetzt werden soll
   PropertyName As String
   'Eigenschaftswert
   PropertyValue As String
   'Für Property-Arrays
   PropertyIndex As Integer
End Type

Public Sub Language_SetFilePath(ByRef tmpLngFilePath As String)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen des Pfades mit den Language-Dateien          |'
'--------------------------------------------------------------------'
   'Pfad bereinigt in Variable speichern
   LngFilePath = FileSystem_ClearPath(tmpLngFilePath)
End Sub

Public Sub Language_ReadFromFile(ByRef tmpLngFile As String)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen einer anderen Sprache                        |'
'--------------------------------------------------------------------'
   'Bei einem Fehler zur Fehlerbehandlung springen
   On Error GoTo NextLine
   
   'Aktuelle Zeilennummer auf 0 setzen
   CurrentLineNumber = 0
   'Wenn Dateiname leer ist, Prozedur beenden
   If Len(tmpLngFile) = 0 Then Exit Sub
   'Freie DateiNummer holen
   FreeFileNumber = FreeFile
   'Datei öffnen
   Open LngFilePath & tmpLngFile For Input As #FreeFileNumber
   'Aktuelle Sprache in Variable speichern
   CurrentLanguage = tmpLngFile
   'Erste Zeile einlesen
   Line Input #FreeFileNumber, FirstLine
   'Wenn Datei keine gültige Language-Datei ist, in Fehlerbehandlung springen
   If (Left(FirstLine, 16) <> "<NG-LanguageFile") And (Right(FirstLine, 1) <> ">") Then GoTo ErrHandler
   'Schleife durch alle Zeilen der geöffneten Datei
   Do While Not EOF(FreeFileNumber)
      'Aktuelle Zeile einlesen und in Variable speichern
      Line Input #FreeFileNumber, CurrentLine
      'Wenn in der Zeile nichts steht, in die nächste Zeile springen
      If Len(CurrentLine) = 0 Then GoTo NextLine
      'Wenn die Zeile kein Kommentar ist
      If Left(CurrentLine, 2) <> "//" Then
         'Aktuelle Zeile vor und nach dem "=" splitten
         CurrentSplitLine = Split(CurrentLine, "=", 2)
         'Aktuelles Control einlesen
         CurrentControl = Trim(CurrentSplitLine(0))
         'Aktuellen Value einlesen
         CurrentValue = Trim(CurrentSplitLine(1))
         CurrentValue = Mid(CurrentValue, 2, Len(CurrentValue) - 2)
         'Aktuellen Value nach Punkten splitten
         CurrentSplitControl = Split(CurrentControl, ".")
         'Wenn in CurrentSplitValue weniger als 3 Werte stehen..
         If UBound(CurrentSplitControl) < 2 Then
            'Wenn eine globale Variable verändert werden soll..
            If LCase(CurrentSplitControl(0)) = "lngspecstrings" Then
               'Auf diese Form bringen: global.lngspecstrings.Eigenschaft
               CurrentSplitControl = Split("global.lngspecstrings." & CurrentSplitControl(1), ".")
            'Wenn die Eigenschaft einer Form verändert werden soll..
            'Form.this.Eigenschaft = ...
            Else
               'Auf diese Form bringen: Form.this.Eigenschaft = ...
               CurrentSplitControl = Split(CurrentSplitControl(0) & ".this." & CurrentSplitControl(1), ".")
            End If
         'Wenn in CurrentSplitValue 4 Werte stehen
         ElseIf UBound(CurrentSplitControl) = 3 Then
            'Die zwei mittleren Split-Strings zusammenfassen
            CurrentSplitControl(1) = CurrentSplitControl(1) & "." & CurrentSplitControl(2)
            'Den letzen Split-String in den vorletzten verschieben
            CurrentSplitControl(2) = CurrentSplitControl(3)
         End If
         'AllControls-Array neu dimensionieren
         ReDim Preserve AllControls(CurrentLineNumber)
         'Container in Variable speichern
         AllControls(CurrentLineNumber).Container = CurrentSplitControl(0)
         'Wenn das Control in einem Control-Array ist..
         If Right(CurrentSplitControl(1), 1) = ")" Then
            'Control-Name in Variable speichern
            AllControls(CurrentLineNumber).Name = Mid(CurrentSplitControl(1), 1, Len(CurrentSplitControl(1)) - 3)
            'Control-Index in Variable speichern
            AllControls(CurrentLineNumber).Index = Val(Mid(CurrentSplitControl(1), InStrRev(CurrentSplitControl(1), "(") + 1, 3))
         'Wenn das Control nicht in einem Array ist..
         Else
            'Control-Name in Variable speichern
            AllControls(CurrentLineNumber).Name = CurrentSplitControl(1)
            'Control-Index auf -1 setzen
            AllControls(CurrentLineNumber).Index = -1
         End If
         'Wenn das Property in einem Array ist..
         If Right(CurrentSplitControl(2), 1) = ")" Then
            'PropertyName in Variable speichern
            AllControls(CurrentLineNumber).PropertyName = Mid(CurrentSplitControl(2), 1, Len(CurrentSplitControl(2)) - 3)
            'Control-Index in Variable speichern
            AllControls(CurrentLineNumber).PropertyIndex = Val(Mid(CurrentSplitControl(2), InStrRev(CurrentSplitControl(2), "(") + 1, 3))
         'Wenn das Property nicht in einem Array ist..
         Else
            'Control-Name in Variable speichern
            AllControls(CurrentLineNumber).PropertyName = CurrentSplitControl(2)
            'Control-Index auf -1 setzen
            AllControls(CurrentLineNumber).PropertyIndex = -1
         End If
         'PropertyValue in Variable speichern
         AllControls(CurrentLineNumber).PropertyValue = CurrentValue
         'Aktuelle Zeilennummer um eins erhöhen
         CurrentLineNumber = CurrentLineNumber + 1
      End If
'Nächste Zeile
NextLine:
   Loop
   'Datei schließen
   Close #FreeFileNumber
   
   'Fehlerbehandlung nicht ausführen
   Exit Sub
   
'Fehlerbehandlung
ErrHandler:
   'Datei schließen
   Close #FreeFileNumber
   'MessageBox anzeigen
   MsgBox LngSpecStrings.LanguageError
End Sub

Public Sub Language_SetControlProperties()
'--------------------------------------------------------------------'
'| Prozedur zum Setzen der Eigenschaften aller Controls             |'
'--------------------------------------------------------------------'
   'Bei einem Fehler zur Fehlerbehandlung springen
   On Error GoTo NextControl
   
   'Es werden Language-Daten verarbeitet
   LngInProcess = True
   
   'Schleife durch alle Einträge im AllControls-Array
   For i = 0 To UBound(AllControls)
      'Referenz auf Control löschen
      Set CurrentWindowControl = Nothing
      'In welchem Container befindet sich das Control ?
      Select Case LCase(AllControls(i).Container)
         'Hauptfenster
         Case "frm_nightgraphix"
            'Referenz auf Fenster setzen
            Set CurrentWindow = frm_nightgraphix
         'Softwareoptionen-Fenster
         Case "frm_optionssoftware"
            'Referenz auf Fenster setzen
            Set CurrentWindow = frm_optionssoftware
         'Hardwareoptionen-Fenster
         Case "frm_optionshardware"
            'Referenz auf Fenster setzen
            Set CurrentWindow = frm_optionshardware
         'WriteRead-Fenster
         Case "frm_writereadscreen"
            'Referenz auf Fenster setzen
            Set CurrentWindow = frm_writereadscreen
         'Fenster zur Versionsauswahl
         Case "frm_choosefileversion"
            'Referenz auf Fenster setzen
            Set CurrentWindow = frm_choosefileversion
         'About-Fenster
         Case "frm_about"
            Set CurrentWindow = frm_about
         'Wenn eine globale Variable verändert werden soll..
         Case "global"
            'Nichts tun, da dieser Fall später behandelt wird
         'Anderes Fenster nicht vorhanden
         Case Else
            'Nächstes Control behandeln
            GoTo NextControl
      End Select
      
      'Wenn die Toolbar verändert werden soll..
      If Split(AllControls(i).Name, ".")(0) = "tlb_toolbar" Then
         'Referenz auf Toolbar setzen
         Set CurrentWindowControl = frm_nightgraphix.tlb_toolbar.Buttons(AllControls(i).Index)
      'Wenn eine globale Variable verändert werden soll..
      ElseIf AllControls(i).Container = "global" Then
         'Globale Variable über CallByName ändern
         CallByName LngSpecStrings, AllControls(i).PropertyName, VbLet, AllControls(i).PropertyValue
         'Mit nächstem Control weitermachen
         GoTo NextControl
      'Wenn die Eigenschaft der Form verändert werden soll..
      ElseIf AllControls(i).Name = "this" Then
         'Referenz auf aktuelle Form setzen
         Set CurrentWindowControl = CurrentWindow
      'Wenn das Control nicht in einem Array ist..
      ElseIf AllControls(i).Index = -1 Then
         'Referenz auf Control setzen
         Set CurrentWindowControl = CurrentWindow.Controls(AllControls(i).Name)
      'Wenn das Control in einem Array ist..
      Else
         'Referenz auf Control setzen
         Set CurrentWindowControl = CurrentWindow.Controls(AllControls(i).Name).Item(AllControls(i).Index)
      End If
      
      'Welche Property soll verändert werden ?
      Select Case LCase(AllControls(i).PropertyName)
         'Caption
         Case "caption"
            'Caption des Controls setzen
            CurrentWindowControl.Caption = AllControls(i).PropertyValue
         'Text
         Case "text"
            'Text des Controls setzen
            CurrentWindowControl.Text = AllControls(i).PropertyValue
         'ToolTip
         Case "tooltiptext"
            'ToolTipText des Controls setzen
            CurrentWindowControl.ToolTipText = AllControls(i).PropertyValue
         'TabCaption
         Case "tabcaption"
            'TabCaption des Controls setzen
            CurrentWindowControl.TabCaption(AllControls(i).PropertyIndex) = AllControls(i).PropertyValue
      End Select
      
'Nächstes Control
NextControl:
   Next i
   
'   'Bei einem Fehler zu NextLoop springen
'   On Error GoTo NextLoop
'   'Variable n auf 1 setzen
'   n = 1
'   'Solange n wahr ist
'   While (n)
'      'Schleife durch alle Fenster
'      For i = 0 To Forms.Count - 1
'         'Wenn die Form nicht das Hauptfenster oder die Softwareoptionen sind..
'         If (Forms(i).Name <> "frm_nightgraphix") And (Forms(i).Name <> "frm_optionssoftware") Then
'            'Form entladen
'            Unload Forms(i)
'         End If
'      Next i
'      'Aus Endlosschleife rausspringen
'      n = 0
''Nächster Durchgang
'NextLoop:
'      'Anderen Events Zeit lassen
'      DoEvents
'   Wend
   
   'Es wird nicht mehr an Language-Daten gearbeitet
   LngInProcess = False
   
   'Fehlerbehandlung nicht ausführen
   Exit Sub
   
'Fehlerbehandlung
ErrHandler:
   'Es wird nicht mehr an Language-Daten gearbeitet
   LngInProcess = False
   'MessageBox anzeigen
   MsgBox LngSpecStrings.LanguageError
End Sub

