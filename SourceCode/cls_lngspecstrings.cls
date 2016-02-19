VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_lngspecstrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'_________________________________________________________________________________'
'|                               KLASSE cls_lngspecstrings                       |'
'| Diese Klasse beinhaltet globale Variablen, die durch die .lng-Datei           |'
'| verändert werden können.                                                      |'
'---------------------------------------------------------------------------------'
'Variablen müssen deklariert werden
Option Explicit

'Suche nach Hardware..
Public LookingForHardware As String
'Anschluss ComX
Public PortComX As String
'Keine NGX-Hardware gefunden
Public NotFoundNGXHardware
'Spalte(n)
Public Columns As String
'Sekunden
Public Seconds As String
'Frames
Public Frames As String
'Wählen Sie, wo das Bild gespeichert werden soll
Public ChooseWhereToSave As String
'Wählen Sie, welches Bild ausgelesen werden soll
Public ChooseWhichToRead As String
'Auslesen - Einstellungen
Public ReadOptions As String
'Schreiben - Einstellungen
Public WriteOptions As String
'Bildnummer
Public PictureNumber As String
'Bild
Public Picture As String
'Wollen Sie die Datei noch speichern?
Public AskToSave As String
'Warnung
Public Warning As String
'Achtung
Public Attention As String
'Fehler
Public Error As String
'Ungültige Rotorgröße für dieses System!
Public InvalidRotorSize As String
'Wollen Sie die alte Rotorgröße beibehalten?
Public AskToKeepOldRotorSize As String
'Ungültiger Offset
Public InvalidOffset As String
'Wollen Sie den alten Offset behalten?
Public AskToKeepOldOffset As String
'Wert für Rotorgröße ist zu klein! Der alte Wert wird beibehalten.
Public RotorSizeTooSmall As String
'Fehler beim Ändern der Sprache aufgetreten
Public LanguageError As String

