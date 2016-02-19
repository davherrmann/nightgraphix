VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_writereadscreen 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Übertragung - Einstellungen"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "frm_writescreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4095
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.ProgressBar prg_xpstyle 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame fra_frame 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Bildnummer (1,2,3)"
      Height          =   1335
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cbo_picture 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lbl_beschriftung 
         BackColor       =   &H00EFEFEF&
         Caption         =   "Wählen Sie, wo das Bild gespeichert werden soll:"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton cmd_weiter 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_abbrechen 
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "frm_writereadscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Welches Frame ist aktiv ?
Private ActiveFrame As Integer
'Variable für Schleifen etc.
Private i As Long

Private Sub cmd_abbrechen_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf cmd_abbrechen                          |'
'--------------------------------------------------------------------'
   'Form verstecken
   Me.Hide
   'Form schließen
   Unload Me
End Sub

Private Sub cmd_weiter_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf cmd_weiter                             |'
'--------------------------------------------------------------------'
   'Ausgewählte Einträge in Variablen schreiben
   NumberOfPicture = cbo_picture.ListIndex + 1
   'Fenster verstecken
   Me.Hide
   
   'Wenn EEPROM ausgelesen werden soll..
   If ReadOrWrite = ReadEEPROM Then
      'EEPROM-Daten von der Hardware anfordern
      Communication_RequestData frm_nightgraphix.msc_seriell
   'Wenn EEPROM beschrieben werden soll..
   Else
      'Daten an MC senden
      Communication_SendData frm_nightgraphix.msc_seriell
   End If
   
   'Fenster schließen
   Unload Me
End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------'
'| Prozedur beim Laden der Form: Initialisierungen                  |'
'--------------------------------------------------------------------'
   'Icon der Form initialisieren
   Icons_Init Me
   'Beschriftung der Form und des Labels ändern
   'Wenn EEPROM ausgelesen wird..
   If ReadOrWrite = ReadEEPROM Then
      'Beschriftung auf "Auslesen - Einstellungen" ändern
      Me.Caption = LngSpecStrings.ReadOptions
      'Beschriftung des Labels ändern
      lbl_beschriftung(0).Caption = LngSpecStrings.ChooseWhereToSave
   'Wenn EEPROM beschrieben wird..
   Else
      'Beschriftung auf "Beschreiben - Einstellungen" ändern
      Me.Caption = LngSpecStrings.WriteOptions
      'Beschriftung des Labels ändern
      lbl_beschriftung(0).Caption = LngSpecStrings.ChooseWhichToRead
   End If
   
   'Wenn keine Animation gewählt ist..
   If NGAnimationRate = 0 Then
      'cbo_picture initialisieren
      For i = 1 To 3
         'Eintrag hinzufügen
         Init_ComboBox cbo_picture, LngSpecStrings.Picture & " " & CStr(i)
      Next i
      'Beschriftung des Frames ändern
      fra_frame(0).Caption = LngSpecStrings.PictureNumber & " (1,2,3)"
   'Wenn Animation gewählt ist..
   Else
      'cbo_picture initialisieren
      For i = 1 To NGAnimationFrames
         'Eintrag hinzufügen
         Init_ComboBox cbo_picture, LngSpecStrings.Picture & " " & CStr(i)
      Next i
      'Beschriftung des Frames ändern
      fra_frame(0).Caption = LngSpecStrings.PictureNumber & " (1-" & CStr(NGAnimationFrames) & ")"
   End If
   
   'Sprache für alle Controls setzen
   If Not LngInProcess Then Language_SetControlProperties
End Sub
