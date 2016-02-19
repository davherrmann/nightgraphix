VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_optionshardware 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "NG1.0 - Optionen - Hardware"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   Icon            =   "frm_optionshardware.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3495
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.ProgressBar prg_xpstyle 
      Height          =   375
      Left            =   840
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame fra_animationframes 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Anim. - Frames"
      Height          =   1095
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
      Begin VB.HScrollBar hsc_animationframes 
         Height          =   255
         Left            =   120
         Max             =   8
         Min             =   2
         TabIndex        =   10
         Top             =   720
         Value           =   2
         Width           =   1335
      End
      Begin VB.TextBox txt_animationframes 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "3 Frames"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fra_animationsrate 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Animationsrate"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
      Begin VB.TextBox txt_animationsrate 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "0"
         Top             =   360
         Width           =   1335
      End
      Begin VB.HScrollBar hsc_animationsrate 
         Height          =   240
         Left            =   120
         Max             =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame fra_liposchwellwert 
      BackColor       =   &H00EFEFEF&
      Caption         =   "LiPo Schwellwert"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      Begin VB.HScrollBar hsc_liposchwellwert 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         Max             =   420
         SmallChange     =   5
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txt_liposchwellwert 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "6.00 Volt"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmd_abbrechen 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1575
   End
End
Attribute VB_Name = "frm_optionshardware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Variable für Schleifen etc.
Private i As Integer

Private Sub cmd_abbrechen_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Drücken des Buttons "Abbrechen"                    |'
'--------------------------------------------------------------------'
   'Form beenden
   Unload Me
End Sub

Private Sub cmd_ok_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Drücken des Buttons "OK"                           |'
'--------------------------------------------------------------------'
   'Wenn LiPo-Schwellwert, Animationsrate oder Animationframes verändert wurde..
   If (LiPoTreshold <> CInt(Mid(txt_liposchwellwert.Text, 1, InStr(1, txt_liposchwellwert.Text, " Volt") - 1))) Or (NGAnimationRate <> hsc_animationsrate.Value) Or (NGAnimationFrames <> hsc_animationframes.Value) Then
      'Ersetzt durch Communication_SendHardwareData
      '      'Schwellwert an µC senden
      '      Communication_SendLiPoTreshold frm_nightgraphix.msc_seriell

      'LiPo-Schwellwert in Variable speichern
      LiPoTreshold = CInt(Mid(txt_liposchwellwert.Text, 1, InStr(1, txt_liposchwellwert.Text, " Volt") - 1))
      'Animationsrate in Variable speichern
      NGAnimationRate = hsc_animationsrate.Value
      'Wenn "N/A" in der Textbox steht..
      If txt_animationframes.Text = "N/A" Then
         'Eine 2 in Variable speichern
         NGAnimationFrames = 2
      'Wenn AnimationFrames ungleich 0 ist..
      Else
         'AnimationFrames in Variable speichern
         NGAnimationFrames = CInt(Mid(txt_animationframes.Text, 1, 1))
      End If
      'Hardware-Daten an den MC senden
      Communication_SendHardwareData frm_nightgraphix.msc_seriell
   End If
   
   'Daten in der Registry speichern
   FileSystem_SaveSettings
 
   'Form beenden
   Unload Me
End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------'
'| Prozedur beim Laden der Form: Initialisierungen                  |'
'--------------------------------------------------------------------'
   'Icon der Form initialisieren
   Icons_Init Me
'Es wird jetzt nicht nur der Treshold, sondern auch die anderen Daten angefordert
'   'Treshold anfordern
'   Communication_RequestLiPoTreshold frm_nightgraphix.msc_seriell
   'Daten der Hardware anfordern (LiPo-Treshold, Animationsrate, TOP/BOT)
   If Not LngInProcess Then Communication_RequestHardwareData frm_nightgraphix.msc_seriell
   
   'In Schleife warten, bis Hardware-Daten empfangen wurden
   Do Until Not NGWaitForHardwareData
      'Wenn Hardware nicht angeschlossen ist, dann Schleife beenden
      If Not Connected2Hardware Then Exit Do
      'Anderen Events Zeit lassen
      DoEvents
   Loop
   
   'Schieberegler auf LiPo-Schwellwert einstellen
   hsc_liposchwellwert.Value = LiPoTreshold
   'LiPo-Schwellwert anzeigen
   txt_liposchwellwert.Text = Replace(Format(CStr(hsc_liposchwellwert.Value / 100), "0.00 Volt"), ",", ".")
   
   'Wenn empfangener Wert für die Animations-Frames größer als 8 oder kleiner als 2 ist..
   If (NGAnimationFrames > 8) Or (NGAnimationFrames < 2) Then
      'Eine 2 in Variable speichern
      NGAnimationFrames = 2
   End If
     
   'Schieberegler auf AnimationFrames-Wert einstellen
   hsc_animationframes.Value = NGAnimationFrames
   'AnimationFrames anzeigen
   txt_animationframes.Text = Format(CStr(hsc_animationframes.Value), "0 " & LngSpecStrings.Frames)
   
   'Wenn empfangener Wert für die Animations-Rate größer als 240 ist..
   If NGAnimationRate > 240 Then
      'Eine 0 in Variable speichern
      NGAnimationRate = 0
   End If
   
   'Schieberegler auf Animationsraten-Wert einstellen
   hsc_animationsrate.Value = NGAnimationRate
   'Wert formatiert in die Textbox schreiben
   txt_animationsrate.Text = Format(CStr(hsc_animationsrate.Value / 4), "00.00 " & LngSpecStrings.Seconds)
   'Textbox bei Bedarf daktivieren
   hsc_animationsrate_Change
   
   'Sprache für alle Controls setzen
   If Not LngInProcess Then Language_SetControlProperties
End Sub

Private Sub hsc_animationsrate_Change()
'--------------------------------------------------------------------'
'| Prozedur beim Verändern des Wertes der Scrollbar                 |'
'--------------------------------------------------------------------'
   'Wert formatiert in die Textbox schreiben
   txt_animationsrate.Text = Format(CStr(hsc_animationsrate.Value / 4), "00.00 " & LngSpecStrings.Seconds)
   
   'Wert des Schiebers auswerten..
   Select Case hsc_animationsrate.Value
      'Wenn der Wert = 0 ist
      Case 0
         'AnimationsFrames-Frame deaktivieren
         fra_animationframes.Enabled = False
         'AnimationsFrames-Textbox deaktivieren
         txt_animationframes.Enabled = False
         'AnimationsFrames-Schieberegler deaktivieren
         hsc_animationframes.Enabled = False
         'Ein "N/A" in die Textbox schreiben
         txt_animationsrate.Text = "N/A"
         txt_animationframes.Text = "N/A"
      'Wenn der Wert nicht 0 ist
      Case Else
         'AnimationsFrames-Frame aktivieren
         fra_animationframes.Enabled = True
         'AnimationsFrames-Textbox aktivieren
         txt_animationframes.Enabled = True
         'AnimationsFrames-Schieberegler aktivieren
         hsc_animationframes.Enabled = True
         'Wert der Animationsframes in die Textbox schreiben
         txt_animationframes.Text = Format(CStr(hsc_animationframes.Value), "0 " & LngSpecStrings.Frames)
   End Select
End Sub

Private Sub hsc_animationsrate_Scroll()
'--------------------------------------------------------------------'
'| Prozedur beim Verschieben der Scrollbar                          |'
'--------------------------------------------------------------------'
   'Wert formatiert in die Textbox schreiben
   txt_animationsrate.Text = Format(CStr(hsc_animationsrate.Value / 4), "00.0 " & LngSpecStrings.Seconds)
End Sub

Private Sub hsc_animationframes_Change()
'--------------------------------------------------------------------'
'| Prozedur beim Verändern des Wertes der Scrollbar                 |'
'--------------------------------------------------------------------'
   'Wert formatiert in die Textbox schreiben
   txt_animationframes.Text = Format(CStr(hsc_animationframes.Value), "0 " & LngSpecStrings.Frames)
End Sub

Private Sub hsc_animationframes_Scroll()
'--------------------------------------------------------------------'
'| Prozedur beim Verschieben der Scrollbar                          |'
'--------------------------------------------------------------------'
   'Wert formatiert in die Textbox schreiben
   txt_animationframes.Text = Format(CStr(hsc_animationframes.Value), "0 " & LngSpecStrings.Frames)
End Sub

Private Sub hsc_liposchwellwert_Change()
'--------------------------------------------------------------------'
'| Prozedur beim Verändern des Wertes der Scrollbar                 |'
'--------------------------------------------------------------------'
   'Wert formatiert in die Textbox schreiben
   txt_liposchwellwert.Text = Replace(Format(CStr(hsc_liposchwellwert.Value / 100), "0.00 Volt"), ",", ".")
End Sub

Private Sub hsc_liposchwellwert_Scroll()
'--------------------------------------------------------------------'
'| Prozedur beim Verschieben der Scrollbar                          |'
'--------------------------------------------------------------------'
   'Wert formatiert in die Textbox schreiben
   txt_liposchwellwert.Text = Replace(Format(CStr(hsc_liposchwellwert.Value / 100), "0.00 Volt"), ",", ".")
End Sub

Private Sub txt_animationframes_GotFocus()
'--------------------------------------------------------------------'
'| Prozedur beim Setzen des Fokus auf txt_animationframes           |'
'--------------------------------------------------------------------'
   'Fokus auf hsc_animationframes setzen
   hsc_animationframes.SetFocus
End Sub

Private Sub txt_animationsrate_GotFocus()
'--------------------------------------------------------------------'
'| Prozedur beim Setzen des Fokus auf txt_animationsrate            |'
'--------------------------------------------------------------------'
   'Fokus auf hsc_animationsrate setzen
   hsc_animationsrate.SetFocus
End Sub

Private Sub txt_liposchwellwert_GotFocus()
'--------------------------------------------------------------------'
'| Prozedur beim Setzen des Fokus auf txt_liposchwellwert           |'
'--------------------------------------------------------------------'
   'Fokus auf hsc_liposchwellwert setzen
   hsc_liposchwellwert.SetFocus
End Sub


