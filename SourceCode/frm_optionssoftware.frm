VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_optionssoftware 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "NG1.0 - Optionen"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "frm_optionssoftware.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   4935
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.ProgressBar prg_xpstyle 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmd_abbrechen 
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3360
      Width           =   1455
   End
   Begin TabDlg.SSTab tab_optionen 
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "System"
      TabPicture(0)   =   "frm_optionssoftware.frx":6852
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fra_rotorgröße"
      Tab(0).Control(1)=   "cmd_saverotorsize"
      Tab(0).Control(2)=   "fra_demogröße"
      Tab(0).Control(3)=   "lbl_beschriftung(1)"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "SW"
      TabPicture(1)   =   "frm_optionssoftware.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_drehrichtung"
      Tab(1).Control(1)=   "fra_offset"
      Tab(1).Control(2)=   "fra_lipocapacity"
      Tab(1).Control(3)=   "lbl_beschriftung(2)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Verbindung"
      TabPicture(2)   =   "frm_optionssoftware.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lst_comports"
      Tab(2).Control(1)=   "lbl_beschriftung(3)"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Sprache"
      TabPicture(3)   =   "frm_optionssoftware.frx":68A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lst_languages"
      Tab(3).Control(1)=   "lbl_beschriftung(4)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "HW"
      TabPicture(4)   =   "frm_optionssoftware.frx":68C2
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "lbl_beschriftung(5)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fra_animationframes"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "fra_animationsrate"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "fra_liposchwellwert"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "RGB"
      TabPicture(5)   =   "frm_optionssoftware.frx":68DE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fra_import"
      Tab(5).Control(1)=   "fra_demoversion"
      Tab(5).Control(2)=   "lbl_beschriftung(6)"
      Tab(5).ControlCount=   3
      Begin VB.Frame fra_liposchwellwert 
         Caption         =   "LiPo Schwellwert"
         Height          =   1095
         Left            =   120
         TabIndex        =   53
         Top             =   780
         Width           =   1455
         Begin VB.TextBox txt_liposchwellwert 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   55
            Text            =   "6.00 Volt"
            Top             =   360
            Width           =   1215
         End
         Begin VB.HScrollBar hsc_liposchwellwert 
            Height          =   255
            LargeChange     =   5
            Left            =   120
            Max             =   420
            SmallChange     =   5
            TabIndex        =   54
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame fra_animationsrate 
         Caption         =   "Animationsrate"
         Height          =   1095
         Left            =   3240
         TabIndex        =   50
         Top             =   780
         Width           =   1575
         Begin VB.HScrollBar hsc_animationsrate 
            Height          =   240
            Left            =   120
            Max             =   240
            TabIndex        =   52
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txt_animationsrate 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   51
            Text            =   "0"
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fra_animationframes 
         Caption         =   "Anim. - Frames"
         Height          =   1095
         Left            =   3240
         TabIndex        =   47
         Top             =   1920
         Width           =   1575
         Begin VB.TextBox txt_animationframes 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   49
            Text            =   "3 Frames"
            Top             =   360
            Width           =   1335
         End
         Begin VB.HScrollBar hsc_animationframes 
            Height          =   255
            Left            =   120
            Max             =   8
            Min             =   2
            TabIndex        =   48
            Top             =   720
            Value           =   2
            Width           =   1335
         End
      End
      Begin VB.Frame fra_import 
         Caption         =   "Bildimport"
         Height          =   1455
         Left            =   -71760
         TabIndex        =   35
         Top             =   780
         Width           =   1575
         Begin VB.OptionButton opt_farbe 
            BackColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   44
            Top             =   720
            Width           =   255
         End
         Begin VB.OptionButton opt_farbe 
            BackColor       =   &H0000FFFF&
            Height          =   255
            Index           =   1
            Left            =   480
            Style           =   1  'Grafisch
            TabIndex        =   43
            Top             =   720
            Width           =   255
         End
         Begin VB.OptionButton opt_farbe 
            BackColor       =   &H0000FF00&
            Height          =   255
            Index           =   2
            Left            =   840
            Style           =   1  'Grafisch
            TabIndex        =   42
            Top             =   720
            Width           =   255
         End
         Begin VB.OptionButton opt_farbe 
            BackColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   1200
            Style           =   1  'Grafisch
            TabIndex        =   41
            Top             =   720
            Width           =   255
         End
         Begin VB.OptionButton opt_farbe 
            BackColor       =   &H00FF00FF&
            Height          =   255
            Index           =   4
            Left            =   120
            Style           =   1  'Grafisch
            TabIndex        =   40
            Top             =   1080
            Width           =   255
         End
         Begin VB.OptionButton opt_farbe 
            BackColor       =   &H00FFFF00&
            Height          =   255
            Index           =   5
            Left            =   465
            Style           =   1  'Grafisch
            TabIndex        =   39
            Top             =   1080
            Width           =   255
         End
         Begin VB.OptionButton opt_farbe 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   6
            Left            =   840
            Style           =   1  'Grafisch
            TabIndex        =   38
            Top             =   1080
            Width           =   255
         End
         Begin VB.OptionButton opt_farbe 
            BackColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   1200
            Style           =   1  'Grafisch
            TabIndex        =   37
            Top             =   1080
            Width           =   255
         End
         Begin VB.CheckBox chk_transparentimport 
            Caption         =   "Transparenz"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbl_beschriftung 
            Caption         =   "Transp. Farbe:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame fra_demoversion 
         Caption         =   "Demo-Version"
         Height          =   855
         Left            =   -74880
         TabIndex        =   31
         Top             =   780
         Width           =   1455
         Begin VB.PictureBox frapic_ledfarbe 
            Appearance      =   0  '2D
            BackColor       =   &H00EFEFEF&
            BorderStyle     =   0  'Kein
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   1215
            TabIndex        =   32
            Top             =   240
            Width           =   1215
            Begin VB.OptionButton opt_demoversion 
               Caption         =   "SW-Version"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   34
               Top             =   0
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton opt_demoversion 
               Caption         =   "RGB-Version"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   33
               Top             =   240
               Width           =   1215
            End
         End
      End
      Begin VB.ListBox lst_languages 
         Height          =   1620
         Left            =   -74760
         TabIndex        =   30
         Top             =   1200
         Width           =   4455
      End
      Begin VB.ListBox lst_comports 
         Height          =   1635
         Left            =   -74760
         Style           =   1  'Kontrollkästchen
         TabIndex        =   28
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Frame fra_drehrichtung 
         Caption         =   "Drehrichtung"
         Height          =   855
         Left            =   -74880
         TabIndex        =   21
         Top             =   780
         Width           =   1455
         Begin VB.PictureBox frapic_drehrichtung 
            Appearance      =   0  '2D
            BorderStyle     =   0  'Kein
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   120
            ScaleHeight     =   495
            ScaleWidth      =   1215
            TabIndex        =   22
            Top             =   240
            Width           =   1215
            Begin VB.OptionButton opt_richtung 
               Caption         =   "Rechts"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton opt_richtung 
               Caption         =   "Links"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   23
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin VB.Frame fra_offset 
         Caption         =   "Offset - Spalte"
         Height          =   855
         Left            =   -74880
         TabIndex        =   19
         Top             =   1740
         Width           =   1455
         Begin VB.TextBox txt_offset 
            Height          =   285
            Left            =   120
            TabIndex        =   20
            Text            =   "0 Spalten"
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame fra_lipocapacity 
         Caption         =   "LiPo Kapazität"
         Height          =   1095
         Left            =   -71640
         TabIndex        =   16
         Top             =   780
         Width           =   1455
         Begin VB.TextBox txt_lipocapacity 
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            MousePointer    =   1  'Pfeil
            TabIndex        =   18
            Text            =   "100 mAh"
            Top             =   360
            Width           =   1215
         End
         Begin VB.HScrollBar hsc_lipocapacity 
            Height          =   255
            LargeChange     =   5
            Left            =   120
            Max             =   500
            Min             =   100
            SmallChange     =   5
            TabIndex        =   17
            Top             =   720
            Value           =   100
            Width           =   1215
         End
      End
      Begin VB.Frame fra_rotorgröße 
         Caption         =   "Rotorgröße"
         Height          =   855
         Left            =   -71640
         TabIndex        =   14
         Top             =   780
         Width           =   1455
         Begin VB.TextBox txt_rotorgröße 
            Height          =   290
            Left            =   120
            TabIndex        =   15
            Text            =   "1300 mm"
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmd_saverotorsize 
         Caption         =   "&Rotorgröße übernehmen"
         Height          =   495
         Left            =   -71640
         TabIndex        =   13
         Top             =   1740
         Width           =   1455
      End
      Begin VB.Frame fra_demogröße 
         Caption         =   "LED-Anzahl"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   4
         Top             =   780
         Width           =   1455
         Begin VB.PictureBox frapic_demogröße 
            Appearance      =   0  '2D
            BorderStyle     =   0  'Kein
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   120
            ScaleHeight     =   1695
            ScaleWidth      =   1215
            TabIndex        =   5
            Top             =   240
            Width           =   1215
            Begin VB.OptionButton opt_demogröße 
               Caption         =   "16 LEDs"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton opt_demogröße 
               Caption         =   "64 LEDs"
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   11
               Top             =   1440
               Width           =   975
            End
            Begin VB.OptionButton opt_demogröße 
               Caption         =   "56 LEDs"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   10
               Top             =   1200
               Width           =   975
            End
            Begin VB.OptionButton opt_demogröße 
               Caption         =   "48 LEDs"
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   9
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton opt_demogröße 
               Caption         =   "40 LEDs"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   8
               Top             =   720
               Width           =   975
            End
            Begin VB.OptionButton opt_demogröße 
               Caption         =   "32 LEDs"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   7
               Top             =   480
               Width           =   975
            End
            Begin VB.OptionButton opt_demogröße 
               Caption         =   "24 LEDs"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   6
               Top             =   240
               Width           =   975
            End
         End
      End
      Begin VB.Label lbl_beschriftung 
         Caption         =   "Beschreibung.."
         Height          =   2055
         Index           =   5
         Left            =   1680
         TabIndex        =   56
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label lbl_beschriftung 
         Caption         =   "Beschreibung.."
         Height          =   2055
         Index           =   6
         Left            =   -73320
         TabIndex        =   46
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label lbl_beschriftung 
         Caption         =   "Beschreibung.."
         Height          =   495
         Index           =   4
         Left            =   -74760
         TabIndex        =   29
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lbl_beschriftung 
         Caption         =   "Beschreibung.."
         Height          =   495
         Index           =   3
         Left            =   -74760
         TabIndex        =   27
         Top             =   600
         Width           =   4455
      End
      Begin VB.Label lbl_beschriftung 
         Caption         =   "Beschreibung.."
         Height          =   2055
         Index           =   2
         Left            =   -73320
         TabIndex        =   26
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label lbl_beschriftung 
         Caption         =   "Beschreibung.."
         Height          =   2055
         Index           =   1
         Left            =   -73320
         TabIndex        =   25
         Top             =   900
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_optionssoftware"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Variable für Schleifen etc.
Private i As Integer
'Variable für die ausgewählte Version
Private SelectedVersion As Integer
'Rückgabewert von Messageboxen
Private RetVal As Long
'Dateinamen der Language-Files
Private LngFileNames() As String

Private Sub Form_Load()
'--------------------------------------------------------------------'
'| Prozedur beim Laden der Form: Initialisierungen                  |'
'--------------------------------------------------------------------'
   'Log: Software-Fenster wurde geladen
   LogEvent "Software-Fenster geladen"
   
   'Wenn Spalten = 0 ist, dann Spalten auf 512 setzen
   If Not Spalten Then Spalten = 512
   
   'Wenn RGB nicht verfügbar ist..
   If Not NGRGBAvailable Then
      'Tab mit RGB-Einstellungen deaktivieren
      tab_optionen.TabEnabled(5) = False
   End If
   
   'Wenn NG nicht mit der Hardware verbunden ist..
   If Not Connected2Hardware Then
      'Tab mit SW-Einstellungen deaktivieren
      tab_optionen.TabEnabled(4) = False
   End If
   
   'Fokus auf ersten Tab setzen
   tab_optionen.Tab = 0
   
   'Log: Icons initialisieren
   LogEvent "Icons werden initialisiert"
   
   'Icon der Form initialisieren
   Icons_Init Me
   'Aktuelles Offset einstellen
   txt_offset.Text = CStr(Offset) & LngSpecStrings.Columns
   'Aktuelle Drehrichtung einstellen
   opt_richtung((CInt(Not NGRotationSystemLeft)) + 1).Value = True
   'Rotorgröße anzeigen
   txt_rotorgröße.Text = CStr(NGRotorSize) & " mm"
   'Demogröße anzeigen
   opt_demogröße(NGDemoSize).Value = True
   'DemoVersion einstellen
   opt_demoversion(NGDemoVersion).Value = True
   'Transparenz einstellen
   chk_transparentimport.Value = NGTransparentImport
   'Transparente Farbe einstellen
   opt_farbe(Color2Number(NGTransparentColor)).Value = True
   
   'Wenn LiPo-Kapazität ungültig ist, auf 100 setzen
   If (NGLiPoCapacity < 100) Or (NGLiPoCapacity > 500) Then NGLiPoCapacity = 100
   'Schieberegler auf LiPo-Kapazität einstellen
   hsc_lipocapacity.Value = NGLiPoCapacity
   'Wert formatiert in die Textbox schreiben
   txt_lipocapacity.Text = Replace(Format(CStr(hsc_lipocapacity.Value), "000 mAh"), ",", ".")

   'Log: Verfügbare Comports
   LogEvent "Verfügbare COM-Ports werden abgefragt"

   'Schleife von 1 bis 16 (durch alle möglichen Comports)
   For i = 1 To 16
      'Ist der Comport verfügbar ?
      If ComPort_Available(frm_nightgraphix.msc_comport, i) Then
         'Comport in die Liste eintragen
         lst_comports.AddItem "COM" & Format(CStr(i), "00")
         'Wenn Comport durchsucht werden soll..
         If NGComportSearch(i) = True Then
            'Item in Listbox anwählen
            lst_comports.Selected(lst_comports.ListCount - 1) = True
         End If
      End If
   Next i
   
   'Log:Alle Lng-Files werden ermittelt
   LogEvent "Alle LngFiles werden ermittelt"
   
   'Alle Language-Dateien ermitteln
   LngFileNames = FileSystem_GetFilesInFolder(LngFilePath, "*.lng")
   'Schleife durch alle Dateinamen
   For i = 0 To UBound(LngFileNames)
      'Dateiname in ListBox eintragen
      lst_languages.AddItem LngFileNames(i)
   Next i
   
   'Wenn NG mit der Hardware verbunden ist..
   If Connected2Hardware Then
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
   End If

   'Log: Sprache wird gesetzt
   LogEvent "Sprache wird gesetzt" & vbCrLf
   
   'Sprache für alle Controls setzen
   If Not LngInProcess Then Language_SetControlProperties

'   'Text für Beschreibung in System-Tab eintragen
'   lbl_beschriftung(1).Caption = "Links können Sie auswählen, welches NGX-System Sie haben." & vbCrLf & "Rechts können Sie Ihre Rotorgröße eintragen, danach klicken Sie auf ""Rotorgröße übernehmen""."
'   'Text für Beschreibung in Hardware-Tab eintragen
'   lbl_beschriftung(2).Caption = "Links können Sie auswählen, ob Ihr Heli links- oder rechtsdrehend ist. Außerdem können Sie die Offset-Spalte auswählen." & vbCrLf & "Rechts können Sie die Kapazität Ihres Akkus eintragen."
'   'Text für Beschreibung in Verbindung-Tab eintragen
'   lbl_beschriftung(3).Caption = "Wählen Sie die Ports aus, die beim automatischen Verbinden auf NGX-Hardware überprüft werden sollen."
'   'Text für Beschreibung in Sprache-Tab eintragen
'   lbl_beschriftung(4).Caption = "Wählen Sie die Sprache, die Sie in NG verwenden möchten."
'   'Text für Beschreibung in RGB-Tab eintragen
'   lbl_beschriftung(5).Caption = "Links können Sie auswählen, ob Sie die Demo in RGB oder SW starten wollen." & vbCrLf & "Rechts können Sie auswählen, ob beim Bildimport eine Farbe transparent wird."
End Sub

Private Sub cmd_ok_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Drücken des Buttons "OK"                           |'
'--------------------------------------------------------------------'
   'LiPo-Kapazität in Variable speichern
   NGLiPoCapacity = hsc_lipocapacity.Value
   
   'Wenn Richtung: Rechts ist..
   If opt_richtung(0).Value = True Then
      'Richtung in Variable schreiben
      NGRotationSystemLeft = False
   'Wenn Richtung: Links ist..
   Else
      'Richtung in Variable schreiben
      NGRotationSystemLeft = True
   End If
   
   'Wenn Demoversion: SW ist..
   If opt_demoversion(0).Value Then
      'Demoversion in Variable speichern
      NGDemoVersion = 0
   'Wenn Demoversion: RGB ist..
   Else
      'Demoversion in Variable speichern
      NGDemoVersion = 1
   End If
   
   'Transparenz beim Import speichern
   NGTransparentImport = chk_transparentimport.Value
   
   'Schleife durch alle Importfarben
   For i = 0 To opt_farbe.UBound
      'Wenn der Button mit dem Index i angewählt ist..
      If opt_farbe(i).Value Then
         'Farbe in Variable speichern
         NGTransparentColor = Number2Color(i)
      End If
   Next i
   
   'Schleife durch alle Demogrößen
   For i = 0 To opt_demogröße.UBound
      'Wenn der Button mit dem Index i angewählt ist..
      If opt_demogröße(i).Value = True Then
         'Demogröße setzen
         NGDemoSize = i
         'Wenn die Rotorgröße nicht dem Wert in der Textbox entspricht..
         If NGRotorSize <> Val(txt_rotorgröße.Text) Then
            'Wenn Rotorgröße ungültig ist..
            If (Val(txt_rotorgröße.Text) - 2 * (i * 8 + 16) * 10 - 2 * 35) / 20 <= 6 Then
               'MessageBox anzeigen
               RetVal = MsgBox(LngSpecStrings.InvalidRotorSize & vbCrLf & LngSpecStrings.AskToKeepOldRotorSize, vbYesNo, "NightGraphiX V1.0 - " & LngSpecStrings.Error)
               'Wenn Nein geklickt wurde..
               If RetVal = vbNo Then
                  'Fokus in die Textbox setzen
                  txt_rotorgröße.SetFocus
                  'Text in Textbox selektieren
                  txt_Rotorgröße_DblClick
                  'Prozedur beenden, damit das Fenster nicht entladen wird
                  Exit Sub
               End If
            Else
               'NGRotorSize speichern
               NGRotorSize = Val(txt_rotorgröße.Text)
               'Rotorgröße im NGRotorSizeArray speichern
               NGRotorSizeArray(i) = NGRotorSize
            End If
         End If
      End If
   Next i
   
   'Wenn Offset verändert wurde..
   If Offset <> CLng(Val(txt_offset.Text)) Then
      'Wenn Offset ungültig ist..
      If Val(txt_offset.Text) > (Spalten - 1) Then
         'MessageBox anzeigen
         RetVal = MsgBox(LngSpecStrings.InvalidOffset & " (>" & CStr(Spalten - 1) & ")" & vbCrLf & LngSpecStrings.AskToKeepOldOffset, vbYesNo, "NightGraphiX V1.0 - " & LngSpecStrings.Error)
         'Wenn Nein geklickt wurde..
         If RetVal = vbNo Then
            'Fokus in die Textbox setzen
            txt_offset.SetFocus
            'Text in Textbox selektieren
            txt_offset_DblClick
            'Prozedur beenden, damit das Fenster nicht entladen wird
            Exit Sub
         End If
      'Wenn Offset gültig ist..
      Else
         'Offset in Variable schreiben
         Offset = CLng(Val(txt_offset.Text))
         'Wenn Demomodus gestartet oder Hardware verbunden ist..
         If NGDemoModus Or Connected2Hardware Then
               'Offset-Pfeil neu setzen
            Draw_SetOffsetArrow frm_nightgraphix.pic_arrowshow, Offset * CDbl(360) / (Spalten - 1)
            'Offset-Pfeil zeichnen
            frm_nightgraphix.Draw_OffsetCursor Offset * CDbl(360) / (Spalten - 1)
         End If
      End If
   End If
   
   'Schleife durch alle verfügbaren Comports
   For i = 0 To lst_comports.ListCount - 1
      'Wenn der Eintrag in der Listbox angewählt ist..
      If lst_comports.Selected(i) Then
         'Comport beim Verbinden durchsuchen
         NGComportSearch(Val(Right(lst_comports.List(i), 2))) = True
      'Wenn der Eintrag in der Listbox nicht ausgewählt ist..
      Else
         'Comport beim Verbinden nicht durchsuchen
         NGComportSearch(Val(Right(lst_comports.List(i), 2))) = False
      End If
   Next i
   
   'Wenn NG mit der HW verbunden ist..
   If Connected2Hardware Then
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
   End If
   
   'Daten in der Registry speichern
   FileSystem_SaveSettings
   
   'Form beenden
   Unload Me
End Sub

Private Sub cmd_abbrechen_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Drücken des Buttons "Abbrechen"                    |'
'--------------------------------------------------------------------'
   'Form beenden
   Unload Me
End Sub

Private Sub cmd_saverotorsize_Click()
'--------------------------------------------------------------------'
'| Prozedur zum Speichern der Rotorgröße der ausgewählten Version   |'
'--------------------------------------------------------------------'
   'Wenn der Hub kleiner als 0 wäre..
   If (Val(txt_rotorgröße.Text) - 2 * (SelectedVersion * 8 + 16) * 10) / 10 < 0 Then
      'Messagebox mit Warnung anzeigen
      MsgBox LngSpecStrings.RotorSizeTooSmall, vbOKOnly & vbCritical, "NightVision V1.0 - " & LngSpecStrings.Warning
      'Prozedur beenden
      Exit Sub
   End If
   'Wert für Rotorgröße in Variable speichern
   NGRotorSizeArray(SelectedVersion) = Val(txt_rotorgröße.Text)
End Sub

Private Sub lst_languages_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf eine Sprache in der ListBox            |'
'--------------------------------------------------------------------'
   'Pfad mit Language-Dateien setzen
   Language_SetFilePath FileSystem_ClearPath(App.Path) & "\Language\"
   'Language-Datei auslesen
   Language_ReadFromFile lst_languages.List(lst_languages.ListIndex)
   'Die Beschriftungen und Tooltips aller Controls setzen
   Language_SetControlProperties
End Sub

Private Sub opt_demogröße_Click(Index As Integer)
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf einen Optionbutton                     |'
'--------------------------------------------------------------------'
   'Rotorgröße der Hardwareversion für die Demo in der Textbox anzeigen
   txt_rotorgröße.Text = CStr(NGRotorSizeArray(Index)) & " mm"
   'Ausgewählte Version in Variable speichern
   SelectedVersion = Index
End Sub

Private Sub txt_lipocapacity_GotFocus()
'--------------------------------------------------------------------'
'| Prozedur beim Erhalten des Fokusses in txt_lipocapacity          |'
'--------------------------------------------------------------------'
   'Fokus auf Slider setzen
   hsc_lipocapacity.SetFocus
End Sub

'#####################Filter für Offset-Feld#######################
Private Sub txt_Rotorgröße_DblClick()
'--------------------------------------------------------------------'
'| Prozedur bei Doppelklick in txt_Rotorgröße                       |'
'--------------------------------------------------------------------'
   'Selektierung von Start bis zum Leerzeichen
   txt_rotorgröße.SelStart = 0
   txt_rotorgröße.SelLength = InStr(1, txt_rotorgröße.Text, " mm") - 1
End Sub

Private Sub txt_Rotorgröße_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------'
'| Prozedur bei Drücken einer Taste in txt_Rotorgröße               |'
'--------------------------------------------------------------------'
   'Wenn Entfernen gedrückt wurde..
   If KeyCode = vbKeyDelete Then
      'Wenn etwas von " mm" gelöscht werden soll..
      If txt_rotorgröße.SelStart >= InStr(1, txt_rotorgröße.Text, " mm") - 1 Then
         'Tastendruck ungültig machen
         KeyCode = 0
      End If
   'Wenn rechte Pfeiltaste gedrückt wird..
   ElseIf (KeyCode = vbKeyRight) Or (KeyCode = vbKeyDown) Then
      'Wenn auf " mm" gesprungen werden soll..
      If txt_rotorgröße.SelStart >= InStr(1, txt_rotorgröße.Text, " mm") - 1 Then
         'Tastendruck ungültig machen
         KeyCode = 0
      End If
   End If
End Sub

Private Sub txt_Rotorgröße_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------------'
'| Prozedur bei Drücken einer Taste in txt_Rotorgröße               |'
'--------------------------------------------------------------------'
   'Wenn das eingegebene Zeichen nicht numerisch oder vbKeyBack ist..
   If (Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = vbKeyBack)) Or ((Len(Mid(txt_rotorgröße.Text, 1, InStr(1, txt_rotorgröße.Text, " mm"))) >= 8) And (KeyAscii <> vbKeyBack)) Then
      'Tastendruck ungültig machen
      KeyAscii = 0
   End If
End Sub

Private Sub txt_Rotorgröße_KeyUp(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------'
'| Prozedur bei Loslassen einer Taste in txt_Rotorgröße             |'
'--------------------------------------------------------------------'
   'Wenn das eingegebene Zeichen vbKeyBack ist..
   If KeyCode = vbKeyBack Then
      'Wenn kein Zeichen mehr zwischen Start und " mm" steht
      If Mid(txt_rotorgröße.Text, 1, InStr(1, txt_rotorgröße.Text, " mm") - 1) = "" Then
         'Eine "0" am Anfang einfügen
         txt_rotorgröße.Text = "0" & txt_rotorgröße.Text
         'Die "0" selektieren
         txt_rotorgröße.SelStart = 0
         txt_rotorgröße.SelLength = 1
      End If
   End If
End Sub

Private Sub txt_Rotorgröße_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur bei Drücken einer Maustaste in txt_Rotorgröße           |'
'--------------------------------------------------------------------'
   'Wenn die aktuelle Cursorposition größer als die Position des " mm" ist..
   If txt_rotorgröße.SelStart >= InStr(1, txt_rotorgröße.Text, " mm") Then
      'Cursorposition vor " mm" setzen
      txt_rotorgröße.SelStart = InStr(1, txt_rotorgröße.Text, " mm") - 1
   End If
End Sub

Private Sub txt_Rotorgröße_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur bei Bewegen der Maus in txt_Rotorgröße                  |'
'--------------------------------------------------------------------'
   'Wenn die linke Maustaste gedrückt ist..
   If Button = vbLeftButton Then
      'Wenn die Startposition des selektierten Textes plus
      'die Länge des selektierten Textes größer als
      'die Position von " mm" ist..
      If txt_rotorgröße.SelStart + txt_rotorgröße.SelLength >= InStr(1, txt_rotorgröße.Text, " mm") Then
         'Länge des selektierten Textes verkleinern
         txt_rotorgröße.SelLength = InStr(1, txt_rotorgröße.Text, " mm") - 1 - txt_rotorgröße.SelStart
      End If
   End If
End Sub

'#####################Filter für Offset-Feld#######################
Private Sub txt_offset_DblClick()
'--------------------------------------------------------------------'
'| Prozedur bei Doppelklick in txt_offset                       |'
'--------------------------------------------------------------------'
   'Selektierung von Start bis zum Leerzeichen
   txt_offset.SelStart = 0
   txt_offset.SelLength = InStr(1, txt_offset.Text, LngSpecStrings.Columns) - 1
End Sub

Private Sub txt_offset_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------'
'| Prozedur bei Drücken einer Taste in txt_offset               |'
'--------------------------------------------------------------------'
   'Wenn Entfernen gedrückt wurde..
   If KeyCode = vbKeyDelete Then
      'Wenn etwas von " mm" gelöscht werden soll..
      If txt_offset.SelStart >= InStr(1, txt_offset.Text, LngSpecStrings.Columns) - 1 Then
         'Tastendruck ungültig machen
         KeyCode = 0
      End If
   'Wenn rechte Pfeiltaste gedrückt wird..
   ElseIf (KeyCode = vbKeyRight) Or (KeyCode = vbKeyDown) Then
      'Wenn auf " mm" gesprungen werden soll..
      If txt_offset.SelStart >= InStr(1, txt_offset.Text, LngSpecStrings.Columns) - 1 Then
         'Tastendruck ungültig machen
         KeyCode = 0
      End If
   End If
End Sub

Private Sub txt_offset_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------------'
'| Prozedur bei Drücken einer Taste in txt_offset               |'
'--------------------------------------------------------------------'
   'Wenn das eingegebene Zeichen nicht numerisch oder vbKeyBack ist..
   If (Not (IsNumeric(Chr(KeyAscii)) Or KeyAscii = vbKeyBack)) Or ((Len(Mid(txt_offset.Text, 1, InStr(1, txt_offset.Text, LngSpecStrings.Columns))) >= 8) And (KeyAscii <> vbKeyBack)) Then
      'Tastendruck ungültig machen
      KeyAscii = 0
   End If
End Sub

Private Sub txt_offset_KeyUp(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------'
'| Prozedur bei Loslassen einer Taste in txt_offset             |'
'--------------------------------------------------------------------'
   'Wenn das eingegebene Zeichen vbKeyBack ist..
   If KeyCode = vbKeyBack Then
      'Wenn kein Zeichen mehr zwischen Start und " mm" steht
      If Mid(txt_offset.Text, 1, InStr(1, txt_offset.Text, LngSpecStrings.Columns) - 1) = "" Then
         'Eine "0" am Anfang einfügen
         txt_offset.Text = "0" & txt_offset.Text
         'Die "0" selektieren
         txt_offset.SelStart = 0
         txt_offset.SelLength = 1
      End If
   End If
End Sub

Private Sub txt_offset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur bei Drücken einer Maustaste in txt_offset           |'
'--------------------------------------------------------------------'
   'Wenn die aktuelle Cursorposition größer als die Position des " mm" ist..
   If txt_offset.SelStart >= InStr(1, txt_offset.Text, LngSpecStrings.Columns) Then
      'Cursorposition vor " mm" setzen
      txt_offset.SelStart = InStr(1, txt_offset.Text, LngSpecStrings.Columns) - 1
   End If
End Sub

Private Sub txt_offset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur bei Bewegen der Maus in txt_offset                  |'
'--------------------------------------------------------------------'
   'Wenn die linke Maustaste gedrückt ist..
   If Button = vbLeftButton Then
      'Wenn die Startposition des selektierten Textes plus
      'die Länge des selektierten Textes größer als
      'die Position von " Spalte(n)" ist..
      If txt_offset.SelStart + txt_offset.SelLength >= InStr(1, txt_offset.Text, LngSpecStrings.Columns) Then
         'Länge des selektierten Textes verkleinern
         txt_offset.SelLength = InStr(1, txt_offset.Text, LngSpecStrings.Columns) - 1 - txt_offset.SelStart
      End If
   End If
End Sub

Private Sub hsc_lipocapacity_Change()
'--------------------------------------------------------------------'
'| Prozedur beim Verändern des Wertes der Scrollbar                 |'
'--------------------------------------------------------------------'
   'Wert formatiert in die Textbox schreiben
   txt_lipocapacity.Text = Replace(Format(CStr(hsc_lipocapacity.Value), "000 mAh"), ",", ".")
End Sub

Private Sub hsc_lipocapacity_Scroll()
'--------------------------------------------------------------------'
'| Prozedur beim Verschieben der Scrollbar                          |'
'--------------------------------------------------------------------'
   'Wert formatiert in die Textbox schreiben
   txt_lipocapacity.Text = Replace(Format(CStr(hsc_lipocapacity.Value), "000 mAh"), ",", ".")
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

