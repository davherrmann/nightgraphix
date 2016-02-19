VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_nightgraphix 
   Caption         =   "NightGraphiX V1.0"
   ClientHeight    =   10635
   ClientLeft      =   2580
   ClientTop       =   450
   ClientWidth     =   10305
   Icon            =   "frm_nightgraphix.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   709
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   687
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.Toolbar tlb_toolbar 
      Align           =   1  'Oben ausrichten
      Height          =   660
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "img_imagelist"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   24
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Neue Datei"
            Object.Tag             =   ""
            ImageIndex      =   18
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Datei öffnen"
            Object.Tag             =   ""
            ImageIndex      =   19
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Bild importieren"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Datei speichern"
            Object.Tag             =   ""
            ImageIndex      =   28
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Datei speichern unter.."
            Object.Tag             =   ""
            ImageIndex      =   35
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Löschen"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Zoom +"
            Object.Tag             =   ""
            ImageIndex      =   32
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Zoom -"
            Object.Tag             =   ""
            ImageIndex      =   33
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "EEPROM schreiben"
            Object.Tag             =   ""
            ImageIndex      =   31
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "EEPROM auslesen"
            Object.Tag             =   ""
            ImageIndex      =   25
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Verbinden"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Trennen"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Demo starten"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Demo beenden"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Optionen - Software"
            Object.Tag             =   ""
            ImageIndex      =   20
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "Optionen - Hardware"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "Über NightGraphiX"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
            Object.Width           =   110
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "NG1.0 Beenden"
            Object.Tag             =   ""
            ImageIndex      =   34
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic_farbwahl 
      Appearance      =   0  '2D
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   630
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   167
      TabIndex        =   38
      Top             =   2820
      Visible         =   0   'False
      Width           =   2535
      Begin VB.OptionButton opt_farbe 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   2115
         Style           =   1  'Grafisch
         TabIndex        =   46
         Top             =   45
         Width           =   255
      End
      Begin VB.OptionButton opt_farbe 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   1830
         Style           =   1  'Grafisch
         TabIndex        =   45
         Top             =   45
         Width           =   255
      End
      Begin VB.OptionButton opt_farbe 
         BackColor       =   &H00FFFF00&
         Height          =   255
         Index           =   5
         Left            =   1545
         Style           =   1  'Grafisch
         TabIndex        =   44
         Top             =   45
         Width           =   255
      End
      Begin VB.OptionButton opt_farbe 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Index           =   4
         Left            =   1260
         Style           =   1  'Grafisch
         TabIndex        =   43
         Top             =   45
         Width           =   255
      End
      Begin VB.OptionButton opt_farbe 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   975
         Style           =   1  'Grafisch
         TabIndex        =   42
         Top             =   45
         Width           =   255
      End
      Begin VB.OptionButton opt_farbe 
         BackColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   690
         Style           =   1  'Grafisch
         TabIndex        =   41
         Top             =   45
         Width           =   255
      End
      Begin VB.OptionButton opt_farbe 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   405
         Style           =   1  'Grafisch
         TabIndex        =   40
         Top             =   45
         Width           =   255
      End
      Begin VB.OptionButton opt_farbe 
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Grafisch
         TabIndex        =   39
         Top             =   45
         Width           =   255
      End
   End
   Begin VB.OptionButton cmd_tool 
      Height          =   495
      Index           =   3
      Left            =   120
      Picture         =   "frm_nightgraphix.frx":1982
      Style           =   1  'Grafisch
      TabIndex        =   37
      ToolTipText     =   "Farbwahl"
      Top             =   2760
      Width           =   495
   End
   Begin VB.PictureBox pic_lupeshow 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   4080
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   36
      Top             =   15000
      Width           =   1920
   End
   Begin VB.Timer tmr_lupe 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   4920
      Top             =   8400
   End
   Begin VB.PictureBox pic_lupemask 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      FillStyle       =   0  'Ausgefüllt
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   5400
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   34
      Top             =   15000
      Width           =   1920
   End
   Begin VB.PictureBox pic_lupebitmap 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'Kein
      FillStyle       =   0  'Ausgefüllt
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   7320
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   33
      Top             =   15000
      Width           =   1920
   End
   Begin VB.PictureBox pic_alphablend 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   405
      Left            =   23760
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   32
      Top             =   23400
      Width           =   405
   End
   Begin VB.Timer tmr_comports 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5400
      Top             =   8400
   End
   Begin VB.OptionButton cmd_tool 
      Height          =   495
      Index           =   6
      Left            =   120
      Picture         =   "frm_nightgraphix.frx":224C
      Style           =   1  'Grafisch
      TabIndex        =   22
      ToolTipText     =   "Löschen"
      Top             =   4680
      Width           =   495
   End
   Begin VB.OptionButton cmd_tool 
      Height          =   495
      Index           =   5
      Left            =   120
      Picture         =   "frm_nightgraphix.frx":2B16
      Style           =   1  'Grafisch
      TabIndex        =   21
      ToolTipText     =   "Invertieren"
      Top             =   4080
      Width           =   495
   End
   Begin VB.OptionButton cmd_tool 
      Height          =   495
      Index           =   4
      Left            =   120
      Picture         =   "frm_nightgraphix.frx":33E0
      Style           =   1  'Grafisch
      TabIndex        =   20
      ToolTipText     =   "Bild importieren"
      Top             =   3480
      Width           =   495
   End
   Begin VB.Timer tmr_offsetarrow 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5880
      Top             =   8400
   End
   Begin ComctlLib.ProgressBar prg_fortschritt 
      Height          =   315
      Left            =   5880
      TabIndex        =   8
      Top             =   10320
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_font 
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      Style           =   1  'Grafisch
      TabIndex        =   13
      ToolTipText     =   "Schriftauswahl"
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox pic_letter 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1650
      Left            =   960
      ScaleHeight     =   110
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   12
      Top             =   12600
      Width           =   1650
   End
   Begin VB.OptionButton cmd_tool 
      Height          =   495
      Index           =   1
      Left            =   120
      Picture         =   "frm_nightgraphix.frx":3CAA
      Style           =   1  'Grafisch
      TabIndex        =   11
      ToolTipText     =   "Kreise zeichnen"
      Top             =   1440
      Width           =   495
   End
   Begin VB.OptionButton cmd_tool 
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "frm_nightgraphix.frx":4974
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Einzelne Pixel setzen"
      Top             =   840
      Value           =   -1  'True
      Width           =   495
   End
   Begin VB.OptionButton cmd_tool 
      Height          =   495
      Index           =   2
      Left            =   120
      Picture         =   "frm_nightgraphix.frx":523E
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Text einfügen"
      Top             =   2040
      Width           =   495
   End
   Begin MSComDlg.CommonDialog cdlg_dialog 
      Left            =   6960
      Top             =   8880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmr_cursor 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6840
      Top             =   8400
   End
   Begin VB.PictureBox pic_bildgroß 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   405
      Left            =   7320
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   3
      Top             =   8400
      Width           =   405
   End
   Begin VB.PictureBox pic_source 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   405
      Left            =   7800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   4
      Top             =   8400
      Width           =   405
   End
   Begin VB.PictureBox pic_bild 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   405
      Left            =   8280
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   5
      Top             =   8400
      Width           =   405
   End
   Begin VB.Timer tmr_microcontroller 
      Left            =   6360
      Top             =   8400
   End
   Begin VB.VScrollBar vsc_scrollen 
      Height          =   9135
      LargeChange     =   210
      Left            =   9840
      Max             =   3394
      SmallChange     =   45
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin VB.HScrollBar hsc_scrollen 
      Height          =   255
      LargeChange     =   210
      Left            =   720
      Max             =   3394
      SmallChange     =   45
      TabIndex        =   6
      Top             =   9840
      Width           =   9135
   End
   Begin MSCommLib.MSComm msc_seriell 
      Left            =   6480
      Top             =   8880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin ComctlLib.StatusBar stb_statusbar 
      Align           =   2  'Unten ausrichten
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   10260
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2778
            MinWidth        =   2778
            Text            =   "System Information"
            TextSave        =   "System Information"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7461
            MinWidth        =   4762
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   7408
            MinWidth        =   7408
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic_rotatearrowbitmap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   285
      Left            =   0
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   14
      Top             =   37500
      Width           =   285
   End
   Begin VB.PictureBox pic_rotatearrowmask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   285
      Left            =   360
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   15
      Top             =   37500
      Width           =   285
   End
   Begin VB.PictureBox pic_arrowbitmap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   195
      Left            =   120
      Picture         =   "frm_nightgraphix.frx":5B08
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   16
      Top             =   37500
      Width           =   180
   End
   Begin VB.PictureBox pic_arrowmask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   195
      Left            =   360
      Picture         =   "frm_nightgraphix.frx":5D1E
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   12
      TabIndex        =   17
      Top             =   37500
      Width           =   180
   End
   Begin MSCommLib.MSComm msc_comport 
      Left            =   5880
      Top             =   8880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   6
   End
   Begin VB.PictureBox pic_rahmen 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9135
      Left            =   720
      ScaleHeight     =   607
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   607
      TabIndex        =   0
      Top             =   720
      Width           =   9135
      Begin VB.CommandButton cmd_cancelimport 
         Caption         =   "&Stop"
         Height          =   375
         Left            =   960
         MousePointer    =   1  'Pfeil
         TabIndex        =   48
         Top             =   8640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmd_finishimport 
         Caption         =   "&Fertig"
         Height          =   375
         Left            =   120
         MousePointer    =   1  'Pfeil
         TabIndex        =   47
         Top             =   8640
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.PictureBox pic_lupe 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   1920
         Left            =   5880
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   128
         TabIndex        =   35
         Top             =   15000
         Width           =   1920
      End
      Begin VB.PictureBox pic_import 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   3600
         ScaleHeight     =   143
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   143
         TabIndex        =   31
         Top             =   3600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox pic_resizeimport 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00000000&
         Height          =   90
         Index           =   0
         Left            =   4440
         MousePointer    =   8  'Größenänderung NW SO
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   30
         Top             =   8280
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox pic_resizeimport 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00000000&
         Height          =   90
         Index           =   1
         Left            =   4440
         MousePointer    =   9  'Größenänderung W O
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   29
         Top             =   8520
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox pic_resizeimport 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00000000&
         Height          =   90
         Index           =   2
         Left            =   4440
         MousePointer    =   6  'Größenänderung NO SW
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   28
         Top             =   8760
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox pic_resizeimport 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00000000&
         Height          =   90
         Index           =   3
         Left            =   4920
         MousePointer    =   6  'Größenänderung NO SW
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   27
         Top             =   8280
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox pic_resizeimport 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00000000&
         Height          =   90
         Index           =   4
         Left            =   4920
         MousePointer    =   9  'Größenänderung W O
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   26
         Top             =   8520
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox pic_resizeimport 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00000000&
         Height          =   90
         Index           =   5
         Left            =   4920
         MousePointer    =   8  'Größenänderung NW SO
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   25
         Top             =   8760
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox pic_resizeimport 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00000000&
         Height          =   90
         Index           =   6
         Left            =   4680
         MousePointer    =   7  'Größenänderung N S
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   24
         Top             =   8280
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox pic_resizeimport 
         Appearance      =   0  '2D
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H00000000&
         Height          =   90
         Index           =   7
         Left            =   4680
         MousePointer    =   7  'Größenänderung N S
         ScaleHeight     =   6
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   6
         TabIndex        =   23
         Top             =   8760
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.PictureBox pic_target 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'Kein
         Height          =   3045
         Left            =   0
         ScaleHeight     =   203
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   211
         TabIndex        =   1
         Top             =   0
         Width           =   3165
         Begin VB.PictureBox pic_arrowshow 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00A0A0A0&
            BorderStyle     =   0  'Kein
            Height          =   285
            Left            =   1560
            ScaleHeight     =   19
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   19
            TabIndex        =   18
            Top             =   1560
            Width           =   285
         End
      End
      Begin ComctlLib.ImageList img_icons 
         Left            =   8040
         Top             =   8400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
      Begin ComctlLib.ImageList img_imagelist 
         Left            =   6840
         Top             =   8160
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   255
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   36
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":5F34
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":5F92
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":5FF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":604E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":60AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":610A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6168
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":61C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6224
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6282
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":62E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":633E
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":639C
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":63FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6458
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":64B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6514
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6572
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":65D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":662E
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":668C
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":66EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6748
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":67A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6804
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6862
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":68C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":691E
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":697C
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":69DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6A38
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6A96
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6AF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6B52
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6BB0
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frm_nightgraphix.frx":6C0E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnu_datei 
      Caption         =   "&Datei"
      Visible         =   0   'False
      Begin VB.Menu mnu_öffnen 
         Caption         =   "&Öffnen"
      End
      Begin VB.Menu mnu_importieren 
         Caption         =   "&Importieren"
      End
      Begin VB.Menu mnu_texteingabe 
         Caption         =   "&Texteingabe"
      End
      Begin VB.Menu mnu_speichern 
         Caption         =   "&Speichern"
      End
      Begin VB.Menu mnu_speichernunter 
         Caption         =   "Speichern &unter..."
      End
      Begin VB.Menu mnu_löschen 
         Caption         =   "&Löschen"
      End
      Begin VB.Menu mnu_bindestrich 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_beenden 
         Caption         =   "&Beenden"
      End
   End
   Begin VB.Menu mnu_bilder 
      Caption         =   "&Bilder"
      Visible         =   0   'False
      Begin VB.Menu mnu_bild 
         Caption         =   "Bild &1"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnu_bild 
         Caption         =   "Bild &2"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnu_bild 
         Caption         =   "Bild &3"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mnu_bild 
         Caption         =   "Bild &4"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mnu_extras 
      Caption         =   "&Extras"
      Visible         =   0   'False
      Begin VB.Menu mnu_nvinfo 
         Caption         =   "&Über NightVision"
      End
      Begin VB.Menu mnu_hilfe 
         Caption         =   "&Hilfe"
      End
   End
End
Attribute VB_Name = "frm_nightgraphix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variablen müssen deklariert werden..
Option Explicit

'VARIABLEN..
'Variable für Schleifen
Private x As Integer
Private y As Integer
'Variable für Schleifen etc.
Private i As Integer

'KONSTANTEN..
'Konstante für den Abstand der blauen Kästchen zum Control
Const Sp = 0&

'Minimale & Maximale Abmaße des Controls
Const Wmin = 66&
Const Wmax = 594&
Const Hmin = 66&
Const Hmax = 594&

'Eckkoordinaten des erlaubten Bewegungsraumes
Const LimX1 = -100&
Const LimY1 = -100&
Const LimX2 = 700&
Const LimY2 = 700&

'Ist die Maustaste gedrückt ? Wird die Maus bewegt ?
Dim DragFlag As Boolean, MoveFlag As Boolean
Dim Fetched As Boolean
Dim StartX&, Starty&
'Referenz auf das Steuerelement
Dim MCtrl As Control

Private Sub Form_Load()
'--------------------------------------------------------------------'
'| Prozedur beim Laden der Form: Initialisierungen                  |'
'--------------------------------------------------------------------'
   'Icon der Form initialisieren
   Icons_Init Me

   'FillArea initialisieren
   Draw_InitFillArea
   
   'Bilder-Menü initialisieren
   Init_Menu mnu_bild
           
   'Pictureboxen initialisieren (Größe, AutoRedraw)
   Init_PictureBoxes pic_source, pic_target, pic_rahmen, pic_bild, pic_bildgroß
   
   'Imagelist aktualisieren
   Init_ImageList img_imagelist
   
   'Toolbar initialisieren
   Init_Toolbar tlb_toolbar
   
   
   '##########ZUM DEBUGGEN MUSS DIESE ZEILE AUSKOMMENTIERT WERDEN !!###############'
   'Scrollrad initialisieren
   'HookScrollWheel_Init Me
   '##########ZUM DEBUGGEN MUSS DIESE ZEILE AUSKOMMENTIERT WERDEN !!###############'
   
   'Wenn die Hardware nicht verbunden ist..
   If Not Connected2Hardware Then
      'Schreiben des EEPROMS deaktivieren
      tlb_toolbar.Buttons(12).Enabled = False
      'Auslesen des EEPROMS deaktivieren
      tlb_toolbar.Buttons(13).Enabled = False
   End If
   
'   'Variablen initialisieren
'   Init_Variables pic_source

'   'Font-PictureBox initialisieren
'   Init_FontPictureBox pic_letter
'
'   'DateiHeader initialisieren
'   Init_FileHeader
'
'   'Arrays initialisieren
'   Init_Arrays
'
   'Werte der ScrollBars initialisieren
   Init_Scrollbars vsc_scrollen, hsc_scrollen, pic_rahmen
      
   'Referenz auf WScript.Shell setzen
   FileSystem_CreateWScriptReference
   
   'Registry-Einträge prüfen
   FileSystem_CheckRegistry4Extension
   
   'Daten aus der Registry auslesen
   FileSystem_GetSettings
   
'   'Eigenschaften in Statusbar schreiben
'   NightGraphix_ShowPanelProperties
   
   '#################################ZEICHNEN##################################'
'   'Modell zeichnen
'   Draw_Modell pic_source

   'Hintergrund grau füllen
   Draw_Background pic_source

   'pic_source refreshen
   Draw_Refresh pic_source

   'Bild von pic_source nach pic_target zoomen
   Draw_Zoom pic_source, pic_target

   'PictureBox refreshen
   Draw_Refresh pic_target

   'Bild klein zoomen
   Draw_ZoomModell (1)

   'Buttons deaktivieren
   NightGraphix_DisableToolButtons
   
   'Pfad mit Language-Dateien setzen
   Language_SetFilePath FileSystem_ClearPath(App.Path) & "\Language\"
   'Language-Datei auslesen
   Language_ReadFromFile CurrentLanguage
   'Die Beschriftungen und Tooltips aller Controls setzen
   Language_SetControlProperties

   
'   'Offset-Pfeil setzen
'   Draw_SetOffsetArrow pic_arrowshow, Offset * CDbl(360) / (Spalten - 1)

'   'Offset-Pfeil zeichnen
'   Draw_OffsetCursor Offset * CDbl(360) / (Spalten - 1)
   '#################################ZEICHNEN##################################'

   'Übergebene Parameter auswerten
   'Wenn eine Datei als Argument übergeben wurde..
   If Len(Command) <> 0 Then
      'Datei öffnen
      FileSystem_OpenDatei cdlg_dialog, pic_source, prg_fortschritt, Command
      'Modell refreshen
      Draw_Refresh pic_source
      'Bild von pic_source nach pic_target zoomen
      Draw_Zoom pic_source, pic_target
      'PictureBox refreshen
      Draw_Refresh pic_target
   End If
End Sub

Private Sub Form_Initialize()
'--------------------------------------------------------------------'
'| Prozedur beim Initialisieren der Form: Initialisierungen         |'
'--------------------------------------------------------------------'
   'XP-Style initialisieren
   Init_XPStyle
End Sub

Private Sub Form_Resize()
   Dim länge As Long
   Dim höhe As Long
   Dim kleiner As Long
   If frm_nightgraphix.Width < 9000 Then frm_nightgraphix.Width = 9000
   If frm_nightgraphix.Height < 6500 Then frm_nightgraphix.Height = 6500
   
   länge = frm_nightgraphix.Width / Screen.TwipsPerPixelX - 86
   höhe = frm_nightgraphix.Height / Screen.TwipsPerPixelY - 131
   
   kleiner = IIf(länge < höhe, länge, höhe)
   pic_rahmen.Height = kleiner
   pic_rahmen.Width = kleiner
   
   hsc_scrollen.Top = pic_rahmen.Top + pic_rahmen.Height
   hsc_scrollen.Width = pic_rahmen.Width
   
   vsc_scrollen.Left = pic_rahmen.Left + pic_rahmen.Width
   vsc_scrollen.Height = pic_rahmen.Height
   
   'Max-Wert der Scrollleisten einstellen
   vsc_scrollen.Max = frm_nightgraphix.pic_target.Height - frm_nightgraphix.pic_rahmen.Height
   'Max-Wert der Scrollleisten einstellen
   hsc_scrollen.Max = frm_nightgraphix.pic_target.Width - frm_nightgraphix.pic_rahmen.Width

   tlb_toolbar.Buttons(23).Width = 110 + frm_nightgraphix.Width / Screen.TwipsPerPixelX - 695
   
   prg_fortschritt.Top = frm_nightgraphix.Height / Screen.TwipsPerPixelY - 52
   prg_fortschritt.Left = frm_nightgraphix.Width / Screen.TwipsPerPixelX - 303
   
   If hsc_scrollen.Visible = False Then
      pic_target.Width = pic_rahmen.Width
      pic_target.Height = pic_rahmen.Height
      Draw_Zoom pic_source, pic_target
      If Connected2Hardware Or NGDemoModus Then
         'Offset-Pfeil neu setzen
         Draw_SetOffsetArrow pic_arrowshow, Offset / Spalten * 360 \ 1
      End If
   End If
   
   If pic_import.Visible Then
      'pic_import an die richtige Stelle schieben
      pic_import.Move (pic_rahmen.Width - pic_import.Width) / 2, (pic_rahmen.Height - pic_import.Height) / 2
      'pic_import halb transparent machen, wenn das Bild schon angeklickt wurde
      If Fetched Then Draw_AlphaBlend pic_target, pic_alphablend, pic_import
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'--------------------------------------------------------------------'
'| Prozedur beim Entladen der Form                                  |'
'--------------------------------------------------------------------'
'   'Wenn die Hardware nicht verbunden ist und nicht im Demo-Modus gearbeitet wird..
'   If (NGDemoModus = False) And (Connected2Hardware = False) Then
'      'Scrollrad-Hooken beenden
'      HookScrollWheel_Exit
'      'Nightvision ohne Nachfrage beenden
'      Exit Sub
'   End If
   
   'Wenn die Datei noch nicht gespeichert ist..
   If (SavedChanges = False) And Not ((NGDemoModus = False) And (Connected2Hardware = False)) Then
      'Messagebox aufrufen
      Select Case MsgBox(LngSpecStrings.AskToSave, vbYesNoCancel, LngSpecStrings.Attention & " - NightGraphix V1.0")
         'Wenn Abbrechen gedrückt wurde..
         Case vbCancel
            'NG nicht beenden..
            Cancel = 1
            'Prozedur beenden
            Exit Sub
         'Wenn Ja gedrückt wurde..
         Case vbYes
            'Klick auf Speichern simulieren
            mnu_speichern_Click
      End Select
   End If
   
   'Scrollrad-Hooken beenden
   HookScrollWheel_Exit
   
   'Brush-Objekte wieder freigeben
   Draw_ReleaseFillArea
   
   'Referenz zu WScript auf Nothing setzen
   FileSystem_DeleteWScriptReference

   'Aktuelle Einstellungen in der Registry speichern
   FileSystem_SaveSettings

   'Andere Fenster auch entladen
   Unload frm_splashscreen
   
   'Wenn der Port noch geöffnet ist..
   If msc_seriell.PortOpen = True Then
      '"X;" senden
      Communication_SendQuit msc_seriell
      'Port schließen
      ComPort_Close msc_seriell
   End If
   
   'Fenster verstecken
   Me.Hide
   'Fenster schließen
   Unload Me
   
   '###Verändern###'
   'End beendet alle Fenster, sollte man aber nicht nehmen,
   'lieber Unload für alle Fenster
   End
   '###Verändern###'
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'--------------------------------------------------------------------'
'| Prozedur beim ausdrücklichen Entladen der Form                   |'
'--------------------------------------------------------------------'
   'Log-Event speichern
   FileSystem_SaveLogFile "Log001"
   'Scrollrad-Hooken beenden
   HookScrollWheel_Exit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------'
'| Prozedur beim Drücken einer Taste in der Form                    |'
'--------------------------------------------------------------------'
   'Wenn Plus gedrückt wurde und das Pixeltool gewählt ist..
   If ((KeyCode = vbKeyAdd) Or (KeyCode = 187)) And (Tool = Pencil) And (Not pic_import.Visible) Then
      'Lupe anschalten
      tmr_lupe.Enabled = True
   End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'--------------------------------------------------------------------'
'| Prozedur beim Loslassen einer Taste in der Form                  |'
'--------------------------------------------------------------------'
   'Wenn Plus gedrückt wurde..
   If (KeyCode = vbKeyAdd) Or (KeyCode = 187) Then
      'Lupe anschalten
      tmr_lupe.Enabled = False
      'Lupe unsichtbar machen
      pic_lupe.Visible = False
      pic_lupe.Top = 1000
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'--------------------------------------------------------------------'
'| Prozedur zum Schreiben von Text bei Tastendruck                  |'
'--------------------------------------------------------------------'
   'Wenn Text-Werkzeug gewählt ist, dann schreiben
   If Tool = Text Then
      'Wenn BackSpace gedrückt wurde..
      If (KeyAscii = vbKeyBack) And (ZeichenAnzahl > 0) Then
         'Zeichenanzahl um 1 verringern
         ZeichenAnzahl = ZeichenAnzahl - 1
         'Zeichenspalte zurücksetzen
         ZeichenSpalte = ZeichenSpalte + (CInt(TextClockWise) + 0.5) * 2 * (UBound(OldLetterArrayRGB(ZeichenAnzahl).Data, 1) + 1)
         
         'Timer für Cursor ausschalten
         tmr_cursor.Enabled = False
         'Cursor ausschalten, wenn noch sichtbar
         If CursorVisible Then Draw_Cursor CursorPosition.x, CursorPosition.y, pic_source, False
         'Cursor-Koordinaten ändern
         CursorPosition.x = CursorPosition.x + (CInt(TextClockWise) + 0.5) * 2 * (UBound(OldLetterArrayRGB(ZeichenAnzahl).Data, 1) + 3)
         'Cursor initialisieren
         Draw_InitCursor CursorPosition.x, CursorPosition.y, pic_source
         'Timer für Cursor anschalten
         tmr_cursor.Enabled = True
         
         'Wenn es RGB-Version ist..
         If RGBVersion = True Then
            'Wenn Text im Uhrzeigersinn geschrieben wurde..
            If TextClockWise Then
               'Schleife durch alle Pixel des alten Hintergrundes
               For x = 0 To UBound(OldLetterArrayRGB(ZeichenAnzahl).Data, 1)
                  For y = 0 To UBound(OldLetterArrayRGB(ZeichenAnzahl).Data, 2)
                     'Zelle mit altem Inhalt füllen
                     Draw_FillCell (x + OldLetterArrayRGB(ZeichenAnzahl).Left + (Spalten - 1)) Mod Spalten + 1, y + OldLetterArrayRGB(ZeichenAnzahl).Top - 1, pic_source, RGB((OldLetterArrayRGB(ZeichenAnzahl).Data(x, y) \ 100) * 255, ((OldLetterArrayRGB(ZeichenAnzahl).Data(x, y) \ 10) Mod 10) * 255, (OldLetterArrayRGB(ZeichenAnzahl).Data(x, y) Mod 10) * 255), True
                  Next y
               Next x
            'Wenn Text gegen den Uhrzeigersinn geschrieben wurde..
            Else
               'Schleife durch alle Pixel des alten Hintergrundes
               For x = UBound(OldLetterArrayRGB(ZeichenAnzahl).Data, 1) To 0
                  For y = 0 To UBound(OldLetterArrayRGB(ZeichenAnzahl).Data, 2)
                     'Zelle mit altem Inhalt füllen
                     Draw_FillCell (x + OldLetterArrayRGB(ZeichenAnzahl).Left - UBound(OldLetterArrayRGB(ZeichenAnzahl).Data, 1) + (Spalten - 1)) Mod Spalten + 1, y + OldLetterArrayRGB(ZeichenAnzahl).Top - 1, pic_source, RGB((OldLetterArrayRGB(ZeichenAnzahl).Data(x, y) \ 100) * 255, ((OldLetterArrayRGB(ZeichenAnzahl).Data(x, y) \ 10) Mod 10) * 255, (OldLetterArrayRGB(ZeichenAnzahl).Data(x, y) Mod 10) * 255), True
                  Next y
               Next x
            End If
         'Wenn es SW-Version ist..
         Else
            'Wenn Text im Uhrzeigersinn geschrieben wurde..
            If TextClockWise Then
               'Schleife durch alle Pixel des alten Hintergrundes
               For x = 0 To UBound(OldLetterArraySW(ZeichenAnzahl).Data, 1)
                  For y = 0 To UBound(OldLetterArraySW(ZeichenAnzahl).Data, 2)
                     'Zelle mit altem Inhalt füllen
                     Draw_FillCell (x + OldLetterArraySW(ZeichenAnzahl).Left + (Spalten - 1)) Mod Spalten + 1, y + OldLetterArraySW(ZeichenAnzahl).Top - 1, pic_source, IIf(OldLetterArraySW(ZeichenAnzahl).Data(x, y), LEDColor, vbWhite), True
                  Next y
               Next x
            'Wenn Text gegen den Uhrzeigersinn geschrieben wurde..
            Else
               'Schleife durch alle Pixel des alten Hintergrundes
               For x = UBound(OldLetterArraySW(ZeichenAnzahl).Data, 1) To 0 Step -1
                  For y = 0 To UBound(OldLetterArraySW(ZeichenAnzahl).Data, 2)
                     'Zelle mit altem Inhalt füllen
                     Draw_FillCell (x + OldLetterArraySW(ZeichenAnzahl).Left - UBound(OldLetterArraySW(ZeichenAnzahl).Data, 1) + (Spalten - 1)) Mod Spalten + 1, y + OldLetterArraySW(ZeichenAnzahl).Top - 1, pic_source, IIf(OldLetterArraySW(ZeichenAnzahl).Data(x, y), LEDColor, vbWhite), True
                  Next y
               Next x
            End If
         End If
         
         'Eigenschaften in StatusPanel anzeigen
         NightGraphix_ShowPanelProperties
      'Bei anderen Tasten..
      ElseIf (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyReturn) Then
         'Wenn noch Zeichen hinpassen..
         If ZeichenSpalte < Spalten Then
            'Text schreiben
            Text_Write pic_source, pic_target, prg_fortschritt, Chr(KeyAscii), IIf(TextClockWise, CursorPosition.x + 1, CursorPosition.x - 1), CursorPosition.y - Letter.FontHeight + 2, LEDColor
            
            'Wenn die Textrichtung im Uhrzeigersinn ist..
            If TextClockWise = True Then
               'Zeichenspalte um 2 mehr als die Buchstabenbreite erhöhen
               ZeichenSpalte = ZeichenSpalte + Letter.Width + 2
            'Wenn die Textrichtung gegen den Uhrzeigersinn ist..
            Else
               'Zeichenspalte um 2 mehr als die Buchstabenbreite erniedrigen
               ZeichenSpalte = ZeichenSpalte - (Letter.Width + 2)
            End If
            
            '###########################TEST###########################
            tmr_cursor.Enabled = False
            Draw_Cursor CursorPosition.x, CursorPosition.y, pic_source, False
            
            'Wenn die Textrichtung im Uhrzeigersinn ist..
            If TextClockWise = True Then
               'Cursor nach links schieben
               CursorPosition.x = CursorPosition.x + Letter.Width + 2
            'Wenn die Textrichtung gegen den Uhrzeigersinn ist..
            Else
               'Cursor nach rechts schieben
               CursorPosition.x = CursorPosition.x - (Letter.Width + 2)
            End If
            
            'Cursor initialisieren
            Draw_InitCursor CursorPosition.x, CursorPosition.y, pic_source
            'Timer für Cursor anschalten
            tmr_cursor.Enabled = True
            '###########################TEST###########################
         End If
         'Eigenschaften in StatusPanel anzeigen
         NightGraphix_ShowPanelProperties
      End If
   End If
End Sub

Private Sub cmd_font_Click()
'--------------------------------------------------------------------'
'| Prozedur zum Öffnen des Font-Dialoges                            |'
'--------------------------------------------------------------------'
   'Font-Dialog öffnen
   Text_OpenFontDialog cdlg_dialog, pic_source, pic_target, tmr_cursor
   'Fokus auf PicBox setzen
   pic_target.SetFocus
   'Zeichenanzahl auf 0 setzen
   ZeichenAnzahl = 0
   'Zeichenspalte auf 0 setzen
   ZeichenSpalte = 0
End Sub

Private Sub cmd_tool_Click(Index As Integer)
'--------------------------------------------------------------------'
'| Prozedur zum Ändern des Werkzeuges                               |'
'--------------------------------------------------------------------'
   'Focus auf PictureBox setzen
   pic_target.SetFocus
   'Altes Tool speichern
   LastTool = Tool
   'Schriftauswahl-Button deaktivieren
   cmd_font.Enabled = False
   'Index überprüfen
   Select Case Index
      Case 0
         'Bleistift
         Tool = Pencil
      Case 1
         'Pinsel
         Tool = LEDCircle
      Case 2
         'Schriftauswahl-Button aktivieren
         cmd_font.Enabled = True
         'Cursor anschalten
         Tool = Text
         'Wenn das alte Tool das Farbwahl-Tool war
         If LastTool = ChooseColor Then
            'Cursor einschalten
            Draw_Cursor CursorPosition.x, CursorPosition.y, pic_source, True
            'Bild refreshen
            Draw_Zoom pic_source, pic_target
            'Cursor ist sichtbar
            CursorVisible = True
         End If
      'Wenn das Farbwahl-Tool gewählt wurde..
      Case 3
         'Aktuelles Tool: Farbwahl
         Tool = ChooseColor
         'Farbwahl-Picturebox sichtbar machen
         pic_farbwahl.Visible = True
         'Wenn RGB-Version..
         If RGBVersion Then
            'Schleife durch alle Optionbuttons
            For i = 0 To opt_farbe.UBound
               'Optionbutton aktivieren
               opt_farbe(i).Enabled = True
               'Farbe des Optionbuttons setzen
               opt_farbe(i).BackColor = Number2Color(i)
            Next i
            'Fokus auf aktuelle Farbe setzen
            opt_farbe(RGBColorIndex).Value = True
         'Wenn SW-Version..
         Else
            'Schleife durch die 4 ersten Optionbuttons
            For i = 0 To 3
               'Optionbutton aktivieren
               opt_farbe(i).Enabled = True
            Next i
            'Schleife durch die restlichen Optionbuttons
            For i = 4 To opt_farbe.UBound
               'Optionbutton deaktivieren
               opt_farbe(i).Enabled = False
               'Farbe auf grau setzen
               opt_farbe(i).BackColor = &H8000000F
            Next i
            'Fokus auf aktuelle Farbe setzen
            opt_farbe(SWColorIndex).Value = True
         End If
         'Fokus auf Picturebox setzen
         pic_farbwahl.SetFocus
      'Wenn "Bild importieren" gewählt wurde..
      Case 4
         'Aktuelles Tool: Bild importieren
         Tool = ImportPicture
         'Bild importieren-Dialog aufrufen
         mnu_importieren_Click
         'Fokus wieder auf Pencil setzen
         cmd_tool(0).Value = True
      'Wenn "Invertieren" gewählt wurde..
      Case 5
         'Aktuelles Tool: Invertieren
         Tool = InvertPicture
         'Invertieren ausführen
         Draw_InvertPicture pic_source
         'Statusanzeige aktualisieren
         NightGraphix_ShowPanelProperties
         'Fokus wieder auf Pencil setzen
         cmd_tool(0).Value = True
      'Wenn "Löschen" gewählt wurde..
      Case 6
         'Aktuelles Tool: Löschen
         Tool = ClearPicture
         'Arrays löschen
         Erase Array_SW
         Erase Array_Red
         Erase Array_Green
         Erase Array_Blue
         'Arrays neu initialisieren
         ReDim Array_SW(1 To Spalten, 1 To Leds)
         ReDim Array_Red(1 To Spalten, 1 To Leds)
         ReDim Array_Green(1 To Spalten, 1 To Leds)
         ReDim Array_Blue(1 To Spalten, 1 To Leds)
         'Modell neu zeichnen
         Draw_Redraw pic_source, pic_target
         'Cursor-Array auf weiß stellen
         Draw_InitCursor CursorPosition.x, CursorPosition.y, pic_source
         'Eigenschaften in StatusBar anzeigen
         NightGraphix_ShowPanelProperties
         'Fokus wieder auf Pencil setzen
         cmd_tool(0).Value = True
End Select
   
   'Timer bei allen Werkzeugen außer Text ausschalten
   If (Tool <> Text) Then
      tmr_cursor.Enabled = False
      'Cursor ausschalten, wenn noch sichtbar
      If CursorVisible = True Then
         'Cursor ausschalten
         Draw_Cursor CursorPosition.x, CursorPosition.y, pic_source, False
         'Bild refreshen
         Draw_Zoom pic_source, pic_target
         'Cursor soll am Anfang wieder gezeichnet werden
         CursorVisible = False
      End If
   End If
End Sub

Private Sub cmd_tool_GotFocus(Index As Integer)
'--------------------------------------------------------------------'
'| Prozedur beim Setzten des Fokus auf eine Tool-Button             |'
'--------------------------------------------------------------------'
   'Focus auf Picturebox setzen
   '(Damit Optionbutton keinen schwarzen Rahmen hat)
   pic_target.SetFocus
End Sub

Private Sub mnu_löschen_Click()
'--------------------------------------------------------------------'
'| Prozedur zum Löschen des Modells                                 |'
'--------------------------------------------------------------------'
   'Modell neu zeichnen
   Draw_Redraw pic_source, pic_target
End Sub

Private Sub mnu_NGinfo_Click()
'--------------------------------------------------------------------'
'| Prozedur zum Anzeigen der Info über NightGraphix                 |'
'--------------------------------------------------------------------'

   'MessageBox mit Info und Copyright anzeigen
   MsgBox "NightGraphiX Software" & _
        vbCrLf & _
        "Version " & App.Major & "." & App.Minor & "." & App.Revision & _
        vbCrLf & _
        vbCrLf & _
        "David Herrmann, Copyright (C) 2007" & _
        vbCrLf & _
        vbCrLf & _
        "Weitere Information unter" & _
        vbCrLf & _
        "http://www.NightGraphiX.de", vbOKOnly, "Über NightGraphiX"
End Sub

Private Sub msc_comport_OnComm()
'--------------------------------------------------------------------'
'| Prozedur bei einem Ereignis von msc_comport                      |'
'--------------------------------------------------------------------'
   'Ereignis auswerten
   Select Case msc_comport.CommEvent
      'Wenn Daten empfangen wurden..
      Case comEvReceive
         'Daten auswerten
         Communication_HandleData msc_comport.Input, msc_comport, pic_source
         'Daten löschen
         msc_comport.InBufferCount = 0
   End Select
End Sub

Private Sub msc_seriell_OnComm()
'--------------------------------------------------------------------'
'| Prozedur beim Empfangen von Daten von der Hardware               |'
'--------------------------------------------------------------------'
   Select Case msc_seriell.CommEvent
      'Wenn Daten empfangen wurden..
      Case comEvReceive
         'Daten auswerten
         Communication_HandleData msc_seriell.Input, msc_seriell, pic_source
   End Select
End Sub

Private Sub opt_farbe_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf einen Farb-Optionbutton                |'
'--------------------------------------------------------------------'
   'Farbe auf gewählte Farbe setzen
   LEDColor = Number2Color(Index)
   'Wenn SW-Version..
   If Not RGBVersion Then
      'Index in Variable speichern
      SWColorIndex = Index
      'Wenn mit Hardware verbunden..
      If Connected2Hardware Then
         'Wenn das obere Blatt angeschlossen ist..
         If (Not NGRotationSystemLeft) = (NGRotationMCRight) Then
            'Aktuelle Farbe als obere Farbe setzen
            NGTopLEDColor = LEDColor
         'Wenn das untere Blatt angeschlossen ist..
         Else
            'Aktuelle Farbe als untere Farbe setzen
            NGBottomLEDColor = LEDColor
         End If
      End If
      'Wenn Demo-Modus ausgeführt wird oder Hardware verbunden ist..
      If NGDemoModus Or Connected2Hardware Then
         'Alle gefüllten Zellen mit der neuen Farbe füllen
         Draw_RefreshNewColor frm_nightgraphix.pic_source
         'Bild refreshen
         Draw_Zoom frm_nightgraphix.pic_source, frm_nightgraphix.pic_target
      End If
   'Wenn RGB-Version
   Else
      'Index in Variable speichern
      RGBColorIndex = Index
   End If
   'Altes Tool wieder anwählen
   cmd_tool(LastTool).Value = True
   'Farbwahl-Picturebox unsichtbar machen
   pic_farbwahl.Visible = False
   'Aktuelle Änderungen in Registry speichern
   FileSystem_SaveSettings
End Sub

Private Sub pic_arrowshow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur beim Mausklick in pic_arrowshow                         |'
'--------------------------------------------------------------------'
   'Mauszustand in Variable schreiben
   MouseButton = Button
   'Timer für Offset-Pfeil ausschalten
   tmr_offsetarrow.Enabled = True
End Sub

Private Sub pic_arrowshow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur wird beim Loslassen einer Maustaste ausgeführt          |'
'--------------------------------------------------------------------'
   'Zustand in Variable speichern
   MouseButton = MouseButton - Button
   'Timer für Offset-Pfeil ausschalten
   tmr_offsetarrow.Enabled = False
End Sub

Private Sub pic_farbwahl_LostFocus()
'--------------------------------------------------------------------'
'| Prozedur, wenn pic_farbwahl den Fokus verliert                   |'
'--------------------------------------------------------------------'
   'Altes Tool wieder anwählen
   cmd_tool(LastTool).Value = True
   'Farbwahl unsichtbar machen
   pic_farbwahl.Visible = False
End Sub

Private Sub pic_lupe_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf pic_lupe                               |'
'--------------------------------------------------------------------'
   'Mausposition ermitteln
   GetCursorPos CurPos
   'Bildschirm-Koordinaten in Client-Daten umrechnen
   ScreenToClient pic_target.hwnd, CurPos
   'Klick an pic_target weitergeben
   pic_target_MouseDown Button, Shift, CSng(CurPos.x), CSng(CurPos.y)
End Sub

Private Sub pic_target_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur zum Füllen der Felder bei Mausklick                     |'
'--------------------------------------------------------------------'
   
   'Wenn Hardware nicht verbunden ist und Demomodus nicht gestartet ist..
   If (Not NGDemoModus) And (Not Connected2Hardware) Then Exit Sub
   
   'Welches Werkzeug ist aktiv ?
   Select Case Tool
      'Stift
      Case Pencil
         'Wenn Zeichnen noch gesperrt ist, Prozedur beenden
         If LockDraw Then Exit Sub
         
         'Füllen ausführen
         Draw_Click pic_source, pic_target, x, y
         
         'Picturebox refreshen
         Draw_Refresh pic_source
      
         'Zoomen
         'Draw_Zoom pic_source, pic_target
         Draw_ZoomArea x - 70 / ZoomX, y - 70 / ZoomY, 140 / ZoomX, 140 / ZoomY, pic_source, pic_target
         
         'Eigenschaften in StatusPanel anzeigen
         NightGraphix_ShowPanelProperties
      'Kreis
      Case LEDCircle
         'Wenn Zeichnen noch gesperrt ist, Prozedur beenden
         If LockDraw Then Exit Sub
         
         'LED und Spalte ausrechnen
         Maths_Long2Cell x * ZoomX, y * ZoomY
         'Wenn Zelle eine andere Farbe als weiß ist..
         If GetPixel(pic_source.hdc, x * ZoomX, y * ZoomY) <> vbWhite Then
            'Kreis weiß füllen
            Draw_Circle pic_source, pic_target, HöheFeld + 1, vbWhite
         'Wenn Zelle weiß ist..
         Else
            'Kreis mit aktueller Farbe füllen
            Draw_Circle pic_source, pic_target, HöheFeld + 1, LEDColor
         End If
         
         'Eigenschaften in StatusPanel anzeigen
         NightGraphix_ShowPanelProperties

         'Änderungen wurden nicht gespeichert
         FileSystem_SavedChanges False
      'Text
      Case Text
         'Wenn rechte Maustaste gedrückt wurde..
         If Button = vbRightButton Then
            'Textrichtung im Uhrzeigersinn
            TextClockWise = True
         'Wenn linke Maustaste gedrückt wurde..
         Else
            'Textrichtung gegen den Uhrzeigersinn
            TextClockWise = False
         End If
         'Cursor ausschalten, wenn noch sichtbar
         If CursorVisible = True Then
            'Cursor ist nicht sichtbar
            CursorVisible = False
            'Cursor ausschalten
            Draw_Cursor CursorPosition.x, CursorPosition.y, pic_source, False
            'Bild refreshen
            Draw_Zoom pic_source, pic_target
         End If
         'Zeichenanzahl auf 0 setzen
         ZeichenSpalte = 0
         'Zeichenanzahl auf 0 setzen
         ZeichenAnzahl = 0
         'CursorPosition beschreiben
         Maths_Long2Cell x * ZoomX, y * ZoomY
         'Wenn im Uhrzeigersinn geschrieben wird
         If TextClockWise Then
            'Y-Position einstellen
            CursorPosition.y = HöheFeld + 1
         'Wenn gegen den Uhrzeigersinn geschrieben wird
         Else
            'Y-Position einstellen
            CursorPosition.y = HöheFeld + Letter.FontHeight
         End If
         'X-Position einstellen
         CursorPosition.x = BreiteFeld + 1
         'Cursor initialisier en
         Draw_InitCursor CursorPosition.x, CursorPosition.y, pic_source
         'Timer für Cursor anschalten
         tmr_cursor.Enabled = True
         'Änderungen wurden nicht gespeichert
         FileSystem_SavedChanges False
   End Select
End Sub

Private Sub mnu_importieren_Click()
'--------------------------------------------------------------------'
'| Prozedur zum Importieren von Bildern                             |'
'--------------------------------------------------------------------'
   'CommonDialog öffnen und Pfad in Variable schreiben
   BildPfad = FileSystem_OpenDialog(cdlg_dialog, Filter_Graphik, "Open")
   'Wurde Abbrechen gewählt ?
   If BildPfad = "" Then Exit Sub
   'Bild in PictureBox laden
   pic_bild.Picture = LoadPicture(BildPfad)
   
   'Inhalt in pic_bildgroß löschen
   pic_bildgroß.Cls
   
   'Größe der pic_alphablend Picbox anpassen
   pic_alphablend.Width = pic_import.Width
   pic_alphablend.Height = pic_import.Height
   
   'Bild von pic_bild in pic_alphablend zoomen
   Draw_Zoom pic_bild, pic_alphablend
   
   'Bild refreshen
   Draw_Refresh pic_bild
   
   'Bild von pic_bild in pic_import übertragen
   Draw_Zoom pic_bild, pic_import
   
   'Button sichtbar machen
   cmd_finishimport.Visible = True

   'pic_import an die richtige Stelle schieben
   pic_import.Move (pic_rahmen.Width - pic_import.Width) / 2, (pic_rahmen.Height - pic_import.Height) / 2
   'pic_import sichtbar machen
   pic_import.Visible = True
   
   'Button "Fertig" sichtbar machen
   cmd_finishimport.Visible = True
   
   'Button "Stop" sichtbar machen
   cmd_cancelimport.Visible = True
   
'   'Bild in große PicBox zoomen
'   Draw_Zoom pic_bild, pic_bildgroß
'
'   'Statusbar-Anzeige ändern
'   stb_statusbar.Panels(2).Text = FileSystem_Path2File(BildPfad) & " wird importiert ..."
'
'   'Bild ins Modell füllen
'   Draw_ImportPicture pic_bildgroß, pic_source, pic_target, prg_fortschritt
'
'   'Statusbar-Anzeige zurücksetzen
'   stb_statusbar.Panels(2).Text = ""
End Sub

Private Sub mnu_öffnen_Click()
'--------------------------------------------------------------------'
'| Prozedur zum Öffnen eines Modells                                |'
'--------------------------------------------------------------------'
   'Datei öffnen
   FileSystem_OpenDatei cdlg_dialog, pic_source, prg_fortschritt
   
   'Modell neu zeichnen
   Draw_Zoom pic_source, pic_target
End Sub

Private Sub mnu_speichern_Click()
'--------------------------------------------------------------------'
'| Prozedur zum Speichern des Modells                               |'
'--------------------------------------------------------------------'
   'Datei speichern
   FileSystem_CreateDatei cdlg_dialog, prg_fortschritt, False
End Sub

Private Sub mnu_speichernunter_Click()
'--------------------------------------------------------------------'
'| Prozedur zum Speichern des Modells an bestimmtem Ort             |'
'--------------------------------------------------------------------'
   'Datei Speichern unter...
   FileSystem_CreateDatei cdlg_dialog, prg_fortschritt, True
End Sub

Private Sub hsc_scrollen_Change()
'--------------------------------------------------------------------'
'| Prozedur zum Verschieben des Modells                             |'
'--------------------------------------------------------------------'
   'pic_target verschieben
   pic_target.Left = -hsc_scrollen.Value
   'Wenn Bildimport-Picturebox sichtbar ist..
   If frm_nightgraphix.pic_import.Visible Then
      'Inhalt der Picbox neu zeichnen
      Draw_AlphaBlend frm_nightgraphix.pic_target, frm_nightgraphix.pic_alphablend, frm_nightgraphix.pic_import
   End If
End Sub

Private Sub hsc_scrollen_Scroll()
'--------------------------------------------------------------------'
'| Prozedur zum Verschieben des Modells                             |'
'--------------------------------------------------------------------'
   'pic_target verschieben
   pic_target.Left = -hsc_scrollen.Value
   'Wenn Bildimport-Picturebox sichtbar ist..
   If frm_nightgraphix.pic_import.Visible Then
      'Inhalt der Picbox neu zeichnen
      Draw_AlphaBlend frm_nightgraphix.pic_target, frm_nightgraphix.pic_alphablend, frm_nightgraphix.pic_import
   End If
End Sub

Public Sub tlb_toolbar_ButtonClick(ByVal Button As ComctlLib.Button)
'--------------------------------------------------------------------'
'| Prozedur der ToolBarButtons                                      |'
'--------------------------------------------------------------------'
   'Welcher Button wurde gedrückt ?
   Select Case Button.Index
      'Neue Datei
      Case 1
         'Wenn die Datei noch nicht gespeichert ist..
         If SavedChanges = False Then
            'Messagebox aufrufen
            Select Case MsgBox(LngSpecStrings.AskToSave, vbYesNoCancel, LngSpecStrings.Attention & " - NightGraphix V1.0")
               'Wenn Abbrechen gedrückt wurde..
               Case vbCancel
                  'Prozedur beenden
                  Exit Sub
               'Wenn Ja gedrückt wurde..
               Case vbYes
                  'Klick auf Speichern simulieren
                  mnu_speichern_Click
            End Select
         End If
         'Datei wurde nicht gespeichert..
         SavedChanges = False
         'Dateiname zurücksetzen
         DateiName = "Untitled"
         'Dateipfad zurücksetzen
         DateiPfad = ""
         'Beschriftung der Form zurücksetzen
         Me.Caption = DateiName & " - NightGraphix V1.0"
         'Anzahl der geschriebenen Zeichen zurücksetzen
         ZeichenAnzahl = 0
         'Array und Modell löschen
         Erase Array_SW
         'Array neu initialisieren
         ReDim Array_SW(1 To Spalten, 1 To Leds)
         'Modell neu zeichnen
         Draw_Redraw pic_source, pic_target
         'Cursor-Array auf weiß stellen
         Draw_InitCursor CursorPosition.x, CursorPosition.y, pic_source
      'Zoom +
      Case 9
         'Wenn noch nicht näher gezoomt ist, näher zoomen
         If Button.Enabled = True Then
            Draw_ZoomModell 0
            'Offset-Pfeil neu setzen
            Draw_SetOffsetArrow pic_arrowshow, Offset / Spalten * 360 \ 1
         End If
      'Zoom -
      Case 10
         'Wenn noch nicht weg gezoomt ist, weg zoomen
         If Button.Enabled = True Then
            Draw_ZoomModell 1
            'Offset-Pfeil neu setzen
            Draw_SetOffsetArrow pic_arrowshow, Offset / Spalten * 360 \ 1
         End If
      'Datei öffnen
      Case 2
         'Datei öffnen
         mnu_öffnen_Click
      'Datei speichern
      Case 4
         'Datei speichern
         mnu_speichern_Click
      'Datei speichern unter
      Case 5
         'Datei speichern unter
         mnu_speichernunter_Click
'      Case "Bild importieren"
'         'Bild importieren
'         mnu_importieren_Click
'      Case "Neu zeichnen"
'         'Bild neu laden (von pic_source nach pic_target)
'         Draw_Zoom pic_source, pic_target
      'Verbinden
      Case 15
         'Trennen-Button sichtbar machen, Verbinden unsichtbar machen
         tlb_toolbar.Buttons(15).Visible = False
         tlb_toolbar.Buttons(16).Visible = True
         'Splashscreen laden
         Load frm_splashscreen
         'Demo-Button, Verbinden-Button und Software-Optionen deaktivieren
         tlb_toolbar.Buttons(15).Enabled = False
         tlb_toolbar.Buttons(17).Enabled = False
         tlb_toolbar.Buttons(20).Enabled = False
         'Splashscreen anzeigen
         frm_splashscreen.Show vbNormal, Me
         'Letzen geöffneten Port auf 0 setzen
         COMPortOpened_Last = 0
         'Aktuell geöffneten Port auf 0 setzen
         COMPortOpened_Now = 0
         'Timer zum Suchen der Hardware aktivieren
         tmr_comports.Enabled = True
         'Form Caption ändern
         frm_splashscreen.lbl_beschriftung.Caption = LngSpecStrings.LookingForHardware
      'Trennen
      Case 16
         'Dateipfad zurücksetzen
         DateiPfad = ""
         'Trennen-Button unsichtbar machen, Verbinden sichtbar machen
         tlb_toolbar.Buttons(16).Visible = False
         tlb_toolbar.Buttons(15).Visible = True
         'Timer für den Cursor beim Schreiben anhalten
         tmr_cursor.Enabled = False
         'Cursor ist nicht mehr sichtbar
         CursorVisible = False
         'Dem MC signalisieren, dass Verbindung beendet wird
         Communication_SendQuit msc_seriell
         'Port schließen
         ComPort_Close msc_seriell
         'RTreshold von msc_comport wieder auf 6 setzen
         msc_comport.RThreshold = 6 'Hier stand eine 18
         'Zum Startscreen (graues Feld) zurückkehren
         NightGraphix_ReInitialize
      'Demo starten
      Case 17
         'LED-Anzahl setzen
         Leds = (NGDemoSize + 2) * 8
         'Spaltenanzahl setzen
         Spalten = 512
         'RGB-Version
         RGBVersion = CBool(NGDemoVersion)
         'Demo starten
         NightGraphix_StartDemo
      'Demo beenden
      Case 18
         'Dateipfad zurücksetzen
         DateiPfad = ""
         'Timer für den Cursor beim Schreiben anhalten
         tmr_cursor.Enabled = False
         'Cursor ist nicht mehr sichtbar
         CursorVisible = False
         'Wenn die Datei noch nicht gespeichert ist..
         If SavedChanges = False Then
            'Messagebox aufrufen
            Select Case MsgBox(LngSpecStrings.AskToSave, vbYesNoCancel, LngSpecStrings.Attention & " - NightGraphix V1.0")
               'Wenn Abbrechen gedrückt wurde..
               Case vbCancel
                  'Prozedur beenden
                  Exit Sub
               'Wenn Ja gedrückt wurde..
               Case vbYes
                  'Klick auf Speichern simulieren
                  mnu_speichern_Click
            End Select
         End If
         'Modell löschen, Startbildschirm (graues Feld) wiederherstellen
         NightGraphix_ReInitialize
      'EEPROM beschreiben
      Case 12
         'Fenster entladen
         Unload frm_writereadscreen
         'Klick auf Pixel-Tool simulieren
         cmd_tool_Click 0
         cmd_tool(0).Value = True
         'EEPROM wird beschrieben
         ReadOrWrite = WriteEEPROM
         'frm_writescreen laden
         Load frm_writereadscreen
         'frm_writescreen anzeigen
         frm_writereadscreen.Show
      'EEPROM auslesen
      Case 13
         'Fenster entladen
         Unload frm_writereadscreen
         'Klick auf Pixel-Tool simulieren
         cmd_tool_Click 0
         cmd_tool(0).Value = True
         'EEPROM wird ausgelesen
         ReadOrWrite = ReadEEPROM
         'Auf das Pixel-Tool umschalten
         cmd_tool_Click (eTool.Pencil)
         'Pixel-Tool Button aktivieren
         cmd_tool(eTool.Pencil).Value = True
         'frm_writescreen laden
         Load frm_writereadscreen
         'frm_writescreen anzeigen
         frm_writereadscreen.Show
         'Modell neu zeichnen
         Draw_Zoom pic_source, pic_target
      'Info über NG anzeigen
      Case 22
         'Fenster entladen
         Unload frm_about
         'Klick auf Pixel-Tool simulieren
         cmd_tool_Click 0
         cmd_tool(0).Value = True
         'About-Box anzeigen
         'mnu_NGinfo_Click
         Load frm_about
         frm_about.Show vbModal, Me
      'Softwareoptionen
      Case 20
         'Log: Software-Optionen werden aufgerufen
         LogEvent "Softwareoptionen werden aufgerufen"
         'Fenster entladen
         Unload frm_optionssoftware
         'Softwareoptionen-Fenster anzeigen
         Load frm_optionssoftware
         frm_optionssoftware.Show vbModal, Me
      'Hardwareoptionen
      Case 21
         'Fenster entladen
         Unload frm_optionshardware
         'Hardwareoptionen-Fenster anzeigen
         Load frm_optionshardware
         frm_optionshardware.Show vbModal, Me
      'NG beenden
      Case 24
         'NG beenden
         Unload Me
   End Select
End Sub

Private Sub tmr_comports_Timer()
'--------------------------------------------------------------------'
'| Prozedur des Timer-Events von tmr_comport                        |'
'--------------------------------------------------------------------'
   'Wenn NG noch nicht mit der Hardware verbunden ist..
   If (Not Connected2Hardware) And (Not CancelConnect) Then
      'Den zuletzt geöffneten ComPort in Variable schreiben
      COMPortOpened_Last = COMPortOpened_Now
      'Variable i setzen
      i = COMPortOpened_Last + 1

      'Schleife, bis der ComPort i verfügbar ist
      Do Until ComPort_Available(msc_comport, i) And NGComportSearch(i)
         'Variable i um 1 inkrementieren
         i = i + 1
         'Anderen Events auch mal Zeit lassen
         DoEvents
         'Wenn i größer als 16 ist (höchstmöglicher ComPort)
         If i > 16 Then
            'Splashscreen schließen
            Unload frm_splashscreen
            'Trennen-Button unsichtbar machen, Verbinden sichtbar machen
            tlb_toolbar.Buttons(16).Visible = False
            tlb_toolbar.Buttons(15).Visible = True
            'Timer ausschalten
            tmr_comports.Enabled = False
            'Status in Statusbar anzeigen
            stb_statusbar.Panels(2).Text = LngSpecStrings.NotFoundNGXHardware
            'Wenn im Demomodus..
            If NGDemoModus Then
               'Demo-beenden wieder aktivieren
               tlb_toolbar.Buttons(17).Enabled = True
               'Verbinden wieder aktivieren
               tlb_toolbar.Buttons(15).Enabled = True
               'Prozedur beenden
               Exit Sub
            End If
            'Toolbarbuttons aktivieren/deaktivieren
            NightGraphix_DisableToolButtons
            'Prozedur beenden
            Exit Sub
         End If
      Loop

      'Aktuellen Port im Splashscreen anzeigen
      frm_splashscreen.lbl_beschriftung.Caption = LngSpecStrings.PortComX & i

      'Aktuellen Port setzen
      COMPortOpened_Now = i
      'Port öffnen
      ComPort_Open msc_comport, COMPortOpened_Now
      'Ein "P:;" zum Verbinden senden
      msc_comport.Output = "P:;"
   'Wenn NG schon mit Hardware verbunden ist oder das Verbinden abgebrochen werden soll..
   Else
      'Splashscreen schließen
      Unload frm_splashscreen
      'Timer ausschalten
      tmr_comports.Enabled = False
      'Wenn das Verbinden abgebrochen werden soll..
      If CancelConnect Then
         'MessageBox anzeigen
         MsgBox "Anzahl der Demo-LEDs stimmt nicht mit der Hardware überein - Verbinden wurde abgebrochen!", vbOKOnly & vbCritical, "NightGraphiX V1.0 - Fehler"
         'Demo-beenden wieder aktivieren
         tlb_toolbar.Buttons(17).Enabled = True
         'Verbinden wieder aktivieren
         tlb_toolbar.Buttons(15).Enabled = True
         'Prozedur beenden
         Exit Sub
      End If
      'Trennen wieder aktivieren
      tlb_toolbar.Buttons(16).Enabled = True
      'NG nicht mehr im Demomodus
      NGDemoModus = False
      'Status in Statusbar anzeigen
      stb_statusbar.Panels(2).Text = "Verbunden mit: " & "NG-" & CStr(Leds) & ", " & IIf((Not NGRotationSystemLeft) = (NGRotationMCRight), "Oben", "Unten") & IIf(RGBVersion, ", RGB", "")
      'COM-Port schließen
      ComPort_Close msc_comport
      'Hub setzen
      Hub = (NGRotorSizeArray((Leds - 16) / 8) - 2 * Leds * 10 - 2 * 35) / 20
      'Variablen initialisieren
      Init_Variables pic_source
      'Namen der Form initialisieren
      Init_Form Me
      'Font-PictureBox initialisieren
      Init_FontPictureBox pic_letter
      'DateiHeader initialisieren
      Init_FileHeader
      'Arrays initialisieren
      Init_Arrays
      'Zeichenfläche initialisieren
      Draw_Constructor pic_source, pic_target, pic_arrowshow
      'Buttons aktivieren
      NightGraphix_EnableToolButtons
      'Demo-Button deaktivieren
      tlb_toolbar.Buttons(18).Enabled = False
      'COM-Port öffnen
      ComPort_Open msc_seriell, NGCOMPort
   End If
End Sub

Private Sub tmr_cursor_Timer()
'--------------------------------------------------------------------'
'| Prozedur zum Anzeigen/Löschen des Cursors                        |'
'--------------------------------------------------------------------'
   
   'Wenn Cursor sichtbar ist..
   If CursorVisible Then
      'Cursor ist nicht sichtbar
      CursorVisible = False
      'Cursor löschen
      Draw_Cursor CursorPosition.x, CursorPosition.y, pic_source, False
   'Wenn Cursor nicht sichtbar ist..
   Else
      'Cursor ist sichtbar
      CursorVisible = True
      'Cursor anzeigen
      Draw_Cursor CursorPosition.x, CursorPosition.y, pic_source, True
   End If
   
   'Bild refreshen
   Draw_Zoom pic_source, pic_target
End Sub

Private Sub tmr_lupe_Timer()
'--------------------------------------------------------------------'
'| Prozedur zum Benutzen der Lupe                                   |'
'--------------------------------------------------------------------'
   'Mausposition ermitteln
   GetCursorPos CurPos
   'Bildschirm-Koordinaten in Client-Daten umrechnen
   ScreenToClient pic_target.hwnd, CurPos
   'Pic_lupe an richtige Position setzen
   pic_lupe.Left = CurPos.x - pic_lupe.Width / 2 + pic_target.Left
   pic_lupe.Top = CurPos.y - pic_lupe.Height / 2 + pic_target.Top
   'Maske zeichnen: Weißer Kreis mit schwarzem Hintergrund
   pic_lupemask.FillColor = vbWhite
   pic_lupemask.BackColor = vbBlack
   pic_lupemask.Circle (129 / 2 - 1, 129 / 2 - 1), 129 / 2 - 1, vbWhite
   'Vergrößerter Bereich von pic_source nach pic_lupebitmap blitten
   BitBlt pic_lupebitmap.hdc, 0, 0, pic_lupe.Width, pic_lupe.Height, pic_source.hdc, (pic_lupe.Left - pic_target.Left + pic_lupe.Width \ 2) * ZoomX - pic_lupe.Width \ 2, (pic_lupe.Top - pic_target.Top + pic_lupe.Height \ 2) * ZoomY - pic_lupe.Height \ 2, BIT_COPY
   'Maske auf pic_lupebitmap blitten
   BitBlt pic_lupebitmap.hdc, 0, 0, pic_lupe.Width, pic_lupe.Height, pic_lupemask.hdc, 0, 0, BIT_AND
   'Maske zeichnen: Schwarzer Kreis mit weißem Hintergrund
   pic_lupemask.FillColor = vbBlack
   pic_lupemask.BackColor = vbWhite
   pic_lupemask.Circle (129 / 2 - 1, 129 / 2 - 1), 129 / 2 - 1, vbYellow
   pic_lupemask.Circle (129 / 2 - 1, 129 / 2 - 1), 127 / 2 - 1, vbBlack
   'pic_lupemask und pic_lupebitmap neu zeichnen
   pic_lupemask.Refresh
   pic_lupebitmap.Refresh
   'Rand-Bereich von pic_target nach pic_lupeshow blitten
   BitBlt pic_lupeshow.hdc, 0, 0, pic_lupe.Width, pic_lupe.Height, pic_target.hdc, pic_lupe.Left - pic_target.Left, pic_lupe.Top - pic_target.Top, BIT_COPY
   'Maske auf pic_lupeshow blitten
   BitBlt pic_lupeshow.hdc, 0, 0, pic_lupe.Width, pic_lupe.Height, pic_lupemask.hdc, 0, 0, BIT_AND
   'Bild von pic_lupebitmap nach pic_lupeshow blitten
   BitBlt pic_lupeshow.hdc, 0, 0, pic_lupe.Width, pic_lupe.Height, pic_lupebitmap.hdc, 0, 0, BIT_Invert
   'pic_lupeshow neu zeichnen
   pic_lupeshow.Refresh
   'gebuffertes Bild auf pic_lupe blitten
   BitBlt pic_lupe.hdc, 0, 0, pic_lupe.Width, pic_lupe.Height, pic_lupeshow.hdc, 0, 0, BIT_COPY
   'pic_lupe neu zeichnen
   pic_lupe.Refresh
   'Wenn die Lupe noch nicht sichbar ist, sichtbar machen
   If Not pic_lupe.Visible Then pic_lupe.Visible = True
End Sub

Private Sub tmr_offsetarrow_Timer()
'--------------------------------------------------------------------'
'| Prozedur zum Setzen des Offset-Cursors                           |'
'--------------------------------------------------------------------'
   'Wenn linke Maustaste gedrückt ist..
   If MouseButton = vbLeftButton Then
      'Mausposition ermitteln
      GetCursorPos CurPos
      'Bildschirm-Koordinaten in Client-Daten umrechnen
      ScreenToClient pic_target.hwnd, CurPos
      'In welchem Winkel befindet sich die Maus ?
      RotateDegree = Maths_GetWinkel(CurPos.x * ZoomX, CurPos.y * ZoomY)
      'Offset-Cursor drehen
      Draw_SetOffsetArrow pic_arrowshow, RotateDegree
      'Eigenschaften in der Statusbar anzeigen
      NightGraphix_ShowPanelProperties
      'Offset-Cursor an die richtige Position setzen
      Draw_OffsetCursor RotateDegree
   End If
End Sub

Private Sub vsc_scrollen_Change()
'--------------------------------------------------------------------'
'| Prozedur zum Verschieben des Modells                             |'
'--------------------------------------------------------------------'
   'pic_target verschieben
   pic_target.Top = -vsc_scrollen.Value
   'Wenn Bildimport-Picturebox sichtbar ist..
   If frm_nightgraphix.pic_import.Visible Then
      'Inhalt der Picbox neu zeichnen
      Draw_AlphaBlend frm_nightgraphix.pic_target, frm_nightgraphix.pic_alphablend, frm_nightgraphix.pic_import
   End If
End Sub

Private Sub vsc_scrollen_Scroll()
'--------------------------------------------------------------------'
'| Prozedur zum Verschieben des Modells                             |'
'--------------------------------------------------------------------'
   'pic_target verschieben
   pic_target.Top = -vsc_scrollen.Value
   'Wenn Bildimport-Picturebox sichtbar ist..
   If frm_nightgraphix.pic_import.Visible Then
      'Inhalt der Picbox neu zeichnen
      Draw_AlphaBlend frm_nightgraphix.pic_target, frm_nightgraphix.pic_alphablend, frm_nightgraphix.pic_import
   End If
End Sub

Private Sub pic_import_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur beim Drücken der Maustaste auf pic_import               |'
'--------------------------------------------------------------------'
   'Wenn der Fertigstellen-Button unsichtbar ist, Prozedur beenden
   If Not cmd_finishimport.Visible Then Exit Sub
   'Picbox soll bei Mausbewegung verschoben werden
   MoveFlag = True
   'Startposition der Maus
   StartX = x
   Starty = y
End Sub

Private Sub pic_import_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur beim Loslassen der Maustaste auf pic_import             |'
'--------------------------------------------------------------------'
   'Picbox soll bei Mausbewegung NICHT verschoben werden
  MoveFlag = False
   'pic_import halb transparent machen
   Draw_AlphaBlend pic_target, pic_alphablend, pic_import
End Sub

Private Sub pic_import_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur beim Bewegen der Maus in pic_import                     |'
'--------------------------------------------------------------------'
   Dim xD&, yD&, TPX&, TPY&
   Static Doing As Boolean
   If Doing Or Not Fetched Then Exit Sub
      DoEvents
      Doing = True
      If MoveFlag Then
         TPX = 1 'Screen.TwipsPerPixelX
         TPY = 1 'Screen.TwipsPerPixelY
         
         xD = pic_import.Left + (x - StartX)
         If xD + pic_import.Width > LimX2 * TPX Then
         xD = LimX2 * TPX - pic_import.Width
      ElseIf xD < LimX1 * TPX Then
         xD = LimX1 * TPX
      End If
      DoEvents
      yD = pic_import.Top + (y - Starty)
      If yD + pic_import.Height > LimY2 * TPY Then
         yD = LimY2 * TPY - pic_import.Height
      ElseIf yD < LimY1 * TPY Then
         yD = LimY1 * TPY
      End If
      
      pic_import.Left = xD
      pic_import.Top = yD
      Call DrawPics
   End If
   Doing = False
   
   'pic_import halb transparent machen
   If MoveFlag Then Draw_AlphaBlend pic_target, pic_alphablend, pic_import
End Sub

Private Sub pic_resizeimport_MouseDown(Index As Integer, Button As _
                               Integer, Shift As Integer, x _
                               As Single, y As Single)
  DragFlag = True
End Sub

Private Sub pic_resizeimport_MouseUp(Index As Integer, Button As _
                             Integer, Shift As Integer, x _
                             As Single, y As Single)
   DragFlag = False
   'pic_import halb transparent machen
   Draw_AlphaBlend pic_target, pic_alphablend, pic_import
End Sub

Private Sub pic_resizeimport_MouseMove(Index As Integer, Button As _
                               Integer, Shift As Integer, x _
                               As Single, y As Single)
                               
  Dim xP&, yP&, x1&, x2&, y1&, y2&, WPM As WINDOWPLACEMENT
  Dim TPX%, TPY&, XNoSize As Boolean, YNoSize As Boolean
  Static MemX1&, MemX2&, MemY1&, Memy2&
  Static Doing As Boolean
    
    If Doing Then Exit Sub
    Doing = True
     
     DoEvents
    If DragFlag Then
      TPX = 1 'Screen.TwipsPerPixelX
      TPY = 1 'Screen.TwipsPerPixelY
      xP = x / TPX
      yP = y / TPY
      
      WPM.Length = Len(WPM)
      Call GetWindowPlacement(MCtrl.hwnd, WPM)
      x1 = WPM.rcNormalPosition.Left
      x2 = WPM.rcNormalPosition.Right
      y1 = WPM.rcNormalPosition.Top
      y2 = WPM.rcNormalPosition.Bottom
    
      With pic_resizeimport(Index)
            
        If Index = 0 Or Index = 1 Or Index = 2 Then
          If x1 + xP > x2 - Wmin Then
            XNoSize = True
            x1 = x2 - Wmin
          ElseIf x2 - (x1 + xP) > Wmax Then
            XNoSize = True
            x1 = x2 - Wmax
          Else
            x1 = x1 + xP
          End If
          If x1 <= LimX1 Then
            XNoSize = True
            x1 = LimX1
          End If
        End If
        DoEvents
        If Index = 4 Or Index = 5 Or Index = 3 Then
          If x2 + xP < x1 + Wmin Then
            XNoSize = True
            x2 = x1 + Wmin
          ElseIf x2 + xP - x1 > Wmax Then
            XNoSize = True
            x2 = x1 + Wmax
          Else
            x2 = x2 + xP
          End If
          
          If x2 > LimX2 Then
            XNoSize = True
            x2 = LimX2
          End If
        End If
        DoEvents
        If Index = 0 Or Index = 6 Or Index = 3 Then
          If y1 + yP > y2 - Hmin Then
            YNoSize = True
            y1 = y2 - Hmin
          ElseIf y2 - (y1 + yP) > Hmax Then
            YNoSize = True
            y1 = y2 - Hmax
          Else
            y1 = y1 + yP
          End If
          
          If y1 <= LimY1 Then
            YNoSize = True
            y1 = LimY1
          End If
        End If
         
        If Index = 7 Or Index = 2 Or Index = 5 Then
          If y2 + yP < y1 + Hmin Then
            YNoSize = True
            y2 = y1 + Hmin
          ElseIf y2 + yP - y1 > Hmax Then
            YNoSize = True
            y2 = y1 + Hmax
          Else
            y2 = y2 + yP
          End If
          
          If y2 > LimY2 Then
            YNoSize = True
            y2 = LimY2
          End If
        End If
    DoEvents
      Select Case Index
        Case 0, 2, 3, 5: y = y + .Top
                         x = x + .Left
                        
        Case 1, 4:       x = x + .Left
                         y = .Top
               
        Case 6, 7:       y = y + .Top
                         x = .Left
      End Select
      
      If Not YNoSize Then .Top = y
      If Not XNoSize Then .Left = x
    End With

    If MemX1 <> x1 Or MemX2 <> x2 Or MemY1 <> y1 _
                   Or Memy2 <> y2 Then
      If TypeOf MCtrl Is ListBox Or TypeOf MCtrl Is ComboBox Then
        '...
      Else
        WPM.rcNormalPosition.Left = x1
        WPM.rcNormalPosition.Top = y1
        WPM.rcNormalPosition.Right = x2
        WPM.rcNormalPosition.Bottom = y2
        Call SetWindowPlacement(MCtrl.hwnd, WPM)
        Call DrawPics
      End If
    End If
    
    MemX1 = x1
    MemX2 = x2
    MemY1 = y1
    Memy2 = y2
    'Flackert, nur bei Bedarf wieder einkommentieren !
    '''pic_import halb transparent machen
    ''Draw_AlphaBlend pic_target, pic_alphablend, pic_import
  End If
  
  Doing = False
End Sub

Private Sub pic_import_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf pic_import                             |'
'--------------------------------------------------------------------'
   'Wenn der Fertigstellen-Button unsichtbar ist, Prozedur beenden
   If Not cmd_finishimport.Visible Then Exit Sub
   Set MCtrl = pic_import
   pic_import.MousePointer = vbSizeAll
   Call DrawPics
   Fetched = True
   'pic_import halb transparent machen
   Draw_AlphaBlend pic_target, pic_alphablend, pic_import
End Sub

Private Sub pic_import_Resize()
'--------------------------------------------------------------------'
'| Prozedur beim Verändern der Größe der Picturebox                 |'
'--------------------------------------------------------------------'
   'Bild neu anpassen
   Draw_Zoom pic_bild, pic_import
   'Größe der pic_alphablend Picbox anpassen
   pic_alphablend.Width = pic_import.Width
   pic_alphablend.Height = pic_import.Height
   'Bild von pic_bild in pic_alphablend zoomen
   Draw_Zoom pic_bild, pic_alphablend
End Sub

Private Sub cmd_finishimport_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf cmd_finishimport                       |'
'--------------------------------------------------------------------'
   'Verhältnis zwischen Source und Target ausrechnen
   ZoomX = pic_source.Width / pic_target.Width
   ZoomY = pic_source.Height / pic_target.Height

   'Bei Fehler weitermachen, da AutoRedraw-Bild (manchmal !!) nicht zur Verfügung steht..
   On Error Resume Next
   'Bild in große PicBox zoomen
   StretchBlt pic_bildgroß.hdc, (pic_import.Left - pic_target.Left) * ZoomX, (pic_import.Top - pic_target.Top) * ZoomY, pic_import.Width * ZoomX, pic_import.Height * ZoomY, pic_bild.hdc, 0, 0, pic_bild.Width, pic_bild.Height, SRCCOPY
   'Fehlerbehandlung wieder ausschalten
   On Error GoTo 0
   
   Draw_Refresh pic_bildgroß
   'Button zum Fertigstellen unsichtbar machen
   cmd_finishimport.Visible = False
   'Button zum Abbrechen unsichtbar machen
   cmd_cancelimport.Visible = False
   'Mauszeiger von pic_import zurücksetzen
   pic_import.MousePointer = 0
   'Schleife durch alle blaue Boxen
   For i = 0 To pic_resizeimport.UBound
      'Picbox unsichtbar machen
      pic_resizeimport(i).Visible = False
   Next i

   'Statusbar-Anzeige ändern
   stb_statusbar.Panels(2).Text = FileSystem_Path2File(BildPfad) & " wird importiert ..."

   'Bild ins Modell füllen
   Draw_ImportPicture pic_alphablend, pic_source, pic_import, pic_target, prg_fortschritt

   'Informationen in Statusbar anzeigen
   NightGraphix_ShowPanelProperties
   
   'PicBox unsichtbar machen
   pic_import.Visible = False
End Sub

Private Sub cmd_cancelimport_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf cmd_cancelimport                       |'
'--------------------------------------------------------------------'
   'Button zum Fertigstellen unsichtbar machen
   cmd_finishimport.Visible = False
   'Button zum Abbrechen unsichtbar machen
   cmd_cancelimport.Visible = False
   'PicBox unsichtbar machen
   pic_import.Visible = False
   'Schleife durch alle blaue Boxen
   For i = 0 To pic_resizeimport.UBound
      'Picbox unsichtbar machen
      pic_resizeimport(i).Visible = False
   Next i
End Sub

Private Sub DrawPics()
  Dim TPX%, TPY&, x%
     
     TPX = 1 'Screen.TwipsPerPixelX
     TPY = 1 'Screen.TwipsPerPixelY
     
     With MCtrl
       pic_resizeimport(0).Left = .Left - Sp * TPX - pic_resizeimport(0).Width
       pic_resizeimport(1).Left = .Left - Sp * TPX - pic_resizeimport(1).Width
       pic_resizeimport(2).Left = .Left - Sp * TPX - pic_resizeimport(2).Width
       pic_resizeimport(0).Top = .Top - Sp * TPY - pic_resizeimport(0).Height
       pic_resizeimport(1).Top = .Top + (.Height - pic_resizeimport(1).Height) / 2
       pic_resizeimport(2).Top = .Top + .Height + Sp * TPY
       
       pic_resizeimport(3).Left = .Left + .Width + Sp * TPX
       pic_resizeimport(4).Left = .Left + .Width + Sp * TPX
       pic_resizeimport(5).Left = .Left + .Width + Sp * TPX
       pic_resizeimport(3).Top = .Top - Sp * TPY - pic_resizeimport(0).Height
       pic_resizeimport(4).Top = .Top + (.Height - pic_resizeimport(1).Height) / 2
       pic_resizeimport(5).Top = .Top + .Height + Sp * TPY
       
       pic_resizeimport(6).Left = .Left + (.Width - pic_resizeimport(6).Width) / 2
       pic_resizeimport(7).Left = .Left + (.Width - pic_resizeimport(6).Width) / 2
       
       pic_resizeimport(6).Top = .Top - Sp * TPY - pic_resizeimport(0).Height
       pic_resizeimport(7).Top = .Top + .Height + Sp * TPY
     End With
     DoEvents
    If Not pic_resizeimport(0).Visible Then
      For x = 0 To 7
        pic_resizeimport(x).Visible = True
      Next x
    End If
End Sub

Public Sub Draw_OffsetCursor(ByRef tmpRotateDegree As Double)
'--------------------------------------------------------------------'
'| Prozedur zum Zeichnen des Offset-Pfeils                          |'
'--------------------------------------------------------------------'
   'Bitmap aus pic_sourceBitmap gedreht in pic_rotateBitmap anzeigen
   Draw_RotatePicture pic_arrowbitmap, pic_rotatearrowbitmap, tmpRotateDegree
   'Bitmap aus pic_sourceBitmap gedreht in pic_rotateBitmap anzeigen
   Draw_RotatePicture pic_arrowmask, pic_rotatearrowmask, tmpRotateDegree
   
   'Picturebox leeren
   pic_arrowshow.Cls
   'Rotierte Maske auf pic_show blitten
   Call BitBlt(pic_arrowshow.hdc, 0, 0, pic_rotatearrowmask.ScaleWidth, pic_rotatearrowmask.ScaleHeight, pic_rotatearrowmask.hdc, 0, 0, BIT_AND)
   'Rotiertes Bitmap auf pic_show blitten
   Call BitBlt(pic_arrowshow.hdc, 0, 0, pic_rotatearrowbitmap.ScaleWidth, pic_rotatearrowbitmap.ScaleHeight, pic_rotatearrowbitmap.hdc, 0, 0, BIT_Invert)
   
   'Picturebox neu zeichnen
   pic_arrowshow.Refresh
End Sub

Public Sub Draw_ZoomModell(Index As Integer)
'--------------------------------------------------------------------'
'| Prozedur zum Zoomen des Modells                                  |'
'--------------------------------------------------------------------'
   
   'Gedrückten Button erkennen
   If Index = 0 Then
      'Zoomen
      Draw_ZoomUp pic_target
      
      'Bild übertragen
      Draw_Zoom pic_source, pic_target
      
      'Scrollleisten in die Mitte scrollen
      hsc_scrollen.Value = hsc_scrollen.Max / 2
      vsc_scrollen.Value = vsc_scrollen.Max / 2
      
      'Scrollleisten sichtbar machen
      hsc_scrollen.Visible = True
      vsc_scrollen.Visible = True
      
      'Buttons sichtbar/unsichtbar machen
      tlb_toolbar.Buttons(9).Enabled = False
      tlb_toolbar.Buttons(10).Enabled = True
   ElseIf Index = 1 Then
      'Zoomen
      Draw_ZoomDown pic_target
      
      'Bild übertragen
      Draw_Zoom pic_source, pic_target
      
      'Scrollleisten auf 0 stellen
      hsc_scrollen.Value = 0
      vsc_scrollen.Value = 0
      
      'Scrollleisten unsichtbar machen
      hsc_scrollen.Visible = False
      vsc_scrollen.Visible = False
      
      'Buttons sichtbar/unsichtbar machen
      tlb_toolbar.Buttons(9).Enabled = True
      tlb_toolbar.Buttons(10).Enabled = False
   End If
   
   'Max-Wert der Scrollleisten einstellen
   frm_nightgraphix.vsc_scrollen.Max = frm_nightgraphix.pic_target.Height - frm_nightgraphix.pic_rahmen.Height
   'Max-Wert der Scrollleisten einstellen
   frm_nightgraphix.hsc_scrollen.Max = frm_nightgraphix.pic_target.Width - frm_nightgraphix.pic_rahmen.Width
End Sub

'Public Sub NightGraphix_ShowPanelProperties()
''--------------------------------------------------------------------'
''| Prozedur zum Anzeigen der Eigenschaften in der Panelbar          |'
''--------------------------------------------------------------------'
'   'Eigenschaften im 2. Panel anzeigen
'   stb_statusbar.Panels(2).Text = "Offset: " & Format(CStr(Offset), "000") & "   Aktive LEDs: " & Format(CStr(Maths_CountActiveLEDs), "00000") & "   Strom: " & Format(CStr(Maths_CountActiveLEDs * CDbl(20) \ Spalten), "000,000") & " Ah"
'End Sub

Public Sub NightGraphix_ShowPanelProperties()
'--------------------------------------------------------------------'
'| Prozedur zum Anzeigen der Eigenschaften in der Panelbar          |'
'--------------------------------------------------------------------'
'Variablen deklarieren und vorbelegen
Dim LED_Strom, Laufzeit, LED_Anzahl As Long
   LED_Anzahl = Format(CStr(Maths_CountActiveLEDs), "00000")
   'wenn die Anzahl der eingeschalteten LED´s 0 ist Sub verlassen (wegen Div / 0)
   If LED_Anzahl = 0 Then
      'Laufzeit auf größer 90min setzten
      stb_statusbar.Panels(2).Text = "(Offset: " & Format(CStr(Offset), "000") & ")  (Laufzeit: >90min)" & "  (" & Format(CStr(Maths_CountActiveLEDs)) & " LED´s)" & "  (" & Format(CStr(Maths_CountActiveLEDs * CDbl(20) \ Spalten)) & " mAh)"
      Exit Sub
   End If
   
   'Berechnung der Laufzeit
   LED_Strom = LED_Anzahl * 20 / Spalten
   Laufzeit = (NGLiPoCapacity / LED_Strom) * 60
   'Ergebniss aufrunden
   Laufzeit = Abs(Int(-Laufzeit))
   
   'wenn die Laufzeit >90min ist keine tatsächliche Laufzeit anzeigen (da nicht relevant)
   If Laufzeit > 90 Then
      'Eigenschaften im 2. Panel anzeigen
      stb_statusbar.Panels(2).Text = "(Offset: " & Format(CStr(Offset), "000") & ") (Laufzeit: >90min)" & " (" & Format(CStr(Maths_CountActiveLEDs)) & " LEDs)" & "  (" & Format(CStr(Maths_CountActiveLEDs * CDbl(20) \ Spalten)) & " mAh)"
   'Wenn die Laufzeit <90 Minuten ist, dann tatsächliche Laufzeit anzeigen
   Else
      'Eigenschaften im 2. Panel anzeigen
      stb_statusbar.Panels(2).Text = "(Offset: " & Format(CStr(Offset), "000") & ") (Laufzeit: ca." & Laufzeit & "min)" & " (" & Format(CStr(Maths_CountActiveLEDs)) & " LEDs)" & "  (" & Format(CStr(Maths_CountActiveLEDs * CDbl(20) \ Spalten)) & " mAh)"
   End If
   
   '*original* stb_statusbar.Panels(2).Text = "Offset: " & Format(CStr(Offset), "000") & "   Aktive LEDs: " & Format(CStr(Maths_CountActiveLEDs), "00000") & "   Strom: " & Format(CStr(Maths_CountActiveLEDs * CDbl(20) \ Spalten), "000,000") & " Ah"
End Sub

Public Sub NightGraphix_EnableToolButtons()
'--------------------------------------------------------------------'
'| Prozedur zum Aktivieren der Buttons auf dem Hauptfenster         |'
'--------------------------------------------------------------------'
   'Option-Buttons am linken Rand aktivieren
   'Schleife durch alle Optionbuttons
   For i = 0 To cmd_tool.UBound
      'Button aktivieren
      cmd_tool(i).Enabled = True
   Next i
   
   'Toolbar-Buttons in der oberen Toolbar aktivieren
   For i = 1 To tlb_toolbar.Buttons.Count
      'Button aktivieren
      tlb_toolbar.Buttons(i).Enabled = True
   Next i
   
   'Demo starten auf Demo beenden umschalten
   tlb_toolbar.Buttons(17).Visible = False
   tlb_toolbar.Buttons(18).Visible = True
   
   'Zoom - deaktivieren
   tlb_toolbar.Buttons(10).Enabled = False
   
   'Wenn der Demo-Modus gestartet ist..
   If NGDemoModus = True Then
      'Trennen deaktivieren
      tlb_toolbar.Buttons(16).Enabled = False
      'EEPROM lesen und schreiben deaktivieren
      tlb_toolbar.Buttons(12).Enabled = False
      tlb_toolbar.Buttons(13).Enabled = False
      'Hardware-Optionen deaktivieren
      tlb_toolbar.Buttons(21).Enabled = False
   End If
End Sub

Public Sub NightGraphix_DisableToolButtons()
'--------------------------------------------------------------------'
'| Prozedur zum Deaktivieren der Buttons auf dem Hauptfenster       |'
'--------------------------------------------------------------------'
   'Pixel-Tool anschalten
   cmd_tool(0).Value = True
   'Option-Buttons am linken Rand deaktivieren
   'Schleife durch alle Optionbuttons
   For i = 0 To cmd_tool.UBound
      'Button deaktivieren
      cmd_tool(i).Enabled = False
   Next i
   'Font-Auswahl deaktivieren
   cmd_font.Enabled = False
   
   'Toolbar-Buttons in der oberen Toolbar deaktivieren
   For i = 1 To tlb_toolbar.Buttons.Count
      'Button deaktivieren
      tlb_toolbar.Buttons(i).Enabled = False
   Next i
   
   'Trennen-Button unsichtbar machen, Verbinden sichtbar machen
   tlb_toolbar.Buttons(16).Visible = False
   tlb_toolbar.Buttons(15).Visible = True

   'Öffnen-Button aktivieren
   tlb_toolbar.Buttons(2).Enabled = True
   'Verbinden aktivieren
   tlb_toolbar.Buttons(15).Enabled = True
   'Demo starten und Demo beenden aktivieren
   tlb_toolbar.Buttons(17).Enabled = True
   tlb_toolbar.Buttons(18).Enabled = True
   'Demo beenden auf Demo starten umschalten
   tlb_toolbar.Buttons(17).Visible = True
   tlb_toolbar.Buttons(18).Visible = False
   'Software-Optionen aktivieren
   tlb_toolbar.Buttons(20).Enabled = True
   'Fragezeichen aktivieren
   tlb_toolbar.Buttons(22).Enabled = True
      
   'Beenden aktivieren
   tlb_toolbar.Buttons(24).Enabled = True
End Sub

Public Sub NightGraphix_ReInitialize()
'--------------------------------------------------------------------'
'| Prozedur zum Wiederherstellen der grauen Fläche                  |'
'--------------------------------------------------------------------'
   'Es wird mit Demo gearbeitet
   NGDemoModus = False
   'Programm-Einstellungen aus der Registry laden
   FileSystem_GetSettings
   'LED-Anzahl setzen
   Leds = (NGDemoSize + 3) * 8
   'Spaltenanzahl setzen
   Spalten = 512
   'RGB-Version
   RGBVersion = False
   'Hub setzen
   Hub = (NGRotorSize - 2 * Leds * 10) / 10
   'Variablen initialisieren
   Init_Variables pic_source
   'Namen der Form initialisieren
   Me.Caption = "NightGraphix V1.0"
   'Font-PictureBox initialisieren
   Init_FontPictureBox pic_letter
   'DateiHeader initialisieren
   Init_FileHeader
   'Arrays initialisieren
   Init_Arrays
   
   pic_source.Cls
   pic_arrowshow.Cls
   'Hintergrund grau füllen
   Draw_Background pic_source

   'pic_source refreshen
   Draw_Refresh pic_source

   'Bild von pic_source nach pic_target zoomen
   Draw_Zoom pic_source, pic_target

   'PictureBox refreshen
   Draw_Refresh pic_target

   If CInt(ZoomX) = 1 Then
      'Bild klein zoomen
      Draw_ZoomModell (1)
   End If

   'Buttons deaktivieren
   NightGraphix_DisableToolButtons
End Sub

Public Sub NightGraphix_StartDemo()
'--------------------------------------------------------------------'
'| Prozedur zum Starten der Demo                                    |'
'--------------------------------------------------------------------'
   'Es wird mit Demo gearbeitet
   NGDemoModus = True
   'Programm-Einstellungen aus der Registry laden
   FileSystem_GetSettings
   'Hub setzen
   Hub = (NGRotorSizeArray((Leds - 16) / 8) - 2 * Leds * 10 - 2 * 35) / 20
   'Variablen initialisieren
   Init_Variables pic_source
   'Namen der Form initialisieren
   Init_Form Me
   'Font-PictureBox initialisieren
   Init_FontPictureBox pic_letter
   'DateiHeader initialisieren
   Init_FileHeader
   'Arrays initialisieren
   Init_Arrays
   'Zeichenfläche initialisieren
   Draw_Constructor pic_source, pic_target, pic_arrowshow
   'Buttons aktivieren
   NightGraphix_EnableToolButtons
End Sub

