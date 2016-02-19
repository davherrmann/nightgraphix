VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_splashscreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'Kein
   Caption         =   "Nightgraphix V1.0 - Hardware suchen"
   ClientHeight    =   1770
   ClientLeft      =   4500
   ClientTop       =   3345
   ClientWidth     =   1770
   ForeColor       =   &H00000000&
   Icon            =   "frm_startoptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.ProgressBar prg_xpstyle 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmr_comports 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   600
      Top             =   4320
   End
   Begin VB.Timer tmr_rotate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   4320
   End
   Begin VB.PictureBox pic_background 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   1500
      Left            =   3120
      Picture         =   "frm_startoptions.frx":1582
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1500
   End
   Begin VB.PictureBox pic_rotateMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   720
      Left            =   2400
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3000
      Width           =   660
   End
   Begin VB.PictureBox pic_rotateBitmap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      Height          =   720
      Left            =   1680
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3000
      Width           =   660
   End
   Begin VB.PictureBox pic_sourceMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   600
      Picture         =   "frm_startoptions.frx":8AF4
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3000
      Width           =   420
   End
   Begin VB.PictureBox pic_sourceBitmap 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   480
      Left            =   120
      Picture         =   "frm_startoptions.frx":95B6
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3000
      Width           =   420
   End
   Begin MSCommLib.MSComm msc_comport 
      Left            =   120
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   8
   End
   Begin VB.PictureBox pic_show 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      Height          =   1455
      Left            =   120
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1455
      Begin VB.Label lbl_beschriftung 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "Hardware wird gesucht.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Shape shp_frame 
      BorderColor     =   &H80000002&
      Height          =   1770
      Left            =   0
      Top             =   0
      Width           =   1770
   End
End
Attribute VB_Name = "frm_splashscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Winkel der Rotation
Private RotateDegree As Long
'Variable für Zähler, Schleifen etc.
Private i As Integer

Private Sub Form_Load()
'--------------------------------------------------------------------'
'| Prozedur beim Laden der Form                                     |'
'--------------------------------------------------------------------'
   'Icon der Form initialisieren
   Icons_Init Me
   'XP-Style-Progressbar nach vorne bringen
   prg_xpstyle.ZOrder 0
   'XP-Style initialisieren
   Init_XPStyle
   'Rotations Timer setzen
   tmr_rotate.Enabled = True
   'Form Caption ändern
   lbl_beschriftung.Caption = LngSpecStrings.LookingForHardware
End Sub

Private Sub tmr_rotate_Timer()
'--------------------------------------------------------------------'
'| Prozedur zum Drehen der Pfeile                                   |'
'--------------------------------------------------------------------'
   'Winkelgröße erhöhen
   RotateDegree = (RotateDegree - 3) Mod 360
   'Bitmap aus pic_sourceBitmap gedreht in pic_rotateBitmap anzeigen
   Draw_RotatePicture pic_sourceBitmap, pic_rotateBitmap, RotateDegree
   'Bitmap aus pic_sourceBitmap gedreht in pic_rotateBitmap anzeigen
   Draw_RotatePicture pic_sourceMask, pic_rotateMask, RotateDegree
   
   'Hintergrundbild auf pic_show blitten
   Call BitBlt(pic_show.hdc, 0, 0, pic_show.ScaleWidth, pic_show.ScaleHeight, pic_background.hdc, 0, 0, BIT_COPY)
   'Rotierte Maske auf pic_show blitten
   Call BitBlt(pic_show.hdc, 27, 10, pic_rotateBitmap.ScaleWidth, pic_rotateBitmap.ScaleHeight, pic_rotateMask.hdc, 0, 0, BIT_AND)
   'Rotiertes Bitmap auf pic_show blitten
   Call BitBlt(pic_show.hdc, 27, 10, pic_rotateBitmap.ScaleWidth, pic_rotateBitmap.ScaleHeight, pic_rotateBitmap.hdc, 0, 0, BIT_Invert)

   'pic_show refreshen
   pic_show.Refresh
End Sub
