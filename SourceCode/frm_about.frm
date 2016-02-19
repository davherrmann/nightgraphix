VERSION 5.00
Begin VB.Form frm_about 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Info über NighGraphiX V1.0"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   Icon            =   "frm_about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   315
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   1  'Fenstermitte
   Begin VB.PictureBox pic_about 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   4710
      Left            =   0
      Picture         =   "frm_about.frx":6852
      ScaleHeight     =   314
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   0
      Width           =   6000
      Begin VB.PictureBox pic_copyright 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   2640
         Picture         =   "frm_about.frx":62874
         ScaleHeight     =   207
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   215
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lbl_link 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.NightGraphiX.de"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF6644&
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lbl_version 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fest Einfach
         Caption         =   "Version 1.0.134"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF7700&
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   3840
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variablen müssen deklariert werden
Option Explicit

Private Sub Form_Load()
'--------------------------------------------------------------------'
'| Prozedur beim Laden der Form: Initialisierungen                  |'
'--------------------------------------------------------------------'
   'Icon der Form initialisieren
   Icons_Init Me
   'Version im Label anzeigen
   lbl_version.Caption = "Version " & CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
End Sub
