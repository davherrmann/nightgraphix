VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_choosefileversion 
   BackColor       =   &H00EFEFEF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "NightGraphiX V1.0 - Neue Datei"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3735
   Icon            =   "frm_choosefileversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.ProgressBar prg_xpstyle 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.OptionButton opt_version 
      BackColor       =   &H00EFEFEF&
      Caption         =   "NG24 - 24 LEDs"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.OptionButton opt_version 
      BackColor       =   &H00EFEFEF&
      Caption         =   "NG32 - 32 LEDs"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton opt_version 
      BackColor       =   &H00EFEFEF&
      Caption         =   "NG40 - 40 LEDs"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.OptionButton opt_version 
      BackColor       =   &H00EFEFEF&
      Caption         =   "NG48 - 48 LEDs"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton opt_version 
      BackColor       =   &H00EFEFEF&
      Caption         =   "NG56 - 56 LEDs"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.OptionButton opt_version 
      BackColor       =   &H00EFEFEF&
      Caption         =   "NG64 - 64 LEDs"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lbl_beschriftung 
      BackColor       =   &H00EFEFEF&
      Caption         =   "Bitte wählen Sie die Version, für die Sie die Datei nutzen wollen:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frm_choosefileversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Alle Variablen müssen deklariert werden
Option Explicit

'Variable für Schleifen
Private i As Integer

Private Sub cmd_ok_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf den OK-Button                          |'
'--------------------------------------------------------------------'
   'Schleife durch alle Optionbuttons
   For i = 0 To opt_version.UBound
      'Wenn der Optionbutton ausgwählt ist..
      If opt_version(i).Value Then
         'Version für eine neue Datei in Variable speichern
         NGNewFileVersion = i
         'Fenster schließen
         Unload Me
         'Prozedur beenden
         Exit Sub
      End If
   Next i
End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------'
'| Prozedur beim Laden der Form                                     |'
'--------------------------------------------------------------------'
   'Icon der Form initialisieren
   Icons_Init Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'--------------------------------------------------------------------'
'| Prozedur beim Schließen der Form                                 |'
'--------------------------------------------------------------------'
   'Wenn die Form über den Schließen-Button geschlossen werden soll..
   If UnloadMode = 0 Then
      'Beenden abbrechen
      Cancel = 1
   End If
End Sub
