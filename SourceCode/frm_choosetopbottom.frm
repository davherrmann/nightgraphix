VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_choosetopbottom 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "NightGraphiX V1.0"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   Icon            =   "frm_choosetopbottom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3975
   StartUpPosition =   2  'Bildschirmmitte
   Begin ComctlLib.ProgressBar prg_xpstyle 
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1080
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
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.OptionButton opt_topbottom 
      Caption         =   "Unterseite"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton opt_topbottom 
      Caption         =   "Oberseite"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Label lbl_beschriftung 
      Caption         =   "Auf welcher Seite befindet sich die angeschlossene Hardware ?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frm_choosetopbottom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variablen müssen deklariert werden..
Option Explicit

Private Sub cmd_ok_Click()
'--------------------------------------------------------------------'
'| Prozedur beim Klicken auf den OK-Button                          |'
'--------------------------------------------------------------------'
   'Seite der angeschlossenen Hardware in Variable speichern
   NGTopSideMC = opt_topbottom(0).Value
   'Hardwaredaten an den MC senden
   Communication_SendHardwareData frm_nightgraphix.msc_seriell
   'Fenster verstecken
   Me.Hide
   'Fenster schließen
   Unload Me
End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------'
'| Prozedur beim Laden der Form                                     |'
'--------------------------------------------------------------------'
   'Icon der Form initialisieren
   Icons_Init Me
End Sub
