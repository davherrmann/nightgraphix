Attribute VB_Name = "mdl_public"
'_________________________________________________________________________________'
'|                               MODUL mdl_draw                                  |'
'| Dieses Modul beinhaltet öffentliche Variablen und Funktionen                  |'
'|                                                                               |'
'---------------------------------------------------------------------------------'

Option Explicit

'VARIABLEN..
'Ist Maustaste gedrückt ?
Public MouseButton As Integer
'Ist Text- oder Graphikversion gewählt ?
Public TextVersion As Boolean

'ENUMERATIONEN..
'Höhe der ToolButtons
Public Enum eToolTop
   'Höhe des Stiftes
   PencilTop = 56
   'Höhe des Pinsels
   BrushTop = 96
   'Höhe des Füll-Tools
   FillTop = 136
   'Höhe der Spraydose
   SprayTop = 176
   'Höhe des Text-Tools
   TextTop = 216
   'Höhe des Radierers
   EraserTop = 256
End Enum

'Referenz auf das Enum-Objekt
Public ToolTop As eToolTop
