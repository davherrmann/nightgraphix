Attribute VB_Name = "mdl_public"
'_________________________________________________________________________________'
'|                               MODUL mdl_draw                                  |'
'| Dieses Modul beinhaltet �ffentliche Variablen und Funktionen                  |'
'|                                                                               |'
'---------------------------------------------------------------------------------'

Option Explicit

'VARIABLEN..
'Ist Maustaste gedr�ckt ?
Public MouseButton As Integer
'Ist Text- oder Graphikversion gew�hlt ?
Public TextVersion As Boolean

'ENUMERATIONEN..
'H�he der ToolButtons
Public Enum eToolTop
   'H�he des Stiftes
   PencilTop = 56
   'H�he des Pinsels
   BrushTop = 96
   'H�he des F�ll-Tools
   FillTop = 136
   'H�he der Spraydose
   SprayTop = 176
   'H�he des Text-Tools
   TextTop = 216
   'H�he des Radierers
   EraserTop = 256
End Enum

'Referenz auf das Enum-Objekt
Public ToolTop As eToolTop
