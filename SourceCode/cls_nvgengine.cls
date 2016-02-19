VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_nvgengine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Referenz auf die Buffer-Klasse
Private NVB As cls_nvbuffer
'Referenz auf das PicSource-Steuerelement
Private PicSource As PictureBox
'Referenz auf das PicTarget-Steuerelement
Private PicTarget As PictureBox
'Werden die gebufferten Tasks gezeichnet ?
Public DrawingBufferedTasks As Boolean
'Task, der gezeichnet werden soll
Private Task As cls_nvtask

Public Sub Init(ByRef tmpPicSource As PictureBox, ByRef tmpPicTarget As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren der Graphik-Engine                   |'
'--------------------------------------------------------------------'
   'Referenz auf die Buffer-Klasse setzen
   Set NVB = New cls_nvbuffer
   'Referenz auf die Task-Klasse setzen
   Set Task = New cls_nvtask
   'Referenz auf PicSource-Steuerelement erstellen
   Set PicSource = tmpPicSource
   'Referenz auf PicTarget-Steuerelement erstellen
   Set PicTarget = tmpPicTarget
End Sub

Private Sub DrawBufferedTasks()
'--------------------------------------------------------------------'
'| Prozedur zum Zeichnen der gebufferten Tasks                      |'
'--------------------------------------------------------------------'
   'Die gebufferten Tasks werden gezeichnet
   DrawingBufferedTasks = True
   'Endlosschleife, bis alle gebufferten Tasks gezeichnet sind
   While (NVB.NumberOfTasks > 0)
      'Anderen Events auch Zeit lassen..
      DoEvents
      'Task aus FIFO-Buffer holen
      Set Task = NVB.GetTask
      'Typ des Tasks auswerten
      Select Case Task.Typ
         'Wenn ein Punkt gezeichnet werden soll..
         Case [Draw Point]
            'Argumente an Draw_Click übergeben
            Draw_Click PicSource, PicTarget, Task.X1, Task.Y1
         'Wenn eine Linie gezeichnet werden soll..
         Case [Draw Line]
            'Argumetne an Draw_BresenhamLine übergeben
            Draw_BresenhamLine Task.X1, Task.Y1, Task.X2, Task.Y2, PicTarget, PicSource
         'Wenn ein Kreis gezeichnet werden soll..
         Case [Draw Circle]
            'Argumente an Draw_Circle übergeben
            Draw_Circle PicSource, PicTarget, Task.Y1, Task.Color
         'Wenn ein Buchstabe gezeichnet werden soll..
         Case [Draw Letter]
            'Argumente an Text_Write übergeben
            Text_Write PicSource, PicTarget, frm_nightvision.prg_fortschritt, Task.Letter, Task.X1, Task.Y1, Task.Color
      End Select
      frm_nightvision.Label1.Caption = CStr(NVB.NumberOfTasks)
   Wend
   'Die gebufferten Tasks werden nicht mehr gezeichnet
   DrawingBufferedTasks = False
End Sub

Public Sub DrawPoint(ByRef X As Double, ByRef Y As Double)
'--------------------------------------------------------------------'
'| Prozedur Zeichnen eines Punktes                                  |'
'--------------------------------------------------------------------'
   'Task-Eigenschaften ändern
   'Typ des Tasks setzen
   Task.Typ = [Draw Point]
   'X-Koordinaten setzen
   Task.X1 = X
   'Y-Koordinaten setzen
   Task.Y1 = Y
   'Farbe setzen
   Task.Color = LEDColor
   
   'Neuen Task im Buffer anmelden
   NVB.AddTask Task
   If NVB.NumberOfTasks > 10 Then NVB.DeleteFirstTask
   frm_nightvision.Label1.Caption = CStr(NVB.NumberOfTasks)

   'Wenn Tasks im Buffer noch nicht gezeichnet werden..
   If Not DrawingBufferedTasks Then
      'Tasks im Buffer verarbeiten
      DrawBufferedTasks
   End If
End Sub



