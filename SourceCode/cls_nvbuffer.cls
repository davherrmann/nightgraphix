VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_nvbuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'_________________________________________________________________________________'
'|                               KLASSE cls_nvbuffer                             |'
'| Diese Klasse erzeugt einen Buffer nach dem FIFO-Prinzip                       |'
'|                                                                               |'
'---------------------------------------------------------------------------------'

'Variablen müssen deklariert werden..
Option Explicit

'VARIABLEN..
'Wieviele Tasks gibt es ?
Private mNumberOfTasks As Long
'Array zum Speichern der Tasks
Private Tasks As Collection
'Variablen für Schleifen (1- oder 2-dimensional)
Private X As Integer
Private Y As Integer

Private Sub Class_Initialize()
'--------------------------------------------------------------------'
'| Prozedur beim Initialisieren der Klasse                          |'
'--------------------------------------------------------------------'
   'Referenz auf Task-Collection setzen
   Set Tasks = New Collection
End Sub

Public Sub AddTask(ByRef Task As cls_nvtask)
'--------------------------------------------------------------------'
'| Prozedur zum Hinzufügen einer neuen Aufgabe                      |'
'--------------------------------------------------------------------'
   'Anzahl der Tasks um eins erhöhen
   mNumberOfTasks = mNumberOfTasks + 1
   'Task zur Task-Collection hinzufügen
   Tasks.Add Task
End Sub

Public Function GetTask() As cls_nvtask
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben einer Aufgabe                           |'
'--------------------------------------------------------------------'
   'Wenn es keinen Task mehr gibt..
   If Tasks.Count < 1 Then
      '"" (ungültig) zurückgeben
      GetTask = ""
   'Wenn es noch Tasks gibt..
   Else
      'Ersten Task zurückgeben
      Set GetTask = Tasks(1)
      'Ersten Task löschen
      DeleteFirstTask
   End If
End Function

Public Function DeleteFirstTask()
'--------------------------------------------------------------------'
'| Prozedur zum Löschen der ersten Aufgabe                          |'
'--------------------------------------------------------------------'
   'Ersten Task löschen
   Tasks.Remove 1
   'Anzahl der Tasks um eins erniedrigen
   mNumberOfTasks = mNumberOfTasks - 1
End Function

Public Function ClearTasks()
'--------------------------------------------------------------------'
'| Prozedur zum Löschen aller Tasks                                 |'
'--------------------------------------------------------------------'
   'Schleife durch alle Tasks der Collection
   For X = 1 To Tasks.Count
      'Task löschen (immer 0, da der Index automatisch erniedrigt wird)
      Tasks.Remove 0
   Next Task
End Function

Public Property Get NumberOfTasks() As Variant
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben der Anzahl an Tasks                     |'
'--------------------------------------------------------------------'
   'Task-Anzahl zurückgeben
   NumberOfTasks = mNumberOfTasks
End Property
