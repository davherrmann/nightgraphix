Attribute VB_Name = "mdl_public"
'_________________________________________________________________________________'
'|                               MODUL mdl_draw                                  |'
'| Dieses Modul beinhaltet öffentliche Variablen und Funktionen                  |'
'|                                                                               |'
'---------------------------------------------------------------------------------'
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Ist Maustaste gedrückt ?
Public MouseButton As Integer

'Handle des gesuchten Fensters
Private RetVal As Long
Private TaskHwnd As Long
Private TaskResult As Long
Private TaskTitle As String

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

'Position des Mauscursors
Public CurPos As POINTAPI
'Eckpunkte, in denen gescrollt werden darf
Public P1 As POINTAPI, P2 As POINTAPI

'APIs..
'Api zum Ermitteln der Mausposition
Public Declare Function GetCursorPos Lib "user32" _
       (lpPoint As POINTAPI) As Long
'Api zum Umrechnen der Bildschirm-Koordinaten in Control-Koordinaten
Public Declare Function ScreenToClient Lib "user32" _
       (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'Api zum Ermitteln der minimalen und maximalen Koordinaten eines Controls
Public Declare Function GetWindowPlacement Lib "user32" _
       (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As _
       Long
'Api zum Setzen der minimalen und maximalen Koordinaten eines Controls
Public Declare Function SetWindowPlacement Lib _
       "user32" (ByVal hwnd As Long, lpwndpl _
       As WINDOWPLACEMENT) As Long
'Api zum Ermitteln eines Fensterhandles
Public Declare Function FindWindow Lib "user32" Alias _
  "FindWindowA" (ByVal lpClassName As String, _
  ByVal lpWindowName As String) As Long
'Api zum Anzeigen eines Fensters
Public Declare Function ShowWindow Lib "user32" _
  (ByVal hwnd As Long, ByVal nCmdShow As Long) _
  As Long
'Api zum Hervorheben eines Fensters
Public Declare Function SetForegroundWindow Lib _
  "user32" (ByVal hwnd As Long) As Long
'API zum Ermitteln des hWnd eines Fensters
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'API zum Ermitteln der Länge des Titels eines Fensters
Private Declare Function GetWindowTextLength Lib "user32" _
        Alias "GetWindowTextLengthA" (ByVal hwnd As Long) _
        As Long
'API zum Ermitteln des Titels eines Fensters
Private Declare Function GetWindowText Lib "user32" Alias _
        "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString _
        As String, ByVal cch As Long) As Long
'API zum Ermitteln des Desktop-Handles
Private Declare Function GetDesktopWindow& Lib "user32" ()
'API zum Ermitteln des Handles des Controls, das den Fokus hat
Public Declare Function GetFocus Lib "user32" () As Long

'Konstanten für das Ermitteln des hWnd eines Fensters
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
'Soll das Fenster normal oder maximiert angezeigt werden ?
Public Const SW_NORMAL = &H1
Public Const SW_MAXIMIZE = &H3

Public Sub Main()
'--------------------------------------------------------------------'
'| Programmstart                                                    |'
'--------------------------------------------------------------------'
   'Oberstes Fenster ermitteln
   TaskHwnd = GetWindow(frm_splashscreen.hwnd, GW_HWNDFIRST)
   'Splashscreen wieder entladen, da er beim Ermitteln des obersten Fensters geladen wurde
   Unload frm_splashscreen
   
   'Schleife durch alle vorhandenen Fenster
   Do
      'Länge des Fenstertextes herausfinden
      TaskResult = GetWindowTextLength(TaskHwnd) + 1
      'Titel mit Leerzeichen füllen
      TaskTitle = Space$(TaskResult)
      'Titel auslesen
      TaskResult = GetWindowText(TaskHwnd, TaskTitle, TaskResult)
      'Titel in Variable schreiben
      TaskTitle = Left$(TaskTitle, Len(TaskTitle) - 1)
      
      'Wenn Titel den übergebenen Parameter enthält..
      If InStr(1, UCase(TaskTitle), UCase("NightGraphiX V1.0")) <> 0 Then
         'Fenster in den Vordergrund bringen
         SetForegroundWindow TaskHwnd
         'Fenster anzeigen
         ShowWindow TaskHwnd, SW_NORMAL
         'Programm beenden
         End
      End If
      
      'Nächstes Fensterhandle ermitteln
      TaskHwnd = GetWindow(TaskHwnd, GW_HWNDNEXT)
   'Solange wiederholen bis alle Handles abgefragt wurden
   Loop Until TaskHwnd = 0
   
   'Hauptfenster starten
   Load frm_nightgraphix
   frm_nightgraphix.Show
End Sub

Public Function CheckIfRunningInIDE() As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Prüfen, ob Programm in der IDE läuft                |'
'--------------------------------------------------------------------'
   'Bei einem Fehler zu ErrHandler gehen
   On Error GoTo ErrHandler

   'Diese Anweisung ergibt einen Fehler, wenn sie in der IDE ausgeführt wird
   Debug.Print 1 / 0
   
   'Programm läuft nicht in der IDE
   CheckIfRunningInIDE = False
   
   'Fehlerbehandlung nicht durchführen
   Exit Function
   
'Fehlerbehandlung
ErrHandler:
   'Programm läuft in der IDE
   CheckIfRunningInIDE = True
End Function
