Attribute VB_Name = "mdl_hookscrollwheel"
'_________________________________________________________________________________'
'|                          MODUL mdl_mausrad                                    |'
'| Dieses Modul beinhaltet Prozeduren zum Hooken der Form, um das Scrollrad      |'
'| der Maus abzufangen.                                                          |'
'---------------------------------------------------------------------------------'
'Variablen müssen deklariert werden
Option Explicit

'APIs..
Public Declare Function SetWindowsHookEx Lib "user32" Alias _
    "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, _
    ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Public Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, _
    ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" _
    (ByVal hHook As Long) As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) _
    As Long
    
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

'TYPEs..
Private Type MOUSEHOOKSTRUCT
  pt As POINTAPI
  hwnd As Long
  wHitTestCode As Long
  dwExtraInfo As Long
End Type

'KONSTANTEN..
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const MK_LBUTTON = &H1
Private Const MK_MBUTTON = &H10
Private Const MK_RBUTTON = &H2

Public Const WH_MOUSE = 7
Private Const WHEEL_DELTA = 120

Public Const GWL_WNDPROC = -4
Public MausPos As POINTAPI

'VARIABLEN..
Private hook As Long
Private nKeys As Long, Delta As Long, XPos As Long, YPos As Long
Private OriginalWindowProc As Long

'ENUMERATIONEN..
Public Enum mButtons
  LBUTTON = &H1
  MBUTTON = &H10
  RBUTTON = &H2
End Enum

Public Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, _
                          lParam As MOUSEHOOKSTRUCT) As Long
    Select Case nCode
      Case Is < 0
        MouseProc = CallNextHookEx(hook, nCode, wParam, lParam)
      Case 0
        If lParam.hwnd = frm_nightgraphix.hwnd Then
          Select Case wParam
            Case WM_MBUTTONDOWN
            Case WM_MBUTTONUP
          End Select
        End If
    End Select
End Function

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'--------------------------------------------------------------------'
'| Prozedur zum Auffangen der Nachrichten vom Hooken                |'
'--------------------------------------------------------------------'
   'Nachrichten auswerten
   Select Case uMsg
      'Wurde das Rad gescrollt..
      Case WM_MOUSEWHEEL
         'Daten des Scrollens auswerten
         nKeys = wParam And 65535
         Delta = wParam / 65536 / WHEEL_DELTA
         XPos = lParam And 65535
         YPos = lParam / 65536
         
         'Daten an die Prozedur MouseWheelRotation übergeben
         HookScrollWheel_MouseWheelRotation Delta, nKeys, XPos, YPos, hwnd
   End Select
   
   WindowProc = CallWindowProc(OriginalWindowProc, hwnd, uMsg, wParam, lParam)
End Function

Public Function HookScrollWheel_Init(Form As Form)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren des Hookens                          |'
'--------------------------------------------------------------------'
'###############HookScrollWheel_Exit nicht vergessen !!!#############!
   hook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, 0, _
                           GetCurrentThreadId)
   OriginalWindowProc = SetWindowLong(Form.hwnd, GWL_WNDPROC, _
                                      AddressOf WindowProc)
End Function

Public Function HookScrollWheel_Exit()
'--------------------------------------------------------------------'
'| Prozedur zum Beenden des Hookens                                 |'
'--------------------------------------------------------------------'
    UnhookWindowsHookEx hook
    SetWindowLong frm_nightgraphix.hwnd, GWL_WNDPROC, OriginalWindowProc
End Function

Public Function HookScrollWheel_MouseWheelRotation(Richtung As Long, Buttons As mButtons, x As Long, y As Long, hwnd As Long)
'--------------------------------------------------------------------'
'| Prozedur beim Scrollen des Mausrades                             |'
'--------------------------------------------------------------------'
   'Variable für den Punkt in der PicBox
   Dim Point_PicBox As POINTAPI
   
   'Wenn NG nicht verbunden ist und sich nicht im Demomodus befindet, beenden
   If Not Connected2Hardware And Not NGDemoModus Then Exit Function
   'Wenn die Maus sich in der Picturebox befindet..
   If HookScrollWheel_MausInBox() Then
      'Wenn größer gezoomt wird und pic_target nicht zu groß ist..
      If (Richtung = 1) And (frm_nightgraphix.pic_target.Width < frm_nightgraphix.pic_source.Width) Then
         'Position des Mauscursors ermitteln
         GetCursorPos Point_PicBox
         'Position des Mauszeigers in pic_target-Koordinaten umwandeln
         ScreenToClient frm_nightgraphix.pic_target.hwnd, Point_PicBox
         
         'Breite von pic_target vergrößern
         frm_nightgraphix.pic_target.Width = frm_nightgraphix.pic_target.Width + 200
         'Höhe von pic_target vergrößern
         frm_nightgraphix.pic_target.Height = frm_nightgraphix.pic_target.Height + 200
         'Modell neu von pic_source in pic_target zoomen
         Draw_Zoom frm_nightgraphix.pic_source, frm_nightgraphix.pic_target
         
         'Linken Rand von pic_target verschieben
         frm_nightgraphix.pic_target.Left = frm_nightgraphix.pic_target.Left - Point_PicBox.x / (frm_nightgraphix.pic_target.Width - 200) * 200
         'Oberen Rand von pic_target verschieben
         frm_nightgraphix.pic_target.Top = frm_nightgraphix.pic_target.Top - Point_PicBox.y / (frm_nightgraphix.pic_target.Height - 200) * 200
      'Wenn kleiner gezoomt wird und pic_target nicht zu klein ist..
      ElseIf (Richtung = -1) And (frm_nightgraphix.pic_target.Width > frm_nightgraphix.pic_rahmen.Width) Then
         'Position des Mauscursors ermitteln
         GetCursorPos Point_PicBox
         'Position des Mauszeigers in pic_target-Koordinaten umwandeln
         ScreenToClient frm_nightgraphix.pic_target.hwnd, Point_PicBox
         
         'Breite von pic_target verkleinern
         frm_nightgraphix.pic_target.Width = frm_nightgraphix.pic_target.Width - 200
         'Höhe von pic_target verkleinern
         frm_nightgraphix.pic_target.Height = frm_nightgraphix.pic_target.Height - 200
         'Modell neu von pic_source in pic_target zoomen
         Draw_Zoom frm_nightgraphix.pic_source, frm_nightgraphix.pic_target
         
         'Linken Rand von pic_target verschieben
         frm_nightgraphix.pic_target.Left = frm_nightgraphix.pic_target.Left + Point_PicBox.x / (frm_nightgraphix.pic_target.Width + 200) * 200
         'Oberen Rand von pic_target verschieben
         frm_nightgraphix.pic_target.Top = frm_nightgraphix.pic_target.Top + Point_PicBox.y / (frm_nightgraphix.pic_target.Height + 200) * 200
         
         'Wenn sich ein Rand von pic_target innerhalb von pic_rahmen befindet, wieder zurücksetzen
         If frm_nightgraphix.pic_target.Left + frm_nightgraphix.pic_target.Width < frm_nightgraphix.pic_rahmen.Width Then frm_nightgraphix.pic_target.Left = frm_nightgraphix.pic_rahmen.Width - frm_nightgraphix.pic_target.Width
         If frm_nightgraphix.pic_target.Top + frm_nightgraphix.pic_target.Height < frm_nightgraphix.pic_rahmen.Height Then frm_nightgraphix.pic_target.Top = frm_nightgraphix.pic_rahmen.Height - frm_nightgraphix.pic_target.Height
         If frm_nightgraphix.pic_target.Left > 0 Then frm_nightgraphix.pic_target.Left = 0
         If frm_nightgraphix.pic_target.Top > 0 Then frm_nightgraphix.pic_target.Top = 0
      End If
      
      'Wenn pic_target genauso groß wie pic_source ist..
      If frm_nightgraphix.pic_target.Width >= frm_nightgraphix.pic_source.Width Then
         'Scrollleisten sichtbar machen
         frm_nightgraphix.hsc_scrollen.Visible = True
         frm_nightgraphix.vsc_scrollen.Visible = True
         
         'Max-Wert der Scrollleisten einstellen
         frm_nightgraphix.vsc_scrollen.Max = frm_nightgraphix.pic_target.Height - frm_nightgraphix.pic_rahmen.Height
         'Max-Wert der Scrollleisten einstellen
         frm_nightgraphix.hsc_scrollen.Max = frm_nightgraphix.pic_target.Width - frm_nightgraphix.pic_rahmen.Width
         
         'Wert der Scrollleisten einstellen
         frm_nightgraphix.vsc_scrollen.Value = Abs(frm_nightgraphix.pic_target.Top)
         'Wert der Scrollleisten einstellen
         frm_nightgraphix.hsc_scrollen.Value = Abs(frm_nightgraphix.pic_target.Left)
         
         'Buttons sichtbar/unsichtbar machen
         frm_nightgraphix.tlb_toolbar.Buttons(9).Enabled = False
         frm_nightgraphix.tlb_toolbar.Buttons(10).Enabled = True
      'Wenn pic_target größer als pic_rahmen ist..
      ElseIf frm_nightgraphix.pic_target.Width > frm_nightgraphix.pic_rahmen.Width Then
         'Scrollleisten sichtbar machen
         frm_nightgraphix.hsc_scrollen.Visible = True
         frm_nightgraphix.vsc_scrollen.Visible = True
         
         'Max-Wert der Scrollleisten einstellen
         frm_nightgraphix.vsc_scrollen.Max = frm_nightgraphix.pic_target.Height - frm_nightgraphix.pic_rahmen.Height
         'Max-Wert der Scrollleisten einstellen
         frm_nightgraphix.hsc_scrollen.Max = frm_nightgraphix.pic_target.Width - frm_nightgraphix.pic_rahmen.Width
         
         'Wert der Scrollleisten einstellen
         frm_nightgraphix.vsc_scrollen.Value = Abs(frm_nightgraphix.pic_target.Top)
         'Wert der Scrollleisten einstellen
         frm_nightgraphix.hsc_scrollen.Value = Abs(frm_nightgraphix.pic_target.Left)
         
         'Buttons sichtbar machen
         frm_nightgraphix.tlb_toolbar.Buttons(9).Enabled = True
         frm_nightgraphix.tlb_toolbar.Buttons(10).Enabled = True
      'Wenn pic_target gleich groß wie pic_rahmen ist..
      Else
         'Scrollleisten unsichtbar machen
         frm_nightgraphix.hsc_scrollen.Visible = False
         frm_nightgraphix.vsc_scrollen.Visible = False
         'Buttons sichtbar/unsichtbar machen
         frm_nightgraphix.tlb_toolbar.Buttons(9).Enabled = True
         frm_nightgraphix.tlb_toolbar.Buttons(10).Enabled = False
      End If
      'Offset-Pfeil neu setzen
      Draw_SetOffsetArrow frm_nightgraphix.pic_arrowshow, Offset / Spalten * 360 \ 1
   End If
   
   'Wenn Bildimport-Picturebox sichtbar ist..
   If frm_nightgraphix.pic_import.Visible Then
      'Inhalt der Picbox neu zeichnen
      Draw_AlphaBlend frm_nightgraphix.pic_target, frm_nightgraphix.pic_alphablend, frm_nightgraphix.pic_import
   End If
End Function

Public Function HookScrollWheel_MausInBox() As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Prüfen, ob die Maus in der PictureBox ist           |'
'--------------------------------------------------------------------'
   'Linke, obere Ecke von pic_rahmen auf der Form ermitteln
   P1.x = frm_nightgraphix.ScaleX(frm_nightgraphix.pic_rahmen.Left, frm_nightgraphix.ScaleMode, vbPixels)
   P1.y = frm_nightgraphix.ScaleY(frm_nightgraphix.pic_rahmen.Top, frm_nightgraphix.ScaleMode, vbPixels)
   'Position auf der Form in absolute Position auf dem Bildschirm umwandeln
   ClientToScreen frm_nightgraphix.hwnd, P1
   
   'Rechte, untere Ecke von pic_rahmen auf der Form ermitteln
   P2.x = frm_nightgraphix.ScaleX(frm_nightgraphix.pic_rahmen.Left + frm_nightgraphix.pic_rahmen.Width, frm_nightgraphix.ScaleMode, vbPixels)
   P2.y = frm_nightgraphix.ScaleY(frm_nightgraphix.pic_rahmen.Top + frm_nightgraphix.pic_rahmen.Height, frm_nightgraphix.ScaleMode, vbPixels)
   'Position auf der Form in absolute Position auf dem Bildschirm umwandeln
   ClientToScreen frm_nightgraphix.hwnd, P2
   
   'Position des Mauscursors ermitteln
   GetCursorPos CurPos

   'Wenn sich der Mauscursor außerhalb von pic_rahmen befindet..
   If CurPos.x <= P1.x Or CurPos.x >= P2.x Or CurPos.y <= P1.y Or CurPos.y >= P2.y Then
      'Maus befindet sich nicht in der PicBox
      HookScrollWheel_MausInBox = False
   'Wenn sich der Mauscursor innerhalb von pic_rahmen befindet..
   Else
      'Maus befindet sich in der PicBox
      HookScrollWheel_MausInBox = True
   End If
End Function
