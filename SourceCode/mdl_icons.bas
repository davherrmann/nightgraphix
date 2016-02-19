Attribute VB_Name = "mdl_icons"
'_________________________________________________________________________________'
'|                               MODUL mdl_text                                  |'
'| Dieses Modul beinhaltet Routinen zum Setzen der Programmicons usw.            |'
'|                                                                               |'
'---------------------------------------------------------------------------------'
'Variablen müssen deklariert werden
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long

Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1
Private Const GWL_HWNDPARENT = (-8)

Public Sub Icons_Init(ByVal tmpForm As Form)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen der Programmicons                            |'
'--------------------------------------------------------------------'
'Vielen Dank an Jonathan Haas für einen Teil dieses Codes ! (ActiveVB)

   Dim nRet         As Long
   Dim nMainhWnd    As Long
   Dim B As Long, c As Long
   Dim big As Long, small As Long
   nRet = GetWindowLong(tmpForm.hwnd, GWL_HWNDPARENT)
   Do While nRet
      nMainhWnd = nRet
      nRet = GetWindowLong(nMainhWnd, GWL_HWNDPARENT)
   Loop
   'Wenn Programm in der IDE läuft..
   If CheckIfRunningInIDE Then
      'Icon aus Ico-Datei laden
      Call ExtractIconEx(FileSystem_ClearPath(App.Path) & "NG.ico", 0, big, small, 1)
   'Wenn Programm als kompilierte Exe läuft..
   Else
      'Icon direkt aus Exe-Datei laden
      Call ExtractIconEx(FileSystem_ClearPath(App.Path) & "NG1.0.exe", 0, big, small, 1)
   End If
   SendMessage nMainhWnd, WM_SETICON, ICON_SMALL, ByVal small 'Alt+Tab-icon setzen
   SendMessage nMainhWnd, WM_SETICON, ICON_BIG, ByVal big
   SendMessage tmpForm.hwnd, WM_SETICON, ICON_SMALL, ByVal small 'Fenstericon setzen
   SendMessage tmpForm.hwnd, WM_SETICON, ICON_BIG, ByVal big
End Sub
