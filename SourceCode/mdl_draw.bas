Attribute VB_Name = "mdl_draw"
'_________________________________________________________________________________'
'|                               MODUL mdl_draw                                  |'
'| Dieses Modul beinhaltet Graphik-Routinen, wie z.B. das Zeichnen des           |'
'| Hintergrundes.                                                                |'
'---------------------------------------------------------------------------------'

Option Explicit

'VARIABLEN..
'Farb- oder SW-Version ?
'Farbversion = True, SW-Version = False
Public RGBVersion As Boolean
'Variable für Schleifen
Private i As Integer
'Variablen für FillCell
Private lngX As Double
Private lngY As Double
'Verhältnis zwischen Target und Source
Public ZoomX As Double
Public ZoomY As Double
'Verhältnis zwischen Gegenkathete und Hypothenuse
Private RatioGkHyp As Double
'Verhältnis zwischen Ankathete und Hypothenuse
Private RatioAkHyp As Double
'Radius des OffsetArrows
Public RadiusOffsetArrow
'In welchem Winkel ist der Offset-Pfeil ?
Public RotateDegree As Double
'Position des OffsetArrows
Private PosOffsetArrow As POINTAPI
'Variablen für ImportPicture
Private XSource As Long
Private YSource As Long
'Variablen für Schleifen..
Private x As Integer
Private y As Integer
Private Z As Integer
'Color-Variable
Public Col As Long
'Variablen für ExtFloodFill
Private lngBrush As Long
Private lngCol As Long
'Winkel fürs Zeichnen
Private sngval As Double
'Type für das Speichern von RGB-Werten
Public Type tRGB
   R As Integer
   G As Integer
   B As Integer
End Type
'RGB-Type für SW-Umwandlung
Private SWRGB As tRGB
'RGB-Type für Array_Anpassen
Public ARRAYRGB As tRGB
'Enumeration für die Werkzeuge
Public Enum eTool
   Pencil = 0
   LEDCircle = 1
   Text = 2
   ChooseColor = 3
   ImportPicture = 4
   InvertPicture = 5
   ClearPicture = 6
End Enum
'Welches Werkzeug wird benutzt ?
Public Tool As eTool
'Welches Werkzeug wurde als letztes benutzt ?
Public LastTool As eTool
'Ist der Cursor sichtbar ?
Public CursorVisible As Boolean
'Position des Cursors
Public CursorPosition As POINTAPI
'Höhe des Cursors
Public CursorHeight As Integer
'Wieviel Zeichen wurden schon geschrieben
Public ZeichenSpalte As Integer
'Größe des Pinsels
Public DrawSize As Integer
'Soll das Zeichnen gesperrt werden ?
Public LockDraw As Boolean

'KONSTANTEN..
'Korrekturwert für das Zeichnen des Modells
Public Const KorrekturWert = 0.02
'Schwellwert für das Umwandeln von Farbe in SW
Private Const Schwellwert = 254
'Masken für die Umrechnung von Long-Farbwerten nach RGB
Public Mask_R As Long
Public Mask_G As Long
Public Mask_B As Long

'API-DEKLARATIONEN..
'Füllen von Bereichen
Public Declare Function ExtFloodFill Lib "gdi32.dll" ( _
     ByVal hdc As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal crColor As Long, _
     ByVal wFillType As Long) As Long
'Erstellen von Brushs
Public Declare Function CreateSolidBrush Lib "gdi32.dll" ( _
     ByVal crColor As Long) As Long
'Objekt auswählen
Public Declare Function SelectObject Lib "gdi32.dll" ( _
     ByVal hdc As Long, _
     ByVal hObject As Long) As Long
'Objekt entfernen
Public Declare Function DeleteObject Lib "gdi32.dll" ( _
     ByVal hObject As Long) As Long
'Bestimmtes Pixel auslesen
Public Declare Function GetPixel Lib "gdi32.dll" ( _
     ByVal hdc As Long, _
     ByVal x As Long, _
     ByVal y As Long) As Long
'Bereich eines Bildes strecken
Public Declare Function StretchBlt Lib "gdi32" ( _
                 ByVal hdc As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal nSrcWidth As Long, _
                 ByVal nSrcHeight As Long, _
                 ByVal dwRop As Long) As Long
'Kopieren eines Bildes
Public Declare Function BitBlt Lib "gdi32" (ByVal _
       NachHdc As Long, ByVal x As Long, ByVal _
       y As Long, ByVal w As Long, ByVal h As Long, ByVal _
       vonHdc As Long, ByVal vonX As Long, ByVal _
       vonY As Long, ByVal Modus As Long) As Long
'Api zum Drehen eines Bildes in einem beliebigen Winkel
'X-, Y- oder Z-Achse, Zoomen möglich
Private Declare Function PlgBlt Lib "gdi32.dll" ( _
         ByVal hdcDest As Long, _
         lpPoint As POINTAPI, _
         ByVal hDCSrc As Long, _
         ByVal nXSrc As Long, _
         ByVal nYSrc As Long, _
         ByVal nWidth As Long, _
         ByVal nHeight As Long, _
         ByVal hbmMask As Long, _
         ByVal xMask As Long, _
         ByVal yMask As Long) As Long
' AlphaBlend: NT/2000/XP/Vista: Included in Windows 2000 and later.
Public Declare Function AlphaBlend& Lib "msimg32.dll" (ByVal hdcDest&, ByVal XDest&, _
  ByVal YDest&, ByVal WidthDest&, ByVal HeightDest&, ByVal hDCSrc&, ByVal xSrc&, _
  ByVal ySrc&, ByVal WidthSrc&, ByVal HeightSrc&, ByVal Blendfunc&)

'Konstanten für API-Funktionen
Public Const SRCCOPY = &HCC0020
Public Const FLOODFILLSURFACE As Long = 1
Public Const FLOODFILLBORDER As Long = 0
'Konstanten für BitBlt
Public Const BIT_COPY = &HCC0020
Public Const BIT_AND = &H8800C6
Public Const BIT_Invert = &H660046
Public Const BIT_BLACK = &H42&

'TYPEs..
'POINTAPI: Speichern von X- und Y-Werten
Public Type POINTAPI
   x As Long
   y As Long
End Type
'RECT: Speichern von Left, Top, Right und Bottom-Werten
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
'WINDOWPLACEMENT: Daten der maximalen und minimalen Größe
Public Type WINDOWPLACEMENT
  Length As Long
  flags As Long
  showCmd As Long
  ptMinPosition As POINTAPI
  ptMaxPosition As POINTAPI
  rcNormalPosition As RECT
End Type

Public Sub Draw_Modell(picsource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Zeichnen des Modells                                |'
'--------------------------------------------------------------------'
   'Wenn RGB-Version..
   If RGBVersion Then
      'Hintergrund dunkel
      picsource.BackColor = RGB(60, 60, 60)
   'Wenn SW-Version
   Else
      'Hintergrund weiß
      picsource.BackColor = vbWhite
   End If

   'Kreise zeichnen
   For i = Hub To Leds + Hub
      picsource.Circle (picsource.Height / 2, picsource.Width / 2), SpaltenAbstand * i, RGB(192, 192, 192)
   Next i
   
   'Linien zeichnen
   For i = 1 To Spalten
      sngval = i * 2 * PI / Spalten
      picsource.Line (picsource.Height / 2 + (Hub) * SpaltenAbstand * Cos(sngval), picsource.Width / 2 + (Hub) * SpaltenAbstand * Sin(sngval))-(picsource.Height / 2 + (Leds + Hub + KorrekturWert) * SpaltenAbstand * Cos(sngval), picsource.Width / 2 + (Leds + Hub + KorrekturWert) * SpaltenAbstand * Sin(sngval)), RGB(192, 192, 192)
   Next i
End Sub

Public Sub Draw_Background(picsource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Füllen des Hintergrundes                            |'
'--------------------------------------------------------------------'
   With picsource
      'Außenkreis grau füllen
      'Brush erstellen
      lngBrush = CreateSolidBrush(RGB(160, 160, 160))
      'Pixel auslesen
      lngCol = GetPixel(.hdc, 1, 1)
      'PicBox auswählen
      SelectObject .hdc, lngBrush
      'Floodfill anwenden
      ExtFloodFill .hdc, 1, 1, lngCol, FLOODFILLSURFACE
      'Brush löschen
      DeleteObject lngBrush
      
      'Innenkreis grau füllen
      'Brush erstellen
      lngBrush = CreateSolidBrush(RGB(160, 160, 160))
      'Pixel auslesen
      lngCol = GetPixel(.hdc, .ScaleWidth / 2, .ScaleHeight / 2)
      'PicBox auswählen
      SelectObject .hdc, lngBrush
      'Floodfill anwenden
      ExtFloodFill .hdc, .ScaleWidth / 2, .ScaleHeight / 2, lngCol, FLOODFILLSURFACE
      'Brush löschen
      DeleteObject lngBrush
   End With
End Sub

Public Sub Draw_Refresh(picbox As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Refreshen der Picturebox                            |'
'--------------------------------------------------------------------'
   'Zeit:
   'Ohne Visible-Schalter:    9-13 Milisekunden (Messung)
   'Mit  Visible-Schalter:    ~1,5 Milisekunden (gemessener Höchstwert: 0,4 ms, 1-2%)
   
   'Picturebox unsichtbar machen (Performance-Steigerung)
   picbox.Visible = False
   'picbox refreshen
   picbox.Refresh
   'Picturebox wieder anzeigen
   picbox.Visible = True
End Sub

Public Sub Draw_Redraw(picsource As PictureBox, pictarget As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Neuzeichnen des Modells                             |'
'--------------------------------------------------------------------'
   'PicBox löschen
   picsource.Cls
   
   'Modell neu zeichnen
   Draw_Modell picsource
   
   'Background füllen
   Draw_Background picsource
   
   'PicBox refreshen und Zoomen
   Draw_Zoom picsource, pictarget
End Sub

Public Sub Draw_Circle(picsource As PictureBox, pictarget As PictureBox, LEDCircle As Integer, tmpCol As Long)
'--------------------------------------------------------------------'
'| Prozedur zum Füllen eines Kreises bei MausKlick                  |'
'--------------------------------------------------------------------'
   'Zeichnen soll gesperrt werden
   LockDraw = True
   
   'Schleife durch alle Spalten
   For i = 1 To Spalten
      'Kreis in pic_source füllen
      Draw_FillCell i, LEDCircle, picsource, tmpCol, True
      'DoEvents
      DoEvents
   Next
   
   'Zeichnen soll erlaubt werden
   LockDraw = False
   
   'Bild von pic_source nach pic_target übertragen
   Draw_Zoom picsource, pictarget
End Sub

Public Sub Draw_ZoomArea(ByRef x As Long, ByRef y As Long, ByRef Width As Long, ByRef Height As Long, ByRef tmpPicSource As PictureBox, ByRef tmpPicTarget As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Übertragen eines Teilbildes von Source nach Target  |'
'--------------------------------------------------------------------'
   'Teilgebiet des Bildes übertragen und Zoomen
   StretchBlt tmpPicTarget.hdc, x, y, Width, Height, tmpPicSource.hdc, x * ZoomX, y * ZoomY, Width * ZoomX, Height * ZoomY, SRCCOPY
   'pictarget refreshen
   Draw_Refresh tmpPicTarget
End Sub

Public Sub Draw_Zoom(picsource As PictureBox, pictarget As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Übertragen des Bildes von picsource nach pictarget  |'
'--------------------------------------------------------------------'
   'Zeit:
   '1-faches StretchBlt:      ~205 Milisekunden
   '4-faches StretchBlt:      ~224 Milisekunden
   'Erste Variante:
   'StretchBlt pictarget.hdc, 0, 0, pictarget.Width, pictarget.Height, picsource.hdc, 0, 0, picsource.Width, picsource.Height, SRCCOPY
   'Zweite Variante wird benutzt, da Mauszeiger dann nicht stehen bleibt:
   'On Error Resume Next
   
   'Bild übertragen und Zoomen:
   'Alle vier Bildviertel einzeln strecken und übertragen
   'Erstes Viertel (0, 0)
   StretchBlt pictarget.hdc, 0, 0, pictarget.Width / 2, pictarget.Height / 2, picsource.hdc, 0, 0, picsource.Width / 2, picsource.Height / 2, SRCCOPY
   'Zweites Viertel (1, 0)
   StretchBlt pictarget.hdc, pictarget.Width / 2, 0, pictarget.Width / 2, pictarget.Height / 2, picsource.hdc, picsource.Width / 2, 0, picsource.Width / 2, picsource.Height / 2, SRCCOPY
   'Drittes Viertel (1, 1)
   StretchBlt pictarget.hdc, pictarget.Width / 2, pictarget.Height / 2, pictarget.Width / 2, pictarget.Height / 2, picsource.hdc, picsource.Width / 2, picsource.Height / 2, picsource.Width / 2, picsource.Height / 2, SRCCOPY
   'Viertes Viertel (0, 1)
   StretchBlt pictarget.hdc, 0, pictarget.Height / 2, pictarget.Width / 2, pictarget.Height / 2, picsource.hdc, 0, picsource.Height / 2, picsource.Width / 2, picsource.Height / 2, SRCCOPY
   
   'pictarget refreshen
   Draw_Refresh pictarget
   
   'Verhältnis zwischen Source und Target ausrechnen
   ZoomX = picsource.Width / pictarget.Width
   ZoomY = picsource.Height / pictarget.Height
End Sub

Public Sub Draw_Click(picsource As PictureBox, pictarget As PictureBox, x As Single, y As Single)
'--------------------------------------------------------------------'
'| Prozedur zum Füllen von Bereichen bei MausKlicks                 |'
'--------------------------------------------------------------------'
   'Zeichnen soll gesperrt werden
   LockDraw = True

   'Verhältnis zwischen Source und Target ausrechnen
   ZoomX = picsource.Width / pictarget.Width
   ZoomY = picsource.Height / pictarget.Height
   
   With picsource
      'Pixel auslesen
      lngCol = GetPixel(picsource.hdc, x * ZoomX, y * ZoomY)
      'Farbe herausfinden
      Col = Draw_GetColor(lngCol, True)
      
      'Ist die Farbe gültig (<> grau) ?
      If Col <> -1 Then
         'Bereich füllen
         Draw_FillArea x * ZoomX, y * ZoomY, picsource, Col
         'Änderungen wurden nicht gespeichert
         FileSystem_SavedChanges False
         'DoEvents
         DoEvents
         'Array anpassen
         Array_Anpassen x * ZoomX, y * ZoomY, Col
      End If
   End With
   'Zeichnen soll erlaubt werden
   LockDraw = False
End Sub

Public Function Number2Color(tmpnumber As Integer) As Long
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben der Farbe bei einem Index von Opt-Button|'
'--------------------------------------------------------------------'

   'Je nach tmpnumber Farbe wählen
   Select Case tmpnumber
      Case 0
         'Farbe: Rot
         Number2Color = vbRed
      Case 1
         'Farbe: Gelb
         Number2Color = vbYellow
      Case 2
         'Farbe: Grün
         Number2Color = vbGreen
      Case 3
         'Farbe: Blau
         Number2Color = vbBlue
      Case 4
         'Farbe: Magenta
         Number2Color = vbMagenta
      Case 5
         'Farbe: Cyan
         Number2Color = vbCyan
      Case 6
         'Farbe: Weiß (löschen)
         Number2Color = vbWhite
      Case 7
         'Farbe: Schwarz (dunkelgrau)
         'RGB(60,60,60), da dunkelgrau genommen werden soll
         Number2Color = RGB(60, 60, 60)
   End Select
End Function

Public Function Color2Number(tmpcolor As Long) As Integer
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben des Index vom Opt-Button aus einer Farbe|'
'--------------------------------------------------------------------'

   'Je nach tmpnumber Farbe wählen
   Select Case tmpcolor
      Case vbRed
         'Farbe: Rot
         Color2Number = 0
      Case vbYellow
         'Farbe: Gelb
         Color2Number = 1
      Case vbGreen
         'Farbe: Grün
         Color2Number = 2
      Case vbBlue
         'Farbe: Blau
         Color2Number = 3
      Case vbMagenta
         'Farbe: Magenta
         Color2Number = 4
      Case vbCyan
         'Farbe: Cyan
         Color2Number = 5
      Case vbWhite
         'Farbe: Weiß (löschen)
         Color2Number = 6
      Case RGB(60, 60, 60)
         'Farbe: Schwarz (dunkelgrau)
         'RGB(60,60,60), da dunkelgrau genommen werden soll
         Color2Number = 7
   End Select
End Function

Public Sub Draw_ZoomDown(pictarget As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum kleiner Zoomen des Modells                          |'
'--------------------------------------------------------------------'
   'Kleiner Zoomen
   'Bildbreite verkleinern
   pictarget.Width = frm_nightgraphix.pic_rahmen.Width
   'Bildhöhe verkleinern
   pictarget.Height = frm_nightgraphix.pic_rahmen.Height
End Sub

Public Sub Draw_ZoomUp(pictarget As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum größer Zoomen des Modells                          |'
'--------------------------------------------------------------------'
   'Größer zoomen
   'Bildbreite vergrößern
   pictarget.Width = frm_nightgraphix.pic_source.Width
   'Bildhöhe vergrößern
   pictarget.Height = frm_nightgraphix.pic_source.Height
End Sub

Public Function Draw_GetColor(ByVal lngCol As Long, ByVal WhiteIfColored As Boolean) As Long
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben der Farbe aus dem Long-Wert             |'
'--------------------------------------------------------------------'
   'Wenn RGB-Version..
   If RGBVersion Then
      'Wenn das Feld dunkelgrau ist, die aktuelle Farbe wählen
      If ((lngCol = RGB(60, 60, 60)) Or (lngCol = 3750969)) Then Draw_GetColor = LEDColor: Exit Function
      
      'Wenn das Feld rot, gelb, grün, blau, cyan, magenta oder schwarz ist, dann
      'weiß füllen
      If (lngCol = RGB(255, 0, 0)) Or (lngCol = RGB(0, 255, 0)) Or (lngCol = RGB(0, 0, 255)) Or (lngCol = RGB(255, 255, 0)) Or (lngCol = RGB(255, 0, 255)) Or (lngCol = RGB(0, 255, 255)) Or (lngCol = RGB(60, 60, 60)) Or (lngCol = 3750969) Or (lngCol = RGB(255, 255, 255)) Then
         'Wenn Feld weiß gefüllt werden soll, wenn es farbig ist..
         If WhiteIfColored And (LEDColor = lngCol) Then
            'Schwarz zurückgeben
            Draw_GetColor = RGB(60, 60, 60)
         'Wenn das Feld immer mit aktueller Farbe gefüllt werden soll..
         Else
            'Aktuelle Farbe zurückgeben
            Draw_GetColor = LEDColor: Exit Function
         End If
      'Wenn das Feld eine andere Farbe hat..
      Else
         'Feld nicht füllen (-1 zurückgeben)
         Draw_GetColor = -1: Exit Function
      End If
   'Wenn SW-Version..
   Else
      'Wenn das Feld weiß ist, die aktuelle Farbe wählen
      If ((lngCol = RGB(255, 255, 255)) Or (lngCol = 16711422)) Then Draw_GetColor = LEDColor: Exit Function
      
      'Wenn das Feld rot, gelb, grün, blau, cyan, magenta oder schwarz ist, dann
      'weiß füllen
      If (lngCol = RGB(255, 0, 0)) Or (lngCol = RGB(0, 255, 0)) Or (lngCol = RGB(0, 0, 255)) Or (lngCol = RGB(255, 255, 0)) Or (lngCol = RGB(255, 0, 255)) Or (lngCol = RGB(0, 255, 255)) Or (lngCol = RGB(60, 60, 60)) Or (lngCol = 3750969) Then
         'Wenn Feld weiß gefüllt werden soll, wenn es farbig ist..
         If WhiteIfColored = True Then
            'Weiß zurückgeben
            Draw_GetColor = vbWhite
         'Wenn das Feld immer mit aktueller Farbe gefüllt werden soll..
         Else
            'Aktuelle Farbe zurückgeben
            Draw_GetColor = LEDColor: Exit Function
         End If
      'Wenn das Feld eine andere Farbe hat..
      Else
         'Feld nicht füllen (-1 zurückgeben)
         Draw_GetColor = -1: Exit Function
      End If
   End If
End Function

Public Function Draw_Color2SW(tmpCol As Long) As Byte
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben von Schwarz oder Weiß aus Farbe         |'
'--------------------------------------------------------------------'
   'Wennn Schwarz zurückgegeben werden soll, wird Draw_Color2SW
   'auf False gesetzt, bei weiß auf True
   
   'RGB-Wert der übergebenen Farbe errechnen
   SWRGB = Draw_Color2RGB_Bool(tmpCol)
   
   'Wenn Mittelwert größer als Schwellwert ist, dann Weiß zurückgeben, sonst schwarz
   If (SWRGB.R + SWRGB.G + SWRGB.B) / 3 < 255 Then             'Schwellwert Then
      Draw_Color2SW = 1
   Else
      Draw_Color2SW = 0
   End If
End Function

Public Function Draw_Color2RGB_Bool(tmpCol As Long) As tRGB
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben von RGB-Werten einer Long-Farbe         |'
'--------------------------------------------------------------------'
   'RGB-Werte errechnen und zurückgeben
   'Rot-Wert
   Draw_Color2RGB_Bool.R = (tmpCol And Mask_R) / 1
   'Grün-Wert
   Draw_Color2RGB_Bool.G = (tmpCol And Mask_G) / &H100&
   'Blau-Wert
   Draw_Color2RGB_Bool.B = (tmpCol And Mask_B) / &H10000
   
   'Werte auf 255 oder 0 runden
   If Draw_Color2RGB_Bool.R > 128 Then
      Draw_Color2RGB_Bool.R = 255
   Else
      Draw_Color2RGB_Bool.R = 0
   End If
   'Werte auf 255 oder 0 runden
   If Draw_Color2RGB_Bool.G > 128 Then
      Draw_Color2RGB_Bool.G = 255
   Else
      Draw_Color2RGB_Bool.G = 0
   End If
   'Werte auf 255 oder 0 runden
   If Draw_Color2RGB_Bool.B > 128 Then
      Draw_Color2RGB_Bool.B = 255
   Else
      Draw_Color2RGB_Bool.B = 0
   End If
End Function

Public Function Draw_Color2RGB_Int(tmpCol As Long) As tRGB
'--------------------------------------------------------------------'
'| Prozedur zum Zurückgeben von RGB-Werten einer Long-Farbe         |'
'--------------------------------------------------------------------'
   'RGB-Werte errechnen und zurückgeben
   'Rot-Wert
   Draw_Color2RGB_Int.R = (tmpCol And Mask_R) / 1
   'Grün-Wert
   Draw_Color2RGB_Int.G = (tmpCol And Mask_G) / &H100&
   'Blau-Wert
   Draw_Color2RGB_Int.B = (tmpCol And Mask_B) / &H10000
End Function

Public Sub Draw_FillCell(ByVal x As Integer, ByVal y As Integer, picsource As PictureBox, Col As Long, UpdateArray As Boolean)
'--------------------------------------------------------------------'
'| Prozedur zum Füllen einer bestimmten Zelle                       |'
'--------------------------------------------------------------------'
   'Wenn ungültige Werte übergeben werden, Prozedur schließen
   If Not ((y >= 1) And (y <= Leds)) Then Exit Sub
   
   'Variablen übernehmen
   lngX = x
   lngY = y
   
   'Zellkoordinaten holen
   x = Maths_GetCellPosition(lngX, lngY, picsource).x
   y = Maths_GetCellPosition(lngX, lngY, picsource).y
   
   'Bereich in picsource füllen
   Draw_FillArea x, y, picsource, Col
   
   'Array anpassen
   If UpdateArray = True Then
      'Array anpassen
      ARRAYRGB = Draw_Color2RGB_Bool(Col)
      If RGB(ARRAYRGB.R, ARRAYRGB.G, ARRAYRGB.B) = 0 Then
         ARRAYRGB.R = 0
         ARRAYRGB.G = 0
         ARRAYRGB.B = 0
      End If
      'RGB-Arrays updaten
      Array_Red(lngX, lngY) = ARRAYRGB.R / 255
      Array_Green(lngX, lngY) = ARRAYRGB.G / 255
      Array_Blue(lngX, lngY) = ARRAYRGB.B / 255
      'Wenn Farbe weiß ist..
      If Draw_Color2SW(Col) = False Then
         'Eine 0 ins Array schreiben
         Array_SW(lngX, lngY) = 0
      'Wenn Farbe nicht weiß ist..
      Else
         'Eine 1 ins Array schreiben
         Array_SW(lngX, lngY) = 1
      End If
   End If
End Sub

Public Sub Draw_InitFillArea()
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren von FillArea                         |'
'--------------------------------------------------------------------'
   'Brush-Objekte für alle Farben erstellen
   'Rot:
   Array_RGBBrush(1, 0, 0) = CreateSolidBrush(vbRed)
   'Grün:
   Array_RGBBrush(0, 1, 0) = CreateSolidBrush(vbGreen)
   'Blau:
   Array_RGBBrush(0, 0, 1) = CreateSolidBrush(vbBlue)
   'Gelb:
   Array_RGBBrush(1, 1, 0) = CreateSolidBrush(vbYellow)
   'Magenta:
   Array_RGBBrush(1, 0, 1) = CreateSolidBrush(vbMagenta)
   'Cyan:
   Array_RGBBrush(0, 1, 1) = CreateSolidBrush(vbCyan)
   'Weiß:
   Array_RGBBrush(1, 1, 1) = CreateSolidBrush(vbWhite)
   'Schwarz:
   Array_RGBBrush(0, 0, 0) = CreateSolidBrush(RGB(60, 60, 60))
End Sub

Public Sub Draw_ReleaseFillArea()
'--------------------------------------------------------------------'
'| Prozedur zum Freigeben der Brush-Objekte                         |'
'--------------------------------------------------------------------'
   'Schleife durch das dreidimensionale Array
   For x = 0 To 1
      For y = 0 To 1
         For Z = 0 To 1
            'Brush-Objekt freigeben
            DeleteObject Array_RGBBrush(x, y, Z)
   Next Z, y, x
End Sub

Public Sub Draw_FillArea(ByVal x As Double, ByVal y As Double, ByRef picbox As PictureBox, tmpCol As Long)
'--------------------------------------------------------------------'
'| Prozedur zum Füllen von Bereichen                                |'
'--------------------------------------------------------------------'
   'Gültigkeit der Farbe prüfen
   If tmpCol = -1 Then Exit Sub
   
   'Farbe in 3-Bit umwandeln
   ARRAYRGB = Draw_Color2RGB_Bool(tmpCol)
   tmpCol = RGB(ARRAYRGB.R, ARRAYRGB.G, ARRAYRGB.B)
   
   'Bereich in PictureBox füllen
   With picbox
      'PictureBox auswählen, mit zuvor erstelltem Brush
      SelectObject .hdc, Array_RGBBrush(ARRAYRGB.R \ 255, ARRAYRGB.G \ 255, ARRAYRGB.B \ 255)
      'Pixel auslesen
      lngCol = GetPixel(.hdc, x, y)
      'Floodfill ausführen
      ExtFloodFill .hdc, x, y, lngCol, FLOODFILLSURFACE
   End With
End Sub

Public Function Draw_GetCell(ByVal x As Integer, ByVal y As Integer, picbox As PictureBox) As Long
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen einer bestimmten Zelle                     |'
'--------------------------------------------------------------------'
   'Hub addieren
   y = y + Hub
   
   'Winkel berechnen
   Winkel = 2 * PI - (((x / Spalten) * 360 - 0.5) / 180 * PI)
   
   'X und Y umrechnen in geometrische Daten
   x = picbox.Height / 2 - SpaltenAbstand * (y - 0.5) * Sin(Winkel)
   y = picbox.Width / 2 - SpaltenAbstand * (y - 0.5) * Cos(Winkel)

   'Zelle auslesen
   Draw_GetCell = GetPixel(picbox.hdc, x, y)
End Function

Public Sub Draw_ImportPicture(ByRef picbild As PictureBox, ByRef picsource As PictureBox, ByRef picimport As PictureBox, ByRef pictarget As PictureBox, progbar As Object)
'--------------------------------------------------------------------'
'| Prozedur zum Importieren eines Bildes                            |'
'--------------------------------------------------------------------'
   'Das Bild befindet sich bereits in picbild
   'Modell neu zeichnen
   'Draw_Redraw picsource, pictarget
   
   'ProgressBar einstellen
   progbar.Max = Leds * CLng(Spalten) / 2
   
   'Schleife liest an jeder Position einer Zelle im Bild
   'ein Pixel aus, es werden also LEDs*Spalten Pixel ausgelesen
   For lngY = 1 To Leds
      For lngX = 1 To Spalten
         
         'DoEvents
         DoEvents
         
         'X und Y-Wert umrechnen
         XSource = Maths_GetCellPosition(lngX, lngY, picsource).x '/ picbild.Width * picsource.Width
         YSource = Maths_GetCellPosition(lngX, lngY, picsource).y '/ picbild.Height * picsource.Height
         
         'Wenn Pixel nicht im Bereich des Bildes liegt, Schleife überspringen
         If (XSource < (picimport.Left - pictarget.Left) * ZoomX) Or (XSource > (picimport.Left - pictarget.Left + picimport.Width) * ZoomX) Or (YSource < (picimport.Top - pictarget.Top) * ZoomY) Or (YSource > (picimport.Top - pictarget.Top + picimport.Height) * ZoomY) Then GoTo Jump
            
         'Prüfen, ob in picsource Feld weiß ist ..
         lngCol = GetPixel(picsource.hdc, XSource, YSource)
         
         'Farbe prüfen
         lngCol = Draw_GetColor(lngCol, False)
         If lngCol = -1 Then GoTo Jump
         
         'Farbe des Pixels auslesen
         'Col = Draw_GetCell(lngX, lngY, picbild)
         Col = GetPixel(picbild.hdc, XSource / ZoomX - (picimport.Left - pictarget.Left), YSource / ZoomY - (picimport.Top - pictarget.Top))
         
         'Farbe in 8-Bit umwandeln
         ARRAYRGB = Draw_Color2RGB_Bool(Col)
         Col = RGB(ARRAYRGB.R, ARRAYRGB.G, ARRAYRGB.B)
         
         'Wenn es eine S/W-Version ist..
         If Not RGBVersion Then
            'Wenn Feld weiß werden soll und vorher schon farbig war..
            If (Not Draw_Color2SW(Col)) And Draw_Color2SW(Draw_GetCell(lngX, lngY, picsource)) Then
               'Prozedur beenden
               GoTo Jump
            End If
            'Feld schwarz oder weiß füllen
            Draw_FillCell lngX, lngY, picsource, IIf(Draw_Color2SW(Col), LEDColor, vbWhite), True
         'Wenn es eine RGB-Version ist..
         Else
            'Wenn Zelle die Transparenzfarbe bekommen soll und Transparenz an ist..
            If (Col = NGTransparentColor) And NGTransparentImport Then
               'Prozedur beenden
               GoTo Jump
            End If
            'Feld farbig füllen
            Draw_FillCell lngX, lngY, picsource, Col, True
         End If
         
'Bei ungültiger Farbe hierher springen
Jump:
         'ProgressBar einstellen
         progbar.Value = CLng(((lngY - 1) * Spalten + lngX) / 2)
      Next lngX
   Next lngY
   
   'ProgressBar zurückstellen
   progbar.Value = 0
   
   'Bild refreshen
   Draw_Zoom picsource, pictarget
   
   'Einen Stern (ungespeichert) in Formtitel einfügen
   FileSystem_SavedChanges False
End Sub

Public Sub Draw_InitCursor(ByVal x As Integer, ByVal y As Integer, ByRef picsource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren des Cursors für die Texteingabe      |'
'--------------------------------------------------------------------'
   'Bei einem Fehler weitermachen..
   On Error Resume Next
   
   'Array für Felder unter dem Cursor anpassen
   ReDim Array_CursorRGB(0 To Letter.FontHeight - 1)
   'Array für Felder des Cursors anpassen
   ReDim Array_CursorSW(0 To Letter.FontHeight - 1)
   'Array für Cursor-Positionen anpassen
   ReDim Array_CursorPos(0 To Letter.FontHeight - 1)
   
   'Zellen auf dem Modell abfragen
   For i = y To y - Letter.FontHeight Step -1
      'Position der Zelle für den Cursor speichern
      Array_CursorPos(i - y + Letter.FontHeight) = Maths_GetCellPosition(x, i + 1, picsource)
      'Zellen auslesen und im Array speichern
      Array_CursorRGB(i - y + Letter.FontHeight) = GetPixel(picsource.hdc, Array_CursorPos(i - y + Letter.FontHeight).x, Array_CursorPos(i - y + Letter.FontHeight).y)
      'Wenn Hintergrund schwarz ist..
      If Array_CursorRGB(i - y + Letter.FontHeight) = 3750969 Then
         'Cursor soll weiß sein
         Array_CursorSW(i - y + Letter.FontHeight) = vbWhite
      'Wenn Hintergrund nicht schwarz ist..
      Else
         'Cursor soll schwarz sein
         Array_CursorSW(i - y + Letter.FontHeight) = vbBlack
      End If
   Next i
End Sub

Public Sub Draw_Cursor(ByVal x As Integer, ByVal y As Integer, ByRef picsource As PictureBox, Draw As Boolean)
'--------------------------------------------------------------------'
'| Prozedur zum Zeichnen/Löschen des Cursors für die Texteingabe    |'
'--------------------------------------------------------------------'
   'Wenn Cursor gelöscht werden soll..
   If Draw = False Then
      For i = y To y - Letter.FontHeight + 1 Step -1
         'Ist die Farbe gültig (<> grau) ?
         If Draw_GetColor(Array_CursorRGB(i - y + Letter.FontHeight - 1), True) <> -1 Then
            'Zellen mit ursprünglichem Inhalt füllen
            Draw_FillArea Array_CursorPos(i - y + Letter.FontHeight - 1).x, Array_CursorPos(i - y + Letter.FontHeight - 1).y, picsource, Array_CursorRGB(i - y + Letter.FontHeight - 1)
         End If
      Next i
   'Wenn Cursor gezeichnet werden soll..
   Else
      For i = y To y - Letter.FontHeight + 1 Step -1
         'Ist die Farbe gültig (<> grau) ?
         If Draw_GetColor(Array_CursorRGB(i - y + Letter.FontHeight - 1), True) <> -1 Then
            'Zellen mit Cursor füllen
            Draw_FillArea Array_CursorPos(i - y + Letter.FontHeight - 1).x, Array_CursorPos(i - y + Letter.FontHeight - 1).y, picsource, Array_CursorSW(i - y + Letter.FontHeight - 1)
         End If
      Next i
   End If
End Sub

Public Sub Draw_RotatePicture(ByRef tmpPicSource As PictureBox, ByRef tmpPicRotate As PictureBox, ByVal Degree As Integer)
'--------------------------------------------------------------------'
'| Prozedur zum Drehen eines Bildes um eine bestimmte Gradzahl      |'
'--------------------------------------------------------------------'
   Dim x As Integer
   Dim NewX As Integer, NewY As Integer
   Dim SinAng1, CosAng1
   Dim PtList(2) As POINTAPI
   
   'Punktliste zurücksetzen
   PtList(0).x = -(tmpPicSource.ScaleWidth / 2)
   PtList(0).y = -(tmpPicSource.ScaleHeight / 2)
   PtList(1).x = tmpPicSource.ScaleWidth / 2
   PtList(1).y = -(tmpPicSource.ScaleHeight / 2)
   PtList(2).x = -(tmpPicSource.ScaleWidth / 2)
   PtList(2).y = (tmpPicSource.ScaleHeight / 2)
   
   'Variablen vorberechnen
   SinAng1 = Sin((Degree - 90) * PI / 180)
   CosAng1 = Cos((Degree - 90) * PI / 180)
   
   'Punkte transformieren
   For x = 0 To 2
      NewX = (PtList(x).x * SinAng1 + PtList(x).y * CosAng1)
      NewY = (PtList(x).y * SinAng1 - PtList(x).x * CosAng1)
      PtList(x).x = NewX + (tmpPicRotate.ScaleWidth / 2)
      PtList(x).y = NewY + (tmpPicRotate.ScaleHeight / 2)
   Next
   
   'Alte Darstellung des Backbuffers löschen
   tmpPicRotate.Cls
   
   'Neue Darstellung in den Backbuffer zeichnen
   Call PlgBlt(tmpPicRotate.hdc, PtList(0), tmpPicSource.hdc, 0, 0, tmpPicSource.ScaleWidth, tmpPicSource.ScaleHeight, 0, 0, 0)
End Sub

Public Sub Draw_RefreshNewColor(ByRef tmpPicSource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Ersetzen der alten Farbe durch eine neue            |'
'--------------------------------------------------------------------'
   'Zeichnen soll gesperrt werden
   LockDraw = True
   'Schleife durch alle Spalten
   For x = 1 To Spalten
      'Schleife durch alle LEDs
      For y = 1 To Leds
         'Farbe der Zelle ermitteln
         Col = Draw_GetCell(x, y, tmpPicSource)
         'Wenn Farbe nicht weiß und nicht die neue Farbe ist
         If (Col <> vbWhite) And (Col <> LEDColor) Then
            'Farbe durch die neue ersetzen
            Draw_FillCell x, y, tmpPicSource, LEDColor, False
         End If
      Next y
   Next x
   'Zeichnen soll erlaubt werden
   LockDraw = False
End Sub

Public Sub Draw_SetOffsetArrow(ByRef tmpPicArrow As PictureBox, ByRef tmpWinkel As Double)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen des Offset-Pfeils                            |'
'--------------------------------------------------------------------'
   'Verhältnis zwischen Gegenkathete und Hypothenuse mit Sin() ausrechnen
   RatioGkHyp = Sin((tmpWinkel - 90) / 180 * PI)
   'Verhältnis zwischen Ankathete und Hypothenuse mit Cos() ausrechnen
   RatioAkHyp = Cos((tmpWinkel - 90) / 180 * PI)
   
   'X-Position des Pfeils bestimmen
   PosOffsetArrow.x = (AbstandX + KreisRadius + RatioAkHyp * (RadiusOffsetArrow - 20 * ZoomX)) / ZoomX - tmpPicArrow.Width / 2
   'Y-Position des Pfeils bestimmen
   PosOffsetArrow.y = (AbstandY + KreisRadius + RatioGkHyp * (RadiusOffsetArrow - 20 * ZoomY)) / ZoomY - tmpPicArrow.Height / 2

   'Position des Offset-Pfeils setzen
   tmpPicArrow.Move PosOffsetArrow.x, PosOffsetArrow.y
   
   'Offset in Variable schreiben
   Offset = tmpWinkel / 360 * (Spalten - 1) \ 1
End Sub

Public Sub Draw_ClearModell()
'--------------------------------------------------------------------'
'| Prozedur zum Löschen des gesamten Modells inkl. Array            |'
'--------------------------------------------------------------------'
   'Schleife durch alle Spalten
   For x = 1 To Spalten
      'Schleife durch alle Zeilen (LEDs)
      For y = 1 To Leds
         'S/W-Array löschen
         Array_SW(x, y) = 0
         'RGB-Arrays löschen
         'Rot-Array
         Array_Red(x, y) = 0
         'Grün-Array
         Array_Green(x, y) = 0
         'Blau-Array
         Array_Blue(x, y) = 0
         'Zelle löschen
         Draw_FillCell x, y, frm_nightgraphix.pic_source, vbWhite, False
      Next y
   Next x
   'pic_source refreshen
   Draw_Refresh frm_nightgraphix.pic_source
   'Bild von pic_source nach pic_target zoomen
   Draw_Zoom frm_nightgraphix.pic_source, frm_nightgraphix.pic_target
   'PictureBox refreshen
   Draw_Refresh frm_nightgraphix.pic_target
End Sub

Public Sub Draw_InvertPicture(ByRef tmpPicSource As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Invertieren des ganzen Bildes        (neu Frank)    |'
'--------------------------------------------------------------------'
   'Zeichnen soll gesperrt werden
   LockDraw = True
   'Wenn RGB-Version
   If RGBVersion Then
      'Schleife durch alle Spalten
      For x = 1 To Spalten
         'Schleife durch alle LEDs
         For y = 1 To Leds
            'Farbe der Zelle ermitteln
            Col = Array_GetRGBCell(x, y, True)
            'Zelle füllen
            Draw_FillCell x, y, tmpPicSource, Col, True
         Next y
      Next x
   'Wenn SW-Version..
   Else
      'Schleife durch alle Spalten
      For x = 1 To Spalten
         'Schleife durch alle LEDs
         For y = 1 To Leds
            'Farbe der Zelle ermitteln
            Col = Draw_GetCell(x, y, tmpPicSource)
            'Wenn Farbe nicht weiß ist,
            If (Col <> vbWhite) Then
               'Zelle weiß zeichnen
               Draw_FillCell x, y, tmpPicSource, vbWhite, True
            Else
               Draw_FillCell x, y, tmpPicSource, LEDColor, True
            End If
         Next y
      Next x
   End If
   'Zeichnen soll erlaubt werden
   LockDraw = False

   'pic_source refreshen
   Draw_Refresh frm_nightgraphix.pic_source
   'Bild von pic_source nach pic_target zoomen
   Draw_Zoom frm_nightgraphix.pic_source, frm_nightgraphix.pic_target
   'PictureBox refreshen
   Draw_Refresh frm_nightgraphix.pic_target
End Sub

Public Sub Draw_Constructor(ByRef tmpPicSource As PictureBox, ByRef tmpPicTarget As PictureBox, ByRef tmpPicArrowShow As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum Initialisieren der gesamten Zeichenfläche           |'
'--------------------------------------------------------------------'
   'PicSource löschen
   tmpPicSource.Cls
   
   'Modell zeichnen
   Draw_Modell tmpPicSource

   'Hintergrund grau füllen
   Draw_Background tmpPicSource

   'pic_source refreshen
   Draw_Refresh tmpPicSource

   'Bild von pic_source nach pic_target zoomen
   Draw_Zoom tmpPicSource, tmpPicTarget
   
   'PictureBox refreshen
   Draw_Refresh tmpPicTarget

   'Offset-Pfeil setzen
   Draw_SetOffsetArrow tmpPicArrowShow, Offset * CDbl(360) / (Spalten - 1)

   'Offset-Pfeil zeichnen
   frm_nightgraphix.Draw_OffsetCursor Offset * CDbl(360) / (Spalten - 1)
End Sub

Public Sub Draw_Destructor()
'--------------------------------------------------------------------'
'| Prozedur zum Schließen der Zeichenfläche                         |'
'--------------------------------------------------------------------'
End Sub

Public Sub Draw_AlphaBlend(ByRef tmpPicTarget As PictureBox, ByRef tmpPicAlphaBlend As PictureBox, ByRef tmpPicImport As PictureBox)
'--------------------------------------------------------------------'
'| Prozedur zum erzeugen einer Transparenz von tmpPicImport         |'
'--------------------------------------------------------------------'
   'ScaleWidth und ScaleHeight
   Dim SW, SH
   With tmpPicImport
      'ScaleWidth und ScaleHeight von tmpPicAlphaBlend in Variablen speichern
      SW = tmpPicAlphaBlend.ScaleWidth
      SH = tmpPicAlphaBlend.ScaleHeight
      'Bildausschnitt aus tmpPicTarget auf tmpPicImport blitten
      Call BitBlt(.hdc, 0, 0, SW, SH, tmpPicTarget.hdc, .Left - tmpPicTarget.Left, .Top - tmpPicTarget.Top, vbSrcCopy)
      'Bild aus tmpPicAlphaBlend auf tmpPicImport blitten, halbe Transparenz
      Call AlphaBlend(.hdc, 0, 0, SW, SH, tmpPicAlphaBlend.hdc, 0, 0, SW, SH, CLng(&H10000 * 128))
      .Refresh
   End With
End Sub
