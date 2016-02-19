VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_nvtask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variablen müssen deklariert werden..
Option Explicit

'ENUMs..
'Unterschiedliche Aufgaben
Public Enum tTaskTyp
   'Ein Punkt soll gezeichnet werden
   [Draw Point]
   'Eine Linie soll gezeichnet werden
   [Draw Line]
   'Ein Kreis soll gezeichnet werden
   [Draw Circle]
   'Ein Buchstabe soll gezeichnet werden
   [Draw Letter]
End Enum

'Typ des Tasks
Private mTyp As tTaskTyp
'Koordinaten des ersten Punkts des Tasks
Private mX1 As Double
Private mY1 As Double
'Farbe des Tasks
Private mColor As Long
'Optional: Koordinaten des ersten Punkts des Tasks
Private mX2 As Double
Private mY2 As Double
'Optional: Was für ein Buchstabe ?
Private mLetter As String

Public Property Get Typ() As tTaskTyp
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen vom Typ des Tasks                          |'
'--------------------------------------------------------------------'
   'Typ des Tasks zurückgeben
   Typ = mTyp
End Property

Public Property Let Typ(ByVal tmpTyp As tTaskTyp)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen vom Typ des Tasks                            |'
'--------------------------------------------------------------------'
   'Typ des Tasks setzen
   mTyp = tmpTyp
End Property

Public Property Get X1() As Double
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen des ersten Punktes des Tasks               |'
'--------------------------------------------------------------------'
   'Point1 des Tasks zurückgeben
   X1 = mX1
End Property

Public Property Let X1(ByVal tmpX1 As Double)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen des zweiten Punktes des Tasks                |'
'--------------------------------------------------------------------'
   'Point1 des Tasks setzen
   mX1 = tmpX1
End Property

Public Property Get Y1() As Double
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen des ersten Punktes des Tasks               |'
'--------------------------------------------------------------------'
   'Point1 des Tasks zurückgeben
   Y1 = mY1
End Property

Public Property Let Y1(ByVal tmpY1 As Double)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen des zweiten Punktes des Tasks                |'
'--------------------------------------------------------------------'
   'Point1 des Tasks setzen
   mY1 = tmpY1
End Property

Public Property Get Color() As Long
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen der Farbe des Tasks                        |'
'--------------------------------------------------------------------'
   'Farbe des Tasks zurückgeben
   Color = mColor
End Property

Public Property Let Color(ByVal tmpColor As Long)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen der Farbe des Tasks                          |'
'--------------------------------------------------------------------'
   'Farbe des Tasks setzen
   mColor = tmpColor
End Property

Public Property Get X2() As Double
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen des zweiten Punktes des Tasks              |'
'--------------------------------------------------------------------'
   'Point2 des Tasks zurückgeben
   X2 = mX2
End Property

Public Property Let X2(ByVal tmpX2 As Double)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen vom zweiten Punktes des Tasks                |'
'--------------------------------------------------------------------'
   'Point2 des Tasks setzen
   mX2 = tmpX2
End Property

Public Property Get Y2() As Double
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen des zweiten Punktes des Tasks              |'
'--------------------------------------------------------------------'
   'Point2 des Tasks zurückgeben
   Y2 = mY2
End Property

Public Property Let Y2(ByVal tmpY2 As Double)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen vom zweiten Punktes des Tasks                |'
'--------------------------------------------------------------------'
   'Point2 des Tasks setzen
   mY2 = tmpY2
End Property

Public Property Get Letter() As String
'--------------------------------------------------------------------'
'| Prozedur zum Auslesen des Buchstabens des Tasks                  |'
'--------------------------------------------------------------------'
   'Buchstaben des Tasks zurückgeben
   Letter = mLetter
End Property

Public Property Let Letter(ByVal tmpLetter As String)
'--------------------------------------------------------------------'
'| Prozedur zum Setzen vom Buchstaben des Tasks                     |'
'--------------------------------------------------------------------'
   'Buchstaben des Tasks setzen
   mLetter = tmpLetter
End Property
