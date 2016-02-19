Attribute VB_Name = "mdl_comport"
'_________________________________________________________________________________'
'|                               MODUL mdl_comport                               |'
'| Dieses Modul beinhaltet Routinen für den COM-Port                             |'
'|                                                                               |'
'---------------------------------------------------------------------------------'
'Variablen müssen deklariert werden
Option Explicit

'VARIABLEN..
'Ist Port geöffnet ?
Public COMPortOpened As Boolean
'Der zuletzt geöffnete Port
Public COMPortOpened_Last As Integer
'Der aktuell geöffnete Port
Public COMPortOpened_Now

Public Function ComPort_Available(ByRef tmpMSComm As MSComm, ByVal tmpComPortNumber As Integer) As Boolean
'--------------------------------------------------------------------'
'| Prozedur zum Testen, ob ein bestimmter Port vorhanden ist        |'
'--------------------------------------------------------------------'
   'Bei einem Fehler die Fehlerbehandlung aufrufen
   On Error GoTo ErrHandler
   
   'Wenn tmpComPortNumber nicht zwischen 1 und 16 liegt..
   If (tmpComPortNumber < 1) Or (tmpComPortNumber > 16) Then
      'In Fehlerbehandlung springen
      GoTo ErrHandler
   End If
   
   'Wenn MSComm-Control noch einen COM-Port geöffnet hat..
   If tmpMSComm.PortOpen = True Then
      'COM-Port schließen
      tmpMSComm.PortOpen = False
   End If
   
   'Anderen Events Zeit lassen
   DoEvents
   
   'COM-Port setzen
   tmpMSComm.CommPort = tmpComPortNumber
   'Versuche COM-Port zu öffnen
   tmpMSComm.PortOpen = True
   
'--Wenn Programm bis hier kommt, ließ sich der COM-Port öffnen--'
   'COM-Port lässt sich öffnen
   ComPort_Available = True
   'Nicht in die Fehlerbehandlung springen
   GoTo CloseCOMPort
   
'Fehlerbehandlung
ErrHandler:
   'Wenn die FehlerNR = 8005 ist..
   '(Anschluss bereits geöffnet)
   If Err.Number = 8005 Then
      'COM-Port ist nicht vorhanden oder schon offen..
      ComPort_Available = False
   'Anderer Fehler..
   Else
      'COM-Port lässt sich nicht öffnen..
      ComPort_Available = False
   End If
   
'COM-Port wieder schließen
CloseCOMPort:
   'Wenn COM-Port noch geöffnet ist..
   If tmpMSComm.PortOpen = True Then
      'COM-Port schließen
      tmpMSComm.PortOpen = False
   End If
End Function

Public Sub ComPort_Open(ByRef tmpMSComm As MSComm, ByVal tmpPortNumber As Integer)
'--------------------------------------------------------------------'
'| Prozedur zum Öffnen eines COM-Ports                              |'
'--------------------------------------------------------------------'
   'Bei einem Fehler die Fehlerbehandlungsroutine ausführen
   'On Error GoTo ErrHandler
   
   'COM-Port schließen
   ComPort_Close tmpMSComm
   'Einstellungen ändern
   tmpMSComm.Settings = "9600,N,8,1"
   'Portnummer ändern
   tmpMSComm.CommPort = tmpPortNumber
   'Port öffnen
   tmpMSComm.PortOpen = True
   'Port ist geöffnet
   COMPortOpened = True
   
   'Fehlerbehandlung nicht ausführen
   Exit Sub
   
'Fehlerbehandlung
ErrHandler:
   'Nur für Simulation, da kein COM-Port geöffnet werden kann
End Sub

Public Sub ComPort_Close(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Schließen eines COM-Ports                           |'
'--------------------------------------------------------------------'
   'Wenn COM-Port noch geöffnet ist..
   If tmpMSComm.PortOpen Then
      'Port schließen
      tmpMSComm.PortOpen = False
   End If
   'Port ist nicht mehr geöffnet
   COMPortOpened = False
End Sub


