Attribute VB_Name = "mdl_comport"
'_________________________________________________________________________________'
'|                               MODUL mdl_comport                               |'
'| Dieses Modul beinhaltet Routinen f�r den COM-Port                             |'
'|                                                                               |'
'---------------------------------------------------------------------------------'
'Variablen m�ssen deklariert werden
Option Explicit

'VARIABLEN..
'Ist Port ge�ffnet ?
Public COMPortOpened As Boolean
'Der zuletzt ge�ffnete Port
Public COMPortOpened_Last As Integer
'Der aktuell ge�ffnete Port
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
   
   'Wenn MSComm-Control noch einen COM-Port ge�ffnet hat..
   If tmpMSComm.PortOpen = True Then
      'COM-Port schlie�en
      tmpMSComm.PortOpen = False
   End If
   
   'Anderen Events Zeit lassen
   DoEvents
   
   'COM-Port setzen
   tmpMSComm.CommPort = tmpComPortNumber
   'Versuche COM-Port zu �ffnen
   tmpMSComm.PortOpen = True
   
'--Wenn Programm bis hier kommt, lie� sich der COM-Port �ffnen--'
   'COM-Port l�sst sich �ffnen
   ComPort_Available = True
   'Nicht in die Fehlerbehandlung springen
   GoTo CloseCOMPort
   
'Fehlerbehandlung
ErrHandler:
   'Wenn die FehlerNR = 8005 ist..
   '(Anschluss bereits ge�ffnet)
   If Err.Number = 8005 Then
      'COM-Port ist nicht vorhanden oder schon offen..
      ComPort_Available = False
   'Anderer Fehler..
   Else
      'COM-Port l�sst sich nicht �ffnen..
      ComPort_Available = False
   End If
   
'COM-Port wieder schlie�en
CloseCOMPort:
   'Wenn COM-Port noch ge�ffnet ist..
   If tmpMSComm.PortOpen = True Then
      'COM-Port schlie�en
      tmpMSComm.PortOpen = False
   End If
End Function

Public Sub ComPort_Open(ByRef tmpMSComm As MSComm, ByVal tmpPortNumber As Integer)
'--------------------------------------------------------------------'
'| Prozedur zum �ffnen eines COM-Ports                              |'
'--------------------------------------------------------------------'
   'Bei einem Fehler die Fehlerbehandlungsroutine ausf�hren
   'On Error GoTo ErrHandler
   
   'COM-Port schlie�en
   ComPort_Close tmpMSComm
   'Einstellungen �ndern
   tmpMSComm.Settings = "9600,N,8,1"
   'Portnummer �ndern
   tmpMSComm.CommPort = tmpPortNumber
   'Port �ffnen
   tmpMSComm.PortOpen = True
   'Port ist ge�ffnet
   COMPortOpened = True
   
   'Fehlerbehandlung nicht ausf�hren
   Exit Sub
   
'Fehlerbehandlung
ErrHandler:
   'Nur f�r Simulation, da kein COM-Port ge�ffnet werden kann
End Sub

Public Sub ComPort_Close(ByRef tmpMSComm As MSComm)
'--------------------------------------------------------------------'
'| Prozedur zum Schlie�en eines COM-Ports                           |'
'--------------------------------------------------------------------'
   'Wenn COM-Port noch ge�ffnet ist..
   If tmpMSComm.PortOpen Then
      'Port schlie�en
      tmpMSComm.PortOpen = False
   End If
   'Port ist nicht mehr ge�ffnet
   COMPortOpened = False
End Sub


