VERSION 5.00
Begin VB.Form Window 
   BackColor       =   &H00000080&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "m2lagfix"
   ClientHeight    =   120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   975
   Icon            =   "Window.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   120
   ScaleWidth      =   975
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   8
      Left            =   720
      Top             =   960
   End
End
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetPrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private UDP_LocalAddress_RX As String
Private UDP_LocalAddress_TX As String
Private UDP_RemoteAddress_RX As String
Private UDP_RemoteAddress_TX As String
Private UDP_Socket_RX As Long
Private UDP_Socket_TX As Long

Private STATS_Enabled As Boolean
Private STATS_RemoteAddress As String

Private LAG_Enabled As Boolean
Private LAG_Packet As String
Private LAG_Tick As Long

Private LINK_Enabled As Boolean

Private STALL_Enabled As Boolean
Private STALL_Tick As Long

Private Function ReadIni(File As String, Section As String, Key As String, Default As String) As String
  Dim Result As Integer
  Dim Temp As String * 1024
  Result = GetPrivateProfileStringA(Section, Key, Default, Temp, Len(Temp), App.Path & "\" & File)
  ReadIni = Left$(Temp, Result)
End Function

Private Sub Form_Load()
  ' Clear Log
  Open App.Path & "\m2laglog.txt" For Output As #1
  Print #1, "-- START @ " & Date & " " & Time
  Close #1
  
  ' Init Winsock
  Winsock.Load
    
  Dim Host As String
  Dim Port As Long
  
  ' Local-RX (m2lagfix)
  Host = ReadIni("m2lagfix.ini", "m2rx", "LocalHost", "127.0.0.1")
  Port = CLng(ReadIni("m2lagfix.ini", "m2rx", "LocalPort", "8000"))
  UDP_LocalAddress_RX = Winsock.WSABuildSocketAddress(Host, Port)
  
  ' Remote-RX (Emulator)
  Host = ReadIni("m2lagfix.ini", "m2rx", "RemoteHost", "127.0.0.1")
  Port = CLng(ReadIni("m2lagfix.ini", "m2rx", "RemotePort", "15612"))
  UDP_RemoteAddress_RX = Winsock.WSABuildSocketAddress(Host, Port)

  ' Local-TX (m2lagfix)
  Host = ReadIni("m2lagfix.ini", "m2tx", "LocalHost", "127.0.0.1")
  Port = CLng(ReadIni("m2lagfix.ini", "m2tx", "LocalPort", "15613"))
  UDP_LocalAddress_TX = Winsock.WSABuildSocketAddress(Host, Port)
  
  ' Remote-RX (Emulator)
  Host = ReadIni("m2lagfix.ini", "m2tx", "RemoteHost", "127.0.0.1")
  Port = CLng(ReadIni("m2lagfix.ini", "m2tx", "RemotePort", "8000"))
  UDP_RemoteAddress_TX = Winsock.WSABuildSocketAddress(Host, Port)

  ' M2Stats (if enabled)
  Host = ReadIni("m2lagfix.ini", "m2stats", "RemoteHost", "127.0.0.1")
  Port = CLng(ReadIni("m2lagfix.ini", "m2stats", "RemotePort", "-1"))
  STATS_RemoteAddress = Winsock.WSABuildSocketAddress(Host, Port)
  STATS_Enabled = Not (STATS_RemoteAddress = "")
  
  ' DELAY fix (if enabled)
  Port = CLng(ReadIni("m2lagfix.ini", "m2rx", "StallDetection", "0"))
  STALL_Enabled = CBool(Port = 1)
  
  If UDP_LocalAddress_RX = "" Or UDP_RemoteAddress_RX = "" Then
    MsgBox "Something went wrong! #ADDR_RX", vbCritical Or vbOKOnly, "m2lagfix"
    Form_Unload 0
  End If
  
  If UDP_LocalAddress_TX = "" Or UDP_RemoteAddress_TX = "" Then
    MsgBox "Something went wrong! #ADDR_TX", vbCritical Or vbOKOnly, "m2lagfix"
    Form_Unload 0
  End If
  
  UDP_Socket_RX = Winsock.ListenUDP(UDP_LocalAddress_RX)
  If UDP_Socket_RX = -1 Then
    MsgBox "Something went wrong! #SOCK_RX", vbCritical Or vbOKOnly, "m2lagfix"
    Form_Unload 0
  End If
  
  UDP_Socket_TX = Winsock.ListenUDP(UDP_LocalAddress_TX)
  If UDP_Socket_TX = -1 Then
    MsgBox "Something went wrong! #SOCK_TX", vbCritical Or vbOKOnly, "m2lagfix"
    Form_Unload 0
  End If
  
  LAG_Enabled = False
  LAG_Packet = ""
  LAG_Tick = 0
  
  LINK_Enabled = False
  
  STALL_Tick = 0
  
  If STATS_Enabled Then
    Window.BackColor = RGB(0, 128, 0)
  Else
    Window.BackColor = RGB(128, 0, 0)
  End If
  
  Timer.Enabled = True

  Debug.Print "UDP_Socket_RX", UDP_Socket_RX, "UDP_Socket_TX", UDP_Socket_TX
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Winsock.Unload
  End
End Sub

Public Sub OnReadTCP(lSocket As Long, sBuffer As String)
  Debug.Print "OnReadTCP", "Else", lSocket, Len(sBuffer)
End Sub

Public Sub OnReadUDP(lSocket As Long, sBuffer As String, sAddress As String)
  Dim sDummy As String
  If lSocket = UDP_Socket_RX Then
    ' incoming, send to emulator
    While Len(sBuffer) >= 3589
      sDummy = Left$(sBuffer, 3589)
      sBuffer = Mid$(sBuffer, 3590)
      Winsock.SendUDP UDP_Socket_RX, sDummy, UDP_RemoteAddress_RX
      If STATS_Enabled Then
        Winsock.SendUDP UDP_Socket_RX, sDummy, STATS_RemoteAddress
      End If
    Wend
    
    LINK_Enabled = (Asc(Mid$(sDummy, 5, 1)) = 2)
    
    LAG_Tick = GetTickCount
    LAG_Packet = sDummy
    LAG_Enabled = LINK_Enabled
  ElseIf lSocket = UDP_Socket_TX Then
    ' outgoing, send to next unit
    While Len(sBuffer) >= 3589
      sDummy = Left$(sBuffer, 3589)
      sBuffer = Mid$(sBuffer, 3590)
      Winsock.SendUDP UDP_Socket_TX, sDummy, UDP_RemoteAddress_TX
    Wend

    STALL_Tick = GetTickCount
    LAG_Enabled = False
  End If
End Sub

Public Sub OnIncoming(lSocket As Long, sNewSocket As Long)
  Debug.Print "OnIncoming", lSocket, sNewSocket
End Sub

Public Sub OnConnected(lSocket As Long)
  Debug.Print "OnConnected", lSocket
End Sub

Public Sub OnConnectError(lSocket As Long, lError As Long)
  Debug.Print "OnConnectError", lSocket, lError
End Sub

Public Sub OnClose(lSocket As Long)
  Debug.Print "OnClose", lSocket
End Sub

Private Sub Timer_Timer()
  ' LAG if enabled
  If LAG_Enabled Then
    ' if waited long enough
    If GetTickCount - LAG_Tick >= 64 Then
      Winsock.SendUDP UDP_Socket_RX, LAG_Packet, UDP_RemoteAddress_RX
      LAG_Tick = GetTickCount
    
      ' Log
      Open App.Path & "\m2laglog.txt" For Append As #1
      Print #1, "LAG-- @ " & Date & " " & Time
      Close #1
    End If
  ElseIf (STALL_Enabled And LINK_Enabled) Then
    ' if waited long enough
    If GetTickCount - STALL_Tick >= 128 Then
      Winsock.SendUDP UDP_Socket_RX, LAG_Packet, UDP_RemoteAddress_RX
      STALL_Tick = GetTickCount
      
      ' Log
      Open App.Path & "\m2laglog.txt" For Append As #1
      Print #1, "STALL @ " & Date & " " & Time
      Close #1
    End If
  End If
End Sub


