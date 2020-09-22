Attribute VB_Name = "modAPI"
Option Explicit

Public Const SND_FILENAME = &H20000     '  name is a file name
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_SYNC = &H0         '  play synchronously (default)

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


