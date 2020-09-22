VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPTElem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Atomic_Number     As Integer ' Atomic Number
Public Symbol            As String  ' Atomic Symbol
Public Name              As String  ' Name
Public Atomic_Weight     As String  ' Atomic Weight
Public Oxidation_States  As String  ' Oxidation states (valence)^b
Public Melting_Point     As String  ' Melting point (°C)
Public Boiling_Point     As String  ' Boiling point (°C)
Public Density           As String  ' Density^4 (g/cm^3)
Public Electronegativity As String  ' Electronegativy
Public Atomic_Radius     As String  ' Atomic Radius (1x10^-10m)

Public Discoverer        As String
Public Country           As String
Public DDate             As String
Public Uses              As String


