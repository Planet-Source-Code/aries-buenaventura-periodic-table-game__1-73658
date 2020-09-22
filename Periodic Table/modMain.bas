Attribute VB_Name = "modMain"
Option Explicit

Public PTElem  As New Collection
Public PTTable As New Collection

Public Counter As Integer

Public Function GetCurPath() As String
    ' get current path
    If Mid$(App.Path, Len(App.Path), 1) = "\" Then
        GetCurPath = Left$(App.Path, Len(App.Path) - 1)
    Else
        GetCurPath = App.Path
    End If
End Function

Public Function RemoveUnneededSpace(Expression As String) As String
' this function will remove unneeded space
' example:
'    1.
'       input : 1    3   4    5
'       output: 1 3 4 5
'    2.
'       input  : 4   3   55 12
'       output : 4 3 55 12

    Dim s       As String
    Dim i       As Integer
    Dim IsSpace As Boolean
    
    Expression = Trim$(Expression)
    
    For i = 1 To Len(Expression)
        If Mid$(Expression, i, 1) = " " Then
            If Not IsSpace Then
                s = s & " "
                IsSpace = True
            End If
        Else
            s = s & Mid$(Expression, i, 1)
            
            IsSpace = False
        End If
    Next i
    
    RemoveUnneededSpace = s
End Function

Public Function Generate_Random_Number(ByVal lowerbound As Integer, upperbound As Integer) As Integer
    Call Randomize
    Generate_Random_Number = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Public Sub PlaySound(ByVal SoundFile As String)
    If Dir$(GetCurPath & "\sound\" & SoundFile) <> "" Then
        sndPlaySound GetCurPath & "\sound\" & SoundFile, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT
    End If
End Sub
