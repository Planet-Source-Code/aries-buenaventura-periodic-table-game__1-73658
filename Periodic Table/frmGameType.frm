VERSION 5.00
Begin VB.Form frmGameType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTutorial 
      Caption         =   "Tutorial"
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   3180
      Width           =   855
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   3180
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   3180
      Width           =   915
   End
   Begin VB.Frame fraTray 
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.PictureBox picLevel 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   3375
         TabIndex        =   4
         Top             =   1500
         Width           =   3375
         Begin VB.OptionButton optLevel 
            Caption         =   "Difficult"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   780
            Width           =   1215
         End
         Begin VB.OptionButton optLevel 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   420
            Width           =   1215
         End
         Begin VB.OptionButton optLevel 
            Caption         =   "Easy"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   60
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.PictureBox picGameType 
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   60
         ScaleHeight     =   1155
         ScaleWidth      =   3495
         TabIndex        =   1
         Top             =   180
         Width           =   3495
         Begin VB.OptionButton optGameType 
            Caption         =   "Click on the symbol"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   180
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton optGameType 
            Caption         =   "Name the Element"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   180
            TabIndex        =   2
            Top             =   540
            Width           =   2415
         End
      End
      Begin VB.Line linLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   3480
         Y1              =   1395
         Y2              =   1395
      End
      Begin VB.Line linLine 
         Index           =   0
         X1              =   120
         X2              =   3480
         Y1              =   1380
         Y2              =   1380
      End
   End
End
Attribute VB_Name = "frmGameType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPlay_Click()
    With frmMain
        If optGameType(1).Value Then
            .GameType = 0
        ElseIf optGameType(2).Value Then
            .GameType = 1
        End If
        
        If optLevel(0).Value Then
            .Level = 1
        ElseIf optLevel(1).Value Then
            .Level = 2
        Else
            .Level = 3
        End If
        
        .Tutorial = False
    End With
    
    Unload Me
    Call frmMain.Start
End Sub

Private Sub cmdTutorial_Click()
    With frmMain
        If optGameType(1).Value Then
            .GameType = 0
        ElseIf optGameType(2).Value Then
            .GameType = 1
        End If
        
        If optLevel(0).Value Then
            .Level = 1
        ElseIf optLevel(1).Value Then
            .Level = 2
        Else
            .Level = 3
        End If
        
        .Tutorial = True
    End With
    
    Unload Me
    Call frmMain.Start
End Sub

Private Sub optGameType_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
    Case Is = 1
        For i = optLevel.LBound To optLevel.UBound
            optLevel(i).Enabled = True
        Next i
        
        cmdTutorial.Enabled = True
    Case Is = 2
        For i = optLevel.LBound To optLevel.UBound
            optLevel(i).Enabled = False
        Next i
        
        cmdTutorial.Enabled = False
    End Select
End Sub
