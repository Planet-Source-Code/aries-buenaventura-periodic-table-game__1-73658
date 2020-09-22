VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Periodic Table (Demo only...) by Aris Buenaventura"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   586
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5340
      Top             =   4140
   End
   Begin VB.PictureBox picElemWinInfo 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   792
      TabIndex        =   11
      Top             =   7095
      Width           =   11880
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   0
         ScaleHeight     =   111
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   429
         TabIndex        =   18
         Top             =   30
         Width           =   6435
         Begin VB.TextBox txtDataCountry 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            TabIndex        =   3
            Top             =   360
            Width           =   3375
         End
         Begin VB.TextBox txtDataUses 
            Appearance      =   0  'Flat
            Height          =   585
            Left            =   1140
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   1035
            Width           =   3375
         End
         Begin VB.TextBox txtDataDate 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            TabIndex        =   6
            Top             =   690
            Width           =   3375
         End
         Begin VB.TextBox txtDataDiscoverer 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1140
            TabIndex        =   1
            Top             =   30
            Width           =   3375
         End
         Begin VB.Label lblUses 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Uses :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   525
            TabIndex        =   7
            Top             =   1260
            Width           =   555
         End
         Begin VB.Label lblDiscoverer 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Discoverer : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   30
            TabIndex        =   0
            Top             =   0
            Width           =   1110
         End
         Begin VB.Label lblCountry 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Country : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   2
            Top             =   420
            Width           =   840
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Da&te : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   540
            TabIndex        =   4
            Top             =   780
            Width           =   600
         End
      End
      Begin VB.PictureBox picElemImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1425
         Left            =   0
         ScaleHeight     =   93
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   0
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   0
      End
      Begin VB.PictureBox picElemInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1680
         Left            =   6480
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   360
         TabIndex        =   12
         Top             =   0
         Width           =   5400
      End
      Begin VB.Label lblElem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   1470
         Width           =   45
      End
   End
   Begin VB.PictureBox picPT 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6555
      Left            =   600
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   749
      TabIndex        =   5
      Top             =   540
      Width           =   11235
      Begin VB.PictureBox picPlayElem 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   1440
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   385
         TabIndex        =   14
         Top             =   120
         Width           =   5775
         Begin VB.CommandButton cmdOk 
            Caption         =   "Ok"
            Default         =   -1  'True
            Height          =   375
            Left            =   780
            TabIndex        =   17
            Top             =   60
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox txtElement 
            Height          =   315
            Left            =   2340
            TabIndex        =   16
            Top             =   60
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmdAnother 
            Caption         =   "Try anohter"
            Height          =   375
            Left            =   2400
            TabIndex        =   15
            Top             =   1140
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblMsgErr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   390
            TabIndex        =   22
            Top             =   1320
            Visible         =   0   'False
            Width           =   195
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblMsgName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Name the highlighted Element : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   60
            TabIndex        =   21
            Top             =   720
            Visible         =   0   'False
            Width           =   2805
         End
         Begin VB.Label lblMsg 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   30
            TabIndex        =   20
            Top             =   1320
            Visible         =   0   'False
            Width           =   165
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.PictureBox picCol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   660
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   749
      TabIndex        =   9
      Top             =   0
      Width           =   11235
   End
   Begin VB.PictureBox picRow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      ScaleHeight     =   473
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   10
      Top             =   0
      Width           =   675
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuGameNewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelElem As New Collection
Dim CurElem As Integer

Public Level    As Integer
Public GameType As Integer
Public Tutorial As Boolean

Private Sub LoadPTInfo()
    Dim i          As Integer
    Dim InFile     As Long
    Dim sTemp      As String
    Dim oTemp      As Object  ' temporary object
    Dim arrTemp()  As String
    Dim arrTemp1() As String
    Dim arrTemp2() As String
    Dim DataKey    As Integer ' 1-[Map],2-[Data]
    
    On Error GoTo ErrHandler
    
    If Dir$(GetCurPath & "\PTsymb.dat", vbHidden Or vbNormal) <> "" Then ' determine if a particular file exist
        InFile = FreeFile ' provide you with an unused file number, so you can
                          ' open a file with a unique identifier.
                          ' note: when you open a file on the disk drive
                          '       (to read it, add to it, change to it, whatever),
                          '       you use a filenumber.
        Open GetCurPath & "\PTsymb.dat" For Input As InFile
            If LOF(InFile) > 0 Then ' LOF (Length Of File) this tells you how large a disk file is, _
                                    ' how many bytes it takes up on the disk
                Do While Not EOF(InFile) ' continue looping until we reached the end of the file
                    Line Input #InFile, sTemp ' It reads in a single "line" of text-all characters up
                                              ' to the "Carriage Return."
                    If Trim$(sTemp) <> "" Then
                        If Left$(Trim$(sTemp), 1) <> "'" Then ' all comments are marked ' (single quotation)
                            Select Case UCase$(Trim$(sTemp))
                            Case Is = "[MAP]"
                                DataKey = 1
                            Case Is = "[DATA]"
                                DataKey = 2
                            End Select
    On Error Resume Next
                            Select Case DataKey
                            Case Is = 1 ' [Map]
                                If UCase$(Trim$(sTemp)) <> "[MAP]" Then
                                    sTemp = RemoveUnneededSpace(Trim$(sTemp))
                                    arrTemp1 = Split(sTemp, " ")
                                    
                                    For i = LBound(arrTemp1()) To UBound(arrTemp1())
                                        Set oTemp = New cPTTable
            
                                        If Left$(arrTemp1(i), 1) <> "#" Then
                                            arrTemp2 = Split(arrTemp1(i), ",")
                                            oTemp.Atomic_Number = Val(arrTemp2(0))
                                            oTemp.Element = arrTemp2(1)
                                        Else
                                            oTemp.Element = ""
                                        End If
                                        
                                        PTTable.Add oTemp
                                    Next i
                                End If
                            Case Is = 2  ' [Data]
                                If UCase$(Trim$(sTemp)) <> "[DATA]" Then
                                    arrTemp = Split(sTemp, ",")
                                    
                                    Set oTemp = New cPTElem
                                    oTemp.Atomic_Number = Val(arrTemp(0))
                                    oTemp.Symbol = arrTemp(1)
                                    oTemp.Name = arrTemp(2)
                                    oTemp.Atomic_Weight = arrTemp(3)
                                    oTemp.Oxidation_States = Replace(arrTemp(4), "@", ",")
                                    oTemp.Melting_Point = arrTemp(5)
                                    oTemp.Boiling_Point = arrTemp(6)
                                    oTemp.Density = arrTemp(7)
                                    oTemp.Electronegativity = arrTemp(8)
                                    oTemp.Atomic_Radius = arrTemp(9)
                                    
                                    oTemp.Discoverer = Replace(arrTemp(10), "@", ",")
                                    oTemp.Country = Replace(arrTemp(11), "@", ",")
                                    oTemp.DDate = arrTemp(12)
                                    oTemp.Uses = Replace(arrTemp(13), "@", ",")
                                    PTElem.Add oTemp
                                End If
                            End Select
                        Else
                            ' this is a comment so let skip it.
                        End If
                    End If
                Loop
            Else
                ' there's nothing to read in the PTsymb.dat because the
                ' length of the file is equal to zero
            End If
        Close #InFile
    Else
        ' if file does not exist then exit the program
        MsgBox GetCurPath & "\" & "PTsymb.dat" & vbCrLf & "File not found.", _
               vbExclamation Or vbOKOnly, "Interactive Graphic"
        Unload Me
    End If
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbOKOnly Or vbCritical, "Interactive Graphic"
    If InFile <> 0 Then Close #InFile
    Unload Me
End Sub

Private Sub cmdAnother_Click()
    PlaySound "click.wav"
    cmdAnother.Visible = False
        
    Call Start
End Sub

Private Sub cmdOk_Click()
    If Trim$(txtElement.Text) = "" Then Exit Sub
    
    Dim i     As Integer
    Dim sTemp As String
    
    If UCase$(Trim$(PTElem(SelElem(1)).Name)) = UCase$(Trim$(txtElement.Text)) Then
        i = SelElem(1)
        
        txtElement.Text = ""
        txtElement.SetFocus
        
        Call Start
        lblMsgErr.Caption = "Right! " & PTElem(i).Symbol & " is the symbol for " & PTElem(i).Name & "."
        PlaySound "applause.wav"
    Else
        lblMsgErr.Caption = "The correct answer is " & PTElem(SelElem(1)).Name & " not " & Trim$(txtElement.Text) & "."
        
        txtElement.Text = ""
        txtElement.SetFocus
        PlaySound "beep.wav"
        Call picPlayElem_Resize
        Call Start
        
        lblMsgErr.Visible = True
    End If
    
    picPlayElem.Refresh
End Sub

Private Sub Form_Load()
    Counter = GetSetting("123456", "123456", "123456", 20)
    MsgBox "trials left : " & Counter
    If Counter <= 0 Then
        MsgBox "trial session ended"
        Exit Sub
    End If
    
    Call LoadPTInfo
    Call ShowElemInfo(1) ' hydrogen
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState <> vbMinimized Then
        picRow.Move 0, 0, picRow.ScaleWidth, _
                    Me.ScaleHeight - picElemWinInfo.ScaleHeight
        picCol.Move picRow.ScaleWidth - 1, 0, Me.ScaleWidth - picRow.ScaleWidth - 1, _
                    picCol.ScaleHeight
        picPT.Move picRow.ScaleWidth - 1, _
                   picCol.ScaleHeight + 20, _
                   Me.ScaleWidth - picRow.ScaleWidth - 1, _
                   Me.ScaleHeight - picCol.ScaleHeight - picElemWinInfo.ScaleHeight - 20
        
        Dim sw As Integer
        Dim sh As Integer
        
        sw = picPT.ScaleWidth / 18
        sh = picPT.ScaleHeight / 9 - 1
        
        picPlayElem.Move sw * 2.25, 0, sw * 9.5, sh * 2.75
        Call picRow_Resize
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrBlink.Enabled = False
    SaveSetting "123456", "123456", "123456", Counter - 1
End Sub

Private Sub mnuGameExit_Click()
    End
End Sub

Private Sub mnuGameNew_Click()
    frmGameType.Show vbModal, Me
End Sub

Private Sub picCol_Resize()
    Dim i     As Integer
    Dim sw    As Single
    Dim sh    As Single
    Dim sTemp As String
    
    sw = picCol.ScaleWidth / 18
    sh = picCol.ScaleHeight - 1
    
    picCol.FontBold = True
    For i = 0 To 18
        If (i >= 7) And (i <= 9) Then
            ' do nothing
        Else
            picCol.Line (i * sw, 0)-((i + 1) * sw, sh), GetColBkColor(i), BF
            picCol.Line (i * sw, 0)-((i + 1) * sw, sh), &H0, B
        End If
        
        Select Case i
        Case 0, 1
            If i = 0 Then
                sTemp = "1A"
            ElseIf i = 1 Then
                sTemp = "2A"
            End If
            
            TextOut picCol.hdc, i * sw + ((i + 1) * sw - i * sw) / 2 - picCol.TextWidth(sTemp) / 2, _
                                (picCol.ScaleHeight - picCol.TextHeight(sTemp)) / 2, sTemp, Len(sTemp)
        Case 2, 3, 4, 5, 6
            If i = 2 Then
                sTemp = "3B"
            ElseIf i = 3 Then
                sTemp = "4B"
            ElseIf i = 4 Then
                sTemp = "5B"
            ElseIf i = 5 Then
                sTemp = "6B"
            ElseIf i = 6 Then
                sTemp = "7B"
            End If
            
            TextOut picCol.hdc, i * sw + ((i + 1) * sw - i * sw) / 2 - picCol.TextWidth(sTemp) / 2, _
                                (picCol.ScaleHeight - picCol.TextHeight(sTemp)) / 2 - picCol.TextHeight(sTemp) / 2, sTemp, Len(sTemp)
            If i = 2 Then
                sTemp = "3A"
            ElseIf i = 3 Then
                sTemp = "4A"
            ElseIf i = 4 Then
                sTemp = "5A"
            ElseIf i = 5 Then
                sTemp = "6A"
            ElseIf i = 6 Then
                sTemp = "7A"
            End If
            
            TextOut picCol.hdc, i * sw + ((i + 1) * sw - i * sw) / 2 - picCol.TextWidth(sTemp) / 2, _
                                (picCol.ScaleHeight - picCol.TextHeight(sTemp)) / 2 + picCol.TextHeight(sTemp) / 2, sTemp, Len(sTemp)
        Case 10, 11
            If i = 10 Then
                sTemp = "1B"
            ElseIf i = 11 Then
                sTemp = "2B"
            End If
            
            TextOut picCol.hdc, i * sw + ((i + 1) * sw - i * sw) / 2 - picCol.TextWidth(sTemp) / 2, _
                                (picCol.ScaleHeight - picCol.TextHeight(sTemp)) / 2, sTemp, Len(sTemp)
        Case 12, 13, 14, 15, 16
            If i = 12 Then
                sTemp = "3A"
            ElseIf i = 13 Then
                sTemp = "4A"
            ElseIf i = 14 Then
                sTemp = "5A"
            ElseIf i = 15 Then
                sTemp = "6A"
            ElseIf i = 16 Then
                sTemp = "7A"
            End If
            
            TextOut picCol.hdc, i * sw + ((i + 1) * sw - i * sw) / 2 - picCol.TextWidth(sTemp) / 2, _
                                (picCol.ScaleHeight - picCol.TextHeight(sTemp)) / 2 - picCol.TextHeight(sTemp) / 2, sTemp, Len(sTemp)
            If i = 12 Then
                sTemp = "3B"
            ElseIf i = 13 Then
                sTemp = "4B"
            ElseIf i = 14 Then
                sTemp = "5B"
            ElseIf i = 15 Then
                sTemp = "6B"
            ElseIf i = 16 Then
                sTemp = "7B"
            End If
            
            TextOut picCol.hdc, i * sw + ((i + 1) * sw - i * sw) / 2 - picCol.TextWidth(sTemp) / 2, _
                                (picCol.ScaleHeight - picCol.TextHeight(sTemp)) / 2 + picCol.TextHeight(sTemp) / 2, sTemp, Len(sTemp)
        Case 17
            sTemp = "0"
            TextOut picCol.hdc, i * sw + ((i + 1) * sw - i * sw) / 2 - picCol.TextWidth(sTemp) / 2, _
                                (picCol.ScaleHeight - picCol.TextHeight(sTemp)) / 2, sTemp, Len(sTemp)
        End Select
    Next i
    
    picCol.Line (7 * sw, 0)-(10 * sw, sh), GetColBkColor(7), BF
    picCol.Line (7 * sw, 0)-(10 * sw, sh), &H0, B
    
    TextOut picCol.hdc, 7 * sw + (10 * sw - 7 * sw) / 2 - picCol.TextWidth("8") / 2, _
                        (picCol.ScaleHeight - picCol.TextHeight(sTemp)) / 2, "8", Len("8")
    TextOut picCol.hdc, 7 * sw + (7 * sw - 7 * sw) / 2 + 2, _
                        (picCol.ScaleHeight - picCol.TextHeight("USA")) / 2 - picCol.TextHeight("USA") / 2, "USA", Len("USA")
    TextOut picCol.hdc, 7 * sw + (7 * sw - 7 * sw) / 2 + 2, _
                        (picCol.ScaleHeight - picCol.TextHeight("Europe")) / 2 + picCol.TextHeight("USA") / 2, "Europe", Len("Europe")
    TextOut picCol.hdc, 10 * sw + (10 * sw - 10 * sw) / 2 - picCol.TextWidth("IUPAC+") - 2, _
                        (picCol.ScaleHeight - picCol.TextHeight("Europe")) / 2 + picCol.TextHeight("IUPAC+") / 2, "IUPAC+", Len("IUPAC+")
End Sub

Private Sub picElemWinInfo_Resize()
    On Error Resume Next
    
    picElemInfo.Move picElemWinInfo.ScaleWidth - picElemInfo.Width
    picInfo.Move picElemImage.ScaleWidth + 5, picInfo.Top, picElemWinInfo.ScaleWidth - picElemImage.ScaleWidth - picElemInfo.ScaleWidth - 10
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    
    txtDataDiscoverer.Move txtDataDiscoverer.Left, txtDataDiscoverer.Top, picInfo.ScaleWidth - lblDiscoverer.Width - 5
    txtDataCountry.Move txtDataCountry.Left, txtDataCountry.Top, picInfo.ScaleWidth - lblDiscoverer.Width - 5
    txtDataDate.Move txtDataDate.Left, txtDataDate.Top, picInfo.ScaleWidth - lblDiscoverer.Width - 5
    txtDataUses.Move txtDataUses.Left, txtDataUses.Top, picInfo.ScaleWidth - lblDiscoverer.Width - 5
End Sub

Private Sub picPlayElem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picPT.MousePointer = vbDefault
End Sub

Private Sub picPlayElem_Resize()
    On Error Resume Next
    
    If lblMsg.Visible Then
        lblMsg.Move 0, (picPlayElem.Height - lblMsg.Height) / 2, _
                    picPlayElem.ScaleWidth, lblMsg.Height
    End If
    
    If lblMsgName.Visible Then
        lblMsgName.Move 0, (picPlayElem.Height - lblMsgName.Height) / 2
        txtElement.Move lblMsgName.Left + lblMsgName.Width, (picPlayElem.ScaleHeight - txtElement.Height) / 2, _
                        picPlayElem.ScaleWidth - lblMsgName.Width - 5
        cmdOk.Move (picPlayElem.ScaleWidth - cmdOk.Width) / 2, lblMsgName.Top + lblMsgName.Height + 8
    End If
    
    If lblMsgErr.Visible Then
        lblMsgErr.Move (picPlayElem.ScaleWidth - lblMsgErr.Width) / 2, lblMsgName.Top - lblMsgErr.Height - 8, picPlayElem.ScaleWidth, lblMsgErr.Height
    End If
    
    If cmdAnother.Visible Then
        cmdAnother.Move (picPlayElem.ScaleWidth - cmdAnother.Width) / 2, _
                        lblMsg.Top + lblMsg.Height + 1
    End If
End Sub

Private Sub picPT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GameType = 1 Then Exit Sub
    
    If picPT.MousePointer = vbCustom Then
        Set picPT.MouseIcon = LoadResPicture(102, vbResCursor)
    End If
End Sub

Private Sub picPT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i       As Integer
    Dim curidx  As Integer
    Dim IsHover As Boolean
    
    Static IsHandCursor As Boolean
    
    curidx = -1
    For i = 1 To PTTable.Count
        If Trim$(PTTable(i).Element) <> "" Then
            If (X >= PTTable(i).X1) And (X <= PTTable(i).X2) Then
                If (Y >= PTTable(i).Y1) And (Y <= PTTable(i).Y2) Then
                    ShowElemInfo PTElem(PTTable(i).Atomic_Number).Atomic_Number
                    curidx = i
                    Exit For
                End If
            End If
        End If
    Next i
    
    If GameType = 1 Then Exit Sub
    
    If curidx <> -1 Then
        For i = 1 To SelElem.Count
            If SelElem(i) = PTElem(PTTable(curidx).Atomic_Number).Atomic_Number Then
                IsHover = True: CurElem = curidx
                Exit For
            End If
        Next i
    End If
    
    If IsHover Then
        If Not IsHandCursor Then
            picPT.MousePointer = vbCustom
            Set picPT.MouseIcon = LoadResPicture(101, vbResCursor)
            IsHandCursor = True
        End If
    Else
        If IsHandCursor Then
            picPT.MousePointer = vbDefault
            IsHandCursor = False
        End If
    End If
End Sub

Private Sub picPT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GameType = 1 Then Exit Sub
    
    If picPT.MousePointer = vbCustom Then
        Dim sTemp As String
        
        Set picPT.MouseIcon = LoadResPicture(101, vbResCursor)
        
        If SelElem.Count > 0 Then
            If PTElem(PTTable(CurElem).Atomic_Number).Atomic_Number = SelElem(1) Then
                lblMsg.Caption = "Right! The symbol for " & _
                                  PTElem(PTTable(CurElem).Atomic_Number).Name & " is " & _
                                  PTElem(PTTable(CurElem).Atomic_Number).Name
                cmdAnother.Visible = False
                Call picPlayElem_Resize
                Call Start
                
                PlaySound "applause.wav"
            Else
                lblMsg.Caption = "The correct symbol for " & _
                                 PTElem(SelElem(1)).Name & " is " & PTElem(SelElem(1)).Symbol & _
                                 ", not " & PTElem(PTTable(CurElem).Atomic_Number).Symbol & "." & vbCrLf & _
                                 PTElem(PTTable(CurElem).Atomic_Number).Symbol & _
                                 " is the symbol for " & _
                                 PTElem(PTTable(CurElem).Atomic_Number).Name
                cmdAnother.Visible = True
                Call picPlayElem_Resize
                Call Clear
                
                PlaySound "beep.wav"
            End If
        End If
    End If
    
    picPlayElem.MousePointer = vbDefault
End Sub

Private Sub picPT_Resize()
    Dim i         As Integer
    Dim j         As Integer
    Dim X         As Integer
    Dim Y         As Integer
    Dim sw        As Single  ' scale width
    Dim sh        As Single  ' scale height
    Dim oTemp     As Object
    Dim arrTemp() As String
    Dim sTemp     As String
    Dim curidx    As Integer
    
    sw = picPT.ScaleWidth / 18
    sh = picPT.ScaleHeight / 9 - 1
    
    For Y = 0 To 8
        For X = 0 To 18
            Set oTemp = New cPTTable
            curidx = Y * 18 + X + 1
            
            If curidx <= 162 Then
                PTTable(curidx).X1 = X * sw
                PTTable(curidx).X2 = (X + 1) * sw
                
                If Y < 7 Then
                    PTTable(curidx).Y1 = Y * sh
                    PTTable(curidx).Y2 = (Y + 1) * sh
                Else
                    PTTable(curidx).Y1 = Y * sh + 8
                    PTTable(curidx).Y2 = (Y + 1) * sh + 8
                End If
            End If
        Next X
    Next Y
    
    picPT.Cls
    picPT.Font.Bold = True
    picPT.Font.Size = 10
    For i = 1 To PTTable.Count
        If PTTable(i).Element <> "" Then
            DrawElem i, False
        End If
    Next i
    
    For j = 1 To SelElem.Count
        For i = 1 To PTTable.Count
            If PTTable(i).Atomic_Number = SelElem(j) Then
                DrawElem i, True
                Exit For
            End If
        Next i
    Next j
    
    picPT.Refresh
End Sub

Private Function GetElemBkColor(ByVal idElem As Integer) As Long
    Select Case idElem
    Case 1, 2, 10, 18, 36, 54, 86, 118
        ' H
        GetElemBkColor = RGB(255, 255, 255)
    Case 3, 11, 19, 37, 55, 87
        ' Li, Na, K, Rb, Cs, Fr
        GetElemBkColor = RGB(255, 255, 153)
    Case 4, 12, 20, 38, 56, 88 To 103
        ' Be, Mg, Ca, Sr, Ba, Ra, N, P, As, Sb, Bi, (), Ac, Th, Pa, U, Np, Pu Am, Cm, Bk, Cf, Es, Fm, Md, No, Lr
        GetElemBkColor = RGB(102, 204, 51)
    Case 21, 39, 22, 40, 72, 104, 23, 41, 73, 105, 24, 42, 74, 106, 25, 43, 75, 105, 25, 43, 75, 107, 7, 15, 33, 51, 83, 115
        ' Sc, Y, Ti, Zr, Hf, Rf, V, Nb, Ta, Db, Cr, Mo, W, Sg, Mn, Tc, Re, Bh, N, P, As, Sb, Bi, ()
        GetElemBkColor = RGB(0, 153, 255)
    Case 6, 14, 32, 50, 82, 114, 57 To 71
        GetElemBkColor = RGB(255, 204, 0)
    Case 5, 13, 31, 49, 81, 113
        GetElemBkColor = RGB(255, 204, 255)
    Case 8, 16, 34, 52, 84, 116
        GetElemBkColor = RGB(204, 204, 255)
    Case 9, 17, 35, 53, 85, 117
        GetElemBkColor = RGB(255, 219, 157)
    Case Else
        GetElemBkColor = RGB(153, 153, 255)
    End Select
End Function

Private Function GetElemForeColor(ByVal idElem As Integer) As Long
    Select Case idElem
    Case 1, 7, 8, 9, 17, 2, 10, 18, 36, 54, 86 ' H
        GetElemForeColor = RGB(255, 0, 0)
    Case 43, 61, 104 To 117, 93 To 103
        GetElemForeColor = RGB(255, 255, 255)
    Case 55, 87, 80, 31, 35
        GetElemForeColor = RGB(0, 51, 153)
    Case Else
        GetElemForeColor = RGB(0, 0, 0)
    End Select
End Function

Private Function GetColBkColor(ByVal idCol As Integer) As Long
    Select Case idCol
    Case 0
        GetColBkColor = RGB(255, 255, 153)
    Case 1
        GetColBkColor = RGB(102, 204, 51)
    Case 2 To 6, 14
        GetColBkColor = RGB(0, 153, 255)
    Case 7 To 11
        GetColBkColor = RGB(153, 153, 255)
    Case 12
        GetColBkColor = RGB(255, 204, 255)
    Case 13
        GetColBkColor = RGB(255, 204, 0)
    Case 15
        GetColBkColor = RGB(204, 204, 255)
    Case 16
        GetColBkColor = RGB(255, 219, 157)
    Case Else
        GetColBkColor = RGB(255, 255, 255)
    End Select
End Function

Private Sub DrawElem(ByVal idElem As Integer, ByVal Highlight As Boolean)
    If Not Highlight Then
        picPT.Line (PTTable(idElem).X1, PTTable(idElem).Y1)- _
               (PTTable(idElem).X2, PTTable(idElem).Y2), _
               GetElemBkColor(PTTable(idElem).Atomic_Number), BF
        picPT.Line (PTTable(idElem).X1, PTTable(idElem).Y1)- _
               (PTTable(idElem).X2, PTTable(idElem).Y2), &H0, B

        picPT.ForeColor = &H0
        picPT.FontUnderline = False
        TextOut picPT.hdc, PTTable(idElem).X1 + (PTTable(idElem).X2 - PTTable(idElem).X1 - _
                           picPT.TextWidth(PTTable(idElem).Atomic_Number)) / 2, _
                           PTTable(idElem).Y1, PTTable(idElem).Atomic_Number, Len(PTTable(idElem).Atomic_Number)
        picPT.ForeColor = GetElemForeColor(PTTable(idElem).Atomic_Number)
        TextOut picPT.hdc, PTTable(idElem).X1 + (PTTable(idElem).X2 - PTTable(idElem).X1 - _
                           picPT.TextWidth(PTTable(idElem).Element)) / 2, _
                           PTTable(idElem).Y1 + (PTTable(idElem).Y2 - PTTable(idElem).Y1 - _
                           picPT.TextHeight(PTTable(idElem).Element)) / 2 + 5, PTTable(idElem).Element, Len(PTTable(idElem).Element)
    Else
        picPT.Line (PTTable(idElem).X1, PTTable(idElem).Y1)- _
                   (PTTable(idElem).X2, PTTable(idElem).Y2), _
                   RGB(153, 255, 102), BF
        picPT.Line (PTTable(idElem).X1, PTTable(idElem).Y1)- _
                   (PTTable(idElem).X2, PTTable(idElem).Y2), &H0, B

        picPT.ForeColor = &HFF0000
        picPT.FontUnderline = True
        TextOut picPT.hdc, PTTable(idElem).X1 + (PTTable(idElem).X2 - PTTable(idElem).X1 - _
                           picPT.TextWidth(PTTable(idElem).Atomic_Number)) / 2, _
                           PTTable(idElem).Y1, PTTable(idElem).Atomic_Number, Len(PTTable(idElem).Atomic_Number)
        TextOut picPT.hdc, PTTable(idElem).X1 + (PTTable(idElem).X2 - PTTable(idElem).X1 - _
                           picPT.TextWidth(PTTable(idElem).Element)) / 2, _
                           PTTable(idElem).Y1 + (PTTable(idElem).Y2 - PTTable(idElem).Y1 - _
                           picPT.TextHeight(PTTable(idElem).Element)) / 2 + 5, PTTable(idElem).Element, Len(PTTable(idElem).Element)
    End If
End Sub

Private Sub ShowElemInfo(ByVal idElem As Integer)
    Dim i     As Integer
    Dim th    As Integer ' text height
    Dim sxbox As Integer
    Dim sybox As Integer
    Dim exbox As Integer
    Dim eybox As Integer
    Dim bval  As Boolean
    
    sxbox = 100
    sybox = 1
    exbox = 215
    eybox = picElemInfo.ScaleHeight - 1
    th = picElemInfo.TextHeight("H")
    
    picElemInfo.Cls
    picElemInfo.DrawWidth = 1
    picElemInfo.Line (sxbox, sybox)-(exbox, eybox), GetElemBkColor(idElem), BF
    picElemInfo.Line (sxbox, sybox)-(exbox, eybox), &H0, B
    
    picElemInfo.FontBold = True
    picElemInfo.FontSize = 12
    SetTextColor picElemInfo.hdc, &H0
    TextOut picElemInfo.hdc, sxbox + 5, 1, PTElem(idElem).Atomic_Number, Len(PTElem(idElem).Atomic_Number)

    picElemInfo.FontSize = 8
    picElemInfo.FontBold = False
    SetTextColor picElemInfo.hdc, &H0
    TextOut picElemInfo.hdc, exbox - picElemInfo.TextWidth(PTElem(idElem).Atomic_Weight) - 6, th * 0 + 2, PTElem(idElem).Atomic_Weight, Len(PTElem(idElem).Atomic_Weight)
    TextOut picElemInfo.hdc, exbox - picElemInfo.TextWidth(PTElem(idElem).Oxidation_States) - 6, th * 1 + 2, PTElem(idElem).Oxidation_States, Len(PTElem(idElem).Oxidation_States)
    TextOut picElemInfo.hdc, exbox - picElemInfo.TextWidth(PTElem(idElem).Melting_Point) - 6, th * 2 + 2, PTElem(idElem).Melting_Point, Len(PTElem(idElem).Melting_Point)
    TextOut picElemInfo.hdc, exbox - picElemInfo.TextWidth(PTElem(idElem).Boiling_Point) - 6, th * 3 + 2, PTElem(idElem).Boiling_Point, Len(PTElem(idElem).Boiling_Point)
    TextOut picElemInfo.hdc, exbox - picElemInfo.TextWidth(PTElem(idElem).Density) - 6, th * 4 + 2, PTElem(idElem).Density, Len(PTElem(idElem).Density)
    TextOut picElemInfo.hdc, exbox - picElemInfo.TextWidth(PTElem(idElem).Electronegativity) - 6, th * 5 + 2, PTElem(idElem).Electronegativity, Len(PTElem(idElem).Electronegativity)
    TextOut picElemInfo.hdc, exbox - picElemInfo.TextWidth(PTElem(idElem).Atomic_Radius) - 6, th * 6 + 2, PTElem(idElem).Atomic_Radius, Len(PTElem(idElem).Atomic_Radius)
    
    picElemInfo.FontBold = True
    picElemInfo.FontSize = 10
    th = picElemInfo.TextHeight("H")
    
    For i = 1 To SelElem.Count
        If SelElem(i) = idElem Then
            bval = True
            Exit For
        End If
    Next i
    
    If bval = True Then
        TextOut picElemInfo.hdc, sxbox + (exbox - sxbox - 1) / 2, picElemInfo.ScaleHeight - th - 4, "?", 1
    Else
        TextOut picElemInfo.hdc, sxbox + (exbox - sxbox - picElemInfo.TextWidth(PTElem(idElem).Name)) / 2, picElemInfo.ScaleHeight - th - 4, PTElem(idElem).Name, Len(PTElem(idElem).Name)
    End If
    
    picElemInfo.FontBold = False
    picElemInfo.FontSize = 24
    th = picElemInfo.TextHeight("H")
    SetTextColor picElemInfo.hdc, GetElemForeColor(idElem)
    
    If (idElem >= 113) And (idElem <= 118) Then
        TextOut picElemInfo.hdc, sxbox + (exbox - sxbox - picElemInfo.TextWidth(PTElem(idElem).Symbol)) / 2, th * 1, PTElem(idElem).Symbol, Len(PTElem(idElem).Symbol)
    Else
        For i = 1 To SelElem.Count
            If SelElem(i) = idElem Then
                bval = True
                Exit For
            End If
        Next i
                
        If bval Then
            TextOut picElemInfo.hdc, sxbox + 5, th * 1, "?", 1
            lblElem.Caption = "?"
        Else
            TextOut picElemInfo.hdc, sxbox + 5, th * 1, PTElem(idElem).Symbol, Len(PTElem(idElem).Symbol)
            lblElem.Caption = PTElem(idElem).Name
        End If
    End If
    
    Dim sx      As Integer
    Dim sy      As Integer
    Dim dx      As Integer
    Dim dy      As Integer
    Dim BmpW    As Integer
    Dim BmpH    As Integer
    Dim TempPic As New StdPicture
    
    If Dir$(GetCurPath & "\pictures\" & idElem & ".jpg") <> "" Then
        Set TempPic = LoadPicture(GetCurPath & "\pictures\" & idElem & ".jpg")
    Else
        If Dir$(GetCurPath & "\pictures\0.jpg") <> "" Then
            Set TempPic = LoadPicture(GetCurPath & "\pictures\0.jpg")
        End If
    End If
    
    If TempPic.Handle <> 0 Then
        BmpW = ScaleX(TempPic.Width, vbHimetric, vbPixels)
        BmpH = ScaleY(TempPic.Height, vbHimetric, vbPixels)
        
        With picElemImage
            If (BmpW <= .ScaleWidth) And (BmpH <= .ScaleHeight) Then
                sx = (.ScaleWidth - BmpW) / 2
                sy = (.ScaleHeight - BmpH) / 2
                dx = BmpW
                dy = BmpH
            ElseIf (BmpW <= .ScaleWidth) And (BmpH > .ScaleHeight) Then
                sx = (.ScaleWidth - BmpW) / 2
                sy = 0
                dx = BmpW
                dy = .ScaleHeight
            ElseIf (BmpW > .ScaleWidth) And (BmpH <= .ScaleHeight) Then
                sx = 0
                sy = (.ScaleHeight - BmpH) / 2
                dx = .ScaleWidth
                dy = BmpH
            ElseIf (BmpW > .ScaleWidth) And (BmpH > .ScaleHeight) Then
                sx = 0
                sy = 0
                dx = .ScaleWidth
                dy = .ScaleHeight
            End If
            
            .Cls
'            .PaintPicture TempPic, sx, sy, dx, dy
        End With
    End If
    
    picElemInfo.FontBold = False
    picElemInfo.FontSize = 8
    picElemInfo.DrawWidth = 3
    th = picElemInfo.TextHeight("H")
    SetTextColor picElemInfo.hdc, &H0
    TextOut picElemInfo.hdc, 0, 4, "Atomic Number", Len("Atomic Number")
    picElemInfo.Line (picElemInfo.TextWidth("Atomic Number") + 2, 10)-(sxbox + 2, 10), &HFF
    TextOut picElemInfo.hdc, 0, 50, "Atomic Symbol", Len("Atomic Symbol")
    picElemInfo.Line (picElemInfo.TextWidth("Atomic Symbol") + 2, 56)-(sxbox + 2, 56), &HFF
    TextOut picElemInfo.hdc, 0, 95, "Name", Len("Name")
    picElemInfo.Line (picElemInfo.TextWidth("Name") + 2, 102)-(sxbox + 2, 102), &HFF
    TextOut picElemInfo.hdc, exbox + 10, 2, "Atomic Weight", Len("Atomic Weight")
    TextOut picElemInfo.hdc, exbox + 10 + picElemInfo.TextWidth("Atomic Weight"), 0, "a", 1
    picElemInfo.Line (exbox - 4, 7)-(exbox + 8, 7), &HFF
    TextOut picElemInfo.hdc, exbox + 10, th + 2, "Oxidation States (Valence)", Len("Oxidation States (Valence)")
    TextOut picElemInfo.hdc, exbox + 10 + picElemInfo.TextWidth("Oxidation States (Valence)"), th - 0.5, "b", 1
    picElemInfo.Line (exbox - 4, th + 9)-(exbox + 8, th + 9), &HFF
    TextOut picElemInfo.hdc, exbox + 10, th * 2 + 2, "Melting Point (°C)", Len("Melting Point (°C)")
    TextOut picElemInfo.hdc, exbox + 10 + picElemInfo.TextWidth("Melting Point (°C)"), th * 2 - 1.5, "c", 1
    picElemInfo.Line (exbox - 4, th * 2 + 9)-(exbox + 8, th * 2 + 9), &HFF
    TextOut picElemInfo.hdc, exbox + 10, th * 3 + 2, "Boiling Point (°C)", Len("Boiling Point (°C)")
    picElemInfo.Line (exbox - 4, th * 3 + 9)-(exbox + 8, th * 3 + 9), &HFF
    TextOut picElemInfo.hdc, exbox + 10, th * 4 + 2, "Density   (g/cm  )", Len("Density   (g/cm  )")
    TextOut picElemInfo.hdc, exbox + 10 + picElemInfo.TextWidth("Density"), th * 4 - 0.5, "d", 1
    TextOut picElemInfo.hdc, exbox + 10 + picElemInfo.TextWidth("Density   (g/cm"), th * 4 - 0.5, "3", 1
    picElemInfo.Line (exbox - 4, th * 4 + 9)-(exbox + 8, th * 4 + 9), &HFF
    TextOut picElemInfo.hdc, exbox + 10, th * 5 + 2, "Electronegativity", Len("Electronegativity")
    picElemInfo.Line (exbox - 4, th * 5 + 9)-(exbox + 8, th * 5 + 9), &HFF
    TextOut picElemInfo.hdc, exbox + 10, th * 6 + 2, "Atomic Radius (1x10     m)", Len("Atomic Radius (1x10     m)")
    TextOut picElemInfo.hdc, exbox + 10 + picElemInfo.TextWidth("Atomic Radius (1x10"), th * 6 - 0.5, "-10", 3
    picElemInfo.Line (exbox - 4, th * 6 + 9)-(exbox + 8, th * 6 + 9), &HFF
    picElemInfo.Refresh
    
    txtDataDiscoverer.Text = PTElem(idElem).Discoverer
    txtDataCountry.Text = PTElem(idElem).Country
    txtDataDate.Text = PTElem(idElem).DDate
    txtDataUses.Text = PTElem(idElem).Uses
End Sub

Private Sub picRow_Resize()
    Dim i         As Integer
    Dim j         As Integer
    Dim X         As Integer
    Dim Y         As Integer
    Dim sw        As Single  ' scale width
    Dim sh        As Single  ' scale height
    Dim oTemp     As Object
    Dim sTemp     As String
    Dim curidx    As Integer
    
    sw = picRow.ScaleWidth
    sh = picPT.ScaleHeight / 9 - 1
    
    picRow.Cls
    picRow.FontBold = True
    picRow.FontSize = 12
    picRow.Line (0, 0)-(picRow.ScaleWidth, picCol.ScaleHeight - 1), &H0, B
    picRow.Line (0, picCol.ScaleHeight - 1)-(picRow.ScaleWidth - 1, _
                 picPT.Top - picCol.ScaleHeight + picCol.ScaleHeight), &H0, B
    For i = 0 To 6
        picRow.Line (0, i * sh + picPT.Top)-(sw, (i + 1) * sh + picPT.Top), &H0, B
        TextOut picRow.hdc, (picRow.ScaleWidth - picRow.TextWidth(CStr(i + 1))) / 2, _
                             picPT.Top + sh / 2 + (i * sh - picRow.TextHeight(CStr(i + 1)) / 2), CStr(i + 1), Len(CStr(i + 1))
    Next i

    picRow.FontSize = 8
    TextOut picRow.hdc, (picRow.ScaleWidth - picRow.TextWidth("Group")) / 2, _
                        (picCol.ScaleHeight - picRow.TextHeight("Group") - 4) / 2, _
                        "Group", 5
    TextOut picRow.hdc, (picRow.ScaleWidth - picRow.TextWidth("Group")) / 2, _
                        picCol.ScaleHeight + (picPT.Top - picCol.ScaleHeight - picRow.TextHeight("Period")) / 2, _
                        "Period", 6
    picRow.Refresh
End Sub

Public Sub Start()
    Dim sTemp As String
    
    cmdOk.Visible = False
    cmdAnother.Visible = False
    txtElement.Visible = False
    lblMsg.Visible = False
    lblMsgName.Visible = False
    lblMsgErr.Visible = False
    
    If GameType = 0 Then
        Dim i         As Integer
        Dim j         As Integer
        Dim cntr      As Integer
        Dim rndval    As Integer
        Dim isexist   As Boolean
        Dim TotalElem As Integer
        
        Call Clear
        
        SelElem.Add Generate_Random_Number(1, 112)
        
        If Level = 1 Then
            TotalElem = 10
        ElseIf Level = 2 Then
            TotalElem = 20
        Else
            TotalElem = 40
        End If
        
        Do While cntr < TotalElem
            isexist = False
            rndval = Generate_Random_Number(1, 112)
            
            For i = 1 To SelElem.Count
                If rndval = SelElem(i) Then
                    isexist = True
                    Exit For
                End If
            Next i
            
            If Not isexist Then
                SelElem.Add rndval
                cntr = cntr + 1
            End If
            DoEvents
        Loop
        
        lblMsg.Visible = True
        If Left$(UCase$(lblMsg.Caption), 5) = "RIGHT" Then
            lblMsg.Caption = lblMsg.Caption & vbCrLf & "Click " & PTElem(SelElem(1)).Name
        Else
            lblMsg.Caption = "Click " & PTElem(SelElem(1)).Name
        End If
                
        Call picPlayElem_Resize
    Else
        Call Clear
        SelElem.Add Generate_Random_Number(1, 112)
        
        lblMsgName.Visible = True
        lblMsg.Caption = "Name the highlighted Element : "
        cmdOk.Visible = True
        txtElement.Visible = True
        lblMsgErr.Visible = True
        txtElement.SetFocus
        
        Call picPlayElem_Resize
    End If
    
    For j = 1 To SelElem.Count
        For i = 1 To PTTable.Count
            If PTTable(i).Atomic_Number = SelElem(j) Then
                DrawElem i, True
                picPT.Refresh
                Sleep 30
                Exit For
            End If
        Next i
    Next j
    
    If GameType = 0 Then
        If Tutorial Then
            tmrBlink.Enabled = True
        Else
            tmrBlink.Enabled = False
        End If
    Else
        tmrBlink.Enabled = True
    End If
    
    picPT.MousePointer = vbDefault
    picPlayElem.Refresh
End Sub

Private Sub Clear()
    Dim j As Integer
    Dim i As Integer
    
    For j = 1 To SelElem.Count
        For i = 1 To PTTable.Count
            If PTTable(i).Atomic_Number = SelElem(j) Then
                DrawElem i, False
                picPT.Refresh
                Sleep 30
                Exit For
            End If
        Next i
    Next j
    
    Set SelElem = Nothing
End Sub

Private Sub tmrBlink_Timer()
    Dim i          As Integer
    Static bSwitch As Boolean
    
    On Error Resume Next
    
    If bSwitch = False Then
        bSwitch = True
    Else
        bSwitch = False
    End If
    
    If PTTable.Count > 0 Then
        For i = 1 To PTTable.Count
            If PTTable(i).Atomic_Number = SelElem(1) Then
                DrawElem i, bSwitch
                Exit For
            End If
        Next i
    End If
End Sub
