VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About IPDG3 version"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10425
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   6240
      Top             =   6000
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   7680
      Top             =   6000
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   10365
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   10425
   End
   Begin VB.PictureBox picIPDG3 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      Picture         =   "frmAbout.frx":1CCA
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   676
      TabIndex        =   14
      Top             =   600
      Width           =   10200
   End
   Begin VB.PictureBox picScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      ScaleHeight     =   205
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   357
      TabIndex        =   0
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Label lblResume 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "View My Resume Online"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   5400
      Width           =   4815
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Location :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "Aurora, IL. USA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "630.236.5584"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Voice :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label lblEmail 
      BackColor       =   &H00000000&
      Caption         =   "info@ipdg3.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   8
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblWebsite 
      BackColor       =   &H00000000&
      Caption         =   "www.ipdg3.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Email :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Website:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      Caption         =   "George Goehring"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Written By:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblExit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   9600
      TabIndex        =   2
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interactive PsyberTechnology Developers Group"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   690
      Left            =   5760
      TabIndex        =   1
      Top             =   2640
      Width           =   4575
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Const mstVersion As String = "Version: 1.0"
Const mstTitle As String = "IPDG3 Spell Checker"
Const mstDescription As String = "IPDG Spell Checker is designed for people like me who can not spell. LOL  We have added some other functionality to the program but use it how ever you want."

'the actual text to scroll. This could also be loaded in from a text file
Const ScrollText As String = mstTitle & _
                             vbCrLf & vbCrLf & mstVersion & _
                             vbCrLf & vbCrLf & mstDescription & _
                             vbCrLf & vbCrLf & "For more information please " & _
                             "visit our website for details..." & _
                             vbCrLf & vbCrLf & "www.ipdg3.com - info@ipdg3.com" & _
                             vbCrLf & vbCrLf & "Voice: 630.236.5584 Aurora, IL. USA"
                             
Dim EndingFlag As Boolean

Private Sub Form_Activate()

    RunMain

End Sub

Private Sub Form_Load()

    picScroll.ForeColor = vbYellow
    picScroll.FontSize = 12

    frmAbout.Caption = "About IPDG3 Version: " & App.Major & "." & App.Minor

End Sub

Private Sub RunMain()
Dim LastFrameTime As Long
Const IntervalTime As Long = 40
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long 'Upper left point of drawing rect
Dim RectHeight As Long

    'show the form
    frmAbout.Refresh

    'Get the size of the drawing rectangle by suppying the DT_CALCRECT constant
    rt = DrawText(picScroll.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)

    If rt = 0 Then 'err
      MsgBox "Error scrolling text", vbExclamation
      EndingFlag = True
    Else
      DrawingRect.Top = picScroll.ScaleHeight
      DrawingRect.Left = 0
      DrawingRect.Right = picScroll.ScaleWidth
      'Store the height of The rect
      RectHeight = DrawingRect.Bottom
      DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
    End If

    Do While Not EndingFlag
      If GetTickCount() - LastFrameTime > IntervalTime Then
                    
        picScroll.Cls
        
        DrawText picScroll.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
        'update the coordinates of the rectangle
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        
        'control the scolling and reset if it goes out of bounds
        If DrawingRect.Top < -(RectHeight) Then 'time to reset
            DrawingRect.Top = picScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
        End If
        
        picScroll.Refresh
        
        LastFrameTime = GetTickCount()
        
      End If
    
      DoEvents
    Loop

Unload Me
Set frmAbout = Nothing

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblExit.ForeColor = vbYellow
    lblName.ForeColor = vbYellow
    lblWebsite.ForeColor = vbYellow
    lblEmail.ForeColor = vbYellow
    lblResume.ForeColor = vbYellow
    Me.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    EndingFlag = True
   
End Sub

Private Sub lblEmail_Click()

    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Are you On-Line right Now..?"   ' Define message.
    Style = vbYesNo + vbInformation ' Define buttons.
    Title = "On-Line Confirmation"  ' Define title.
    
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)

    If Response = vbYes Then
      sendemail
    Else
      MsgBox "This Will launch your default web browser and goto www.ipdg3.com so you must be on-line ...", vbInformation, "Connect to use this option"
    End If

End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblEmail.ForeColor = vbRed

End Sub

Private Sub lblExit_Click()

    Beep

    EndingFlag = True

End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblExit.ForeColor = vbRed

End Sub

Private Sub lblName_Click()

    MsgBox "Thanx for d/l and reviewing my Digital Brochure...", , "Thanx"
    
End Sub

Private Sub lblName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblName.ForeColor = vbRed
    Me.MousePointer = vbArrowQuestion

End Sub

Private Sub lblResume_Click()

    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Are you On-Line right Now..?"   ' Define message.
    Style = vbYesNo + vbInformation ' Define buttons.
    Title = "On-Line Confirmation"  ' Define title.
    
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)

    If Response = vbYes Then
      gotoResume
    Else
      MsgBox "This will launch your default web browser and goto my on-line RESUME you must be on-line ...", vbInformation, "Connect to use this option"
    End If

End Sub

Private Sub lblResume_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblResume.ForeColor = vbRed

End Sub

Private Sub lblWebsite_Click()

    Dim Msg, Style, Title, Help, Ctxt, Response, MyString
    Msg = "Are you On-Line right Now..?"   ' Define message.
    Style = vbYesNo + vbInformation ' Define buttons.
    Title = "On-Line Confirmation"  ' Define title.
    
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)

    If Response = vbYes Then
      gotoIPDG3
    Else
      MsgBox "This Will launch your default web browser and goto www.ipdg3.com so you must be on-line ...", vbInformation, "Connect to use this option"
    End If

End Sub

Private Sub lblWebsite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblWebsite.ForeColor = vbRed

End Sub

Private Sub Timer1_Timer()
Const MESSAGE = "Q: What can IPDG3 do for you...?       A: Choose one of the topics from the listbox below..."

Static done_before As Boolean
Static msg_width As Single
Static X As Single

    If Not done_before Then
        msg_width = Picture1.TextWidth(MESSAGE)
        done_before = True
        X = Picture1.ScaleWidth
    End If

    Picture1.Cls
    Picture1.CurrentX = X
    Picture1.CurrentY = 0
    Picture1.Print MESSAGE
    
    X = X - 30
    If X < -msg_width Then X = Picture1.ScaleWidth

End Sub

Private Sub Timer2_Timer()
Dim X As Integer
Dim Y As Integer
Dim z As Integer
       
    Y = Rnd() * 256
    X = Rnd() * 256
    z = Rnd() * 256
        
    Picture1.ForeColor = RGB(X, Y, z)

End Sub

