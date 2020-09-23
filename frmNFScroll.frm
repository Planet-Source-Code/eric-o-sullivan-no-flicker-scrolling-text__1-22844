VERSION 5.00
Begin VB.Form frmNFScroll 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "No-Flicker Scrolling Text"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2055
      Left            =   0
      Picture         =   "frmNFScroll.frx":0000
      ScaleHeight     =   2055
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Timer timScroll 
      Interval        =   1
      Left            =   0
      Top             =   2040
   End
End
Attribute VB_Name = "frmNFScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim NextLine As String
Dim GotInfo As String

'display the background colour

'set the display text
NextLine = Chr(13) & Chr(10)
GotInfo = "No-Flicker Scrolling Text Project  v " & App.Major & "." & App.Minor
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "CompApp Technologies™"
GotInfo = GotInfo & NextLine & "Copyright 2000"
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "Support"
GotInfo = GotInfo & NextLine & "If there are any problems with this"
GotInfo = GotInfo & NextLine & "product, please don't hesitate to "
GotInfo = GotInfo & NextLine & "contact us."
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "E-mail"
GotInfo = GotInfo & NextLine & "DiskJunky@hotmail.com"
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "Web Site"
GotInfo = GotInfo & NextLine & "http://www.compapp.co-ltd.com"
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & "Programmer"
GotInfo = GotInfo & NextLine & "Eric O'Sullivan"
GotInfo = GotInfo & NextLine & ""
GotInfo = GotInfo & NextLine & ""

Call EnterText(GotInfo)
End Sub

Private Sub timScroll_Timer()
'scroll the credits up wards.

'in nanoseconds
Const TimePerPixel = 60

Dim Speed As Integer
Static BackArea As Rect
Static Tick As Long

If Tick = 0 Then
    'set the co-ordinates of the background
    picText.Cls
    BackArea.Top = 0
    BackArea.Left = 0
    BackArea.Right = (picText.ScaleWidth / Screen.TwipsPerPixelX)
    BackArea.Bottom = (picText.ScaleHeight / Screen.TwipsPerPixelY)
    
    'set the background of the text
    Call LoadOldBack(picText, BackArea)
    Tick = GetTickCount
End If

'if X nanoseconds have elapsed, move text up one pixel
If (Tick + TimePerPixel) < GetTickCount Then
    'move one pixel at a time
    Speed = Screen.TwipsPerPixelY
    
    'move the text up one pixel (Speed)
    Call MoveText(picText, BackArea, Speed, vbCentreAlign)
    
    'wait until you can move the text up one pixel
    Tick = GetTickCount
End If
End Sub

