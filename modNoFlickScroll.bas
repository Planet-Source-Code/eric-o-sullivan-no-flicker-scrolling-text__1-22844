Attribute VB_Name = "modNoFlickScroll"
'This module was created by Eric O'Sullivan
'
'I created this module not as a control but as a standard module
'because I wanted to be able to use this code without compiling
'another .ocx and adding it to some of my already large projects.
'
'-------------------------------------------------------------------------------
'The module works by creating two off-screen bitmaps at run-time.
'
'The first bitmap (we'll call it the permanant bitmap) holds the
'background where the text is going to be displayed. This bitmap
'will not be updated unless you specify it in the parameters of the
'"ShowText" procedure. Be warned though; you must prepare the
'area where the text is going to be shown before updating the
'permanent bitmap or else the bitmap will contain a picture you
'didn't want to dislpay. I suggest you do this by using Cls or
're-loading (or re-painting) a picture onto a form - This should clear
'the area for the bitmap to take a snap shot of the area.
'This bitmap must also be deleted if you are not using the module.
'The bitmap takes up memory and must be "cleaned" up before the
'program is terminated. This isn't strictly necessary if you are
'terminating the program, but it's good practice to do so. You can do
'this by calling the "DeleteOldBack" procedure in the QueryUnload
'event of the main form.
'
'The second bitmap (we'll call this one the temperory bitmap) is also
'created dynamically. This bitmap is used to construct the text and
'the permanent bitmap before displaying the result on the screen.
'First the contents of the permanent bitmap are copied into the
'temperory bitmap, using BitBlt.
'Next the font is created for the text and the alignment of the text
'within the area specified is set. The font settings are taken from the
'Font property of the form passed to the "ShowText" procedure. The
'colour can be specified as part of the parameters also.
'Once the font has been created, we draw the text onto the
'temperory bitmap. This now contains the contents of the permanent
'bitmap and the drawn text on top of it.
'We now copy the entire contents of the temperory bitmap onto the
'specified area of the screen.
'
'In conclusion; the result of all this is that the specified area does
'not flicker because the text is updated off-screen before the
'results are copied onto the form, whole.
'
'The text is moved up by decreasing the distance of the start of the
'text to the top of the picutre box. This will include a negitive value
'if the text includes enough lines (seperated by the carriage return
'and line-feed characters (chr(13) & chr(10)) ).
'-------------------------------------------------------------------------------
'
'
'Please free to e-mail me with any queries or improvments made to
'this code. I hope it is usefull to you and you've learnt from using it
'as I did making it ;)
'You can contact me at : DiskJunky@hotmail.com
' - Enjoy!
'===========================================================


Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
Public Declare Function SetRect Lib "user32" (lpRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LogBrush) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hwnd As Long, Fill As Rect, HBrush As Long) As Integer
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LogFont) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Enum AlignText
    vbLeftAlign = 1
    vbCentreAlign = 2
    vbRightAlign = 3
End Enum

Public Type LogFont
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(1 To 32) As Byte
End Type

Public Type Rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type LogBrush
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type

'DrawText constants
Public Const DT_CENTER = &H1
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_LEFT = &H0
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_TOP = &H0
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

'CreateBrushIndirect constants
Public Const BS_HATCHED = 2
Public Const BS_HOLLOW = Null
Public Const BS_PATTERN = 3
Public Const BS_SOLID = 0

'BitBlt constants
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)

'LogFont constants
Public Const LF_FACESIZE = 32
Public Const FW_BOLD = 700
Public Const FW_DONTCARE = 0
Public Const FW_EXTRABOLD = 800
Public Const FW_EXTRALIGHT = 200
Public Const FW_HEAVY = 900
Public Const FW_LIGHT = 300
Public Const FW_MEDIUM = 500
Public Const FW_NORMAL = 400
Public Const FW_SEMIBOLD = 600
Public Const FW_THIN = 100
Public Const DEFAULT_CHARSET = 1
Public Const OUT_CHARACTER_PRECIS = 2
Public Const OUT_DEFAULT_PRECIS = 0
Public Const OUT_DEVICE_PRECIS = 5
Public Const OUT_OUTLINE_PRECIS = 8
Public Const OUT_RASTER_PRECIS = 6
Public Const OUT_STRING_PRECIS = 1
Public Const OUT_STROKE_PRECIS = 3
Public Const OUT_TT_ONLY_PRECIS = 7
Public Const OUT_TT_PRECIS = 4
Public Const CLIP_CHARACTER_PRECIS = 1
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const CLIP_EMBEDDED = 128
Public Const CLIP_LH_ANGLES = 16
Public Const CLIP_MASK = &HF
Public Const CLIP_STROKE_PRECIS = 2
Public Const CLIP_TT_ALWAYS = 32
Public Const WM_SETFONT = &H30
Public Const LF_FULLFACESIZE = 64
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_QUALITY = 0
Public Const PROOF_QUALITY = 2

'GetDeviceCaps constants
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

'store information on where what was displayed in the selected
'area before the text was displayed.
Private hDcMemOld As Long
Private hDcBmpOld As Long
Private hDcOldPtr As Long
Private OldRect As Rect

Private TextInfo() As String
Private TextLen As Integer

Public Sub MoveText(MyPic As PictureBox, DisplayRect As Rect, Speed As Integer, Alignment As AlignText, Optional Reset As Boolean)
'This procedure scrolls the text in the picture box

Static TopText As Long
Static LoadedTopText As Boolean

If (Not LoadedTopText) And (Not Reset) Then
    'only do this once
    TopText = MyPic.ScaleHeight
    LoadedTopText = True
End If

'display the text starting at the specified point
Call ShowAllText(MyPic, DisplayRect, (TopText / Screen.TwipsPerPixelY), Alignment)

If TopText < -(MyPic.TextHeight(TextInfo(0)) * (TextLen + 1)) Then
    'if it's time to re-start the scrolling process, set the TopText
    'variable.
    TopText = MyPic.ScaleHeight
Else
    'else continue scrolling the text from it's current position.
    TopText = TopText - Speed
End If
End Sub

Public Sub ShowAllText(MyPic As PictureBox, MyRect As Rect, StartHeight As Long, Align As AlignText)
'This procedure loads all the text onto a bitmap befoe blitting it
'onto the screen. I'm assuming the values in MyRect are in pixels
'and that the value in StartHeight is also in pixels.

'bitmap to copy text from, onto the form
'-----------
Dim hDcMemNew As Long
Dim hDcBmpNew As Long
Dim hDcNewPtr As Long
Dim NewRect As Rect
'-----------

Dim Result As Long
Dim StartLineHeight As Long
Dim Width As Long
Dim Height As Long

'Create off-screen bitmap and select it (also storing where it is so I
'can delete it later once the text is displayed)
Result = GetClientRect(MyPic.hwnd, NewRect)
hDcMemNew = CreateCompatibleDC(MyPic.hdc)
hDcBmpNew = CreateCompatibleBitmap(MyPic.hdc, (NewRect.Right - NewRect.Left), (NewRect.Bottom - NewRect.Top))
hDcNewPtr = SelectObject(hDcMemNew, hDcBmpNew)

'copy the old background to the off-screen bitmap
Result = BitBlt(hDcMemNew, 0, 0, (NewRect.Right - NewRect.Left), (NewRect.Bottom - NewRect.Top), hDcMemOld, 0, 0, SRCCOPY)

'-------------------------
StartLineHeight = StartHeight

For Counter = 0 To TextLen
    'display each line in the picture box if able to display
    
    'set where this line is going to go
    StartLineHeight = StartHeight + ((MyPic.TextHeight(TextInfo(Counter)) * (Counter + 0)) / Screen.TwipsPerPixelY)
    
    'see if we are able to display it
    If (StartLineHeight >= (-1 * (MyPic.TextHeight(TextInfo(Counter)) / Screen.TwipsPerPixelY))) And (StartLineHeight <= MyRect.Bottom) And (TextInfo(Counter) <> "") Then
        Call AddLineToBitmap(MyPic, TextInfo(Counter), hDcMemNew, Align, (MyRect.Right - MyRect.Left), StartLineHeight)
    End If
Next Counter
'-------------------------

'set the width and height of the bitmap to display onto.
Width = MyRect.Right - MyRect.Left
Height = MyRect.Bottom - MyRect.Top

'copy the newly built bitmap onto the picturebox
Result = BitBlt(MyPic.hdc, MyRect.Left, MyRect.Top, Width, Height, hDcMemNew, 0, 0, SRCCOPY)

'remove the bitmap object from memory before exiting the procedure
Junk = SelectObject(hDcMemNew, hDcNewPtr)
Junk = DeleteObject(hDcBmpNew)
Junk = DeleteDC(hDcMemNew)
End Sub

'ByRef mypic As PictureBox, MyString As String, Colour As Long, Top As Integer, Left As Integer, Width As Integer, Height As Integer, Optional TopText As Long, Optional Reset As Boolean)
Public Sub AddLineToBitmap(MyPic As PictureBox, MyString As String, hDcBitmap As Long, TextAlignment As AlignText, TotalWidth As Long, TextTop As Long)
'displays the text onto a form (centered)

Static OldBackLoaded As Boolean

Dim MyRect As Rect
Dim TextRect As Rect
Dim TextLen As Long
Dim Result As Long
Dim Junk As Long
Dim GotAutoRedraw As Boolean

'font details
Dim HDcOldFont As Long
Dim hDcFont As Long
Dim FontStruc As LogFont
Dim Counter As Integer
Dim Indent As Long
Dim Colour As Long

If (MyString = "") Or (Not MyPic.Visible) Then
    'exit if there is no string to display
    Exit Sub
End If

'record how many characters we're going to display
TextLen = Len(MyString)

'set the text co-ordinates
TextRect.Left = 0
TextRect.Top = TextTop
TextRect.Right = TotalWidth 'MyRect.Right
TextRect.Bottom = TextRect.Top + (MyPic.TextHeight(MyString) / Screen.TwipsPerPixelY) 'MyRect.Bottom

'Create details about the font using the forms' font details
'====================

'convert point size to pixels
FontStruc.lfHeight = -((MyPic.FontSize * GetDeviceCaps(MyPic.hdc, LOGPIXELSY)) / 72) 'mypic.FontSize
FontStruc.lfCharSet = DEFAULT_CHARSET
FontStruc.lfClipPrecision = CLIP_DEFAULT_PRECIS
FontStruc.lfEscapement = 0

'move the name of the font into the array
For Counter = 1 To Len(MyPic.FontName)
    FontStruc.lfFaceName(Counter) = Asc(Mid(MyPic.FontName, Counter, 1))
Next Counter
FontStruc.lfFaceName(Counter) = 0   'this has to be a Null terminated string

FontStruc.lfItalic = MyPic.FontItalic
FontStruc.lfUnderline = MyPic.FontUnderline
FontStruc.lfStrikeOut = MyPic.FontStrikethru
FontStruc.lfOrientation = 0
FontStruc.lfOutPrecision = OUT_DEFAULT_PRECIS
FontStruc.lfPitchAndFamily = DEFAULT_PITCH
FontStruc.lfQuality = PROOF_QUALITY

If MyPic.FontBold Then
    FontStruc.lfWeight = FW_BOLD
Else
    FontStruc.lfWeight = FW_NORMAL
End If

FontStruc.lfWidth = 0
Colour = MyPic.ForeColor
hDcFont = CreateFontIndirect(FontStruc)
HDcOldFont = SelectObject(hDcBitmap, hDcFont)
'====================

'set the alignment of the text in the area provided
Select Case TextAlignment
Case vbLeftAlign
    Indent = DT_LEFT
Case vbCentreAlign
    Indent = DT_CENTER
Case vbRightAlign
    Indent = DT_RIGHT
End Select


'Draw the text into the off-screen bitmap before copying the
'new bitmap (with the text) onto the screen.
Result = SetBkMode(hDcBitmap, TRANSPARENT)
Result = SetTextColor(hDcBitmap, Colour)
Result = DrawText(hDcBitmap, MyString, TextLen, TextRect, Indent)

'clean up by deleting the off-screen bitmap and font
Junk = SelectObject(hDcBitmap, HDcOldFont)
Junk = DeleteObject(hDcFont)
End Sub

Public Sub LoadOldBack(ByVal TheForm As PictureBox, AreaToLoad As Rect)
'This procedure will load a section of the form into an off-screen
'bitmap for use in the ShowText procedure.

Dim Result As Long

'if the form is not visible, then skip this procedure
If Not TheForm.Visible Then
    Exit Sub
End If

'if the bitmap already exists, then delete it before creating a new one
If hDcMemOld <> 0 Then
    Call DeleteOldBack
End If

'Create off-screen bitmap and select it (also storing where it is so I
'can delete it later once the text is displayed)
OldRect = AreaToLoad
hDcMemOld = CreateCompatibleDC(TheForm.hdc)
hDcBmpOld = CreateCompatibleBitmap(TheForm.hdc, (OldRect.Right - OldRect.Left), (OldRect.Bottom - OldRect.Top))
hDcOldPtr = SelectObject(hDcMemOld, hDcBmpOld)

'copy the old background to the off-screen bitmap
TheForm.AutoRedraw = False
Result = BitBlt(hDcMemOld, 0, 0, (OldRect.Right - OldRect.Left), (OldRect.Bottom - OldRect.Top), TheForm.hdc, OldRect.Left, OldRect.Top, SRCCOPY)
End Sub

Public Sub DeleteOldBack()
'This will remove the bitmap that stored what was displayed before
'the text was written to the screen, from memory.
Dim Junk As Long

If hDcMemOld = 0 Then
    'there is nothing to delete. Exit the sub-routine
    Exit Sub
End If

Junk = SelectObject(hDcMemOld, hDcOldPtr)
Junk = DeleteObject(hDcBmpOld)
Junk = DeleteDC(hDcMemOld)

hDcMemOld = 0
hDcBmpOld = 0
hDcOldPtr = 0
End Sub

Public Function IsNotEqual(Rect1 As Rect, Rect2 As Rect) As Boolean
'This compares two Rect structures and returns False if they both
'are equal. Otherwise the function returns True

IsEqual = False

If (Rect1.Left <> Rect2.Left) Or (Rect1.Bottom <> Rect2.Bottom) Or (Rect1.Top <> Rect2.Top) Or (Rect1.Right <> Rect2.Right) Then
    IsEqual = True
End If
End Function

Public Sub EnterText(Alltext As String)
'This procedure will take a string and parse each line into
'a seperate element of an array of strings. This then can be
'used to format each line of text in the picture box for
'scrolling.

Dim Counter As Integer
Dim LastLinePos As Integer
Dim NextLine As String * 2

NextLine = Chr(13) & Chr(10)

'reset array
TextLen = 0
ReDim TextInfo(TextLen)

For Counter = 0 To (Len(Alltext) - 1)
    If Mid(Alltext, Counter + 1, 2) = NextLine Then
        'add a new element to the array
        ReDim Preserve TextInfo(TextLen)
        
        TextInfo(TextLen) = Mid(Alltext, LastLinePos + 1, (Counter - LastLinePos))
        If Left(TextInfo(TextLen), 2) = NextLine Then
            'remove the nextline characters
            TextInfo(TextLen) = Right(TextInfo(TextLen), Len(TextInfo(TextLen)) - 2)
        End If
        TextLen = TextLen + 1
        
        LastLinePos = Counter '- 1
    End If
Next Counter

'this was incremented when an array element was last added, but the
'number currently in this variable is out of sync by one.
If TextLen > 0 Then
    TextLen = TextLen - 1
End If
End Sub

