VERSION 5.00
Begin VB.Form frmMsgBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMsgBox.frx":0000
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmMsgBox.frx":000C
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox UserInputTxt 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I need to know the dimensions of the user's usable screen ... so
'declare the SystemParametersInfo api
'--> Min reqmts: NT 3.1 or +, Win95 or + [used by all msgbox/input boxes]
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48 'Desktop Area with task bar consideration.
'Type structure to hold results of query
Private Type User_Vis_Screen_Rect
    Left As Long
    Top As Long
    Right As Long   'Width = Right - Left
    Bottom As Long  'Height = Bottom - Top
End Type
'assign the type
Private ScreenDimensions As User_Vis_Screen_Rect 'used to keep the actual screen size results (in pixels)
    Dim GetScreenData As Long   ' API call requires a long number
    Dim StoreDimensions As User_Vis_Screen_Rect  'A good place to store results for later work

'declare the GetSystemMetrics api and declare needed constants
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'--> Min reqmts: NT 3.1 or +, Win95 or +
'--> constants needed for our purposes(?)
    Const SM_CXSCREEN = 0 'X Size of screen in pixels
    Const SM_CYSCREEN = 1 'Y Size of Screen in pixels

'declare the api to draw the form border
'--> Min reqmts: NT 3.1 or +, Win95 or +
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'declare api to round the form and clip for transparency
'--> Min reqmts: NT 3.1(1) or +, Win95 or +
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
'--> Min reqmts: NT 3.1 or +, Win95 or +
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, ByVal RectY2 As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'declare api for draging the form around (mousedown anywhere on form ... except buttons)
'--> Min reqmts: NT 3.1 or +, Win95 or +
Private Declare Function ReleaseCapture Lib "user32" () As Long
'--> Min reqmts: NT 3.1 or +, Win95 or +
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'declare api for waiting without using a timer ... used by ShowTimedMessageBox box only
'--> Min reqmts: NT 3.1 or +, Win95 or +
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'declare api for playing system sounds
'--> Min reqmts: NT 3.1 or +, Win95 or +
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
'--> constants needed for our purposes(machine registered sounds)
    Private Const MB_IconAsterisk As Long = &H10&
    Private Const MB_IconQuestion As Long = &H20&
    Private Const MB_IconExclamation As Long = &H30&
    Private Const MB_IconInformation As Long = &H40&

'declare api for getting/printing system icons
'--> Min reqmts: NT 3.1 or +, Win95 or +
Private Declare Function DuplicateIcon Lib "shell32.dll" (ByVal hInst As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'--> constants needed for our purposes(machine registered icons) & misc
    Private Const IDI_APPLICATION = 32512&
    Private Const IDI_ASTERISK = 32516&
    Private Const IDI_EXCLAMATION = 32515&
    Private Const IDI_HAND = 32513&
    Private Const IDI_ERROR = IDI_HAND
    Private Const IDI_INFORMATION = IDI_ASTERISK
    Private Const IDI_QUESTION = 32514&
    Private Const IDI_WARNING = IDI_EXCLAMATION
    Private Const IDI_WINLOGO = 32517
    Private Const DI_MASK = &H1
    Private Const DI_IMAGE = &H2
    Private Const DI_NORMAL = DI_MASK Or DI_IMAGE

'misc declarations needed
Dim NumToUnload As Integer, ColorSchemeUsed As Integer, BtnClicked As Integer
Dim BtnRepaint As Integer, OverButton As Integer, BoxMode As Integer
Dim Draggable As Boolean, RetryOK As Boolean
Dim BtnSpecs() As Integer, Tallest As Integer, Longest As Integer
Dim ButtonText(3) As String ' contains text to reprint
Dim BtnTextData(3, 1) As Integer 'needed to center later
Dim BtnXY(3, 3) As Integer 'stores button outline location as X1,Y1,X2,Y2
Dim SColorScheme(3, 2)
Dim TitleColor As Long, PromptColor As Long, BtnColorReg As Long, BtnColorMO As Long
Dim MyBtnFont As String, MyBtnFontSize As Integer, MyBtnFontBold As Boolean, MyBtnFontItalic As Boolean

Private Function GetColor(ColorString As String) As Long
    GetColor = RGB(Val(Mid(ColorString, 1, 3)), Val(Mid(ColorString, 4, 3)), Val(Mid(ColorString, 7, 3)))
End Function

Private Sub DoIconThing(IconNum As Integer)
    Dim hIcon As Long, hDuplIcon As Long, WhichIcon As Long
    Select Case IconNum
        'Case 1
            'see the readme.txt file for explanation of why this remains here if the following is not clear
            'use icon from this form... Useful for about boxes... not much else
            'temporarily store icon in picturebox
            'frmMsgBox.TempPic.Picture = Me.Icon
            'coerce to a 32x32 icon
            'Me.PaintPicture TempPic.Picture, 6, 6, 32, 32
            'get outta here
            'Exit Sub
        Case 1
            WhichIcon = IDI_APPLICATION
        Case 2
            WhichIcon = IDI_EXCLAMATION
        Case 3
            WhichIcon = IDI_ERROR
        Case 4
            WhichIcon = IDI_INFORMATION
        Case 5
            WhichIcon = IDI_QUESTION
        Case 6
            WhichIcon = IDI_WINLOGO
    End Select
    ' Open the icon
    hIcon = LoadIcon(ByVal 0&, WhichIcon)
    ' Duplicate the icon
    hDuplIcon = DuplicateIcon(ByVal 0&, hIcon)
    ' Draw the result on the form & try to coerce to 32x32
    DrawIconEx Me.hdc, 6, 6, hDuplIcon, 32, 32, 0, 0, DI_NORMAL
    ' Destroy handles
    DestroyIcon hIcon
    DestroyIcon hDuplIcon

End Sub
Private Sub PrintTitle(MsgText As String, MsgFont As String, MsgFontSize As Integer, MsgFontBold As Boolean, MsgFontItalic As Boolean, XOffSet As Integer, YOffSet As Integer)
    With Me
        .Font = MsgFont
        .FontSize = MsgFontSize
        .FontBold = MsgFontBold
        .FontItalic = MsgFontItalic
        .ForeColor = TitleColor
    End With
    'TitleLines will allow text to be left-justified if the title has been wordwrapped earlier
    TitleLines = Split(MsgText, vbCrLf) 'get total lines of text to show with vbcrlf removed
    Me.CurrentY = 5
    Me.CurrentX = 8 + XOffSet
    For i = 0 To UBound(TitleLines)
        Me.Print TitleLines(i) 'print text that will fit on form
        Me.CurrentX = 8 + XOffSet
    Next
End Sub
Private Sub PrintPrompt(MsgText As String, MsgFont As String, MsgFontSize As Integer, MsgFontBold As Boolean, MsgFontItalic As Boolean, TopOffSet As Integer)
    With Me
        .Font = MsgFont
        .FontSize = MsgFontSize
        .FontBold = MsgFontBold
        .FontItalic = MsgFontItalic
        .ForeColor = PromptColor
    End With
    Me.CurrentY = TopOffSet
    'PromptLines will allow text to be left-justified if the Prompt has been wordwrapped earlier
    PromptLines = Split(MsgText, vbCrLf) 'get total lines of text to show with vbcrlf removed
    Me.CurrentX = 8
    For i = 0 To UBound(PromptLines)
        Me.Print PromptLines(i) 'print text that will fit on form
        Me.CurrentX = 8 'reset for justification purposes
    Next
End Sub
Private Function SplitText(MsgText As String, MsgFont As String, MsgFontSize As Integer, MsgFontBold As Boolean, MsgFontItalic As Boolean, OffSet As Integer) As String
    'set up to reflect font changes
    With Me
        .Font = MsgFont
        .FontSize = MsgFontSize
        .FontBold = MsgFontBold
        .FontItalic = MsgFontItalic
    End With
    Dim AllLines, NewWords, NewText()
    Dim Msg: Msg = ""
    AllLines = Split(MsgText, vbCrLf) 'get total lines of text to parse with vbcrlf removed
    For i = 0 To UBound(AllLines)
        If Me.TextWidth(AllLines(i)) > Me.ScaleWidth - OffSet Then
            'resplit till fits
            NewWords = Split(AllLines(i), " ")
                For X = 0 To UBound(NewWords)
                    If Me.TextWidth(TempMsg & " " & NewWords(X)) < Me.ScaleWidth - OffSet Then
                        TempMsg = TempMsg & NewWords(X) & " "
                    Else
                        Msg = Msg & TempMsg & vbCrLf
                        TempMsg = NewWords(X) & " "
                    End If
                Next X
                Msg = Msg & TempMsg & vbCrLf
        Else
            'NewText(i) = AllLines(i) & vbCrLf
            Msg = Msg & AllLines(i) & vbCrLf
        End If
    Next i
    'return new string
    If Right(Msg, 2) = vbCrLf Then
        Msg = Mid(Msg, 1, Len(Msg) - 2)
    End If
    SplitText = Msg
End Function
Private Function GetTextWidth(TText As String, _
                               TFont As String, _
                               TFontSize As Integer, _
                               TFontBold As Boolean, _
                               TFontItalic As Boolean) As Integer
    'return the width of the Text
    With Me
        .Font = TFont
        .FontSize = TFontSize
        .FontBold = TFontBold
        .FontItalic = TFontItalic
    End With
    GetTextWidth = Me.TextWidth(TText)
End Function
Private Function GetTextHeight(TText As String, _
                               TFont As String, _
                               TFontSize As Integer, _
                               TFontBold As Boolean, _
                               TFontItalic As Boolean) As Integer
    'return the height of the Text
    With Me
        .Font = TFont
        .FontSize = TFontSize
        .FontBold = TFontBold
        .FontItalic = TFontItalic
    End With
    GetTextHeight = Me.TextHeight(TText)
End Function


Private Sub GetScreenInfo()
    
    'Call the SystemParametersInfo API
    GetScreenData = SystemParametersInfo(SPI_GETWORKAREA, vbNull, StoreDimensions, 0)
    'was that good for you too?
    If GetScreenData Then
        'the API call was successful... returns dimensions in pixel terms
        ScreenDimensions.Left = StoreDimensions.Left
        ScreenDimensions.Right = StoreDimensions.Right
        ScreenDimensions.Top = StoreDimensions.Top
        ScreenDimensions.Bottom = StoreDimensions.Bottom
        'note: on my 800x600 monitor w/ taskbar at bottom...
        'ScreenDimensions.Left = 0
        'ScreenDimensions.Right = 800
        'ScreenDimensions.Top = 0
        'ScreenDimensions.Bottom = 572
        'therefore:
        'Total Available Width = ScreenDimensions.Right - ScreenDimensions.Left
        'Total Available Height = ScreenDimensions.Bottom - ScreenDimensions.Top
    Else
        'API call failed
        'try less sophisticated way
        ScreenDimensions.Left = 0
        ScreenDimensions.Right = Int(Screen.Width / Screen.TwipsPerPixelX) 'total screen width in pixels
        ScreenDimensions.Top = 0
        ScreenDimensions.Bottom = Int(Screen.Height / Screen.TwipsPerPixelY)
    End If
End Sub
Private Sub PlayAnnoyingSound(SoundNumber As Integer)
    '--> This sub is available to all msg/input boxes <--
    '... all sounds may not be registered in your user's computer... if a sound is not...
    '... this sub degrades seamlessly (no sound is played)
    Select Case SoundNumber
        Case 0
            'this is the same as beep... but I left it here anyway
            MessageBeep 0
        Case 1
            MessageBeep MB_IconAsterisk
        Case 2
            MessageBeep MB_IconQuestion
        Case 3
            MessageBeep MB_IconExclamation
        Case 4
            MessageBeep MB_IconInformation
    End Select
End Sub
Private Sub DrawGradient(ColorScheme As Integer, Top2Bot As Boolean, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    If Top2Bot Then
        RedStrt = 255: GrnStrt = 255: BluStrt = 255
        Red = SColorScheme(ColorScheme, 0): Grn = SColorScheme(ColorScheme, 1): Blu = SColorScheme(ColorScheme, 2)
    Else
        'swap values
        RedStrt = SColorScheme(ColorScheme, 0): GrnStrt = SColorScheme(ColorScheme, 1): BluStrt = SColorScheme(ColorScheme, 2)
        Red = 255: Grn = 255: Blu = 255
    End If
    Trips = Y2 - Y1
    If Trips < 1 Then Exit Sub 'prevent error... skip the gradient
    Dist = (Y2 - Y1) / 255
    For Y = 0 To Trips
        UseRed = (RedStrt / 255) * (255 - (Y / Dist)) + (Red / 255) * (Y / Dist)
        UseGreen = (GrnStrt / 255) * (255 - (Y / Dist)) + (Grn / 255) * (Y / Dist)
        UseBlue = (BluStrt / 255) * (255 - (Y / Dist)) + (Blu / 255) * (Y / Dist)
        Line (X1, Y1 + Y)-(X2, Y1 + Y), RGB(UseRed, UseGreen, UseBlue)
    Next
    
End Sub
Private Sub DoFormStuff(ColorScheme As Integer, TitleHeight As Integer, RepeatGradient As Boolean)
    '--> This sub is used by all msg/input boxes <--
    'Draw a black box around the virtual title bar
    Me.Line (0, 0)-(Me.ScaleWidth, TitleHeight + 8), &H0&, B
    '--> color scheme gradient in RGB format in virtual title bar
    '... this is matched (sort of) to the graphics provided... if you change those...
    '... change values below also.  The scheme is also used to choose the border color
    'the following is passed to the dimensioned array... so it can be used for buttons later
    SColorScheme(0, 0) = 255: SColorScheme(0, 1) = 255: SColorScheme(0, 2) = 255 ' white
    SColorScheme(1, 0) = 214: SColorScheme(1, 1) = 214: SColorScheme(1, 2) = 214 'silver
    SColorScheme(2, 0) = 173: SColorScheme(2, 1) = 198: SColorScheme(2, 2) = 239 'blue
    SColorScheme(3, 0) = 206: SColorScheme(3, 1) = 214: SColorScheme(3, 2) = 189 'olive
    'draw the title gradient
    DrawGradient ColorScheme, True, 0, 0, Me.ScaleWidth - 1, TitleHeight + 7
    If RepeatGradient Then
        DrawGradient ColorScheme, True, 1, TitleHeight + 9, Me.ScaleWidth - 1, Me.ScaleHeight - 1
    End If
    '--> Draw form border
    'Draw the form border according to the colorscheme
    Me.ForeColor = RGB(SColorScheme(ColorScheme, 0), SColorScheme(ColorScheme, 1), SColorScheme(ColorScheme, 2))
    'correct if it's white/white
    If Me.ForeColor = RGB(255, 255, 255) Then
        Me.ForeColor = RGB(192, 192, 192) ' change to a gray
    End If
    RoundRect Me.hdc, 0, 0, (Me.Width / Screen.TwipsPerPixelX) - 1, (Me.Height / Screen.TwipsPerPixelY) - 1, CLng(25), CLng(25)
    'if you want to get a more 3d look... change the Me.Forecolor again here... but this looks ok to me.
    RoundRect Me.hdc, 1, 1, (Me.Width / Screen.TwipsPerPixelX) - 2, (Me.Height / Screen.TwipsPerPixelY) - 2, CLng(25), CLng(25)
    '-->clip rounded corners transparent
    SetWindowRgn Me.hWnd, CreateRoundRectRgn(0, 0, (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY), 25, 25), True
End Sub
Private Function GetBoxLeft(ByRef BoxLeft As Integer) As Integer
    '--> This function is used by all msg/input boxes <--
    'Returns an integer used to locate the form horizontally
    'the SystemParametersInfo API has already been called to set the max X and Y
    'values of the form... so we don't need to call again since we stored that info in
    'User_Vis_Screen_Rect earlier
    Dim AvailableHSpace
    'Establish usable screen space
    AvailableHSpace = ScreenDimensions.Right - ScreenDimensions.Left
    'now locate the form somewhere depending on BoxLeft
    If BoxLeft = -1 Then 'center the form horizontally
        GetBoxLeft = Int(((AvailableHSpace * Screen.TwipsPerPixelX) - Me.Width)) / 2 + (ScreenDimensions.Left * Screen.TwipsPerPixelX) 'add in the offset on left... if any exists
        Exit Function
    End If
    'handle BoxLeft values > -1
    Dim MaxX
    'make sure won't leak off if toolbar is on left side
    If BoxLeft < ScreenDimensions.Left Then BoxLeft = ScreenDimensions.Left
    
    'check to make sure form won't leak off right screen edge
    MaxX = ScreenDimensions.Right - Me.ScaleWidth
    'if it does... move it back to the left
    If BoxLeft > MaxX Then BoxLeft = MaxX
    GetBoxLeft = BoxLeft * Screen.TwipsPerPixelX
End Function
Private Function GetBoxTop(ByRef BoxTop As Integer) As Integer
    '--> This function is used by all msg/input boxes <--
    'Returns an integer used in locating the form vertically
    'the SystemParametersInfo API has already been called to set the max X and Y
    'values of the form... so we don't need to call again since we stored that info in
    'User_Vis_Screen_Rect earlier
    
    'Establish usable screen space
    AvailableVSpace = ScreenDimensions.Bottom - ScreenDimensions.Top
    'calculate form top
    If BoxTop = -1 Then 'center the form vertically
        GetBoxTop = Int(((AvailableVSpace * Screen.TwipsPerPixelY) - Me.Height)) / 2 + (ScreenDimensions.Top = 0 * Screen.TwipsPerPixelY) 'add in the top offset... if any exists
        Exit Function
    End If
    Dim MaxY
    'check to make sure form won't leak off top screen edge if toolbar is on top
    If BoxTop < ScreenDimensions.Top Then BoxTop = ScreenDimensions.Top
    'check to make sure form won't leak off bottom screen edge
    MaxY = ScreenDimensions.Bottom - Me.ScaleHeight
    'if it does... move it back up
    If BoxTop > MaxY Then BoxTop = MaxY
    GetBoxTop = BoxTop * Screen.TwipsPerPixelY
End Function
Private Sub GetBtnSpecs(Btn1Text As String, Btn2Text As String, Btn3Text As String, BtnFont As String, BtnFontSize As Integer, BtnFontBold As Boolean, BtnFontItalic As Boolean)
    'this sub calculates the height and width of the text that appears in the buttons
    'and the default X for the top right close button
    'close button and pass the variables for later use
    ButtonText(0) = Trim(Btn1Text): ButtonText(1) = Trim(Btn2Text): ButtonText(2) = Trim(Btn3Text): ButtonText(3) = "X"
    'this could be called more than once... so clear the array
    ReDim BtnSpecs(3, 3)
    'retrieve the width/height of the button text' used temporarily... reset later
    BtnSpecs(0, 2) = GetTextHeight(Trim(Btn1Text), BtnFont, BtnFontSize, BtnFontBold, BtnFontItalic)
    BtnSpecs(1, 2) = GetTextHeight(Trim(Btn2Text), BtnFont, BtnFontSize, BtnFontBold, BtnFontItalic)
    BtnSpecs(2, 2) = GetTextHeight(Trim(Btn3Text), BtnFont, BtnFontSize, BtnFontBold, BtnFontItalic)
    'BtnSpecs(3, 2) = GetTextHeight(ButtonText(3), "Arial", 10, True, False)
    
    BtnSpecs(0, 3) = GetTextWidth(Trim(Btn1Text), BtnFont, BtnFontSize, BtnFontBold, BtnFontItalic)
    BtnSpecs(1, 3) = GetTextWidth(Trim(Btn2Text), BtnFont, BtnFontSize, BtnFontBold, BtnFontItalic)
    BtnSpecs(2, 3) = GetTextWidth(Trim(Btn3Text), BtnFont, BtnFontSize, BtnFontBold, BtnFontItalic)
    'BtnSpecs(3, 3) = GetTextWidth(ButtonText(3), "Arial", 10, True, False)
End Sub
Private Function CalcHorizSpacing(MaxMsgboxWidth As Integer, Longest As Integer, BtnsNeeded As Integer) As Integer
    'this returns an integer representing the offset needed to center lower buttons horizontally
    'calculates the button border spacing... NOT the text spacing!
    CalcHorizSpacing = Int(((MaxMsgboxWidth - 0) - ((Longest + 8) * BtnsNeeded)) / (BtnsNeeded + 1))
End Function
Private Sub DrawBtnRectangle(StartLeft As Integer, Tallest As Integer, Longest As Integer, Index As Integer)
    'draw the button outline for the lower buttons
    Me.ForeColor = &H0&
    Me.CurrentY = UserInputTxt.Top + UserInputTxt.Height + 5
    RoundRect Me.hdc, StartLeft, Me.CurrentY, StartLeft + Longest + 8, Me.CurrentY + Tallest + 8, CLng(6), CLng(6)
    'save this for gradient,text positioning and button mapping
    BtnSpecs(Index, 0) = StartLeft 'x position
    BtnSpecs(Index, 1) = Me.CurrentY 'y position
    BtnSpecs(Index, 2) = Tallest + 8 'height
    BtnSpecs(Index, 3) = Longest + 8 'width
    'Dim BtnXY(3, 3) As Integer 'stores button outline location as X1,Y1,X2,Y2
    BtnXY(Index, 0) = StartLeft 'x1 position
    BtnXY(Index, 1) = Me.CurrentY 'y1 position
    BtnXY(Index, 2) = StartLeft + Longest + 8 'x2 position
    BtnXY(Index, 3) = Me.CurrentY + Tallest + 8 'y2 position
End Sub
Private Sub SetToNoButton(Index As Integer)
    'this sub sets/resets the BtnSpecs array to zero... which maps it's x,y coords to 0,0
    'which is never visible to the user since 0,0 is in the area clipped by the round window
    BtnSpecs(Index, 0) = 0 ' set x = 0
    BtnSpecs(Index, 1) = 0 ' set y = 0
    BtnSpecs(Index, 2) = 0 ' set height = 0
    BtnSpecs(Index, 3) = 0 ' set width = 0
End Sub
Private Sub PrintLowerBtnText(BtnFont As String, BtnFontSize As Integer, BtnFontBold As Boolean, BtnFontItalic As Boolean)
    'this sub prints the text for the lower buttons and stores the needed info for rollover effects
    'initialize font related stuff for the form
    With Me
        .Font = BtnFont
        .FontSize = BtnFontSize
        .FontBold = BtnFontBold
        .FontItalic = BtnFontItalic
        .ForeColor = BtnColorReg
    End With
    'center text vertically... 1 time... all buttons are in a single row
    Me.CurrentY = UserInputTxt.Top + UserInputTxt.Height + 5 + Int((BtnSpecs(0, 2) - BtnTextData(0, 1)) / 2)
    'center text horizontally... every time
    Me.CurrentX = BtnSpecs(0, 0) + Int((BtnSpecs(0, 3) - BtnTextData(0, 0)) / 2)
        BtnTextData(0, 0) = Me.CurrentX 'reuse the BtnTextData array to keep track of where to print rollover updates
        BtnTextData(0, 1) = Me.CurrentY
    'center text horizontally... every time
    Me.Print Trim(ButtonText(0));
    Me.CurrentX = BtnSpecs(1, 0) + Int((BtnSpecs(1, 3) - BtnTextData(1, 0)) / 2)
        BtnTextData(1, 0) = Me.CurrentX 'reuse the BtnTextData array to keep track of where to print rollover updates
        BtnTextData(1, 1) = Me.CurrentY
    Me.Print Trim(ButtonText(1));
    Me.CurrentX = BtnSpecs(2, 0) + Int((BtnSpecs(2, 3) - BtnTextData(2, 0)) / 2)
        BtnTextData(2, 0) = Me.CurrentX 'reuse the BtnTextData array to keep track of where to print rollover updates
        BtnTextData(2, 1) = Me.CurrentY
    Me.Print Trim(ButtonText(2))
    
End Sub
Private Sub MoveCloseButton(THeight As Integer)
    'initialize font related stuff for the form
    With Me
        .Font = "Arial"
        .FontSize = 8
        .FontBold = True
        .FontItalic = False
        .ForeColor = BtnColorReg
    End With
    'how wide/tall is the X?
    XWidth = GetTextWidth("X", Me.Font, Me.FontSize, True, False)
    XHeight = GetTextHeight("X", Me.Font, Me.FontSize, True, False)
    BtnTextData(3, 0) = Me.ScaleWidth - XWidth - 12 'the X's X position
    BtnTextData(3, 1) = Int((THeight - XHeight) / 2) + 5 'the X's Y position
    'BtnXY(3, 3) stores button outline location as X1,Y1,X2,Y2
    BtnXY(3, 0) = BtnTextData(3, 0) - 4 'x1 position of rectangle
    BtnXY(3, 1) = BtnTextData(3, 1) - 2  'y1 position
    BtnXY(3, 2) = BtnXY(3, 0) + XWidth + 8 'x2 position
    BtnXY(3, 3) = BtnXY(3, 1) + XHeight + 4 'y2 position
    'draw outline rectangle
    Me.ForeColor = &H0&
    RoundRect Me.hdc, BtnXY(3, 0), BtnXY(3, 1), BtnXY(3, 2), BtnXY(3, 3), CLng(6), CLng(6)
    'draw top half gradient for close btn
    DrawGradient (ColorSchemeUsed), False, BtnXY(3, 0) + 2, BtnXY(3, 1) + 1, BtnXY(3, 2) - 2, BtnXY(3, 1) + Int((BtnXY(3, 3) - BtnXY(3, 1)) / 2)
    'reverse colors and draw bottom half of gradient
    DrawGradient (ColorSchemeUsed), True, BtnXY(3, 0) + 2, BtnXY(3, 1) + Int((BtnXY(3, 3) - BtnXY(3, 1)) / 2) + 1, BtnXY(3, 2) - 2, BtnXY(3, 3) - 2
    Me.CurrentX = BtnTextData(3, 0)
    Me.CurrentY = BtnTextData(3, 1)
    'print the x
    Me.ForeColor = BtnColorReg
    ButtonText(3) = "X"
    Me.Print ButtonText(3)
End Sub
Private Sub SetBtnSpecs(HorizSpacing As Integer)
    If BtnSpecs(0, 3) Then
        BtnSpecs(0, 0) = HorizSpacing ' set x for The left button box left
        DrawBtnRectangle BtnSpecs(0, 0), Tallest, Longest, 0 'draw btn border
    Else
        SetToNoButton 0 ' shouldn't be possible for input box
    End If
    
    If BtnSpecs(1, 3) Then
        BtnSpecs(1, 0) = BtnSpecs(0, 0) + Longest + 8 + HorizSpacing
        DrawBtnRectangle BtnSpecs(1, 0), Tallest, Longest, 1
    Else
        SetToNoButton 1
    End If
    
    If BtnSpecs(2, 3) Then
        BtnSpecs(2, 0) = BtnSpecs(1, 0) + Longest + 8 + HorizSpacing
        DrawBtnRectangle BtnSpecs(2, 0), Tallest, Longest, 2
    Else
        SetToNoButton 2
    End If
End Sub
Private Sub DrawLowerButtonGradient(ColorScheme As Integer, Index As Integer)
    'this sub draws the lower button gradients before the text is applied
    'otherwise... it would overwrite the button text
    'The SColorScheme(X, Y) values were passed earlier
    'start with darkest... as if lighting to center
    MyR = SColorScheme(ColorScheme, 0): MyG = SColorScheme(ColorScheme, 1): MyB = SColorScheme(ColorScheme, 2) ' start with darkest... as if lighting to center
    'draw top & bottom lines so won't overwrite border... then use gradient for balance
    Me.Line (BtnSpecs(0, 0) + 2, BtnSpecs(0, 1) + 1)-(BtnSpecs(0, 0) + BtnSpecs(0, 3) - 2, BtnSpecs(0, 1) + 1), RGB(MyR, MyG, MyB)
    Me.Line (BtnSpecs(0, 0) + 2, BtnSpecs(0, 1) + BtnSpecs(0, 2) - 2)-(BtnSpecs(0, 0) + BtnSpecs(0, 3) - 2, BtnSpecs(0, 1) + BtnSpecs(0, 2) - 2), RGB(MyR, MyG, MyB)
    Me.Line (BtnSpecs(1, 0) + 2, BtnSpecs(1, 1) + 1)-(BtnSpecs(1, 0) + BtnSpecs(1, 3) - 2, BtnSpecs(1, 1) + 1), RGB(MyR, MyG, MyB)
    Me.Line (BtnSpecs(1, 0) + 2, BtnSpecs(1, 1) + BtnSpecs(1, 2) - 2)-(BtnSpecs(1, 0) + BtnSpecs(1, 3) - 2, BtnSpecs(1, 1) + BtnSpecs(1, 2) - 2), RGB(MyR, MyG, MyB)
    Me.Line (BtnSpecs(2, 0) + 2, BtnSpecs(2, 1) + 1)-(BtnSpecs(2, 0) + BtnSpecs(2, 3) - 2, BtnSpecs(2, 1) + 1), RGB(MyR, MyG, MyB)
    Me.Line (BtnSpecs(2, 0) + 2, BtnSpecs(2, 1) + BtnSpecs(2, 2) - 2)-(BtnSpecs(2, 0) + BtnSpecs(2, 3) - 2, BtnSpecs(2, 1) + BtnSpecs(2, 2) - 2), RGB(MyR, MyG, MyB)
    
    DrawGradient ColorScheme, False, BtnSpecs(0, 0) + 1, BtnSpecs(0, 1) + 2, BtnSpecs(0, 0) + BtnSpecs(0, 3) - 1, BtnSpecs(0, 1) + (Int(Tallest / 2) + 4)
    DrawGradient ColorScheme, True, BtnSpecs(0, 0) + 1, BtnSpecs(0, 1) + (Int(Tallest / 2) + 5), BtnSpecs(0, 0) + BtnSpecs(0, 3) - 1, BtnSpecs(0, 1) + BtnSpecs(0, 2) - 3
    
    DrawGradient ColorScheme, False, BtnSpecs(1, 0) + 1, BtnSpecs(1, 1) + 2, BtnSpecs(1, 0) + BtnSpecs(1, 3) - 1, BtnSpecs(1, 1) + (Int(Tallest / 2) + 4)
    DrawGradient ColorScheme, True, BtnSpecs(1, 0) + 1, BtnSpecs(1, 1) + (Int(Tallest / 2) + 5), BtnSpecs(1, 0) + BtnSpecs(1, 3) - 1, BtnSpecs(1, 1) + BtnSpecs(1, 2) - 3
    
    DrawGradient ColorScheme, False, BtnSpecs(2, 0) + 1, BtnSpecs(2, 1) + 2, BtnSpecs(2, 0) + BtnSpecs(2, 3) - 1, BtnSpecs(2, 1) + (Int(Tallest / 2) + 4)
    DrawGradient ColorScheme, True, BtnSpecs(2, 0) + 1, BtnSpecs(2, 1) + (Int(Tallest / 2) + 5), BtnSpecs(2, 0) + BtnSpecs(2, 3) - 1, BtnSpecs(2, 1) + BtnSpecs(2, 2) - 3
End Sub
Private Sub DoScreenLimits(BoxLeft As Integer, BoxTop As Integer)
    'this sub stops form from ending up in unruly places... and sets .left/.top for the form
    Dim UsersScreenWidth As Integer
    Dim UsersScreenHeight As Integer
        UsersScreenWidth = GetSystemMetrics(SM_CXSCREEN) 'returns pixel width of the user screen
        UsersScreenHeight = GetSystemMetrics(SM_CYSCREEN) 'returns pixel height of the user screen
        'Limit accidental passed variable values to screen size
        If BoxLeft < -1 Then BoxLeft = 0
        If BoxLeft > UsersScreenWidth Then BoxLeft = UsersScreenWidth
        If BoxTop < -1 Then BoxTop = 0
        If BoxTop > UsersScreenHeight Then BoxTop = UsersScreenHeight
            'set left position in screen pixels
            Me.Left = GetBoxLeft((BoxLeft))
            'set top position in screen pixels
            Me.Top = GetBoxTop((BoxTop))
End Sub
Private Sub SetBtnsToNaught(BoxMode As Integer)
    
End Sub

Public Function ShowMsgBox(MsgTitle As String, MsgPrompt As String, Optional MsgMode As Integer = 1, _
                             Optional ByVal MsgTitleFontColor As String = "000000000", Optional ByVal MsgPromptFontColor As String = "000000000", _
                             Optional ByVal ColorScheme As Integer = 0, Optional ByVal RepeatGradient As Boolean = False, _
                             Optional ByVal ShowIconNum As Integer = 0, _
                             Optional ByVal BoxLeft As Integer = -1, Optional ByVal BoxTop As Integer = -1, _
                             Optional ByVal DefaultAnswer As String = vbNullString, _
                             Optional ByVal Btn1Text As String = "OK", Optional ByVal Btn2Text As String = vbNullString, Optional ByVal Btn3Text As String = vbNullString, _
                             Optional ByVal Dragit As Boolean = True, _
                             Optional ByVal PlaySound As Integer = -1, _
                             Optional ByVal MsgTitleFont As String = "Arial", Optional ByVal MsgTitleFontSize As Integer = 8, Optional ByVal MsgTitleFontBold As Boolean = False, Optional ByVal MsgTitleFontItalic As Boolean = False, _
                             Optional ByVal MsgPromptFont As String = "Arial", Optional ByVal MsgPromptFontSize As Integer = 8, Optional ByVal MsgPromptFontBold As Boolean = False, Optional ByVal MsgPromptFontItalic As Boolean = False, _
                             Optional ByVal BtnFontColorReg As String = "000000000", Optional ByVal BtnFontColorMOver As String = "255000000", _
                             Optional ByVal BtnFont As String = "Arial", _
                             Optional ByVal BtnFontSize As Integer = 8, _
                             Optional ByVal BtnFontBold As Boolean = False, _
                             Optional ByVal BtnFontItalic As Boolean = False, Optional ByVal TimeDelay As Integer = 5) As String
    
    BoxMode = MsgMode '0 = input box, 1=msgbox, 2 turns off lower buttons, 3 turns off all buttons
    'play an annoying sound if you must!
    If PlaySound > -1 Then
        PlayAnnoyingSound ((PlaySound))
    End If
    RetryOK = True
    Draggable = Dragit 'enable/disable dragging
    'pass colorscheme value so can be used by other subs
    If ColorScheme > 3 Or ColorScheme < 0 Then ColorScheme = 0 'limit to available schemes
    ColorSchemeUsed = ColorScheme
    'check font colors for mistakes... format is [Red 255][Green 255][Blue 255] in a string as "255255255"
    If Len(MsgTitleFontColor) <> 9 Or IsNumeric(MsgTitleFontColor) = False Then MsgTitleFontColor = "000000000" 'return it to black
    If Len(MsgPromptFontColor) <> 9 Or IsNumeric(MsgPromptFontColor) = False Then MsgPromptFontColor = "000000000"  'return it to black
    If Len(BtnFontColorReg) <> 9 Or IsNumeric(BtnFontColorReg) = False Then MsgTitleFontColor = "000000000"  'return it to black
    If Len(BtnFontColorMOver) <> 9 Or IsNumeric(BtnFontColorMOver) = False Then MsgPromptFontColor = "255000000"  'return it to red
    If ShowIconNum > 6 Then ShowIconNum = 0 'limit icon possibilities
    'convert string colors to long values now
    TitleColor = GetColor(MsgTitleFontColor)
    PromptColor = GetColor(MsgPromptFontColor)
    BtnColorReg = GetColor(BtnFontColorReg)
    BtnColorMO = GetColor(BtnFontColorMOver)
    'pass font defaults
    MyBtnFont = BtnFont: MyBtnFontSize = BtnFontSize: MyBtnFontBold = BtnFontBold: MyBtnFontItalic = BtnFontItalic
    UserInputTxt.Text = DefaultAnswer 'set the default answer
    'get usable screen dimensions
    GetScreenInfo
    'get width of Title Text
    TWidth = GetTextWidth(MsgTitle, MsgTitleFont, MsgTitleFontSize, MsgTitleFontBold, MsgTitleFontItalic)
    'get the total width of the prompt text
    PWidth = GetTextWidth(MsgPrompt, MsgPromptFont, MsgPromptFontSize, MsgPromptFontBold, MsgPromptFontItalic)
    'set the maximum to 90% of height and width of users screen... sort of arbitrary I guess
    Dim MaxMsgboxWidth As Integer, MaxMsgboxHeight As Integer, SaveWidth As Integer, SaveHeight As Integer
    MaxMsgboxWidth = Int((ScreenDimensions.Right - ScreenDimensions.Left) * 0.9)
        SaveWidth = MaxMsgboxWidth
    MaxMsgboxHeight = Int((ScreenDimensions.Bottom - ScreenDimensions.Top) * 0.9)
        SaveHeight = MaxMsgboxHeight
    'start with smallest width possible
    If BoxMode = 3 Then
        XW = 68
    Else
        XW = 88
    End If
    If PWidth >= TWidth + 88 Then
        TargetWidth = PWidth + 16
    Else
        TargetWidth = TWidth + XW '<---- needed to keep title from writing over close X
    End If

    Select Case MaxMsgboxWidth
        Case Is >= TargetWidth
            MaxMsgboxWidth = TargetWidth
        Case Else
            MaxMsgboxWidth = MaxMsgboxWidth
    End Select
ResizeTheStinkinBoxAgain:
    'set the messagebox width
    Me.Width = MaxMsgboxWidth * Screen.TwipsPerPixelX
    'check to make sure really long text won't leak off form... or over button
    '-----------------
    
    If TWidth + 40 > MaxMsgboxWidth Then
        'I need to wordwrap the title... this can funk up stuff if it gets too tall!
        MsgTitle = SplitText(MsgTitle, MsgTitleFont, MsgTitleFontSize, MsgTitleFontBold, MsgTitleFontItalic, 88)
    End If
    'get height of title text
    THeight = GetTextHeight(MsgTitle, MsgTitleFont, MsgTitleFontSize, MsgTitleFontBold, MsgTitleFontItalic)
    If PWidth + 16 > MaxMsgboxWidth Then 'wordwrap the prompt
        MsgPrompt = SplitText(MsgPrompt, MsgPromptFont, MsgPromptFontSize, MsgPromptFontBold, MsgPromptFontItalic, 18)
    End If
    'get height of prompt text
    PHeight = GetTextHeight(MsgPrompt, MsgPromptFont, MsgPromptFontSize, MsgPromptFontBold, MsgPromptFontItalic)
    'skip button sizing if BoxMode = 2 or 3
    'need to initialize the array or out of range
    ReDim BtnSpecs(3, 3)
    If BoxMode > 1 Then GoTo SkipBtns
    '-----------------
    'calc height/width (space) needed for buttons
    'since this function could generate an input box...
    'force atleast an ok button so the user can respond even if it is not used later
ResizeForButtons:
    If Len(Btn1Text) = 0 Then 'some wise guy passed an empty string that defeated the default
        Btn1Text = "OK": ButtonText(0) = "OK"
    End If
    GetBtnSpecs Btn1Text, Btn2Text, Btn3Text, BtnFont, BtnFontSize, BtnFontBold, BtnFontItalic
    'initialize tallest/longest button text values
    Tallest = 0: Longest = 0
    
    For X = 0 To 2
        'sort for largest and save temporarily
        If BtnSpecs(X, 2) > Tallest Then Tallest = BtnSpecs(X, 2)
        If BtnSpecs(X, 3) > Longest Then Longest = BtnSpecs(X, 3)
        BtnTextData(X, 0) = BtnSpecs(X, 3)
        BtnTextData(X, 1) = BtnSpecs(X, 2)
    Next
    'how many bottom buttons are requested
    Dim BtnsNeeded As Integer
    BtnsNeeded = 0
    For X = 0 To 2
        If Len(ButtonText(X)) > 1 Then BtnsNeeded = BtnsNeeded + 1
    Next
    'add space for button borders and offsets on the form
    'so the width needed for the button text is:
    Dim BtnWidthNeeded As Integer
    BtnWidthNeeded = (BtnsNeeded * (Longest + 8)) + 16
    'check BtnWidthNeeded vs Savewidth... does it need resizing?
    If BtnWidthNeeded <= SaveWidth And BtnWidthNeeded > MaxMsgboxWidth Then
        MaxMsgboxWidth = BtnWidthNeeded
        GoTo ResizeTheStinkinBoxAgain 'resize form so they will fit
    Else
        'this is another bad possibility: button text that is wider than the screen...
        'I can't win the resize battle you want to fight with such long text in the buttons
        'so I'll just reset the defaults and forget about it!
        If BtnWidthNeeded > SaveWidth Then
            If Len(Btn1Text) Then
                Btn1Text = "OK"
            End If
            If Len(Btn2Text) Then
                Btn2Text = "Cancel"
            End If
            If Len(Btn3Text) Then
                Btn3Text = "Retry"
            End If
            GoTo ResizeForButtons
        End If
    End If
SkipBtns:
    '==========================
    'calc offset needed for icon if shown
    If ShowIconNum = 0 Then
        TitleOffset = 0
        PromptOffSet = 0
    Else
        TitleOffset = 40
        PromptOffSet = 42
    End If
    'change the offset if title is taller than the Icon offset
    If THeight >= 44 Then
        If ShowIconNum > 0 Then
            InBoxOffSet = PromptOffSet - 32
        End If
    Else
        InBoxOffSet = PromptOffSet
    End If
    If PromptOffSet < THeight + 10 Then PromptOffSet = THeight + 10
    'move the input text box
    'depending on icon situation
    If ShowIconNum = 0 Then
        UserInputTxt.Move 8, (PromptOffSet + PHeight + 13 + InBoxOffSet), Me.ScaleWidth - 15
    Else
        UserInputTxt.Move 8, (PromptOffSet + PHeight + 13), Me.ScaleWidth - 15
    End If
    '--- a cheesy but effective method for moving buttons up the screen if a msgbox
    If BoxMode Then
        UserInputTxt.Top = UserInputTxt.Top - UserInputTxt.Height
        UserInputTxt.Visible = False 'might as well make it invisible now
    End If
    '============
    '--- set the form height
    '... & check to see if this exceeds the max height... if so limit it
    '... this is a bad deal if it is taller than screen... so this should make sure the
    '... top close button is atleast still visible so window can be closed
    Me.Height = (PromptOffSet + PHeight + 10 + UserInputTxt.Height + 10 + Tallest + 15) * Screen.TwipsPerPixelY
    'another cheesy move up
    If BoxMode Then
        If BoxMode = 1 Then
            Me.Height = Me.Height - UserInputTxt.Height * Screen.TwipsPerPixelY
        Else
            Me.Height = Me.Height - ((UserInputTxt.Height + 15) * Screen.TwipsPerPixelY)
        End If
    End If
    'finally, adjust if no lower buttons
    'If BoxMode > 1 Then Me.Height = Me.Height - 35
    If Me.Height > SaveHeight * Screen.TwipsPerPixelY Then
        Me.Height = SaveHeight * Screen.TwipsPerPixelY
    End If
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++ FINALLY!  All of the error checking/re-sizing should be done... ++
    '++ So we can do the real work                                      ++
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'paint and clip the form
    DoFormStuff ColorSchemeUsed, (THeight), RepeatGradient
    'paint the icon if needed
    If ShowIconNum > 0 Then DoIconThing ShowIconNum
    'print the title text
    PrintTitle MsgTitle, MsgTitleFont, MsgTitleFontSize, MsgTitleFontBold, MsgTitleFontItalic, (TitleOffset), (THeight - 19)
    'print the prompt text (adjusted if title text has been wrapped to 2nd (or more) line)
    PrintPrompt MsgPrompt, MsgPromptFont, MsgPromptFontSize, MsgPromptFontBold, MsgPromptFontItalic, (PromptOffSet)
    'calc height/width (space) needed for buttons
    'save this to the BtnSpecs(3, 3) integer array for mapping reasons that may(?) make sense later
    'the array format is: X,Y,Height,Width
    'ie.             Btn0,X,Y,Height,Width
    '................Btn1,X,Y,Height,Width... and so on (Btn 4 is the close X)
    If BoxMode > 1 Then
        'map btns as needed (out of sight so won't respond to mouse over/click events)
        Select Case BoxMode
            Case 2
                For X = 0 To 2
                    SetToNoButton (X)
                Next
            Case Is >= 3
                For X = 0 To 3
                    SetToNoButton (X)
                Next
        End Select
        GoTo NoLowBtns
    End If
    Dim HorizSpacing As Integer
    HorizSpacing = CalcHorizSpacing(MaxMsgboxWidth, Longest, BtnsNeeded)
    'pass info needed to draw lower button rectangles and draw them
    SetBtnSpecs HorizSpacing
    '----- draw lower button gradients
    DrawLowerButtonGradient ColorScheme, 0
    '-----print button text
    PrintLowerBtnText BtnFont, BtnFontSize, BtnFontBold, BtnFontItalic
NoLowBtns:
    '----- locate & Print the X close button?
    If BoxMode <> 3 And BoxMode <> 4 Then
        MoveCloseButton (THeight)
    End If
    '----------------------
    'put the form on the screen somewhere!
    DoScreenLimits BoxLeft, BoxTop
    
    BtnClicked = -1 ' set default response
    If BoxMode < 3 Then
        frmMsgBox.Show vbModal, Form1
    Else
        frmMsgBox.Show
        frmMsgBox.Refresh 'make sure the message box is displayed first
    End If
    'report which button clicked
    Select Case BtnClicked
        Case -1
            'BoxMode = 3 non-modal and no owner!
            'THIS IS IMPORTANT!!!!
            'Eliminating all of the buttons (BoxMode = 3) shows this form non-modally
            'and with no owner... so you can go about other things while this form is
            'displayed... but there are some caveats to using this mode:
            '... 1) You can not show the form again for any purpose before it is unloaded...
            '       if you attempt to do so... you'll get an error
            '... 2) You must REMEMBER to unload the form before your calling program terminates
            '...    or the form will remain on your users screen until the computer is shut down!
            If BoxMode = 3 Then
                GoTo SkipHide
            End If
            If BoxMode = 4 Then
                'timed message box
                frmMsgBox.Show , Form1
                frmMsgBox.Refresh 'make sure the message box is displayed first
                Timerloop = 0 'initialize the delay period to zero
                Do Until Timerloop >= TimeDelay
                    'TimerLbl.Caption = TimerCaption & " " & DelaySeconds - Timerloop & " Seconds"
                    DoEvents
                    Sleep 1000 ' wait 1 second
                    Timerloop = Timerloop + 1
                Loop
                
            End If
        Case 0
            'button in OK position clicked
            If BoxMode = 0 Then ShowMsgBox = UserInputTxt.Text
            If BoxMode = 1 Then ShowMsgBox = "0"
        Case 1
            'button in Cancel position clicked
            If BoxMode = 0 Then ShowMsgBox = vbNullString
            If BoxMode = 1 Then ShowMsgBox = "1"
        Case 2
            'button in Re-try position clicked
            If BoxMode = 0 Then
                'you don't need to handle this as the code is now written...
                'this possibility is trapped in the Form_MouseUp sub
            End If
            If BoxMode = 1 Then ShowMsgBox = "2"
        Case 3
            'the top right X close btn clicked
            If BoxMode = 1 Then ShowMsgBox = "3"
    End Select
    
    Form_Unload 0
SkipHide:

End Function

Private Sub Form_DblClick()
Form_Unload 1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Draggable And OverButton = -1 Then
        'allow form to be moved
        ReleaseCapture
        SendMessage Me.hWnd, &HA1, 2, 0&
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'clear button rollover effects
    Select Case Y
        Case BtnXY(3, 1) To BtnXY(3, 3)
            CheckForCloseBtnMouseOver (X)
        Case BtnXY(0, 1) To BtnXY(0, 3)
            CheckForLowBtnMouseOver (X)
        Case Else
            Me.MousePointer = 0
            If BtnRepaint > -1 Then
                Repaint_If_Needed BtnRepaint
            End If
    End Select
End Sub
Private Sub CheckForLowBtnMouseOver(X As Integer)
    Select Case X
        Case BtnXY(0, 0) To BtnXY(0, 2)
            If BtnRepaint > -1 And BtnRepaint <> 0 Then Repaint_If_Needed BtnRepaint
            OverButton = 0: BtnRepaint = 0
            With Me
                .CurrentX = BtnTextData(0, 0)
                .CurrentY = BtnTextData(0, 1)
            End With
        Case BtnXY(1, 0) To BtnXY(1, 2)
            If BtnRepaint > -1 And BtnRepaint <> 1 Then Repaint_If_Needed BtnRepaint
            OverButton = 1: BtnRepaint = 1
            With Me
                .CurrentX = BtnTextData(1, 0)
                .CurrentY = BtnTextData(1, 1)
            End With
        Case BtnXY(2, 0) To BtnXY(2, 2)
            If BtnRepaint > -1 And BtnRepaint <> 2 Then Repaint_If_Needed BtnRepaint
            OverButton = 2: BtnRepaint = 2
            With Me
                .CurrentX = BtnTextData(2, 0)
                .CurrentY = BtnTextData(2, 1)
            End With
        Case Else
            Me.MousePointer = 0
            If BtnRepaint > -1 Then
                Repaint_If_Needed BtnRepaint
                
            End If
            Exit Sub
        End Select
        With Me
            .MousePointer = 99
            .Font = MyBtnFont
            .FontSize = MyBtnFontSize
            .FontBold = MyBtnFontBold
            .FontItalic = MyBtnFontItalic
            .ForeColor = BtnColorMO
        End With
        If OverButton = -1 Then
            Me.MousePointer = 0
            Exit Sub
        End If
        Me.Print ButtonText(OverButton)
End Sub
Private Sub CheckForCloseBtnMouseOver(X As Single)
    Select Case X
        Case BtnXY(3, 0) To BtnXY(3, 2)
            OverButton = 3: BtnRepaint = 3
            With Me
                .CurrentX = BtnTextData(3, 0) 'the X's X position
                .CurrentY = BtnTextData(3, 1)
                .MousePointer = 99
                .Font = "Arial"
                .FontSize = 8
                .FontBold = True
                .ForeColor = BtnColorMO
                .FontItalic = False
            End With
            Me.Print "X"
        Case Else
            If BtnRepaint > -1 Then
                Repaint_If_Needed BtnRepaint
                Exit Sub
            End If
        End Select
        
End Sub
Private Sub Repaint_If_Needed(Index As Integer)
    'repaint one of the buttons
    Select Case Index
        Case 0 To 2
            'set defaults for repainting the lower button text back to default
            With Me
                .Font = MyBtnFont
                .FontSize = MyBtnFontSize
                .FontBold = MyBtnFontBold
                .ForeColor = BtnColorReg
                .FontItalic = MyBtnFontItalic
            End With
            Me.CurrentX = BtnTextData(Index, 0)
            Me.CurrentY = BtnTextData(Index, 1)
            Me.Print ButtonText(Index)
        Case 3
            'set defaults for repainting the upper button text back to default
            With Me
                .Font = "Arial"
                .FontSize = 8
                .FontBold = True
                .ForeColor = BtnColorReg
                .FontItalic = False
            End With
            Me.CurrentX = BtnTextData(Index, 0)
            Me.CurrentY = BtnTextData(Index, 1)
            Me.Print ButtonText(Index)
            
    End Select
    BtnRepaint = -1: OverButton = -1
    Me.MousePointer = 0
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'intercept a mouse_up from a form drag
    If OverButton = -1 Then Exit Sub
    'a button on the form has been clicked... figure out what to do
    If BoxMode = 0 Then
        If RetryOK And OverButton = 2 Then
            'in this condition, the 3rd button on an input box (ReTry) has been clicked...
            'this MouseUp event is only intercepted in the Function ShowInputBox which sets RetryOK=True
            Repaint_If_Needed 2
            UserInputTxt.Text = ""
            UserInputTxt.SetFocus
            Exit Sub 'don't hide the form... or vars get passed too soon.
        End If
    End If
    'return the button clicked
    BtnClicked = OverButton
    Me.Hide 'have to do this so function will get the variables needed to respond
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'release resources
    Set frmMsgBox = Nothing
    Unload Me
    
End Sub
