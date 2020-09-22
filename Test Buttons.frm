VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Message/Input Box Replacement Demo"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   Begin VB.CommandButton Command2 
      Caption         =   "UnLoad Nag Message Box"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   19
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Nag Message Box"
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Top             =   3600
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show Single Close Button Message Box"
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   3120
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Frame Frame2 
      Caption         =   " Message Box With ColorScheme "
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   4455
      Begin VB.CommandButton TestMessageBox 
         Caption         =   "Show Message Box"
         Height          =   375
         Left            =   2520
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Show Icon"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Olive"
         Height          =   255
         Left            =   1200
         TabIndex        =   12
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Blue"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Silver"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Caption         =   "White"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton TestTimedBox 
      Caption         =   "Show Timed Message Box"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   " Input Box With ColorScheme "
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CheckBox Check1 
         Caption         =   "Show Icon"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Olive"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Blue"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Silver"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton TestInputBox 
         Caption         =   "Show Input Box"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "White"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Line Line2 
      X1              =   64
      X2              =   48
      Y1              =   280
      Y2              =   280
   End
   Begin VB.Line Line1 
      X1              =   48
      X2              =   48
      Y1              =   272
      Y2              =   280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Waiting On Your Decision..."
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dim some strings for passing
Dim TitleMsg As String, PromptMsg As String
Dim Crap As String
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ GENERALLY SPEAKING... all functions are called by passing some or all of the following
'++ to the functions... here is a brief explanation of what means what.
'++ Note that not all functions do not use all of the parameters below...
'++ but most do... so examine the function to see what actually exists for your needs
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Call the function using this form:
'YourVariable = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg,[optional vars])
    'MsgTitle As String, <- Text to display in "Title Bar"
    'MsgPrompt As String, <- Text to display in prompt area
    'Optional MsgMode As Integer = 1, <- Flag that sets alot of stuff...
    '...0 = Input Box with (1-3) lower buttons, and an Upper "X" close button
    '...1 = Message Box with (1-3) lower buttons, and an Upper "X" close button
    '...2 = Message Box with No Lower buttons... Only an Upper "X" close button
    '...3 = Message Box with NO BUTTONS at all
    '...4 = Message Box that unloads after a specified time
    'Optional ByVal MsgTitleFontColor As String = "000000000", <- Forecolor of Title Bar text
    'Optional ByVal MsgPromptFontColor As String = "000000000", <- Forecolor of prompt text
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '++ Be careful with the colors in the strings above...   +
    '++ The format is RGB passed in a string as:             +
    '++ "RRRGGGBBB"... where the values passed must be       +
    '++ 3 digit positive integers as:                        +
    '++ the minimum -> 000 to 255 <- the maximum             +
    '++ SO, RED would be passed as: "255000000"              +
    '++ Blue as: "000000255"                                 +
    '++ Green as: "000255000".                               +
    '++ Of course, you may combine values as: "128202113"    +
    '++ The string passed must be 9 characters in length     +
    '++ Or the function will default to black... "000000000" +
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Optional ByVal ColorScheme As Integer = 0, <-0=White,1=Silver,2=Blue,3=Oliveish
    '... if this value > 1 then gradient is applied to title bar
    'Optional ByVal RepeatGradient As Boolean = False, <- Set to true to repeat gradient in body
    'Optional ByVal ShowIconNum As Integer = 0, <-What Icon to show: 0=None,(1=Custom...see readme),1=Application type,2=!,3=X,4=Info,5=?,6=Windows icon
    '... the Custom icon is the icon for the message box form set in the properties window
    '... See the read me.txt file
    'Optional ByVal BoxLeft As Integer = -1, <- MsgBox horizontal location: -1=center, other = me.left
    'Optional ByVal BoxTop As Integer = -1,<- MsgBox vertical location: -1=center, other = me.top
    'Optional ByVal DefaultAnswer As String = vbNullString, <- Provides a default answer for the input box
    'Optional ByVal Btn1Text As String = "OK", <- the caption for the first button
    'Optional ByVal Btn2Text As String = vbNullString, <- the caption for the 2nd button
    'Optional ByVal Btn3Text As String = vbNullString, <- the caption for the 3rd button
    '--> See the read me.txt file for more about the buttons!
    'Optional ByVal Dragit As Boolean = True, <-True = draggable form / False = NoDrag
    'Optional ByVal PlaySound As Integer = -1, <- Value must be > -2 and < 5
    '++ if enabled here... this plays a sound registered on your user's computer
    '++ -1 = Don't play sound  0 = Beep 1 = MB_IconAsterisk 2 = MB_IconQuestion
    '++ 3 = MB_IconExclamation  4 = MB_IconInformation
    '++ See the PlayAnnoyingSound sub for more info.
    'Optional ByVal MsgTitleFont As String = "Arial", <- Sets title bar font
    'Optional ByVal MsgTitleFontSize As Integer = 8, <- Sets title bar font size
    'Optional ByVal MsgTitleFontBold As Boolean = False, <- Sets title bar font bold
    'Optional ByVal MsgTitleFontItalic As Boolean = False, <- Sets title bar font italic
    'Optional ByVal MsgPromptFont As String = "Arial", <- Sets prompt font
    'Optional ByVal MsgPromptFontSize = 8, <- Sets prompt font size
    'Optional ByVal MsgPromptFontBold As Boolean = False, <- Sets prompt font bold
    'Optional ByVal MsgPromptFontItalic As Boolean = False), <- Sets prompt font italic
    
Private Sub Command1_Click()
    Label1.Caption = "Waiting On MessageBox Response..."
    'show a message box no close button
    TitleMsg = "Shareware Notice!"
                PromptMsg = "This Software" & vbCrLf & "Is Not Registered!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 3, , , 3, , 2, 10000, 10000, , , , , False, , , 10, , , , 10)
    Msg = "This form has no owner, so it will stay here until you unload it!" & vbCrLf & "No matter what you do with the demo form." & vbCrLf & vbCrLf & "Click the -UnLoad Nag Message Box- button to remove the message box"
    Title = "Important!"
    MsgBox Msg, vbOKOnly + vbExclamation, Title
    SetOne
End Sub

Private Sub Command2_Click()
    Unload frmMsgBox
    SetMost
End Sub

Private Sub Form_Load()
    SetMost
    TestTimedBox_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'this just makes sure you don't leave a nag screen hanging
    If Command2.Enabled = True Then
        Unload frmMsgBox
    End If
End Sub

Private Sub TestInputBox_Click()
    'Display an input box
    'this is accomplished by setting the MsgMode flag to 0...
    'ie.: Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0,...)
    Label1.Caption = "Waiting On InputBox Response..."
    Select Case Check1.Value
        Case 1
            If Option1.Value Then
                TitleMsg = "White ColorScheme (0) Demo... With Icon 1"
                PromptMsg = "You can see that this input box has red title text (size=10)(bold=true)" & vbCrLf & "The prompt text is blue... (size=12)... but could be any font / size / color / bold / italic." & vbCrLf & "Dragging is enabled.  You may drag this input box anywhere using any part of the input box except the text box or buttons." & vbCrLf & vbCrLf & "The icon above (#2) is the icon associated with this form by your computer." & vbCrLf & "This icon has been parsed from your computer... this form has only one (1!) control!" & vbCrLf & "(Hint... it's the text box...)"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0, "255000000", "000000255", , , 1, , , "Enter something!", "OK", "Cancel If You Like", "Clear Entry", , , , 10, True, , , 12)
            End If
            If Option2.Value Then
                TitleMsg = "Silver(1) ColorScheme Demo... With Icon 2 for those of you that want or need multiple line title bars... here you have it!" & vbCrLf & "Unless you have a huge monitor... this should be line 3 in the title!"
                PromptMsg = "This input box has a multiple line title bar.  The top 2 lines auto-wordwrapped... the bottom line was added." & vbCrLf & vbCrLf & "Some things to notice about this input box:" & vbCrLf & "... 1,2, or 3 (in this case... 2) color matched buttons that you provide the captions for... with hover effects" & vbCrLf & "... An (optional) color matched gradient can repeated in this prompt area" & vbCrLf & "... The form and buttons resize as needed..." & vbCrLf & "... The button font properties and hover colors can be set! (This is Courier 12 Bold & Italic)" & vbCrLf & vbCrLf & "All of this... and there is STILL ONLY 1 Control"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0, "000000255", , 1, True, 2, , , "Type some crap here", , "Screw It", , , 0, , 14, , , , 10, , , "255000000", "000000255", "Courier", 12, True, True)
            End If
            If Option3.Value Then
                TitleMsg = "LightBlue(2) ColorScheme Demo... With Icon 3"
                PromptMsg = "Sure... the same input box... but a single button you provide the caption for." & vbCrLf & "You control the font attributes of the Title Bar..." & vbCrLf & "And  the font attributes (font/color/bold/italic/size) of the prompt text... AS I GUESS YOU CAN SEE?" & vbCrLf & vbCrLf & "Drag is enabled on this input box... Drag it by any part except buttons (or text box)!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0, , "000000255", 2, True, 3, , , "My Answer is...", "Customized YES Button", , , , 2, "Comic Sans MS", 10, True, , , 10, , , , , , 14)
            End If
            If Option4.Value Then
                TitleMsg = "Olive...ish(3) ColorScheme Demo... With Icon 4"
                PromptMsg = "You can make your own color schemes." & vbCrLf & "Change the fonts/icons and play annoying sounds!." & vbCrLf & "What Else Could You Want?" & vbCrLf & vbCrLf & "How about:" & vbCrLf & "Put the message box where ever you want..." & vbCrLf & "Enable/Disable Dragging." & vbCrLf & "Won't leak off your screen."
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0, , "064128128", 3, True, 4, 2000, 2000, "Don't set there... type something!", "Use This", "I Give Up!", "Let Me ReTry!", False, , , 10)
            End If
        Case Else
            If Option1.Value Then
                TitleMsg = "White ColorScheme (0) Demo... WithOut Icon"
                PromptMsg = "Whats New In This Build???" & vbCrLf & vbCrLf & "••• Now... only 1 control on form (the text box)... all else is owner drawn" & vbCrLf & "••• Dramatically improved loading speed" & vbCrLf & "••• Optimized (smoother/faster/more accurate) gradient code... still non-API" & vbCrLf & "••• Much improved resize code..." & vbCrLf & "••• Multiple line title bars supported..." & vbCrLf & "••• Added wordwrap for title and prompt text..." & vbCrLf & "••• more font property support and lots more!  See the readme.txt file!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0, "000000255", , , , , , , "Enter something!", , "Cancel", "Retry", , , , 10, True, , , 10, True)
            End If
            If Option2.Value Then
                TitleMsg = "Silver(1) ColorScheme Demo... WithOut Icon"
                PromptMsg = "This is an input box... with two buttons that you provide the captions for." & vbCrLf & "The form and buttons resize as needed..." & vbCrLf & "Notice that the buttons and border follow the color scheme." & vbCrLf & vbCrLf & "The repeat gradient is NOT implemented on this input box."
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0, "000000255", , 1, , , , , "Type your crap here", , "Screw It", , , 0, , 10)
            End If
            If Option3.Value Then
                TitleMsg = "LightBlue(2) ColorScheme Demo... WithOut Icon"
                PromptMsg = "Looks just like the Windows message box doesn't it?" & vbCrLf & "You control the font attributes of the Title Bar..." & vbCrLf & "And  the font attributes (font/color/bold/italic/size) of the prompt text... AS I GUESS YOU CAN SEE?" & vbCrLf & vbCrLf & "Drag is disabled on this input box... No matter how hard you click... it ain't goin' nowhere!!!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0, , "000000255", 2, , , , , "My Answer is...", "Customized YES Button", , , False, 2, "Comic Sans MS", 10, True, , , 10)
            End If
            If Option4.Value Then
                TitleMsg = "Olive...ish(3) ColorScheme Demo... WithOut Icon"
                PromptMsg = "YOU can put your input box where ever you like." & vbCrLf & "Dock it left top/bottom/center... top/left/center..." & vbCrLf & "... you get the idea!" & vbCrLf & vbCrLf & "If you don't believe me... type:" & vbCrLf & "TopRight" & vbCrLf & "In the text box below and click the Do It button!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0, , "064128128", 3, , , 0, , "Don't set there... type something!", "Do It", "Screw It!", "Let Me ReTry!", False, , , 10)
                    If UCase(Crap) = "TOPRIGHT" Then
                        PromptMsg = "See there!"
                        Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 0, , "064128128", 3, , , 9000, -20, "Don't set there... type something!", "Do It", "Screw It!", "Let Me ReTry!", False, , , 10)
                    End If
            End If
    End Select
    'report the input box results
    DoEvents 'allow windows to repaint... a nice thing to do before doing the following!
    TitleMsg = "Your Response Is In!"
    PromptMsg = "Your Text Was:" & vbCrLf & Crap
    Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, , , 3, True, 4, , , , , , , , 0)
    If Crap = vbNullString Then Crap = "VBNullString"
    Label1.Caption = "Returned: " & Chr$(34) & Crap & Chr$(34)
End Sub
Private Sub TestMessageBox_Click()
    'show message box
    'this is accomplished by setting the MsgMode flag to 1...
    'ie.: Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1,...)
    Label1.Caption = "Waiting On MessageBox Response..."
    
    
    Select Case Check2.Value
        Case 1
            If Option5.Value Then
                TitleMsg = "Do We Have A Situation Here!"
                PromptMsg = "You have requested a file you do not have clearance for!" & vbCrLf & vbCrLf & "Should I Call The FBI???"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, "255000000", , , , 5, , , , "Yes", "No", , , 4, , 10, True, , , 10)
            End If
            If Option6.Value Then
                TitleMsg = "You Are An Un-Registered User!"
                PromptMsg = "What Do You Want To Do SlimeBall?" & vbCrLf & vbCrLf & "Turn Your Self In... Run For Cover" & vbCrLf & vbCrLf & " Or... Turn Your Computer Off!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, "000000255", , 1, True, 6, , , , "Turn Me In!", "Run", "GoodNight!", , 0, , 14, , , , 10, , , , , "Courier", 10)
            End If
            If Option7.Value Then
                TitleMsg = "Are You Nuts... Or What?"
                PromptMsg = "You CAN'T choose the Close Button Now!" & vbCrLf & "I guess they didn't have that many computers in the barn... did they?" & vbCrLf & vbCrLf & "Why don't your give it another try CowBoy?"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, , , 2, True, 3, , , , "I Get It Now!", , , , 2, "Comic Sans MS", 10, True, , , 10, , , , , , 14)
            End If
            If Option8.Value Then
                TitleMsg = "Software Update Manager"
                PromptMsg = "Your software is out of date." & vbCrLf & vbCrLf & "An Update Is Available And Recommended."
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, , "064128128", 3, True, 4, 2000, 2000, , "Update Now", "Remind Me Later", , False, , , 10)
            End If
        Case Else
            If Option5.Value Then
                TitleMsg = "What's New???"
                PromptMsg = "Whats New In This Build???" & vbCrLf & vbCrLf & "••• Now... only 1 control on form (the text box)... all else is owner drawn" & vbCrLf & "••• Dramatically improved loading speed" & vbCrLf & "••• Optimized (smoother/faster/more accurate) gradient code... still non-API" & vbCrLf & "••• Much improved resize code..." & vbCrLf & "••• Multiple line title bars supported..." & vbCrLf & "••• Added wordwrap for title and prompt text..." & vbCrLf & "••• more font property support and lots more!  See the readme.txt file!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, "000000255", , , , , , , , "OK... I Get It!", , , , , , 10, True, , , 10, True, , , , , , True)
            End If
            If Option6.Value Then
                TitleMsg = "Don't you wish I wasn't bored with writing demo code?"
                PromptMsg = "It's a message box! (for crying-out-loud!)" & vbCrLf & vbCrLf & "You know how they work..." & vbCrLf & vbCrLf & "BUT... This one is even better!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, "000000255", , 1, , , , , , "I Agree", , , , 0, , 10, True, , , 10)
            End If
            If Option7.Value Then
                TitleMsg = "LightBlue(2) ColorScheme Demo... WithOut Icon"
                PromptMsg = "Looks just like the Windows message box doesn't it?" & vbCrLf & "You control the font attributes of the Title Bar..." & vbCrLf & "And  the font attributes (font/color/bold/italic/size) of the prompt text... AS I GUESS YOU CAN SEE?" & vbCrLf & vbCrLf & "Drag is disabled on this input box... No matter how hard you click... it ain't goin' nowhere!!!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, , "000000255", 2, , , , , , "Customized YES Button", , , False, 2, "Comic Sans MS", 10, True, , , 10)
            End If
            If Option8.Value Then
                TitleMsg = "Olive...ish(3) ColorScheme Demo... WithOut Icon"
                PromptMsg = "YOU can put your input box where ever you like." & vbCrLf & "Dock it left top/bottom/center... top/left/center..." & vbCrLf & "... you get the idea!" & vbCrLf & vbCrLf & "If you don't believe me..." & vbCrLf & "Click The MOVE ME button!"
                Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, , , 3, , , 0, , , "Don't Move Me", "MOVE ME", , False, , , 10)
                    If Crap = "1" Then
                        TitleMsg = "I'm Down Here Now!... but Not behind your toolbar!"
                        PromptMsg = "See there!"
                        Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 1, , , 3, , , -2, 2000, , , , , False, , , 10)
                    End If
            End If
    End Select
    'report the input box results
    Label1.Caption = "Returned: " & Chr$(34) & Crap & Chr$(34)

End Sub
Private Sub TestTimedBox_Click()
    'show a message box that goes away by itself
    TitleMsg = "Welcome To The UpDated Drop-In MsgBox/Input Box Demo!"
    PromptMsg = "Simply Drop This Form Into Your Project And Instantly Add Much Cooler" & vbCrLf & "Message And Input Boxes In Which You Can Control Almost Everything!" & vbCrLf & vbCrLf & "Only One Function To Call... Multiple Results!  Much More Flexible Than The" & vbCrLf & "Windows Message/Input Boxes Your Used To!" & vbCrLf & vbCrLf & "Multiple Line Title Bars, Color Schemes (XP or not), Gradients, Button Text," & vbCrLf & "Button Rollover Colors, Dock/Drag Your Message/Input Boxes, Nag Screens," & vbCrLf & "Sounds, Icons, and Lots More!" & vbCrLf & vbCrLf & "Give It A Try!  It's Newer, Faster, Lighter & Better Than Before..." & vbCrLf & "Oh yeah... NOW There's Only One Control On The Form!" & vbCrLf & vbCrLf & "This message box will close in 10 Seconds.  To see it again... try the" & vbCrLf & "TIMED MESSAGE BOX button."
    Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 4, , "064128128", 1, True, 4, , , , , , , , , , 10, True, , , 10, , , , , , , , , 10)
    
End Sub
Private Sub Command4_Click()
    'screw it... you must be done playing!
    End
End Sub

Private Sub Command5_Click()
    'show a message box with X close button only
    TitleMsg = "Buttons... We Don't Need No Stinkin' Buttons!"
    PromptMsg = "YOU DON'T NEED ALL THOSE BUTTONS!" & vbCrLf & vbCrLf & "Simply set the BoxMode Flag to 2" & vbCrLf & "... and the lower buttons are gone!"
    Crap = frmMsgBox.ShowMsgBox(TitleMsg, PromptMsg, 2, , , 3, , , , , , , , , False, , , 10, , , , 10)
                    
End Sub
Private Sub SetMost()
    TestInputBox.Enabled = True
    TestMessageBox.Enabled = True
    TestTimedBox.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = False
    Command4.Enabled = True
    Command5.Enabled = True
End Sub
Private Sub SetOne()
    TestInputBox.Enabled = False
    TestMessageBox.Enabled = False
    TestTimedBox.Enabled = False
    Command1.Enabled = False
    Command2.Enabled = True
    Command4.Enabled = False
    Command5.Enabled = False
End Sub

