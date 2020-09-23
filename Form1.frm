VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Test"
   ClientHeight    =   4560
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4560
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTestAlwaysActive 
      BackColor       =   &H00C0C000&
      Caption         =   "Test ""Always Active"" Options"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3705
      Width           =   4200
   End
   Begin VB.TextBox txtCaption 
      Height          =   315
      Left            =   1140
      TabIndex        =   5
      Text            =   "La Volpe Rules the Den !!"
      Top             =   3345
      Width           =   4215
   End
   Begin VB.ComboBox cboMenuBar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2835
      List            =   "Form1.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2970
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0006CDD2&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4845
      TabIndex        =   0
      ToolTipText     =   "Disabling close button prevents closing form. Click here instead"
      Top             =   405
      Width           =   555
   End
   Begin VB.ListBox lstOptions 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      ItemData        =   "Form1.frx":0054
      Left            =   45
      List            =   "Form1.frx":007C
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   765
      Width           =   5355
   End
   Begin VB.CheckBox chkActivate 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Use subclass procedues to customize this form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   450
      Width           =   5280
   End
   Begin VB.ComboBox cboTBar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Form1.frx":016D
      Left            =   45
      List            =   "Form1.frx":017A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2970
      Width           =   2535
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Caption         =   "Bottom Aligned Object -- just for testing purposes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4215
      Width           =   5445
   End
   Begin VB.Data Data1 
      Align           =   1  'Align Top
      Caption         =   "Top Aligned Object -- just for testing purposes"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Enabled         =   0   'False
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Width           =   5445
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CAPTION"
      Height          =   225
      Left            =   150
      TabIndex        =   6
      Top             =   3405
      Width           =   945
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&New Project"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open Project"
         Index           =   1
         Begin VB.Menu mnuOpen 
            Caption         =   "From Floppy"
            Index           =   0
            Begin VB.Menu mnuOpen0 
               Caption         =   "A or B Drive"
            End
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "From CD"
            Index           =   1
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "From Harddrive"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Edit"
      Index           =   1
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Project"
      Enabled         =   0   'False
      Index           =   2
   End
   Begin VB.Menu mnuMain 
      Caption         =   "F&ormat"
      Index           =   3
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Debug"
      Index           =   4
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Run"
      Index           =   5
      Begin VB.Menu mnuRun 
         Caption         =   "Start"
         Index           =   0
      End
      Begin VB.Menu mnuRun 
         Caption         =   "Start wtih &Full Compile"
         Index           =   1
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Q&uery"
      Index           =   6
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Diagram"
      Index           =   7
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is a draft in preparation of a full-blown custom window skinner
' This first draft does not skin per se; its sole purpose is to
' see if I was able to keep Windows from drawing over the non-client area.
' To do that, I was forced to continuously toggle certain key window styles
' at specific moments. This, in my opinion, creates a very "noisy" application;
' meaning, that there are many more window messages being processed due to the
' regular removal/addition of window styles.  The real way to skin an app
' would be to create your own window class & control everything... However;
' since so many of us like the ease of using VB forms, I'll continue to
' try to force VB to do what I want vs. what it wants ;)

' If you would be so kind as to abuse this project to see if you can find
' any instances where windows draws over the nonclient area, menus beep at
' you when they shouldn't and provide suggestions, I would definitely
' appreciate it.  When you find bugs please include the O/S you were using.
' I have not been able to make this crash without hitting the End button or
' putting in an End statement. So far, the subclassing seems sound.

' NOTE TO SELF (Top 10 TO DOs)
' 1. Menu font changes need to resize nonclient area as needed
' 2. Add user-defined min/max window size restrictions
' 3. Finish the routines to add other titlebar buttons
' 4. Allow not showing captions on window (skins), but show in taskbar
' 5. Add other typical sublcassing functions like
'       - system tray routines
'       - mouse over menu feedback
'       - possibly provide feedback for API-created menus
'       - ideas?
' 6. Tackle the MDI menubar problems and make compatible with MDIs
' 7. Add true skinning to include real-time user-drawn overrides
' 8. Include routines to automatically scale non-rectangular window
'       shape regions if applied to the window (done just need to add to project)
' 9. Offer some way to allow user to customize how menubar items look
'   when hilighted and/or selected (current routines are too simple & restrictive)
'       - Note: no intention to create vertical menu bars
' 10. The final step is to include routines for owner-drawn menus
'       -- also remove all default window drawing...
'        i.e, the final cut will not have vertical captions, will not
'        have default gradient titlebars, will not have any default
'        window painting at all. This was done just so I can get the
'        titlebar out of the way to more easily tell if windows was attempting
'        to paint the non-client area. The borders, titlebars, menubars, buttons
'        etc will be supplied by the user. Should you want to keep
'        this project for your playtime, save it in a good place 'cause it
'        won 't stay on PSC when the final cut is released.

' For XP users; don't get hung up on the fact that min/max/close buttons are
' drawn in Win95/98 style. In a skinned application, these will be user-defined
' icons/jpgs/whatever. Just the fact that I can tell when the mouse is
' over one of them and that they work is all that is important.  If you hover
' over those with the vertical caption, the tooltip should always appear if
' the button is enabled, no matter where it is physically located on the
' window. That was a bit creative on my part.

' A request for a specific suggestion....
' Most skinned forms have a set/non-adjustable menubar rectangle somewhere in
' or around the titlebar region.  This is simple enough to account for by
' having users dictate the rectangle dimensions.  However, since there is no
' true menubar in this case and keeping MDI children in mind, what suggestions
' could you make to bring a maximized MDI child's menubar onto a static
' MDI parent's "menu bar"? Kinda looking for ideas in visual design,
' not how-to's? I have 2 ideas but really don't like either of them.

' the classes are pretty heavily commented. API declarations exist in each
' class/module at the moment as I haven't decided how to finalize this project
' (i.e., as classes, as a DLL, or as a usercontrol).  The final cut will most
' likely have all APIs used by more than one class/module pulled out and made
' Public in a single module.

Option Explicit
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTCAPTION As Long = 2
Private Const MF_SEPARATOR As Long = &H800&
Private Const MF_DISABLED As Long = &H2&
Private Const MF_GRAYED As Long = &H1&
Private Const MF_DEFAULT As Long = &H1000&


' optional statement & used only if real-time overriding drawing is performed
Implements CustomWindowCalls

' required & must be initialized somewhere
Private LaVolpe_Window As clsCustomWndow

Private gradQuality As Byte ' used for gradient quality while resizing window


Private Sub cboMenuBar_Click()
If LaVolpe_Window Is Nothing Then Exit Sub

' modify menubar font, colors & also the titlebar font

Dim fColor(0 To 3) As Long, fIndex As Long
Dim hColor(0 To 1) As Long
Dim tFont As StdFont, mFont As StdFont

Set mFont = LaVolpe_Window.Font_MenuBar

Select Case cboMenuBar.ListIndex
Case 0: 'defaults
    fColor(fcDisabled) = vbGrayText
    fColor(fcSelected) = vbMenuText
    hColor(0) = vbHighlightText
    hColor(1) = vb3DShadow
    Set tFont = Nothing
    Set mFont = Nothing
Case 1: ' blue menubar
    fColor(fcDisabled) = RGB(52, 52, 52)
    fColor(fcSelected) = vbBlue
    hColor(0) = &H800080
    hColor(1) = vbCyan
    Set tFont = LaVolpe_Window.Font_TBar
    tFont.Name = "Book Antiqua"
Case 2: ' red menubar
    fColor(fcDisabled) = RGB(92, 92, 92)
    fColor(fcSelected) = &H40C0&
    hColor(0) = vbWhite
    hColor(1) = RGB(92, 92, 92)
    Set tFont = LaVolpe_Window.Font_TBar
    tFont.Name = "Comic Sans MS"
Case 3: ' green menubar
    fColor(fcDisabled) = RGB(0, 128, 0)
    fColor(fcSelected) = RGB(255, 64, 255)
    hColor(0) = &H8080&
    hColor(1) = vbGreen
    Set tFont = LaVolpe_Window.Font_TBar
    tFont.Name = "Arial"
End Select
fColor(fcInActive) = vbGrayText
fColor(fcEnabled) = vbMenuText

If Not tFont Is Nothing Then
    If cboMenuBar.ListIndex = 3 Then
        tFont.Size = 12
        mFont.Size = 10
    Else
        tFont.Size = 8
        mFont.Size = 8
    End If
    tFont.Bold = True
End If

If Not mFont Is Nothing Then mFont.Name = tFont.Name
LaVolpe_Window.Font_MenuBar = mFont
LaVolpe_Window.Font_TBar = tFont

For fIndex = 0 To 3
    LaVolpe_Window.FontColor_Menubar(fIndex) = fColor(fIndex)
Next
If cboMenuBar.ListIndex Mod 2 = 0 Then
    LaVolpe_Window.MenuSelect_FlatStyle = False
Else
    LaVolpe_Window.MenuSelect_FlatStyle = True
End If
LaVolpe_Window.SetMenuSelectColors hColor(0), hColor(1)
End Sub

Private Sub cboTBar_Click()
If LaVolpe_Window Is Nothing Then Exit Sub

' modify titlebar colors

Select Case cboTBar.ListIndex
Case 0 ' standard blue
    LaVolpe_Window.FontColor_TBar(True) = vbActiveTitleBarText
    LaVolpe_Window.FontColor_TBar(False) = vbInactiveTitleBarText
Case 1: ' red titlebar
    LaVolpe_Window.FontColor_TBar(True) = vbWhite
    LaVolpe_Window.FontColor_TBar(False) = vbGrayText
Case 2: ' green titlebar
    LaVolpe_Window.FontColor_TBar(True) = vbWhite
    LaVolpe_Window.FontColor_TBar(False) = vbGrayText
End Select
End Sub

Private Sub chkActivate_Click()
If chkActivate = 0 Then
    Set LaVolpe_Window = Nothing    ' should unsublass & clean up
Else
    Test_SetOptions ' set options & re-subclass
End If
cmdTestAlwaysActive.Enabled = (chkActivate = 1)
End Sub

Private Sub cmdTestAlwaysActive_Click()
If lstOptions.Selected(10) = True Then
    Form2.Show 'not shown with an owner which would normally gray out the current form
Else
    If lstOptions.Selected(11) = False Then
        MsgBox "Neither of the options to keep the form looking active has been selected", vbInformation + vbOKOnly, "Oops"
    Else
        If MsgBox("Will try to display a Window's Explorer window to take the " & vbNewLine & _
            "focus off this thread. If the window doesn't display, simply " & vbNewLine & _
            "click elsewhere on your desktop to force a change in the focused thread.", vbOKCancel + vbInformation, "Testing") = vbOK Then
                Shell "explorer.exe", vbNormalFocus
        End If
    End If
End If
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub CustomWindowCalls_EnterExitSizing(ByVal BeginSizing As Boolean, UserRedrawn As Boolean)
' Example of using gradients at a lower quality for faster drawing
' during resizing
If BeginSizing Then
    gradQuality = 5
Else
    ' done resizing; reset gradient quality to best & refresh window
    gradQuality = 0
    ' by not setting UserRedraw to True, class will automatically redraw window
End If
End Sub

Private Sub CustomWindowCalls_UserButtonClick(ByVal ID As String)

' not used in this sample project

Select Case LCase(ID)
Case "logo"
'    LaVolpe_Window.TitleBar_AlwaysActive(False) = True
'    MsgBox "Got the title bar button"
'    LaVolpe_Window.TitleBar_AlwaysActive(False) = True
'    Unload Me
End Select
End Sub

Private Sub CustomWindowCalls_UserDrawnButton(ByVal ID As String, ByVal State As Integer, ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long)
' not used in this sample project
End Sub

Private Sub CustomWindowCalls_UserDrawnMenuBar(ByVal hDC As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal HasFocus As Boolean, Modified As Boolean)

Dim fColor As Long
Select Case cboMenuBar.ListIndex
Case 0: Exit Sub
Case 1: ' blue bar
    fColor = RGB(128, 128, 255)
Case 2: ' red bar
    fColor = RGB(255, 128, 128)
Case 3: ' green bar
    fColor = RGB(128, 255, 128)
Case 4:
End Select
    If HasFocus Then
        GradientFill fColor, vbWhite, hDC, 0, 0, Cx, Cy, gradQuality
    Else
        GradientFill RGB(221, 221, 221), RGB(241, 241, 241), hDC, 0, 0, Cx, Cy, gradQuality
    End If
    Modified = True
End Sub

Private Sub CustomWindowCalls_UserDrawnTitleBar(ByVal hDC As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal HasFocus As Boolean, Modified As Boolean)
Dim fColor As Long, tColor As Long
Select Case cboTBar.ListIndex
Case 0: Exit Sub
Case 1: ' red titlebar
    fColor = RGB(51, 0, 0): tColor = vbRed
Case 2: ' green titlebar
    fColor = RGB(0, 51, 0): tColor = vbGreen
End Select
    If LaVolpe_Window.VerticalCaption Then Cy = -Cy
    If HasFocus Then
        GradientFill fColor, tColor, hDC, 0, 0, Cx, Cy, gradQuality
    Else
        GradientFill RGB(221, 221, 221), RGB(221, 221, 221), hDC, 0, 0, -Cx, Cy, gradQuality
    End If
    Modified = True
End Sub

Private Sub Form_Load()

cboMenuBar.ListIndex = 1
cboTBar.ListIndex = 1

Dim hSysM As Long, poM As Long
hSysM = GetSystemMenu(hWnd, 0)
poM = CreatePopupMenu()
AppendMenu hSysM, MF_SEPARATOR Or MF_DISABLED, 50, ""
AppendMenu hSysM, &H10&, poM, "System Menu Add On"
AppendMenu poM, MF_DEFAULT, 1, "Item #1"
AppendMenu poM, MF_DISABLED Or MF_GRAYED, 2, "Item #2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set LaVolpe_Window = Nothing
End Sub

Private Sub lstOptions_ItemCheck(Item As Integer)
If LaVolpe_Window Is Nothing Then Exit Sub
If Len(lstOptions.Tag) > 0 Then Exit Sub
'Center Caption
'Vertical Caption
'Disable Close
'Disable Maximize
'Disable Minimize
'Hide Window Icon
'Cannot be Resized
'Cannot be Moved
'Hide Disabled Buttons
'Take off Windows Taskbar
'Look Active while in Thread
'Always Look Active
With LaVolpe_Window
    Select Case Item
    Case 0: .CenteredCaption = (lstOptions.Selected(Item) = True)
    Case 1: .VerticalCaption = (lstOptions.Selected(Item) = True)
    Case 2: .EnableTBarBtns(smClose) = (lstOptions.Selected(Item) = False)
    Case 3: .EnableTBarBtns(smMaximize) = (lstOptions.Selected(Item) = False)
    Case 4: .EnableTBarBtns(smMinimize) = (lstOptions.Selected(Item) = False)
    Case 5: .EnableTBarBtns(smSysIcon) = (lstOptions.Selected(Item) = False)
    Case 6: .EnableTBarBtns(smSize) = (lstOptions.Selected(Item) = False)
    Case 7: .EnableTBarBtns(smMove) = (lstOptions.Selected(Item) = False)
    Case 8: .HideDisabledButtons = (lstOptions.Selected(Item) = True)
    Case 9: .ShowInTaskBar = (lstOptions.Selected(Item) = False)
    Case 10
        lstOptions.Tag = "NoRecurse"
        If lstOptions.Selected(Item) = True Then
            lstOptions.Selected(Item + 1) = False
        End If
        .AlwaysActive(False) = (lstOptions.Selected(Item) = True)
        lstOptions.Tag = ""
    Case 11
        lstOptions.Tag = "NoRecurse"
        If lstOptions.Selected(Item) = True Then
            lstOptions.Selected(Item - 1) = False
        End If
        .AlwaysActive(True) = (lstOptions.Selected(Item) = True)
        lstOptions.Tag = ""
    End Select
End With
End Sub

Private Sub mnuMain_Click(Index As Integer)
Select Case Index
Case 0
Case 1: Debug.Print "Got menubar item "; mnuMain(Index).Caption
Case 2: Debug.Print "Got menubar item "; mnuMain(Index).Caption
Case 3: Debug.Print "Got menubar item "; mnuMain(Index).Caption
Case 4: Debug.Print "Got menubar item "; mnuMain(Index).Caption
Case 5
Case 6: Debug.Print "Got menubar item "; mnuMain(Index).Caption
Case 7: Debug.Print "Got menubar item "; mnuMain(Index).Caption
End Select
End Sub

Private Sub mnuOpen0_Click()
Debug.Print "got deep nested menu item"
End Sub

Private Sub Test_SetOptions()

Dim sysButtonIDs As Long
'Center Caption
'Vertical Caption
'Disable Close
'Disable Maximize
'Disable Minimize
'Hide Window Icon
'Cannot be Resized
'Cannot be Moved
'Hide Disabled Buttons
'Take off Windows Taskbar

If Not LaVolpe_Window Is Nothing Then Set LaVolpe_Window = Nothing
If lstOptions.Selected(11) = True Then lstOptions.Selected(10) = False

Set LaVolpe_Window = New clsCustomWndow
With LaVolpe_Window
    .CenteredCaption = lstOptions.Selected(0)
    .VerticalCaption = lstOptions.Selected(1)
    If lstOptions.Selected(2) Then sysButtonIDs = sysButtonIDs Or smClose
    If lstOptions.Selected(3) Then sysButtonIDs = sysButtonIDs Or smMaximize
    If lstOptions.Selected(4) Then sysButtonIDs = sysButtonIDs Or smMinimize
    If lstOptions.Selected(5) Then sysButtonIDs = sysButtonIDs Or smSysIcon
    If lstOptions.Selected(6) Then sysButtonIDs = sysButtonIDs Or smSize
    If lstOptions.Selected(7) Then sysButtonIDs = sysButtonIDs Or smMove
    .EnableTBarBtns(sysButtonIDs) = False
    .HideDisabledButtons = (lstOptions.Selected(8) = True)
    If lstOptions.Selected(11) = True Then
        .AlwaysActive(True) = True
    ElseIf lstOptions.Selected(10) = True Then
        .AlwaysActive(False) = True
    End If
    .ShowInTaskBar = (lstOptions.Selected(9) = False)
End With
Call cboMenuBar_Click
Call cboTBar_Click
Call txtCaption_Change
LaVolpe_Window.BeginCustomWindow Me, Me
End Sub

Private Sub txtCaption_Change()
If Not LaVolpe_Window Is Nothing Then
    ' if the window is not subclassed, you must pass the hWnd parameter
    ' otherwise hWnd is optional. Suggest always passing it, but if you
    ' want to test whether or not the window is subclassed you can
    ' call the property IsSubclassed(hWnd)
    LaVolpe_Window.SetCaption txtCaption.Text, hWnd
Else
    Me.Caption = txtCaption.Text
End If
End Sub
