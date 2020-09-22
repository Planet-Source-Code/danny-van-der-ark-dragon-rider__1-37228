VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Dragon Rider"
   ClientHeight    =   5820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7680
   Icon            =   "Intro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Intro.frx":324A
   ScaleHeight     =   388
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   Begin VB.PictureBox pIntro 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   0
      Picture         =   "Intro.frx":94A8E
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   0
      Width           =   7680
      Begin VB.Timer tFade 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7080
         Top             =   180
      End
      Begin VB.PictureBox pText 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3840
         Left            =   3990
         Picture         =   "Intro.frx":F4AD2
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   512
         TabIndex        =   1
         Top             =   -255
         Visible         =   0   'False
         Width           =   7680
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "danny@slave-studios.co.uk"
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   1035
      TabIndex        =   7
      Top             =   5565
      Width           =   6060
   End
   Begin VB.Label lHelp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1020
      Left            =   180
      TabIndex        =   6
      Top             =   3885
      Width           =   7320
   End
   Begin VB.Label ButCommand 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Orbus Multiserif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   5670
      TabIndex        =   5
      Top             =   4980
      Width           =   1515
   End
   Begin VB.Label ButCommand 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "&PLAY"
      BeginProperty Font 
         Name            =   "Orbus Multiserif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   420
      TabIndex        =   4
      Top             =   4980
      Width           =   1515
   End
   Begin MediaPlayerCtl.MediaPlayer mp1 
      Height          =   585
      Left            =   2280
      TabIndex        =   3
      Top             =   4485
      Visible         =   0   'False
      Width           =   3090
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -500
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dragon Rider is (C) Copyright 2002 by Danny E.L. Diablo,--"
      ForeColor       =   &H00400000&
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   5355
      Width           =   6675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ButCommand_Click(Index As Integer)

 ButCommand(0).Enabled = False
 ButCommand(1).Enabled = False

 MP1.Stop
 tFade.Enabled = False
 
 If Index = 0 Then
   'start
    Level.Show vbModal
    
   'recover
    Form1.Visible = True
    ButCommand(0).Enabled = True
    ButCommand(1).Enabled = True
    
 Else
   'Exit
    Unload Me
    End
 End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyP Then Call ButCommand_Click(0)
 If KeyCode = vbKeySpace Then Call ButCommand_Click(0)
 
 If KeyCode = vbKeyX Then Call ButCommand_Click(1)
 If KeyCode = vbKeyQ Then Call ButCommand_Click(1)
 If KeyCode = vbKeyEscape Then Call ButCommand_Click(1)

End Sub

Private Sub Form_Load()

'Quick start
 Randomize
 
 lHelp.Caption = "Use 1...4 to fire, use cursor keys to navigate, press 'M' to kill the music, or 'Q' to quit.       Don't press 'E' to toggle the Level Edit mode (1..5 for objects, A...D for characters, P for player-start or S to save map) , or 'F' for FPS counter, nor should you press 'D' for debug information"
 
'Check for Single or dual screen setup & center forms...
 If Screen.Width / Screen.Height > 1.5 Then
   'Dual screen setup
    Form1.Move Screen.Width / 4 - Form1.Width / 2, Screen.Height / 2 - Form1.Height / 2
 Else
   'Single screen setup
    Form1.Move Screen.Width / 2 - Form1.Width / 2, Screen.Height / 2 - Form1.Height / 2
 End If
 
'Start Background Music
'MP1.FileName = App.Path & "\Music\LowQ\002_Theme - Mortal Kombat Annihilation.sm.mp3"
'MP1.Play
 
'Clear the intro-header background
 pIntro.Cls
 
'Enable FadeIn IRQ
 Call tFade_Timer

'start GAME inmediately (debug)
'  Call ButCommand_Click(0)
 
'## This IRQ fades in the intro screen

 pIntro.Refresh
 
End Sub

Private Sub tFade_Timer()
Static Merge As Integer

'first time init
 If tFade.Enabled = False Then Merge = 0
 
'prevent re-entry
 tFade.Enabled = False
 
'Blend in...
 Merge = Merge + 1
 
'## This IRQ fades in the intro screen
 FoxAlphaBlend pIntro.HDC, 0, 0, pIntro.ScaleWidth, pIntro.ScaleHeight, pText.HDC, 0, 0, Merge, 0, FOX_USE_MASK
 pIntro.Refresh
 
'Check if blend finished...
 If Merge <= 254 Then
   'Continue fading in...
    tFade.Enabled = True
 Else
   'Print image at 100% opacity
    FoxAlphaBlend pIntro.HDC, 0, 0, pText.ScaleWidth, pText.ScaleHeight, pText.HDC, 0, 0, 255, 0, FOX_USE_MASK
 End If
  
End Sub
