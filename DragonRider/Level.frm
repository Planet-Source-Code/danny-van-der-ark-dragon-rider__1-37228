VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Level 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Level 1"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "Orbus Multiserif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Level.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Level.frx":203A
   ScaleHeight     =   508
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   786
   Begin VB.Timer tstart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   930
   End
   Begin VB.Timer tKey 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   45
      Top             =   30
   End
   Begin VB.Frame Frame5 
      Caption         =   "Monsters"
      Height          =   4500
      Left            =   7725
      TabIndex        =   7
      Top             =   45
      Visible         =   0   'False
      Width           =   1455
      Begin VB.PictureBox pMonster 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00C000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3840
         Left            =   105
         Picture         =   "Level.frx":9387E
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   320
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   4800
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "VFX"
      Height          =   4500
      Left            =   9255
      TabIndex        =   16
      Top             =   45
      Visible         =   0   'False
      Width           =   1155
      Begin VB.PictureBox pVFX 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00C000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1920
         Left            =   105
         Picture         =   "Level.frx":CF8C2
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   640
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   345
         Visible         =   0   'False
         Width           =   9600
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Collision Buffer"
      Height          =   675
      Left            =   30
      TabIndex        =   14
      Top             =   5955
      Visible         =   0   'False
      Width           =   1815
      Begin VB.PictureBox pCollision 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   720
         Left            =   165
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Pass Buffer"
      Height          =   870
      Left            =   30
      TabIndex        =   12
      Top             =   6675
      Visible         =   0   'False
      Width           =   1845
      Begin VB.PictureBox pPass 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   30000
         Left            =   150
         Picture         =   "Level.frx":10B906
         ScaleHeight     =   2000
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   384
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
         Width           =   5760
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Code Buffer"
      Height          =   750
      Left            =   1980
      TabIndex        =   9
      Top             =   6690
      Visible         =   0   'False
      Width           =   1920
      Begin VB.PictureBox pCode 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   945
         Left            =   165
         ScaleHeight     =   63
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   86
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Timer tFPS 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   30
      Top             =   465
   End
   Begin VB.Frame Frame4 
      Caption         =   "BCK Buffer"
      Height          =   750
      Left            =   1965
      TabIndex        =   5
      Top             =   5940
      Visible         =   0   'False
      Width           =   1920
      Begin VB.PictureBox pBACK 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C00000&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   165
         ScaleHeight     =   18
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   86
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Player"
      Height          =   1530
      Left            =   4005
      TabIndex        =   1
      Top             =   5940
      Visible         =   0   'False
      Width           =   6195
      Begin VB.PictureBox pPlayer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00C000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   2
         Left            =   1635
         Picture         =   "Level.frx":115AB1
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox pPlayerFin 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1110
         Left            =   2460
         ScaleHeight     =   74
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   214
         TabIndex        =   18
         Top             =   255
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.PictureBox pPlayer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00C000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   0
         Left            =   105
         Picture         =   "Level.frx":1175F5
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.PictureBox pPlayer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillColor       =   &H00C000C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   1
         Left            =   855
         Picture         =   "Level.frx":119139
         ScaleHeight     =   48
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   48
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.PictureBox VIEW 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3840
      Left            =   960
      MousePointer    =   99  'Custom
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   1440
      Width           =   5760
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ammo"
      Height          =   4500
      Left            =   10470
      TabIndex        =   4
      Top             =   45
      Visible         =   0   'False
      Width           =   1215
      Begin VB.PictureBox pAmmo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1980
         Left            =   75
         Picture         =   "Level.frx":11AC7D
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   384
         TabIndex        =   11
         Top             =   360
         Width           =   5820
      End
   End
   Begin VB.Label lHP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   4680
      TabIndex        =   25
      Top             =   1065
      Width           =   1935
   End
   Begin MediaPlayerCtl.MediaPlayer MP1 
      Height          =   1260
      Left            =   7770
      TabIndex        =   24
      Top             =   4605
      Visible         =   0   'False
      Width           =   1950
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
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
      PlayCount       =   99
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
      Volume          =   -1399
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lDebug 
      BackStyle       =   0  'Transparent
      Caption         =   "Debug Information:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   960
      TabIndex        =   23
      Top             =   5295
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Label lEdit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(edit mode status)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2340
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Label lFps 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "25 fps (25)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   5865
      TabIndex        =   19
      Top             =   5325
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lScore 
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE 00000"
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   20
      Top             =   1065
      Width           =   1935
   End
   Begin VB.Label lScore 
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE 00000"
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   21
      Top             =   1095
      Width           =   1935
   End
End
Attribute VB_Name = "Level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bullet_Damage(ByVal BulId As Integer, ByVal TileId As Integer)
Static rDam As Single

'## This little piggy applies 'normal' damage (ie. not-fire) to the target
'## All damage a bullet can still do is transfered to the target, depending
'## on the amount of damage left in the bullet, this may or may not kill
'## the target. If the target is a dead-body (corps) the bullet loses no
'## power.

With Bullet(BulId)

    'Apply damage to monster
     rDam = TILE(TileId).HitPoints
     TILE(TileId).HitPoints = TILE(TileId).HitPoints - .Damage

    'Subtract power lost on monster from bullet power,
    'unless bullet hit a dead body
     If Not TILE(TileId).IsDead Then
        .Damage = .Damage - rDam
     Else
        'Subtract 1 from the bullet for wear and tear
        .Damage = .Damage - 1
     End If
End With

End Sub

Private Sub Bullet_DamageFire(ByVal BulId As Integer, ByVal TileId As Integer)
Static rDam As Single

'## This routine takes care of the complex damage fire can do.
'##
'## 1. Fire weapons don't do damage directly!
'##    They set the target on fire doing indirect damage:
'##    - a VFX is created, displaying the fire FX and draining their hitpoints.
'##    - An invisible fire-bullet is created and locked to the tile which
'##      can hurt other tiles when they come in contact (ie. spreading the fire).
'## 2. All damage that can be applied to the target will be subtracted from the
'##    bullets remaining damage.
'## 3. When a fire-bullet doesn't have enough damage left to kill the target,
'##    then the target will live.

With Bullet(BulId)
                            
    'Calculate the max damage that can be applied.
     rDam = TILE(TileId).HitPoints
     
     If .Damage >= rDam Then
       'Bullet lives - monster will die in time (indirect damage from vfx)
       'convert max damage (monster hitpoints) to vfx & bullet lives on with less damage still left in it
       
       'substract damage from bullet (possibly .damage will be 0 if .damage and .hitpoints are equal!)
        .Damage = .Damage - rDam
     
       'Create 'Damaging VFX' (love that word ;)
        Create_VFX VFX_FIRE_S, TILE(TileId).X, TILE(TileId).Y, TileId, 0.1, Ammo(.AmmoType).Damage   ''rDam
        
                                
     Else
       'Bullet dies - monster lives
       'convert max bullet damage left to vfx and
       'kill bullit - monster will get damage but survive!!
       
       'Create (weak) damaging effect
        Create_VFX VFX_FIRE_S, TILE(TileId).X, TILE(TileId).Y, TileId, 0.1, Ammo(.AmmoType).Damage '.damage
                                 
       'Destroy the bullet (it's out of damage)
        .Damage = 0
     End If
                             
    'In either case, this dude's been hit by fire and will
    'start screaming and kicking and panic - without delay ;)
     If Not TILE(TileId).IsDead Then
        TILE(TileId).Program = ANIM_PANIC
        TILE(TileId).ProgramMode = 0
     End If
                             
    'And this monster will 'carry an invisible fire bullet' that
    'doesn't affect him, but sets everything that touches him on fire.

     Create_Bullet AMMO_FIREINVISIBLE, TileId

End With

End Sub


Private Function Bullet_Animation(ByVal BulId As Integer) As Boolean

Static DoFlip As Boolean
Static ffDir As Single
Static mS As Single

'## This function takes care of the bullet animation
'## And checks it's lifetime
'## If returns TRUE if for whatever reason the bullet dies

'Innocent till proven dead
 Bullet_Animation = False

 With Bullet(BulId)
    'Update sprite animation
     .Frame = .Frame + Ammo(.AmmoType).AnimSpeed
     If .Frame >= Ammo(.AmmoType).Frames Then .Frame = 0
        
     DoFlip = False
    
    'Check Angle
     If .Angle > 360 Then .Angle = .Angle - 360
     If .Angle < -360 Then .Angle = .Angle + 360
    
    'Limit movement to speed
     If .Speed > Ammo(.AmmoType).MaxSpeed Then .Speed = Ammo(.AmmoType).MaxSpeed
     If .Speed < -Ammo(.AmmoType).MaxSpeed Then .Speed = -Ammo(.AmmoType).MaxSpeed
    
    'Update position based on speed & direction
     .X = .X + Cos((.Angle - 90) / 180 * PI) * .Speed
     .Y = .Y + Sin((.Angle - 90) / 180 * PI) * .Speed
     
    'Check if bullets ricochet's  (how the f*ck do you spell this ?!)
    '13000000 comes from rgb(200,200,200)
     If Ammo(.AmmoType).DoesRicochet Then
        If GetPixel(pPass.HDC, .X, .Y) < 13000000 Then
           .Angle = .Angle + Int(Rnd(1) * 90) + 135
        End If
     End If
            
    'Check if bullet is out of screen!
     If .Y - SCROLL_OFFSET < -32 Then Bullet_Animation = True
     If .X < -16 Then Bullet_Animation = True
     If .X > VIEW.ScaleWidth + 16 Then Bullet_Animation = True
     If .Y - SCROLL_OFFSET > VIEW.ScaleHeight + 32 Then Bullet_Animation = True
                
    'Check bullet lifetime (it might die of natural causes ;)
     .LifeTime = .LifeTime - 1
     If .LifeTime <= 0 Then Bullet_Animation = True
      
 End With

End Function

Private Sub Bullet_Erase(ByVal BulId As Integer)
Static tel As Integer

'## This Kills the bullet and removes all links to it.

 Bullet(BulId).Alive = False
 
 For tel = 1 To MAX_TILES
     If TILE(tel).LastBulletId = BulId Then
        TILE(tel).LastBulletId = -1
        Exit For
     End If
 Next tel

'Check link with player
 If PLAYER.LastBulletId = BulId Then PLAYER.LastBulletId = -1
 
End Sub

Private Function Bullet_CollisionCheck(ByVal BulId As Integer, ByVal TileId As Integer) As Boolean

Static X1, X2, Y1, Y2, W, H, X1off, Y1off, X2off, Y2off As Single

'## This routine checks for collision detection between the alpha channels
'## of the given bullet and Tile, depending on what kind of bullet it is
'## (visible or not)

'Get bullet data
 With Bullet(BulId)
      X1 = .X
      Y1 = .Y

      W = Ammo(.AmmoType).Width
      H = Ammo(.AmmoType).Height
     
      X1off = Ammo(.AmmoType).imgXoff + 32 'alpha channel
      Y1off = Ammo(.AmmoType).imgYoff
 End With

'Get Tile/Monster data
 With TILE(TileId)
      X2 = .X
      Y2 = .Y
     
      X2off = MONSTER(.ObjectId).imgXoff + 32 'alpha channel
      Y2off = MONSTER(.ObjectId).imgYoff
                      
 End With
                      'Check collision if it's a visible bullet!
'reset
 Bullet_CollisionCheck = False
                       
'Let check
 If Ammo(Bullet(BulId).AmmoType).Visible Then
    If CollisionDetect(X1, Y1, pAmmo, X1off, Y1off, W, H, _
                       X2, Y2, pMonster, X2off, Y2off, _
                       pCollision) Then Bullet_CollisionCheck = True
 Else
   'Bullet is invisible so do a bounding-box collision detection
    If ((X1 + W > X2) And (X1 < X2 + W) And (Y1 + H > Y2) And (Y1 < Y2 + H)) Then Bullet_CollisionCheck = True
 End If
 
'Take care of some other administration
 If Bullet_CollisionCheck Then
 
   'Prevent this tile being hit by this bullet again.
    TILE(TileId).LastBulletId = BulId
                          
   'Incase the target dies because of colliding with this bullet, record how.
    TILE(TileId).CauseOfDeath = Ammo(Bullet(BulId).AmmoType).WeaponType

 End If
 
End Function

Private Function Bullet_CollisionPlayer(ByVal BulId As Integer) As Boolean

Static X1, X2, Y1, Y2, W1, H1, W2, H2 As Single

'## This routine checks for collision between Player and a bullet fired
'## by a monster. Note, no pixel collision!

'Get bullet data
 With Bullet(BulId)
      X1 = .X
      Y1 = .Y

      W1 = Ammo(.AmmoType).Width
      H1 = Ammo(.AmmoType).Height
     
 End With

'Get Player data
 With PLAYER
      X2 = .X
      Y2 = .Y
     
      W2 = .hW
      H2 = .hH
 End With
                      'Check collision if it's a visible bullet!
'reset
 Bullet_CollisionPlayer = False
                       
'Let's do a bounding box collision check
 If ((X1 + W2 > X2) And (X1 < X2 + W2) And (Y1 + H2 > Y2) And (Y1 < Y2 + H2)) Then Bullet_CollisionPlayer = True
 
 
'Take care of some other administration
 If Bullet_CollisionPlayer Then
 
   'Prevent this tile being hit by this bullet again.
    PLAYER.LastBulletId = BulId

 End If
 
End Function


Private Sub Bullet_GroupKillBonus(ByVal BulId As Integer)

'## This routine is called when bullet hits some monster
'## it also checks if the bullet has killed multiple victims,
'## if so, a special bonus is awarded. sounds sick heh?!

'Add another kill-score to this bullet for possible player bonus later
 Bullet(BulId).NumKills = Bullet(BulId).NumKills + 1

With Bullet(BulId)
    '5 kills is 250 bonus to score
     If .NumKills = 5 Then
        'Add bonus score
        PLAYER.Score = PLAYER.Score + 250
        'display bonus
        Create_VFX VFX_BONUS_250, .X, .Y, 0
     End If
                             
    '10 or more kills is another 500 bonus to score!
     If .NumKills >= 10 Then
       'add bonus score
        PLAYER.Score = PLAYER.Score + 500
       'Display bonus
        Create_VFX VFX_BONUS_500, .X, .Y, 0
       'Reset Group kill bonus posibility
        .NumKills = 0
     End If
    
End With

End Sub


Private Sub Bullet_StickToTile(ByVal BulId As Integer)

'## This routine simply sticks the bullet to it's sticky tile
               
With Bullet(BulId)
    'Stick it, if it's a sticker
     If .StickToTileId > 0 Then
        .X = TILE(.StickToTileId).X
        .Y = TILE(.StickToTileId).Y
        '.Xdir = 0
        '.Ydir = 0
     End If
End With

End Sub

Private Sub DebugPrint(ByVal Txt As String)

 lDebug.Caption = lDebug.Caption & vbCrLf & Txt

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X, Y As Integer

Select Case KeyCode
    
   'PLAYER GAME CONTROLLS
    Case vbKey1
'        PLAYER.Fire1 = True
         If EDITMODE Then
            EDITPAINT = MAP_EMPTY
            lEdit.Caption = "> EMPTY"
         End If
    Case vbKey2
'        PLAYER.Fire2 = True
         If EDITMODE Then
            EDITPAINT = MAP_ROCK
            lEdit.Caption = "> ROCK"
         End If
    Case vbKey3
         If EDITMODE Then
            EDITPAINT = MAP_STRUC
            lEdit.Caption = "> STRUCTURE"
         End If
    
    Case vbKeyM
         MP1.Stop
   
    Case vbKey0
         'save a screenshot ;)
         SavePicture VIEW.Image, App.Path & "\SShot.bmp"
   
   'DEBUG CONTROLLS
    Case vbKeyEscape 'End Game
         STOPGAME = True
    
    Case vbKeyQ      'End Game
         STOPGAME = True
    
    Case vbKeyF 'Show FPS
         lFps.Visible = Not lFps.Visible
         
         FPSmin = 100   'reset
         tFPS.Enabled = lFps.Visible
    
   'EDIT MODE
    Case vbKeyE 'Toggle Editmode
         EDITMODE = Not EDITMODE
         lEdit.Visible = EDITMODE
         lEdit.Caption = "* Map-Code EditMode *"
         SCROLL_SPEED = 0
    
    Case vbKeyX 'Clear map!
         If EDITMODE Then
            For X = 0 To 24
                For Y = 0 To Int(pBACK.ScaleHeight / 16)
                    MAP(X, Y) = MAP_EMPTY
                Next Y
             Next X
             pCode.Cls
             lEdit.Caption = "!MAP CLEARED!"
         End If
         
    Case vbKeyS 'Save New Code-Map
         pCode.Refresh
         SavePicture pCode.Image, App.Path & "\Images\Level01\LV01_Background_code.bmp"
         lEdit.Caption = "CODE-MAP SAVED"
    
    Case vbKeyA
         EDITPAINT = 11
         lEdit.Caption = "> Soldier"

    Case vbKeyB
         EDITPAINT = 12
         lEdit.Caption = "> Farmer"
    
    Case vbKeyC
         EDITPAINT = 13
         lEdit.Caption = "> Sheep"
    
    Case vbKeyD
         EDITPAINT = 14
         lEdit.Caption = "> Priest LV1"
         
         lDebug.Visible = Not lDebug.Visible
         
    Case vbKeyP
         EDITPAINT = 99
         lEdit.Caption = "Player Start Position"
         
    Case vbKeyR 'restart game
         STOPGAME = True
         PLAYER.GameOn = False
         tFPS.Enabled = False
         tKey.Enabled = False
         
        'Wait 2 seconds, alowing everything else to stop
         FPSLimit.LimitFrames 50
 

         Call Form_Load

         
End Select
    
    

End Sub

Private Sub ShowPass()

Dim X, Y As Integer
Dim Gx, Gy As Integer
Dim Col As Long

'Show Code
For Y = SCROLL_OFFSET / 16 To SCROLL_OFFSET / 16 + VIEW.ScaleHeight / 16
    For X = 0 To Int(VIEW.ScaleWidth / 16) - 1
        
        Gx = X * 16 + 2
        Gy = (Y - SCROLL_OFFSET / 16) * 16 + 2
        
        Col = RGB(0, 100, 0)
        
        If MAP(X, Y) = MAP_ROCK Then Col = RGB(200, 200, 200)
        If MAP(X, Y) = MAP_STRUC Then Col = RGB(200, 200, 0)
        If MAP(X, Y) = MAP_OBJ Then Col = RGB(200, 0, 200)
        
        Col = pCode.Point(X, Y)
        
        VIEW.Line (Gx, Gy)-(Gx + 12, Gy + 12), Col, B
    Next X
Next Y


 
End Sub


Private Sub START_Game()

'GAME STARTS
 
'Start player & IRQ's...
 PLAYER.GameOn = True
 
 tFPS.Enabled = False
 tKey.Enabled = True        'Allow A-Sync keyboard input
  
'Player the music ;)
 PlaySound SND_Music1

'Let's do it ;)
 Call IRQ
 
 Unload Me
 
End Sub






Private Sub tFPS_Timer()

'Debug FPS timer
 If FPS < FPSmin Then FPSmin = FPS
 lFps.Caption = FPS & " fps (" & FPSmin & ")"
 
'Clear
 FPS = 0

End Sub


Private Sub tKey_Timer()

'## This IRQ get's the asyn keystates for player controllllllll...
If PLAYER.GameOn Then

'Prevent re-entry
 tKey.Enabled = False 'CLI

'cursor keys increase the players speed unto a certain maximum
' If GetAsyncKeyState(&H27) < 0 Then PLAYER.Xdir = PLAYER.Xdir + PLAYER.Accel * 5 'right
' If GetAsyncKeyState(&H25) < 0 Then PLAYER.Xdir = PLAYER.Xdir - PLAYER.Accel * 5 'left
' If GetAsyncKeyState(&H28) < 0 Then PLAYER.Ydir = PLAYER.Ydir + PLAYER.Accel    'down
' If GetAsyncKeyState(&H26) < 0 Then PLAYER.Ydir = PLAYER.Ydir - PLAYER.Accel 'up
 
'When player turns left/right it also moves slightly forware to prevent it turning
'around it's perfect center - making it look very unnaturally
 If GetAsyncKeyState(&H27) < 0 Then
    PLAYER.Angle = PLAYER.Angle + 6 'right
    PLAYER.Speed = PLAYER.Speed + PLAYER.Accel * 0.2
 End If
 If GetAsyncKeyState(&H25) < 0 Then
    PLAYER.Angle = PLAYER.Angle - 6 'left
    PLAYER.Speed = PLAYER.Speed + PLAYER.Accel * 0.2
 End If
 If GetAsyncKeyState(&H28) < 0 Then PLAYER.Speed = PLAYER.Speed - PLAYER.Accel    'down
 If GetAsyncKeyState(&H26) < 0 Then PLAYER.Speed = PLAYER.Speed + PLAYER.Accel    'up
    
 
'vbKeySpace
 If GetAsyncKeyState(&H20) < 0 Then PLAYER.Weapon(1).FireButton = True
 
'vbKey1 ....
 If GetAsyncKeyState(&H31) < 0 Then PLAYER.Weapon(1).FireButton = True
 If GetAsyncKeyState(&H32) < 0 Then PLAYER.Weapon(2).FireButton = True
 If GetAsyncKeyState(&H33) < 0 Then PLAYER.Weapon(3).FireButton = True
 If GetAsyncKeyState(&H34) < 0 Then PLAYER.Weapon(4).FireButton = True
 If GetAsyncKeyState(&H35) < 0 Then PLAYER.Weapon(5).FireButton = True

'Resume
 tKey.Enabled = True 'RTS

End If

End Sub

Private Sub tstart_Timer()

'## This routine is called one to enable the form to init and then
'   start the game using this timer....

'Disbale timer
 tstart.Enabled = False
 
'Start game!
 START_Game
 
End Sub



Private Sub VIEW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Gx, Gy As Integer
Static Col As Long

If Not EDITMODE Then
Else
 'Editmode!
  Gx = Int(X / 16)
  Gy = Int((Y + SCROLL_OFFSET) / 16)
  
  If Gx < 0 Then Gx = 0
  If Gx > 24 Then Gx = 24
  If Gy < 0 Then Gy = 0
  If Gy > Int(pBACK.ScaleHeight / 16) - 1 Then Gy = Int(pBACK.ScaleHeight / 16) - 1
  
 'Draw result on MAP and CODE IMAGE !
  Select Case EDITPAINT
        Case MAP_EMPTY
             Col = RGB(0, 0, 0)
        Case MAP_ROCK
             Col = RGB(255, 255, 255)
        Case MAP_STRUC
             Col = RGB(0, 0, 255)
        Case 11     'Soldier
             Col = RGB(255, 0, 255)
        Case 12     'Farmer
             Col = RGB(0, 255, 0)
        Case 13     'Sheep
             Col = RGB(0, 255, 128)
        Case 14     'Priest
             Col = RGB(255, 0, 100)
        Case 99     'Player start position
             Col = RGB(255, 0, 0)
  End Select
  
  If Button = vbLeftButton Then
     'MAP(Gx, Gy) = EDITPAINT
     pCode.PSet (Gx, Gy), Col
  End If
  If Button = vbRightButton Then
     'MAP(Gx, Gy) = MAP_EMPTY
     pCode.PSet (Gx, Gy), Col
  End If
  
 'Refresh the code map
  pCode.Refresh
  
End If

End Sub

Private Sub Form_Load()


'Move & scale game screen
 Level.Move Form1.Left, Form1.Top, Form1.Width, Form1.Height
 
'Init all bits 'n pieces
 INIT_Monsters
 INIT_Background
 INIT_Player
 INIT_Ammo
 INIT_Sound
 INIT_Effects
 
'Show static level until game starts
'Show SOMETHING at start
 Call Update_Background
 Call Update_Player
 Call Update_Panels
 
 STOPGAME = False
 
'Hide intro screen
 Form1.Visible = False
 
'Start playin!
 tstart.Enabled = True
 
End Sub


Private Sub INIT_Background()
Dim fName As String
Dim X, Y As Integer
Dim Pix As Long
Dim nMon As Integer
Dim nRock As Integer
Dim nMon2 As Integer

'Loads & Sets up the scrolling background for this level.
 fName = App.Path & "\Images\Level01\LV01_Background.jpg"
 pBACK.Picture = LoadPicture(fName)
 
 SCROLL_SPEED = 0
 SCROLL_OFFSET = pBACK.ScaleHeight - VIEW.ScaleHeight
 SCROLL_NUMROWS = pBACK.ScaleHeight / 16
 
'Load the rock-map for player collision with map
 fName = App.Path & "\Images\Level01\LV01_Background_pass.jpg"
 pPass.Picture = LoadPicture(fName)
 
'Now load in the data-code map and get our information form there!
 fName = App.Path & "\Images\Level01\LV01_Background_code.bmp"
 pCode.Picture = LoadPicture(fName)
 
'Now analyse the 'code'
 nMon = 0
 nMon2 = 0
 nRock = 0

 For X = 0 To pCode.ScaleWidth
     For Y = 0 To pCode.ScaleHeight
     
         Pix = GetPixel(pCode.HDC, X, Y)
         Select Case Pix
            
            Case RGB(255, 0, 0)              'Player start
            
                 MAP(X, Y) = MAP_EMPTY
                 PLAYER.X = X * 16
                 PLAYER.Y = Y * 16
                
            Case RGB(255, 255, 255)          'Rock
            
                 MAP(X, Y) = MAP_ROCK
                 nRock = nRock + 1
            
            Case RGB(0, 255, 0)              'Civillian
            
                 MAP(X, Y) = MAP_EMPTY
                 nMon2 = nMon2 + 1
                 Create_Monster X * 16, Y * 16, 3
            
            Case RGB(255, 0, 255)            'Soldier
            
                 MAP(X, Y) = MAP_EMPTY
                 nMon = nMon + 1
                 Create_Monster X * 16, Y * 16, 1
            
            Case RGB(0, 255, 128)           'Sheep
                 MAP(X, Y) = MAP_EMPTY
                 nMon = nMon + 1
                 Create_Monster X * 16, Y * 16, 5
            
            Case RGB(255, 0, 100)           'Priest
                 MAP(X, Y) = MAP_EMPTY
                 nMon = nMon + 1
                 Create_Monster X * 16, Y * 16, 7
            
            Case Else
                 
                 MAP(X, Y) = MAP_EMPTY
         
         End Select
     Next Y
 Next X
 
 DebugPrint (nMon & " Solders, " & nMon2 & " Farmers")
 DebugPrint (nRock & " Rocks")
 
End Sub


Private Sub Create_Monster(ByVal X As Integer, ByVal Y As Integer, ByVal MonsterID As Integer)
Static tId As Integer

'## This routine creates a specific monster on the given grid-coordinates.
 
'Check if the house isn't full already...
 If NUM_TILES >= MAX_TILES Then Exit Sub
 
 tId = Find_Tile
 
 If tId <= 0 Then Exit Sub
 
 NUM_TILES = NUM_TILES + 1

 With TILE(tId)
     .ObjectId = MonsterID
     .Alive = True
     .IsDead = MONSTER(.ObjectId).IsDead
     .LastBulletId = -1
     .HitPoints = MONSTER(.ObjectId).HitPoints
     .XStart = Int(X / 16)   'map coordinates
     .YStart = Int(Y / 16)
     .X = X + Int(Rnd(1) * 8) - 4
     .Y = Y + Int(Rnd(1) * 8) - 4
     .Frame = Int(Rnd(1) * MONSTER(.ObjectId).Frames)
     .Program = MONSTER(.ObjectId).Program 'Set to default animation program & mode
     .ProgramMode = 0
     .Angle = Int(Rnd(1) * 359)
     .WeaponAmmo = MONSTER(.ObjectId).WeaponAmmo
     .WeaponReloadTime = 0
 End With

End Sub

Private Sub INIT_Effects()
Dim fName As String

'## Creates the default database for Visual Effects.
'  fName = App.Path & "\Images\Effects\100-Effects-Pack-01.bmp"
'  pVFX.Picture = LoadPicture(fName)

 With EFFECT(1)

    .Name = "Fire Dying Down"
    
    .Frames = 5
    .AnimSpeed = 0.5
    .Type = VFX_TYPE_SPRITE
    .Loops = 3
    
    .LockedToTile = False
    .Program = 0
    .Text = ""
    
    .imgXoff = 0
    .imgYoff = 0
    .W = 32
    .H = 32
    
 End With

 With EFFECT(2)

    .Name = "Lightning bodyshock"
    
    .Frames = 5
    .Loops = 1
    .AnimSpeed = 0.5
    .Type = VFX_TYPE_SPRITE
    
    .LockedToTile = True
    .Program = 1 'minor random factor to vfx position
    .Text = ""
    
    .imgXoff = 0
    .imgYoff = 1 * 32
    .W = 32
    .H = 32
    
 End With

 With EFFECT(3)

    .Name = "Small Fire (damage)"
    
    .Frames = 5
    .Loops = 50
    .AnimSpeed = 0.3
    .Type = VFX_TYPE_SPRITE
    
    .LockedToTile = True
    .Program = 2
    .Text = ""
    
    .imgXoff = 0
    .imgYoff = 2 * 32
    
    .W = 32
    .H = 32
    
 End With
 
 With EFFECT(4)

    .Name = "Bonus 250"
    
    .Frames = 50
    .AnimSpeed = 1
    .Loops = 0
    
    .Type = VFX_TYPE_TEXT
    
    .LockedToTile = False
    .Program = 3
    .Text = "+250"
    
    .imgXoff = 0
    .imgYoff = 0
    
    .W = VIEW.TextWidth(.Text)
    .H = VIEW.TextHeight(.Text)
    
 End With

 With EFFECT(5)

    .Name = "Bonus 500"
    
    .Frames = 50
    .AnimSpeed = 1
    .Loops = 0
    
    .Type = VFX_TYPE_TEXT
    
    .LockedToTile = False
    .Program = 3
    .Text = "+500"
    
    .imgXoff = 0
    .imgYoff = 0
    
    .W = VIEW.TextWidth(.Text)
    .H = VIEW.TextHeight(.Text)
    
 End With

End Sub

Private Sub Create_VFX(ByVal EffectID As Integer, _
                       ByVal X As Integer, ByVal Y As Integer, _
                       Optional ByVal TileId As Integer, _
                       Optional ByVal ProgVal1 As Single, _
                       Optional ByVal ProgVal2 As Single, _
                       Optional ByVal Xoff As Single, _
                       Optional ByVal Yoff As Single)

Static vId As Integer

'## This routine creates a new tile which holds a vfx.
'## Optionally this fx can be locked to another tile (for example a flame attached
'## to a running soldier).

vId = Find_Vfx

'Make sure we don't go over the limit
If vId > MAX_VFX Then vId = -1

'create jammies...
If vId > 0 Then

   With VFX(vId)
        .Alive = True
        
        .Frame = -1         'It starts at -1 because at Update_vfx it already gets increased by 1, and thus missing the first frame :(
        .Loop = EFFECT(EffectID).Loops
        
        .EffectID = EffectID
        
       'Program specific data / settings
        .ProgVal1 = ProgVal1
        .ProgVal2 = ProgVal2
        
        'Coordination
        .X = X
        .Y = Y
        .Xdir = 0
        .Ydir = 0
        .Xoff = Xoff
        .Yoff = Yoff
                
       'If tileId is < 0 then it's the player number (-1 = player1, -2 = player2 etc.)
        If EFFECT(EffectID).LockedToTile Then
           .Follow = True
           .TileId = TileId
        Else
           .Follow = False
        End If
        
          
   End With
   
   NUM_VFX = NUM_VFX + 1
   
Else
  'All VFX are used - MAX_VFX is reached, can't create effect!
  
End If


End Sub


Private Sub Update_VFX()
Static tel As Integer
Static vFound As Integer
Static vDied As Integer
Static DoDie As Boolean

'## This routine updates all running VFX

If NUM_VFX <= 0 Then Exit Sub

 vFound = 0
 vDied = 0
   
 For tel = 1 To MAX_VFX
        
     DoDie = False
          
     If VFX(tel).Alive Then
          
        vFound = vFound + 1
          
        With VFX(tel)
          
           'Update Frame  & Sprite Animation (random between 0.75 and 1.25 animspeed)
            .Frame = .Frame + EFFECT(.EffectID).AnimSpeed 'Int(Rnd(1) * (0.25 * EFFECT(.EffectID).AnimSpeed)) + EFFECT(.EffectID).AnimSpeed * 0.75
            If .Frame > EFFECT(.EffectID).Frames - 1 Then
               'Check if this is a loop animation
               If .Loop > 0 Then
                  .Loop = .Loop - 1
                  .Frame = 0
               Else
                  'This is not a loopable effect OR it has finished it's last loop
                  'This effect dies now!
                  DoDie = True
               End If
            End If
           
           'Check if this VFX is locked to a tile
            If .Follow Then
            
              'Check if it's locked to the PLAYER or a MONSTER
               If .TileId < 0 Then
                  .X = PLAYER.X
                  .Y = PLAYER.Y
               Else
                 'Check if that tile is still alive, if not - disconnect it!
                  If Not TILE(.TileId).Alive Then
                    'Disconnect it - coordinates remain the same!
                     .Follow = False
                  Else
                     'Stick this vfx to the tile
                     .X = TILE(.TileId).X
                     .Y = TILE(.TileId).Y
                  End If
               End If
            End If
               
           'Check VFX program
               '-> Move this to a seperate sub before it grows too large!
               '-> And it shouldn't be in the .follow else statement!!!
            Select Case EFFECT(.EffectID).Program
                    Case 0
                        'Default program - no special animation
                    Case 1
                        'Program 1 - small random factor to position (-3 to +3)
                        .X = .X + Rnd(1) * 2 - 1
                        .Y = .Y + Rnd(1) * 1 - 1
                    Case 2
                    
                        'Program 3 - VFX doing damage while dying out!!
                        '---------------------------------------------------
                        'This effect (used when a flame hits a monster, the
                        'monster is set alight, and keeps burning for the
                        'duration of the effect and SUSTAINING DAMAGE !!!!
                        
                        'ProgVal1 = the amount of damage it does each frame
                        'ProgVal2 = the amount of total damage the effect can deliver
                        
                       'Give a minor random shake on the effect
                        '.X = .X + Rnd(1) - 0.1
                        '.Y = .Y + Rnd(1) - 0.1
                        If .TileId > 0 Then
                          'Substract effect damage from monster
                           TILE(.TileId).HitPoints = TILE(.TileId).HitPoints - .ProgVal1
                    
                          'Check if effect is dying out
                           .ProgVal2 = .ProgVal2 - .ProgVal1
                           If .ProgVal2 < 0 Then DoDie = True
                        
                          'Check monster is dying out, if so,
                          'the effect disappears the next loop ;)
                          '(no fuel is no fire)
                           If TILE(.TileId).HitPoints <= 0 Then
                             'Disconnect
                              .Follow = False
                              .TileId = 0
                              .Loop = 1
                           End If
                        End If
                   Case 3
                        'Program 3 - Slides the vfx up slowly (used for text effects)
                        .Y = .Y - 1
    
             End Select
             
        End With
     End If
      
    'Check if this VFX has died out somehow...
     If DoDie Then
        VFX(tel).Alive = False
        VFX(tel).TileId = 0
        VFX(tel).Follow = False
        vDied = vDied + 1
     End If
              
    'Check if we've had all of 'em...
     If vFound >= NUM_VFX Then Exit For
       
 Next tel
   
'Update
 NUM_VFX = NUM_VFX - vDied

 If NUM_VFX < 0 Then MsgBox "NUM_VFX Dipped !!"

End Sub


Private Sub Draw_VFX()
Static tel As Integer
Static vFound As Integer
Static Xoff, Yoff As Integer

'## This routine Prints all VFX to screen

If NUM_VFX > 0 Then

   vFound = 0
   
   For tel = 1 To NUM_VFX
        
      'PAINT SPRITE EFFECTS
       If VFX(tel).Alive And EFFECT(VFX(tel).EffectID).Type = VFX_TYPE_SPRITE Then

          vFound = vFound + 1
         
         'Paint this bitch
          With VFX(tel)
           Xoff = EFFECT(.EffectID).imgXoff + Int(.Frame) * 64
           Yoff = EFFECT(.EffectID).imgYoff
           
           FoxAlphaMask VIEW.HDC, .X - 16 + .Xoff, .Y - 16 - SCROLL_OFFSET + .Yoff, _
                        EFFECT(.EffectID).W, EFFECT(.EffectID).H, _
                        pVFX.HDC, Xoff, Yoff, _
                        pVFX.HDC, Xoff + 32, Yoff, 0, FOX_USE_MASK
          End With
       End If
       
      'PAINT TEXT EFFECTS
       If VFX(tel).Alive And EFFECT(VFX(tel).EffectID).Type = VFX_TYPE_TEXT Then

          vFound = vFound + 1
         
         'text this...
          With VFX(tel)
           
            VIEW.CurrentX = .X - EFFECT(.EffectID).W / 2
            VIEW.CurrentY = .Y - EFFECT(.EffectID).H / 2 - SCROLL_OFFSET
            VIEW.ForeColor = RGB(255, 230, 40)
            VIEW.Print EFFECT(.EffectID).Text
           
          End With
       End If
      
      'Check if we've had all of 'em...
       If vFound >= NUM_VFX Then Exit For
       
   Next tel
End If

End Sub

Private Sub INIT_Player()

'Measure & setup the player
 With PLAYER
    'Wait till game starts
     .GameOn = False
     
     'Init Weapon Banks
     .Weapon(1).AmmoType = AMMO_LIGHTNING
     .Weapon(1).Ammo = 500
     .Weapon(1).ReloadTime = 0
     .Weapon(1).FireButton = False
     
     .Weapon(2).AmmoType = AMMO_FIREBREATH
     .Weapon(2).Ammo = 500
     .Weapon(2).ReloadTime = 0
     .Weapon(2).FireButton = False
     
     .Weapon(3).AmmoType = AMMO_LIGHTNINGBOLT
     .Weapon(3).Ammo = 500
     .Weapon(3).ReloadTime = 0
     .Weapon(3).FireButton = False
     
     .Weapon(4).AmmoType = AMMO_PLASMA
     .Weapon(4).Ammo = 500
     .Weapon(4).ReloadTime = 0
     .Weapon(4).FireButton = False
     
     .Weapon(5).AmmoType = AMMO_NONE
     .Weapon(5).Ammo = 0
     .Weapon(5).ReloadTime = 0
     .Weapon(5).FireButton = False
     
     .Speed = 0
     .MaxSpeed = 4
     .Accel = 0.6
     
     .Angle = 0
     
     .HitPoints = 180
     
     .LastBulletId = -1
     
    'PLayer dimensions
     .W = pPlayer(0).ScaleWidth
     .H = pPlayer(0).ScaleHeight
 
     .hW = .W / 2
     .hH = .H / 2
      
 End With
 
End Sub

Private Sub INIT_Monsters()

'## This sub now CREATES hard-coded monsters and INITS them
'   Whenever I make a level designer, this routine loads & inits monsters.

'Note a Monster always has an ODD ID number,
'the next ID is the same guy - only dead! (ie. a dead body).

'Create default soldier
 With MONSTER(1)
     .Id = 1
     .Name = "Soldier"
     .IsDead = False
     .HitPoints = 10
     .SpeedWalk = 0.3
     .SpeedRun = 0.75
     .Vision = 200
     .BonusDead = 10
     .Width = 32
     .Height = 32
     .Destroyable = True
     .Frames = 4
     .AnimSpeed = 0.5
     .Program = ANIM_ATTACK
     .imgXoff = 0 * 32
     .imgYoff = 0 * 32
     
     .WeaponType = AMMO_NONE
     .WeaponAmmo = 0
     
     .SndDeath = SND_DeathCry
     .SndAmbient = SND_NONE
 End With

 With MONSTER(2)
     .Id = 2
     .Name = "Soldier"
     .IsDead = True
     .HitPoints = 1
     .SpeedWalk = 0
     .SpeedRun = 0
     .Vision = 0
     .BonusDead = 1
     .Width = 32
     .Height = 32
     .Destroyable = True
     .Frames = 1
     .AnimSpeed = 0
     .Program = ANIM_DEAD
     .imgXoff = 2 * 32
     .imgYoff = 1 * 32
     
     .WeaponType = AMMO_NONE
     .WeaponAmmo = 0
     
     .SndDeath = SND_NONE 'dead men don't cry
     .SndAmbient = SND_NONE
 
 End With

'Create default Farmer / civillian
 With MONSTER(3)
     .Id = 3
     .Name = "Farmer"
     .IsDead = False
     .HitPoints = 3
     .SpeedWalk = 0.2
     .SpeedRun = 1
     .Vision = 110
     .BonusDead = 5
     .Width = 32
     .Height = 32
     .Destroyable = True
     .Frames = 4
     .AnimSpeed = 0.3
     .Program = ANIM_FEAR
     .imgXoff = 0 * 32
     .imgYoff = 2 * 32
     
     .WeaponType = AMMO_NONE
     .WeaponAmmo = 0
     
     .SndDeath = SND_DeathCry
     .SndAmbient = SND_NONE
     
 End With

 With MONSTER(4)
     .Id = 4
     .Name = "DEAD Farmer"
     .IsDead = True
     .HitPoints = 1
     .SpeedWalk = 0
     .SpeedRun = 0
     .Vision = 0
     .BonusDead = 1
     .Width = 32
     .Height = 32
     .Destroyable = True
     .Frames = 1
     .AnimSpeed = 0
     .Program = ANIM_DEAD
     .imgXoff = 2 * 32
     .imgYoff = 3 * 32
     
     .WeaponType = AMMO_NONE
     .WeaponAmmo = 0
     
     .SndDeath = SND_NONE
     .SndAmbient = SND_NONE
 End With

'Create default Farmer / civillian
 With MONSTER(5)
     .Id = 5
     .Name = "Sheep"
     .IsDead = False
     .HitPoints = 8
     .SpeedWalk = 0.2
     .SpeedRun = 1
     .Vision = 14
     .BonusDead = 2
     .Width = 32
     .Height = 32
     .Destroyable = True
     .Frames = 4
     .AnimSpeed = 0.2
     .Program = ANIM_FEAR
     .imgXoff = 0 * 32
     .imgYoff = 4 * 32
     
     .WeaponType = AMMO_NONE
     .WeaponAmmo = 0
     
     .SndDeath = SND_SheepDeathCry
     .SndAmbient = SND_SheepAmbient
     
 End With

 With MONSTER(6)
     .Id = 6
     .Name = "DEAD Sheep"
     .IsDead = True
     .HitPoints = 2
     .SpeedWalk = 0
     .SpeedRun = 0
     .Vision = 0
     .BonusDead = 1
     .Width = 32
     .Height = 32
     .Destroyable = True
     .Frames = 1
     .AnimSpeed = 0
     .Program = ANIM_DEAD
     .imgXoff = 0 * 32
     .imgYoff = 5 * 32
     
     .WeaponType = AMMO_NONE
     .WeaponAmmo = 0
     
     .SndDeath = SND_NONE
     .SndAmbient = SND_NONE
 End With

'Create default Priest
 With MONSTER(7)
     .Id = 7
     .Name = "Priest"
     .IsDead = False
     .HitPoints = 8
     .SpeedWalk = 0.2
     .SpeedRun = 1
     .Vision = 120
     .BonusDead = 15
     .Width = 32
     .Height = 32
     .Destroyable = True
     .Frames = 4
     .AnimSpeed = 0.3
     .Program = ANIM_ATTACK
     .imgXoff = 0 * 32
     .imgYoff = 6 * 32
     
     .WeaponType = AMMO_PLASMA
     .WeaponAmmo = 25
     
     .SndDeath = SND_DeathCry
     .SndAmbient = SND_NONE
     
 End With

 With MONSTER(8)
     .Id = 8
     .Name = "DEAD Priest"
     .IsDead = True
     .HitPoints = 1
     .SpeedWalk = 0
     .SpeedRun = 0
     .Vision = 0
     .BonusDead = 1
     .Width = 32
     .Height = 32
     .Destroyable = True
     .Frames = 1
     .AnimSpeed = 0
     .Program = ANIM_DEAD
     .imgXoff = 2 * 32
     .imgYoff = 6 * 32
     
     .WeaponType = AMMO_NONE
     .WeaponAmmo = 0
     
     .SndDeath = SND_NONE
     .SndAmbient = SND_NONE
 End With




 NUM_MONSTERS = 8
 
 NUM_TILES = 0
 
End Sub

Private Sub INIT_Sound()

'disabled to keep the .zip filesize down -
'send me an email if you want the soundpack!
Exit Sub

'## This routine Inits DX7 - loads up wav files (mak 64k) into memory...
 SetupDX7Sound Me
 
 CreateBuffers
 
'Load up all the wav files into the soundbuffers in memory.
' MP1.FileName = App.Path & "\Music\LowQ\001_Fluke - Absurd (Headrillaz Mix).smloop.mp3"
' MP1.Play
 
'Sound Effects
 DX7LoadSound SndWavBuffer(SND_BreathWeapon), App.Path & "\Sound\100-Weapon-FireBreath-01.wav"

'Sheep
 DX7LoadSound SndWavBuffer(SND_SheepDeathCry), App.Path & "\Sound\210-Sheep-DeathCry-01.wav"
 DX7LoadSound SndWavBuffer(SND_SheepAmbient + 0), App.Path & "\Sound\210-Sheep-Ambient-01.wav"
 DX7LoadSound SndWavBuffer(SND_SheepAmbient + 1), App.Path & "\Sound\210-Sheep-Ambient-02.wav"
 DX7LoadSound SndWavBuffer(SND_SheepAmbient + 2), App.Path & "\Sound\210-Sheep-Ambient-03.wav"
 DX7LoadSound SndWavBuffer(SND_SheepAmbient + 3), App.Path & "\Sound\210-Sheep-Ambient-04.wav"
 DX7LoadSound SndWavBuffer(SND_SheepAmbient + 4), App.Path & "\Sound\210-Sheep-Ambient-05.wav"
 DX7LoadSound SndWavBuffer(SND_SheepAmbient + 5), App.Path & "\Sound\210-Sheep-Ambient-06.wav"

'Human death cries
 DX7LoadSound SndWavBuffer(SND_DeathCry + 0), App.Path & "\Sound\200-Voice-DeathCry-01.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 1), App.Path & "\Sound\200-Voice-DeathCry-02.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 2), App.Path & "\Sound\200-Voice-DeathCry-03.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 3), App.Path & "\Sound\200-Voice-DeathCry-04.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 4), App.Path & "\Sound\200-Voice-DeathCry-05.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 5), App.Path & "\Sound\200-Voice-DeathCry-06.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 6), App.Path & "\Sound\200-Voice-DeathCry-07.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 7), App.Path & "\Sound\200-Voice-DeathCry-08.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 8), App.Path & "\Sound\200-Voice-DeathCry-09.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 9), App.Path & "\Sound\200-Voice-DeathCry-10.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 10), App.Path & "\Sound\200-Voice-DeathCry-11.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 11), App.Path & "\Sound\200-Voice-DeathCry-12.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 12), App.Path & "\Sound\200-Voice-DeathCry-13.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 13), App.Path & "\Sound\200-Voice-DeathCry-14.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 14), App.Path & "\Sound\200-Voice-DeathCry-15.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 15), App.Path & "\Sound\200-Voice-DeathCry-16.wav"
 DX7LoadSound SndWavBuffer(SND_DeathCry + 16), App.Path & "\Sound\200-Voice-DeathCry-17.wav"
End Sub


Private Sub INIT_Ammo()
Dim fName As String

'Load the latest ammo-pack
' fName = App.Path & "\Images\Ammo\101-ammo-Pack.bmp"
' pAmmo.Picture = LoadPicture(fName)

'## Creates the ammunition database
 With Ammo(AMMO_FIREBREATH)

    .Name = "Dragon Fire"
    .WeaponType = WEAPON_FIRE
    
    .Damage = 85
    
    .Frames = 3
    .AnimSpeed = 1
    .AnimLifeTime = 75
    
    .Visible = True
    .LocksToTile = False
    
    .MaxSpeed = 7
    
    .Height = 32
    .Width = 32
    .imgXoff = 0 * 32
    .imgYoff = 0 * 32
    
    .PackAmount = 250
    .ReloadTime = 20
    .DoesRicochet = False
    .DestroyedOnImpact = False
    
    .SoundFire = SND_BreathWeapon
    
 End With

 With Ammo(AMMO_LIGHTNING)

    .Name = "Lightning Bolt Small"
    .WeaponType = WEAPON_LIGHTNING
    
    .AnimLifeTime = 50
    .MaxSpeed = 10 'was 15
    .Damage = 10
    .Frames = 4
    .AnimSpeed = 1
    .Visible = True
    .LocksToTile = False
    .Height = 32
    .Width = 32
    .imgXoff = 0 * 32
    .imgYoff = 2 * 32
    .PackAmount = 200
    .ReloadTime = 5
    .DoesRicochet = True
    .DestroyedOnImpact = True
 
    .SoundFire = 0
 
 End With
 
 With Ammo(AMMO_LIGHTNINGBOLT)

    .Name = "Lightning Bolt BIG"
    .WeaponType = WEAPON_LIGHTNING
    
    .AnimLifeTime = 75
    .MaxSpeed = 10
    .Damage = 50
    .Frames = 4
    .AnimSpeed = 1
    .Visible = True
    .LocksToTile = False
    .Height = 32
    .Width = 32
    .imgXoff = 0 * 32
    .imgYoff = 1 * 32
    .PackAmount = 200
    .ReloadTime = 50
    .DoesRicochet = False
    .DestroyedOnImpact = False
 
    .SoundFire = 0
 
 End With

 With Ammo(AMMO_FIREINVISIBLE)
 
    .Name = "Invisible SMALL ammo that passes on SMALL fires"
    .WeaponType = WEAPON_FIRE
    
    .AnimLifeTime = 125
    .MaxSpeed = 0
    .Damage = 5
    .Frames = 0
    .AnimSpeed = 1      'else it will never die uit
    .Visible = False
    .LocksToTile = True
    .Height = 10
    .Width = 10
    .imgXoff = 0
    .imgYoff = 0
    .PackAmount = 250
    .ReloadTime = 3
    .DoesRicochet = False
    .DestroyedOnImpact = True
    .SoundFire = 0
    
 End With
 
 With Ammo(AMMO_PLASMA)
 
    .Name = "Plasma Bolt"
    .WeaponType = WEAPON_MAGIC
    
    .AnimLifeTime = 75
    .MaxSpeed = 6
    .Damage = 5
    .Frames = 5
    .AnimSpeed = 1      'else it will never die uit
    .Visible = True
    .LocksToTile = False
    .Height = 32
    .Width = 32
    .imgXoff = 0
    .imgYoff = 3 * 32
    .PackAmount = 5
    .ReloadTime = 75
    .DoesRicochet = False
    .DestroyedOnImpact = True
    .SoundFire = 0

 End With

 NUM_AMMO = 5
 
End Sub

Private Sub IRQ()

Do 'main unendless loop!

 FPS = FPS + 1

 If Not EDITMODE Then

   'Start game processes
    Call Update_Background
    Call Update_Bullets        'Bullets is high in the list cause it can kill monsters & players
    Call Update_Player
    Call Update_Monsters
    Call Update_VFX
    Call Update_Panels
 
   'Layers are drawn back to front: Monsters, Weapons, Player
    Call Draw_Monsters
    Call Draw_VFX
    Call Draw_Bullets
    Call Draw_Player

 Else
    'Show Code map
     Call Update_Background
     ShowPass

 End If
 
'Refresh the viewport
 VIEW.Refresh
  
'If the viewport moved, move it back!
 VIEW.Move 64, 96, VIEW.ScaleWidth, VIEW.ScaleHeight

'Limit game to certain fps
 FPSLimit.LimitFrames 20
 
Loop Until STOPGAME = True

'Stop all IRQ's
 tFPS.Enabled = False
 tKey.Enabled = False
 
 Unload Me
 
End Sub


Private Sub Update_Background()


'Scroll-Loop the backhround
 If SCROLL_OFFSET >= 0 Then
 
   'Lock the players position in relation to the viewport
    SCROLL_OFFSET = PLAYER.Y - VIEW.ScaleHeight * 0.725
    
   'Check for level boundaries
    If SCROLL_OFFSET < 0 Then
       SCROLL_OFFSET = 0
       SCROLL_SPEED = 0
    End If
    
    If SCROLL_OFFSET >= pBACK.ScaleHeight - VIEW.ScaleHeight Then
       SCROLL_OFFSET = pBACK.ScaleHeight - VIEW.ScaleHeight
       SCROLL_SPEED = 0
    End If
    
 End If
    
 SCROLL_ROW = Int(SCROLL_OFFSET / 16)
 
'Paint background
 BitBlt VIEW.HDC, 0, 0, VIEW.ScaleWidth, VIEW.ScaleHeight, pBACK.HDC, 0, SCROLL_OFFSET, vbSrcCopy
 
End Sub

Private Sub Update_Panels()

'UPDATES THE SCORE & LIFE PANEL
  
'Update score panel
 lScore(0).Caption = "Score " & Proper(PLAYER.Score, 5)
 lScore(1).Caption = lScore(0).Caption
  
'Update life-bar
 lHP.Caption = "Hitpoints: " & PLAYER.HitPoints
  
'Update the debug panel
 lDebug.Caption = "Tiles: " & Proper(NUM_TILES, 3) & "  VFX: " & Proper(NUM_VFX, 2) & "  BUL: " & Proper(NUM_BULLETS, 2)
  
End Sub

Private Sub Update_Player()

Static Gx, Gy As Integer
Static rXoff, rYoff As Integer
Static pXoff, pYoff As Integer
Static tel As Integer

'Check players vital statistics
 If PLAYER.HitPoints <= 0 Then
    MsgBox "Game over dude!"
    Unload Me
 End If

'Limit accelaration to player maximum speed
 If PLAYER.Speed > PLAYER.MaxSpeed Then PLAYER.Speed = PLAYER.MaxSpeed
 If PLAYER.Speed < -PLAYER.MaxSpeed Then PLAYER.Speed = -PLAYER.MaxSpeed
 
'Apply friction on player speed
 PLAYER.Speed = PLAYER.Speed * 0.95

'Check PLayer Angle
 If PLAYER.Angle > 360 Then PLAYER.Angle = PLAYER.Angle - 360
 If PLAYER.Angle < 360 Then PLAYER.Angle = PLAYER.Angle + 360
 
'Calculate new player position based on it's speed
'I subtract 90 to the player's angle because our 0-degrees is due North not due East!
 PLAYER.X = PLAYER.X + Cos((PLAYER.Angle - 90) / 180 * PI) * PLAYER.Speed
 PLAYER.Y = PLAYER.Y + Sin((PLAYER.Angle - 90) / 180 * PI) * PLAYER.Speed
 
'Check for viewport limitations
 If PLAYER.Y > VIEW.ScaleHeight - PLAYER.hH + SCROLL_OFFSET Then PLAYER.Y = VIEW.ScaleHeight - PLAYER.hH + SCROLL_OFFSET
 If PLAYER.Y < PLAYER.H Then PLAYER.Y = PLAYER.H
 If PLAYER.X > VIEW.ScaleWidth Then PLAYER.X = VIEW.ScaleWidth
 If PLAYER.X < 0 Then PLAYER.X = 0
 
'Check player collision detection using it's alpha channel
 If CollisionDetect(0, 0, pPlayerFin, _
    70 + 16, 16, PLAYER.W - 32, PLAYER.H - 32, _
    0, 0, _
    pPass, PLAYER.X - PLAYER.hW + 16, PLAYER.Y - PLAYER.hH + 16, pCollision) Then
    
   'Reverse direction
    PLAYER.Speed = -PLAYER.Speed * 0.4
    
    'MsgBox "Boem!"
    
    PLAYER.X = PLAYER.Xprev
    PLAYER.Y = PLAYER.Yprev
 End If

'Check player Firebuttons, ammo, reload times, AND SHOOT!
 For tel = 1 To 5
     If PLAYER.Weapon(tel).ReloadTime > 0 Then PLAYER.Weapon(tel).ReloadTime = PLAYER.Weapon(tel).ReloadTime - 1
     
     If PLAYER.Weapon(tel).FireButton Then
        PLAYER.Weapon(tel).FireButton = False
        If PLAYER.Weapon(tel).Ammo > 0 And PLAYER.Weapon(tel).ReloadTime = 0 Then
           Create_Bullet_Player tel
        End If
     End If
 Next tel
 
End Sub

Private Sub Draw_Player()

Static X, Y As Integer

'Calculate proper position
 X = PLAYER.X - 35
 Y = PLAYER.Y - 35 - SCROLL_OFFSET
 
'Rotate Player Color, Alpha and Shadow map!
 pPlayerFin.Cls
 FoxRotate pPlayerFin.HDC, 35, 35, PLAYER.W, PLAYER.H, pPlayer(0).HDC, 0, 0, PLAYER.Angle, 0, FOX_ANTI_ALIAS
 FoxRotate pPlayerFin.HDC, 35 + 70, 35, PLAYER.W, PLAYER.H, pPlayer(1).HDC, 0, 0, PLAYER.Angle, 0, FOX_ANTI_ALIAS
 FoxRotate pPlayerFin.HDC, 35 + 140, 35, PLAYER.W, PLAYER.H, pPlayer(2).HDC, 0, 0, PLAYER.Angle, 0, FOX_ANTI_ALIAS
 
'Paint The Shadow (note: the shadow image is done in color rgb(0,5,0) the key around it is (0,0,0).
 FoxAlphaBlend VIEW.HDC, X + 24, Y + 20, 70, 70, pPlayerFin.HDC, 140, 0, 120, RGB(0, 0, 0), FOX_USE_MASK
 
'Paint Player
 FoxAlphaMask VIEW.HDC, X, Y, _
              70, 70, pPlayerFin.HDC, 0, 0, _
              pPlayerFin.HDC, 70, 0, 0, FOX_USE_MASK
 
 PLAYER.Xprev = PLAYER.X
 PLAYER.Yprev = PLAYER.Y

End Sub


Private Sub Update_Monsters()

'## This updates the movement of enemies ('monsters') and their action
'   as well as taking care of dieing monsters, etc.

Static tel As Integer
Static tFound, tDied As Integer
Static Xoff, Yoff As Integer
Static DoDead As Boolean

'First check if there are any monsters left
 If NUM_MONSTERS <= 0 Then Exit Sub

 tFound = 0
 tDied = 0
 
 For tel = 1 To MAX_TILES
     If TILE(tel).Alive = True And TILE(tel).Y > SCROLL_OFFSET - 32 Then
        
        tFound = tFound + 1
       
        With TILE(tel)
        
             DoDead = False
        
            'Check if monster is damaged / still alive
             If .HitPoints <= 0 Then
                
                'Dude is dead! Kill it & score for player
                 DoDead = True
                 tDied = tDied + 1
                 
                 PLAYER.Score = PLAYER.Score + MONSTER(.ObjectId).BonusDead
                 
                'Check how he was killed and leave a body (or not)
                 Select Case .CauseOfDeath
                   'Case WEAPON_NONE
                   'Case WEAPON_BLOOD
                   'Case WEAPON_MAGIC
                    Case WEAPON_FIRE
                         
                         If .IsDead Then
                            Yoff = MONSTER(.ObjectId).imgYoff
                         Else
                            Yoff = MONSTER(.ObjectId).imgYoff + 32
                         End If
                         
                         Xoff = 2 * 64
                         
                         VFX_BurnMap .X, .Y, _
                             MONSTER(.ObjectId).Width, _
                             MONSTER(.ObjectId).Height, _
                             Int(Rnd(1) * 360), _
                             pMonster, Xoff, Yoff
                        
                   'Case WEAPON_ICE
                   'Case WEAPON_ACID
                    Case WEAPON_LIGHTNING
                        'Shock 'em!
                         Create_VFX VFX_LIGHTNING_S, .X, .Y, tel

                        'If body is already dead, just re-create it so it can be burned
                        '(which has a small random position factor-making it shake a little ;)
                         If .IsDead Then
                            Create_Monster .X, .Y, .ObjectId
                         Else
                            Create_Monster .X, .Y, .ObjectId + 1
                         End If
                         
                         
                    Case Else
                        'Note: 'normal' death has 2 possible frames, chosen randomly
                         If .IsDead Then
                            Yoff = MONSTER(.ObjectId).imgYoff
                         Else
                            Yoff = MONSTER(.ObjectId).imgYoff + 32
                         End If
                         
                         Xoff = 0 + 2 * 64

                 End Select
                             
             End If
             
            'Check if tile needs to be erased
             If DoDead Then .Alive = False
            
            'IF MONSTER IS ALIVE....
             If .Alive And Not .IsDead Then
             
               'Update tile's reload time
                If MONSTER(.ObjectId).WeaponType > AMMO_NONE Then
                   If .WeaponReloadTime > 0 Then .WeaponReloadTime = .WeaponReloadTime - 1
                End If
             
                Monster_Program tel
                   
               'Animate the sprite
                .Frame = .Frame + MONSTER(.ObjectId).AnimSpeed
                If .Frame >= MONSTER(.ObjectId).Frames Then .Frame = 0
                   
               'Check ambient sound effects
                If MONSTER(.ObjectId).SndAmbient > SND_NONE Then
                   If Rnd(1) < 0.0005 Then PlaySound MONSTER(.ObjectId).SndAmbient
                End If
             End If
             
            
        End With

        
       'Check if we've had em all...
        If tFound >= NUM_TILES Then Exit For
     End If
 Next tel

'Clean up tiles when one or more died
'(can't do this in previous loop because it kills AND creates monsters/tiles)
NUM_TILES = NUM_TILES - tDied
 

End Sub

Private Sub Monster_Program(ByVal TileId As Integer)

'## This routine pre-checks and executes a monster-animation-program.
'   This determines their behaviour, if they attack the player, run away,
'   or whatever. It also takes care of basic collision correction, etc.

'First check if the monster is visible
 If TILE(TileId).Y >= SCROLL_OFFSET - 16 Then
            
    Select Case TILE(TileId).Program ' MONSTER(TILE(TileId).ObjectId).Program
    
        Case ANIM_ATTACK          'Monster attacks when it has weapons, else it runs away
             Monster_Attack TileId
        Case ANIM_FEAR            'Monster runs away in fear
             Monster_Fear TileId
        Case ANIM_DEAD
            'errr.. he's dead so,... don't do anything!
            'Damn, easiest animation I ever codec!!! Play Dead!
        Case ANIM_PANIC
             Monster_Panic TileId
        Case ANIM_PRIEST
             Monster_Priest TileId
    End Select
 End If
 
End Sub
Private Sub Monster_Attack(ByVal TileId As Integer)

Static rS, mS As Single             'Random speed, monster speed
Static Xo, Yo, XdO, YdO, pX, pY, Xnew, Ynew As Single
Static Xok, Yok As Boolean
Static BulId As Integer
Static D As Single

'##-----------------------------------------------------------------------
'## ANIMATION:   ATTACK     (soldiers)
'##-----------------------------------------------------------------------
'## This Program makes the monster run towards the Player and attack him.
'## When the monster runs out of ammo, the enemy might panic and switch to
'## the 'FLEE' program and run away.
'##
'## MODES
'## 0 - Default, Enemy is minding his own business (wandering around as if
'##     in guard mode) until he sees the player.
'## 1 - Attack, Monster tries to run towards the player and fires ammo
'##     If enemy runs out of ammo, he switches to the Monster_Fear program

 With TILE(TileId)
 
    Xo = .X
    Yo = .Y
    
    XdO = .Xdir
    YdO = .Ydir
    
    pX = PLAYER.X
    pY = PLAYER.Y
    
   'First check in wich mode we are
    Select Case .ProgramMode
        Case 0
             
             mS = MONSTER(.ObjectId).SpeedWalk
            
            'Monster doesn't do anything, perhaps wander around in rnd direction...
             .Xdir = Rnd(1) * (0.5 * mS)
             .Ydir = Rnd(1) * (0.5 * mS)
             
            'Check distance to player
             If Dist(.X, .Y, pX, pY) <= MONSTER(.ObjectId).Vision Then
               '20% chance that monster see player - should be an Alertness property ;)
                If Rnd(1) > 0.2 Then .ProgramMode = 1
             End If
             
        Case 1
             
             mS = MONSTER(.ObjectId).SpeedRun
             
            '10% chance that monster will not change to new direction
             If Rnd(1) < 0.1 Then
                .Xdir = XdO
                .Ydir = YdO
             End If
             
            '1% chance that monster will just stop and do nothing
            ' If Rnd(1) < 0.01 Then
            '    .Xdir = 0
            '    .Ydir = 0
            '    .ProgramMode = 0
            ' End If
            
            '2% chance that monster will just flee of fear
            ' If Rnd(1) < 0.02 Then
            '    .Program = ANIM_FEAR
            '    .ProgramMode = 1
            ' End If
            
            '2% chance that monster will reverse direction
            ' If Rnd(1) < 0.02 Then
            '    .Xdir = .Xdir * -1
            '    .Ydir = .Ydir * -1
            ' End If
             
             
' If Monster sees Player it will AND monster has ammo, it will attack
' If monster sees Player and has no ammo, it stays at a safe distance
' if monster is too close to the player - it will run away from player

            'Get distance between monster & player
             D = Dist(.X, .Y, pX, pY)
             
            'If Monster is close to the player he will FIRE and Back off!
            'Store the bullet's id, so this monster is immune to his own bullet.
             If D < 100 And .WeaponAmmo > 0 Then ' Ammo(MONSTER(.ObjectId).Weapon).ReloadCurTime = 0 Then
               
               '##TODO: 100 should be weapon range!
             
               'Fire a weapon (if it has been reloaded)
                BulId = Create_Bullet_Monster(AMMO_PLASMA, TileId)
                If BulId > 0 Then
                   .LastBulletId = BulId
                End If
                
                'turn the other cheek
                .Xdir = .Xdir * -1
                .Ydir = .Ydir * -1
                
             End If
             
             If .WeaponReloadTime = 0 Then
             
               'Make monster run towards the player
                If pX < .X Then
                   .Xdir = .Xdir - 0.25 * mS
                End If
                If pX > .X Then
                   .Xdir = .Xdir + 0.25 * mS
                End If
                 
               'Limit movement to speed
                If .Xdir > mS Then .Xdir = mS
                If .Xdir < -mS Then .Xdir = -mS
                
               'vertical...
                If pY < .Y Then
                   .Ydir = .Ydir - 0.25 * mS
                End If
                If pY > .Y Then
                   .Ydir = .Ydir + 0.25 * mS
                End If
        
               'Limit movement to speed
                If .Ydir > mS Then .Ydir = mS
                If .Ydir < -mS Then .Ydir = -mS
             
             End If

    End Select
   
   'General Movement & collision checks
   '-----------------------------------
    
    Xok = True
    Yok = True
    
   'Prevent Walking out of frame
    If .X + .Xdir < 0 Then
       .X = 0
       .Xdir = .Xdir * -1
    End If
    If .X + .Xdir > VIEW.ScaleWidth Then
       .X = VIEW.ScaleWidth
       .Xdir = .Xdir * -1
    End If
    
    If .Y + .Ydir < 0 Then
       .Y = 0
       .Ydir = .Ydir * -1
    End If
    If .Y + .Ydir > pBACK.ScaleHeight Then
       .Y = pBACK.ScaleHeight
       .Ydir = .Ydir * -1
    End If
    
   'Check Horizontal collision
    If MAP(Int((.X + .Xdir) / 16), Int(.Y / 16)) = MAP_ROCK Then Xok = False
               
   'Check Vertical collision
    If MAP(Int(.X / 16), Int((.Y + .Ydir) / 16)) = MAP_ROCK Then Yok = False
   
   'Check combined collision
    If MAP(Int((.X + .Xdir) / 16), Int((.Y + .Ydir) / 16)) = MAP_ROCK Then
       Xok = False
       Yok = False
    End If
    
   'Update coordinates
    If Xok Then
       .X = .X + .Xdir
    Else
       .Xdir = 0
    End If
    
    If Yok Then
       .Y = .Y + .Ydir
    Else
       .Ydir = 0
    End If
    
    
 End With

End Sub

Private Sub Monster_Priest(ByVal TileId As Integer)

Static rS, mS As Single             'Random speed, monster speed
Static Xo, Yo, XdO, YdO, pX, pY, Xnew, Ynew As Single
Static Xok, Yok As Boolean

'##-----------------------------------------------------------------------
'## ANIMATION:   Priest         (priests ;)
'##-----------------------------------------------------------------------
'## This Program makes the priest move towards the player and fire a lightning
'## bolt at him.
'##
'## MODES
'## 0 - Default, Priest is minding his own business (wandering around as if
'##     in guard mode) until he sees the player.
'## 1 - Attack, Priest tries to run towards the player and attacks with lightning.
'##


 With TILE(TileId)
 
    Xo = .X
    Yo = .Y
    
    XdO = .Xdir
    YdO = .Ydir
    
    pX = PLAYER.X
    pY = PLAYER.Y
    
   'First check in wich mode we are
    Select Case .ProgramMode
        Case 0
             
             mS = MONSTER(.ObjectId).SpeedWalk
            
            'Monster doesn't do anything, perhaps wander around in rnd direction...
             .Xdir = Rnd(1) * (1 * mS) - 0.5 * mS
             .Ydir = Rnd(1) * (1 * mS) - 0.5 * mS
             
            'Check distance to player
             If Dist(.X, .Y, pX, pY) <= MONSTER(.ObjectId).Vision Then
               '10% chance that monster see player - should be an Alertness property ;)
                If Rnd(1) > 0.1 Then .ProgramMode = 1
             End If
             
        Case 1
             
             mS = MONSTER(.ObjectId).SpeedRun
       
            'Make monster run towards the player
             If pX < .X Then
                .Xdir = .Xdir - 0.25 * mS
             End If
             If pX > .X Then
                .Xdir = .Xdir + 0.25 * mS
             End If
             
             If .Xdir > mS Then .Xdir = mS
             If .Xdir < -mS Then .Xdir = -mS
            
            'vertical...
             If pY < .Y Then
                .Ydir = .Ydir - 0.25 * mS
             End If
             If pY > .Y Then
                .Ydir = .Ydir + 0.25 * mS
             End If
    
             If .Ydir > mS Then .Ydir = mS
             If .Ydir < -mS Then .Ydir = -mS
             
            '10% chance that monster will not change to new direction
             If Rnd(1) < 0.1 Then
                .Xdir = XdO
                .Ydir = YdO
             End If
             
            '1% chance that monster will just stop and do nothing
             If Rnd(1) < 0.01 Then
                .Xdir = 0
                .Ydir = 0
                .ProgramMode = 0
             End If
            
            '2% chance that monster will just flee of fear
             If Rnd(1) < 0.02 Then
                .Program = ANIM_FEAR
                .ProgramMode = 1
             End If
            
            '2% chance that monster will reverse direction
             If Rnd(1) < 0.02 Then
                .Xdir = .Xdir * -1
                .Ydir = .Ydir * -1
             End If
             
            'If Monster is too close to the player he will back off
             If Dist(.X, .Y, pX, pY) < 64 Then
                .Xdir = .Xdir * -1
                .Ydir = .Ydir * -1
             End If

    End Select
   
   'General Movement & collision checks
   '-----------------------------------

'##TODO: All this lot should be moved to a seperate func - since it should be the same
'   for all anims, right?!

    Xok = True
    Yok = True
    
   'Prevent Walking out of frame
    If .X + .Xdir < 0 Then
       .X = 0
       .Xdir = .Xdir * -1
    End If
    If .X + .Xdir > VIEW.ScaleWidth Then
       .X = VIEW.ScaleWidth
       .Xdir = .Xdir * -1
    End If
    
    If .Y + .Ydir < 0 Then
       .Y = 0
       .Ydir = .Ydir * -1
    End If
    If .Y + .Ydir > pBACK.ScaleHeight Then
       .Y = pBACK.ScaleHeight
       .Ydir = .Ydir * -1
    End If
    
   'Check Horizontal collision
    If MAP(Int((.X + .Xdir) / 16), Int(.Y / 16)) = MAP_ROCK Then Xok = False
               
   'Check Vertical collision
    If MAP(Int(.X / 16), Int((.Y + .Ydir) / 16)) = MAP_ROCK Then Yok = False
   
   'Check combined collision
    If MAP(Int((.X + .Xdir) / 16), Int((.Y + .Ydir) / 16)) = MAP_ROCK Then
       Xok = False
       Yok = False
    End If
    
   'Update coordinates
    If Xok Then
       .X = .X + .Xdir
    Else
       .Xdir = 0
    End If
    
    If Yok Then
       .Y = .Y + .Ydir
    Else
       .Ydir = 0
    End If
    
    
 End With

End Sub






Private Sub Monster_Fear(ByVal TileId As Integer)

Static rS, mS As Single             'Random speed, monster speed
Static Xo, Yo, XdO, YdO, pX, pY, Xnew, Ynew As Single
Static Xok, Yok As Boolean

'##-----------------------------------------------------------------------
'## ANIMATION:   FLEE
'##-----------------------------------------------------------------------
'## This Program makes the monster run for his life as soon as he spots
'## any danger. This is used for civilians (ie. a farmer) or when a
'## more brave monster suddenly turns scared (ie. because he's out of
'## ammo).
'##
'## MODES
'## 0 - Default, Monster is minding his own business (stands still ;)
'##     until he see's the player.
'## 1 - Flee, Monster tries to run away from player.
'##



 With TILE(TileId)
 
    Xo = .X
    Yo = .Y
    
    XdO = .Xdir
    YdO = .Ydir
    
    pX = PLAYER.X
    pY = PLAYER.Y
    
   'First check in wich mode we are
    Select Case .ProgramMode
        Case 0
             
             mS = MONSTER(.ObjectId).SpeedWalk
            
            'Monster doesn't do anything, perhaps wander around in rnd direction...
             .Xdir = Rnd(1) * (0.25 * mS)
             .Ydir = Rnd(1) * (0.25 * mS)
             
            'Check distance to player
             If Dist(.X, .Y, pX, pY) <= MONSTER(.ObjectId).Vision Then
               '50/50 chance that monster see player - should be an Alertness property ;)
                If Rnd(1) > 0.5 Then .ProgramMode = 1
             End If
             
        Case 1
             
             mS = MONSTER(.ObjectId).SpeedRun
       
            'Make monster walk away from the player
             If pX < .X Then
                .Xdir = .Xdir + 0.25 * mS
             End If
             If pX > .X Then
                .Xdir = .Xdir - 0.25 * mS
             End If
             
             If .Xdir > mS Then .Xdir = mS
             If .Xdir < -mS Then .Xdir = -mS
            
            'vertical...
             If pY < .Y Then
                .Ydir = .Ydir + 0.25 * mS
             End If
    
             If pY > .Y Then
                .Ydir = .Ydir - 0.25 * mS
             End If
    
             If .Ydir > mS Then .Ydir = mS
             If .Ydir < -mS Then .Ydir = -mS
             
            'Gamble on a percentage that monster WILL change direction
             If Rnd(1) < 0.9 Then
                .Xdir = XdO
                .Ydir = YdO
             End If
       

    End Select
   
   'General Movement & collision checks
   '-----------------------------------
    
    Xok = True
    Yok = True
    
   'Prevent Walking out of frame
    If .X + .Xdir < 0 Then
       .X = 0
       .Xdir = 0 '.Xdir * -1
    End If
    If .X + .Xdir > VIEW.ScaleWidth Then
       .X = VIEW.ScaleWidth
       .Xdir = .Xdir * -1
    End If
    
    If .Y + .Ydir < 0 Then
       .Y = 0
       .Ydir = 0 '.Ydir * -1
    End If
    If .Y + .Ydir > pBACK.ScaleHeight Then
       .Y = pBACK.ScaleHeight
       .Ydir = .Ydir * -1
    End If
    
   'Check Horizontal collision
    If MAP(Int((.X + .Xdir) / 16), Int(.Y / 16)) = MAP_ROCK Then Xok = False
               
   'Check Vertical collision
    If MAP(Int(.X / 16), Int((.Y + .Ydir) / 16)) = MAP_ROCK Then Yok = False
   
   'Check combined collision
    If MAP(Int((.X + .Xdir) / 16), Int((.Y + .Ydir) / 16)) = MAP_ROCK Then
       Xok = False
       Yok = False
    End If
    
   'Update coordinates
    If Xok Then
       .X = .X + .Xdir
    Else
       .Xdir = 0
    End If
    
    If Yok Then
       .Y = .Y + .Ydir
    Else
       .Ydir = 0
    End If
    
    
 End With
 
 
End Sub

Private Sub Monster_Panic(ByVal TileId As Integer)

Static rS, mS As Single             'Random speed, monster speed
Static Xo, Yo, XdO, YdO, pX, pY, Xnew, Ynew As Single
Static Xok, Yok As Boolean

'##-----------------------------------------------------------------------
'## ANIMATION:   PANIC
'##-----------------------------------------------------------------------
'## This Panic mode makes the monster run in random directions, completely
'## out of control at double speed! It's used ie. for people who are on fire.
'##
'## NOTE: Monster has no mind and runs in a random direction at
'##       a random time - REGARDLESS of the position of the PLAYER !!
'##
'##
'## MODES
'## 0 - Default, Initialisation mode. Monster get's a random direction at high speed
'## 1 - (internal mode only) Once initialised there's a small chance he changes direction.
'##


 With TILE(TileId)
 
    Xo = .X
    Yo = .Y
    
    XdO = .Xdir
    YdO = .Ydir
    
    pX = PLAYER.X
    pY = PLAYER.Y
    
   'Let's take care of this patient...
             
   'They get a Times Two bonus on their running speed when their in panic mode !!! ;)
    mS = MONSTER(.ObjectId).SpeedRun
       
    Select Case .ProgramMode
        Case 0
             'Init phase, give the dog a direction
              .Xdir = Rnd(1) * (2 * mS) - mS
              .Ydir = Rnd(1) * (2 * mS) - mS
              
             'Little Panic bonus
              .Xdir = .Xdir + 0.5 * mS
              .Ydir = .Ydir - 0.5 * mS
              
             'Switch to next phase
              .ProgramMode = 1
        Case 1
             'percent change that player runs in any direction
              If Rnd(1) < 0.01 Then .Xdir = Rnd(1) * (2 * mS) - mS
              If Rnd(1) < 0.01 Then .Ydir = Rnd(1) * (2 * mS) - mS
    End Select
    
    'Prevent that the monster was travelling at warp speed by previous action
    'ie. going faster than he's allowed to (run * panic bonus)
         
    If .Xdir > mS Then .Xdir = mS
    If .Xdir < -mS Then .Xdir = -mS
            
    If .Ydir > mS Then .Ydir = mS
    If .Ydir < -mS Then .Ydir = -mS
             
   'General Movement & collision checks
   '-----------------------------------
    
    Xok = True
    Yok = True
    
   'Prevent Walking out of frame
    If .X + .Xdir < 0 Then
       .X = 0
       .Xdir = .Xdir * -1
    End If
    If .X + .Xdir > VIEW.ScaleWidth Then
       .X = VIEW.ScaleWidth
       .Xdir = .Xdir * -1
    End If
    
    If .Y + .Ydir < 0 Then
       .Y = 0
       .Ydir = .Ydir * -1
    End If
    If .Y + .Ydir > pBACK.ScaleHeight Then
       .Y = pBACK.ScaleHeight
       .Ydir = .Ydir * -1
    End If
    
   'Check Horizontal collision
    If MAP(Int((.X + .Xdir) / 16), Int(.Y / 16)) = MAP_ROCK Then Xok = False
               
   'Check Vertical collision
    If MAP(Int(.X / 16), Int((.Y + .Ydir) / 16)) = MAP_ROCK Then Yok = False
   
   'Check combined collision
    If MAP(Int((.X + .Xdir) / 16), Int((.Y + .Ydir) / 16)) = MAP_ROCK Then
       Xok = False
       Yok = False
    End If
    
   'Update coordinates
    If Xok Then
       .X = .X + .Xdir
    Else
       .Xdir = 0
    End If
    
    If Yok Then
       .Y = .Y + .Ydir
    Else
       .Ydir = 0
    End If
    
    
 End With
 
 
End Sub


Private Sub Draw_Monsters()

Static tel, tFound As Integer
Static Xoff, Yoff As Integer

'First check if there are any monsters left
 If NUM_TILES <= 0 Then Exit Sub

 tFound = 0
 
 For tel = 1 To MAX_TILES

     If TILE(tel).Alive Then
        tFound = tFound + 1
        
       'Draw only when it's visible
        If TILE(tel).Y > SCROLL_OFFSET - 16 And TILE(tel).Y < SCROLL_OFFSET + VIEW.ScaleHeight + 16 Then
        
           Xoff = MONSTER(TILE(tel).ObjectId).imgXoff + Int(TILE(tel).Frame) * 64
           Yoff = MONSTER(TILE(tel).ObjectId).imgYoff

           BitBlt VIEW.HDC, TILE(tel).X - 16, TILE(tel).Y - 16 - SCROLL_OFFSET, _
                  MONSTER(TILE(tel).ObjectId).Width, MONSTER(TILE(tel).ObjectId).Height, _
                  pMonster.HDC, Xoff + 32, Yoff, SRCAND
                  
           BitBlt VIEW.HDC, TILE(tel).X - 16, TILE(tel).Y - 16 - SCROLL_OFFSET, _
                  MONSTER(TILE(tel).ObjectId).Width, MONSTER(TILE(tel).ObjectId).Height, _
                  pMonster.HDC, Xoff, Yoff, SRCPAINT

'           FoxAlphaMask VIEW.HDC, TILE(Tel).X - 16, TILE(Tel).Y - 16 - SCROLL_OFFSET, _
'                        MONSTER(TILE(Tel).ObjectId).Width, MONSTER(TILE(Tel).ObjectId).Height, _
'                        pMonster.HDC, Xoff, Yoff, pMonster.HDC, Xoff + 32, Yoff, , FOX_USE_MASK

        End If
     End If
    'Check if we've got 'em all...
     If tFound >= NUM_TILES Then Exit For
 Next tel
 
End Sub


Private Sub Create_Bullet_Player(ByVal WeaponNum As Integer)
Static bID As Integer

'## This sub creates a new bullet made by the player
'## NOTE that RELOAD and AMOUNT OF AMMO checks have already been done!!
 
 bID = Find_Bullet()
 
 If bID <= 0 Then
    Exit Sub       'We're out of ammo-slots OR no real weapon selected
 End If
 
'ok, let's make a bullet
 NUM_BULLETS = NUM_BULLETS + 1

 With Bullet(bID)
     .Alive = True
     .AmmoType = PLAYER.Weapon(WeaponNum).AmmoType
     
     .FiredByPlayer = True
     
     .X = PLAYER.X
     .Y = PLAYER.Y
         
     .Angle = PLAYER.Angle
         
     .Speed = Ammo(PLAYER.Weapon(WeaponNum).AmmoType).MaxSpeed
     
     'Add a little variation to it ;)
     .Xdir = .Xdir + (Rnd(1) * 3 - 1.5)
     .Ydir = .Ydir + (Rnd(1) * 3 - 1.5)
     
     .Damage = Ammo(PLAYER.Weapon(WeaponNum).AmmoType).Damage
     .Frame = 1
     .NumKills = 0
     .LifeTime = Ammo(PLAYER.Weapon(WeaponNum).AmmoType).AnimLifeTime
     .StickToTileId = -1

 End With
           
'Reset the reload clock
 PLAYER.Weapon(WeaponNum).ReloadTime = Ammo(PLAYER.Weapon(WeaponNum).AmmoType).ReloadTime
           
'Play sound effect
 PlaySound Ammo(PLAYER.Weapon(WeaponNum).AmmoType).SoundFire

End Sub

Private Sub Create_Bullet(ByVal AmmoType As Integer, Optional ByVal TargetTileId As Integer)
Static bID As Integer

 Exit Sub

'## THIS ROUTINE IS ASS AT THE MOMENT! perhaps not this one, but the one(s)
'   that call or use this one....

'## This sub creates a new bullet shot by neither player nor monsters..
 bID = Find_Bullet()
 
 If bID <= 0 Or AmmoType = AMMO_NONE Then
    Exit Sub
 End If
 
  'ok, let's make a bullet
   NUM_BULLETS = NUM_BULLETS + 1

   With Bullet(bID)
     .Alive = True
     .AmmoType = AmmoType
     .FiredByPlayer = True
     
     .X = TILE(TargetTileId).X
     .Y = TILE(TargetTileId).Y
         
     .Speed = 0
     .Angle = 0
     
     .Damage = Ammo(AmmoType).Damage
     .Frame = Int(Rnd(1) * Ammo(AmmoType).Frames)
     .NumKills = 0
     .LifeTime = Ammo(AmmoType).AnimLifeTime
     
     If Ammo(AmmoType).LocksToTile Then
        .StickToTileId = TargetTileId
     Else
        .StickToTileId = -1
     End If
     
   End With

'Play sound effect
 PlaySound Ammo(AmmoType).SoundFire

End Sub

Private Function Create_Bullet_Monster(ByVal AmmoType As Integer, ByVal TileId As Integer) As Integer
Static bID As Integer
Static dX, dY, Delta As Single

'## This sub creates a new bullet Fired by an Enemy - It returns the new bullets
'   ID so that the monster doesn't get hit by his own bullet ;)
 bID = Find_Bullet()
 
 Create_Bullet_Monster = bID
 
 If bID <= 0 Then
     Exit Function  'We're out of ammo-slots / MAX_BULLETS reached!
 End If
 
'Check if weapon has been reloaded AND (double-check) monster still has ammo left
 If TILE(TileId).WeaponReloadTime = 0 And TILE(TileId).WeaponAmmo > 0 Then  'Ammo(AmmoType).ReloadCurTime = 0 Then
 
   'ok, let's make a bullet
    NUM_BULLETS = NUM_BULLETS + 1

   'Reset bullet reload
    TILE(TileId).WeaponReloadTime = Ammo(MONSTER(TILE(TileId).ObjectId).WeaponType).ReloadTime

   'Remove a bullet from his ammo
    TILE(TileId).WeaponAmmo = TILE(TileId).WeaponAmmo - 1

    With Bullet(bID)
     .Alive = True
     .AmmoType = AmmoType
     .FiredByPlayer = False
     
     .X = TILE(TileId).X
     .Y = TILE(TileId).Y
     
     .Speed = Ammo(AmmoType).MaxSpeed
     
     .Angle = 180   'currently only shoot down!!
          
     .Damage = Ammo(AmmoType).Damage
     .Frame = 1
     .NumKills = 0
     .LifeTime = Ammo(AmmoType).AnimLifeTime
     .StickToTileId = -1
     
    End With
           
   'Play sound effect
    PlaySound Ammo(AmmoType).SoundFire
 Else
   'Weapon is still reloading...
    Create_Bullet_Monster = -1
 End If

End Function


Private Function Find_Vfx() As Integer
Static tel As Integer

'## This function scans for an empty tile slot slot and returns it's position
 Find_Vfx = -1
 
 For tel = 1 To MAX_VFX
     If Not VFX(tel).Alive Then
        Find_Vfx = tel
        Exit For
        Exit Function
     End If
 Next tel

End Function
Private Function Find_Tile() As Integer
Static tel As Integer

'## This function scans for an empty tile slot slot and returns it's position
 Find_Tile = -1
 
 For tel = 1 To MAX_TILES
     If Not TILE(tel).Alive Then
        Find_Tile = tel
        Exit For
        Exit Function
     End If
 Next tel

End Function

Private Function Find_Bullet() As Integer
Static tel As Integer

'## This function scans for 'dead bullet' slots and returns it's position
 Find_Bullet = -1
 
 For tel = 1 To MAX_BULLETS
     If Not Bullet(tel).Alive Then
        Find_Bullet = tel
        Exit For
        Exit Function
     End If
 Next tel

End Function


Private Sub Update_Bullets()
Static tel, sTel As Integer
Static bFound As Integer       'Number of live bullets found in loop
Static bDied As Integer        'Number of bullets died in loop
Static IsDead As Boolean       'Check if the bullet died
Static tFound As Integer       'Amount of objects found / checked

'## This baby updates anything to do with bullets, etc.

'Check if there's anything to do here at all ;)
 If NUM_BULLETS <= 0 Then Exit Sub
 
'init
 bDied = 0
 bFound = 0
 
'Check me bullets
 For tel = 1 To MAX_BULLETS
     IsDead = False
     
     If Bullet(tel).Alive And Bullet(tel).FiredByPlayer Then
        bFound = bFound + 1
        
        With Bullet(tel)
            
            'Check if player bullet hits any monsters
             If NUM_TILES > 0 Then
                tFound = 0
                
                Bullet_StickToTile (tel)
                
                For sTel = 1 To MAX_TILES
                
                    IsDead = False
                    
                    'Check if we got a live one
                    If TILE(sTel).Alive Then tFound = tFound + 1
                    
                    'Check if tile is alive and not hit by this bullet before,
                    'And if it's a sticker - make sure the tile doesn't get hurt by
                    'the invisible bullet stuck to it
                    If TILE(sTel).Alive And TILE(sTel).LastBulletId <> tel And sTel <> .StickToTileId Then
                       
                       If TILE(sTel).LastBulletId = tel Then MsgBox "Noway!"
                       
                      'Check if bullet collides with tile / monster
                       If Bullet_CollisionCheck(tel, sTel) Then
                          
                         'The dude's got hit,- Sound a death-cry!
                          PlaySound MONSTER(TILE(sTel).ObjectId).SndDeath

                         'Check what kind of ammo we're dealing with
                          If Ammo(.AmmoType).WeaponType = WEAPON_FIRE Then
                             Bullet_DamageFire tel, sTel
                          Else
                             Bullet_Damage tel, sTel
                          End If
                          
                         'Check if this bullet earned some bonus if the target was alive
                          If TILE(sTel).Alive And Not TILE(sTel).IsDead Then
                             Bullet_GroupKillBonus tel
                          End If
 
                         'Check if this ammo is destroyed when it hits something
                          If Ammo(.AmmoType).DestroyedOnImpact Then IsDead = True
                          
                         'Check if bullet has any power left
                          If .Damage <= 0 Then IsDead = True
                          
                       End If
                    End If
                    
                   'Check if we've had em all or the bullet died.
                    If IsDead Then sTel = MAX_TILES
                    If tFound >= NUM_TILES Then sTel = MAX_TILES
                    
                Next sTel
             Else
               'No monsters alive - so no checking necesary!
             End If
        
            'Update bullet animation if it's still alive
             If Not IsDead Then IsDead = Bullet_Animation(tel)
                
            'Check if the bullet survived, if not, kill it!
             If IsDead Then
               'Erase this bullet and all links to it
                Bullet_Erase (tel)
                bDied = bDied + 1
             End If '/if num_tiles > 0
        End With
     End If 'if bullet.Alive
     
    'Check for bullets fired by monsters...
     If Bullet(tel).Alive And Not Bullet(tel).FiredByPlayer Then
        bFound = bFound + 1
        
        IsDead = False
        
        With Bullet(tel)
            'Check if bullet is anywhere near the player and has not hit him before
             If PLAYER.LastBulletId <> tel Then
                If Bullet_CollisionPlayer(tel) Then

                  'Playsound player getting hit...
                  'Playsound bullet hit...
                  '##TODO: Make sure player is immune against fire damage!
                        
                  'Do Damage to Player
                   PLAYER.HitPoints = PLAYER.HitPoints - .Damage
                   
                  'Show some VFX that player has been hit
                   Create_VFX VFX_LIGHTNING_S, .X, .Y, -1, , , .X - PLAYER.X, .Y - PLAYER.Y
                   
                  'Give the screen a little nudge
                   VIEW.Move VIEW.Left + .Xdir / 2, VIEW.Top + .Ydir / 2, VIEW.ScaleWidth, VIEW.ScaleHeight
                   
                  'Check if this ammo is destroyed when it hits something
                   If Ammo(.AmmoType).DestroyedOnImpact Then IsDead = True
                          
                  'Check if bullet has any power left
                   If .Damage <= 0 Then IsDead = True
                End If '/collision
             End If '/LastBullet Id
        End With
       
       'Update bullet animation if it's still alive
        If Not IsDead Then IsDead = Bullet_Animation(tel)
                
       'Check if the bullet survived, if not, kill it!
        If IsDead Then
          'Erase this bullet and all links to it
           Bullet_Erase (tel)
           bDied = bDied + 1
        End If '/if num_tiles > 0

     End If
    
    'Make sure we're not looking for more bullets than necesary
     If bFound >= NUM_BULLETS Then Exit For
     
 Next tel

'Update number of active bullets
 NUM_BULLETS = NUM_BULLETS - bDied

End Sub

Private Sub Draw_Bullets()
Static tel As Integer
Static bFound As Integer       'Number of live bullets found in loop
Static X, Y As Integer
Static Xoff, Yoff As Integer   'Image offset

'## This baby draws all bullets

 bFound = 0

'Check me bullets and paint LIVE and VISIBLE ammo!
 For tel = 1 To MAX_BULLETS
     If Bullet(tel).Alive And Ammo(Bullet(tel).AmmoType).Visible Then
        bFound = bFound + 1
        
        X = Bullet(tel).X - 16
        Y = Bullet(tel).Y - 16 - SCROLL_OFFSET
        
        Xoff = Ammo(Bullet(tel).AmmoType).imgXoff + Int(Bullet(tel).Frame) * 64
        Yoff = Ammo(Bullet(tel).AmmoType).imgYoff
        
       'Draw this baby
        FoxAlphaMask VIEW.HDC, X, Y, 32, 32, pAmmo.HDC, Xoff, Yoff, pAmmo.HDC, Xoff + 32, Yoff, , FOX_USE_MASK
        
     End If
    
    'Make sure we're not looking for more bullets than necesary
     If bFound > NUM_BULLETS Then Exit For
        
 Next tel

End Sub

Private Sub PlaySound(ByRef SoundId As Integer)

'disabled to keep the .zip filesize down -
'send me an email if you want the soundpack!
Exit Sub

 Select Case SoundId
 
'   Case SND_Music1
'       'A death cry is a random one out of 11
'        PlayMusicBuffer SoundId, 75, 50, 1
    
   'Sound effects
    Case SND_BreathWeapon
         PlaySoundAnyBuffer2 SoundId, 100, 50, 0
    
    Case SND_LightningBolt
    
    Case SND_SheepDeathCry
        'play the ambient sound - just 1 deathcry is driving me crazy!
         PlaySoundAnyBuffer2 SoundId + Int(Rnd(1) * SND_NUM_SheepAmbient), 90, 50, 0
    
    Case SND_SheepAmbient
        'A sheep-ambient sound is a random one out of x
         PlaySoundAnyBuffer2 SoundId + Int(Rnd(1) * SND_NUM_SheepAmbient), 90, 50, 0
    
    Case SND_DeathCry
        'A death cry is a random one out of x
         PlaySoundAnyBuffer2 SoundId + Int(Rnd(1) * SND_NUM_DeathCry), 100, 50, 0
    
    Case Else
        'No sound
 End Select

End Sub

Private Sub VFX_BurnMap(ByVal X As Integer, ByVal Y As Integer, _
                        ByVal W As Integer, ByVal H As Integer, _
                        ByVal Angle As Integer, _
                        ByRef SrcPic As PictureBox, _
                        ByVal Xoff As Integer, ByVal Yoff As Integer)

'## This function burns a given sprite on the map background.
'## Wich will stay there until the game or level re-starts.

'Draw this baby at a random angle
 
 Select Case Int(Rnd(1) * 5) + 1
        Case 1
             FoxAlphaMask pBACK.HDC, X - W / 2, Y - W / 2, W, H, SrcPic.HDC, Xoff, Yoff, SrcPic.HDC, Xoff + 32, Yoff, , FOX_USE_MASK Or FOX_TURN_90DEG
        Case 2
             FoxAlphaMask pBACK.HDC, X - W / 2, Y - W / 2, W, H, SrcPic.HDC, Xoff, Yoff, SrcPic.HDC, Xoff + 32, Yoff, , FOX_USE_MASK Or FOX_TURN_180DEG
        Case 3
             FoxAlphaMask pBACK.HDC, X - W / 2, Y - W / 2, W, H, SrcPic.HDC, Xoff, Yoff, SrcPic.HDC, Xoff + 32, Yoff, , FOX_USE_MASK Or FOX_FLIP_X
        Case 4
             FoxAlphaMask pBACK.HDC, X - W / 2, Y - W / 2, W, H, SrcPic.HDC, Xoff, Yoff, SrcPic.HDC, Xoff + 32, Yoff, , FOX_USE_MASK Or FOX_TURN_270DEG
        Case 5
             FoxAlphaMask pBACK.HDC, X - W / 2, Y - W / 2, W, H, SrcPic.HDC, Xoff, Yoff, SrcPic.HDC, Xoff + 32, Yoff, , FOX_USE_MASK
        Case Else
            MsgBox "Not good"
 End Select

' FoxRotate pPlayerFin.HDC, 35, 35, PLAYER.W, PLAYER.H, pPlayer(0).HDC, 0, 0, PLAYER.Angle, , FOX_ANTI_ALIAS
' FoxRotate pBACK.HDC, X, Y, W, H, SrcPic.HDC, Xoff, yOFF, Angle, RGB(0, 0, 0), FOX_ANTI_ALIAS + FOX_USE_MASK
 
End Sub






