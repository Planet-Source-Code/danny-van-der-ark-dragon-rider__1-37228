Attribute VB_Name = "Defines"


'Graphics functions and constants
 Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'GDI declarations for collision detection
' Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDC As Long) As Long
' Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
' Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
 Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
' Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
' Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Sound playback ------------------------------------------------------
'Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal uFlags As Long) As Long

'LEVEL SCREEN SIZE
 Public Const LEVEL_WIDTH = 384
 Public Const LEVEL_HEIGHT = 2000

'Game Sounds
 Public Const SND_NONE = 0
 Public Const SND_Music1 = 1
 Public Const SND_Music2 = 2
 
 Public Const SND_BreathWeapon = 10
 Public Const SND_LightningBolt = 11
 
 Public Const SND_SheepDeathCry = 20
 Public Const SND_SheepAmbient = 21
 
 Public Const SND_DeathCry = 30            '30 to 49 are reserved for human death-cries!
 
 Public Const SND_NUM_SheepAmbient = 6     'current number of ambient sheep sounds :)
 Public Const SND_NUM_DeathCry = 17        'current number of death cries
 
 Public Const WAV_DEFAULT = "\Sound\Default.wav"    'Default sound to init sound buffers

'Keyboard
 Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Mathics, Sine & Cosine Table
 Public Const PI = 3.14159265358979 'Trig
 Public Const PIdiv18 = PI / 18
 Public Sine(35) As Single, CoSn(35) As Single

'Animation
 Public Const ANIM_ATTACK = 1                'Brave but foolish man, attacks player
 Public Const ANIM_FEAR = 2                  'Panic, enemy runs away
 Public Const ANIM_DEAD = 3                  'for Dead Bodies
 Public Const ANIM_PANIC = 4                 'running around like he lost his mind
 Public Const ANIM_PRIEST = 5                'Attack the player with divine magic
 
'Weapon Types
 Public Const WEAPON_NONE = 0                'None or unkown
 Public Const WEAPON_BLOOD = 1               'Blunt or sharp weapon (axe, sword, arrow)
 Public Const WEAPON_MAGIC = 2               'General Magic
 Public Const WEAPON_FIRE = 3                'Fire (normal or magical)
 Public Const WEAPON_ICE = 4                 'Ice (normal or magical)
 Public Const WEAPON_ACID = 5                'Acid
 Public Const WEAPON_LIGHTNING = 6           'Lightning (probably magical;)
 
'Ammo types
 Public Const AMMO_NONE = 0
 Public Const AMMO_FIREBREATH = 1
 Public Const AMMO_LIGHTNING = 2
 Public Const AMMO_LIGHTNINGBOLT = 3
 Public Const AMMO_FIREINVISIBLE = 4
 Public Const AMMO_PLASMA = 5
 
'VISUAL EFFECTS
 Public Const VFX_LIGHTNING_S = 2
 Public Const VFX_FIRE_M = 1
 Public Const VFX_FIRE_S = 3
 Public Const VFX_BONUS_250 = 4
 Public Const VFX_BONUS_500 = 5
  
'VFX TYPES
 Public Const VFX_TYPE_SPRITE = 1       'VFX is displayed using Sprites
 Public Const VFX_TYPE_TEXT = 2         'VFX is displayed using text
 Public Const VFX_TYPE_PARTICLE = 3     'VFX is displayed using particles
  
'Constants
 Public Const SRCAND = &H8800C6         'for masks
 Public Const SRCPAINT = &HEE0086       'to paint over masks on next blit

'App Maxes -----------------------------------------------------------
 Global Const MAX_BULLETS = 15          'Max of simultanious bullets in game
 Global Const MAX_AMMO = 10             'Max of different ammo / weapons
 Global Const MAX_MONSTERS = 11         'Max # of different monsters in database
 Global Const MAX_TILES = 250           'Max # of tile sprites in the game (monsters/characters)
 Global Const MAX_EFFECTS = 5           'Max # of pre-defined effects in database
 Global Const MAX_VFX = 50              'Max # of VFX-sprites active in the game
 Global Const MAX_SOUND_BUFFERS = 50    'Max # of sound-files loaded in memory
 Global Const MAX_PLAYBACK_BUFFERS = 50 'Max # of sound playback buffers (for simultanious playback)
 
'Publics -----------------------------------------------------------
 Public SCROLL_SPEED As Single          'Background scrolling speed.
 Public SCROLL_OFFSET As Single         'scrolling offset.
 Public SCROLL_ROW As Integer           'current top row visible.
 Public SCROLL_NUMROWS As Integer       'Length op current map
 
 Public FPS As Integer                  'Frames per second counter
 Public FPSmin As Integer               'Lowest FPS measured

 Public FPSLimit As New clsFrameLimiter 'Geoffrey Hazen's Class to limit the FPS
 
 Public STOPGAME As Boolean
 
 Public EDITMODE As Boolean
 Public EDITPAINT As Integer
 
 Public MAP(24, 256)                 'Maximum map size (384 x 4096 pixels)
 
 Global Const MAP_EMPTY = 0          'Nothing in particular
 Global Const MAP_ROCK = 1           'Undestroyable, unmovable obstacle
 Global Const MAP_STRUC = 2          'Destroyable, unmovable obstacle
 Global Const MAP_OBJ = 3            'Destroyable, movable obstacle

'Types -----------------------------------------------------------


'Note: Effect describes default data for a certain pre-defined visual effect.
'      Like an explosion, a fire, etc.
'      the actual entity is stored within a Tile
'Note: VFX is the current database of active Effects in the Game.
'
Type DR_EFFECT

     Name As String                 'Internal easy name for fx.
     
    'Effects Data
     Frames As Integer              'Duration
     Loops As Integer               'Number of (sprite animation) loops
     AnimSpeed As Single            'Sprite Anim speed (frame increase)
     
     LockedToTile As Boolean        'This effect is (position-)locked to a tile
     
     Text As String                 'Text (if it's a text-effect)
     Program As Integer             'Effects sub-program
     Type As Integer                'constant defining type of vfx (sprite, particles, text, etc.)
     
    'Pixel data
     W As Integer
     H As Integer
     
     imgXoff As Integer             'Offset in pVFX picture box
     imgYoff As Integer
     
End Type

Global EFFECT(MAX_EFFECTS) As DR_EFFECT

'Note: Effect describes default data for a certain pre-defined visual effect.
'      Like an explosion, a fire, etc.
'      the actual entity is stored within a Tile
'Note: VFX is the current database of active Effects in the Game.

Type DR_VFX
     EffectID As Integer            'Pointer to Effect
     
     Alive As Boolean               'VFX still active
     
     Frame As Single                'Current sprite-anim frame
     Loop As Integer                'Current loop

    'Behaviour
     Follow As Boolean              'Follow a Tile (monster) or stationairy?
     TileId As Integer              'Id of tile to follow
     
    'Program data
     ProgVal1 As Single            'Data tags for program specific data.
     ProgVal2 As Single

    'Pixel data
     X As Single
     Y As Single
     Xdir As Single                'Motion direction, if any.
     Ydir As Single
     Xoff As Single                'Offset to position
     Yoff As Single

End Type

Global VFX(MAX_VFX) As DR_VFX
Global NUM_VFX As Integer

Type DR_PlayerWeapon

     AmmoType As Integer
     Ammo As Integer
     ReloadTime As Integer
     FireButton As Boolean

End Type

Type DR_Player

     Name As String                 'Player Name
     GameOn As Boolean              'Start game (on/off)

     X As Integer                   'Absolute XY Position
     Y As Integer
     
     Speed As Single                'Current cruising speed
     MaxSpeed As Single             'Maximum fly speed
     
     Accel As Single                'Keyboard acceleration increase
     
     Xdir As Single                 'XY Direction / speed
     Ydir As Single
     
     Angle As Double                'Rotation Angle
     
     Xprev As Single                'prev coordinates in case player get's stuck
     Yprev As Single
     
     W As Integer                   'Player width / height
     H As Integer
     
     hW As Integer                  'Player Halve width / Halve height
     hH As Integer
     
    'Gameplay data
     Score As Long                  'Player Score
     HitPoints As Integer           'Player hitpoints (0= dead!)
     LastBulletId As Integer        'Id of bullet that last hit the player
     
    'Weapon data
     Weapon(5) As DR_PlayerWeapon
     
End Type

Global PLAYER As DR_Player

Type DR_Bullet

    AmmoType As Integer         'Ammo Type / ID
    
    Alive As Boolean            'Is Bullet 'alive' ?!
    
    X As Single                 'Final XY position
    Y As Single
    
    Angle As Single             'Bullet's direction
    Speed As Single             'Bullet's current speed
    
    Xdir As Single              'XY Direction / speed
    Ydir As Single
    
    Damage As Single            'Amount of damage applied / still left
    
    NumKills As Integer         'Number of kills by this bullet (of bonus')
   
   'Behaviour data
   'ReloadTime As Integer       '0 if ready to fire, >0 when recharging
    Frame As Single             'sprite animation frame counter
    
    FiredByPlayer As Boolean     'Bullet fired by Player or not (=by Monster)
    
    LifeTime As Integer         'Bullet lifetime
    
    StickToTileId As Integer    'Id of tile to which this bullet sticks to (like fire)

End Type

Global Bullet(MAX_BULLETS) As DR_Bullet
Global NUM_BULLETS As Integer    'Current number of bullets active in the game

Type DR_Ammo

    Name As String                'Name of ammo
    WeaponType As Integer         'Weapon type (acid, fire, blood, etc.)
    
    imgXoff As Integer            'XY offset in ammo image
    imgYoff As Integer
    
    Frames As Integer             'Number of animated frames
    AnimSpeed As Single           'Ammo Sprite Animaition speed
    
    MaxSpeed As Single               'Flying speed
    
   '##TODO: Range As Integer            'Range of ammo
    
    Visible As Boolean            'Is the bullet visible ?!
    
    LocksToTile As Boolean        'Does this bullet lock to the tile/monster it hits ?!
    
   'Dimensions
    Width As Integer              'Sprite Dimensions
    Height As Integer
    
   'General score & options
    Damage As Single              'Damage performed
    
    ReloadTime As Integer         'Frames to pause before firing again
    
    DoesRicochet As Boolean       'Does the bullet ricochet (like lightning-or a rubber ball?!;)
    DestroyedOnImpact As Boolean  'Destroys bullet on impact
    
   'CanHitMonster As Boolean      'Can it hit creatures
   'CanHitStructure As Boolean    'Can it hit structures
   'MonsterBonus As Integer       'Bonus score when monster hit with this ammo
   'StructureBonus As Integer     'Bonus score when structure hit with this ammo
    PackAmount As Integer         'Amount of bullets default with ammo-pack
   
   'Animation Data
   'AnimType As Integer           'Animation program / flight path type
   'AnimNumber As Integer         'Number of sub-bullets / particles
    AnimLifeTime As Integer       'Lifetime (or -1 = forever / untill destroyed)
    
   'Audio data
    SoundFire As Integer          'Sound when weapon fired / first used
    SoundImpact As Integer        'Sound when it hits something

End Type

Global Ammo(MAX_AMMO) As DR_Ammo    'Ammunition database
Global NUM_AMMO As Integer          'Current number of different ammo in database

'##-----------------------------------------------------------------------------------------------------

Type DR_Tile

    ObjectId As Integer           'Object Id
    Alive As Boolean              'Object still alive / active?
    IsDead As Boolean             'Object is Dead (meaning it is .Alive because it's
                                  'still in the game, but not actively - now I find
                                  'out .Alive should have been called .Enabled, but
                                  'don't feel like renaming all that cr*p...
    
    LastBulletId As Integer       'To prevent being hit by the same bullet more than 1 time
    
    IsVFX As Boolean              'Tile can be a VFX or a Monster
    VFXid As Integer              'Vfx Id
    
    XStart As Integer             'Grid Xpos start position
    YStart As Integer             'Grid Ypos
    
    X As Single                   'Pixel Xpos current position
    Y As Single                   'Pixel YPos
    
    Xdir As Single                'XY Direction / speed
    Ydir As Single
    
    Angle As Single               'orientation Angle
    
    HitPoints As Single           'Current amount of hitpoints
    CauseOfDeath As Integer       'Weapon Code which identifies How the patient died.
    
    Frame As Single               'Current sprite animation frame
    
    Program As Integer            'Current behaviour program running
    ProgramMode As Integer        'program sub-mode
    
   'Weapon status
    WeaponReloadTime As Integer   '0 if ready to fire, >0 when recharging
    WeaponAmmo As Integer         'Current number of bullets
    
End Type

Global TILE(MAX_TILES) As DR_Tile 'Object database
Global NUM_TILES As Integer       'Current number of active tiles

'DR Object describes all gameplay details of an object that can be placed on a
'tile. This includes score, animation & fx, sound effects, etc.

Type DR_Object
    Id As Integer           'Identifier
    Name As String          'Internal name
    
    IsDead As Boolean       'For 'Dead' object (ie. sheep).
   
   'Visuals
    Frames As Integer       'Number of animated frames
    AnimSpeed As Single     'Sprite Animation Speed
    Program As Integer      'Default / Initial Behaviour program
    
    WeaponType As Integer   'Weapon carried
    WeaponAmmo As Integer   'Default Number of bullets
    
   'Behaviour
    SpeedWalk As Single     'Normal movement speed
    SpeedRun As Single      'Running movement speed
    
    HitPoints As Single     'Amount of hitpoints, 0=destroyed.
    Vision As Integer       'determines how far Monsters' can see (In pixels)
    
   'Scores
    BonusHit As Integer     'Score on succesfull hit
    BonusDead As Integer    'Score when destroyed
    
   'Tech Data
    Destroyable As Boolean  'Can be destroyed or not
    
    Width As Integer        'Sprite Dimensions
    Height As Integer
    imgXoff As Integer      'XY offset in ammo image
    imgYoff As Integer
    
   'Passable As Boolean     'Can one walk over this tile
   'Flyable As Boolean      'Can one fly over this tile
   'Invisible As Boolean    'Visible or not.
   'Sound effects
    SndDeath As Integer     'Monsters Deathcry
    SndAmbient As Integer   'Ambient sound (random interval)
   
   'Visual effects
   
   'Some objects will be passable when destroyed.
   'Each situation has a sound effect.
   'Each situation has a visual effect / animation.

End Type


Global MONSTER(MAX_MONSTERS) As DR_Object
Global NUM_MONSTERS As Integer      'Current number of active monsters


Public Function Proper(ByVal Number As Single, ByVal digits As Integer) As String

'##-----------------------------------------------------------------------------------------------------
'## This function returns a string holding the given number preprened with zeroes
'##-----------------------------------------------------------------------------------------------------

Dim sstr As String

sstr = Trim(Str(Number))

If Len(sstr) >= digits Then
   'if the string is bigger than the given number of digits...
   Proper = sstr
Else
   Proper = String(digits - Len(sstr), "0") & sstr
End If


End Function

Function CollisionDetect(ByVal X1 As Integer, ByVal Y1 As Integer, _
                         picMask As PictureBox, _
                         ByVal Xoff As Integer, ByVal Yoff As Integer, _
                         ByVal Width As Integer, ByVal Height As Integer, _
                         ByVal X2 As Integer, ByVal Y2 As Integer, _
                         picMask1 As PictureBox, _
                         ByVal mXoff As Integer, ByVal mYoff As Integer, _
                         picBlank As PictureBox) As Boolean
                         
'============================================================
'== Desciption
'==
'== Name    : CollisionDetect
'==
'== Author  : Richard Lowe * NOTE: now Heavily adjusted & optimized
'== Date    : July 99
'== Contact : riklowe@hotmail.com
'==
'== Inputs
'== X1       X position in pixels of the first sprite mask
'== Y1       Y position in pixels of the first sprite mask
'== picMask  Picturebox Object of the first sprite mask
'== Xoff     Internal X Offset in picMask
'== Yoff     Internal Y Offset in picMask
'== Width    Width of picMask
'== Height   Height of picMask
'== X2       X position in pixels of the second sprite mask
'== Y2       Y position in pixels of the second sprite mask
'== picMask  Picturebox Object of the second sprite mask
'== picBlank Picturebox Object of a blank sprite
'==
'== Returns
'== TRUE     If The pixels of sprites intersect
'== FALSE    If The pixels of sprites do not intersect
'==
'== Notes
'== Remove or comment out the section of code marked *** in a real program
'== It is only included here to dislay the contents of the memory DC
'==
'============================================================
    
'------------------------------------------------------------
'This section of code calculates the overlapping mask section
'size, and defines the X and Y coordinates to be used to copy
'from each of the sprites into the memory DC.
'
'These calcs have to take into account the orientation of the
'two sprites
'------------------------------------------------------------
'Dim iMaskWidth, iM1SrcX, im2srcx, im1srcy, im2srcy As Integer
'Dim iStartBlankWidth, iStartBlankHeight As Integer
'Dim iDestX, iDestY, iMaskHeight As Integer
'Dim hMemDC, hNewBMP, hPrevBMP, tmpobj As Long
'Dim C, R As Integer
'Dim blnCollision As Boolean

'Innocent till proven guilty
 CollisionDetect = False

'Extent collision check first (cheap)
 If Not ((X1 + Width > X2) And (X1 < X2 + Width) And (Y1 + Height > Y2) And (Y1 < Y2 + Height)) Then Exit Function

    If X1 <= X2 Then
        imaskwidth = X1 + Width - X2
        iM1SrcX = Width - imaskwidth
        im2srcx = 0
        iDestX = 0
        iStartBlankWidth = imaskwidth
    Else
        imaskwidth = X2 + Width - X1
        iM1SrcX = 0
        im2srcx = Width - imaskwidth
        iDestX = 0
        iStartBlankWidth = imaskwidth
    End If
    
    If Y1 <= Y2 Then
        imaskheight = Y1 + Height - Y2
        im1srcy = Height - imaskheight
        im2srcy = 0
        iDesty = 0
        iStartBlankHeight = imaskheight
    Else
        imaskheight = Y2 + Height - Y1
        im1srcy = 0
        im2srcy = Height - imaskheight
        iDesty = 0
        iStartBlankHeight = imaskheight
    End If
    
'------------------------------------------------------------
'draw the two sprite in the collision box.
'------------------------------------------------------------
    
   'Change 1st vbsrcpaint with vbnotsrccopy adjusted because one of the alpha's is negative !
   'this also clears the picblank box - AND I don't need to DC create a dozen other maps and stuff..
    BitBlt picBlank.HDC, iDestX, iDesty, imaskwidth, imaskheight, picMask.HDC, iM1SrcX + Xoff, im1srcy + Yoff, vbNotSrcCopy
    BitBlt picBlank.HDC, iDestX, iDesty, imaskwidth, imaskheight, picMask1.HDC, im2srcx + mXoff, im2srcy + mYoff, vbSrcPaint
    
   'For debug only!
    picBlank.Refresh
    
    
'------------------------------------------------------------
'Examine the memory DC, and see if it contains any non white
'pixels. If so, set Collision = true and exit
'------------------------------------------------------------
    
    'To account for minor (compression) artifacts in color/alpha channels:
    'rgb(255,255,255) is ok (value = 16777316)
    'rgb(200,200,200) or lower is wrong (value ~13000000)
    
    For C = 0 To imaskheight - 1
        For r = 0 To imaskwidth - 1
            If GetPixel(picBlank.HDC, r, C) < 13000000 Then
                CollisionDetect = True
                Exit For
            Else
            End If
        Next
        
        If CollisionDetect = True Then Exit For
    Next

'------------------------------------------------------------

End Function

Public Function Dist(ByVal X1 As Single, ByVal Y1 As Single, _
                      ByVal X2 As Single, ByVal Y2 As Single)

 Dist = Sqr((X1 - X2) * (X1 - X2) + (Y1 - Y2) * (Y1 - Y2))


End Function


' Builds a SINE / COSINE TABLE!
Private Sub Math_BTT()
    For i = 0 To 35
        Sine(i) = Sin(i * PIdiv18)
        CoSn(i) = Cos(i * PIdiv18)
    Next
End Sub





