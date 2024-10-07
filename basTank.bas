Attribute VB_Name = "basTank"
Option Explicit

' Saved Information Settings

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' Game Color Settings
Public SkyColor As Long
Public SkyChange As Long
Public Sky As Long
Public TerrainColor As Long
Public TerrainChange As Long
Public Terrain As Long
Public BulletColor As Long
Public BulletChange As Long
Public Bullet As Long
Public RandomColor As Boolean

' Game Draw Settings
Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Type PointAPI
    X As Double
    Y As Double
End Type

Public Type cColor
    R As Integer
    G As Integer
    B As Integer
End Type

Public LineShiftMin As Integer
Public LineLenMin As Integer
Public LineVerMin As Integer
Public LineShiftMax As Integer
Public LineLenMax As Integer
Public LineVerMax As Integer
Public LineShift As Integer
Public LineLen As Integer
Public LineVer As Integer
Public RandomTerrain As Boolean

Global Const TANK_WIDTH = 16
Global Const TANK_OFSET = 5
Global Const DRAWTANK_WIDTH = 28
Global Const PAINT_TAG = "/p"

' Mathematical Settings
Global Const PI = 3.14159265358979
Global Const START_POWER = 300
Global Const NOSE_LEN = 11

Type TankAim
    Direction As Integer
    Power As Integer
End Type

' Form Position Settings
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

' Game Progress Settings
Public GameCondition As Integer

Global Const MSG_INACTIVE = "Press F2 to begin a new game"
Global Const MSG_PAUSED = "Paused"

Public IsInGame As Boolean

' Border Settings
Private Const BK_PICT = "CAMO.JFB"
Global Const CAMO_TAG = "/c"
Global Const BK_WIDE = 300
Global Const BK_HIGH = 162

' Runtime Settings
Public Players() As PlayerInfo
Public DieOrder As String
Public TankOrder As String
Public TankNumber As Integer
Public TankHasWon As Boolean

Type PlayerInfo
    PlayerName As String
    PlayerScore As Long
    TankAim As TankAim
    TankPos As PointAPI
    TankColor As Long
    TankState As Integer
End Type

Global Const STATE_ALIVE = 0
Global Const STATE_DYING = 2
Global Const STATE_DEAD = 1
Global Const STATE_GONE = 3

Global Const PNT_HITTANK = 10
Global Const PNT_ALIVE = 50
Global Const PNT_SUICIDE = -50
Global Const MIN_SCORE = -1000
Global Const MAX_PLAYERS = 8
Global Const TX_PLAYER = "Player "
Global Const TX_COUNT = " Players"

'Statusbar Settings
Global Const MSG_NOGAME = "No game currently open"
Global Const MSG_PLAYER_TURN = "'s Turn"
Global Const MSG_HIT = " scored against "
Global Const MSG_HITRUIN = " hit the runis of "
Global Const MSG_SUICIDE = " just commited suicide! (-50 points)"
Global Const MSG_TERRAIN = " hit the terrain."
Global Const MSG_ALIVE = " won the round!!! (50 Points)"
Global Const MSG_FIRE = " just fired a shot"
Global Const MSG_OUTBOUND = " fired out of bounds"

' Sound Settings
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public FireSound() As String
Public TerrainSound() As String
Public ScoreSound() As String
Public LaunchSound() As String

Private Const MSG_SNDERR = "Tank '99 was unable to load all necessary sounds. Either sounds are missing or the sound folder had been moved. Sounds will not work properly until this is fixed."

Private Const FL_FIRE = "FIRE*.WAV"
Private Const FL_TERRAIN = "BLAST*.WAV"
Private Const FL_SCORE = "SCORE*.WAV"
Private Const FL_LAUNCH = "LAUNCH*.WAV"
Private Const FOLDER_SOUND = "SOUND\"
Global Const WAV_NOTHING = "END"
Global Const WAV_NEGATIVE = "NEGATIVE"
Global Const WAV_SPLASH = "SPLASH"
Global Const WAV_FORMLOAD = "FORMLOAD"
Global Const WAV_CLICK = "CLICK"
Global Const WAV_TICK = "TICK"
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

' Various Misc. Game Settings
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private AppPath As String

Private Const MSG_APPOPEN = "Only one instance of Tank '99 can run at a time."
Private Const ILLEGAL_OBJECTS = "MenuCommonDialog"

Public Sub PlaySound(ByVal WAVName As String, Flags As Long)
Dim rc As Long
    WAVName = AppPath & FOLDER_SOUND & WAVName & ".wav"
    If Dir(WAVName) = "" Then SoundErr: Exit Sub
    rc = sndPlaySound(WAVName, Flags)
    If InStr(1, WAVName, WAV_NOTHING) > 0 Then Exit Sub
    If Flags = 17 Then Exit Sub
    If rc = 0 And Len(WAVName) > 0 Then SoundErr
End Sub

Public Sub ClickCtl(Optional MkSound As Boolean = True, Optional Flags As Long = SND_ASYNC Or SND_NOSTOP)
    If MkSound Then PlaySound WAV_CLICK, Flags
End Sub

Public Function Cosine(ByVal i As Double) As Double
    Cosine = Cos(i * (PI / 180))
End Function

Public Function Sine(ByVal i As Double) As Double
    Sine = Sin(i * (PI / 180))
End Function

Public Sub Wait(ByVal Seconds As Double)
Dim Start As Double, TotalTime As Double, Finish As Double
    Start = Timer
    Do While Timer < Start + Seconds
        DoEvents
    Loop
    Finish = Timer
    TotalTime = Finish - Start
End Sub

Public Sub CheckPlayerNames()
Dim i As Integer
    For i = LBound(Players) To UBound(Players)
        If Not Len(Players(i).PlayerName) > 0 Then
            Players(i).PlayerName = TX_PLAYER & i + 1
        End If
    Next i
End Sub

Public Function PlayerAlive(PlayerNum As Integer) As Boolean
    PlayerAlive = InStr(1, TankOrder, Str$(PlayerNum)) > 0
End Function

Sub Main()
Dim i As Integer
    If App.PrevInstance Then
        MsgBox MSG_APPOPEN, vbCritical
        End
    End If
    AppPath = App.Path
    If Not Right(AppPath, 1) = "\" Then AppPath = AppPath & "\"
    GetSounds
    ShowCurs False
    PlaySound WAV_SPLASH, SND_ASYNC Or SND_LOOP
    Wait 2
    frmSplash.Show vbModal
    GetSettings
    Wait 0.5
    ShowCurs True
    frmSettings.Show vbModal
    frmTank.Show vbModal
    End
End Sub

Public Function ColorBetween(StartCol As Long, EndCol As Long, Fraction As Currency) As Long
Dim ResultColor As cColor, Col1 As cColor, Col2 As cColor
    Col1 = ColorRGB(StartCol)
    Col2 = ColorRGB(EndCol)
    ResultColor.R = Fraction * (Col1.R - Col2.R) + Col2.R
    ResultColor.G = Fraction * (Col1.G - Col2.G) + Col2.G
    ResultColor.B = Fraction * (Col1.B - Col2.B) + Col2.B
    ColorBetween = RGB(ResultColor.R, ResultColor.G, ResultColor.B)
End Function

Public Function ColorRGB(ByVal l As Long) As cColor
    ColorRGB.R = l Mod 256
    ColorRGB.G = ((l And &HFF00FF00) / 256&)
    ColorRGB.B = (l And &HFF0000) / (256& * 256&)
End Function

Public Sub NextTurn(Optional Reset As Boolean)
Static TankPos As Integer
    If Len(ReplaceChars(TankOrder, " ")) <= 1 Then Exit Sub
    If Reset Then TankPos = -1
    Do
        TankPos = TankPos + 1
        If TankPos = Len(TankOrder) Then TankPos = 0
        If TankPos > Len(TankOrder) Then TankPos = 1
    Loop Until GetTurn(TankPos) > -1
    TankNumber = GetTurn(TankPos)
End Sub

Public Function GetTurn(TurnNum As Integer) As Integer
Dim TmpTurn As String
    TmpTurn = Mid(TankOrder, TurnNum + 1, 1)
    GetTurn = Val(TmpTurn)
    If TmpTurn = " " Then GetTurn = -1
End Function

Public Sub TankDied(TankNum As Integer)
    TankOrder = ReplaceChars(TankOrder, Trim$(Str$(TankNum)), " ")
    If Not InStr(1, DieOrder, Trim$(Str$(TankNum))) > 0 Then DieOrder = DieOrder & Trim$(Str$(TankNum))
End Sub

Public Function ReplaceChars(ByVal Chars As String, Optional ByVal ReplaceChr As String, Optional ByVal ReplaceWith As String) As String
Dim ChrCnt As Long
    If ReplaceChr = "" Then ReplaceChr = " "
    ChrCnt = 1
    Do
        ChrCnt = InStr(ChrCnt, Chars, ReplaceChr)
        If ChrCnt = 0 Then Exit Do
        Chars = Left$(Chars, ChrCnt - 1) & ReplaceWith & Right(Chars, Len(Chars) + 1 - Len(ReplaceChr) - ChrCnt)
        ChrCnt = ChrCnt + Len(ReplaceWith)
    Loop
    ReplaceChars = Chars
End Function

Public Function GetProfileString(Section As String, Key As String, Default As String) As String
Dim rc As Long
    GetProfileString = Space$(100)
    rc = GetPrivateProfileString(Section, Key, Default, GetProfileString, Len(GetProfileString), AppPath & "DIGITANK.INI")
    GetProfileString = Trim$(GetProfileString)
    GetProfileString = Left(GetProfileString, Len(GetProfileString) - 1)
End Function

Public Sub WriteProfileString(Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, AppPath & "DIGITANK.INI"
End Sub

Public Sub GetSettings()
Dim i As Integer
    ReDim Players(0 To Val(GetProfileString("Players", "Count", "7")))
    For i = LBound(Players) To UBound(Players)
        Players(i).PlayerName = GetProfileString("Players", "PlayerName" & i, TX_PLAYER & i + 1)
        Players(i).TankColor = CLng(GetProfileString("Players", "PlayerColor" & i, CStr(QBColor(i + 1))))
    Next i
    RandomTerrain = CBool(GetProfileString("TerrainSettings", "RandomTerrain", "True"))
    LineShiftMin = CInt(GetProfileString("TerrainSettings", "MinShift", "10"))
    LineShiftMax = CInt(GetProfileString("TerrainSettings", "MaxShift", "120"))
    LineLenMin = CInt(GetProfileString("TerrainSettings", "MinLen", "5"))
    LineLenMax = CInt(GetProfileString("TerrainSettings", "MaxLen", "250"))
    LineVerMin = CInt(GetProfileString("TerrainSettings", "MinVer", "200"))
    LineVerMax = CInt(GetProfileString("TerrainSettings", "MaxVer", "330"))
    RandomColor = CBool(GetProfileString("ColorSettings", "RandomColors", "True"))
    SkyColor = CLng(GetProfileString("ColorSettings", "SkyMin", "16776960"))
    SkyChange = CLng(GetProfileString("ColorSettings", "SkyMax", "11513600"))
    TerrainColor = CLng(GetProfileString("ColorSettings", "TerrainMin", "65280"))
    TerrainChange = CLng(GetProfileString("ColorSettings", "TerrainMax", "44800"))
    BulletColor = CLng(GetProfileString("ColorSettings", "BulletMax", "32255"))
    BulletChange = CLng(GetProfileString("ColorSettings", "BulletMin", "57855"))
End Sub

Public Sub SaveSettings()
Dim i As Integer
    WriteProfileString "Players", "Count", UBound(Players)
    For i = LBound(Players) To UBound(Players)
        WriteProfileString "Players", "PlayerName" & i, Players(i).PlayerName
        WriteProfileString "Players", "PlayerColor" & i, CStr(Players(i).TankColor)
    Next i
    WriteProfileString "TerrainSettings", "RandomTerrain", CStr(RandomTerrain)
    WriteProfileString "TerrainSettings", "MinShift", CStr(LineShiftMin)
    WriteProfileString "TerrainSettings", "MaxShift", CStr(LineShiftMax)
    WriteProfileString "TerrainSettings", "MinLen", CStr(LineLenMin)
    WriteProfileString "TerrainSettings", "MaxLen", CStr(LineLenMax)
    WriteProfileString "TerrainSettings", "MinVer", CStr(LineVerMin)
    WriteProfileString "TerrainSettings", "MaxVer", CStr(LineVerMax)
    WriteProfileString "ColorSettings", "RandomColors", CStr(RandomColor)
    WriteProfileString "ColorSettings", "SkyMin", CStr(SkyColor)
    WriteProfileString "ColorSettings", "SkyMax", CStr(SkyChange)
    WriteProfileString "ColorSettings", "TerrainMin", CStr(TerrainColor)
    WriteProfileString "ColorSettings", "TerrainMax", CStr(TerrainChange)
    WriteProfileString "ColorSettings", "BulletMax", CStr(BulletColor)
    WriteProfileString "ColorSettings", "BulletMin", CStr(BulletChange)
End Sub

Public Function GetRank(TankNum As Integer) As Integer
Dim i As Integer
    GetRank = 1
    For i = LBound(Players) To UBound(Players)
        If i = TankNum Then
            If i = UBound(Players) Then Exit Function
            i = i + 1
        End If
        GetRank = GetRank + Abs(CInt(Players(i).PlayerScore >= Players(TankNum).PlayerScore))
    Next i
End Function

Public Function GetHiPlayer() As Integer
Dim i As Integer, HiScore As Integer
    HiScore = Players(LBound(Players)).PlayerScore
    For i = LBound(Players) + 1 To UBound(Players)
        If Players(i).PlayerScore > HiScore Then
            HiScore = Players(i).PlayerScore
            GetHiPlayer = i
        End If
    Next i
End Function

Public Sub EnableContainerObjects(ContainerForm As Form, ContainerName As Control, Enabled As Boolean)
Dim i As Integer
    For i = 0 To ContainerForm.Controls.Count - 1
        If Not InStr(1, ILLEGAL_OBJECTS, TypeName(ContainerForm.Controls(i))) > 0 Then
            If ContainerForm.Controls(i).Container.Name = ContainerName.Name Then
                ContainerForm.Controls(i).Enabled = Enabled
            End If
        End If
    Next i
End Sub

Public Sub StopSound()
    PlaySound WAV_NOTHING, SND_ASYNC
End Sub

Public Sub CamoForm(FormName As Object, UseAutoRedraw As Boolean)
Dim CamoPic As IPictureDisp, SavedRedraw As Integer
Dim X As Integer, Y As Integer
    SavedRedraw = FormName.AutoRedraw
    FormName.AutoRedraw = UseAutoRedraw
    FormName.ScaleMode = vbPixels
    Set CamoPic = LoadPicture(AppPath & BK_PICT)
    For Y = 1 To Int(FormName.ScaleHeight / BK_HIGH) + 1
        For X = 1 To Int(FormName.ScaleWidth / BK_WIDE) + 1
            FormName.PaintPicture CamoPic, (X - 1) * BK_WIDE, (Y - 1) * BK_HIGH, BK_WIDE, BK_HIGH
        Next X
    Next Y
    FormName.AutoRedraw = SavedRedraw
End Sub

Private Sub GetSounds()
    GetSoundsFromFolder FL_FIRE, FireSound
    GetSoundsFromFolder FL_TERRAIN, TerrainSound
    GetSoundsFromFolder FL_SCORE, ScoreSound
    GetSoundsFromFolder FL_LAUNCH, LaunchSound
End Sub

Private Sub GetSoundsFromFolder(MaskString As String, ByRef StoreArray() As String)
Dim FileStr As String, SoundFolder As String, i As Integer
    Erase StoreArray
    SoundFolder = AppPath & FOLDER_SOUND
    FileStr = Dir(SoundFolder & MaskString)
    Do While FileStr > ""
        i = i + 1
        ReDim Preserve StoreArray(1 To i)
        StoreArray(i) = Left(FileStr, InStr(1, FileStr, ".") - 1)
        FileStr = Dir()
    Loop
End Sub

Public Sub RandomArraySound(ArrayName() As String, SoundMode As Long)
    On Error GoTo IsError
    PlaySound ArrayName(RandomNumber(LBound(ArrayName), UBound(ArrayName))), SoundMode
    Exit Sub
IsError:
    SoundErr
End Sub

Sub SoundErr()
Static PrevErr  As Boolean
    If PrevErr Then Exit Sub
    PrevErr = True
    MsgBox MSG_SNDERR, vbCritical
End Sub

Public Function RandomNumber(Min, Max)
    Randomize Timer
    RandomNumber = Rnd * (Max - Min) + Min
End Function

Public Sub AddToScore(AddScore As Integer, TankNum As Integer)
    Players(TankNum).PlayerScore = Players(TankNum).PlayerScore + AddScore
    If Players(TankNum).PlayerScore < MIN_SCORE Then Players(TankNum).PlayerScore = MIN_SCORE
End Sub

Public Function GetDirection(Direction As Integer) As Integer
Dim GeneralValue As Integer
    On Error GoTo BadAngle
    GeneralValue = Direction / Abs(Direction)
    GetDirection = GeneralValue * (90 - Abs(Direction))
    Exit Function
BadAngle:
    Err.Clear
    GeneralValue = 1
    Resume Next
End Function

Public Function BulletHitPlayer(PointPos As PointAPI, PlayNum As Integer) As Boolean
    If PointPos.X > Players(PlayNum).TankPos.X + 3 And PointPos.X < Players(PlayNum).TankPos.X + DRAWTANK_WIDTH - 3 Then
        If PointPos.Y > Players(PlayNum).TankPos.Y - 12 And PointPos.Y < Players(PlayNum).TankPos.Y Then
            BulletHitPlayer = True
        End If
    End If
End Function

Public Function GetNoseTip(PlayerNum As Integer) As PointAPI
Dim StartPoint As PointAPI
    StartPoint = GetTankFillPoint(PlayerNum)
    GetNoseTip.X = StartPoint.X + Sine(Players(PlayerNum).TankAim.Direction) * NOSE_LEN
    GetNoseTip.Y = StartPoint.Y - Cosine(Players(PlayerNum).TankAim.Direction) * NOSE_LEN
End Function

Public Function GetTankFillPoint(TankNum As Integer) As PointAPI
    GetTankFillPoint.X = Players(TankNum).TankPos.X + DRAWTANK_WIDTH / 2 - 1
    GetTankFillPoint.Y = Players(TankNum).TankPos.Y - 4
End Function

Public Sub MakeFormTop(FormName As Form, Topmost As Boolean)
    SetWindowPos FormName.hwnd, -(CInt(Topmost) + 2), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub ShowCurs(Optional Show As Boolean = True)
Dim sCursor As Long, sShow As Integer
    sShow = -Show * 2 - 1
    Do
        sCursor = ShowCursor(Show)
    Loop Until Abs(sCursor) * sShow = sCursor And sCursor <> 0
End Sub
