VERSION 5.00
Begin VB.Form frmTank 
   BackColor       =   &H00999999&
   BorderStyle     =   0  'None
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   11970
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Tank.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   579
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   798
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "/p"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctControls 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00007F7F&
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   150
      ScaleHeight     =   885
      ScaleWidth      =   11685
      TabIndex        =   3
      Tag             =   "/in"
      Top             =   7650
      Width           =   11715
      Begin VB.HScrollBar scrAngle 
         Height          =   255
         LargeChange     =   10
         Left            =   150
         Max             =   90
         Min             =   -90
         TabIndex        =   6
         Top             =   510
         Width           =   1575
      End
      Begin VB.HScrollBar scrPower 
         Height          =   255
         LargeChange     =   100
         Left            =   1950
         Max             =   1000
         TabIndex        =   5
         Top             =   510
         Width           =   1575
      End
      Begin VB.CommandButton cmdFire 
         Caption         =   "FIRE!!!"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3690
         TabIndex        =   4
         Top             =   150
         Width           =   1515
      End
      Begin VB.Label lblAngle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Angle:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Top             =   150
         Width           =   450
      End
      Begin VB.Label lblAng 
         Alignment       =   2  'Center
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "???"
         ForeColor       =   &H00007F00&
         Height          =   255
         Left            =   870
         TabIndex        =   17
         Top             =   150
         Width           =   855
      End
      Begin VB.Label lblPower 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Power:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1950
         TabIndex        =   16
         Top             =   150
         Width           =   495
      End
      Begin VB.Label lblPow 
         Alignment       =   2  'Center
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "???"
         ForeColor       =   &H00007F00&
         Height          =   255
         Left            =   2670
         TabIndex        =   15
         Top             =   150
         Width           =   855
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00007F00&
         Height          =   300
         Left            =   6150
         TabIndex        =   14
         Top             =   120
         Width           =   1395
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRankCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   5475
         TabIndex        =   13
         Top             =   525
         Width           =   435
      End
      Begin VB.Label lblRank 
         Alignment       =   2  'Center
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00007F00&
         Height          =   300
         Left            =   6150
         TabIndex        =   12
         Top             =   495
         Width           =   1395
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblScoreCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   5475
         TabIndex        =   11
         Top             =   150
         Width           =   465
      End
      Begin VB.Label lblHiPlayerCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hi Player:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   7800
         TabIndex        =   10
         Top             =   120
         Width           =   675
      End
      Begin VB.Label lblHiPlayer 
         Alignment       =   2  'Center
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?????"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00007F00&
         Height          =   300
         Left            =   8625
         TabIndex        =   9
         Top             =   120
         Width           =   2895
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblHiScoreCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hi Score:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   7800
         TabIndex        =   8
         Top             =   570
         Width           =   660
      End
      Begin VB.Label lblHiScore 
         Alignment       =   2  'Center
         BackColor       =   &H00CCCCCC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "?????"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00007F00&
         Height          =   300
         Left            =   8625
         TabIndex        =   7
         Top             =   495
         Width           =   2895
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox pctField 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   6915
      Left            =   150
      ScaleHeight     =   459
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   779
      TabIndex        =   0
      Tag             =   "/in"
      Top             =   150
      Width           =   11715
      Begin VB.Label lblCondition 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%CONDITION%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   5025
         TabIndex        =   1
         Top             =   2925
         Width           =   1665
      End
   End
   Begin VB.PictureBox Tank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00999999&
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   2
      Left            =   840
      Picture         =   "Tank.frx":030A
      ScaleHeight     =   180
      ScaleWidth      =   240
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "/p"
      Top             =   180
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Tank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00999999&
      BorderStyle     =   0  'None
      Height          =   90
      Index           =   1
      Left            =   540
      Picture         =   "Tank.frx":0384
      ScaleHeight     =   90
      ScaleWidth      =   240
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "/p"
      Top             =   255
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Tank 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00999999&
      BorderStyle     =   0  'None
      Height          =   90
      Index           =   0
      Left            =   240
      Picture         =   "Tank.frx":03C8
      ScaleHeight     =   90
      ScaleWidth      =   240
      TabIndex        =   21
      TabStop         =   0   'False
      Tag             =   "/p"
      Top             =   255
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00999999&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   150
      TabIndex        =   2
      Tag             =   "/in"
      Top             =   7200
      Width           =   11715
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Game"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFilePause1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileEnd 
         Caption         =   "&End Game"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFilePause 
         Caption         =   "&Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFilePause2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "&About..."
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSettings 
         Caption         =   "&Settings..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFilePause3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmTank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const BTN_FIRE = "FIRE!!!"
Const BTN_RESET = "RESET"

Private Sub ShowTurn()
    GetTurn TankNumber
    scrAngle.Value = Players(TankNumber).TankAim.Direction
    scrPower.Value = Players(TankNumber).TankAim.Power
    lblStatus.ForeColor = GetTankColor(TankNumber)
    lblHiPlayer = Players(GetHiPlayer).PlayerName
    lblHiScore = Players(GetHiPlayer).PlayerScore
    lblRank = GetRank(TankNumber)
    lblScore = Players(TankNumber).PlayerScore
    If Not TankHasWon Then PrintMessage MSG_PLAYER_TURN, Players(TankNumber).PlayerName
End Sub

Private Sub cmdFire_Click()
Static InProgress As Boolean, i As Integer
    Select Case cmdFire.Caption
    Case BTN_FIRE
        ChangeCondition 3
        ConditionForm
        RandomArraySound FireSound, SND_SYNC
        RandomArraySound LaunchSound, SND_ASYNC
        FireBullet TankNumber
    Case BTN_RESET
        IsInGame = True
        For i = LBound(Players) To UBound(Players)
            Players(i).TankState = 0
        Next i
        TankHasWon = False
        TankOrder = ReplaceChars(DieOrder, " ") & ReplaceChars(TankOrder, " ")
        DieOrder = ""
        NextTurn True
        DrawGround
        cmdFire.Caption = BTN_FIRE
        ChangeCondition 1
        ConditionForm
    End Select
End Sub

Private Sub DrawBullet(Position As PointAPI, Color As Long)
    pctField.FillColor = Color
    pctField.Circle (Position.X, Position.Y), 1, Color
End Sub

Private Sub FireBullet(PlayerNum As Integer)
Dim BulletPos As PointAPI, PrevBulletPos As PointAPI
Dim UpMov As Currency, RightMov As Currency
Dim i As Integer, IsSuicide As Boolean, TankWon As Integer
Dim RemoveTank As Boolean
    BulletPos = GetNoseTip(PlayerNum)
    BulletPos.Y = BulletPos.Y - 5
    UpMov = (Sine(-Players(PlayerNum).TankAim.Direction + 90) * Players(PlayerNum).TankAim.Power) / 90
    RightMov = (Cosine(-Players(PlayerNum).TankAim.Direction + 90) * Players(PlayerNum).TankAim.Power) / 90
    PrintMessage MSG_FIRE, Players(TankNumber).PlayerName
    pctField.DrawWidth = 1
    Do
        Do While GameCondition = 2
            DoEvents
        Loop
        pctField.AutoRedraw = True
        UpMov = UpMov - 0.1
        PrevBulletPos = MovePoint(BulletPos.X, BulletPos.Y)
        BulletPos = MovePoint(BulletPos.X + RightMov, BulletPos.Y - UpMov)
        DrawBullet PrevBulletPos, Sky
        For i = LBound(Players) To UBound(Players)
            If BulletHitPlayer(BulletPos, i) And Players(i).TankState < 3 Then
                IsSuicide = (i = TankNumber)
                RemoveTank = (Players(i).TankState > 0)
                If IsSuicide Then
                    PrintMessage MSG_SUICIDE, Players(TankNumber).PlayerName
                    AddToScore PNT_SUICIDE, TankNumber
                Else
                    If Players(i).TankState = 0 Then
                        PrintMessage MSG_HIT & Players(i).PlayerName, Players(TankNumber).PlayerName
                        AddToScore PNT_HITTANK, TankNumber
                    Else
                        PrintMessage MSG_HITRUIN & Players(i).PlayerName, Players(TankNumber).PlayerName
                    End If
                End If
                RandomArraySound ScoreSound, SND_ASYNC
                TankDied i
                Players(i).TankState = STATE_DYING
                DrawTanks
                Wait 1
                Players(i).TankState = STATE_DEAD
                If RemoveTank Then Players(i).TankState = STATE_GONE
                If Len(Trim$(TankOrder)) = 1 Then
                    TankWon = Val(TankOrder)
                    AddToScore PNT_ALIVE, TankWon
                    PrintMessage MSG_ALIVE, Players(TankWon).PlayerName
                    TankHasWon = True
                    ResetForm
                    Exit Sub
                End If
                DrawTanks
                GoTo NextPlayer
            End If
        Next i
        If pctField.Point(BulletPos.X, BulletPos.Y) = Terrain Then
            pctField.FillColor = Sky
            PrintMessage MSG_TERRAIN, Players(TankNumber).PlayerName
            pctField.Circle (BulletPos.X, BulletPos.Y), 10, Sky
            RandomArraySound TerrainSound(), SND_ASYNC
            GoTo NextPlayer
        End If
        If BulletPos.X > pctField.Width Then BulletPos.X = 0
        If BulletPos.X < 0 Then BulletPos.X = pctField.Width
        If BulletPos.Y > pctField.Height Then
            PrintMessage MSG_OUTBOUND, Players(TankNumber).PlayerName
            GoTo NextPlayer
        End If
        DrawBullet BulletPos, Bullet
        DoEvents
    Loop
NextPlayer:
    DrawTanks
    Wait 1
    NextTurn
    ShowTurn
    ChangeCondition 1
End Sub

Sub ResetForm()
    cmdFire.Caption = BTN_RESET
    ChangeCondition 1
End Sub

Private Sub Form_Load()
    ConditionForm
    PaintPictures
    CamoForm Me, True
End Sub

Private Sub lblAng_Change()
    Players(TankNumber).TankAim.Direction = scrAngle.Value
    DoEvents
    DrawTanks
End Sub

Private Sub lblCondition_Change()
Static InChange As Boolean
    lblCondition.Visible = Len(lblCondition) > 1
    If InChange Then Exit Sub
    InChange = True
    lblCondition = UCase(lblCondition)
    InChange = False
End Sub

Private Sub lblStatus_Change()
Static InChange As Boolean
    If InChange Then Exit Sub
    InChange = True
    lblStatus = UCase(lblStatus)
    InChange = False
End Sub

Private Sub mnuFileAbout_Click()
    PlaySound WAV_FORMLOAD, SND_ASYNC
    frmAbout.Show vbModal
End Sub

Private Sub mnuFileEnd_Click()
Dim i As Integer, ScoreForm As New frmScores
    PlaySound WAV_FORMLOAD, SND_ASYNC
    ScoreForm.Show vbModal
    IsInGame = False
    DieOrder = ""
    TankOrder = ""
    GameCondition = 0
    ConditionForm
    pctField.BackColor = 0
End Sub

Private Sub mnuFileExit_Click()
Dim ScoreForm As New frmScores
    If GameCondition > 0 Then
        PlaySound WAV_FORMLOAD, SND_ASYNC Or SND_NOSTOP
        ScoreForm.Show vbModal
    End If
    Unload Me
End Sub

Private Sub DrawGround()
Dim Point1 As PointAPI, Point2 As PointAPI
Dim TankPoint As PointAPI, PntTank As Boolean
Dim LineNum As Integer, Finished As Boolean
Dim TanksDone As Integer, i As Integer, rc As Long
Dim StartY As Integer
    LineShift = RandomNumber(LineShiftMax, LineShiftMin)
    LineLen = RandomNumber(LineLenMax, LineLenMin)
    LineVer = RandomNumber(LineVerMax, LineVerMin)
    pctField.Picture = LoadPicture()
    pctField.DrawWidth = 2
    If Not RandomTerrain Then
        LineShift = LineShiftMin
        LineLen = LineLenMin
        LineVer = LineVerMin
    End If
    Sky = SkyColor
    Terrain = TerrainColor
    Bullet = BulletColor
    If RandomColor Then
        i = Int(Rnd * 101)
        Sky = ColorBetween(SkyColor, SkyChange, i / 100)
        Terrain = ColorBetween(TerrainColor, TerrainChange, i / 100)
        Bullet = ColorBetween(BulletColor, BulletChange, i / 100)
    End If
    pctField.Cls
    pctField.BackColor = Sky
    pctField.ForeColor = Terrain
    pctField.FillColor = Terrain
    PaintPictures
    For i = LBound(Players) To UBound(Players)
        Players(i).TankPos = MovePoint(Int(RandomNumber(CDbl(rc) + DRAWTANK_WIDTH, (pctField.ScaleWidth - DRAWTANK_WIDTH) / (UBound(Players) + 1) * (i + 1))), RandomHeight)
        rc = Players(i).TankPos.X
    Next i
    Point1 = MovePoint(-1, RandomHeight)
    StartY = Point1.Y
    If Not IsInGame Then TankOrder = ""
    Do
        PntTank = False
        Point2 = Point1
        Point1 = MovePoint(Point2.X + Int(Rnd * LineLen / 2) + LineLen / 2, RandomHeight)
        If TanksDone < UBound(Players) + 1 Then
            If Point1.X > Players(TanksDone).TankPos.X Then
                PntTank = True
                TankPoint = Players(TanksDone).TankPos
            End If
            Players(TanksDone).TankAim.Direction = -(Players(TanksDone).TankPos.X / pctField.ScaleWidth) * 180 + 90
            Players(TanksDone).TankAim.Power = START_POWER
        End If
        If PntTank Then
            pctField.Line (Point2.X, Point2.Y)-(TankPoint.X, TankPoint.Y)
            Point1 = MovePoint(TankPoint.X + DRAWTANK_WIDTH, TankPoint.Y)
            pctField.Line (TankPoint.X, TankPoint.Y)-(Point1.X, Point1.Y)
            Point2 = Point1
            TanksDone = TanksDone + 1
            If Not IsInGame Then TankOrder = TankOrder & Trim$(Str$(TanksDone - 1))
        End If
        If Point1.X > pctField.ScaleWidth Then
            Point1.X = pctField.ScaleWidth
            Point1.Y = StartY
            Finished = True
        End If
        If Not PntTank Then pctField.Line (Point1.X, Point1.Y)-(Point2.X, Point2.Y)
    Loop Until Finished
    rc = FloodFill(pctField.hdc, 5, pctField.Height - 5, Terrain)
    DrawTanks
    TankNumber = Val(Left$(TankOrder, 1))
    ShowTurn
End Sub

Private Function MovePoint(X As Double, Y As Double) As PointAPI
    MovePoint.X = X
    MovePoint.Y = Y
End Function

Private Sub DrawTanks()
Dim rc As Long, i As Integer
Dim FillPoint As PointAPI, CurrentState As Integer
Dim StartPoint As PointAPI, EndPoint As PointAPI
    pctField.DrawWidth = 2
    For i = LBound(Players) To UBound(Players)
        If GameCondition = 0 Then Exit Sub
        CurrentState = Players(i).TankState
        StartPoint = GetTankFillPoint(i)
        pctField.Line (Players(i).TankPos.X + 1, Players(i).TankPos.Y - 2)-(Players(i).TankPos.X + DRAWTANK_WIDTH - 1, Players(i).TankPos.Y - 20), Sky, BF
        EndPoint = GetNoseTip(i)
        If Players(i).TankState < 3 Then rc = BitBlt(pctField.hdc, Players(i).TankPos.X + TANK_OFSET, Players(i).TankPos.Y - Tank(CurrentState).Height - 2, Tank(CurrentState).Width, Tank(CurrentState).Height, Tank(CurrentState).hdc, 0, 0, vbSrcCopy)
        FillPoint = GetTankFillPoint(i)
        pctField.FillColor = Players(i).TankColor
        FloodFill pctField.hdc, FillPoint.X, FillPoint.Y, Sky
        If CurrentState = 0 Then pctField.Line (StartPoint.X, StartPoint.Y)-(EndPoint.X, EndPoint.Y), pctField.FillColor
    Next i
End Sub

Private Function GetColor(PlayerNumber As Integer) As Long
Dim TankPoint As PointAPI
    TankPoint = MovePoint(Players(PlayerNumber).TankPos.X + TANK_OFSET + TANK_WIDTH / 2, Players(PlayerNumber).TankPos.Y - 4)
    GetColor = pctField.Point(TankPoint.X, TankPoint.Y)
End Function

Private Sub PaintPictures()
Dim CtlName
    For Each CtlName In Controls
        If CtlName.Tag = PAINT_TAG Then
            CtlName.BackColor = Sky
            CtlName.Refresh
        End If
    Next CtlName
End Sub

Private Sub mnuFileNew_Click()
    If GameCondition > 0 Then Exit Sub
    lblCondition = ""
    GameCondition = 1
    DrawGround
    ConditionForm
End Sub

Private Sub ChangeCondition(NewCondition As Integer)
    GameCondition = NewCondition
    ConditionForm
End Sub

Private Sub mnuFilePause_Click()
Static InPause As Boolean
    InPause = Not InPause
    If InPause Then ChangeCondition 2 Else: ChangeCondition 1
End Sub

Private Sub mnuFileSettings_Click()
    PlaySound WAV_FORMLOAD, SND_ASYNC
    frmSettings.Show vbModal
End Sub

Private Sub ConditionForm()
Dim InPause As Boolean, InGame As Boolean, InFire As Boolean
Dim Message As String, i As Integer
Static InCondition As Boolean
    If InCondition Then Exit Sub
    InCondition = True
    InGame = (GameCondition > 0)
    InPause = (GameCondition = 2)
    InFire = (GameCondition = 3)
    mnuFilePause.Checked = InPause
    mnuFilePause.Enabled = InGame
    EnableContainerObjects Me, pctControls, InGame And Not (InFire Or InPause)
    If Not (IsInGame Or InGame) Then
        EraseData True
    End If
    If InPause Then
        Message = MSG_PAUSED
        InCondition = False
        Exit Sub
    End If
    If Not InGame Then
        IsInGame = False
        pctField.Cls
        Message = MSG_INACTIVE
        PrintMessage MSG_NOGAME
        lblStatus.ForeColor = 0
    Else
        ShowTurn
    End If
    mnuFileEnd.Enabled = InGame
    mnuFileNew.Enabled = Not InGame
    mnuFileSettings.Enabled = Not InGame
    mnuFileAbout.Enabled = Not InGame
    lblCondition = Message
    InCondition = False
End Sub

Private Sub EraseData(EraseScores As Boolean)
Dim i As Integer
    For i = LBound(Players) To UBound(Players)
        Players(i).TankState = 0
        If EraseScores Then Players(i).PlayerScore = 0
    Next i
    DieOrder = ""
    TankOrder = ""
End Sub

Private Function InResetMode() As Boolean
    InResetMode = (cmdFire.Caption = BTN_RESET)
End Function

Private Sub PrintMessage(Message As String, Optional PlayerName As String)
    lblStatus = PlayerName & Message
End Sub

Private Sub pctControls_Click()
    If GameCondition <> 1 Then PlaySound WAV_NEGATIVE, SND_SYNC
End Sub

Private Sub pctField_KeyDown(KeyCode As Integer, Shift As Integer)
Debug.Print Shift
    If GameCondition <> 1 Then Exit Sub
    If KeyCode = vbKeyUp Then
        AddToScroll scrPower, 9 * Shift + 1
    ElseIf KeyCode = vbKeyDown Then
        AddToScroll scrPower, -(9 * Shift + 1)
    ElseIf KeyCode = vbKeyLeft Then
        AddToScroll scrAngle, -(9 * Shift + 1)
    ElseIf KeyCode = vbKeyRight Then
        AddToScroll scrAngle, 9 * Shift + 1
    End If
End Sub

Private Sub AddToScroll(ScrollName As HScrollBar, Amount As Integer)
Dim AddAmt As Integer
    AddAmt = ScrollName.Value + Amount
    If AddAmt < ScrollName.Min Then AddAmt = ScrollName.Min
    If AddAmt > ScrollName.Max Then AddAmt = ScrollName.Max
    ScrollName.Value = AddAmt
End Sub

Private Sub scrAngle_Change()
Dim NewDirection As Integer
    PlaySound WAV_TICK, SND_ASYNC Or SND_NOSTOP
    scrAngle.Refresh
    NewDirection = Abs(GetDirection(scrAngle.Value))
    If lblAng = Trim$(Str$(NewDirection)) Then lblAng = ""
    lblAng = NewDirection
End Sub

Private Function GetTankColor(TankNum As Integer) As Long
    GetTankColor = pctField.Point(Players(TankNum).TankPos.X + 10, Players(TankNum).TankPos.Y - 5)
End Function

Private Sub scrPower_Change()
    PlaySound WAV_TICK, SND_ASYNC Or SND_NOSTOP
    scrPower.Refresh
    lblPow = Abs(scrPower.Value)
    Players(TankNumber).TankAim.Power = scrPower.Value
End Sub

Private Function RandomHeight() As Double
    RandomHeight = RandomNumber(LineVer + LineShift, LineVer - LineShift)
    If RandomHeight > pctField.ScaleHeight - 10 Then
        RandomHeight = pctField.ScaleHeight - 10
    ElseIf RandomHeight < 20 Then
        RandomHeight = 20
    End If
End Function
