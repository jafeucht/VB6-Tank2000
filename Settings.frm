VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSettings 
   BackColor       =   &H00999999&
   BorderStyle     =   0  'None
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   511
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSaveSettings 
      Caption         =   "&Save As Default"
      Height          =   390
      Left            =   4050
      TabIndex        =   40
      Top             =   3240
      Width           =   1665
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   7125
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   390
      Left            =   5820
      TabIndex        =   0
      Top             =   3240
      Width           =   1665
   End
   Begin VB.PictureBox Settings 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00999999&
      ForeColor       =   &H80000008&
      Height          =   2565
      Index           =   0
      Left            =   150
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   1
      Tag             =   "/in"
      Top             =   600
      Width           =   7365
      Begin VB.PictureBox pctTankColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   5025
         ScaleHeight     =   26
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   134
         TabIndex        =   31
         Top             =   1350
         Width           =   2040
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   3450
         TabIndex        =   5
         Top             =   1950
         Width           =   1215
      End
      Begin VB.TextBox txtPlayerName 
         Height          =   315
         Left            =   3450
         MaxLength       =   20
         TabIndex        =   4
         Top             =   855
         Width           =   3615
      End
      Begin VB.ComboBox cboPlayers 
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1905
         Width           =   2640
      End
      Begin VB.ListBox lstPlayers 
         Height          =   1230
         Left            =   300
         TabIndex        =   2
         Top             =   600
         Width           =   2640
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Color:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   4020
         TabIndex        =   30
         Top             =   1425
         Width           =   885
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   3450
         TabIndex        =   8
         Top             =   555
         Width           =   945
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Information:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   3300
         TabIndex        =   7
         Top             =   180
         Width           =   1305
      End
      Begin VB.Shape fmInfo 
         Height          =   1890
         Left            =   3300
         Tag             =   "/out"
         Top             =   480
         Width           =   3915
      End
      Begin VB.Label lblPlayers 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Players:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   150
         Width           =   555
      End
      Begin VB.Shape fmPlayers 
         Height          =   1890
         Left            =   150
         Tag             =   "/out"
         Top             =   480
         Width           =   2940
      End
   End
   Begin VB.PictureBox Settings 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00999999&
      ForeColor       =   &H80000008&
      Height          =   2565
      Index           =   1
      Left            =   150
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   11
      Tag             =   "/in"
      Top             =   600
      Visible         =   0   'False
      Width           =   7365
      Begin VB.PictureBox pctColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1080
         ScaleHeight     =   285
         ScaleWidth      =   1035
         TabIndex        =   42
         Top             =   1530
         Width           =   1065
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change..."
         Height          =   315
         Index           =   2
         Left            =   2220
         TabIndex        =   41
         Top             =   1530
         Width           =   1065
      End
      Begin VB.PictureBox pctRandom 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00999999&
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   3420
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   228
         TabIndex        =   32
         Tag             =   "/out/c"
         Top             =   570
         Width           =   3450
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   5
            Left            =   960
            ScaleHeight     =   285
            ScaleWidth      =   1035
            TabIndex        =   51
            Top             =   1140
            Width           =   1065
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "Change..."
            Height          =   315
            Index           =   5
            Left            =   2100
            TabIndex        =   50
            Top             =   1140
            Width           =   1065
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   3
            Left            =   960
            ScaleHeight     =   285
            ScaleWidth      =   1035
            TabIndex        =   47
            Top             =   180
            Width           =   1065
         End
         Begin VB.PictureBox pctColor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   4
            Left            =   960
            ScaleHeight     =   285
            ScaleWidth      =   1035
            TabIndex        =   46
            Top             =   660
            Width           =   1065
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "Change..."
            Height          =   315
            Index           =   3
            Left            =   2100
            TabIndex        =   45
            Top             =   180
            Width           =   1065
         End
         Begin VB.CommandButton cmdChange 
            Caption         =   "Change..."
            Height          =   315
            Index           =   4
            Left            =   2100
            TabIndex        =   44
            Top             =   660
            Width           =   1065
         End
         Begin VB.Label lblBullet 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bullet:"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   1200
            Width           =   435
         End
         Begin VB.Label lblSky 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sky:"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   240
            Width           =   315
         End
         Begin VB.Label lblGround 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ground:"
            ForeColor       =   &H0000FF00&
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   720
            Width           =   570
         End
      End
      Begin VB.CheckBox chkRandom 
         BackColor       =   &H00999999&
         Caption         =   "Random Colors"
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   300
         TabIndex        =   19
         Top             =   1980
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change..."
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   16
         Top             =   1050
         Width           =   1065
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change..."
         Height          =   315
         Index           =   0
         Left            =   2220
         TabIndex        =   15
         Top             =   570
         Width           =   1065
      End
      Begin VB.PictureBox pctColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   1080
         ScaleHeight     =   285
         ScaleWidth      =   1035
         TabIndex        =   14
         Top             =   1050
         Width           =   1065
      End
      Begin VB.PictureBox pctColor 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1080
         ScaleHeight     =   285
         ScaleWidth      =   1035
         TabIndex        =   13
         Top             =   570
         Width           =   1065
      End
      Begin VB.Label lblBulletCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bullet:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   300
         TabIndex        =   43
         Top             =   1590
         Width           =   435
      End
      Begin VB.Label lblGroundCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ground:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   300
         TabIndex        =   18
         Top             =   1110
         Width           =   570
      End
      Begin VB.Label lblSkyCol 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sky:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   630
         Width           =   315
      End
      Begin VB.Label lblColors 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colors:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   480
      End
      Begin VB.Shape shpFrame 
         Height          =   1965
         Index           =   0
         Left            =   120
         Tag             =   "/out"
         Top             =   420
         Width           =   6915
      End
   End
   Begin VB.PictureBox Settings 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00999999&
      ForeColor       =   &H80000008&
      Height          =   2565
      Index           =   2
      Left            =   150
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   489
      TabIndex        =   21
      Tag             =   "/in"
      Top             =   600
      Visible         =   0   'False
      Width           =   7365
      Begin VB.CheckBox chkRandomTerrain 
         BackColor       =   &H00999999&
         Caption         =   "Random Terrain"
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   300
         TabIndex        =   29
         Top             =   1980
         Width           =   1590
      End
      Begin VB.TextBox txtTexture 
         Height          =   315
         Left            =   1620
         TabIndex        =   28
         Top             =   1560
         Width           =   1590
      End
      Begin VB.TextBox txtShift 
         Height          =   315
         Left            =   1620
         TabIndex        =   26
         Top             =   1080
         Width           =   1590
      End
      Begin VB.TextBox txtHorizon 
         Height          =   315
         Left            =   1620
         TabIndex        =   24
         Top             =   600
         Width           =   1590
      End
      Begin VB.PictureBox pctRandomTerrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00999999&
         ForeColor       =   &H80000008&
         Height          =   1665
         Left            =   3480
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   33
         Tag             =   "/out/c"
         Top             =   600
         Width           =   3390
         Begin VB.TextBox txtMaxTexture 
            Height          =   315
            Left            =   1575
            TabIndex        =   36
            Top             =   1125
            Width           =   1590
         End
         Begin VB.TextBox txtMaxShift 
            Height          =   315
            Left            =   1575
            TabIndex        =   35
            Top             =   675
            Width           =   1590
         End
         Begin VB.TextBox txtMaxHorizon 
            Height          =   315
            Left            =   1575
            TabIndex        =   34
            Top             =   225
            Width           =   1590
         End
         Begin VB.Label lblMaxTexture 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Terrain Texture:"
            ForeColor       =   &H00FFFFC0&
            Height          =   195
            Left            =   225
            TabIndex        =   39
            Top             =   1125
            Width           =   1125
         End
         Begin VB.Label lblMaxShift 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Terrain Shift:"
            ForeColor       =   &H00FFFFC0&
            Height          =   195
            Left            =   225
            TabIndex        =   38
            Top             =   675
            Width           =   900
         End
         Begin VB.Label lblMaxHeight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horizon Height:"
            ForeColor       =   &H00FFFFC0&
            Height          =   195
            Left            =   225
            TabIndex        =   37
            Top             =   225
            Width           =   1095
         End
      End
      Begin VB.Label lblTexture 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terrain Texture:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   300
         TabIndex        =   27
         Top             =   1575
         Width           =   1125
      End
      Begin VB.Label lblShift 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terrain Shift:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   300
         TabIndex        =   25
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblHorizon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Horizon Height:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   300
         TabIndex        =   23
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblOptions 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terrain options:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1095
      End
      Begin VB.Shape shpFrame 
         Height          =   1965
         Index           =   1
         Left            =   120
         Tag             =   "/out"
         Top             =   420
         Width           =   6915
      End
   End
   Begin VB.Label Tabs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00007F7F&
      BackStyle       =   0  'Transparent
      Caption         =   "&Terrain"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2190
      TabIndex        =   20
      Tag             =   "/out"
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Tabs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00007F7F&
      BackStyle       =   0  'Transparent
      Caption         =   "&Colors"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1170
      TabIndex        =   10
      Tag             =   "/out"
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Tabs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00007F7F&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Players"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Tag             =   "/in"
      Top             =   240
      Width           =   900
   End
   Begin VB.Shape shpFrame 
      Height          =   3690
      Index           =   2
      Left            =   60
      Tag             =   "/in"
      Top             =   45
      Width           =   7575
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboPlayers_Click()
Static InChange As Boolean
Dim PrevUBound As Integer, i As Integer
    If InChange Then Exit Sub
    ClickCtl Visible
    InChange = True
    PrevUBound = UBound(Players)
    ReDim Preserve Players(0 To cboPlayers.ListIndex + 1)
    For i = PrevUBound + 1 To UBound(Players)
        Players(i).PlayerName = GetProfileString("Players", "PlayerName" & i, TX_PLAYER & i + 1)
        Players(i).TankColor = CLng(GetProfileString("Players", "PlayerColor" & i, CStr(QBColor(i + 1))))
    Next i
    PrintData
    InChange = False
End Sub

Private Sub chkComputer_Click()
    ClickCtl
End Sub

Private Sub chkRandom_Click()
    ClickCtl
    EnableContainerObjects Me, pctRandom, CBool(Abs(chkRandom.Value))
    RandomColor = CBool(chkRandom.Value)
End Sub

Private Sub chkRandomTerrain_Click()
    ClickCtl
    EnableContainerObjects Me, pctRandomTerrain, CBool(Abs(chkRandomTerrain.Value))
    RandomTerrain = CBool(chkRandomTerrain.Value)
End Sub

Private Sub cmdChange_Click(Index As Integer)
    ClickCtl
    cDialog.Color = pctColor(Index).BackColor
    cDialog.ShowColor
    pctColor(Index).BackColor = cDialog.Color
    Select Case Index
        Case 0: SkyColor = cDialog.Color
        Case 1: TerrainColor = cDialog.Color
        Case 2: BulletColor = cDialog.Color
        Case 3: SkyChange = cDialog.Color
        Case 4: TerrainChange = cDialog.Color
        Case 5: BulletChange = cDialog.Color
    End Select
End Sub

Private Sub cmdClose_Click()
    ClickCtl True, SND_SYNC
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
Dim ListPos As Integer
    ClickCtl
    ListPos = lstPlayers.ListIndex
    Players(ListPos).PlayerName = txtPlayerName.Text
    Players(ListPos).TankColor = pctTankColor.BackColor
    PrintData
    If ListPos = lstPlayers.ListCount - 1 Then
        cmdClose.Default = True
    Else
        lstPlayers.ListIndex = ListPos + 1
    End If
End Sub

Private Sub cmdSaveSettings_Click()
    ClickCtl
    SaveSettings
End Sub

Private Sub Form_Load()
    PrintData
    CamoForm Me, True
    MakeCamo
End Sub

Private Sub MakeCamo()
Dim i As Integer
    For i = 0 To Controls.Count - 1
        If InStr(1, Controls(i).Tag, CAMO_TAG) > 0 Then
            If TypeName(Controls(i)) = "PictureBox" Then CamoForm Controls(i), True
        End If
    Next i
End Sub

Private Sub PrintData()
Dim i As Integer
    lstPlayers.Clear
    cboPlayers.Clear
    For i = 2 To MAX_PLAYERS
        cboPlayers.AddItem i & TX_COUNT
    Next i
    CheckPlayerNames
    For i = LBound(Players) To UBound(Players)
        lstPlayers.AddItem Players(i).PlayerName
    Next i
    lstPlayers.ListIndex = 0
    pctColor(0).BackColor = SkyColor
    pctColor(1).BackColor = TerrainColor
    pctColor(2).BackColor = BulletColor
    pctColor(3).BackColor = SkyChange
    pctColor(4).BackColor = TerrainChange
    pctColor(5).BackColor = BulletChange
    chkRandom.Value = -RandomColor
    txtTexture = LineLenMin
    txtShift = LineShiftMin
    txtHorizon = LineVerMin
    txtMaxTexture = LineLenMax
    txtMaxShift = LineShiftMax
    txtMaxHorizon = LineVerMax
    chkRandomTerrain = Abs(CInt(RandomTerrain))
    cboPlayers.ListIndex = UBound(Players) - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopSound
End Sub

Private Sub lstPlayers_Click()
    On Error Resume Next
    ClickCtl Visible
    txtPlayerName = Players(lstPlayers.ListIndex).PlayerName
    txtPlayerName.SelLength = Len(txtPlayerName)
    txtPlayerName.SetFocus
    pctTankColor.BackColor = Players(lstPlayers.ListIndex).TankColor
End Sub

Private Sub pctTankColor_DblClick()
    cDialog.Color = pctTankColor.BackColor
    ClickCtl
    cDialog.ShowColor
    pctTankColor.BackColor = cDialog.Color
End Sub

Private Sub Tabs_Click(Index As Integer)
Dim i As Integer
    ClickCtl
    Tabs(Index).BackStyle = 1
    Tabs(Index).BorderStyle = 1
    For i = 0 To Tabs.Count - 1
        If i <> Index Then
            Settings(i).Visible = False
            Tabs(i).BackStyle = 0
            Tabs(i).BorderStyle = 0
        End If
    Next i
    Wait 0.25
    Settings(Index).Visible = True
End Sub

Private Sub txtHorizon_Change()
    If IsNumeric(txtHorizon) Then LineVerMin = Val(txtHorizon)
    CheckIsNumeric txtHorizon
End Sub

Private Sub txtMaxHorizon_Change()
    If IsNumeric(txtMaxHorizon) Then LineVerMax = Val(txtMaxHorizon)
    CheckIsNumeric txtMaxHorizon
End Sub

Private Sub txtMaxShift_Change()
    If IsNumeric(txtMaxShift) Then LineShiftMax = Val(txtMaxShift)
    CheckIsNumeric txtMaxShift
End Sub

Private Sub txtMaxTexture_Change()
    If IsNumeric(txtMaxTexture) Then LineLenMax = Val(txtMaxTexture)
    CheckIsNumeric txtMaxTexture
End Sub

Private Sub txtPlayerName_Change()
    cmdRefresh.Default = True
End Sub

Private Sub CheckIsNumeric(ByRef CheckCtl As TextBox)
Dim NewBackColor As Long
    CheckCtl.ForeColor = Abs(CInt(IsNumeric(CheckCtl)) + 1) * 255
End Sub

Private Sub txtShift_Change()
    If IsNumeric(txtShift) Then LineShiftMin = Val(txtShift)
    CheckIsNumeric txtShift
End Sub

Private Sub txtTexture_Change()
    If IsNumeric(txtTexture) Then LineLenMin = Val(txtTexture)
    CheckIsNumeric txtTexture
End Sub
