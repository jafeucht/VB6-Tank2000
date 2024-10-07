VERSION 5.00
Begin VB.Form frmScores 
   BorderStyle     =   0  'None
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   3720
      Width           =   6135
   End
   Begin VB.PictureBox pctFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3435
      Left            =   240
      ScaleHeight     =   227
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   403
      TabIndex        =   0
      Tag             =   "/out"
      Top             =   180
      Width           =   6075
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   0
         Left            =   -15
         TabIndex        =   1
         Top             =   -15
         Visible         =   0   'False
         Width           =   6075
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Shape shpField 
      Height          =   4140
      Left            =   45
      Tag             =   "/in"
      Top             =   45
      Width           =   6375
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Ranks() As Integer

Const RANKED = " ranked #"
Const RKWITH = " with "
Const PERIOD = " points."

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer, MaxScore As Integer, RankVal As Integer
Dim IsUsedScore As Boolean, j As Integer, k As Integer
Dim PosBonus As Integer
    CamoForm Me, True
    CamoForm pctFrame, True
    RankVal = -32766
    ReDim Ranks(1 To UBound(Players) + 1)
    For j = LBound(Ranks) To UBound(Ranks)
        RankVal = -30000
        For i = LBound(Players) To UBound(Players)
            If Players(i).PlayerScore > RankVal Then
                IsUsedScore = False
                For k = LBound(Ranks) To UBound(Ranks)
                    If Ranks(k) = i + 1 Then IsUsedScore = True
                Next k
                If Not IsUsedScore Then
                    RankVal = Players(i).PlayerScore
                    MaxScore = i + 1
                End If
            End If
        Next i
        Ranks(j) = MaxScore
    Next j
    PosBonus = 1
    For i = LBound(Ranks) To UBound(Ranks)
        Load lblScore(i)
        ScaleMode = 3
        lblScore(i).Caption = Players(Ranks(i) - 1).PlayerName & RANKED & Str$(i) & RKWITH & Players(Ranks(i) - 1).PlayerScore & PERIOD
        lblScore(i).ForeColor = Players(Ranks(i) - 1).TankColor
        lblScore(i).Top = lblScore(i).Height * (i - 1) - (i - 1) - 1
        lblScore(i).Visible = True
        PosBonus = -5
    Next i
End Sub
