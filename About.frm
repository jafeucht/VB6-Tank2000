VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "About.frx":0000
   ScaleHeight     =   221
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   372
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   435
      Left            =   1980
      TabIndex        =   0
      Top             =   2580
      Width           =   1620
   End
   Begin VB.Shape shpField 
      Height          =   3225
      Left            =   45
      Tag             =   "/in"
      Top             =   45
      Width           =   5490
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub
