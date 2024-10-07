VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum LineStyle
    Neutral
    Inward
    Outward
End Enum

Enum TagTest
    Two = -1
    None = 0
    Inset = 1
    Outset = 2
End Enum

Const VALID_CHARS = "012"

Dim InTag As String
Dim OutTag As String
Dim LnkFrm
Dim Msk As String
Dim Targets() As Object
Dim NColor As Long
Dim SColor As Long
Dim HColor As Long
Dim AdjustX As Double, AdjustY As Double

Property Get HighlightColor() As OLE_COLOR
    HighlightColor = HColor
End Property

Property Let HighlightColor(NewColor As OLE_COLOR)
    HColor = NewColor
End Property

Property Get ShadowColor() As OLE_COLOR
    ShadowColor = SColor
End Property

Property Let ShadowColor(NewColor As OLE_COLOR)
    SColor = NewColor
End Property

Property Get NeutralColor() As OLE_COLOR
    NeutralColor = NColor
End Property

Property Let NeutralColor(NewColor As OLE_COLOR)
    NColor = NewColor
End Property

Property Get LinkForm()
    Set LinkForm = LnkFrm
End Property

Property Set LinkForm(NewLink)
    Set LnkFrm = NewLink
End Property

Property Get InsetTag() As String
    InsetTag = InTag
End Property

Property Let InsetTag(NewTag As String)
    InTag = NewTag
End Property

Property Get OutsetTag() As String
    OutsetTag = OutTag
End Property

Property Let OutsetTag(NewTag As String)
    OutTag = NewTag
End Property

Property Get Mask() As String
    Mask = Msk
End Property

Property Let Mask(NewMask As String)
    Msk = NewMask
    PaintTargetControls
End Property

Private Sub FindTargetControls()
Dim ActiveObject, i As Integer, NumCtls As Integer
Dim ValidTarget As Boolean
    On Error Resume Next
    If IsEmpty(LnkFrm) Then Exit Sub
    For Each ActiveObject In LnkFrm.Controls
        ValidTarget = False
        ValidTarget = CheckForTag(ActiveObject) > 0
        If ValidTarget Then
            NumCtls = UBound(Targets)
            ReDim Preserve Targets(1 To NumCtls + 1)
            Set Targets(UBound(Targets)) = ActiveObject
        End If
    Next ActiveObject
End Sub

Public Sub PaintTargetControls()
Dim i As Integer, j As Integer
Dim LineMode As LineStyle
    On Error Resume Next
    FindTargetControls
    For i = 1 To Len(Msk)
        LineMode = CInt(Left(Right(Msk, i), 1))
        For j = 1 To UBound(Targets)
            DrawLine LineMode, Targets(j), i
        Next j
    Next i
    LnkFrm.AutoRedraw = True
End Sub

Public Sub PaintSingleControl(ControlName As Control)
Dim i As Integer, LineMode As LineStyle
    For i = 1 To Len(Msk)
        LineMode = CInt(Left(Right(Msk, i), 1))
        DrawLine LineMode, ControlName, i
    Next i
End Sub

Private Function CheckForTag(TestObj) As TagTest
Dim InsetPresent As Boolean, OutsetPresent As Boolean
    InsetPresent = InStr(1, TestObj.Tag, InTag) > 0
    OutsetPresent = InStr(1, TestObj.Tag, OutTag) > 0
    If InsetPresent Then CheckForTag = Inset
    If OutsetPresent Then CheckForTag = Outset
    If Not (InsetPresent Or OutsetPresent) Then CheckForTag = None
    If InsetPresent And OutsetPresent Then CheckForTag = Two
End Function

Public Function CheckMask(TempMask As String) As String
Dim i As Integer, CharPos As Integer
Dim Character As String, StartLen As Integer
    StartLen = Len(TempMask)
    If StartLen = 0 Then Exit Function
    Do
        i = i + 1
        Character = Right(Left(TempMask, i), 1)
        CharPos = InStr(1, VALID_CHARS, Character)
        If Not CharPos > 0 Then
            TempMask = Left(TempMask, i - 1) & Right(TempMask, Len(TempMask) - i)
            i = i - 1
        End If
    Loop Until i = Len(TempMask)
    CheckMask = TempMask
End Function

Private Sub DrawLine(LStyle As LineStyle, ControlName As Control, Level As Integer)
Dim TopLeft As Long, BottomRight As Long, TagTest As Integer
Dim SavedScaleMode As Integer, SavedTopLeft As Long
    On Error Resume Next
    If Not ControlName.Visible Then Exit Sub
    SavedScaleMode = ControlName.Container.ScaleMode
    ControlName.Container.ScaleMode = vbTwips
    TagTest = CheckForTag(ControlName)
    If TagTest < 1 Then Exit Sub
    If LStyle = Neutral Then
        TopLeft = NColor
        BottomRight = NColor
    ElseIf LStyle = Inward Then
        TopLeft = SColor
        BottomRight = HColor
    ElseIf LStyle = Outward Then
        TopLeft = HColor
        BottomRight = SColor
    End If
    If TagTest = Inset Then
        SavedTopLeft = TopLeft
        TopLeft = BottomRight
        BottomRight = SavedTopLeft
    End If
    ControlName.Container.CurrentX = ControlName.Left - (AdjustX * Level)
    ControlName.Container.CurrentY = ControlName.Top - (AdjustY * Level)
    ControlName.Container.Line -(ControlName.Left + ControlName.Width + (AdjustX * (Level - 1)), ControlName.Top - (AdjustY * Level)), TopLeft
    ControlName.Container.Line -(ControlName.Left + ControlName.Width + (AdjustX * (Level - 1)), ControlName.Top + ControlName.Height + (AdjustY * (Level - 1))), BottomRight
    ControlName.Container.Line -(ControlName.Left - (AdjustX * Level), ControlName.Top + ControlName.Height + (AdjustY * (Level - 1))), BottomRight
    ControlName.Container.Line -(ControlName.Left - (AdjustX * Level), ControlName.Top - (AdjustY * Level)), TopLeft
    ControlName.Container.ScaleMode = SavedScaleMode
End Sub

Public Function RandomBorderMask(Length As Integer) As String
Dim i As Integer
    Randomize Timer
    For i = 1 To Length
        RandomBorderMask = RandomBorderMask & Int(Rnd * 3)
    Next i
End Function

Private Sub Class_Initialize()
    FindTargetControls
    InTag = "/in"
    OutTag = "/out"
    AdjustX = Screen.TwipsPerPixelX
    AdjustY = Screen.TwipsPerPixelY
    NColor = vb3DFace
    HColor = vb3DHighlight
    SColor = vb3DShadow
End Sub
