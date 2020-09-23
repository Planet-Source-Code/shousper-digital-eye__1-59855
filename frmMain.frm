VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Eye"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBuf 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.PictureBox picStat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   150
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   150
      Width           =   2625
   End
   Begin VB.Timer timFPS 
      Interval        =   1000
      Left            =   2925
      Top             =   4350
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Enum ACTIONS
    d_cw = 0
    d_ccw = 1
    
    rad_grow = 2
    rad_shrink = 4
    
    inn_grow = 8
    inn_shrink = 16
    
    out_grow = 32
    out_shrink = 64
End Enum

Private Type Sector
    sWidthAngle As Single
    sDirection As Single
    sInnerRadius As Single
    sOuterRadius As Single
    bAnimating As Boolean
    sAction As ACTIONS
    iAniTick As Integer
    iAniMax As Integer
End Type

Private Const PI As Single = 3.14159265358979
Private Const RAD As Single = 180 / PI

Const CEN_X As Integer = 400
Const CEN_Y As Integer = 240

Dim MIN_THICKNESS As Long
Dim MAX_THICKNESS As Long

Dim MIN_RADIUS As Long
Dim MAX_RADIUS As Long

Dim MAX_SECTORS As Long
Dim MAX_MOVING As Long
Dim cur_Moving As Long

Dim MOVE_AMOUNT As Long

Dim Sectors() As Sector
Dim FPSLimit As New clsFrameLimiter
Dim fps As Long

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMain.Hide
Unload Me
End
End Sub

Private Sub picStat_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo wait:

Select Case KeyCode
    'Sectors
    Case Is = vbKeyQ
        MAX_SECTORS = MAX_SECTORS + 1
        AddSector
    Case Is = vbKeyA
        MAX_SECTORS = MAX_SECTORS - 1
        If MAX_SECTORS < 2 Then MAX_SECTORS = 1
        If MAX_SECTORS <= MAX_MOVING Then MAX_SECTORS = MAX_MOVING + 1
        RemoveSector
        
    'Moving
    Case Is = vbKeyW
        MAX_MOVING = MAX_MOVING + 1
        If MAX_MOVING >= MAX_SECTORS Then MAX_MOVING = MAX_SECTORS - 1
        AddMoving 1
    Case Is = vbKeyS
        MAX_MOVING = MAX_MOVING - 1
        If MAX_MOVING < 2 Then MAX_MOVING = 1
        RemoveMoving 1
    
    'MoveAmount
    Case Is = vbKeyE
        MOVE_AMOUNT = MOVE_AMOUNT + 1
    Case Is = vbKeyD
        MOVE_AMOUNT = MOVE_AMOUNT - 1
    
    'MaxRad
    Case Is = vbKeyR
        MAX_RADIUS = MAX_RADIUS + 1
    Case Is = vbKeyF
        MAX_RADIUS = MAX_RADIUS - 1
        If MAX_RADIUS < 2 Then MAX_RADIUS = 1
    
    'MinRad
    Case Is = vbKeyT
        MIN_RADIUS = MIN_RADIUS + 1
    Case Is = vbKeyG
        MIN_RADIUS = MIN_RADIUS - 1
        If MIN_RADIUS < 2 Then MIN_RADIUS = 1
    
    'MaxThick
    Case Is = vbKeyY
        MAX_THICKNESS = MAX_THICKNESS + 1
    Case Is = vbKeyH
        MAX_THICKNESS = MAX_THICKNESS - 1
        If MAX_THICKNESS < 2 Then MAX_THICKNESS = 1
    
    'MinThick
    Case Is = vbKeyU
        MIN_THICKNESS = MIN_THICKNESS + 1
    Case Is = vbKeyJ
        MIN_THICKNESS = MIN_THICKNESS - 1
        If MIN_THICKNESS < 2 Then MIN_THICKNESS = 1
End Select

frmMain.FillStyle = vbSolid
frmMain.Circle (CEN_X, CEN_Y), 467, vbWhite
frmMain.FillStyle = vbTransparent

DrawText

Exit Sub
wait:
DoEvents
Resume Next
End Sub

Private Sub Form_Load()
frmMain.Show

MIN_THICKNESS = 50
MAX_THICKNESS = 120

MIN_RADIUS = 50
MAX_RADIUS = 220

MAX_SECTORS = 25
MAX_MOVING = 15

MOVE_AMOUNT = 2

ReDim Preserve Sectors(MAX_SECTORS)

For i = 1 To MAX_SECTORS
    With Sectors(i)
        .bAnimating = False
        .iAniMax = RealRnd(5, MIN_RADIUS)
        .iAniTick = 0
        .sAction = rad_grow
        .sDirection = RealRnd(0, 360)
        .sInnerRadius = RealRnd(MIN_RADIUS, MAX_RADIUS - MIN_THICKNESS)
        .sOuterRadius = RealRnd(MIN_RADIUS + MIN_THICKNESS, MAX_RADIUS)
        .sWidthAngle = RealRnd(15, 60)
    End With
Next i

For i = 1 To MAX_MOVING
    Sectors(i).bAnimating = True
    cur_Moving = cur_Moving + 1
Next i

EnterLoop
End Sub

Private Sub timFPS_Timer()
frmMain.Caption = "Digital Eye - FPS: " & fps
fps = 0
End Sub

Private Function RealRnd(ByVal sngLow As Single, ByVal sngHigh As Single) As Single
DoEvents
Randomize (CDbl(Now()) + Timer)
RealRnd = (Rnd * (sngHigh - sngLow)) + sngLow
End Function

Private Sub DrawSector(s As Sector, col As Long)
Dim StartAng As Single, EndAng As Single
Dim ox1 As Single, oy1 As Single
Dim ox2 As Single, oy2 As Single
Dim ix1 As Single, iy1 As Single
Dim ix2 As Single, iy2 As Single

StartAng = s.sDirection - (s.sWidthAngle / 2)
EndAng = s.sDirection + (s.sWidthAngle / 2)

FormatAngle StartAng
FormatAngle EndAng

StartAng = StartAng / RAD: EndAng = EndAng / RAD

ox1 = CEN_X + Cos(StartAng) * s.sOuterRadius
oy1 = CEN_Y - Sin(StartAng) * s.sOuterRadius
ox2 = CEN_X + Cos(EndAng) * s.sOuterRadius
oy2 = CEN_Y - Sin(EndAng) * s.sOuterRadius

ix1 = CEN_X + Cos(StartAng) * s.sInnerRadius
iy1 = CEN_Y - Sin(StartAng) * s.sInnerRadius
ix2 = CEN_X + Cos(EndAng) * s.sInnerRadius
iy2 = CEN_Y - Sin(EndAng) * s.sInnerRadius

'Angle Lines
frmMain.Line (ix1, iy1)-(ox1, oy1), col
frmMain.Line (ix2, iy2)-(ox2, oy2), col

'Arcs
frmMain.Circle (CEN_X, CEN_Y), s.sOuterRadius, col, StartAng, EndAng, 1
frmMain.Circle (CEN_X, CEN_Y), s.sInnerRadius, col, StartAng, EndAng, 1

DoEvents
End Sub

Private Function ShiftEnergy(ByRef s As Sector)
Dim i As Integer

change:
DoEvents
i = RealRnd(1, MAX_SECTORS)
If Sectors(i).bAnimating = True And MAX_SECTORS - 1 <> MAX_MOVING Then GoTo change:

With Sectors(i)
    .iAniMax = RealRnd(5, 50)
ReGen:
    Select Case CLng(RealRnd(1, 8))
        Case Is = 1
            .sAction = d_cw
        Case Is = 2
            .sAction = d_ccw
        Case Is = 3
            If (.sOuterRadius + .iAniMax * MOVE_AMOUNT) >= MAX_RADIUS Or (.sInnerRadius + .iAniMax * MOVE_AMOUNT) >= MAX_RADIUS Then GoTo ReGen:
            .sAction = rad_grow
        Case Is = 4
            If (.sOuterRadius - .iAniMax * MOVE_AMOUNT) <= MIN_RADIUS Or (.sInnerRadius - .iAniMax * MOVE_AMOUNT) <= MIN_RADIUS Then GoTo ReGen:
            .sAction = rad_shrink
        Case Is = 5
            If (.sInnerRadius - .iAniMax * MOVE_AMOUNT) <= MIN_RADIUS Then GoTo ReGen:
            .sAction = inn_shrink
        Case Is = 6
            If (.sInnerRadius + .iAniMax * MOVE_AMOUNT) >= MAX_RADIUS - MIN_RADIUS Then GoTo ReGen:
            .sAction = inn_grow
        Case Is = 7
            If (.sOuterRadius - .iAniMax * MOVE_AMOUNT) <= MIN_RADIUS * 2 Then GoTo ReGen:
            .sAction = out_shrink
        Case Is = 8
            If (.sOuterRadius + .iAniMax * MOVE_AMOUNT) >= MAX_RADIUS + MIN_RADIUS Then GoTo ReGen:
            .sAction = out_grow
    End Select

    .bAnimating = True
    .iAniTick = 0
End With

s.bAnimating = False
End Function

Private Function ProcessAnimation(s As Sector)
If s.bAnimating Then
    Select Case s.sAction
        Case Is = ACTIONS.d_cw
            s.sDirection = s.sDirection - MOVE_AMOUNT
        Case Is = ACTIONS.d_ccw
            s.sDirection = s.sDirection + MOVE_AMOUNT
        Case Is = ACTIONS.rad_grow
            s.sInnerRadius = s.sInnerRadius + MOVE_AMOUNT
            s.sOuterRadius = s.sOuterRadius + MOVE_AMOUNT
        Case Is = ACTIONS.rad_shrink
            s.sInnerRadius = s.sInnerRadius - MOVE_AMOUNT
            s.sOuterRadius = s.sOuterRadius - MOVE_AMOUNT
        Case Is = ACTIONS.inn_grow
            s.sInnerRadius = s.sInnerRadius + MOVE_AMOUNT
        Case Is = ACTIONS.inn_shrink
            s.sInnerRadius = s.sInnerRadius - MOVE_AMOUNT
        Case Is = ACTIONS.out_grow
            s.sOuterRadius = s.sOuterRadius + MOVE_AMOUNT
        Case Is = ACTIONS.out_shrink
            s.sOuterRadius = s.sOuterRadius - MOVE_AMOUNT
    End Select

    If s.iAniTick >= s.iAniMax Then
        ShiftEnergy s
    Else
        s.iAniTick = s.iAniTick + 1
    End If
    cur_Moving = cur_Moving + 1
End If

FormatAngle s.sDirection

If s.sOuterRadius <= MIN_RADIUS + MIN_THICKNESS Then s.iAniTick = s.iAniMax: s.sOuterRadius = MIN_RADIUS + MIN_THICKNESS
If s.sOuterRadius >= MAX_RADIUS Then s.iAniTick = s.iAniMax: s.sOuterRadius = MAX_RADIUS
If s.sInnerRadius <= MIN_RADIUS Then s.iAniTick = s.iAniMax: s.sInnerRadius = MIN_RADIUS
If s.sInnerRadius >= MAX_RADIUS - MIN_THICKNESS Then s.iAniTick = s.iAniMax: s.sInnerRadius = MAX_RADIUS - MIN_THICKNESS

If s.sInnerRadius > s.sOuterRadius Then
    Dim t1 As Single
    t1 = s.sInnerRadius
    s.sInnerRadius = s.sOuterRadius
    s.sOuterRadius = t1
End If

If s.sOuterRadius - s.sInnerRadius > MAX_THICKNESS Then
    s.sInnerRadius = s.sOuterRadius - MAX_THICKNESS
End If

If s.sOuterRadius - s.sInnerRadius < MIN_THICKNESS Then
    s.sOuterRadius = s.sInnerRadius + MIN_THICKNESS
End If
End Function

Private Sub EnterLoop()
Do While frmMain.Visible
    AddMoving MAX_MOVING - cur_Moving
    cur_Moving = 0
    For i = 1 To MAX_SECTORS
        On Error GoTo wait:
        If Sectors(i).bAnimating Then DrawSector Sectors(i), frmMain.BackColor
        ProcessAnimation Sectors(i)
        DrawSector Sectors(i), vbBlack
    Next i
    DrawText
    
    fps = fps + 1
    FPSLimit.LimitFrames 60
    DoEvents
Loop

Exit Sub

wait:

If Err.Number = 9 Then
    AddSector
Else
    MsgBox Err.Number & ", " & Err.Description
    Exit Sub
End If

Sleep 5
DoEvents
Resume Next
End Sub

Private Sub FormatAngle(ByRef ang As Single)
If ang < 0 Then ang = 360 + ang
If ang > 360 Then ang = ang - 360
End Sub

Public Function AddSector()
Sleep 5
ReDim Preserve Sectors(MAX_SECTORS) As Sector
With Sectors(MAX_SECTORS)
    .bAnimating = False
    .iAniMax = RealRnd(5, 50)
    .iAniTick = 0
    .sAction = d_cw
    .sDirection = RealRnd(0, 360)
    .sInnerRadius = RealRnd(MIN_RADIUS, MAX_RADIUS)
    .sOuterRadius = RealRnd(MIN_RADIUS, MAX_RADIUS)
    .sWidthAngle = RealRnd(15, 60)
End With
DoEvents
End Function

Public Function RemoveSector()
Sleep 5
DrawSector Sectors(MAX_SECTORS + 1), frmMain.BackColor
If Sectors(MAX_SECTORS + 1).bAnimating Then cur_Moving = cur_Moving - 1
DoEvents
ReDim Preserve Sectors(MAX_SECTORS) As Sector
DoEvents
End Function

Public Function AddMoving(ByVal x As Long)
Dim t1 As Long
If x < 0 Then RemoveMoving Abs(x)
For i = 1 To x
ReGen:
    t1 = RealRnd(1, MAX_SECTORS)
    If Sectors(t1).bAnimating Then GoTo ReGen Else Sectors(t1).bAnimating = True
    DoEvents
Next i
End Function

Public Function RemoveMoving(ByVal x As Long)
Dim t1 As Long
For i = 1 To x
ReGen:
    t1 = RealRnd(1, MAX_SECTORS)
    If Not Sectors(t1).bAnimating Then GoTo ReGen Else Sectors(t1).bAnimating = False
    DoEvents
Next i
End Function

Public Function DrawText()
With picStat
    .Cls

    .CurrentX = 5: .CurrentY = 5
    picStat.Print "[Q/A] Sectors: " & MAX_SECTORS
    'picStat.Print "Sectors: " & MAX_SECTORS
    .CurrentX = 5: .CurrentY = 20
    picStat.Print "[W/S] Moving Sectors: " & MAX_MOVING
    
    
    .CurrentX = 5: .CurrentY = 40
    picStat.Print "[E/D] Movement Speed: " & MOVE_AMOUNT
    
    
    .CurrentX = 5: .CurrentY = 60
    picStat.Print "[R/F] Max Radius: " & MAX_RADIUS
    .CurrentX = 5: .CurrentY = 75
    picStat.Print "[T/G] Min Radius: " & MIN_RADIUS
    
    
    .CurrentX = 5: .CurrentY = 95
    picStat.Print "[Y/H] Max Thickness: " & MAX_THICKNESS
    .CurrentX = 5: .CurrentY = 110
    picStat.Print "[U/J] Min Thickness: " & MIN_THICKNESS
End With
End Function
