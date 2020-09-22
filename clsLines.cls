VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RGB_COLOR
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type RECT
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Private cMaxColor As RGB_COLOR, cMinColor As RGB_COLOR, cStepColor As RGB_COLOR
Private cData() As Long
Private cStepX As Long, cStepY As Long
Private Const cConvX As Long = 15, cConvY As Long = 15
Private cRange As RECT
Private cShapes As Long, cPoints As Long
Private cUniColor As Boolean, cFixColorStep As Boolean, cTarget As Form
Private u As Long, v As Long, w As Long
Private Colour As Long
Private UDColor As Boolean
Private UDR As Long, UDG As Long, UDB As Long
Private UDRS As Long, UDGS As Long, UDBS As Long

Public Sub Init( _
                DrawTarget As Form, TotalShapes As Long, TotalPoints As Long, _
                Optional X1 As Long = -1, Optional Y1 As Long = -1, _
                Optional X2 As Long = -1, Optional Y2 As Long = -1, _
                Optional MinR As Long = 0, Optional MinG As Long = 0, Optional MinB As Long = 0, _
                Optional MaxR As Long = 255, Optional MaxG As Long = 255, Optional MaxB As Long = 255, _
                Optional StepR As Long = 5, Optional StepG As Long = 5, Optional StepB As Long = 5, _
                Optional StepX As Long = 10, Optional StepY As Long = 10, _
                Optional UseUniColor As Boolean = True, Optional FixColorStep As Boolean = False, Optional UserDefineColor As Boolean = False, _
                Optional UserDefineR As Long = -1, Optional UserDefineG As Long = -1, Optional UserDefineB As Long = -1, _
                Optional UserDefineRStep As Long = -1, Optional UserDefineGStep As Long = -1, Optional UserDefineBStep = -1 _
               )

    With cMaxColor
        .Red = MaxR
        .Green = MaxG
        .Blue = MaxB
    End With
    With cMinColor
        .Red = MinR
        .Green = MinG
        .Blue = MinB
    End With
    With cStepColor
        .Red = StepR
        .Green = StepG
        .Blue = StepB
    End With
    With cRange
        If X1 <> -1 Then .X1 = X1 Else .X1 = 0
        If X2 <> -1 Then .X2 = X2 Else .X2 = Screen.Width / cConvX
        If Y1 <> -1 Then .Y1 = Y1 Else .Y1 = 0
        If Y2 <> -1 Then .Y2 = Y2 Else .Y2 = Screen.Height / cConvY
    End With
    cStepX = StepX: cStepY = StepY
    cShapes = TotalShapes: cPoints = TotalPoints
    cUniColor = UseUniColor
    cFixColorStep = FixColorStep
    UDColor = UserDefineColor
    UDR = UserDefineR
    UDG = UserDefineG
    UDB = UserDefineB
    UDRS = UserDefineRStep
    UDGS = UserDefineGStep
    UDBS = UserDefineBStep
    Set cTarget = DrawTarget
    ReDim cData(cShapes - 1, cPoints - 1, 9)
    Generate
End Sub

Public Sub Run()
    Draw LBound(cData, 1), vbBlack
    Move
    Draw UBound(cData, 1)
End Sub

Public Sub Repaint()
    For u = LBound(cData, 1) To UBound(cData, 1)
        Draw u
    Next
End Sub

Private Sub Generate()
    Randomize Timer
    u = LBound(cData, 1)
    For v = LBound(cData, 2) To UBound(cData, 2)
        cData(u, v, 0) = Int(Rnd(1) * (cRange.X2 - cRange.X1)) + cRange.X1
        cData(u, v, 1) = Int(Rnd(1) * (cRange.Y2 - cRange.Y1)) + cRange.Y1
        cData(u, v, 2) = Int(Rnd(1) * cStepX)
        cData(u, v, 3) = Int(Rnd(1) * cStepY)
        If Not UDColor Then
            cData(u, v, 4) = Int(Rnd(1) * (cMaxColor.Red - cMinColor.Red)) + cMinColor.Red
            cData(u, v, 5) = Int(Rnd(1) * (cMaxColor.Green - cMinColor.Green)) + cMinColor.Green
            cData(u, v, 6) = Int(Rnd(1) * (cMaxColor.Blue - cMinColor.Blue)) + cMinColor.Blue
        Else
            cData(u, v, 4) = UDR
            cData(u, v, 5) = UDG
            cData(u, v, 6) = UDB
        End If
        If cFixColorStep = False Then
            cData(u, v, 7) = Int(Rnd(1) * cStepColor.Red)
            cData(u, v, 8) = Int(Rnd(1) * cStepColor.Green)
            cData(u, v, 9) = Int(Rnd(1) * cStepColor.Blue)
        Else
            cData(u, v, 7) = UDRS
            cData(u, v, 8) = UDGS
            cData(u, v, 9) = UDBS
        End If
    Next
    For u = LBound(cData, 1) + 1 To UBound(cData, 1)
        MoveShape u
    Next
End Sub

Private Sub Move()
    For u = LBound(cData, 1) To UBound(cData, 1) - 1
        For v = LBound(cData, 2) To UBound(cData, 2)
            For w = LBound(cData, 3) To UBound(cData, 3)
                cData(u, v, w) = cData(u + 1, v, w)
            Next
        Next
    Next
    u = UBound(cData, 1)
    MoveShape u
End Sub

Private Sub MoveShape(ShapeIndex As Long)
    For v = LBound(cData, 2) To UBound(cData, 2)
        If cData(u - 1, v, 0) + cData(u - 1, v, 2) > cRange.X2 Or cData(u - 1, v, 0) + cData(u - 1, v, 2) < cRange.X1 Then cData(u, v, 2) = -cData(u - 1, v, 2) Else cData(u, v, 2) = cData(u - 1, v, 2)
        If cData(u - 1, v, 1) + cData(u - 1, v, 3) > cRange.Y2 Or cData(u - 1, v, 1) + cData(u - 1, v, 3) < cRange.Y1 Then cData(u, v, 3) = -cData(u - 1, v, 3) Else cData(u, v, 3) = cData(u - 1, v, 3)
        If cData(u - 1, v, 4) + cData(u - 1, v, 7) > cMaxColor.Red Or cData(u - 1, v, 4) + cData(u - 1, v, 7) < cMinColor.Red Then cData(u, v, 7) = -cData(u - 1, v, 7) Else cData(u, v, 7) = cData(u - 1, v, 7)
        If cData(u - 1, v, 5) + cData(u - 1, v, 8) > cMaxColor.Green Or cData(u - 1, v, 5) + cData(u - 1, v, 8) < cMinColor.Green Then cData(u, v, 8) = -cData(u - 1, v, 8) Else cData(u, v, 8) = cData(u - 1, v, 8)
        If cData(u - 1, v, 6) + cData(u - 1, v, 9) > cMaxColor.Blue Or cData(u - 1, v, 6) + cData(u - 1, v, 9) < cMinColor.Blue Then cData(u, v, 9) = -cData(u - 1, v, 9) Else cData(u, v, 9) = cData(u - 1, v, 9)
        cData(u, v, 0) = cData(u - 1, v, 0) + cData(u, v, 2)
        cData(u, v, 1) = cData(u - 1, v, 1) + cData(u, v, 3)
        cData(u, v, 4) = cData(u - 1, v, 4) + cData(u, v, 7)
        cData(u, v, 5) = cData(u - 1, v, 5) + cData(u, v, 8)
        cData(u, v, 6) = cData(u - 1, v, 6) + cData(u, v, 9)
    Next
End Sub

Private Sub Draw(ShapeIndex As Long, Optional Color As Long = -1)
    u = ShapeIndex
    v = LBound(cData, 2)
    If Color = -1 Then Colour = RGB(cData(u, v, 4), cData(u, v, 5), cData(u, v, 6)) Else Colour = Color
    cTarget.PSet (cData(u, v, 0) * cConvX, cData(u, v, 1) * cConvY), Colour
    For v = LBound(cData, 2) + 1 To UBound(cData, 2)
        If Not cUniColor Then Colour = RGB(cData(u, v, 4), cData(u, v, 5), cData(u, v, 6))
        cTarget.Line -(cData(u, v, 0) * cConvX, cData(u, v, 1) * cConvY), Colour
    Next
    v = LBound(cData, 2)
    If Not cUniColor Then Colour = RGB(cData(u, v, 4), cData(u, v, 5), cData(u, v, 6))
    cTarget.Line -(cData(u, v, 0) * cConvX, cData(u, v, 1) * cConvY), Colour
End Sub
