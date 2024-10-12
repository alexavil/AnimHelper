VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelfAnimateMultipleEntryExit 
   Caption         =   "Choose an animation"
   ClientHeight    =   1548
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "SelfAnimateMultipleEntryExit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelfAnimateMultipleEntryExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Select Case ComboBox1.Value
    Case "Appear"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectAppear, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Blinds"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectBlinds, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Box"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectBox, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Checkerboard"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectCheckerboard, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Circle"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectCircle, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Diamond"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectDiamond, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Dissolve In"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectDissolve, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Fly In"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectFly, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Peek In"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectPeek, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Plus"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectPlus, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Random Bars"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectRandomBars, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Spilt"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectSplit, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Strips"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectStrips, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Wedge"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectWedge, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Wheel"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectWheel, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Wipe"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectWipe, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Expand"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectExpand, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Fade"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectFade, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Swivel"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectFadedSwivel, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Zoom"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectFadedZoom, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Basic Zoom"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectZoom, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Center Revolve"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectCenterRevolve, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Compress"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectCompress, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Grow & Turn"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectGrowAndTurn, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Rise Up"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectRiseUp, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Spinner"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectSpinner, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Stretch"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectStretch, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Basic Swivel"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectSwivel, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Boomerang"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectBoomerang, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Bounce"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectBounce, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Credits"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectCredits, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Drop"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectDrop, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Flip"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectFlip, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Float"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectFloat, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Pinwheel"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectPinwheel, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Spiral In"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectSpiral, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Whip"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectWhip, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    End Select
    SelfAnimateMultipleEntryExit.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Appear", "Blinds", "Box", "Checkerboard", "Circle", "Diamond", "Dissolve In", "Fly In", "Peek In", "Plus", "Random Bars", "Split", "Strips", "Wedge", "Wheel", "Wipe", "Expand", "Fade", "Swivel", "Zoom", "Basic Zoom", "Center Revolve", "Compress", "Grow & Turn", "Rise Up", "Spinner", "Stretch", "Basic Swivel", "Boomerang", "Bounce", "Credits", "Drop", "Flip", "Float", "Pinwheel", "Spiral In", "Whip")
End Sub

