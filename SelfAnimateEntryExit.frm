VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelfAnimateEntryExit 
   Caption         =   "Choose an animation"
   ClientHeight    =   1548
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "SelfAnimateEntryExit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelfAnimateEntryExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Select Case ComboBox1.Value
    Case "Appear"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectAppear, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Blinds"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBlinds, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Box"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBox, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Checkerboard"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectCheckerboard, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Circle"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectCircle, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Diamond"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectDiamond, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Dissolve In"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectDissolve, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Fly In"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectFly, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Peek In"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectPeek, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Plus"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectPlus, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Random Bars"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectRandomBars, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Spilt"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectSplit, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Strips"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectStrips, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Wedge"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectWedge, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Wheel"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectWheel, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Wipe"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectWipe, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Expand"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectExpand, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Fade"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectFade, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Swivel"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectFadedSwivel, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Zoom"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectFadedZoom, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Basic Zoom"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectZoom, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Center Revolve"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectCenterRevolve, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Compress"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectCompress, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Grow & Turn"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectGrowAndTurn, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Rise Up"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectRiseUp, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Spinner"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectSpinner, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Stretch"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectStretch, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Basic Swivel"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectSwivel, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Boomerang"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBoomerang, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Bounce"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBounce, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Credits"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectCredits, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Drop"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectDrop, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Flip"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectFlip, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Float"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectFloat, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Pinwheel"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectPinwheel, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Spiral In"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectSpiral, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    Case "Whip"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectWhip, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
        If CheckBox1.Value = True Then Effect.Exit = True
    End Select
    SelfAnimateEntryExit.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Appear", "Blinds", "Box", "Checkerboard", "Circle", "Diamond", "Dissolve In", "Fly In", "Peek In", "Plus", "Random Bars", "Split", "Strips", "Wedge", "Wheel", "Wipe", "Expand", "Fade", "Swivel", "Zoom", "Basic Zoom", "Center Revolve", "Compress", "Grow & Turn", "Rise Up", "Spinner", "Stretch", "Basic Swivel", "Boomerang", "Bounce", "Credits", "Drop", "Flip", "Float", "Pinwheel", "Spiral In", "Whip")
End Sub
