VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MultipleWithATriggerEntryExit 
   Caption         =   "Choose an animation"
   ClientHeight    =   1716
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "MultipleWithATriggerEntryExit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MultipleWithATriggerEntryExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox2_Click()

End Sub

Private Sub CommandButton1_Click()
Select Case ComboBox1.Value
    Case "Appear"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectAppear, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Blinds"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectBlinds, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Box"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectBox, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Checkerboard"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectCheckerboard, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Circle"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectCircle, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Diamond"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectDiamond, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Dissolve In"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectDissolve, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Fly In"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectFly, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Peek In"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectPeek, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Plus"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectPlus, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Random Bars"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectRandomBars, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Spilt"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectSplit, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Strips"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectStrips, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Wedge"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectWedge, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Wheel"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectWheel, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Wipe"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectWipe, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Expand"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectExpand, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Fade"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectFade, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Swivel"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectFadedSwivel, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Zoom"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectFadedZoom, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Basic Zoom"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectZoom, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Center Revolve"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectCenterRevolve, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Compress"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectCompress, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Grow & Turn"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectGrowAndTurn, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Rise Up"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectRiseUp, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Spinner"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectSpinner, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Stretch"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectStretch, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Basic Swivel"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectSwivel, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Boomerang"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectBoomerang, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Bounce"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectBounce, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Credits"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectCredits, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Drop"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectDrop, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Flip"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectFlip, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Float"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectFloat, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Pinwheel"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectPinwheel, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Spiral In"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectSpiral, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    Case "Whip"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectWhip, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
        If CheckBox1.Value = True Then Effect.Exit = True
        Next sh
    End Select
    MultipleWithATriggerEntryExit.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Appear", "Blinds", "Box", "Checkerboard", "Circle", "Diamond", "Dissolve In", "Fly In", "Peek In", "Plus", "Random Bars", "Split", "Strips", "Wedge", "Wheel", "Wipe", "Expand", "Fade", "Swivel", "Zoom", "Basic Zoom", "Center Revolve", "Compress", "Grow & Turn", "Rise Up", "Spinner", "Stretch", "Basic Swivel", "Boomerang", "Bounce", "Credits", "Drop", "Flip", "Float", "Pinwheel", "Spiral In", "Whip")
End Sub


