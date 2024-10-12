VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelfAnimateMultipleEmphasis 
   Caption         =   "Choose an animation"
   ClientHeight    =   1572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "SelfAnimateMultipleEmphasis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelfAnimateMultipleEmphasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton1_Click()
Select Case ComboBox1.Value
    Case "Fill Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectChangeFillColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Font Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectChangeFontColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Grow (Shrink)"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectGrowShrink, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Line Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectChangeLineColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Spin"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectSpin, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Transparency"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectTransparency, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Bold Flash"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectBoldFlash, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Brush Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectBrushOnColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Complimentary Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectComplementaryColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Contrasting Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectContrastingColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Darken"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectDarken, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Desaturate"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectDesaturate, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Lighten"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectLighten, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Pulse"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectPulse, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Underline"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectBrushOnUnderline, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Grow with Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectGrowWithColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Shimmer"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectShimmer, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Teeter"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectTeeter, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Bold Reveal"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectBoldReveal, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    Case "Wave"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectWave, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
Next sh
    End Select
    SelfAnimateMultipleEmphasis.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Fill Color", "Font Color", "Grow (Shrink)", "Line Color", "Spin", "Transparency", "Bold Flash", "Brush Color", "Complimentary Color", "Contrasting Color", "Darken", "Desaturate", "Lighten", "Pulse", "Underline", "Grow with Color", "Shimmer", "Teeter", "Bold Reveal", "Wave")
End Sub


