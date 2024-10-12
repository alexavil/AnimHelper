VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelfAnimateEmphasis 
   Caption         =   "Choose an animation"
   ClientHeight    =   1572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "SelfAnimateEmphasis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelfAnimateEmphasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton1_Click()
Select Case ComboBox1.Value
    Case "Fill Color"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectChangeFillColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Font Color"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectChangeFontColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Grow/Shrink"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectGrowShrink, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Line Color"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectChangeLineColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Spin"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectSpin, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Transparency"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectTransparency, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Bold Flash"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBoldFlash, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Brush Color"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBrushOnColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Complimentary Color"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectComplementaryColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Contrasting Color"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectContrastingColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Darken"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectDarken, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Desaturate"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectDesaturate, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Lighten"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectLighten, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Pulse"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectPulse, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Underline"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBrushOnUnderline, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Grow with Color"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectGrowWithColor, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Shimmer"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectShimmer, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Teeter"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectTeeter, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Bold Reveal"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBoldReveal, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    Case "Wave"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectWave, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)

    End Select
    SelfAnimateEmphasis.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Fill Color", "Font Color", "Grow/Shrink", "Line Color", "Spin", "Transparency", "Bold Flash", "Brush Color", "Complimentary Color", "Contrasting Color", "Darken", "Desaturate", "Lighten", "Pulse", "Underline", "Grow with Color", "Shimmer", "Teeter", "Bold Reveal", "Wave")
End Sub

