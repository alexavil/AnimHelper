VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AnimateWithATriggerEmphasis 
   Caption         =   "Choose an animation"
   ClientHeight    =   1572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "AnimateWithATriggerEmphasis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AnimateWithATriggerEmphasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton1_Click()
Select Case ComboBox1.Value
    Case "Fill Color"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectChangeFillColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Font Color"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectChangeFontColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Grow (Shrink)"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectGrowShrink, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Line Color"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectChangeLineColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Spin"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectSpin, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Transparency"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectTransparency, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Bold Flash"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBoldFlash, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Brush Color"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBrushOnColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Complimentary Color"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectComplementaryColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Contrasting Color"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectContrastingColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Darken"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectDarken, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Desaturate"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectDesaturate, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Lighten"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectLighten, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Pulse"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectPulse, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Underline"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBrushOnUnderline, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Grow with Color"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectGrowWithColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Shimmer"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectShimmer, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Teeter"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectTeeter, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Bold Reveal"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectBoldReveal, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    Case "Wave"
    
    
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectWave, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))

    End Select
    AnimateWithATriggerEmphasis.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Fill Color", "Font Color", "Grow (Shrink)", "Line Color", "Spin", "Transparency", "Bold Flash", "Brush Color", "Complimentary Color", "Contrasting Color", "Darken", "Desaturate", "Lighten", "Pulse", "Underline", "Grow with Color", "Shimmer", "Teeter", "Bold Reveal", "Wave")
End Sub




