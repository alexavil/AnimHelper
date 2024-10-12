VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MultipleWithATriggerEmphasis 
   Caption         =   "Choose an animation"
   ClientHeight    =   1572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "MultipleWithATriggerEmphasis.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MultipleWithATriggerEmphasis"
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
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectChangeFillColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Font Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectChangeFontColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Grow (Shrink)"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectGrowShrink, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Line Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectChangeLineColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Spin"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectSpin, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Transparency"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectTransparency, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Bold Flash"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectBoldFlash, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Brush Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectBrushOnColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Complimentary Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectComplementaryColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Contrasting Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectContrastingColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Darken"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectDarken, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Desaturate"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectDesaturate, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Lighten"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectLighten, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Pulse"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectPulse, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Underline"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectBrushOnUnderline, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Grow with Color"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectGrowWithColor, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Shimmer"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectShimmer, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Teeter"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectTeeter, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Bold Reveal"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectBoldReveal, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    Case "Wave"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectWave, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
Next sh
    End Select
    MultipleWithATriggerEmphasis.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Fill Color", "Font Color", "Grow (Shrink)", "Line Color", "Spin", "Transparency", "Bold Flash", "Brush Color", "Complimentary Color", "Contrasting Color", "Darken", "Desaturate", "Lighten", "Pulse", "Underline", "Grow with Color", "Shimmer", "Teeter", "Bold Reveal", "Wave")
End Sub



