Attribute VB_Name = "Module1"
Sub AnimHelper()
If ActiveWindow.Selection.Type = 0 Then
Msg = "Select an object to animate!"
Title = "AnimHelper"
Style = vbCritical
    Response = MsgBox(Msg, Style, Title)
End If
If ActiveWindow.Selection.Type = 2 Then
SelectorSelf.Show vbModal
End If
End Sub

Sub TriggerHelper()
If ActiveWindow.Selection.Type = 0 Then
Msg = "Select an object to animate!"
Title = "AnimHelper"
Style = vbCritical
    Response = MsgBox(Msg, Style, Title)
End If
If ActiveWindow.Selection.Type = 2 Then
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        Msg = "Select a trigger for the animation!"
        Title = "AnimHelper"
        Style = vbCritical
        Response = MsgBox(Msg, Style, Title)
    Else
    Trigger.Show vbModal
    End If
End If
End Sub

Sub TriggerHelperMulti()
If ActiveWindow.Selection.Type = 0 Then
Msg = "Select at least one object to animate!"
Title = "AnimHelper"
Style = vbCritical
    Response = MsgBox(Msg, Style, Title)
End If
If ActiveWindow.Selection.Type = 2 Then
    If ActiveWindow.Selection.ShapeRange.Count < 2 Then
        Msg = "Select a trigger for the animation!"
        Title = "AnimHelper"
        Style = vbCritical
        Response = MsgBox(Msg, Style, Title)
    End If
    If ActiveWindow.Selection.ShapeRange.Count >= 2 Then
    TriggerMultiple.Show vbModal
    End If
End If
End Sub

Sub AnimHelperMulti()
If ActiveWindow.Selection.Type = 0 Then
Msg = "Select at least one object to animate!"
Title = "AnimHelper"
Style = vbCritical
    Response = MsgBox(Msg, Style, Title)
End If
If ActiveWindow.Selection.Type = 2 Then
SelectorMultipleSelf.Show vbModal
End If
End Sub


Sub Help()
Msg = "Self-Animate: Select an object to assign it an animation and set the trigger to itself." + vbCrLf + "Animate with a Trigger: Select an object, and then a trigger. The animation will be assigned to your first object with the second one as the trigger." + vbCrLf + "Self-Animate Multiple: Each selected object will be assigned an animation with the trigger set to itself." + vbCrLf + "Animate Multiple with a Trigger: Select multiple objects, and then a trigger. The animation will be assigned to all of your selected objects with the last one as the trigger."
Title = "AnimHelper"
Style = vbQuestion
    Response = MsgBox(Msg, Style, Title)
End Sub
