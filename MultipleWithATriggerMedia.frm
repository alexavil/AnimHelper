VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MultipleWithATriggerMedia 
   Caption         =   "Choose an animation"
   ClientHeight    =   1572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "MultipleWithATriggerMedia.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MultipleWithATriggerMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton1_Click()
Select Case ComboBox1.Value
    Case "Pause"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectMediaPause, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
    Next sh
    Case "Play"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectMediaPlay, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
    Next sh
    Case "Stop"
    For Each sh In ActiveWindow.Selection.ShapeRange
    If sh.Id = ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count).Id Then Exit For
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(sh, msoAnimEffectMediaStop, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(ActiveWindow.Selection.ShapeRange.Count))
    Next sh
    End Select
    MultipleWithATriggerMedia.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Pause", "Play", "Stop")
End Sub



