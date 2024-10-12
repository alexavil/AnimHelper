VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelfAnimateMultipleMedia 
   Caption         =   "Choose an animation"
   ClientHeight    =   1572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "SelfAnimateMultipleMedia.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelfAnimateMultipleMedia"
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
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectMediaPause, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
    Next sh
    Case "Play"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectMediaPlay, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
    Next sh
    Case "Stop"
    For Each sh In ActiveWindow.Selection.ShapeRange
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddEffect(sh, msoAnimEffectMediaStop, msoAnimateLevelNone, msoAnimTriggerOnShapeClick)
    Next sh
    End Select
    SelfAnimateMultipleMedia.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Pause", "Play", "Stop")
End Sub


