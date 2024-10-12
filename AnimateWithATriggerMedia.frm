VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AnimateWithATriggerMedia 
   Caption         =   "Choose an animation"
   ClientHeight    =   1572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "AnimateWithATriggerMedia.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AnimateWithATriggerMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub CommandButton1_Click()
Select Case ComboBox1.Value
    Case "Pause"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectMediaPause, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))
    Case "Play"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectMediaPlay, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))
    Case "Stop"
        Set Effect = ActiveWindow.Selection.SlideRange(1).TimeLine.InteractiveSequences.Add.AddTriggerEffect(ActiveWindow.Selection.ShapeRange(1), msoAnimEffectMediaStop, msoAnimTriggerOnShapeClick, ActiveWindow.Selection.ShapeRange(2))
    End Select
    AnimateWithATriggerMedia.Hide
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("Pause", "Play", "Stop")
End Sub




