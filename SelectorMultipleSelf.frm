VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectorMultipleSelf 
   Caption         =   "Select animation type"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "SelectorMultipleSelf.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectorMultipleSelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
SelfAnimateMultipleEntryExit.Show
End Sub

Private Sub CommandButton2_Click()
SelfAnimateMultipleEmphasis.Show
End Sub

Private Sub CommandButton3_Click()
Hide
End Sub

Private Sub Media_Click()
SelfAnimateMultipleMedia.Show
End Sub

Private Sub UserForm_Initialize()
For Each sh In ActiveWindow.Selection.ShapeRange
If sh.Type = msoMedia Or msoWebVideo Then Media.Enabled = True
Next sh
End Sub
