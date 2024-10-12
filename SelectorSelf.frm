VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectorSelf 
   Caption         =   "Select animation type"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "SelectorSelf.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectorSelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
SelfAnimateEntryExit.Show
End Sub

Private Sub CommandButton2_Click()
SelfAnimateEmphasis.Show
End Sub

Private Sub CommandButton3_Click()
Hide
End Sub

Private Sub Media_Click()
SelfAnimateMedia.Show
End Sub

Private Sub UserForm_Initialize()
If ActiveWindow.Selection.ShapeRange(1).Type = msoMedia Or msoWebVideo Then Media.Enabled = True
End Sub

