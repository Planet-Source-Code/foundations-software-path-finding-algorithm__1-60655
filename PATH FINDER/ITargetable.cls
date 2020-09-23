VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITargetable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetRow() As Long
End Function

Public Function GetCol() As Long
End Function

Public Sub UpdateDirection(ByVal Direction As Long)
End Sub

Public Sub MovePlayer()
End Sub

Public Sub PaintCharacter(ByRef GamePicture As PictureBox, ByRef Image As PictureBox, ByRef ImageMask As PictureBox)
End Sub

