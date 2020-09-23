VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHumanPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Implements ITargetable

Private CharacterX As Long
Private CharacterY As Long
Private Counter As Integer

Private StoredDirection As Long
Private CurrentDirection As Long
Private CurrentRow As Long
Private CurrentCol As Long
Private GameServer As IGameServer
Private CurrentSpriteImagePath As String
Private CurrentSpriteMaskImagePath As String
Private Identifier As String

Private Sub Class_Initialize()
    CharacterX = 32
    CharacterY = 32
    Counter = 1

    Identifier = "HERO"
    CurrentSpriteImagePath = ""
    CurrentSpriteMaskImagePath = ""
    CurrentDirection = vbKeyRight
    StoredDirection = vbKeyRight
    CurrentRow = 2
    CurrentCol = 2
    Set GameServer = frmMain
End Sub

Private Function ITargetable_GetCol() As Long
    ITargetable_GetCol = CLng(CharacterX / 32) + 1
End Function

Private Function ITargetable_GetRow() As Long
    ITargetable_GetRow = CLng(CharacterY / 32) + 1
End Function

Private Sub ITargetable_MovePlayer()
    Dim mStep As Integer

    mStep = 1
    If (StoredDirection = vbKeyUp And Not (GameServer.CollisionDetected(CharacterX, CharacterY - mStep))) Then
        CurrentDirection = vbKeyUp
    ElseIf (StoredDirection = vbKeyDown And Not (GameServer.CollisionDetected(CharacterX, CharacterY + mStep))) Then
        CurrentDirection = vbKeyDown
    ElseIf (StoredDirection = vbKeyLeft And Not (GameServer.CollisionDetected(CharacterX - mStep, CharacterY))) Then
        CurrentDirection = vbKeyLeft
    Else
        If (StoredDirection = vbKeyRight And Not (GameServer.CollisionDetected(CharacterX + mStep, CharacterY))) Then
            CurrentDirection = vbKeyRight
        End If
    End If

    If (CurrentDirection = vbKeyUp) Then
        If Not (GameServer.CollisionDetected(CharacterX, CharacterY - mStep)) Then
            Counter = IIf(Counter = 3, 1, Counter + 1)
            CharacterY = CharacterY - mStep
        Else
            Counter = 2
        End If
        CurrentSpriteImagePath = App.Path & "\IMAGES\" & Identifier & "UP" & Counter & ".BMP"
        CurrentSpriteMaskImagePath = App.Path & "\IMAGES\MASK" & Identifier & "UP" & Counter & ".BMP"
    ElseIf (CurrentDirection = vbKeyDown) Then
        If Not (GameServer.CollisionDetected(CharacterX, CharacterY + mStep)) Then
            Counter = IIf(Counter = 3, 1, Counter + 1)
            CharacterY = CharacterY + mStep
        Else
            Counter = 2
        End If
        CurrentSpriteImagePath = App.Path & "\IMAGES\" & Identifier & "DOWN" & Counter & ".BMP"
        CurrentSpriteMaskImagePath = App.Path & "\IMAGES\MASK" & Identifier & "DOWN" & Counter & ".BMP"
    ElseIf (CurrentDirection = vbKeyLeft) Then
        If Not (GameServer.CollisionDetected(CharacterX - mStep, CharacterY)) Then
            Counter = IIf(Counter = 3, 1, Counter + 1)
            CharacterX = CharacterX - mStep
        Else
            Counter = 2
        End If
        CurrentSpriteImagePath = App.Path & "\IMAGES\" & Identifier & "LEFT" & Counter & ".BMP"
        CurrentSpriteMaskImagePath = App.Path & "\IMAGES\MASK" & Identifier & "LEFT" & Counter & ".BMP"
    Else
        If (CurrentDirection = vbKeyRight) Then
            If Not (GameServer.CollisionDetected(CharacterX + mStep, CharacterY)) Then
                Counter = IIf(Counter = 3, 1, Counter + 1)
                CharacterX = CharacterX + mStep
            Else
                Counter = 2
            End If
            CurrentSpriteImagePath = App.Path & "\IMAGES\" & Identifier & "RIGHT" & Counter & ".BMP"
            CurrentSpriteMaskImagePath = App.Path & "\IMAGES\MASK" & Identifier & "RIGHT" & Counter & ".BMP"
        End If
    End If
    
    CurrentRow = ITargetable_GetRow
    CurrentCol = ITargetable_GetCol
End Sub

Private Sub ITargetable_PaintCharacter(GamePicture As PictureBox, Image As PictureBox, ImageMask As PictureBox)
    Set ImageMask.Picture = LoadPicture(CurrentSpriteMaskImagePath)
    Set Image.Picture = LoadPicture(CurrentSpriteImagePath)
    Call BitBlt(GamePicture.hdc, CharacterX, CharacterY, ImageMask.ScaleWidth, ImageMask.ScaleHeight, ImageMask.hdc, 0, 0, vbSrcAnd)
    Call BitBlt(GamePicture.hdc, CharacterX, CharacterY, Image.ScaleWidth, Image.ScaleHeight, Image.hdc, 0, 0, vbSrcPaint)
End Sub

Private Sub ITargetable_UpdateDirection(ByVal Direction As Long)
    StoredDirection = Direction
End Sub

Private Sub ITargetable_UpdatePosition(ByVal Location As CLocation)
    CurrentRow = Location.Row
    CurrentCol = Location.Col
End Sub
