VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComputerPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ITargetable

Private CharacterX As Long
Private CharacterY As Long
Private StoredDirection As Long
Private CurrentDirection As Long
Private Counter As Long
Private GameServer As IGameServer

Private CurrentRow As Long
Private CurrentCol As Long
Private PreviousRow As Long
Private PreviousCol As Long

Private CurrentSpriteImagePath As String
Private CurrentSpriteMaskImagePath As String
Private Identifier As String

Private Sub Class_Initialize()
    Counter = 1
    Set GameServer = frmMain
    Identifier = "MAN"
    CurrentSpriteImagePath = App.Path & "\IMAGES\MANDOWN1.BMP"
    CurrentSpriteMaskImagePath = App.Path & "\IMAGES\MASKMANDOWN1.BMP"

    StoredDirection = vbKeyLeft
    CurrentDirection = vbKeyLeft
    CharacterX = 320
    CharacterY = 320
    PreviousRow = 10
    PreviousCol = 10
    CurrentRow = 10
    CurrentCol = 10
End Sub

Private Function ITargetable_GetCol() As Long
    ITargetable_GetCol = CLng(CharacterX / 32) + 1
End Function

Private Function ITargetable_GetRow() As Long
    ITargetable_GetRow = CLng(CharacterY / 32) + 1
End Function

Private Sub ITargetable_MovePlayer()
    Dim dir As Long
    
    dir = SearchAnalyzer.SuggestDirection(Me)
'    If (SearchAnalyzer.GetOpenRouteCount(CurrentRow, CurrentCol) > 1) Then
'        If (IsAbruptTurn(dir)) Then
'            dir = AssignRandomMove(dir)
'        End If
'    End If
    Call Move(dir)
End Sub

Private Function IsAbruptTurn(ByRef dir As Long) As Boolean
    Dim abrupt As Boolean
    Dim rowIncr As Integer
    Dim colIncr As Integer

    abrupt = False
    rowIncr = 0
    colIncr = 0

    If (dir = vbKeyUp) Then
        rowIncr = rowIncr - 1
    ElseIf (dir = vbKeyDown) Then
        rowIncr = rowIncr + 1
    ElseIf (dir = vbKeyLeft) Then
        colIncr = colIncr - 1
    Else
        If (dir = vbKeyRight) Then
            colIncr = colIncr + 1
        End If
    End If
    
    If (CurrentRow + rowIncr = PreviousRow And CurrentCol + colIncr = PreviousCol) Then
        abrupt = True
    End If
    IsAbruptTurn = abrupt
End Function

Public Function AssignRandomMove(ByRef excluded As Long) As Long
    Dim result As Long
    Dim X As Long
    Dim directions(4) As Long
    
    result = 0
    directions(1) = vbKeyUp
    directions(2) = vbKeyDown
    directions(3) = vbKeyLeft
    directions(4) = vbKeyRight
    For X = 1 To 4
        If (directions(X) <> excluded) Then
            If (IsOpenCell(directions(X))) Then
                result = directions(X)
            End If
        End If
    Next X
    AssignRandomMove = result
End Function

Private Sub ITargetable_PaintCharacter(GamePicture As PictureBox, Image As PictureBox, ImageMask As PictureBox)
    Set ImageMask.Picture = LoadPicture(CurrentSpriteMaskImagePath)
    Set Image.Picture = LoadPicture(CurrentSpriteImagePath)
    Call BitBlt(GamePicture.hdc, CharacterX, CharacterY, ImageMask.ScaleWidth, ImageMask.ScaleHeight, ImageMask.hdc, 0, 0, vbSrcAnd)
    Call BitBlt(GamePicture.hdc, CharacterX, CharacterY, Image.ScaleWidth, Image.ScaleHeight, Image.hdc, 0, 0, vbSrcPaint)
End Sub

Public Function IsOpenCell(ByRef dir As Long) As Boolean
    Dim rowIncr As Integer
    Dim colIncr As Integer
    
    rowIncr = 0
    colIncr = 0
    
    If (dir = vbKeyUp) Then
        rowIncr = rowIncr - 1
    ElseIf (dir = vbKeyDown) Then
        rowIncr = rowIncr + 1
    ElseIf (dir = vbKeyLeft) Then
        colIncr = colIncr - 1
    Else
        If (dir = vbKeyRight) Then
            colIncr = colIncr + 1
        End If
    End If

    IsOpenCell = (SearchAnalyzer.IsOpenRoute(CurrentRow + rowIncr, CurrentCol + colIncr))
End Function

Private Sub ITargetable_UpdateDirection(ByVal Direction As Long)
    StoredDirection = Direction
End Sub

Private Sub Move(ByVal dir As Long)
    Dim mStep As Integer

    mStep = 1
    
    StoredDirection = dir
    
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

