VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   723
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   820
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3270
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3870
      Picture         =   "Form1.frx":0C42
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picImageMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2670
      Picture         =   "Form1.frx":0EC4
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   6210
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2100
      Top             =   6270
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   6090
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":1B06
      Top             =   900
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6870
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   180
      Width           =   2685
   End
   Begin VB.PictureBox picGame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   0
      Picture         =   "Form1.frx":1B0C
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   1
      Top             =   0
      Width           =   5760
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   0
      Picture         =   "Form1.frx":6DB4E
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5760
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Implements IGameServer

Private Counter As Long

Private HumanPlayer As ITargetable
Private ComputerPlayer As ITargetable

Private Sub Form_Load()
    Counter = 1
    Set SearchAnalyzer = New CSearchAnalyzer
    Set HumanPlayer = New CHumanPlayer
    Set ComputerPlayer = New CComputerPlayer
End Sub

Private Function IGameServer_CollisionDetected(ByVal MoveX As Long, ByVal MoveY As Long) As Boolean
    Dim c As Long
    Dim r As Long

    For r = 0 To picMask.Height - 1
        For c = 0 To picMask.Width - 1
            If (GetPixel(picMask.hdc, c, r) = vbBlack) Then
                If (GetPixel(picGrid.hdc, MoveX + c, MoveY + r) = vbBlack) Then
                    IGameServer_CollisionDetected = True
                    Exit Function
                End If
            End If
        Next
    Next
    IGameServer_CollisionDetected = False
End Function

Private Sub Timer1_Timer()
    Dim TargetTileRow
    Dim TargetTileCol
    
    Call HumanPlayer.MovePlayer
       
    Call SearchAnalyzer.PerformTargetSurveillance(HumanPlayer)
    
    If (Counter = 3) Then
        Counter = 1
        Call ComputerPlayer.MovePlayer
    Else
        Counter = Counter + 1
    End If
    
    Call UpdateImage

'    If (CharacterX Mod 32 = 0 And CharacterY Mod 32 = 0) Then
'        Call SearchAnalyzer.PerformTargetSurveillance(HumanPlayer)
'    End If
'
 '   TargetTileRow = CLng(CharacterY / 32) + 1
 '   TargetTileCol = CLng(CharacterX / 32) + 1
 '   Text1.Text = "[" & TargetTileRow & "][" & TargetTileCol & "]"
    
 '   Dim mLocation As CLocation
 '   Set mLocation = New CLocation
 '   mLocation.Row = TargetTileRow
 '   mLocation.Col = TargetTileCol
 '   Call ComputerPlayer.UpdatePosition(mLocation)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight) Then
        Call HumanPlayer.UpdateDirection(KeyCode)
    End If
End Sub

Private Sub UpdateImage()
    Call BitBlt(picGame.hdc, 0, 0, picGrid.ScaleWidth, picGrid.ScaleHeight, picGrid.hdc, 0, 0, vbSrcCopy)
    Call HumanPlayer.PaintCharacter(picGame, picImage, picImageMask)
    Call ComputerPlayer.PaintCharacter(picGame, picImage, picImageMask)
    Call picGame.Refresh
End Sub

Private Function IGameServer_GetMapElementType(ByVal Row As Long, ByVal Col As Long) As Long
    If (GetPixel(picGrid.hdc, ((Col - 1) * TILE_WIDTH), ((Row - 1) * TILE_HEIGHT)) = vbBlack) Then
        IGameServer_GetMapElementType = CLOSED_CELL
    Else
        IGameServer_GetMapElementType = VACANT_CELL
    End If
End Function
