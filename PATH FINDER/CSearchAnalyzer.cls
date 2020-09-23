VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSearchAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private GameServer As IGameServer
Private SearchQueue As CQueue
Private SearchGrid(NUM_ROWS, NUM_COLS) As Long

Private Sub Class_Initialize()
    Set SearchQueue = New CQueue
    Set GameServer = frmMain
    Call InitSearchArray
End Sub

Private Sub InitSearchArray()
    Dim r As Long
    Dim c As Long
    
    For r = 1 To NUM_ROWS
        For c = 1 To NUM_COLS
           SearchGrid(r, c) = GameServer.GetMapElementType(r, c)
        Next c
    Next r
        
    frmMain.Text2.Text = ""
    For r = 1 To NUM_ROWS
        For c = 1 To NUM_COLS
            frmMain.Text2.Text = frmMain.Text2.Text & SearchGrid(r, c)
        Next c
        frmMain.Text2.Text = frmMain.Text2.Text & vbCrLf
    Next r
End Sub

Public Sub PerformTargetSurveillance(ByVal Target As ITargetable)
    Dim mLocation As CLocation

    Call ResetSearchArray(Target)
    Set mLocation = New CLocation
    mLocation.Row = Target.GetRow()
    mLocation.Col = Target.GetCol()
    Call SearchQueue.Enqueue(mLocation)
    Call PopulateSearchArray
End Sub
    
Private Sub ResetSearchArray(ByVal Target As ITargetable)
    Dim r As Long
    Dim c As Long

    Call SearchQueue.Clear
    For r = 1 To NUM_ROWS
        For c = 1 To NUM_COLS
            If (SearchGrid(r, c) <> CLOSED_CELL) Then
                SearchGrid(r, c) = VACANT_CELL
            End If
        Next c
    Next r
    SearchGrid(Target.GetRow(), Target.GetCol()) = TARGET_CELL
End Sub

Private Sub PopulateSearchArray()
    Dim mLocation  As CLocation

    Set mLocation = SearchQueue.Dequeue()
    If Not (mLocation Is Nothing) Then
        Call PopulateCells(mLocation)
        '==============================
        ' Call PrintIt
        '==============================
        Call PopulateSearchArray
    End If
End Sub

Private Sub PopulateCells(ByVal Cell As CLocation)
    Call PlotCell(Cell, -1, 0, SOUTH_CELL)
    Call PlotCell(Cell, 1, 0, NORTH_CELL)
    Call PlotCell(Cell, 0, 1, WEST_CELL)
    Call PlotCell(Cell, 0, -1, EAST_CELL)
End Sub

Private Sub PlotCell( _
  ByVal Current As CLocation, _
  ByVal RowIncrement As Integer, _
  ByVal ColIncrement As Integer, _
  ByVal Direction As Integer)
  
    Dim r As Long
    Dim c As Long
    Dim m As CLocation
  
    r = Current.Row + RowIncrement
    c = Current.Col + ColIncrement
    If (SearchGrid(r, c) = VACANT_CELL) Then
        SearchGrid(r, c) = Direction
        Set m = New CLocation
        m.Row = r
        m.Col = c
        Call SearchQueue.Enqueue(m)
    End If
End Sub

'==============================================================================
'      METHOD : SuggestDirection
' DESCRIPTION :
'==============================================================================
Public Function SuggestDirection(ByVal Source As ITargetable) As Long
    Dim dir As Long
    Dim info As Integer
    
    dir = 0
    info = SearchGrid(Source.GetRow(), Source.GetCol())
    If (info = SOUTH_CELL) Then
        dir = vbKeyDown
    ElseIf (info = NORTH_CELL) Then
        dir = vbKeyUp
    ElseIf (info = WEST_CELL) Then
        dir = vbKeyLeft
    Else
        If (info = EAST_CELL) Then
            dir = vbKeyRight
        End If
    End If
    SuggestDirection = dir
End Function
'==============================================================================
'      METHOD : PrintIt
' DESCRIPTION : Allows you to see how the search path algorithm works
'==============================================================================
Private Sub PrintIt()
    Dim r, c As Long
    Dim mStr As String

    mStr = ""
    For r = 1 To NUM_ROWS
        For c = 1 To NUM_COLS
            If (SearchGrid(r, c) = VACANT_CELL) Then
                mStr = mStr & " " & vbTab
            ElseIf (SearchGrid(r, c) = CLOSED_CELL) Then
                mStr = mStr & "#" & vbTab
            ElseIf (SearchGrid(r, c) = TARGET_CELL) Then
                mStr = mStr & "T" & vbTab
            ElseIf (SearchGrid(r, c) = NORTH_CELL) Then
                mStr = mStr & "^" & vbTab
            ElseIf (SearchGrid(r, c) = SOUTH_CELL) Then
                mStr = mStr & "V" & vbTab
            ElseIf (SearchGrid(r, c) = EAST_CELL) Then
                mStr = mStr & ">" & vbTab
            Else
                If (SearchGrid(r, c) = WEST_CELL) Then
                    mStr = mStr & "<" & vbTab
                End If
            End If
        Next c
        mStr = mStr & vbCrLf & vbCrLf
    Next r
    MsgBox mStr
End Sub

Public Function GetOpenRouteCount(ByRef r As Long, ByRef c As Long) As Integer
    Dim count As Integer
    
    count = 0
    If (IsOpenRoute(r + 1, c)) Then count = count + 1
    If (IsOpenRoute(r - 1, c)) Then count = count + 1
    If (IsOpenRoute(r, c + 1)) Then count = count + 1
    If (IsOpenRoute(r, c - 1)) Then count = count + 1
    GetOpenRouteCount = count
End Function

Public Function IsOpenRoute(ByVal r As Long, ByVal c As Long) As Boolean
    IsOpenRoute = (SearchGrid(r, c) <> CLOSED_CELL)
End Function
