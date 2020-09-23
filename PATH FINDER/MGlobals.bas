Attribute VB_Name = "MGlobals"
Option Explicit

Public Const NUM_ROWS = 12
Public Const NUM_COLS = 12

Public Const TILE_WIDTH = 32
Public Const TILE_HEIGHT = 32

Public Const VACANT_CELL = 0
Public Const CLOSED_CELL = 1
Public Const TARGET_CELL = 2
Public Const SOUTH_CELL = 3
Public Const NORTH_CELL = 4
Public Const WEST_CELL = 5
Public Const EAST_CELL = 6

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public SearchAnalyzer As CSearchAnalyzer
