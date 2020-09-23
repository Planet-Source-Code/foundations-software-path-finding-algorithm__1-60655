VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private List As Collection

Private Sub Class_Initialize()
    Set List = New Collection
End Sub

Public Sub Enqueue(ByRef Location As CLocation)
    Call List.Add(Location)
End Sub

Public Function Dequeue() As CLocation
    Dim Location As CLocation
    
    Set Location = Nothing
    If (List.Count > 0) Then
        Set Location = List.Item(1)
        Call List.Remove(1)
    End If
    Set Dequeue = Location
End Function

Public Sub Clear()
    While (List.Count > 0)
        Call List.Remove(1)
    Wend
End Sub
