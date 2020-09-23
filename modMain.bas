Attribute VB_Name = "modMain"
Private Type FXTile
    Walkable As Integer
    FXType As Integer
    Layer As Integer
    Event As String
End Type

Public Type Map
    Tiles(287) As FXTile
    sname As String
End Type

Public ThisMap As Map

Function FindPart(lzStr As String, mPart As String) As Integer
Dim TPos As Integer
    TPos = InStr(lzStr, mPart)
    If TPos Then
        FindPart = 1
    Else
        FindPart = 0
    End If
    
End Function
