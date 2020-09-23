Attribute VB_Name = "Collisions"
Option Explicit
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Dim CollisionMap() As Boolean
Dim CMHeight As Integer
Dim CMWidth As Integer

Sub CreateCollisionMap() 'Convert a BMP Mask into an 2d Array of _ booleans 0 = white, 1 = black
Dim X, Y As Integer
Dim Tempdata As Long
    
    CMHeight = FrmPics.MapMask.Height
    CMWidth = FrmPics.MapMask.Width
    ReDim CollisionMap(CMWidth - 1, CMHeight - 1)

    For Y = 0 To CMHeight - 1
        For X = 0 To CMWidth - 1
            DoEvents
            Tempdata = GetPixel(FrmPics.MapMask.hdc, X, Y)
            If Tempdata = 0 Then CollisionMap(X, Y) = True
        Next X
    Next Y

End Sub

Sub SaveCollisionMap(Path As String)
On Error GoTo ErrSub:
    Open Path For Binary Access Write As #1
        Put #1, , CMWidth
        Put #1, , CMHeight
        Put #1, , CollisionMap
    Close #1
ErrSub:
End Sub

Sub LoadCollisionMap(Path As String)
        Open Path For Binary Access Read As #1
        Get #1, , CMWidth
        Get #1, , CMHeight
        ReDim CollisionMap(CMWidth - 1, CMHeight - 1)
        Get #1, , CollisionMap
    Close #1
End Sub

Function CheckPoint(X As Long, Y As Long) As Boolean
    If X < 0 Then
        CheckPoint = True
        Exit Function
    End If
    
    If Y < 0 Then
        CheckPoint = True
        Exit Function
    End If
    
    If X >= CMWidth - 1 Then
        CheckPoint = True
        Exit Function
    End If
    
    If Y >= CMHeight - 1 Then
        CheckPoint = True
        Exit Function
    End If
    
    If CollisionMap(X, Y) = True Then CheckPoint = True
End Function

Sub AllPlayersBGCollision()
Dim X As Integer
Dim X1, X2, X3, X4 As Long
Dim Y1, Y2, Y3, Y4 As Long
Dim TempSpeed As Single

For X = 1 To NumberOfPlayers
DoEvents

    If Players(X).Visible = True Then
        X1 = Players(X).X + (Players(X).Width / 2) 'top middle
        Y1 = Players(X).Y
       
        X2 = Players(X).X + Players(X).Width ' right middle
        Y2 = Players(X).Y + (Players(X).Height / 2)
       
        X3 = Players(X).X + (Players(X).Width / 2) 'bottom middle
        Y3 = Players(X).Y + Players(X).Height
       
        X4 = Players(X).X 'left middle
        Y4 = Players(X).Y + (Players(X).Height / 2)
    
        If Players(X).Yspeed > 0 Then 'going down
            TempSpeed = BGCollisionBottom(X3, Y3, Players(X).Yspeed)
            If TempSpeed < Players(X).Yspeed Then 'hit
                Players(X).Yspeed = Players(X).Yspeed * -1 * 0.2 'bounce
            Else
                Players(X).Yspeed = TempSpeed
            End If
        ElseIf Players(X).Yspeed < 0 Then 'going up
            TempSpeed = BGCollisionTop(X1, Y1, Players(X).Yspeed)
            Players(X).Yspeed = TempSpeed
        End If
    
    
        If Players(X).Xspeed < 0 Then 'going left
            TempSpeed = BGCollisionLeft(X4, Y4, Players(X).Xspeed)
            Players(X).Xspeed = TempSpeed
        Else 'going right
            TempSpeed = BGCollisionRight(X2, Y2, Players(X).Xspeed)
            Players(X).Xspeed = TempSpeed
        End If
    End If
Next X
End Sub


Function BGCollisionTop(ByVal X1 As Long, ByVal Y1 As Long, ByVal Yspeed As Single) As Single
Dim Y As Integer
Dim Tempdata As Boolean
    For Y = 0 To Abs(Int(Yspeed))
        Tempdata = CheckPoint(X1, Y1 - Y)
        If Tempdata = True Then
            If Y = 1 Then Y = 0
            BGCollisionTop = Y * -1
            Exit Function
        End If
    Next Y
    BGCollisionTop = Yspeed
End Function

Function BGCollisionBottom(ByVal X1 As Long, ByVal Y1 As Long, ByVal Yspeed As Single) As Single
Dim Y As Integer
Dim Tempdata As Boolean
    For Y = 0 To Int(Yspeed)
        Tempdata = CheckPoint(X1, Y1 + Y)
        If Tempdata = True Then
            If Y = 1 Then Y = 0
            BGCollisionBottom = Y
            Exit Function
        End If
    Next Y
    BGCollisionBottom = Yspeed
End Function

Function BGCollisionLeft(ByVal X1 As Long, ByVal Y1 As Long, ByVal Xspeed As Single) As Single
Dim Y As Integer
Dim Tempdata As Boolean
    For Y = 0 To Abs(Int(Xspeed))
        Tempdata = CheckPoint(X1 - Y, Y1)
        If Tempdata = True Then
            If Y = 1 Then Y = 0
            BGCollisionLeft = (Y) * -1
            Exit Function
        End If
    Next Y
    BGCollisionLeft = Xspeed
End Function


Function BGCollisionRight(ByVal X1 As Long, ByVal Y1 As Long, ByVal Xspeed As Single) As Single
Dim Y As Integer
Dim Tempdata As Boolean
    For Y = 0 To Int(Xspeed)
        Tempdata = CheckPoint(X1 + Y, Y1)
        If Tempdata = True Then
            If Y = 1 Then Y = 0
            BGCollisionRight = Y
            Exit Function
        End If
    Next Y
    BGCollisionRight = Xspeed
End Function


Function IsOnGround(X1 As Long, Y1 As Long) As Boolean
    If CheckPoint(X1, Y1) = True Then IsOnGround = True
End Function


