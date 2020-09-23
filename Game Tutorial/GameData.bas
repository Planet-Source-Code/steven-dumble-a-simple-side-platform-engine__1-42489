Attribute VB_Name = "GameData"
Option Base 1
Option Explicit

Type ViewData
    CurrentTopX As Long
    CurrentTopY As Long
    ScreenWidth As Integer
    ScreenHeight As Integer
End Type

Global ViewPort As ViewData

Type PlayerData
    X As Long
    Y As Long
    Xspeed As Single
    Yspeed As Single
    Visible As Boolean
    Width As Integer
    Height As Integer
End Type

Global Players() As PlayerData
Global NumberOfPlayers As Integer

Public Const MAX_SPEED As Single = 10
Public Const GRAVITY As Single = 1
Public Const FRICTION As Single = 0.4
Public Const ACCELERATION As Single = 4

'Input Data
Global Adown As Boolean
Global Ddown As Boolean
Global Wdown As Boolean
Global Sdown As Boolean

Sub Main()
        
    Load FrmPics
    Load FrmMain
   
    LoadMap App.Path & "\MAPS\", "Map1"
    
    NewPlayer
    
      
    FrmMain.Show
    DoEvents
    
    FrmMain.GameTimer = True

End Sub


Sub UnloadGame()
    Unload FrmMain
    Unload FrmPics
End Sub

Sub ApplyGravity()
Dim X As Integer
    For X = 1 To NumberOfPlayers
        'add gravity onto the Yspeed
        If Players(X).Visible = True Then
                Players(X).Yspeed = Players(X).Yspeed + GRAVITY
        End If
    Next X
End Sub

Sub AddFriction()
    Dim X As Integer
    For X = 1 To NumberOfPlayers
        If Players(X).Visible = True Then
            If IsOnGround(Players(X).X + (Players(X).Width / 2), Players(X).Y + Players(X).Height + 1) Then
                If Players(X).Xspeed > 0 Then
                    Players(X).Xspeed = Players(X).Xspeed - FRICTION
                    If Players(X).Xspeed < 0 Then Players(X).Xspeed = 0
                End If
                
                If Players(X).Xspeed < 0 Then
                    Players(X).Xspeed = Players(X).Xspeed + FRICTION
                    If Players(X).Xspeed > 0 Then Players(X).Xspeed = 0
                End If
            End If
        End If
    Next X
End Sub

Function GetTopOfTerrain(X As Long, Y As Long) As Long
Dim i As Integer
    For i = 0 To 100
        If CheckPoint(X, Y - i) = False Then
            GetTopOfTerrain = i
            Exit Function
        End If
    Next i
End Function

Sub MoveAllPlayers()
    Dim X As Integer
    Dim Temp As Integer
    For X = 1 To NumberOfPlayers
        If Players(X).Visible = True Then
            Players(X).X = Players(X).X + Players(X).Xspeed
            Players(X).Y = Players(X).Y + Players(X).Yspeed
            If IsOnGround(Players(X).X + (Players(X).Width / 2), Players(X).Y + Players(X).Height) Then
                Players(X).Y = Players(X).Y - GetTopOfTerrain(Players(X).X + (Players(X).Width / 2), Players(X).Y + Players(X).Height)
            End If
        End If
    Next X
End Sub

Sub DoInput()

'Key Boolean On/Off Checks

    If Adown = True Then
        If Players(1).Xspeed < (MAX_SPEED * -1) Then
            Players(1).Xspeed = MAX_SPEED * -1
        Else
            If Players(1).Xspeed - ACCELERATION < (MAX_SPEED * -1) Then
                    Players(1).Xspeed = MAX_SPEED * -1
            Else
                    Players(1).Xspeed = Players(1).Xspeed - ACCELERATION
            End If
          End If
    End If
    
If Ddown = True Then
    If Players(1).Xspeed > MAX_SPEED Then
        Players(1).Xspeed = MAX_SPEED
    Else
        If Players(1).Xspeed + ACCELERATION > MAX_SPEED Then
            Players(1).Xspeed = MAX_SPEED
        Else
            Players(1).Xspeed = Players(1).Xspeed + ACCELERATION
        End If
    End If
End If
    
If Wdown = True Then
    If IsOnGround(Players(1).X + (Players(1).Width / 2), Players(1).Y + Players(1).Height + 1) Then
        Players(1).Yspeed = Players(1).Yspeed - (ACCELERATION * 4)
        If Players(1).Yspeed < (MAX_SPEED * -1) Then Players(1).Yspeed = MAX_SPEED * -1
    End If
End If
          
End Sub

Function IsAreaVisible(TopX As Long, TopY As Long, Width As Integer, Height As Integer) As Boolean
Dim X1, X2, X3, X4 As Long
Dim Y1, Y2, Y3, Y4 As Long
    
    X1 = TopX 'top left
    Y1 = TopY

    X2 = X1 + Width 'top right
    Y2 = Y1

    X3 = X1
    Y3 = Y1 + Height 'bottom left
   
    X4 = X1 + Width 'bottom right
    Y4 = Y1 + Height
    
    If X1 > ViewPort.CurrentTopX And X1 < ViewPort.CurrentTopX + ViewPort.ScreenWidth Then
        If Y1 > ViewPort.CurrentTopY And Y1 < ViewPort.CurrentTopY + ViewPort.ScreenWidth Then
            IsAreaVisible = True
            Exit Function
        End If
    End If
    
    If X2 > ViewPort.CurrentTopX And X1 < ViewPort.CurrentTopX + ViewPort.ScreenWidth Then
        If Y2 > ViewPort.CurrentTopY And Y1 < ViewPort.CurrentTopY + ViewPort.ScreenWidth Then
            IsAreaVisible = True
            Exit Function
        End If
    End If
    
    If X3 > ViewPort.CurrentTopX And X1 < ViewPort.CurrentTopX + ViewPort.ScreenWidth Then
        If Y3 > ViewPort.CurrentTopY And Y1 < ViewPort.CurrentTopY + ViewPort.ScreenWidth Then
            IsAreaVisible = True
            Exit Function
        End If
    End If

    If X4 > ViewPort.CurrentTopX And X1 < ViewPort.CurrentTopX + ViewPort.ScreenWidth Then
        If Y4 > ViewPort.CurrentTopY And Y1 < ViewPort.CurrentTopY + ViewPort.ScreenWidth Then
            IsAreaVisible = True
            Exit Function
        End If
    End If
    
End Function

Function FileExists(FullFileName As String) As Boolean
    On Error GoTo ErrSub:
        Open FullFileName For Input As #1
        Close #1
        FileExists = True
    Exit Function
ErrSub:
        FileExists = False
    Exit Function
End Function

Function NewPlayer() As Integer
    NumberOfPlayers = NumberOfPlayers + 1
    NewPlayer = NumberOfPlayers
    ReDim Preserve Players(NewPlayer)
        With Players(NewPlayer)
            .Visible = True
            .X = 10
            .Y = 10
            .Xspeed = 0
            .Yspeed = 0
            .Width = FrmPics.Character.ScaleWidth
            .Height = FrmPics.Character.ScaleHeight
        End With
End Function

Sub LoadMap(Path As String, MapName As String)
    FrmPics.MapPicture = LoadPicture(Path & MapName & ".MAP")
    
    If FileExists(Path & MapName & ".DAT") = False Then  'mask data doesn't exists... must create it
        FrmPics.MapMask = LoadPicture(Path & MapName & ".MSK")
        MsgBox "This is the first time this map has been used. The program will need to create a collision map for it. This can take a now long time. Please Wait.", vbInformation + vbOKOnly, "Game Tutorial."
        CreateCollisionMap
        SaveCollisionMap Path & MapName & ".DAT"
    Else
        LoadCollisionMap Path & MapName & ".DAT"
    End If

End Sub
 

