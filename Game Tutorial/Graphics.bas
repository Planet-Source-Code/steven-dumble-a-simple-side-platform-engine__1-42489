Attribute VB_Name = "Graphics"
Option Base 1
Option Explicit

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Constants for Raster Operations used by BitBlt function.
Const SRCAND = &H8800C6      ' dest = source AND dest
Const SRCCOPY = &HCC0020     ' dest = source
Const SRCPAINT = &HEE0086    ' dest = source OR dest

Function DrawBackGround(CurrentTopX As Long, CurrentTopY, Height As Integer, Width As Integer, SourceHDC As Long, DestHDC As Long)
    BitBlt DestHDC, 0, 0, ViewPort.ScreenWidth, ViewPort.ScreenHeight, SourceHDC, CurrentTopX, CurrentTopY, SRCCOPY
End Function

Function SwapBuffers()
    BitBlt FrmMain.PlayArea.hdc, 0, 0, ViewPort.ScreenWidth, ViewPort.ScreenHeight, FrmMain.BackBuffer.hdc, 0, 0, SRCCOPY
End Function

Function DrawPlayer(Xpos As Long, Ypos As Long, Width As Integer, Height As Integer, PicSourceHDC As Long, MaskSourceHDC As Long)
    ' Draw the sprite mask directly onto the backbuffer
    BitBlt FrmMain.BackBuffer.hdc, Xpos, Ypos, Width, Height, MaskSourceHDC, 0, 0, SRCAND
    
    ' Draw the sprite over top of the mask.
    BitBlt FrmMain.BackBuffer.hdc, Xpos, Ypos, Width, Height, PicSourceHDC, 0, 0, SRCPAINT
End Function

Function DrawAllPlayers()
Dim X As Integer
Dim LocalX As Long
Dim LocalY As Long

    For X = 1 To NumberOfPlayers
        LocalX = Players(X).X - ViewPort.CurrentTopX
        LocalY = Players(X).Y - ViewPort.CurrentTopY
        
        'check if on screen
            If IsAreaVisible(Players(X).X, Players(X).Y, Players(X).Width, Players(X).Height) Then
                If Players(X).Visible = True Then
                    DrawPlayer LocalX, LocalY, Players(X).Width, Players(X).Height, FrmPics.Character.hdc, FrmPics.CharacterMask.hdc
                End If
            End If
  
    Next X
End Function

Function AdjustBackGround()
Dim CenteredX, CenteredY As Long
    ' Calculate the center of the Viewer.
    CenteredX = (FrmMain.PlayArea.ScaleWidth - Players(1).Width) \ 2
    CenteredY = (FrmMain.PlayArea.ScaleHeight - Players(1).Height) \ 2
        
    
    If (Players(1).X > CenteredX) Then
            If (Players(1).X < (FrmPics.MapPicture.ScaleWidth - Players(1).Width - CenteredX)) Then
                ViewPort.CurrentTopX = Players(1).X - CenteredX
            Else
                ViewPort.CurrentTopX = FrmPics.MapPicture.ScaleWidth - ViewPort.ScreenWidth
            End If
    End If
        
        
    If (Players(1).Y > CenteredY) Then
            If (Players(1).Y < (FrmPics.MapPicture.ScaleHeight - Players(1).Height - CenteredY)) Then
                ViewPort.CurrentTopY = Players(1).Y - CenteredY
            Else
                ViewPort.CurrentTopY = FrmPics.MapPicture.ScaleHeight - ViewPort.ScreenHeight
            End If
    End If
End Function

