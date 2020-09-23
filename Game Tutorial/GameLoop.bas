Attribute VB_Name = "GameLoop"

Sub DoGameLoop()
        
    ApplyGravity
    
    DoInput
      
    AddFriction
    
    AllPlayersBGCollision

    MoveAllPlayers
    
    AdjustBackGround

    DrawBackGround ViewPort.CurrentTopX, ViewPort.CurrentTopY, ViewPort.ScreenHeight, ViewPort.ScreenWidth, FrmPics.MapPicture.hdc, FrmMain.BackBuffer.hdc
    
    DrawAllPlayers
    
    SwapBuffers ' draw the main screen

    DoEvents
    
End Sub

