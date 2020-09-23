VERSION 5.00
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Game Tutorial By Steven Dumble"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer GameTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   120
      Top             =   6960
   End
   Begin VB.PictureBox BackBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6615
      Left            =   0
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1013
      TabIndex        =   1
      Top             =   8760
      Width           =   15255
   End
   Begin VB.PictureBox PlayArea 
      Height          =   6615
      Left            =   0
      ScaleHeight     =   437
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1013
      TabIndex        =   0
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'set up the view port
    ViewPort.CurrentTopX = 0
    ViewPort.CurrentTopY = 0
    ViewPort.ScreenHeight = PlayArea.ScaleHeight
    ViewPort.ScreenWidth = PlayArea.ScaleWidth
    
End Sub

Private Sub GameTimer_Timer()
    DoEvents
    DoGameLoop
    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadGame
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Players(1).Visible = True Then
        If KeyCode = 65 Then Adown = True
        If KeyCode = 68 Then Ddown = True
        If KeyCode = 87 Then Wdown = True
        If KeyCode = 83 Then Sdown = True
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Players(1).Visible = True Then
        If KeyCode = 65 Then Adown = False
        If KeyCode = 68 Then Ddown = False
        If KeyCode = 87 Then Wdown = False
        If KeyCode = 83 Then Sdown = False
    End If
End Sub

