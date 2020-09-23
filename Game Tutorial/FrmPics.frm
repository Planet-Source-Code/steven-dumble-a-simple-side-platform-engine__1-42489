VERSION 5.00
Begin VB.Form FrmPics 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   226
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox MapPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   0
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   3
      Top             =   0
      Width           =   6015
   End
   Begin VB.PictureBox MapMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1335
      Left            =   0
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   2
      Top             =   1440
      Width           =   6015
   End
   Begin VB.PictureBox Character 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   6480
      Picture         =   "FrmPics.frx":0000
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   1
      Top             =   600
      Width           =   660
   End
   Begin VB.PictureBox CharacterMask 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   690
      Left            =   6480
      Picture         =   "FrmPics.frx":13F2
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   0
      Top             =   1560
      Width           =   660
   End
End
Attribute VB_Name = "FrmPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
