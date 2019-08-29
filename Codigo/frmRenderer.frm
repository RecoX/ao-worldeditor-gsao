VERSION 5.00
Begin VB.Form frmRenderer 
   Caption         =   "Renderizando....."
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   7080
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   360
      ScaleHeight     =   5715
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Image Smallpic 
      Height          =   5535
      Left            =   480
      Stretch         =   -1  'True
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

