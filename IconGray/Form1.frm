VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " MFG  <Gray Icon Converter>"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Image3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   3
      Top             =   2160
      Width           =   720
   End
   Begin VB.PictureBox Image2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   360
      Picture         =   "Form1.frx":0ECA
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   720
   End
   Begin VB.PictureBox Image1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   360
      Picture         =   "Form1.frx":1D94
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   240
      Width           =   720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Turn Gray"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Convert Icons into a Better Gray Scale Color Gradient Look. "
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   3000
      Width           =   4260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

RenderIconGrayscale Image1.hdc, Image1.Picture.Handle
RenderIconGrayscale Image2.hdc, Image2.Picture.Handle
RenderIconGrayscale Image3.hdc, Image3.Picture.Handle

End Sub

