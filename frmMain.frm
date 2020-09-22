VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "LCDRAM"
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timUpdate 
      Interval        =   2500
      Left            =   3360
      Top             =   1920
   End
   Begin VB.PictureBox picEmpty 
      Height          =   615
      Left            =   2640
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   9
      Left            =   2040
      Picture         =   "frmMain.frx":0500
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   8
      Left            =   1560
      Picture         =   "frmMain.frx":0A87
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   7
      Left            =   1080
      Picture         =   "frmMain.frx":0F8E
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   6
      Left            =   600
      Picture         =   "frmMain.frx":14B2
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   5
      Left            =   120
      Picture         =   "frmMain.frx":1A2A
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   4
      Left            =   2040
      Picture         =   "frmMain.frx":1F77
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   3
      Left            =   1560
      Picture         =   "frmMain.frx":249C
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   3
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   2
      Left            =   1080
      Picture         =   "frmMain.frx":29F5
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   1
      Left            =   600
      Picture         =   "frmMain.frx":2F82
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   1
      Top             =   1920
      Width           =   495
   End
   Begin VB.PictureBox picLCD 
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":3484
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LCDRAM"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15
      TabIndex        =   13
      Top             =   0
      Width           =   4675
   End
   Begin VB.Image imgF4 
      Height          =   540
      Left            =   2520
      Picture         =   "frmMain.frx":39D3
      Top             =   960
      Width           =   420
   End
   Begin VB.Image imgF3 
      Height          =   540
      Left            =   2880
      Picture         =   "frmMain.frx":3F22
      Top             =   960
      Width           =   420
   End
   Begin VB.Image imgF2 
      Height          =   540
      Left            =   3240
      Picture         =   "frmMain.frx":4471
      Top             =   960
      Width           =   420
   End
   Begin VB.Image imgF1 
      Height          =   540
      Left            =   3600
      Picture         =   "frmMain.frx":49C0
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Top             =   360
      Width           =   135
   End
   Begin VB.Image imgT4 
      Height          =   540
      Left            =   2520
      Picture         =   "frmMain.frx":4F0F
      Top             =   360
      Width           =   420
   End
   Begin VB.Image imgT3 
      Height          =   540
      Left            =   2880
      Picture         =   "frmMain.frx":545E
      Top             =   360
      Width           =   420
   End
   Begin VB.Image imgT2 
      Height          =   540
      Left            =   3240
      Picture         =   "frmMain.frx":59AD
      Top             =   360
      Width           =   420
   End
   Begin VB.Image imgT1 
      Height          =   540
      Left            =   3600
      Picture         =   "frmMain.frx":5EFC
      Top             =   360
      Width           =   420
   End
   Begin VB.Image imgMB2 
      Height          =   345
      Left            =   4080
      Picture         =   "frmMain.frx":644B
      Top             =   1080
      Width           =   465
   End
   Begin VB.Image imgMB1 
      Height          =   345
      Left            =   4080
      Picture         =   "frmMain.frx":68C4
      Top             =   480
      Width           =   465
   End
   Begin VB.Image imgFree 
      Height          =   345
      Left            =   540
      Picture         =   "frmMain.frx":6D3D
      Top             =   1080
      Width           =   1785
   End
   Begin VB.Image imgTotal 
      Height          =   345
      Left            =   120
      Picture         =   "frmMain.frx":7597
      Top             =   480
      Width           =   2205
   End
   Begin VB.Label lblBack 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'LCDRAM by CArNi4
'
'This program shows you how to obtain the amount of free and total RAM
'and how to show this in a picture based lay-out
'The font I used to make the picture's is LCDFont,
'I forgot where I downloaded it :'(. But you can search at:
'       www.flamingtext.com
'       www.fontflood.com
'
'You are free to use everything in this program in your own programs,
'as long as you VOTE for me and thank me in your program

Private Sub Form_Load()
Me.Height = lblBack.Height + 5

lblBack.BackColor = RGB(210, 210, 210)
lblTitle.BackColor = RGB(190, 190, 190)

timUpdate_Timer
End Sub

Private Sub lblClose_Click()
Unload Me
End Sub

Private Sub timUpdate_Timer()
GetRAMInfo
ShowRAMInfo
End Sub
