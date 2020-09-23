VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Magix Chat Server"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3555
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H009ED3D3&
      Caption         =   "Reset Name "
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H009ED3D3&
      Caption         =   "Exit Magix Chat"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H009ED3D3&
      Caption         =   "More info on magix Chat HOST"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009ED3D3&
      Caption         =   "Run Chat Server"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Magix Chat - Private Chat"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   2760
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   240
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4680
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show
Me.Hide
End Sub

Private Sub Command2_Click()
MsgBox "Made By Steven Dorman"
MsgBox "Magix Is part of HgP's Internet / network based programs"
MsgBox "You can chat with multiple people by opening the program again and using a diffrent port"
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
On Error Resume Next
Open "C:\magix.nfo" For Output As #1
Print #1, ""
Close #1
MsgBox "Name reset to nothing"
End Sub
