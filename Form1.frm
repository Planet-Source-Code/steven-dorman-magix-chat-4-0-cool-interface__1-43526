VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Magix Chat"
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
      Caption         =   "Reset Name"
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
      TabIndex        =   4
      Top             =   2640
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
      TabIndex        =   3
      Top             =   2160
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H009ED3D3&
      Caption         =   "More About Magix Chat"
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
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009ED3D3&
      Caption         =   "Connect To Chat Host"
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
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   240
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4440
      Y1              =   1920
      Y2              =   1920
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
      TabIndex        =   1
      Top             =   240
      Width           =   2760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form2.Show
Dim a, b As String
a = InputBox("Enter the IP to connect to")
b = InputBox("Enter the Port to use")
Form2.sock.Connect a, b
Me.Hide
End Sub

Private Sub Command2_Click()
MsgBox "Made by steven dorman"
MsgBox "Magix chat is a personal chat server to chat with someone privately"
MsgBox "You may open this program more than once and chat with more than one person"
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

