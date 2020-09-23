VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Magix Chat"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6060
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      BackColor       =   &H009ED3D3&
      Caption         =   "Emote"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H009ED3D3&
      Caption         =   "last message"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H009ED3D3&
      Caption         =   "Save Name"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox nick 
      BackColor       =   &H009ED3D3&
      Height          =   285
      Left            =   7680
      TabIndex        =   8
      Text            =   "ChaTTer"
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H009ED3D3&
      Caption         =   "Connect to diffent host"
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H009ED3D3&
      Caption         =   "Note"
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H009ED3D3&
      Caption         =   "Clear Chat History"
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   2400
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009ED3D3&
      Caption         =   "Send"
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox outtext 
      BackColor       =   &H009ED3D3&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   5400
      Width           =   6255
   End
   Begin VB.ListBox chat 
      BackColor       =   &H009ED3D3&
      Height          =   4740
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   7095
   End
   Begin VB.Line Line3 
      X1              =   7560
      X2              =   9000
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line2 
      X1              =   7560
      X2              =   8880
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick name/Alias"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   7560
      X2              =   7560
      Y1              =   0
      Y2              =   5280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Chat  Server Magix by HGP"
      BeginProperty Font 
         Name            =   "Lucida Blackletter"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public old As String
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Sub Command1_Click()
On Error Resume Next
Dim a As String
old = outtext.Text
a = "<" & nick.Text & "> " & outtext.Text
sock.SendData a
DoEvents
chat.AddItem "<" & nick.Text & "> " & outtext.Text
outtext.Text = ""
If chat.ListCount = 25 Then
chat.Clear
End If
End Sub

Private Sub Command2_Click()
chat.Clear
End Sub

Private Sub Command3_Click()
Dim a As String
a = InputBox("Enter text for the Note")
chat.AddItem a

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim a As String
a = "User has disconected from chat"
sock.SendData a

DoEvents
sock.Close



Dim c, b As String
c = InputBox("Enter the IP to connect to")
b = InputBox("Enter the Port to use")
Form2.sock.Connect c, b

End Sub

Private Sub Command5_Click()
Open "C:\magix.nfo" For Output As #1
Print #1, nick.Text
Close #1
MsgBox "Saved!"
End Sub

Private Sub Command6_Click()
outtext.Text = old
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim a As String
old = outtext.Text
a = "<" & outtext.Text & "> "
sock.SendData a
DoEvents
chat.AddItem "<" & outtext.Text & "> "
outtext.Text = ""
If chat.ListCount = 25 Then
chat.Clear
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim a As String
Open "C:\magix.nfo" For Input As #1
Input #1, a
Close #1
nick.Text = a

If App.PrevInstance Then
        
     
        MsgBox "A previous Instance of Magix Chat has been Detected.Please use a diffrent port with the other instance"
        End If

End Sub

Private Sub Form_Terminate()
On Error Resume Next
Dim a As String
a = "User has disconected from chat"
sock.SendData a

DoEvents
sock.Close
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim a As String
a = "User has disconected from chat"
sock.SendData a

DoEvents
sock.Close
Variable = sndPlaySound(App.Path & "\logoff.wav", 1)
End
End Sub

Private Sub outtext_KeyPress(KeyAscii As Integer)
If GetAsyncKeyState(vbKeyReturn) Then
Command1_Click
End If
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
Dim dat As String
sock.GetData dat
If dat = "User has disconected from chat" Then
chat.AddItem dat
Variable = sndPlaySound(App.Path & "\logoff.wav", 1)
Else
chat.AddItem dat
Variable = sndPlaySound(App.Path & "\type.wav", 1)
End If
End Sub

