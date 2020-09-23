VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Magix Chat (HOST)"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6090
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      BackColor       =   &H009ED3D3&
      Caption         =   "Emote"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H009ED3D3&
      Caption         =   "Last message"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H009ED3D3&
      Caption         =   "Save Name"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox nick 
      BackColor       =   &H009ED3D3&
      Height          =   285
      Left            =   7560
      TabIndex        =   12
      Text            =   "ChaTTEr"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H009ED3D3&
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H009ED3D3&
      Caption         =   "Note"
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H009ED3D3&
      Caption         =   "Clear Chat History"
      Height          =   495
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock sock 
      Left            =   4440
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009ED3D3&
      Caption         =   "Send"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox outtext 
      BackColor       =   &H009ED3D3&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   5160
      Width           =   5295
   End
   Begin VB.ListBox chat 
      BackColor       =   &H009ED3D3&
      Height          =   4155
      ItemData        =   "Form2.frx":59F3
      Left            =   360
      List            =   "Form2.frx":59F5
      TabIndex        =   0
      Top             =   840
      Width           =   6975
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   9360
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line6 
      X1              =   9360
      X2              =   9360
      Y1              =   4560
      Y2              =   5880
   End
   Begin VB.Line Line5 
      X1              =   7440
      X2              =   9480
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Your name/ Alias"
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Line Line4 
      X1              =   7440
      X2              =   9360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      X1              =   7440
      X2              =   7440
      Y1              =   4560
      Y2              =   0
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   9600
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label yip 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   5640
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Your IP is:"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   5640
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   3240
      X2              =   3240
      Y1              =   5640
      Y2              =   6000
   End
   Begin VB.Label hip 
      BackStyle       =   0  'Transparent
      Caption         =   "0.0.0.0"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No One Is Connected"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   1695
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
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   6135
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
Label2.Caption = "No one is connected"
hip.Caption = "0.0.0.0"
Dim a, b As String
a = "User has disconected from chat"
sock.SendData a

DoEvents
sock.Close
b = InputBox("Enter the port to use")
sock.LocalPort = b
sock.Listen

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
Dim b As String
b = InputBox("Enter The Port To Use")
sock.LocalPort = b
sock.Listen
yip.Caption = sock.LocalIP
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

Private Sub sock_ConnectionRequest(ByVal requestID As Long)
If sock.State <> sckClosed Then sock.Close
sock.Accept requestID
Label2.Caption = "CONNECTED"
chat.AddItem "connected"
hip.Caption = sock.RemoteHostIP
Variable = sndPlaySound(App.Path & "\alert.wav", 1)

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

