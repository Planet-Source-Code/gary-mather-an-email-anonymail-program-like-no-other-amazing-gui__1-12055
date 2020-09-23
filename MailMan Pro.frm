VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Fake mailer"
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   Icon            =   "MailMan Pro.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "MailMan Pro.frx":0442
   ScaleHeight     =   5205
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3525
      Picture         =   "MailMan Pro.frx":466D8
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   75
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   3795
      Picture         =   "MailMan Pro.frx":46A8E
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   75
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   90
      Picture         =   "MailMan Pro.frx":46E44
      ScaleHeight     =   390
      ScaleWidth      =   690
      TabIndex        =   15
      Top             =   4710
      Width           =   690
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   90
      Picture         =   "MailMan Pro.frx":47CBE
      ScaleHeight     =   390
      ScaleWidth      =   690
      TabIndex        =   14
      Top             =   4230
      Width           =   690
   End
   Begin VB.TextBox SUBJECT 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   915
      TabIndex        =   6
      Top             =   2985
      Width           =   3045
   End
   Begin VB.TextBox MAIL_TO 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   915
      TabIndex        =   4
      Top             =   1680
      Width           =   3045
   End
   Begin VB.TextBox FROM 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   915
      TabIndex        =   2
      Top             =   360
      Width           =   3045
   End
   Begin VB.TextBox STATUS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4860
      Width           =   3225
   End
   Begin VB.TextBox SMTP_HOST 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5295
      TabIndex        =   1
      Text            =   "mail.btinternet.com"
      Top             =   1620
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox DATA 
      Appearance      =   0  'Flat
      Height          =   1185
      Left            =   915
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3630
      Width           =   3045
   End
   Begin VB.TextBox RCPT_TO 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   915
      TabIndex        =   5
      Top             =   2325
      Width           =   3045
   End
   Begin VB.TextBox MAIL_FROM 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   915
      TabIndex        =   3
      Top             =   1020
      Width           =   3045
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   2835
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "smtp.kabelfoon.nl"
      RemotePort      =   25
      LocalPort       =   6000
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sender's name:"
      Height          =   255
      Left            =   915
      TabIndex        =   13
      Top             =   120
      Width           =   1125
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   255
      Left            =   915
      TabIndex        =   12
      Top             =   2730
      Width           =   645
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver's name:"
      Height          =   255
      Left            =   915
      TabIndex        =   11
      Top             =   1425
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Body:"
      Height          =   255
      Left            =   915
      TabIndex        =   10
      Top             =   3375
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiver's e-mail address:"
      Height          =   255
      Left            =   915
      TabIndex        =   9
      Top             =   2085
      Width           =   1950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sender's e-mail address:"
      Height          =   255
      Left            =   915
      TabIndex        =   8
      Top             =   765
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Progress
Dim Green_Light As Boolean
Dim DATAFile As String

Private Sub Form_Load()
Progress = 0

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Form_Terminate()
Unload Me
Me.Hide


End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MAIL_RESET_Click()

End Sub

Private Sub Picture1_Click()
'all that below just tests to see if the user has entered text in all the boxes
If FROM.Text = "" Then
Form2.Show
GoTo 1
End If
If MAIL_FROM.Text = "" Then
Form2.Show
GoTo 1
End If
If MAIL_TO.Text = "" Then
Form2.Show
GoTo 1
End If
If RCPT_TO.Text = "" Then
Form2.Show
GoTo 1
End If
If SUBJECT.Text = "" Then
Form2.Show
GoTo 1
End If

'This is where we open the connection to the server and send all the data
Winsock1.Close
Winsock1.Connect SMTP_HOST, "25" 'port 25
Do While Winsock1.State <> sckConnected 'finds out if connected
DoEvents
STATUS.Text = "Connecting to " & SMTP_HOST & ". Please wait." 'adds status to a textbox
Loop
STATUS.Text = "Connected to " & SMTP_HOST & "."

Do While Green_Light = False
DoEvents
STATUS.Text = "Waiting for reply..."
Loop
Winsock1.SendData "MAIL FROM: " & MAIL_FROM & Chr$(13) & Chr$(10) 'it then sends the data out of the text boxes

Do While Progress <> 1
DoEvents
STATUS.Text = "Sending data. (1 of 3)"
Loop
Winsock1.SendData "RCPT TO: " & RCPT_TO & Chr$(13) & Chr$(10)

Do While Progress <> 2
DoEvents
STATUS.Text = "Sending data. (2 of 3)"
Loop
Winsock1.SendData "DATA" & Chr$(13) & Chr$(10)

Do While Progress <> 3
DoEvents
STATUS.Text = "Setting up body transfer..."
Loop
Winsock1.SendData "FROM: " & FROM & " <" & MAIL_FROM & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "TO: " & MAIL_TO & " <" & RCPT_TO & ">" & Chr$(13) & Chr$(10)
Winsock1.SendData "SUBJECT: " & SUBJECT & Chr$(13) & Chr$(10)
Winsock1.SendData Chr$(13) & Chr$(10)
Winsock1.SendData DATA & Chr$(13) & Chr$(10)

Winsock1.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)

Do While Progress <> 4
DoEvents
STATUS.Text = "Sending data. (3 of 3)"
Loop
Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10)
STATUS.Text = "Done"
Winsock1.Close
SMTP_HOST = ""
FROM = ""
MAIL_FROM = ""
MAIL_TO = ""
RCPT_TO = ""
SUBJECT = ""
DATA = ""
STATUS = ""
1:
End Sub

Private Sub Picture2_Click()
Winsock1.Close 'this closes the connection
SMTP_HOST = ""
FROM = ""
MAIL_FROM = ""
MAIL_TO = ""
RCPT_TO = ""
SUBJECT = ""
DATA = ""
STATUS = "" 'making all the textboxes blank
End Sub

Private Sub Picture3_Click()
End 'if you are a REAL beginer this just closes the application
End Sub

Private Sub Picture4_Click()
Me.WindowState = 1 'we then minimize the form by using the windowstate = ( 0 for normal, 1 for minimised, and 3 for maximized)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData DATAFile 'this just recieves the data telling us if sucefful
Reply = Mid(DATAFile, 1, 3)

If Reply = 250 Or Reply = 354 Then
Progress = Progress + 1
End If
If Reply = 220 Then
Green_Light = True
End If
End Sub



