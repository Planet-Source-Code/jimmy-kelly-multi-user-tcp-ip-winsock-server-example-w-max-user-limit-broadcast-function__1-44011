VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Winsock Client"
   ClientHeight    =   4875
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   325
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   453
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   3870
      Width           =   6045
   End
   Begin VB.TextBox txtConsole 
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6045
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
   Begin VB.Menu mnuInet 
      Caption         =   "Internet"
      Begin VB.Menu mnuconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
      End
   End
   Begin VB.Menu mnuUsername 
      Caption         =   "Username"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================
' || Multi-user Winsock Client               ||
' =============================================
' || ©Backwoods Interactive 2003             ||
' || Programmed by James J. Kelly Jr.        ||
' =============================================
' ||                                         ||
'\||/                                       \||/

'Feel free to use this for whatever you wish

Dim Username As String

Private Sub Form_Load()

'Get the username from registry
Username = GetSetting("WSCLIENT", "Settings", "UserName")

'If username is empty
'set it to internal default
If Username = "" Then Username = "~DEFAULTUSER~"

End Sub

Private Sub Form_Resize()

'Set the controls size

On Error Resume Next 'In case the window gets to small

txtConsole.Width = Me.ScaleWidth - txtConsole.Left * 2
txtConsole.Height = Me.ScaleHeight - txtMsg.Height - 3 - _
txtConsole.Top * 2

txtMsg.Top = txtConsole.Height + 2
txtMsg.Width = Me.ScaleWidth - txtMsg.Left * 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

'Let everybody know you quit
If Winsock.State = 7 Then Winsock.SendData Username & " has left"
'Close connection
Winsock.Close

End Sub

Private Sub mnuconnect_Click()

'Connect to a server

Dim Result As String

'Get server ip\host from user
Result = InputBox("Enter the ip address or hostname to connect too.", _
"Connect")

'Check to see if it is valid
If Result <> "" Then
'Let everybody know your leaving
If Winsock.State = 7 Then If Winsock.State = 7 Then _
Winsock.SendData Username & " has left"
'Close connection
Winsock.Close
'Set the port
Winsock.RemotePort = 1110
'Set the ip\host
Winsock.RemoteHost = Result
'Connect
Winsock.Connect
End If

End Sub

Private Sub mnuDisconnect_Click()

'Let everybody know your leaving
If Winsock.State = 7 Then Winsock.SendData Username & " has left"
'Disconnect
Winsock.Close

End Sub

Private Sub mnuExit_Click()

Unload Me
End

End Sub

Private Sub mnuUsername_Click()

'Change username

Dim Result As String

'Get new username from user
Result = InputBox("Enter the alias you wish to use", _
"Username", Username)

'Check to see if its valid
If Result <> "" Then
'Let everybody know you changed
If Winsock.State = 7 Then _
Winsock.SendData Username & " is now known as " & Result
'Change username
Username = Result
'Save to registry
SaveSetting "WSCLIENT", "Settings", "UserName", Username
End If

End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)

'Send data if "enter" was pressed
If KeyAscii = 13 Then
 Winsock.SendData "<" & Username & "> " & txtMsg.Text
End If

End Sub

Private Sub Winsock_Close()

'Notify the user of disconnection
txtConsole.Text = txtConsole.Text & "Connection with " & _
Winsock.RemoteHostIP & " lost" & vbCrLf

End Sub

Private Sub Winsock_Connect()

'Notify the user of connection
txtConsole.Text = txtConsole.Text & "Connection with " & _
Winsock.RemoteHostIP & " established" & vbCrLf

End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)

'Process recieved data
Dim Result As String
Winsock.GetData Result, vbString

'Add it
txtConsole.Text = txtConsole.Text + Result & vbCrLf

End Sub

