VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Winsock Server"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   519
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
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
      Top             =   5040
      Width           =   7785
   End
   Begin VB.TextBox txtConsole 
      Height          =   4950
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7785
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================
' || Multi-user Winsock Server               ||
' =============================================
' || ©Backwoods Interactive 2003             ||
' || Programmed by James J. Kelly Jr.        ||
' =============================================
' ||                                         ||
'\||/                                       \||/

'This is a simple sample code for building a
'mutli user Winsock server.
'Feel free to use the code for whatever you
'wish. However, credits are greatly appreciated.

'Function List:
'ADDLINE - Adds a line to textbox
'BROADCAST - Broadcasts a message
'GetConCount - Retrieves the number of connected computers
'Wait - Suspends execution
'KickUser - Kicks a user according to they're hostname or ip

Option Explicit

'Declarations
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Program variables
Const MaxUsers As Long = 10

Function Wait(ByVal Waitlen As Long, ByVal UseAPI As Boolean)

'This delays execution of a vb app.
'If you have trouble using it turn UseAPI to false.
'However, turning off UseAPI requires more
'CPU.

On Error GoTo ErrShow

If UseAPI = False Then
Dim WS As Long
WS = GetTickCount + Waitlen
 Do While GetTickCount < WS
  DoEvents
 Loop
Else
Sleep Waitlen
End If

Exit Function

ErrShow:
AddLine "Error : " & _
Err.Number & " : " & Err.Description

End Function

Private Sub Form_Load()

On Error GoTo ErrShow
'Close just in case
Winsock(0).Close
'Make it so that they're is a fixed port to connect
'too. I like to call this the "join" port
Winsock(0).Bind 1110, Winsock(0).LocalIP
'Listen for connections
Winsock(0).Listen

'Show the form
Me.Show

'Exit the sub
Exit Sub

ErrShow:
MsgBox "Error : " & _
Err.Number & " : " & Err.Description, vbOKOnly + vbCritical, _
"Error " & Err.Number

Unload Me
End

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

'Unload everything to clear up memory

Dim i As Long

'Set a variable containing the
'number of controls
Dim WSCnt As Long
WSCnt = Winsock.Count - 1

'Is control count greater than 0?
If WSCnt > 0 Then
'Start loop if so
 For i = 1 To WSCnt
 'Is it something it can unload?
  If Not Winsock(i) Is Nothing Then
  'If so then do so
   Winsock(i).Close 'Close just in case
   Unload Winsock(i) 'Unload it from memory
  End If
 Next i
End If

'Unload the form
Unload Me
'End the program
End

End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)

'Make it so that when "Enter" is pressed
'broadcast the data
If KeyAscii = 13 Then

If LCase(Left(txtMsg.Text, 6)) = "/kick " Then
 KickUser Mid(txtMsg.Text, 7, Len(txtMsg.Text))
 Exit Sub
End If

If LCase(Left(txtMsg.Text, 5)) = "/help" Then
 AddLine "/help - displays known commands"
 AddLine "/KickUser [IP\HOST] - Kicks a user according to " + _
 "they're IP address or hostname"
 Exit Sub
End If

 BroadCast txtMsg.Text
 txtMsg.Text = ""
End If

End Sub

Private Sub Winsock_Close(Index As Integer)

On Error GoTo ErrShow
'Display a connection was closed
AddLine "Computer " & Winsock(Index).RemoteHostIP & " has diconnected"
'Close the connection just in case
Winsock(Index).Close
'Unload control from memory
Unload Winsock(Index)

'Exit the sub
Exit Sub

ErrShow:
AddLine "Error : " & _
Err.Number & " : " & Err.Description

End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)

On Error GoTo ErrShow
Dim i As Long

'Set a variable on the number of
'Winsock controls
Dim WinCnt As Long
WinCnt = Winsock.Count - 1

'Check to see if they're are any non-existant
'winsock controls first

'If they're arent any skip this
If WinCnt > 0 Then
'Loop thru the controls
 For i = 0 To WinCnt
 'If a empty control is found then
 'take it
  If Winsock(i) Is Nothing Then
  'Load control into memory
   Load Winsock(i)
   'Close connection in case
   Winsock(i).Close
   'Accept request
   Winsock(i).Accept requestID
   'Set variable to the selected control
   WinCnt = i
   'Goto final steps
   GoTo FinishConnection
  End If
 Next i
End If

'Wait!, how many users are they're?
'Lets check before adding a new user
If GetConCount > MaxUsers Then
Winsock(0).Close
Winsock(0).Listen
GoTo ConMany
End If

'If one wasnt found then
'make a new one

'Set variable to the new one
WinCnt = Winsock.Count
'Load control into memory
Load Winsock(WinCnt)
'Close the connection in case
Winsock(WinCnt).Close
'Accept connection
Winsock(WinCnt).Accept requestID

FinishConnection:
'Alert user a computer has connected
AddLine "Computer " & Winsock(WinCnt).RemoteHostIP & " has connected"

Exit Sub

'The following is all error messages

ErrShow:
'General errors
AddLine "Error : " & _
Err.Number & " : " & Err.Description

Exit Sub

ConMany:
'Server full error
AddLine "A computer attempted to connect. " + _
"However, server is full. Connection refused"

Exit Sub

End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

On Error GoTo ErrShow

'Declare a buffer variable
Dim WSData As String
'Process the data and fill the buffer
Winsock(Index).GetData WSData

'Add a line to the console
AddLine Winsock(Index).RemoteHostIP & " :: " & WSData
'Send the data back for everybody to see
BroadCast WSData

Exit Sub

ErrShow:
AddLine "Error : " & _
Err.Number & " : " & Err.Description

End Sub

Function BroadCast(ByVal StrData As String)

'Broadcast function
'Broadcasts data to all
'connected clients

On Error GoTo ErrShow
Dim i As Long

'Set a variable containing the
'number of controls
Dim WinCnt As Long
WinCnt = Winsock.Count - 1

'Loop thru and send data
For i = 0 To WinCnt
'Check to see if control exists
 If Not Winsock(i) Is Nothing Then
 'If so is it connected?
  If Winsock(i).State = 7 Then
  'If so then send it
   Winsock(i).SendData StrData
  End If
 End If
'Wait 10 MS before sending next message
Wait 10, True
Next i

Exit Function

ErrShow:
AddLine "Error : " & _
Err.Number & " : " & Err.Description

End Function

Function GetConCount() As Long

'GetConCount function
'Retrieves the number of connected
'clients

On Error GoTo ErrShow
Dim i As Long

'Set up function variables

'Control count
Dim WinCnt As Long
'Connection Count
Dim ConCnt As Long
'Set connection count to 0 by default
ConCnt = 0
'Set the control count
WinCnt = Winsock.Count - 1

'Loop thru to find connected users
For i = 0 To WinCnt
'Is control connected?
 If Winsock(i).State = 7 Then
 'If so then add it to connection count
  ConCnt = ConCnt + 1
 End If
Next i

'Return the connection count
GetConCount = ConCnt

Exit Function

ErrShow:
AddLine "Error : " & _
Err.Number & " : " & Err.Description

End Function

Function AddLine(ByVal StrData As String)

'AddLine function
'A quick function for adding
'stuff to a textbox

'Add the line
txtConsole.Text = txtConsole.Text & StrData & vbCrLf

End Function

Function KickUser(ByVal IP As String)

'Kick user function
'Kicks a user with they're hostname or ip address

On Error GoTo ErrShow
Dim i As Long

'Make a variable containing
'the number of controls
Dim WSCnt As Long
WSCnt = Winsock.Count - 1

'Is the number of controls greater
'than 0?
If WSCnt > 0 Then
'If so then start loop
 For i = 1 To WSCnt
 'Is the connections ip or hostname equal to the IP param?
  If Winsock(i).RemoteHostIP = IP Or _
  Winsock(i).RemoteHost = IP Then
  'If so then close and unload the control
   Winsock(i).Close
   Unload Winsock(i)
  End If
 Next i
End If

Exit Function

ErrShow:
ErrShow:
AddLine "Error : " & _
Err.Number & " : " & Err.Description

End Function
