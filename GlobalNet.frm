VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GlobalNET BBS Server"
   ClientHeight    =   4215
   ClientLeft      =   150
   ClientTop       =   810
   ClientWidth     =   10305
   Icon            =   "GlobalNet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Finger 
      Left            =   4200
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Index           =   0
      Left            =   3720
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3240
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   8280
      TabIndex        =   17
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Unload"
      Height          =   255
      Left            =   8280
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   9240
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Listen"
      Height          =   255
      Left            =   8280
      TabIndex        =   9
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Kick"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2085
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "GlobalNet.frx":0E42
      Top             =   3120
      Width           =   10095
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtf2 
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"GlobalNet.frx":0F00
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtf1 
      Height          =   2535
      Left            =   4560
      TabIndex        =   8
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"GlobalNet.frx":0FD6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9840
      TabIndex        =   21
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9840
      TabIndex        =   20
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Receive:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8280
      TabIndex        =   19
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Send:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8280
      TabIndex        =   18
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Controls:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   8400
      TabIndex        =   15
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Global Users:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Socket Properties:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8280
      TabIndex        =   13
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Message of the Day:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Menu Server 
      Caption         =   "Server"
      Begin VB.Menu Listen 
         Caption         =   "Listen"
      End
      Begin VB.Menu Closee 
         Caption         =   "Close"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu Unloaddd 
         Caption         =   "Unload All Sockets"
      End
      Begin VB.Menu dash4 
         Caption         =   "-"
      End
      Begin VB.Menu QuitServer 
         Caption         =   "Quit Server"
      End
   End
   Begin VB.Menu Helpandstuph 
      Caption         =   "Help & Stuph"
      Begin VB.Menu help 
         Caption         =   "Help"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu Version 
         Caption         =   "Version"
      End
      Begin VB.Menu goodies 
         Caption         =   "Goodies"
         Begin VB.Menu Telnett 
            Caption         =   "Telnet To Server"
         End
         Begin VB.Menu MakeAnnoucenemt 
            Caption         =   "Make An Announcement"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NotReg As Boolean
Dim GU As Integer
Dim Sett As Integer
Dim FNameNew As String
Dim Fingerindex As String


Private Sub Finger2(Name As String)
Finger.SendData "/w " & Name & vbCrLf
End Sub

Private Sub about_Click()
Form2.Show
Form2.Text1.Text = "This program was c0ded in VisualBasic 6.0 By JadaCyrus@hotmail.com . Please do not copy code straight from this project. But you can use it as a reference. Also Im looking for some really good VB Programmers to help me build the next version. If your interested contact: JadaCyrus@hotmail.com . Well Thats all folks..Peace. Click to Close"
Form2.Label2.Caption = " "
Form2.Timer1.Enabled = True


End Sub

Private Sub Closee_Click()
Command1_Click

End Sub

Private Sub Command1_Click()
Winsock1.Close
rtf2.SelColor = vbRed
rtf2.SelText = "Server Status: "
rtf2.SelColor = vbRed
rtf2.SelColor = RGB(202, 202, 202)
rtf2.SelText = " Port Closed" & vbCrLf
rtf2.SelColor = RGB(202, 202, 202)

End Sub

Private Sub Command2_Click()
On Error GoTo errorwhocares
If List1.ListCount = 1 Then
List1.ListIndex = 0
List1.RemoveItem List1.ListIndex
For i = 1 To GU
Winsock2(i).Close
Unload Winsock2(i)
Next i
GU = GU - 1
Text2.Text = 0
Else
TempIp = Winsock2(List1.ListIndex + 1).RemoteHostIP
Winsock2(List1.ListIndex + 1).Close
List1.RemoveItem List1.ListIndex
Unload Winsock2(List1.ListIndex + 1)
GU = GU - 1
Text2.Text = Text2.Text - 1
For i = 1 To GU
Winsock2(i).SendData "User of Global Net: " & TempIp & " has been kicked from the server!" & vbCrLf
Next i
End If
errorwhocares:
'a
End Sub

Private Sub Command3_Click()
Winsock1.Close
Winsock1.LocalPort = 28
Winsock1.Listen
rtf2.SelColor = vbRed
rtf2.SelText = "Server Status: "
rtf2.SelColor = vbRed
rtf2.SelColor = vbGreen
rtf2.SelText = " Listening..." & vbCrLf
rtf2.SelColor = vbGreen

End Sub

Private Sub Command4_Click()
On Error GoTo errordebugthis
For i = 1 To Text2.Text
Winsock2(i).Close
Unload Winsock2(i)
rtf2.SelText = vbCrLf & "(" & i & ")"
rtf2.SelColor = vbGreen
rtf2.SelBold = True
rtf2.SelText = "Unloaded from the system" & vbclrf
rtf2.SelColor = vbRed
rtf2.SelBold = True
Next i
Text2.Text = "0"
List1.Clear
Exit Sub
errordebugthis:
'blah

End Sub

Private Sub Finger_Connect()
Finger.SendData "/w " & FNameNew & vbCrLf
End Sub

Private Sub Finger_DataArrival(ByVal bytesTotal As Long)
Dim FingReturn As String
Winsock2(Fingerindex).SendData FingReturn & vbCrLf
rtf2.SelText = FingReturn & vbclrf

End Sub

Private Sub Form_Load()
GU = 0
Sett = 0
NotReg = True
Winsock1.LocalPort = 28
Winsock1.Listen
File1(0).Path = "C:\Articles\"
Dim MessageRecord As String
MessageRecord = Time & " / " & Day(Now) & "/" & Month(Now) & "/" & Year(Now)

End Sub

Private Sub ProxyDetect_Connect()

End Sub

Private Sub ProxyDetect_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

 End Sub

Private Sub Help_Click()
Form2.Show
Form2.Text1.Text = "This server still has alot of Bugs that need to be corrected. But other than that the code is pretty good. The server commands are case senstitive and should be used with flawless syntax.naturally..If you don't know how to operate the server ur a moron.Thats all..Click to close"
Form2.Label2.Caption = " "
Form2.Timer1.Enabled = True

End Sub

Private Sub Listen_Click()
Command3_Click

End Sub

Private Sub MakeAnnoucenemt_Click()
AskAnnounce = InputBox("Message:")
For i = 1 To GU
 For a = 1 To 25
 Winsock2(GU).SendData vbCrLf
 Next a
Winsock2(GU).SendData "*******GLOBAL ANNOUNCEMENT********" & vbCrLf
Winsock2(GU).SendData "* System Adminstrator JadaCyrus:  *" & vbCrLf
Winsock2(GU).SendData AskAnnounce & vbCrLf
Winsock2(GU).SendData "*******END GLOBAL ANNOUNCEMENT****" & vbCrLf
Next i

End Sub

Private Sub QuitServer_Click()
rtf2.SelColor = vbRed
rtf2.SelText = "Server Status: "
rtf2.SelColor = vbRed
rtf2.SelColor = RGB(32, 32, 32)
rtf2.SelText = " Shutting Down..." & vbCrLf
rtf2.SelColor = RGB(32, 32, 32)
Command1_Click
Command4_Click
List1.Clear

End Sub

Private Sub rtf1_Change()
rtf1.SelStart = Len(rtf1)

End Sub

Private Sub telnettoserver_Click()
X = Shell("telnet localhost 28", vbNormalFocus)

End Sub

Private Sub Telnett_Click()
X = Shell("telnet localhost 28", vbNormalFocus)


End Sub

Private Sub Unloaddd_Click()
Command4_Click

End Sub

Private Sub Version_Click()
MsgBox vbclrf & "Version:" & vbCrLf & "1.2.0"

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''   This is the ear of the program.                          ''
''   One of the most important parts.                         ''
''   It listens for connections on prt 28.                    ''
''   When a connection is made it loads it...                 ''
''   into Winsock2(). Winsock2 holds all of the..             ''
''   connections made. And has the power to drop conenctions  ''
''   After the connection is loaded into the server Winsock1  ''
''   closes itself and begins listening again for other       ''
''   connections.                                             ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo WhoGivesAShit
NotReg = True
GU = GU + 1
Text2.Text = Text2.Text + 1
Load Winsock2(Text2.Text)
Load Text3(GU)
Load Text4(GU)
Load File1(GU)
Text3(GU).Left = Text3(0).Left
Text3(GU).Top = Text3(0).Top
Text4(GU).Left = Text4(0).Left
Text4(GU).Top = Text4(0).Top
Text3(GU).Visible = False
Text4(GU).Visible = False
File1(GU).Visible = False
File1(GU).Path = File1(0).Path
File1(GU).Top = File1(0).Top
File1(GU).Left = File1(0).Left
File1(GU).Width = File1(0).Width
File1(GU).Height = File1(0).Height
Winsock2(GU).Accept requestID
rtf1.SelColor = vbRed
rtf1.SelText = vbCrLf & "(" & GU & ")"
rtf1.SelColor = vbRed
rtf1.SelBold = True
rtf1.SelColor = vbWhite
rtf1.SelText = " Connection Established " & vbCrLf
rtf1.SelColor = vbWhite
rtf1.SelBold = True
rtf2.SelColor = vbRed
rtf2.SelText = "Connection: "
rtf2.SelColor = vbRed
rtf2.SelColor = vbGreen
rtf2.SelText = Winsock2(GU).RemoteHostIP & vbCrLf
rtf2.SelColor = vbGreen
Winsock2(GU).SendData "##########################################################" & vbCrLf
Winsock2(GU).SendData "# NX1 Global Net Servers are for # This BBS Was written  #" & vbCrLf
Winsock2(GU).SendData "# public use. Logging is on.     # and developed by:     #" & vbCrLf
Winsock2(GU).SendData "# NX1 Global Net Administrators  # JadaCyrus@hotmail.com #" & vbCrLf
Winsock2(GU).SendData "# reserve the right to kill any  # and use of it w/out   #" & vbCrLf
Winsock2(GU).SendData "# and all connected user(s).     # explicit permission   #" & vbCrLf
Winsock2(GU).SendData "# without warning or notice.     # from the author is    #" & vbCrLf
Winsock2(GU).SendData "# Version 1.0                    # Strictly Prohibted    #" & vbCrLf
Winsock2(GU).SendData "##########################################################" & vbCrLf
Winsock2(GU).SendData "You Will BE Identified By Your Ip/HostName" & vbCrLf
Winsock2(GU).SendData ":::::::::::::::::::::::::::::::::::::::::::::::" & vbCrLf
Winsock2(GU).SendData "!!!!>>>>>>>>>Type !help for help<<<<<<<<<!!!!!" & vbCrLf
Winsock2(GU).SendData ":::::::::::::::::::::::::::::::::::::::::::::::" & vbCrLf
Winsock2(GU).SendData "The time is: " & Time
Winsock2(GU).SendData "All Commands Are Case Sensitive" & vbCrLf
Winsock2(GU).SendData "*** Exclamation Characters , '!' , are the universal delimiters" & vbCrLf
Winsock2(GU).SendData "@@Message of the Day@@" & vbCrLf
Winsock2(GU).SendData Text1.Text & vbCrLf
Winsock2(GU).SendData "@@/MOTD@@" & vbCrLf
List1.AddItem Winsock2(GU).RemoteHostIP
Winsock1.Close
Winsock1.LocalPort = 28
Winsock1.Listen
 Exit Sub
WhoGivesAShit:
 Winsock1.Close
 Winsock1.LocalPort = 28
 Winsock1.Listen
 rtf2.SelText = "Error!" & vbCrLf
 rtf2.SelColor = vbRed
 Winsock2(GU).Close
 Unload Winsock2(GU)
 Unload Text3(GU)
 Unload Text4(GU)
 Unload File1(GU)
 Text2.Text = Text2.Text - 1
 GU = GU - 1


End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
rtf2.SelText = Description & vbCrLf
rtf2.SelColor = vbRed

End Sub

Private Sub Winsock2_Close(Index As Integer)
If List1.ListCount = 1 Then
List1.ListIndex = 0
List1.RemoveItem List1.ListIndex
GU = GU - 1
Text2.Text = Text2.Text - 1
Winsock2(Index).Close
Unload Winsock2(Index)
rtf1.SelText = vbCrLf & "(" & Index & ")"
rtf1.SelColor = vbRed
rtf1.SelBold = True
rtf1.SelText = " Has Disconnected from the system"
rtf1.SelColor = RGB(95, 239, 33)
rtf1.SelBold = True
Else
List1.RemoveItem Index - 1
rtf1.SelText = vbCrLf & "(" & Index & ")"
rtf1.SelColor = vbRed
rtf1.SelBold = True
rtf1.SelText = " Has Disconnected from the system"
rtf1.SelColor = RGB(95, 239, 33)
rtf1.SelBold = True
GU = GU - 1
Text2.Text = Text2.Text - 1
Winsock2(Index).Close
Unload Winsock2(Index)
End If

     
End Sub

Private Sub Winsock2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''         This is the heart of the server itself                         ''
''         It is where the majority of the code is written                ''
''         It computes all commands made by the users and returns the     ''
''         correct response for each individual command to its right user ''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
Dim Data As String
Winsock2(Index).GetData Data, vbString
 buffer = buffer & Data
  Winsock2(Index).SendData Data
  ProgressBar2.Value = ProgressBar2.Value + 1
  Label7.Caption = Label7.Caption + 1
  If ProgressBar2.Value = 100 Then
  ProgressBar2.Value = 1
  End If
 Text4(Index).Text = Text4(Index).Text & Data
 '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!BACKSPACE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
 If InStr(Data, Chr(8)) Then 'if the backspace key is pressed
 stringg = Left(Text4(Index).Text, Len(Text4(Index).Text) - 1) 'get the string
 Text4(Index).Text = stringg 'replace the tempbuffer with the string
 Stringg2 = Left(stringg, Len(stringg) - 1)
 Text4(Index).Text = Stringg2
 rtf1.SelColor = RGB(176, 9, 9)
 rtf1.SelText = Stringg2 & vbCrLf
 rtf1.SelColor = RGB(176, 9, 9)
 End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!ENTER!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If InStr(Data, Chr(13)) Then 'The user hits return!
rtf1.SelText = rtf1.Text & vbCrLf
rtf1.SelColor = vbRed
rtf1.SelText = "(" & Index & ")"
rtf1.SelColor = vbRed
rtf1.SelBold = True
rtf1.SelColor = RGB(6, 57, 210)
rtf1.SelText = Text4(Index).Text
rtf1.SelColor = RGB(6, 57, 210)

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!HELP!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If InStr(Text4(Index).Text, "!help") Then 'The user typed !help
  Text4(Index).Text = "" 'clear the saved buffer
  Winsock2(Index).SendData "Commands Are:" & vbCrLf 'send
  Winsock2(Index).SendData "!list - lists other users" & vbCrLf 'the
  Winsock2(Index).SendData "!post !message !title - posts an article" & vbCrLf 'list
  Winsock2(Index).SendData "!view - lists all articles" & vbCrLf 'of
  Winsock2(Index).SendData "!read <ItemNumber> - reads a specific article" & vbCrLf 'commands
  Winsock2(Index).SendData "!finger <host> <name> - Finger Lookup Service" & vbCrLf
  Winsock2(Index).SendData "!reply !ItemNumber !Message - Replies to an existing post" & vbCrLf
  Winsock2(Index).SendData "!msg !UserIndex !Message - Delivers a small text message to another user on the system" & vbCrLf
  Winsock2(Index).SendData "!MOTD - Displays the Message of the Day" & vbCrLf
  Winsock2(Index).SendData "!CLS - Clears all data on screen" & vbCrLf
  
End If 'end the user typed !help
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!CLS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If InStr(Text4(Index).Text, "!CLS") Then
For i = 1 To 300
Winsock2(Index).SendData vbCrLf
Next i
Text4(Index).Text = ""
End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!MOTD!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If InStr(Text4(Index).Text, "!MOTD") Then
Winsock2(Index).SendData Text1.Text & vbCrLf
End If
 '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!FINGER!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If InStr(Text4(Index).Text, "!finger ") Then
Winsock2(Index).SendData "Finger Request Sent..Waiting for data feedback" & vbCrLf
Fhost = Split(Text4(Index).Text, " ")(1)
FName = Split(Text4(Index).Text, " ")(2)
FNameNew = Left(FName, Len(FName) - 2)
Finger.RemoteHost = Fhost
Finger.RemotePort = 79
Finger.Connect
Fingerindex = Index
End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!MSG!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If InStr(Text4(Index).Text, "!msg ") Then
On Error GoTo ErrorMsg
userindex = Split(Text4(Index).Text, "!")(2)
Message = Split(Text4(Index).Text, "!")(3)
NewMessage = Left(Message, Len(Message) - 2)
For a = 1 To 25
Winsock2(userindex).SendData vbCrLf
Next a
Winsock2(userindex).SendData "**Incoming Message From: " & Winsock2(Index).RemoteHostIP & " ***" & vbCrLf
Winsock2(userindex).SendData "--------------------------------------------------------------------------------" & vbCrLf
Winsock2(userindex).SendData NewMessage & vbCrLf
Winsock2(userindex).SendData "--------------------------------------------------------------------------------" & vbCrLf
Winsock2(Index).SendData "***Message Sent Successfully***" & vbCrLf
End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!REPLY!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If InStr(Text4(Index).Text, "!reply ") Then
 On Error GoTo ErrorReply
  ItemNo = Split(Text4(Index).Text, "!")(2)
  RepMessage = Split(Text4(Index).Text, "!")(3)
   File1(Index).ListIndex = ItemNo
   RepMessageNew = Left(RepMessage, Len(RepMessage) - 2)
   Open "C:\Articles\" & File1(Index).FileName For Input As #4
   TempMessage = Input(LOF(4), (4))
   Close #4
   Open "C:\Articles\" & File1(Index).FileName For Output As #3
   Print #3, TempMessage
   Print #3, "***[Reply By: " & Winsock2(Index).RemoteHostIP & " On: " & Time & " / " & Day(Now) & "/" & Month(Now) & "/" & Year(Now) & "]***"
   Print #3, RepMessageNew
   Print #3, "--------------------------------------------------------"
   Close #3
 Winsock2(Index).SendData "Your Reply has been added to Item Number " & ItemNo & vbCrLf
End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!POST!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If InStr(Text4(Index).Text, "!post ") Then
 On Error GoTo ErrorPost
  Message = Split(Text4(Index).Text, "!")(2)
  Title = Split(Text4(Index).Text, "!")(3)
     '!post !message !title
     NewTitle = Left(Title, Len(Title) - 2)
     Text4(Index).Text = "C:\Articles\" & NewTitle & ".txt"
     Close #2
  Trim (Title)
Open Text4(Index).Text For Output As #2
  Print #2, "--------------------------------------------------------------------------------"
  Print #2, "-Message Posted By: " & Winsock2(Index).RemoteHostIP & " On: " & Time & " / " & Day(Now) & "/" & Month(Now) & "/" & Year(Now)
  Print #2, "--------------------------------------------------------------------------------"
  Print #2, "-Message Text:"
  Print #2, Message
  Print #2, "--------------------------------------------------------------------------------"
  Print #2, "-End Of File"
  Print #2, "--------------------------------------------------------------------------------"
Close #2
  Winsock2(Index).SendData "Message Archived Successfully!" & vbCrLf
File1(Index).Refresh
End If

  
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!READ!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
If InStr(Text4(Index).Text, "!read ") Then
On Error GoTo errorfileopen
  Trim (Text4(Index).Text)
  Leng = Len(Text4(Index).Text)
   FileNumber = Right(Text4(Index).Text, Leng - 6)
    rtf1.SelText = TheFileName & vbCrLf
    File1(Index).ListIndex = FileNumber
    RFILENAME = File1(Index).FileName
   Open File1(Index).Path & "\" & RFILENAME For Input As #1
  FileData = Input(LOF(1), (1))
  For a = 1 To 25
  Winsock2(Index).SendData vbCrLf
  Next a
  Winsock2(Index).SendData FileData & vbCrLf
  Close #1
  Text4(Index).Text = ""
End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!VIEW!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


If InStr(Text4(Index).Text, "!view") Then
On Error GoTo ErrorFileList
  Text4(Index).Text = ""
  File1(Index).ListIndex = -1
  For i = 1 To File1(Index).ListCount
  File1(Index).ListIndex = File1(Index).ListIndex + 1
  Winsock2(Index).SendData File1(Index).ListIndex & ". " & File1(Index).FileName & vbCrLf
 Next i
 File1(Index).ListIndex = 1 'return the index back to negative one
End If

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!LIST!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


If InStr(Text4(Index).Text, "!list") Then
  Text4(Index).Text = ""
  Winsock2(Index).SendData "List Of Users:" & vbCrLf
  Winsock2(Index).SendData "Ip Address:    User Index" & vbCrLf
 For i = 1 To GU
  Winsock2(Index).SendData Winsock2(i).RemoteHostIP & "    " & i & vbCrLf
Next i
End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End If
Exit Sub
ErrorFileList:
  Winsock2(Index).SendData "-ERROR Listing File Archive. This could be an internal Server error" & vbCrLf
ErrorPost:
  Winsock2(Index).SendData "-ERROR Writing To Disk-Correct your synatx-" & vbCrLf
    Text4(Index).Text = ""
errorfileopen:
  Winsock2(Index).SendData "-ERROR Opening File-Please correct your syntax" & vbCrLf
  Text4(Index).Text = ""

ErrorReply:
   Winsock2(Index).SendData "ERROR Appending Reply to File" & vbCrLf
    Text4(Index).Text = ""
ErrorMsg:
    Winsock2(Index).SendData "ERROR Delivering Message to user-Check syntax and try again-" & vbCrLf
    
End Sub

Private Sub Winsock2_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
If List1.ListCount = 1 Then
List1.ListIndex = 0
List1.RemoveItem List1.ListIndex
Text2.Text = Text2.Text - 1
GU = GU - 1
rtf1.SelText = "(" & Index & ")"
rtf1.SelColor = vbRed
rtf1.SelText = Description & vbCrLf
rtf1.SelColor = vbGreen
Winsock2(Index).Close
Unload Winsock2(Index)
Else
List1.RemoveItem Index - 1
Text2.Text = Text2.Text - 1
GU = GU - 1
rtf1.SelText = "(" & Index & ")"
rtf1.SelColor = vbRed
rtf1.SelText = Description & vbCrLf
rtf1.SelColor = vbGreen
Winsock2(Index).Close
Unload Winsock2(Index)
End If

End Sub

Private Sub Winsock2_SendComplete(Index As Integer)
ProgressBar1.Value = 0

End Sub

Private Sub Winsock2_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
If bytesSent > 100 Then
Label6.Caption = Label6.Caption + bytesSent
ProgressBar1.Value = 99
Else
ProgressBar1.Value = bytesSent
Label6.Caption = Label6.Caption + bytesSent
If ProgressBar1.Value >= 100 Then
ProgressBar1.Value = 1
End If
End If




End Sub
