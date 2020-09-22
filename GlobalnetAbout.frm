VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10470
   LinkTopic       =   "Form2"
   ScaleHeight     =   5190
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   2400
   End
   Begin VB.TextBox Text1 
      Height          =   1485
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "GlobalnetAbout.frx":0000
      Top             =   3360
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   5
      Height          =   4935
      Left            =   0
      Top             =   240
      Width           =   10455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   4455
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   10095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "       Info Screen - [w00t]"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim user As Integer
'Variables
'Switch to turn drah on and off.
Dim MoveScreen As Boolean
'Vars to get the mouse position on the form.
'   you are draging.
Dim MousA As Integer
Dim MousB As Integer
'Vars for moving the form.
Dim CurrX As Integer
Dim CurrY As Integer
Dim StringPos As Integer
Private Sub Form_Load()
StringPos = 0
Label2.Caption = " "


End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

MoveScreen = True ' make form movable while the mouse is down.
    'Get the initial coordinates of the mouse on the form.
    MousA = X
    MousB = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If the mouse is down on the form then.
    If MoveScreen Then
        'Calculate the new x,y position for the form.
        '   NB. This is dependant on the X and Y vars on the Form_MouseMove,
        '   you can use objects MouseMove function. i.e. a Label or Textbox.
        CurrX = Form2.Left - MousA + X
        CurrY = Form2.Top - MousB + Y

        'Move the form to the new X,Y.
        Form2.Move CurrX, CurrY     ' move form.
        End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveScreen = False

End Sub

Private Sub Label2_Click()
Me.Hide
Form1.Show
Timer1.Enabled = False
Label2.Caption = " "
StringPos = "0"

End Sub

Private Sub Timer1_Timer()
StringPos = StringPos + 1
Label2.Caption = Mid(Text1.Text, 1, StringPos)
PlayWav "poorbeep.wav", SND_ASYNC
If Label2.Caption = Text1.Text Then
Timer1.Enabled = False
End If


End Sub
Public Function PlayWav(strPath As String, sndVal As sndConst)
    sndPlaySound strPath, sndVal
End Function

