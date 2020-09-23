VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Click on Menu - Stacked_Shit@hotmail.com"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Left            =   2520
      Top             =   4680
   End
   Begin VB.Timer Timer1 
      Left            =   4560
      Top             =   4680
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H8000000E&
      Caption         =   "Goto Chat Rooms"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H8000000E&
      Caption         =   "Send Email"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H8000000E&
      Caption         =   "About MSN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H8000000E&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H8000000E&
      Caption         =   "Always On Top"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H8000000E&
      Caption         =   "Add A Contact"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H8000000E&
      Caption         =   "Start Netmeeting"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H8000000E&
      Caption         =   "Send IM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H8000000E&
      Caption         =   "Send file or photo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H8000000E&
      Caption         =   "My Hotmail Inbox"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H8000000E&
      Caption         =   "Sign Out"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H8000000E&
      Caption         =   "Sign In"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "MSN Messenger"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   8775
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H8000000E&
      Caption         =   "Offline Messages"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H8000000E&
      Caption         =   "Color Effect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H8000000E&
      Caption         =   "Disable Voice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H8000000E&
      Caption         =   "Imvironments"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H8000000E&
      Caption         =   "Message Archive"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H8000000E&
      Caption         =   "Enable Voice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H8000000E&
      Caption         =   "Sign Out and close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000E&
      Caption         =   "Account info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000E&
      Caption         =   "Edit contact info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000E&
      Caption         =   "My Profiles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000E&
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Yahoo! Messenger"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   8775
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000E&
         Caption         =   "About Y! Messenger"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program by : Stacked_shit@hotmail.com"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1200
      TabIndex        =   27
      Top             =   105
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   70
      Width           =   6615
   End
   Begin VB.Image Image7 
      Height          =   405
      Left            =   8280
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   345
   End
   Begin VB.Image Image6 
      Height          =   330
      Left            =   8640
      Picture         =   "Form1.frx":0364
      Top             =   45
      Width           =   345
   End
   Begin VB.Image Image5 
      Height          =   390
      Left            =   3000
      Picture         =   "Form1.frx":06BA
      Top             =   7680
      Width           =   360
   End
   Begin VB.Image Image4 
      Height          =   405
      Left            =   2640
      Picture         =   "Form1.frx":0A1D
      Top             =   7680
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   2280
      Picture         =   "Form1.frx":0D81
      Top             =   7680
      Width           =   315
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   1920
      Picture         =   "Form1.frx":10BA
      Top             =   7680
      Width           =   345
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   0
      Picture         =   "Form1.frx":1410
      Top             =   0
      Width           =   9705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public xpos As Long, ypos As Long
Private Sub Command1_Click()
Dim i As Long
i = FindWindow("YahooBuddyMain", vbNullString)
If i = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(i, "&About Yahoo! Messenger...")
End Sub

Private Sub Command10_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("ImClass", vbNullString)
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
    Call RunMenubystring(YahooBuddy, "Enable &Voice")
End Sub

Private Sub Command11_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("IMClass", vbNullString)
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
    Call RunMenubystring(YahooBuddy, "Color &Effects...")
End Sub

Private Sub Command12_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("yahoobuddymain", vbNullString)
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
    Call RunMenubystring(YahooBuddy, "Offl&ine Messages")
End Sub

Private Sub Command13_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(msnwindow, "S&ign In...")
End Sub

Private Sub Command14_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(msnwindow, "Sig&n Out")
End Sub

Private Sub Command15_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(msnwindow, "My Ho&tmail Inbox")
End Sub

Private Sub Command16_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(msnwindow, "Send a &File or Photo")
End Sub

Private Sub Command17_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(msnwindow, "&Send and Instant Message")
End Sub

Private Sub Command18_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If

Call RunMenubystring(msnwindow, "Start Netmeeting...")
End Sub

Private Sub Command19_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If

Call RunMenubystring(msnwindow, "&Add a Contact...")
End Sub


Private Sub Command2_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("YahooBuddyMain", vbNullString)
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(YahooBuddy, "&Disconnect")
End Sub

Private Sub Command20_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If

Call RunMenubystring(msnwindow, "Always o&n Top")
End Sub

Private Sub Command21_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(msnwindow, "&Options...")
End Sub

Private Sub Command22_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If

Call RunMenubystring(msnwindow, "&About MSN Messenger")
End Sub

Private Sub Command23_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If

Call RunMenubystring(msnwindow, "Send &E-mail...")
End Sub

Private Sub Command24_Click()
Dim msnwindow As Long
msnwindow = FindWindow("MSBLClass", vbNullString)
If msnwindow = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find MSN Messenger Window , Check If Its Loaded", vbInformation, "MSN Messenger Options"
Call StayOnTop(Me)
End If

Call RunMenubystring(msnwindow, "&Go To Chat Rooms")
End Sub

Private Sub Command3_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("YahooBuddyMain", vbNullString)
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
    Call RunMenubystring(YahooBuddy, "&My Profiles...")
End Sub

Private Sub Command4_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("YahooBuddyMain", vbNullString)
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(YahooBuddy, "&Edit My Contact Info...")
End Sub

Private Sub Command5_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("YahooBuddyMain", vbNullString)
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
    Call RunMenubystring(YahooBuddy, "&Account Info")
End Sub

Private Sub Command6_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("YahooBuddyMain", vbNullString)
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
    Call RunMenubystring(YahooBuddy, "C&lose")
End Sub

Private Sub Command7_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("ImClass", vbNullString)
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(YahooBuddy, "Enable &Voice")
End Sub

Private Sub Command8_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("yahoobuddymain", vbNullString)
    
If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
Call RunMenubystring(YahooBuddy, "Me&ssage Archive")
End Sub

Private Sub Command9_Click()
Dim YahooBuddy&
    YahooBuddy = FindWindow("IMClass", vbNullString)
    If YahooBuddy = 0 Then
Call dontstayontop(Me)
MsgBox "I Couldnt Find Yahoo Messenger Window , Check If Its Loaded", vbInformation, "Yahoo Messenger Options"
Call StayOnTop(Me)
End If
    Call RunMenubystring(YahooBuddy, "I&MVironment")
End Sub

Private Sub Form_Load()
Call StayOnTop(Me)
Timer1.interval = 1
Timer1.Enabled = True
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xpos = X
ypos = Y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Move X + (Me.Left - xpos), Y + (Me.Top - ypos)
End If
End Sub

Private Sub Image6_Click()
Unload Me
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Picture = Image3.Picture
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Picture = Image2.Picture
End Sub

Private Sub Image7_Click()
Me.WindowState = 1
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Picture = Image5.Picture
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Picture = Image4.Picture
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xpos = X
ypos = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Move X + (Me.Left - xpos), Y + (Me.Top - ypos)
End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xpos = X
ypos = Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Move X + (Me.Left - xpos), Y + (Me.Top - ypos)
End If
End Sub

Private Sub Timer1_Timer()
If Label2.Left < 4920 Then
DoEvents
Label2.Left = Label2.Left + 50
    If Label2.Left > 4910 Then
    Timer2.interval = 1
    Timer2.Enabled = True
    Timer1.Enabled = False
    End If
End If
End Sub

Private Sub Timer2_Timer()
If Label2.Left > 1200 Then
DoEvents
Label2.Left = Label2.Left - 50
    If Label2.Left < 1210 Then
    Timer1.interval = 1
    Timer1.Enabled = True
    Timer2.Enabled = False
    End If
End If
End Sub
