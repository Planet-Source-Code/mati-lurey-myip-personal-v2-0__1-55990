VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyIP Personal :: Version 2.0"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   3465
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   600
      Top             =   3480
   End
   Begin VB.CommandButton cmdReconnect 
      Caption         =   "Reconnect"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1680
      TabIndex        =   5
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Timer tmrReconnect 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1080
      Top             =   3480
   End
   Begin MyIP.DL DL 
      Left            =   2040
      Top             =   2760
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MSWinsockLib.Winsock sckWinsock 
      Left            =   1560
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "www.google.com"
      RemotePort      =   80
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin MyIP.SysTray SysTray1 
      Left            =   1440
      Top             =   2760
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin VB.TextBox txtWANAddress 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "localhost"
      ToolTipText     =   "Double click to copy to clipboard"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtWANIp 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "127.0.0.1"
      ToolTipText     =   "Double click to copy to clipboard"
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change MyIP Settings"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1680
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox txtLANAddress 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Click Listbox above to get a LAN Address"
      ToolTipText     =   "Double click to copy to clipboard"
      Top             =   3840
      Width           =   3255
   End
   Begin VB.ListBox lstLAN 
      Height          =   1425
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Click to copy to textbox"
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox txtCPUName 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "localhost"
      ToolTipText     =   "Double click to copy to clipboard"
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtCPUAddress 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "127.0.0.1"
      ToolTipText     =   "Double click to copy to clipboard"
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "Software firewalls may hide a computer"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   1920
      Width           =   3195
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   960
      MouseIcon       =   "frmMain.frx":014A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":0454
      ToolTipText     =   "PSCode.com: The largest public source code database on the Internet "
      Top             =   6600
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1800
      MouseIcon       =   "frmMain.frx":27BE
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":2AC8
      Stretch         =   -1  'True
      ToolTipText     =   "GRC.com: ShieldsUP Tests your Security"
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Vote for this Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      MouseIcon       =   "frmMain.frx":C4F2
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Tag             =   "http://www.grc.com/x/ne.dll?rh1dkyd2"
      ToolTipText     =   "I spent a couple of hours on this; Please vote or leave feedback!"
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Line Line5 
      X1              =   1920
      X2              =   1920
      Y1              =   7080
      Y2              =   7320
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   3360
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Links:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Ports for Weakness"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmMain.frx":C7FC
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Tag             =   "http://www.grc.com/x/ne.dll?rh1dkyd2"
      ToolTipText     =   "Make sure your computer is secure!"
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Host Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5070
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "IP Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   4710
      Width           =   840
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Information and services provided by:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3360
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "WAN Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "The following are machines connected via LAN"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1710
      Width           =   3315
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3360
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "LAN Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1320
      Width           =   1425
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   360
      MouseIcon       =   "frmMain.frx":CB06
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":CE10
      ToolTipText     =   "IPChicken.com: Get your real IP!"
      Top             =   6120
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Host Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   870
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IP Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   510
      Width           =   840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3360
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Computer Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1950
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuTraySettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnuTraySep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API Code for Launching Default Web Browser
'Used For Hyperlinks
'Special Thanks to 'CoDe ReD CrYsTaL' for this API
'Visit original link here: http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=29557&lngWId=1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Const SW_SHOWNORMAL = 1

'Actually Unload the Form or just hide it?
Public blnUnload As Boolean

'API Class Module to Place Form 'Always On Top'
'Special Thanks to 'Aaron Chan'
'Visit original link here: http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=12144&lngWId=1
Dim OnTop As New clsOnTop
Dim intLAN As Integer
Public Sub Hyperlink(URL As String)
'Wrapper Procedure for ShellExecute
'Paramater 'URL' is launched in Web Browser
ShellExecute Me.hWnd, vbNullString, URL, vbNullString, "C:\", 1
End Sub


Private Sub cmdReconnect_Click()
sckWinsock.Close
sckWinsock.Connect
End Sub

Private Sub cmdRefresh_Click()
lstLAN.Clear

On Error Resume Next

Dim i As Integer
Dim n As String
Dim ip As String
Dim cPing As New clsPing
For i = GetSetting(App.Title, "Settings", "LAN_Start", 0) To GetSetting(App.Title, "Settings", "LAN_End", 5)
    n = i
    Do While Len(n) < 2
    n = "0" & n
    DoEvents
    Loop
    
    ip = "192.168.0.1" & n
    If cPing.Ping(ip) = True Then
    lstLAN.AddItem ip
    End If
DoEvents
Next i
End Sub

Private Sub Command1_Click()
mnuTraySettings_Click
End Sub

Private Sub DL_Complete(Data As Variant)
'Parse information and Display
Dim intFind1 As Integer, intFind2 As Integer, strTemp As String

'Get the IP Address
intFind1 = InStr(1, Data, "<font face=""Verdana, Arial, Helvetica, sans-serif"" size=""5"" color=""#0000FF""><b>") + Len("<font face=""Verdana, Arial, Helvetica, sans-serif"" size=""5"" color=""#0000FF""><b>")
intFind2 = InStr(intFind1, Data, "<br>")
strTemp = Mid(Data, intFind1, intFind2 - intFind1)

'Cut out the spaces and linebreaks
strTemp = Replace(strTemp, Chr(10), "")
strTemp = Replace(strTemp, Chr(13), "")
strTemp = Trim(strTemp)

'Display IP Address
txtWANIp.Text = strTemp

'Get the Host Name/Address
intFind1 = InStr(1, Data, "Address:") + Len("Address:")
intFind2 = InStr(intFind1, Data, "</font>")
strTemp = Mid(Data, intFind1, intFind2 - intFind1)

'Cut out the spaces and linebreaks
strTemp = Replace(strTemp, Chr(10), "")
strTemp = Replace(strTemp, Chr(13), "")
strTemp = Trim(strTemp)

'Display Host Name/Address
txtWANAddress.Text = strTemp
End Sub

Private Sub Form_Load()
'System Tray Functionality
'Thank You Michael Hollifield
SysTray1.AddToSystemTray Me, mnuTray, mnuTray, "MyIP Personal v2.0"

'Get Local Address
txtCPUAddress.Text = sckWinsock.LocalIP
'Get Local Name
txtCPUName.Text = sckWinsock.LocalHostName
'Test Internet Connection
'Will Try to connect to google.com on http (port 80)
sckWinsock.Close
sckWinsock.Connect

'Scan all the LAN Addresses
cmdRefresh_Click

Load frmSettings
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Click on 'X' only hide the application
'Click on 'Close' display Exit Dialog

If blnUnload = False Then
    Cancel = True
    Hide
End If

End Sub


Private Sub Image1_Click()
'Launch IPChicken.com - Used for WAN Information
Hyperlink "http://www.ipchicken.com"
End Sub


Private Sub Image2_Click()
'Launch GRC.com - Used for ShieldsUP Port Scanner
Hyperlink "http://www.grc.com"
End Sub


Private Sub Image3_Click()
'Launch Pscode.com - Used for Resources
Hyperlink "http://www.pscode.com"
End Sub


Private Sub Label10_Click()
'Launch ShieldsUP provided by GRC.com
'This is absolutely *SAFE*
'Remote computer will do a quick port scan of your computer
'And tell you the results
Hyperlink "https://grc.com/x/ne.dll?bh0bkyd2"
End Sub

Private Sub Label12_Click()
'After you are done trying MyIP go ahead and vote or write some feedback
'I dont care what kind of comments; Constructive Criticsim also Appreciated
'MyIP is Freeware and May be used an Unlimited number of times
'On an Unlimited Number of a Computers as long is it remains FREE
Hyperlink "http://pscode.com/vb/scripts/ShowCode.asp?txtCodeId=55990&lngWId=1"
End Sub


Private Sub lstLAN_Click()
txtLANAddress.Text = lstLAN.Text
End Sub

Private Sub mnuTrayExit_Click()

'Exit the Program

'Set Unload = true (do not hide)
blnUnload = True

If GetSetting(App.Title, "Settings", "ExitConfirm", True) = True Then
    'By default show a confirm dialog
    frmExit.Show
Else
    'Override; Exit immediatly
    End
End If

End Sub

Private Sub mnuTraySettings_Click()
frmSettings.LoadUp
frmSettings.Show
End Sub

Private Sub mnuTrayShow_Click()
'If settings tell it to it will display 'OnTop'
'(default = on)
If GetSetting(App.Title, "Settings", "AlwaysOnTop", True) = True Then
    OnTop.MakeTopMost hWnd
    Else
    OnTop.MakeNormal hWnd
End If

Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height - 300

'Refresh the Display

Me.Show
End Sub

Private Sub sckWinsock_Connect()
sckWinsock.Close
'Good! Now download http://www.ipchicken.com and parse it
DL.Download "http://www.ipchicken.com"
End Sub

Private Sub sckWinsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
sckWinsock.Close
'Set timer to reconnect in one minute (if settings say to)
'By default is TRUE
If GetSetting(App.Title, "Settings", "AutoReconnect", True) = True Then
    'One minute timer
    tmrReconnect.Enabled = True
    
    txtWANIp.Text = "Internet Connection Failed!"
    txtWANAddress.Text = "Trying again in one minute..."
    cmdReconnect.Visible = True
End If

'Alert user
frmError.Left = Screen.Width - frmError.Width
frmError.Top = Screen.Height - frmError.Height - 300
OnTop.MakeTopMost frmError.hWnd
frmError.Show

'Text Alerts
txtWANIp.Text = "Internet Connection Failed!"
txtWANAddress.Text = "Press Reconnect to try again"
cmdReconnect.Visible = True
End Sub

Private Sub Timer1_Timer()
intLAN = intLAN + 1
If intLAN >= GetSetting(App.Title, "Settings", "LANMinutes", 2) Then
intLAN = 0
cmdRefresh_Click
End If

End Sub

Private Sub tmrReconnect_Timer()
'Attempt a reconnect
sckWinsock.Connect
tmrReconnect.Enabled = False
End Sub


Private Sub txtCPUAddress_DblClick()
'Copy to clipboard
Clipboard.Clear
Clipboard.SetText txtCPUAddress.Text
End Sub

Private Sub txtCPUAddress_GotFocus()
txtCPUAddress.SelStart = 0
txtCPUAddress.SelLength = Len(txtCPUAddress.Text)
End Sub


Private Sub txtCPUAddress_KeyPress(KeyAscii As Integer)
'Since 'locked' textbox's cant be copied using CTRL-C I will do it this way
If KeyAscii <> 3 Then
KeyAscii = 0
End If
End Sub

Private Sub txtCPUName_DblClick()
'Copy to clipboard
Clipboard.Clear
Clipboard.SetText txtCPUName.Text
End Sub

Private Sub txtCPUName_GotFocus()
txtCPUName.SelStart = 0
txtCPUName.SelLength = Len(txtCPUName.Text)
End Sub

Private Sub txtCPUName_KeyPress(KeyAscii As Integer)
'Since 'locked' textbox's cant be copied using CTRL-C I will do it this way
If KeyAscii <> 3 Then
KeyAscii = 0
End If
End Sub


Private Sub txtLANAddress_DblClick()
'Copy to clipboard
Clipboard.Clear
Clipboard.SetText txtLANAddress.Text
End Sub

Private Sub txtLANAddress_GotFocus()
txtLANAddress.SelStart = 0
txtLANAddress.SelLength = Len(txtLANAddress.Text)
End Sub

Private Sub txtLANAddress_KeyPress(KeyAscii As Integer)
'Since 'locked' textbox's cant be copied using CTRL-C I will do it this way
If KeyAscii <> 3 Then
KeyAscii = 0
End If
End Sub


Private Sub txtWANAddress_DblClick()
'Copy to clipboard
Clipboard.Clear
Clipboard.SetText txtWANAddress.Text
End Sub

Private Sub txtWANAddress_GotFocus()
txtWANAddress.SelStart = 0
txtWANAddress.SelLength = Len(txtWANAddress.Text)
End Sub

Private Sub txtWANAddress_KeyPress(KeyAscii As Integer)
'Since 'locked' textbox's cant be copied using CTRL-C I will do it this way
If KeyAscii <> 3 Then
KeyAscii = 0
End If
End Sub


Private Sub txtWANIp_DblClick()
'Copy to clipboard
Clipboard.Clear
Clipboard.SetText txtWANIp.Text
End Sub

Private Sub txtWANIp_GotFocus()
txtWANIp.SelStart = 0
txtWANIp.SelLength = Len(txtWANIp.Text)
End Sub

Private Sub txtWANIp_KeyPress(KeyAscii As Integer)
'Since 'locked' textbox's cant be copied using CTRL-C I will do it this way
If KeyAscii <> 3 Then
KeyAscii = 0
End If
End Sub


