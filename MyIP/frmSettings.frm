VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MyIP Personal :: Settings"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7020
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Text            =   "2"
      Top             =   3120
      Width           =   255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Scan again every           minutes"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   5655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      TabIndex        =   12
      Text            =   "05"
      Top             =   2770
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Text            =   "01"
      Top             =   2770
      Width           =   255
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Automatically reconnect within one minute if WAN does not respond"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   5655
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Make MyIP Always On Top"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Display confirmation dialog before exiting"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "to 192.168.0.1"
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   2775
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Scan from 192.168.0.1"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2770
      Width           =   1665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Local Area Network:"
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
      TabIndex        =   8
      Top             =   2400
      Width           =   1680
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "WAN Settings:"
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
      TabIndex        =   6
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "General Settings:"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3360
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub LoadUp()
On Error Resume Next

'Confirmation on Exit
If GetSetting(App.Title, "Settings", "ExitConfirm", True) = True Then
    Check1.Value = vbChecked
Else
    Check1.Value = 0
End If

'Always On Top
If GetSetting(App.Title, "Settings", "AlwaysOnTop", True) = True Then
    Check2.Value = vbChecked
Else
    Check2.Value = 0
End If

'Reconnect WAN In One Minute
If GetSetting(App.Title, "Settings", "AutoReconnect", True) = True Then
    Check3.Value = vbChecked
Else
    frmMain.tmrReconnect.Enabled = False
    Check3.Value = 0
End If

'Re-Scan LAN
If GetSetting(App.Title, "Settings", "LANScan", True) = True Then
    Check4.Value = vbChecked
    frmMain.Timer1.Enabled = True
Else
    Check4.Value = 0
    frmMain.Timer1.Enabled = False
End If

'Starting LAN
Text1 = GetSetting(App.Title, "Settings", "LANStart", "00")
'Ending Lan
Text2 = GetSetting(App.Title, "Settings", "LANEnd", "05")
'ReScan Minuts
Text3 = GetSetting(App.Title, "Settings", "LANMinutes", "2")


End Sub

Private Sub Check1_Click()
Command2.Enabled = True
End Sub

Private Sub Check1_Validate(Cancel As Boolean)
Command2.Enabled = True
End Sub


Private Sub Check2_Click()
Command2.Enabled = True
End Sub

Private Sub Check2_Validate(Cancel As Boolean)
Command2.Enabled = True
End Sub


Private Sub Check3_Click()
Command2.Enabled = True
End Sub


Private Sub Check3_Validate(Cancel As Boolean)
Command2.Enabled = True
End Sub


Private Sub Check4_Click()
Command2.Enabled = True
End Sub

Private Sub Check4_Validate(Cancel As Boolean)
Command2.Enabled = True
End Sub


Private Sub Command1_Click()
Command2_Click
Hide
End Sub

Private Sub Command2_Click()
On Error Resume Next

'Confirmation on Exit
If Check1.Value = vbChecked Then
    SaveSetting App.Title, "Settings", "ExitConfirm", True
Else
    SaveSetting App.Title, "Settings", "ExitConfirm", False
End If

'Always On Top
If Check2.Value = vbChecked Then
    SaveSetting App.Title, "Settings", "AlwaysOnTop", True
Else
    SaveSetting App.Title, "Settings", "AlwaysOnTop", False
End If

'Reconnect WAN In One Minute
If Check3.Value = vbChecked Then
    SaveSetting App.Title, "Settings", "AutoReconnect", True
Else
    SaveSetting App.Title, "Settings", "AutoReconnect", False
End If

'Re-Scan LAN
If Check4.Value = vbChecked Then
    SaveSetting App.Title, "Settings", "LANScan", True
Else
    SaveSetting App.Title, "Settings", "LANScan", False
End If

'Starting LAN
SaveSetting App.Title, "Settings", "LANStart", Text1
'Ending Lan
SaveSetting App.Title, "Settings", "LANEnd", Text2
'ReScan Minuts
SaveSetting App.Title, "Settings", "LANMinutes", Text3

LoadUp

Command2.Enabled = False
End Sub

Private Sub Text1_Change()
Command2.Enabled = True
End Sub


Private Sub Text2_Change()
Command2.Enabled = True
End Sub


Private Sub Text3_Change()
Command2.Enabled = True
End Sub


