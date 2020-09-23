VERSION 5.00
Begin VB.Form frmExit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shutdown MyIP Personal"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4500
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
   Icon            =   "frmExit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "&Do Not Ask again before Shutting Down"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Are you sure you want to close MyIP Personal? Your computer will no longer display IP information."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Check1.Value = True Then
'Information Dialog about undoing the effect
MsgBox "This dialog will not be shown again. If you want MyIP to ask you before closing you will need to go into the settings menu", vbInformation
'Save setting to registry
SaveSetting App.Title, "Settings", "ExitConfirm", False
End If

'Shutdown the application
End
End Sub

Private Sub Command2_Click()
Hide
frmMain.blnUnload = False
End Sub
