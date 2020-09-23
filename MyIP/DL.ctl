VERSION 5.00
Begin VB.UserControl DL 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   465
   InvisibleAtRuntime=   -1  'True
   MaskColor       =   &H00C0C0C0&
   ScaleHeight     =   435
   ScaleWidth      =   465
End
Attribute VB_Name = "DL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Event Declarations:
Event Progress(Max As Long, Min As Long)
Event Complete(Data As Variant)


Public Sub Download(URL As String)
    On Error Resume Next
    AsyncRead URL, 1
End Sub


Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
Dim inData As String
On Error Resume Next
Open AsyncProp.Value For Binary As #1
inData = Space(FileLen(AsyncProp.Value))
Get #1, , inData
Close
RaiseEvent Complete(inData)
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
RaiseEvent Progress(AsyncProp.BytesMax, AsyncProp.BytesRead)
End Sub


