VERSION 5.00
Begin VB.UserControl download_control 
   Alignable       =   -1  'True
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "download_control.ctx":0000
   ScaleHeight     =   510
   ScaleWidth      =   480
   ToolboxBitmap   =   "download_control.ctx":08CA
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   120
      Top             =   360
   End
End
Attribute VB_Name = "download_control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event DownloadErrors(strError As String)
Public Event DownloadEvents(strEvent As String)
Public Event DowloadComplete()
Public Event DownloadProgress(intPercent As Long, info As String)

Dim dl As New download_class, iWaitDownloadFinish As Boolean
Dim iParar As Boolean
Private Sub Timer1_Timer()
On Error Resume Next

If iParar = True Then
RaiseEvent DownloadEvents("Download parado")
Timer1.Enabled = False
Exit Sub
End If

If dl.BytesReceived > j Then
RaiseEvent DownloadProgress(Format((dl.BytesReceived / dl.FileSize) * 100, "#"), dl.BytesReceivedOF)
j = dl.BytesReceived
End If

If dl.SC = True Then
Timer1.Enabled = False
RaiseEvent DowloadComplete
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
iWaitDownloadFinish = PropBag.ReadProperty("WaitDownloadFinish", True)
End Sub

Private Sub UserControl_Resize()
    Width = 480
    Height = 510
End Sub

Public Sub Cancel()
iParar = True
    dl.StopAll
End Sub

Public Function DownloadFile(strURL As String, strDestination As String, Optional StartFrom As Long = 0) As Boolean
Dim KillIfExist As Boolean

If StartFrom > 0 Then KillIfExist = True Else KillIfExist = False

Set dl = New download_class
dl.url = strURL
RaiseEvent DownloadEvents("Conectando ao Servidor...")
dl.FLAG = 1
dl.StartFrom = StartFrom
dl.OutputDir = strDestination
dl.Connect True
aa = dl.Download(, KillIfExist)
If Val(dl.HEAD) = -1 Then
RaiseEvent DownloadErrors("Erro ao tentar se conectar com o servidor.")
DownloadFile = False
RaiseEvent DownloadEvents("")
Exit Function
End If
RaiseEvent DownloadEvents("")

Debug.Print strURL

If aa = -1 Then DownloadFile = False: Exit Function
If aa = -2 Then DownloadFile = False: Exit Function
On Error Resume Next
Dim j As Integer

If iWaitDownloadFinish = True Then
Do
If iParar = True Then Exit Do
DoEvents
If dl.BytesReceived > j Then
RaiseEvent DownloadProgress(Format((dl.BytesReceived / dl.FileSize) * 100, "#"), dl.BytesReceivedOF)
j = dl.BytesReceived
End If

If dl.SC = True Then Exit Do
Loop

DownloadFile = True
RaiseEvent DowloadComplete

Else
Timer1.Enabled = True
End If

End Function

Function Percent() As Variant
On Error Resume Next
Percent = Format((dl.BytesReceived / dl.FileSize) * 100, "#")
End Function

Property Let WaitDownloadFinish(b As Boolean)
iWaitDownloadFinish = b
End Property

Property Get WaitDownloadFinish() As Boolean
WaitDownloadFinish = iWaitDownloadFinish
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "WaitDownloadFinish", iWaitDownloadFinish, True
End Sub
