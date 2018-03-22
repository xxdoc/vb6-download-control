VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event DownloadErrors(strError As String)
Public Event DownloadEvents(strEvent As String)
Public Event DowloadComplete()
Public Event DownloadProgress(intPercent As Long, info As String)

Dim dl As New download_class

Private Sub UserControl_Resize()
    Width = 1020
    Height = 945
End Sub

Public Sub Cancel()
    dl.StopAll
End Sub

Public Function DownloadFile(strURL As String, strDestination As String) As Boolean
Set dl = New sasusk3_download_class
dl.url = strURL
RaiseEvent DownloadEvents("Conectando ao Servidor...")
dl.FLAG = 1
dl.OutputDir = strDestination
dl.Connect True
aa = dl.Download(, True)
Debug.Print strURL
'Form3.List1.AddItem dl.SC & strDestination
If aa = -1 Then DownloadFile = False: Exit Function
If aa = -2 Then DownloadFile = False: Exit Function

Do
DoEvents
RaiseEvent DownloadProgress(Val(dl.Percent), dl.BytesReceivedOF)
If dl.SC = True Then Exit Do
Loop
DownloadFile = True
RaiseEvent DowloadComplete

End Function

Function Porcentagem() As String
Porcentagem = dl.Percent
End Function



