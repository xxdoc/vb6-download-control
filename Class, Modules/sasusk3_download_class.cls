VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "download_class"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Download Class using CSocketMaster


Public WithEvents wsHead As CSocketMaster
Attribute wsHead.VB_VarHelpID = -1
Public WithEvents ws As CSocketMaster
Attribute ws.VB_VarHelpID = -1

Dim bin As String

Private Type download_info
FileSize As Long
FileName As String
url As String
Server As String
uPath As String
End Type

Dim di As download_info, SD As Boolean, iPorcentagem As Long, iBytesReceived As Long, iBytesReceivedOF As String, iOutputDir As String
Dim v As Variant, s As String, uPort As Long, BRi As Long
Public HEAD As String, KBs As String, FLAG As Long, SC As Boolean
Dim iStartFrom As Long

Private Sub Class_Initialize()
Set wsHead = New CSocketMaster
Set ws = New CSocketMaster
End Sub

Private Sub Class_Terminate()
IO.FileClose
Set wsHead = Nothing
Set ws = Nothing
End Sub

Property Get StartFrom() As Long
StartFrom = iStartFrom
End Property

Property Let StartFrom(i As Long)
iStartFrom = i
End Property

Private Sub ws_Connect()
gg = "GET " & di.uPath & " HTTP/1.1 \n"
gg = gg & "Accept: */*\n"
gg = gg & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1) ; FDM; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022)\n"
If iStartFrom > 0 Then gg = gg & "Range: bytes=" & iStartFrom & "-" & vbCrLf
gg = gg & "Host: " & di.Server & "\n"
gg = gg & "Connection: Keep-Alive\n\n"
gg = Replace(gg, "\n", vbCrLf)
ws.SendData gg
End Sub

Function DownloadState() As String
DownloadState = ws.State
End Function

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim Buff As String

ws.GetData Buff
'bin = bin & buff

If SD = True Then
If InStr(1, Buff, vbCrLf & vbCrLf) Then
v = Split(Buff, vbCrLf & vbCrLf, 2)
s = v(UBound(v))
Buff = s
SD = False
GoTo DADS
End If
Else
DADS:
IO.FileWrite Buff

'If LOF(1) > 1 Then Seek #FReE, LOF(1)

'Form3.List1.AddItem iBytesReceived & " )))" & di.FileSize
'Me.Caption = "SALVANDO SAIDA!!! " & buff
'Close #1

iBytesReceived = iBytesReceived + Len(Buff) + iStartFrom
If iBytesReceived = 0 Then iBytesReceived = 1
iPercent = Format(iBytesReceived / di.FileSize * 100)
curPorcentagem = iPercent
KBs = iBytesReceived - BRi
BRi = iBytesReceived

iBytesReceivedOF = Format(iBytesReceived, "###,###,##0") & " bytes de " & Format(di.FileSize, "###,###,##0") & " bytes"
curCopiadoERestante = iBytesReceivedOF

If iBytesReceived >= di.FileSize Then IO.FileClose: SC = True: Exit Sub
End If
End Sub




Private Sub wsHead_Connect()
'GET /bartman2.gif HTTP/1.1
'Accept: */*
'Referer: http://192.168.254.3:81/
'Accept -Language: pt -br
'Accept -Encoding: gzip , deflate
'User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1) ; FDM; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022)
'Host: 192.168.254.3:81
'Connection: Keep -Alive
gg = "HEAD " & di.uPath & " HTTP/1.1 \n"
gg = gg & "Accept: */*\n"
gg = gg & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1) ; FDM; .NET CLR 2.0.50727; .NET CLR 3.0.04506.648; .NET CLR 3.5.21022)\n"
If iStartFrom > 0 Then gg = gg & "Range: bytes=" & iStartFrom & "-" & vbCrLf
gg = gg & "Host: " & di.Server & "\n"
gg = gg & "Connection: Keep-Alive\n\n"
gg = Replace(gg, "\n", vbCrLf)

wsHead.SendData gg
'bin = ""
End Sub

Private Sub wsHead_DataArrival(ByVal bytesTotal As Long)
Dim Buff As String
wsHead.GetData Buff
bin = bin & Buff
If InStr(1, Buff, vbCrLf & vbCrLf) Then
v = Split(bin, vbCrLf & vbCrLf)
HEAD = v(LBound(v))
v = Split(HEAD, vbCrLf)
If InStr(1, v(0), "404") Then
HEAD = "-1"
wsHead.CloseSck
Exit Sub
End If

v = Split(HEAD, "Content-Length:", 2)
s = v(UBound(v))
v = Split(s, vbCrLf, 2)
s = v(LBound(v))
'HTTP/1.1 404 Not Found
di.FileSize = s
If iStartFrom > 0 Then If InStr(1, HEAD, "Accept-Ranges: bytes", vbTextCompare) Then IO.Append = True

wsHead.CloseSck
End If
End Sub

Property Get url() As String
url = di.url
End Property

Property Get FileName() As String
FileName = di.FileName
End Property

Property Get FileSize() As String
FileSize = di.FileSize
End Property

Property Get Server() As String
Server = di.Server
End Property

Property Get uPath() As String
uPath = di.uPath
End Property

Property Get BytesReceived() As String
BytesReceived = iBytesReceived
End Property

Property Get Percent() As String
Percent = iPercent
End Property

Property Get BytesReceivedOF() As String
BytesReceivedOF = iBytesReceivedOF
End Property

Property Get OutputDir() As String
OutputDir = iOutputDir
End Property


Property Let OutputDir(s As String)
 iOutputDir = s
End Property

Property Let url(uri As String)

v = Split(uri, "://", 2)
s = v(UBound(v))
v = Split(s, "/", 2)
s = v(LBound(v))
di.Server = s  'url server

v = Split(uri, "/")
s = v(UBound(v))
di.FileName = s

v = Split(uri, "://", 2)
s = v(UBound(v))
v = Split(s, "/", 2)
s = v(UBound(v))
di.uPath = "/" & s

di.url = uri
End Property

Property Let Port(i As Long)
uPort = i
End Property

Function Connect(Optional WaitConnect As Boolean = False)
If uPort = 0 Then uPort = 80
wsHead.CloseSck
ws.CloseSck


iPorcentagem = 0
iBytesReceived = 0
iBytesReceived = 0

wsHead.Connect di.Server, uPort
If WaitConnect = True Then
Do
DoEvents
'If wsHead.State = sckError Then Exit Do
If Len(HEAD) > 0 Then Exit Do
Loop
End If
End Function

Function IsConnected() As Boolean
If wsHead.State = sckError Then
IsConnected = False
Else
IsConnected = True
End If
End Function

Function Download(Optional WaitDownload As Boolean = False, Optional KillIfExist As Boolean = False) As Long
On Error Resume Next
If HEAD = "-1" Then wsHead.CloseSck: ws.CloseSck: Download = -1: Exit Function
If uPort = 0 Then uPort = 80
curPorcentagem = 0
ws.CloseSck
SD = True
SC = False

If FLAG = 1 Then
If KillIfExist = True Then Kill iOutputDir
IO.File = iOutputDir
Else
'v = Split(iOutputDir, "\")
's = v(UBound(v))
'k = Mid(iOutputDir, 1, Len(iOutputDir) - Len(s))
'iOutputDir = k
If KillIfExist = True Then Kill certo(iOutputDir) & di.FileName
IO.File = certo(iOutputDir) & di.FileName
End If

If IO.OpenFile = -1 Then Download = "-2": Exit Function 'não foi possivel criar o arquivo...

ws.Connect di.Server, uPort

If WaitDownload = True Then
Do
DoEvents
If SC = True Then Exit Do
Loop
End If
End Function

Private Function certo(s As String) As String
If Right(s, 1) = "\" Then
certo = s
Else
certo = s & "\"
End If
End Function

Private Sub wsHead_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
HEAD = "-1"
End Sub

Function StopAll()
HEAD = "-1"
ws.CloseSck
wsHead.CloseSck
IO.FileClose
End Function

