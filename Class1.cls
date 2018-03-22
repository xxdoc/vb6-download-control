VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Classe de acesso a arquivo

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Const GENERIC_WRITE = &H40000000
Private Const GENERIC_READ = &H80000000
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_APPEND_DATA = (&H4)
Private Const FILE_END = 2

Dim iFile As String, iAppend As Boolean
Public handle As Long

Property Let File(s As String)
iFile = s
End Property

Property Get File() As String
File = iFile
End Property

Property Let Append(s As Boolean)
iAppend = s
End Property

Property Get Append() As Boolean
Append = iAppend
End Property

Function OpenFile() As Long
If iFile = "" Then OpenFile = -1: Exit Function

If iAppend = True Then
handle = CreateFile(iFile, GENERIC_WRITE, FILE_SHARE_WRITE, ByVal 0&, 4, 0, 0)
SetFilePointer handle, 0, 0, 2 'seek 2 end
Else
handle = CreateFile(iFile, GENERIC_WRITE, FILE_SHARE_WRITE, ByVal 0&, 2, 0, 0)
End If

OpenFile = handle
End Function

Function FileWrite(Buff As String) As Long
Dim numwritten As Long
numwritten = 5
'If Handle = 0 Then FileWrite = -1
Call WriteFile(handle, ByVal Buff, ByVal Len(Buff), numwritten, ByVal 0&)
End Function

Function FileClose()
CloseHandle handle
handle = 0
End Function




