VERSION 5.00
Object = "*\A..\Projekt2.vbp"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download Parcial - Teste"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   81
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar>>"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin Project2.download_control download_control 
      Left            =   2880
      Top             =   120
      _extentx        =   847
      _extenty        =   900
      waitdownloadfinish=   0   'False
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim arquivoLocal As String, arquivoURL As String


arquivoLocal = "c:\ldm.jpg"
arquivoURL = "http://goliveira.com/ldm.jpg"


If Dir(arquivoLocal) <> "" Then

If MsgBox("Deseja resumir o arquivo?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
Dim StartFrom As Long
StartFrom = FileLen(arquivoLocal)

download_control.DownloadFile arquivoURL, arquivoLocal, StartFrom
Else
GoTo doInicio:
End If

Else
doInicio:
download_control.DownloadFile arquivoURL, arquivoLocal
End If

Me.Caption = Me.Caption & " - Baixando..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
download_control.Cancel
End Sub

Private Sub download_control_DowloadComplete()
MsgBox "Download Concluído!", vbInformation + vbOKOnly, Me.Caption
End Sub

Private Sub download_control_DownloadErrors(strError As String)
MsgBox strError
End Sub

Private Sub download_control_DownloadEvents(strEvent As String)
Label2.Caption = strEvent
End Sub

Private Sub download_control_DownloadProgress(intPercent As Long, info As String)
pb.Value = intPercent
Label1.Caption = info
End Sub
