VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definitions Update"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   480
   End
   Begin FreeAntivirus.FileDownloader FD1 
      Left            =   3480
      Top             =   1200
      _ExtentX        =   1799
      _ExtentY        =   1667
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   465
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Update status :"
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1080
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FD1_DowloadComplete()
MsgBox "Download Completed Sucessfully"
Timer1.Enabled = False
frmMain.Enabled = True
frmOptions.Enabled = True
Unload Me
End Sub

Private Sub FD1_DownloadErrors(strError As String)
If MsgBox("Error Downloading Definitions : " & strError & vbCrLf & vbCrLf & "Do You Want To Try Again?", vbOKCancel) = vbOK Then
Download
Else
End
End If

End Sub

Private Sub FD1_DownloadEvents(strEvent As String)
Label2.Caption = strEvent

End Sub

Private Sub FD1_DownloadProgress(intPercent As String)
Label3.Caption = intPercent & "%"
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub
Sub Download()
Call FD1.DownloadFile("http://ajaybnl.x10.mx/av.mdb", App.Path & "\av.mdb", "", "")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
FD1.Cancel
frmMain.Enabled = True
frmMain.Visible = True
frmOptions.Enabled = True

End Sub

Private Sub Timer1_Timer()
Me.Visible = True
Download
frmMain.Enabled = False
frmOptions.Enabled = False
Timer1.Enabled = False
End Sub
