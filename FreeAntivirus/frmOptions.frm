VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realtime Scanner Options"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Global Options :"
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   5295
      Begin VB.CommandButton Command4 
         Caption         =   "Default"
         Height          =   255
         Left            =   4320
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Default"
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Text            =   $"frmOptions.frx":01CA
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Text            =   "EXE COM DLL DOC BAT PIF JS VBS ASP JAR SH PL PHP SQL WRI RTF HTML HTM"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "1"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skip folders :"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scan files type :"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mb"
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Scan files below :"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Realtime Scanner :"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   5295
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2520
         Top             =   600
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Disable"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enable"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Realtime Scanner Active"
         BeginProperty Font 
            Name            =   "Vrinda"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   270
         Left            =   3255
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Default Action For Realtime Scanner :"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   5295
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove Virus"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Quarentine"
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Do Nothing"
         Height          =   255
         Left            =   3720
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NONE"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last File Scanned :"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   1365
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Our Residental Scanner's Registry
Private Const PROJECT_KEY    As String = "FreeAV.Scanner\Clsid\"
'Enable Realtime
Private Sub Command1_Click()
    'If LenB(ReadReg(HKEY_CLASSES_ROOT, PROJECT_KEY, "", "")) = 0 Then
        Shell "regsvr32.exe" & " /s """ & App.Path & "\scanner.dll""", vbNormalNoFocus
   ' End If
   DoEvents
   
    A = ReadReg(HKEY_CLASSES_ROOT, PROJECT_KEY, vbNullString, vbNullString)
    If LenB(A) = 0 Then
        MsgBox "Error Registering Realtime Monitor!", vbCritical
        Exit Sub
    End If
    WriteRegString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks", A & "", ""
    A = ReadReg(HKEY_CLASSES_ROOT, PROJECT_KEY, vbNullString, vbNullString)
    If LenB(A) Then
        Label3.Caption = "Realtime Scanner Active"
        Exit Sub
    End If
    Form_Load
End Sub
'Disable Realtime
Private Sub Command2_Click()

    Shell "regsvr32.exe" & " /u /s """ & App.Path & "\scanner.dll""", vbNormalNoFocus
    DoEvents
    A = ReadReg(HKEY_CLASSES_ROOT, PROJECT_KEY, vbNullString, vbNullString)
    DeleteRegValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks", A & ""
    A = ReadReg(HKEY_CLASSES_ROOT, PROJECT_KEY, vbNullString, vbNullString)
    If LenB(A) = 0 Then
        Label3.Caption = "Disabled!"
        Exit Sub
    End If
    Form_Load
End Sub
'Default set
Private Sub Command3_Click()
    Text2.Text = "EXE COM DLL DOC BAT PIF JS VBS ASP JAR SH PL PHP SQL WRI RTF HTML HTM"
End Sub

Private Sub Command4_Click()
Text3.Text = "%WINDIR%\system32\drvstore\,%WINDIR%\system32\mui\,%WINDIR%\system32\drivers\,%WINDIR%\pchealth\,%WINDIR%\system32\dllcache\,%WINDIR%\winsxs\;%WINDIR%\Microsoft.Net"
End Sub

'Main Loadings
Private Sub Form_Load()
    A = ReadReg(HKEY_CLASSES_ROOT, PROJECT_KEY, vbNullString, vbNullString)
    If ReadReg(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks", A & vbNullString, "A") = "A" Then
        Label3.Caption = "Disabled!"
    Else
        Label3.Caption = "Realtime Scanner Active"
    End If
    A = GetSetting("FreeAV.Scanner", "RealtimeSettings", "Action", "2")
    If A = "1" Then
        Option1.Value = True
    End If
    If A = "2" Then
        Option2.Value = True
    End If
    If A = "3" Then
        Option3.Value = True
    End If
    Text2.Text = GetSetting("FreeAV.Scanner", "Settings", "Fileset", "EXE COM DLL DOC BAT PIF JS VBS ASP JAR SH PL PHP SQL WRI RTF HTML HTM")
    Text1.Text = GetSetting("FreeAV.Scanner", "Settings", "MaxFile", "1")
    Text3.Text = GetSetting("FreeAV.Scanner", "Settings", "Exclude", "%WINDIR%\system32\drvstore\,%WINDIR%\system32\mui\,%WINDIR%\system32\drivers\,%WINDIR%\pchealth\,%WINDIR%\system32\dllcache\,%WINDIR%\winsxs\;%WINDIR%\Microsoft.Net")
End Sub

Private Sub Option1_Click()
    Set1
End Sub
Private Sub Option2_Click()
    Set1
End Sub
Private Sub Option3_Click()
    Set1
End Sub
Private Sub Set1()
Dim v As Integer
    If Option1.Value Then
        v = 1
    End If
    If Option2.Value Then
        v = 2
    End If
    If Option3.Value Then
        v = 3
    End If
    SaveSetting "FreeAV.Scanner", "RealtimeSettings", "Action", v
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, _
                          Shift As Integer)
    SaveSetting "FreeAV.Scanner", "Settings", "MaxFile", CStr(Val(Text1.Text))
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, _
                          Shift As Integer)
    SaveSetting "FreeAV.Scanner", "Settings", "Fileset", Text2.Text
End Sub

Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SaveSetting "FreeAV.Scanner", "Settings", "Fileset", Text2.Text

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
    SaveSetting "FreeAV.Scanner", "Settings", "Exclude", Text3.Text

End Sub

Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SaveSetting "FreeAV.Scanner", "Settings", "Exclude", Text3.Text
End Sub

Private Sub Timer1_Timer()
    Label2.Caption = GetSetting("FreeAV.Scanner", "Settings", "LastFileScanned", "")
End Sub
':)Code Fixer V3.0.9 (5/12/2009 7:27:46 PM) 10 + 174 = 184 Lines Thanks Ulli for inspiration and lots of code.


