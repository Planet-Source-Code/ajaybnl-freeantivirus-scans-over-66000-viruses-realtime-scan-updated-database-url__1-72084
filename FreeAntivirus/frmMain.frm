VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   ClientHeight    =   5115
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":33E2
   ScaleHeight     =   5115
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   3960
      Width           =   210
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   3960
      Width           =   225
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFF80&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   6000
      Picture         =   "frmMain.frx":8965
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   5760
      Picture         =   "frmMain.frx":BD47
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1440
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "..."
      Height          =   255
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   630
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "c:\"
      Top             =   615
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6540
      TabIndex        =   11
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   5085
      TabIndex        =   10
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
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
      Left            =   5490
      TabIndex        =   9
      Top             =   4260
      Width           =   675
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   5
      Top             =   4635
      Width           =   5055
   End
   Begin VB.Menu mnu 
      Caption         =   "Menu"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu manualscan 
         Caption         =   "Manual Scanner"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu rto 
         Caption         =   "Options"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu extgui 
         Caption         =   "Exit Gui"
      End
      Begin VB.Menu exitt 
         Caption         =   "Exit & Unload Realtime"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private LoadedViaTray   As Boolean
Private BaloonTimer     As Long
Private VirusFound      As Boolean
Private I               As Long, Tmr As Long

'Log
Private Sub Addtext(ByVal T As String)
On Error GoTo err
    Text2.Text = Text2.Text & IIf(Len(Text2.Text) = 0, "", vbNewLine) & T
    Text2.SelStart = Len(Text2.Text)
    'Text2.SelLength = 0
DoEvents

Exit Sub
err:
err.Clear
Text2.Text = ""
End Sub
Private Sub Command1_Click()
    Text1.Text = BrowseFolder(0, "Select Location")
End Sub

'Start Stop
Private Sub Command2_Click()
On Error Resume Next
Dim I As Long
    If Command2.Caption = "Start" Then
        Command2.Caption = "Stop"
        Text2.Text = vbNullString
        Addtext "Scanning " & Text1.Text
        'Addtext String$(50, "#")
         Text1.Enabled = False
        Command1.Enabled = False
        FStart = True
        If Dir(Text1.Text, vbDirectory) = "" Then
        Addtext "Invalid Path!"
        Else
        FindFiles Text1.Text
        End If
        FStart = False
        Addtext ""
        Addtext "Total Infected Files Found : " & I
        'Addtext String$(50, "#")
       Text1.Enabled = True
        Command1.Enabled = True
        
        Command2.Caption = "Start"
    Else
    Text1.Enabled = True
        Command1.Enabled = True
        FStart = False
        Command2.Caption = "Start"
        Addtext ""
        Addtext "Process Stopped By User"
        'Addtext String$(50, "#")
    End If
    Label10.Caption = vbNullString
    err.Clear
End Sub

'Complete Exit
Private Sub exitt_Click()
    LoadedViaTray = False
    Timer1.Enabled = False
    DoEvents
    frmOptions.Hide
    DoEvents
    frmOptions.Command2.Value = True
    DoEvents
    DoEvents
    Unload frmOptions
    Unload frmUpdate
    Unload Me
End Sub

'ExitApp
Private Sub extgui_Click()
    'Unload frmOptions
RemoveFromTray
   End
End Sub

'Main
Private Sub Form_Load()
Dim F As String


'Actions Loading
    A = GetSetting(App.Title, "Settings", "Action", "2")
    If A = "1" Then
        Option1.Value = True
    ElseIf A = "2" Then
        Option2.Value = True
    ElseIf A = "3" Then
        Option3.Value = True
    End If
    
    
    'Load DB
    LoadRecordset
    If LenB(Dir(DBF_NAME)) = 0 Then
    If MsgBox("The Definitions file is not present! Do you want to update it?" & vbCrLf & "It will take approx '8 MB'", vbOKCancel, "Definitions Download") = vbOK Then
    frmUpdate.Show
    Me.Visible = False
    Exit Sub
    End If
    
    '    CreateDatabase DBF_NAME, dbLangGeneral, dbEncrypt
    '    Createtables
    End If

    'Count Virus Sig
    Label10.Caption = "Total Virus Signatures in Database : " & CountRecords()
    
    'detect command line
    If LenB(Command$()) Then
    ' startup mode
        If InStr(1, Command$(), "/r", vbTextCompare) > 0 Then
Start1:
            AddToTray Me, mnu, , , , Picture1.Picture
            frmOptions.Hide
            frmOptions.Command1.Value = True
            Unload frmOptions
            Me.Visible = False
            Timer1.Enabled = True
        Else
'if command to scan file
            F = Command$()
            ScanFile F
            Unload Me
            End
        End If
'no command specified
    Else
        If LenB(ReadReg(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "FreeAntivirus", "")) = 0 Then
            If MsgBox("FreeAntivirus is not set to startup! You can enable realtime protection only when it is set to start at boot!" & vbNewLine & _
          vbNewLine & _
          "Do you want to set it to start at boot?", vbOKCancel) = vbOK Then
                WriteRegString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "FreeAntivirus", App.Path & "\" & App.EXEName & ".exe /r"
                GoTo Start1
            End If
        End If
        LoadedViaTray = True
        AddToTray Me, mnu, , , , Picture1.Picture
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    
    If LoadedViaTray Then
        Me.Visible = False
        Cancel = 1
    Else
        RemoveFromTray
        Unload Me
    End If
    
End Sub

Private Sub Label1_Click()
Command2_Click
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub manualscan_Click()
    Me.Visible = True
    LoadedViaTray = True
End Sub
Private Sub Option1_Click()
    SaveSetting App.Title, "Settings", "Action", "1"
End Sub
Private Sub Option2_Click()
    SaveSetting App.Title, "Settings", "Action", "2"
End Sub
Private Sub Option3_Click()
    SaveSetting App.Title, "Settings", "Action", "3"
End Sub
'Process Each File (Called from finding files)
Public Sub Process(File As String, _
                   Path As String)
                   'debug.print "Processing File " & Path & File
'AddData File, Hash.HashFile(Path & File, MD2)
Dim A As String
Dim H As String
    Label10.Caption = "Scanning " & Path
'calculate hash md2

        H = HashFile(Path & File, MD2)
        If LenB(H) > 0 Then
            A = Match(H)
'virus found
            If LenB(A) > 0 Then
            'Play sound on first virus found
                If Not VirusFound Then
                    VirusFound = True
                    PLaySound App.Path & "\detected.wav"
                End If
                
                'If Virus Found in Memory then kill it
                If isProcess(File) = True Then KillProcess File
                
                If Option1.Value Then
                    Kill Path & File
                    Addtext "Infected File : " & File & " is a " & A & " - Deleted"
                ElseIf Option2.Value Then
                    Name Path & File As Path & File & ".quad"
                    Addtext "Infected File : " & File & " is a " & A & " - Quarentined"
                ElseIf Option3.Value Then
                    Addtext "Infected File : " & File & " is a " & A & " - Skipped"
                End If
                I = I + 1
            End If
            Else ' Hasfile got zero
            Addtext "Locked File : " & File & " - Skipped"
        End If
End Sub
Private Sub rto_Click()
    frmOptions.Visible = True
End Sub
'Find Exect File and Scan it
Private Sub ScanFile(F As String)
'On Error GoTo err
''Dim WD As String
Dim F1 As String
'WD = Environ$("WINDIR") & "\"
'File has path
    If InStr(1, F, "\") > 0 Then
        Process Replace$(F, GetPath(F), ""), GetPath(F)
    Else 'no path
        F1 = GetFileFromName(F)
        If LenB(F1) Then
            Process Replace$(F1, GetPath(F1), ""), GetPath(F1)
        End If
    End If
End Sub



'show green icon if realtime module is loaded else red
Private Sub Timer1_Timer()
    A = ReadReg(HKEY_CLASSES_ROOT, PROJECT_KEY, vbNullString, vbNullString)
    If ReadReg(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks", A & vbNullString, "A") = "A" Then
'if explorer is terminated it will reload icon in tray
    Tmr = Tmr + 1
    If Tmr > 5 Or Tmr = 1 Then
        RemoveFromTray
        AddToTray Me, mnu, , , , Picture2.Picture
        Tmr = 2
        End If
        
        If BaloonTimer = 0 Or BaloonTimer > 120 Then
            AddToTrayToolTip Me, mnu, "Realtime protection is not loaded!" & vbNewLine & _
             "You are not protected from viruses!!!" & vbNewLine & _
             "Please load it from options.", "Warning", 2, Picture2.Picture
            BaloonTimer = 1
        End If
        BaloonTimer = BaloonTimer + 1
    Else
'if explorer is terminated it will reload icon in tray
        
       Tmr = Tmr + 1
    If Tmr > 5 Or Tmr = 1 Then
        RemoveFromTray
        AddToTray Me, mnu, , , , Picture1.Picture
        Tmr = 2
        End If
        
'Show that you are skipping virus action!
        If LenB(GetSetting("FreeAV.Scanner", "Logs", "LastVirusSkipped", "")) Then
         If Not VirusFound Then
                    VirusFound = True
                    PLaySound App.Path & "\detected.wav"
        End If
            AddToTrayToolTip Me, mnu, "Virus found and stopped but cannot take action due to settings!" & vbNewLine & _
             GetSetting("FreeAV.Scanner", "Logs", "LastVirusSkipped", ""), "Virus Found!", 5, Picture1.Picture
            SaveSetting "FreeAV.Scanner", "Logs", "LastVirusSkipped", ""
        End If
    End If
End Sub

''
':)Code Fixer V3.0.9 (5/12/2009 7:27:44 PM) 17 + 294 = 311 Lines Thanks Ulli for inspiration and lots of code.


