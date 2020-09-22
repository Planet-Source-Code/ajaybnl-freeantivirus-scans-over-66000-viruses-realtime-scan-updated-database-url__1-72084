Attribute VB_Name = "basTrayCode"
'This sample was downloaded from http://www.nekhbet.tk
'Made by Trambitas Sorin @ 19.01.2005
'For questions please contact me at TrimbitasSorin@Yahoo.com
'A small part of this code was taken from a sample
'found at http://www.vb-helper.com
Option Explicit
Private OldWindowProc          As Long
Private TheForm                As Form
Private TheMenu                As Menu
Private TheData                As NOTIFYICONDATA
''Private Const WM_SYSCOMMAND    As Long = &H112
''Private Const SC_MOVE          As Long = &HF010
''Private Const SC_RESTORE       As Long = &HF120
''Private Const SC_SIZE          As Long = &HF000
Private Const WM_USER          As Long = &H400
''Private Const WM_LBUTTONUP     As Long = &H202
''Private Const WM_MBUTTONUP     As Long = &H208
Private Const WM_RBUTTONUP     As Long = &H205
Private Const TRAY_CALLBACK    As Double = (WM_USER + 1001&)
Private Const GWL_WNDPROC      As Long = (-4)
''Private Const GWL_USERDATA     As Long = (-21)
Private Const NIF_ICON         As Long = &H2
''Private Const NIF_TIP          As Long = &H4
Private Const NIM_ADD          As Long = &H0
Private Const NIF_INFO         As Long = &H10
''Private Const NIIF_INFO        As Long = &H1
Private Const NIF_MESSAGE      As Long = &H1
Private Const NIM_MODIFY       As Long = &H1
Private Const NIM_DELETE       As Long = &H2
Private Const WM_NULL          As Long = &H0
Private Const WM_MOUSEMOVE     As Long = &H200
Private Type NOTIFYICONDATA
    cbSize                         As Long
    hwnd                           As Long
    uID                            As Long
    uFlags                         As Long
    uCallbackMessage               As Long
    hIcon                          As Long
    szTip                          As String * 128
    dwState                        As Long
    dwStateMask                    As Long
    szInfo                         As String * 256
    uTimeout                       As Long
    szInfoTitle                    As String * 64
    dwInfoFlags                    As Long
End Type
Private IsTrayIconActive       As Boolean
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal Msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                                                                       lpData As NOTIFYICONDATA) As Long
''Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long
Private Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

'Add the form's icon to the tray.
Public Sub AddToTray(frm As Form, _
                     mn_Mnu As Menu, _
                     Optional awa As String = "none", _
                     Optional aqa As String = "none", _
                     Optional TQS As Byte = 1, _
                     Optional stdPicture1 As StdPicture)
    If stdPicture1 Is Nothing Then
        Set stdPicture1 = frm.Icon
    End If
    If Not IsTrayIconActive Then
        Set TheForm = frm
        Set TheMenu = mn_Mnu
        OldWindowProc = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf NewWindowProc)
        With TheData
            If (aqa = "none") And (awa = "none") Then
                .uID = 0
            Else
                .uID = vbNull
            End If
            .hwnd = frm.hwnd
            .cbSize = Len(TheData)
            .hIcon = stdPicture1.Handle
            .uFlags = NIF_ICON Or NIF_MESSAGE
            .uCallbackMessage = TRAY_CALLBACK
            .cbSize = Len(TheData)
        End With
        Shell_NotifyIcon NIM_ADD, TheData
        If awa <> "none" Then
            If aqa <> "none" Then
                ShowPopUp awa, aqa
                Sleep TQS * 1000
            End If
        End If
        IsTrayIconActive = True
    End If
End Sub
'Show a tooltip attached at the icon from tray
Public Sub AddToTrayToolTip(formName As Form, _
                            menuName As Menu, _
                            TipMsg As String, _
                            TipTitle As String, _
                            Optional TipTimeOutInSeconds As Byte = 1, _
                            Optional stdPicture1 As StdPicture)
    If stdPicture1 Is Nothing Then
        Set stdPicture1 = formName.Icon
    End If
    If IsTrayIconActive Then
        RemoveFromTray
        AddToTray formName, menuName, TipMsg, TipTitle, TipTimeOutInSeconds, stdPicture1
        RemoveFromTray
        AddToTray formName, menuName, , , , stdPicture1
    End If
End Sub
'The replacement window process
Private Function NewWindowProc(ByVal lngHwnd As Long, _
                               ByVal Msg As Long, _
                               ByVal wParam As Long, _
                               ByVal lParam As Long) As Long
'debug.print Msg
Const WM_NCDESTROY As Long = &H82
    If Msg = WM_NCDESTROY Then
        RemoveFromTray
    Else
        If Msg = TRAY_CALLBACK Then
            If lParam = WM_RBUTTONUP Then
                SetForegroundWindow TheForm.hwnd
                TheForm.PopupMenu TheMenu
                If Not (TheForm Is Nothing) Then
                    PostMessage TheForm.hwnd, WM_NULL, ByVal 0&, ByVal 0&
                End If
                Exit Function
            End If
        End If
    End If
    NewWindowProc = CallWindowProc(OldWindowProc, lngHwnd, Msg, wParam, lParam)
End Function
'Remove the icon from the system tray.
Public Sub RemoveFromTray()
    If IsTrayIconActive Then
        With TheData
        .uFlags = 0
        End With
        Shell_NotifyIcon NIM_DELETE, TheData
        SetWindowLong TheForm.hwnd, GWL_WNDPROC, OldWindowProc
        Set TheForm = Nothing
        IsTrayIconActive = False
        End If
End Sub
'Show Tooltip
Private Sub ShowPopUp(ByVal Message As String, _
                      ByVal Title As String)
    With TheData
        .cbSize = Len(TheData)
        .hwnd = frmMain.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = TheForm.Icon
        .szTip = Title & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Message & vbNullChar
        .szInfoTitle = Title & vbNullChar
        .dwInfoFlags = NIF_INFO
    End With
    Shell_NotifyIcon NIM_MODIFY, TheData
End Sub
''
''Public Sub InitializeTrayModule()
''
''
''
''IsTrayIconActive = False
''End Function
''
''
'''Set a tray tip.
''Public Sub SetTrayTip(ByVal tip As String)
''
''
''
''If IsTrayIconActive Then
''
''With TheData
''.szTip = tip & vbNullChar
''.uFlags = NIF_TIP
''End With
''Shell_NotifyIcon NIM_MODIFY, TheData
''End If
''End Sub
''
':)Code Fixer V3.0.9 (5/12/2009 7:27:47 PM) 85 + 158 = 243 Lines Thanks Ulli for inspiration and lots of code.


