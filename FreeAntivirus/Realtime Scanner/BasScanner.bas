Attribute VB_Name = "basScanner"
Option Explicit
Public Const E_NOTIMPL = &H80004001
Public Const PAGE_EXECUTE_READWRITE = &H40&
Public Const S_FALSE = 1
Public Const S_OK = 0

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function VirtualProtect Lib "Kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long

Private Declare Function lstrlenA Lib "Kernel32" (ByVal lpString As Long) As Long
Private Declare Function lstrlenW Lib "Kernel32" (ByVal lpString As Long) As Long

Private Declare Function lstrcpyA Lib "Kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function lstrcpyW Lib "Kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

'System Calls This When Any File Executes
Public Function Execute(this As Long, pei As olelib.SHELLEXECUTEINFO) As HRESULTS
    'Dim strInfo As String

    'strInfo = strInfo + "cbSize" + Str(pei.cbSize) + vbCrLf
    'strInfo = strInfo + "fMask" + Str(pei.fMask) + vbCrLf
    'strInfo = strInfo + "hInstApp" + Str(pei.hInstApp) + vbCrLf
    'strInfo = strInfo + "hwnd" + Str(pei.hWnd) + vbCrLf
    'strInfo = strInfo + "lpDirectory" + StrFromPtr(pei.lpDirectory, True) + vbCrLf
    'strInfo = strInfo + "lpFile" + StrFromPtr(pei.lpFile, True) + vbCrLf
    'strInfo = strInfo + "lpParameters" + StrFromPtr(pei.lpParameters, True) + vbCrLf
    'strInfo = strInfo + "lpVerb" + StrFromPtr(pei.lpVerb, True) + vbCrLf
    'strInfo = strInfo + "nShow" + Str(pei.nShow) + vbCrLf

    'MsgBox StrFromPtr(pei.lpFile, True) & " '" & StrFromPtr(pei.lpParameters, True) & "'"
LoadRecordset
If Dir(DBF_NAME) = "" Then
Shell "Regsvr32.exe " & " /u /s " & App.Path & "\" & App.EXEName & ".dll"
Exit Function
End If
'MsgBox StrFromPtr(pei.lpFile, True)
'If File has Parameters then scan Parameters
If LenB(StrFromPtr(pei.lpParameters, True)) > 0 Then
If ScanFile(StrFromPtr(pei.lpFile, True)) = False And ScanFile(StrFromPtr(pei.lpParameters, True)) = False Then
Execute = 1
Else
Execute = 0
End If
' Scan file if file present
ElseIf LenB(StrFromPtr(pei.lpFile, True)) > 0 Then
If ScanFile(StrFromPtr(pei.lpFile, True)) = False Then
Execute = 1
Else
Execute = 0
End If
Else
Execute = 1
End If
End Function
'Find Exect File and Scan it
Function ScanFile(F As String) As Boolean
'On Error Resume Next
Dim F1 As String
'File has path
If InStr(1, F, "\") > 0 Then
ScanFile = ScanForVirus(F)
Else 'no path
F1 = GetFileFromName(F)
If F1 <> "" Then
ScanFile = ScanForVirus(F1)
End If
End If
Exit Function
err:
MsgBox "SF:" & err.Description & vbCrLf & F
err.Clear
End Function
'Scans for virus and action
Private Function ScanForVirus(File As String) As Boolean
Dim L As Long
'On Error Resume Next
SaveSetting "FreeAV.Scanner", "Settings", "LastFileScanned", File
Dim A As String, H As String
If MatchExcludes(GetPath(File)) = True Then
If MatchFileSet(File) = True Then
H = HashFile(File, MD2)
If Len(H) > 0 Then
A = Match(H)
If Len(A) > 0 Then
ScanForVirus = True
If isProcess(Right(File, (Len(File) - InStrRev(File, "\")))) = True Then KillProcess Right(File, (Len(File) - InStrRev(File, "\")))
A = GetSetting("FreeAV.Scanner", "RealtimeSettings", "Action", "2")
On Error Resume Next
If A = "1" Then Kill File
If A = "2" Then Name File As File & ".quad"
If A = "3" Then SaveSetting "FreeAV.Scanner", "Logs", "LastVirusSkipped", File
err.Clear
On Error GoTo err
End If
End If
End If
End If
Exit Function
err:
err.Clear
End Function

Public Function SwapVtableEntry(pObj As Long, EntryNumber As Integer, ByVal lpfn As Long) As Long

    Dim lOldAddr As Long
    Dim lpVtableHead As Long
    Dim lpfnAddr As Long
    Dim lOldProtect As Long

    CopyMemory lpVtableHead, ByVal pObj, 4
    lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
    CopyMemory lOldAddr, ByVal lpfnAddr, 4

    Call VirtualProtect(lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
    CopyMemory ByVal lpfnAddr, lpfn, 4
    Call VirtualProtect(lpfnAddr, 4, lOldProtect, lOldProtect)

    SwapVtableEntry = lOldAddr

End Function

Public Function StrFromPtr(ByVal lpString As Long, Optional fUnicode As Boolean = False) As String
    On Error Resume Next
    If fUnicode Then
        StrFromPtr = String(lstrlenW(lpString), Chr(0))
        lstrcpyW StrPtr(StrFromPtr), ByVal lpString
    Else
        StrFromPtr = String(lstrlenA(lpString), Chr(0))
        lstrcpyA ByVal StrFromPtr, ByVal lpString
    End If
End Function
