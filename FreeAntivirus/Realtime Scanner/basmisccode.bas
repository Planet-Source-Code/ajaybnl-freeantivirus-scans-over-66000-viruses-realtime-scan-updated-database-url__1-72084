Attribute VB_Name = "basMiscCode"
Option Explicit
Private Const SND_ASYNC   As Long = &H1    '  play asynchronously
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                                                             ByVal uFlags As Long) As Long
'Extended Dir
Private Function Dirf(F As String) As String
    On Error Resume Next
    If LenB(Dir(F)) Then
        Dirf = Dir(F)
    ElseIf LenB(Dir(F, vbHidden)) Then
        Dirf = Dir(F, vbHidden)
    ElseIf LenB(Dir(F, vbSystem)) Then
        Dirf = Dir(F, vbSystem)
    ElseIf LenB(Dir(F, vbHidden + vbSystem)) Then
        Dirf = Dir(F, vbHidden + vbSystem)
    End If
    err.Clear
End Function
'--------------------------------------------------------------------------------------------------------------------------
'Virus Scanner Files manupulations (Like if you launched "notepad" it will find the right executable
'Find in Locations and addind executable extensions
Public Function GetFileFromName(ByVal F As String)
'debug.print "Getting File Name " & F
Dim I  As Long
Dim i2 As Long
Dim F1 As String
Dim S  As Variant
Dim S1 As Variant
    On Error GoTo err
    If InStr(1, F, "\") <= 0 Then
        S = Array(vbNullString, ".exe", ".bat", ".com", ".pif")
        S1 = Array(Environ$("WINDIR") & "\", Environ$("WINDIR") & "\system32\", Environ$("WINDIR") & "\system\")
        For I = 0 To 3
            For i2 = 0 To 2
                F1 = S1(i2) & F & S(I)
                If LenB(Dirf(F1)) Then
                    GetFileFromName = F1
                End If
            Next i2
        Next I
    End If
Exit Function
err:
    MsgBox "GFFN: " & err.Description & vbNewLine & F
    err.Clear
End Function
Public Function GetPath(A2 As String) As String

Dim A3 As Integer
Dim a4 As Integer
    On Error Resume Next
    For a4 = 0 To Len(A2)
        For A3 = 0 To a4
            If Left$(Right$(A2, a4), A3) = "\" Then
                GetPath = Replace$(A2, Right$(A2, a4), vbNullString)
                GoTo end1
            End If
        Next A3
    Next a4
end1:
    On Error GoTo 0
End Function


Public Function MatchExcludes(ByVal F As String) As Boolean
DoEvents
Dim FS As String, F2 As String
    'On Error GoTo Err
      Dim I As Long
MatchExcludes = True
      'Match File Extension
    FS = GetSetting("FreeAV.Scanner", "Settings", "Exclude", "%WINDIR%\system32\drvstore\,%WINDIR%\system32\mui\,%WINDIR%\system32\drivers\,%WINDIR%\pchealth\,%WINDIR%\system32\dllcache\,%WINDIR%\winsxs\;%WINDIR%\Microsoft.Net")
    For I = 0 To UBound(Split(FS, ",")) - 1
    F2 = Split(FS, ",")(I)
    F2 = Replace(F2, "%WINDIR%", Environ("WINDIR"), , , vbTextCompare)
    Debug.Print "Matching " & F & " With " & F2
    
        If UCase$(F & "\") = UCase$(F2) Then
            MatchExcludes = False
            Exit For
        
        End If
    Next I
    
    If MatchExcludes = False Then
    DoEvents
    End If
Exit Function
err:
    MatchExcludes = False
End Function

Public Function MatchFileSet(ByVal F As String) As Boolean
Debug.Print "Matching File " & F
Dim FS As String
    'On Error GoTo err
      Dim I As Long

      'Match File Extension
    FS = GetSetting("FreeAV.Scanner", "Settings", "Fileset", "EXE COM DLL DOC BAT PIF JS VBS ASP JAR SH PL PHP SQL WRI RTF HTML HTM")
    For I = 0 To UBound(Split(FS, " ")) - 1
        If UCase$(Right$(F, Len(F) - InStrRev(F, "."))) = UCase$(Split(FS, " ")(I)) Then
            MatchFileSet = True
            Exit For
        End If
    Next I
    If MatchFileSet = True Then
    'Match Filesize
    If FileLen(F) > (Int(GetSetting("FreeAV.Scanner", "Settings", "MaxFile", "1")) * 1000000) Then
        MatchFileSet = False
    End If
    End If
    
    If MatchFileSet = True Then
    'Skip System Libraries
    Dim C As String
    GetVersionInfo F, C
       
    If InStr(1, C, "Microsoft Corporation", vbTextCompare) Then
    MatchFileSet = False
    End If
    C = vbNullString
    End If
Exit Function
err:
    err.Clear
    MatchFileSet = True
End Function
Public Sub PLaySound(F As String)
    sndPlaySound F, SND_ASYNC
End Sub
':)Code Fixer V3.0.9 (5/12/2009 7:27:47 PM) 6 + 139 = 145 Lines Thanks Ulli for inspiration and lots of code.


