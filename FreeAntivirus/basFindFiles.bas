Attribute VB_Name = "basFindFiles"
Option Explicit
'just for manual scanning

Private Const MAX_PATH                     As Integer = 260
Type FILETIME
    dwLowDateTime                              As Long
    dwHighDateTime                             As Long
End Type
Type WIN32_FIND_DATA
    dwFileAttributes                           As Long
    ftCreationTime                             As FILETIME
    ftLastAccessTime                           As FILETIME
    ftLastWriteTime                            As FILETIME
    nFileSizeHigh                              As Long
    nFileSizeLow                               As Long
    dwReserved0                                As Long
    dwReserved1                                As Long
    cFileName                                  As String * MAX_PATH
    cAlternate                                 As String * 14
End Type
Private Const FILE_ATTRIBUTE_DIRECTORY     As Long = &H10
Public FStart                              As Boolean
Public Const PROJECT_KEY                   As String = "FreeAV.Scanner\Clsid"
Public A                                   As String
Public I                                   As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                                                              lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "Kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                                                            lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                              ByVal lpOperation As String, _
                                                                              ByVal lpFile As String, _
                                                                              ByVal lpParameters As String, _
                                                                              ByVal lpDirectory As String, _
                                                                              ByVal nShowCmd As Long) As Long
Public Sub FindFiles(DirPath As String, _
                     Optional FileSpec As String = "*.*")
Dim FindData       As WIN32_FIND_DATA
Dim FindHandle     As Long
Dim FindNextHandle As Long
'Dim filestring As String
    DirPath = Trim$(DirPath)
    If Right$(DirPath, 1) <> "\" Then
        DirPath = DirPath & "\"
    End If
    If Not FStart Then
        Exit Sub
    End If
    FindHandle = FindFirstFile(DirPath & FileSpec, FindData)
    DoEvents
    If FindHandle <> 0 Then
        If (FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            If Left$(FindData.cFileName, 1) <> "." Then
                If Left$(FindData.cFileName, 2) <> ".." Then
                   If MatchExcludes(DirPath & Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1)) Then
                    FindFiles DirPath & Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1), FileSpec
                    End If
                End If
            End If
        ElseIf Len(Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1)) > 0 Then
        If MatchFileSet(DirPath & Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1)) Then
            frmMain.Process Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1), DirPath
            End If
        End If
    End If
' Now loop and find the rest of the files
    If FindHandle <> 0 Then
        Do
            DoEvents
            If Not FStart Then
                Exit Sub
            End If
            FindNextHandle = FindNextFile(FindHandle, FindData)
            If FindNextHandle <> 0 Then
                If (FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                    If Left$(FindData.cFileName, 1) <> "." Then
                        If Left$(FindData.cFileName, 2) <> ".." Then
                         If MatchExcludes(DirPath & Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1)) Then
                            FindFiles DirPath & Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1), FileSpec
                            End If
                        End If
                    End If
                ElseIf Len(Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1)) > 0 Then
                If MatchFileSet(DirPath & Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1)) Then
                    frmMain.Process Left$(FindData.cFileName, InStr(1, FindData.cFileName, vbNullChar) - 1), DirPath
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
    End If
    FindClose FindHandle
End Sub
':)Code Fixer V3.0.9 (5/12/2009 7:27:44 PM) 48 + 74 = 122 Lines Thanks Ulli for inspiration and lots of code.


