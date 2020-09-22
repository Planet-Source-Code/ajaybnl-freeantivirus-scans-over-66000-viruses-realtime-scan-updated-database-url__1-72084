VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl FileDownloader 
   BackStyle       =   0  'Transparent
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9450
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   9450
   ToolboxBitmap   =   "FileDownload.ctx":0000
   Begin InetCtlsObjects.Inet Inet 
      Left            =   5640
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image Image 
      Height          =   945
      Left            =   0
      Picture         =   "FileDownload.ctx":0312
      Top             =   0
      Width           =   1020
   End
End
Attribute VB_Name = "FileDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2004 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Private Declare Function GetDiskFreeSpace Lib "Kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Public Event DownloadErrors(strError As String)
Public Event DownloadEvents(strEvent As String)
Public Event DowloadComplete()
Public Event DownloadProgress(intPercent As String)

Private CancelSearch As Boolean

Private Sub UserControl_Resize()
    Width = 1020
    Height = 945
End Sub

Public Sub Cancel()
    CancelSearch = True
End Sub

Public Function DownloadFile(strURL As String, strDestination As String, Optional UserName As String = Empty, Optional Password As String = Empty) As Boolean
    Const CHUNK_SIZE As Long = 1024
    Const ROLLBACK As Long = 4096

    Dim bData() As Byte
    Dim blnResume As Boolean
    Dim intFile As Integer
    Dim lngBytesReceived As Long
    Dim lngFileLength As Long
    Dim strFile As String
    Dim strHeader As String
    Dim strHost As String
    
On Local Error GoTo InternetErrorHandler
    
    CancelSearch = False

    strFile = ReturnFileOrFolder(strDestination, True)
    strHost = ReturnFileOrFolder(strURL, True, True)

StartDownload:

    If blnResume Then
        RaiseEvent DownloadEvents("Resuming download")
        lngBytesReceived = lngBytesReceived - ROLLBACK
        If lngBytesReceived < 0 Then lngBytesReceived = 0
    Else
        RaiseEvent DownloadEvents("Getting file information")
    End If

    DoEvents
    
    With Inet
        .URL = strURL
        .UserName = UserName
        .Password = Password
    
        .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
        
        While .StillExecuting
            DoEvents
            If CancelSearch = True Then GoTo ExitDownload
        Wend

        strHeader = .GetHeader
    End With
    
    Select Case Mid(strHeader, 10, 3)
        Case "200"
            If blnResume Then
                Kill strDestination
                RaiseEvent DownloadErrors("The server is unable to resume this download.")
                CancelSearch = True
                GoTo ExitDownload
            End If
        Case "206"
        
        Case "204"
            RaiseEvent DownloadErrors("Nothing to download!")
            CancelSearch = True
            GoTo ExitDownload
        Case "401"
            RaiseEvent DownloadErrors("Authorization failed!")
            CancelSearch = True
            GoTo ExitDownload
        Case "404"
            RaiseEvent DownloadErrors("The file, " & """" & Inet.URL & """" & " was not found!")
            CancelSearch = True
            GoTo ExitDownload
        Case vbCrLf
            RaiseEvent DownloadErrors("Cannot establish connection.")
            CancelSearch = True
            GoTo ExitDownload
        Case Else
            strHeader = Left(strHeader, InStr(strHeader, vbCr))
            If strHeader = Empty Then strHeader = "<nothing>"
            RaiseEvent DownloadErrors("The server returned the following response:" & vbCr & vbCr & strHeader)
            CancelSearch = True
            GoTo ExitDownload
    End Select

    If blnResume = False Then
        strHeader = Inet.GetHeader("Content-Length")
        lngFileLength = Val(strHeader)
        If lngFileLength = 0 Then
            GoTo ExitDownload
        End If
    End If

    If Mid(strDestination, 2, 2) = ":\" Then
        If DiskFreeSpace(Left(strDestination, InStr(strDestination, "\"))) < lngFileLength Then
            RaiseEvent DownloadErrors("There is not enough free space on disk for this file.")
            GoTo ExitDownload
        End If
    End If

    DoEvents
    
    If blnResume = False Then lngBytesReceived = 0

On Local Error GoTo FileErrorHandler

    strHeader = ReturnFileOrFolder(strDestination, False)
    If Dir(strHeader, vbDirectory) = Empty Then
        MkDir strHeader
    End If

    intFile = FreeFile()

    Open strDestination For Binary Access Write As #intFile

    If blnResume Then Seek #intFile, lngBytesReceived + 1
    Do
        bData = Inet.GetChunk(CHUNK_SIZE, icByteArray)
        Put #intFile, , bData
        If CancelSearch Then Exit Do
        lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
        RaiseEvent DownloadProgress(Round((lngBytesReceived / lngFileLength) * 100))
        DoEvents
    Loop While UBound(bData, 1) > 0

    Close #intFile

ExitDownload:

    If lngBytesReceived = lngFileLength Then
        If CancelSearch = False Then RaiseEvent DowloadComplete
        DownloadFile = True
    Else
        If Dir(strDestination) = Empty Then
            CancelSearch = True
        Else
            If CancelSearch = False Then
                RaiseEvent DownloadErrors("The connection with the server was reset.")
            End If
        End If
        If Not Dir(strDestination) = Empty Then Kill strDestination
        DownloadFile = False
    End If

CleanUp:

    Inet.Cancel
    
    Exit Function

InternetErrorHandler:
    
    If err.Number = 9 Then Resume Next
    RaiseEvent DownloadErrors("Error: " & err.Description & " occurred.")
    err.Clear
    GoTo ExitDownload
    
FileErrorHandler:

    RaiseEvent DownloadErrors("Cannot write file to disk." & vbCr & vbCr & "Error " & err.Number & ": " & err.Description)
    CancelSearch = True
    err.Clear
    GoTo ExitDownload
End Function

Private Function ReturnFileOrFolder(FullPath As String, ReturnFile As Boolean, Optional IsURL As Boolean = False) As String
    Dim intDelimiterIndex As Integer

    intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
    
    If intDelimiterIndex = 0 Then
        ReturnFileOrFolder = FullPath
    Else
        ReturnFileOrFolder = IIf(ReturnFile, Right(FullPath, Len(FullPath) - intDelimiterIndex), Left(FullPath, intDelimiterIndex))
    End If
End Function

Private Function DiskFreeSpace(strDrive As String) As Double
    Dim SectorsPerCluster As Long
    Dim BytesPerSector As Long
    Dim NumberOfFreeClusters As Long
    Dim TotalNumberOfClusters As Long
    Dim FreeBytes As Long
    Dim spaceInt As Integer

    strDrive = QualifyPath(strDrive)

    GetDiskFreeSpace strDrive, SectorsPerCluster, BytesPerSector, NumberOFreeClusters, TotalNumberOfClusters

    DiskFreeSpace = NumberOFreeClusters * SectorsPerCluster * BytesPerSector
End Function

Private Function QualifyPath(strPath As String) As String
    QualifyPath = IIf(Right(strPath, 1) = "\", strPath, strPath & "\")
End Function
