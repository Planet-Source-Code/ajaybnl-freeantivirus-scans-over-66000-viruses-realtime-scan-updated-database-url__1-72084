Attribute VB_Name = "basVersionInfo"

 Option Explicit

 ' API
 Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
 (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, _
 lpData As Any) As Long
 Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias _
 "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
 Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock _
 As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
 Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 _
 As Any, ByVal lpString2 As Any) As Long
 Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, _
 Source As Any, ByVal Length As Long)

 Public Type VS_FIXEDFILEINFO
 dwSignature As Long
 dwStrucVersion As Long
 dwFileVersionMS As Long
 dwFileVersionLS As Long
 dwProductVersionMS As Long
 dwProductVersionLS As Long
 dwFileFlagsMask As Long
 dwFileFlags As Long
 dwFileOS As Long
 dwFileType As Long
 dwFileSubtype As Long
 dwFileDateMS As Long
 dwFileDateLS As Long
 End Type
Public Const MAXGETHOSTSTRUCT = 1024
Public Const GMEM_FIXED = &H0
 Private Const VFT_APP = &H1
 Private Const VFT_DLL = &H2
 Private Const VFT_DRV = &H3
 Private Const VFT_VXD = &H5
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Dim m_lngMemoryHandle As Long, m_lngMemoryPointer As Long




 Public Function HiWord(ByVal dwValue As Long) As Long
 Dim hexstr As String
 hexstr = Right("00000000" & Hex(dwValue), 8)
 HiWord = CLng("&H" & Left(hexstr, 4))
 End Function

 Public Function LoWord(ByVal dwValue As Long) As Long
 Dim hexstr As String
 hexstr = Right("00000000" & Hex(dwValue), 8)
 LoWord = CLng("&H" & Right(hexstr, 4))
 End Function

 ' Swap de 2 valeurs de type 'byte' avec XOR
 Public Sub SwapByte(byte1 As Byte, byte2 As Byte)
 byte1 = byte1 Xor byte2
 byte2 = byte1 Xor byte2
 byte1 = byte1 Xor byte2
 End Sub

 ' Creation d'une chaine Hexadecimale pour représenter un nombre
 Public Function FixedHex(ByVal hexval As Long, ByVal nDigits As Long) As String
 FixedHex = Right("00000000" & Hex(hexval), nDigits)
 End Function
Private Function AllocateMemory() As Long

    m_lngMemoryHandle = GlobalAlloc(GMEM_FIXED, MAXGETHOSTSTRUCT)
    If m_lngMemoryHandle <> 0 Then
        m_lngMemoryPointer = GlobalLock(m_lngMemoryHandle)
        If m_lngMemoryPointer <> 0 Then
            GlobalUnlock (m_lngMemoryHandle)
            AllocateMemory = m_lngMemoryPointer
        Else 'NOT M_LNGMEMORYPOINTER...
            GlobalFree (m_lngMemoryHandle)
            AllocateMemory = m_lngMemoryPointer '0
        End If
    Else 'NOT M_LNGMEMORYHANDLE...
        AllocateMemory = m_lngMemoryHandle '0
    End If

End Function
Private Sub FreeMemory()

    If m_lngMemoryHandle <> 0 Then
        m_lngMemoryHandle = 0
        m_lngMemoryPointer = 0
        GlobalFree m_lngMemoryHandle
    End If

End Sub
 Public Sub GetVersionInfo(ByVal sFileName As String, sCopyright As String)
 On Error GoTo err
 
 Dim vffi As VS_FIXEDFILEINFO
 Dim buffer() As Byte
 Dim pData As Long
 Dim nDataLen As Long
 Dim cpl(0 To 3) As Byte
 Dim cplstr As String
 Dim retval As Long

 'pData = AllocateMemory
 
 nDataLen = GetFileVersionInfoSize(sFileName, pData)
If nDataLen = 0 Then
'FreeMemory
Exit Sub
End If

 ' Récupération de la 'Version' du fichier
 ' ---------------------------------------
 ' Make the buffer large enough to hold the version info resource.
 ReDim buffer(0 To nDataLen - 1) As Byte
 
 retval = GetFileVersionInfo(sFileName, 0, nDataLen, buffer(0))

 retval = VerQueryValue(buffer(0), "\", pData, nDataLen)
 
 CopyMemory vffi, ByVal pData, nDataLen

 retval = VerQueryValue(buffer(0), "\VarFileInfo\Translation", pData, nDataLen)
 
 CopyMemory cpl(0), ByVal pData, 4
 
 SwapByte cpl(0), cpl(1)
 SwapByte cpl(2), cpl(3)
 
 cplstr = FixedHex(cpl(0), 2) & FixedHex(cpl(1), 2) & FixedHex(cpl(2), 2) & _
 FixedHex(cpl(3), 2)

 retval = VerQueryValue(buffer(0), "\StringFileInfo\" & cplstr & "\LegalCopyright", _
 pData, nDataLen)
 
 sCopyright = Space(nDataLen)
 retval = lstrcpy(sCopyright, pData)
sCopyright = Replace(sCopyright, Chr(0), "")
 
 Exit Sub
err:
err.Clear
 'FreeMemory

 End Sub
