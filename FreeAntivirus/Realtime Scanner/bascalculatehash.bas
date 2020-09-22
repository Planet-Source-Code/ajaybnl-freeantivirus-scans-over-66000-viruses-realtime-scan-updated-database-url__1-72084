Attribute VB_Name = "basCalculateHash"
Option Explicit
'Calculate File Hash or crc

Public Type tagInitCommonControlsEx
    lngSize                             As Long
    lngICC                              As Long
End Type
' .... AdvancedAPI
' .... Constant
Private Const PROV_RSA_FULL         As Integer = 1
Private Const CRYPT_NEWKEYSET       As Long = &H8
Private Const ALG_CLASS_HASH        As Long = 32768
Private Const ALG_TYPE_ANY          As Integer = 0
Private Const ALG_SID_MD2           As Integer = 1
Private Const ALG_SID_MD4           As Integer = 2
Private Const ALG_SID_MD5           As Integer = 3
Private Const ALG_SID_SHA1          As Integer = 4
' .... Enum
Public Enum HashAlgorithm
    MD2 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
    MD4 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
    MD5 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
    SHA1 = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private MD2, MD4, MD5, SHA1
#End If
' .... Oter Constant
Private Const HP_HASHVAL            As Integer = 2
Private Const HP_HASHSIZE           As Integer = 4
''Private Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, _
                                                                                              ByVal pszContainer As String, _
                                                                                              ByVal pszProvider As String, _
                                                                                              ByVal dwProvType As Long, _
                                                                                              ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, _
                                                                 ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, _
                                                             ByVal Algid As Long, _
                                                             ByVal hKey As Long, _
                                                             ByVal dwFlags As Long, _
                                                             ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, _
                                                           pbData As Byte, _
                                                           ByVal dwDataLen As Long, _
                                                           ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, _
                                                               ByVal dwParam As Long, _
                                                               pbData As Any, _
                                                               pdwDataLen As Long, _
                                                               ByVal dwFlags As Long) As Long
Public Function HashFile(ByVal Filename As String, _
                         Optional ByVal Algorithm As HashAlgorithm = MD5) As String
                         'debug.print "Hashing File " & Filename
Dim txtInfo      As String
Dim hCtx         As Long
Dim hHash        As Long
Dim lFile        As Long
Dim lRes         As Long
Dim lLen         As Long
Dim lIdx         As Long
Dim abHash()     As Byte
Const BLOCK_SIZE As Long = 32 * 1024&   ' 32K
Dim lCount       As Long
Dim lBlocks      As Long
Dim lLastBlock   As Long
    On Error Resume Next
'txtInfo.Text = Empty
' .... Check if the file exists
    'If Len(Dire(Filename)) = 0 Then
     '   GoTo err
    'End If
' .... Get default provider context handle
    lRes = CryptAcquireContext(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0)
    If lRes = 0 Then
        If err.LastDllError = &H80090016 Then
' .... There's no default keyset container
' .... Get the provider context and create a default keyset container
            lRes = CryptAcquireContext(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET)
        End If
    End If
    If lRes <> 0 Then
' .... Create the hash
        lRes = CryptCreateHash(hCtx, Algorithm, 0, 0, hHash)
        If lRes <> 0 Then
' .... Get a file handle
            lFile = FreeFile
' .... Open the file
            'Debug.Print FileLen(Filename)
            Open Filename For Binary As lFile
            If err.Number = 0 Then
' .... Init Const Block Size = 32x32 Kb ;)
                ReDim abBlock(1 To BLOCK_SIZE) As Byte
' .... Calculate how many full blocks the file contains
                lBlocks = LOF(lFile) \ BLOCK_SIZE
                txtInfo = txtInfo & "Block Size: " & lBlocks & vbNewLine
' .... Calculate the remaining data length
                lLastBlock = LOF(lFile) - lBlocks * BLOCK_SIZE
' .... Calculate total Size
'lblSize.Caption = "File Size: " & FormatSize(LOF(lFile))
                txtInfo = txtInfo & "Block Remaining: " & lLastBlock & vbNewLine
'PB.Max = lBlocks
' .... Hash the blocks
                For lCount = 1 To lBlocks
                    Get lFile, , abBlock
' .... Add the chunk to the hash
                    lRes = CryptHashData(hHash, abBlock(1), BLOCK_SIZE, 0)
' .... Stop the loop if CryptHashData fails
                    If lRes = 0 Then
                        Exit For
                    End If
'DoEvents
'PB.Value = lCount
                Next lCount
' .... Is there more data?
                If lLastBlock > 0 Then
                    If lRes <> 0 Then
' .... Get the last block
                        ReDim abBlock(1 To lLastBlock) As Byte
                        Get lFile, , abBlock
' .... Hash the last block
                        lRes = CryptHashData(hHash, abBlock(1), lLastBlock, 0)
                    End If
                End If
' .... Close the file
                Close lFile
            End If
            If lRes <> 0 Then
' .... Get the hash lenght
                lRes = CryptGetHashParam(hHash, HP_HASHSIZE, lLen, 4, 0)
                If lRes <> 0 Then
' .... Initialize the buffer
                    ReDim abHash(0 To lLen - 1) As Byte
' .... Get the hash value
                    lRes = CryptGetHashParam(hHash, HP_HASHVAL, abHash(0), lLen, 0)
                    If lRes <> 0 Then
'PB.Max = UBound(abHash)
' .... Convert value to hex string
                        For lIdx = 0 To UBound(abHash)
                            HashFile = HashFile & Right$("0" & Hex$(abHash(lIdx)), 2)
'DoEvents
'PB.Value = lIdx
'txtInfo = txtInfo & "Hex String: " & Right$("0" & Hex$(abHash(lIdx)), 2) & vbCrLf
                        Next lIdx
                    End If
                End If
            End If
' .... Release the hash handle
            CryptDestroyHash hHash
        End If
    End If
' .... Release the provider context
    CryptReleaseContext hCtx, 0
'PB.Value = 0
'txtInfo = txtInfo & "": txtInfo = txtInfo & "Finish!"
' .... Raise an error if lRes = 0
    If lRes = 0 Then

    End If

err.Clear
End Function
''
''Private Function FormatSize(size As Variant) As String
''
''
''
''On Local Error Resume Next
''If size >= 1073741824 Then
''If size <= 1099511627776# Then
''
''FormatSize = Format$(((size / 1024) / 1024) / 1024, "#") & " GB"
''Exit Function
''
''
''
''End If
''
''End If
''If size >= 1048576 Then
''If size <= 1073741824 Then
''
''FormatSize = Format$((size / 1024) / 1024, "#") & " MB"
''Exit Function
''
''
''
''End If
''
''End If
''If size >= 1024 Then
''If size <= 1048576 Then
''
''FormatSize = Format$(size / 1024, "#") & " KB"
''Exit Function
''
''
''
''End If
''
''End If
''If size < 1024 Then
''FormatSize = size & " bytes"
''
''End If
''End Function
''
':)Code Fixer V3.0.9 (5/12/2009 7:27:44 PM) 71 + 175 = 246 Lines Thanks Ulli for inspiration and lots of code.


