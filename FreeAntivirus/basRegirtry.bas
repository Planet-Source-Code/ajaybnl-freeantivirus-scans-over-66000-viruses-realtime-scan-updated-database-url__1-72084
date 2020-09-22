Attribute VB_Name = "basRegistry"
Option Explicit
'Private st1
'Private st2
'Private st3
'Private stt
''Private Const PTSUBKEY                            As String = "software\"
'Private APPKEY
Public Msg                                As String
Public Enum ROOT_HKEY
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_USERS = &H80000003
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_USERS, HKEY_LOCAL_MACHINE, HKEY_CURRENT_CONFIG, HKEY_DYN_DATA, HKEY_PERFORMANCE_DATA
#End If
Private HKEYS                             As New Collection
Private REGTYPES                          As New Collection
Private REGOUT                            As New Collection
''Private Const REG_NONE                            As Integer = 0
Private Const REG_SZ                      As Integer = 1
''Private Const REG_EXPAND_SZ               As Integer = 2
''Private Const REG_BINARY                  As Integer = 3
Private Const REG_DWORD                   As Integer = 4
''Private Const REG_DWORD_LITTLE_ENDIAN             As Integer = 4
''Private Const REG_DWORD_BIG_ENDIAN                As Integer = 5
''Private Const REG_LINK                            As Integer = 6
''Private Const REG_MULTI_SZ                As Integer = 7
''Private Const REG_RESOURCE_LIST                   As Integer = 8
''Private Const REG_FULL_RESOURCE_DESCRIPTOR        As Integer = 9
''Private Const REG_RESOURCE_REQUIREMENTS_LIST      As Integer = 10
Private Const REG_OPTION_NON_VOLATILE     As Integer = 0
Private Const REG_CREATED_NEW_KEY         As Long = &H1
Private Const REG_OPENED_EXISTING_KEY     As Long = &H2
Private Const KEY_QUERY_VALUE             As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS      As Long = &H8
Private Const KEY_NOTIFY                  As Long = &H10
Private Const READ_CONTROL                As Long = &H20000
Private Const STANDARD_RIGHTS_ALL         As Long = &H1F0000
''Private Const STANDARD_RIGHTS_EXECUTE             As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_READ        As Long = (READ_CONTROL)
''Private Const STANDARD_RIGHTS_REQUIRED            As Long = &HF0000
Private Const SYNCHRONIZE                 As Long = &H100000
''Private Const KEY_READ                    As Double = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_SET_VALUE               As Long = &H2
Private Const KEY_CREATE_SUB_KEY          As Long = &H4
Private Const KEY_CREATE_LINK             As Long = &H20
Private Const STANDARD_RIGHTS_WRITE       As Long = (READ_CONTROL)
Private Const KEY_WRITE                   As Double = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS              As Double = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS               As Long = 0
''Private Const ERROR_ACCESS_DENIED                 As Long = 5
''Private Const ERROR_MORE_DATA                     As Long = 234
''Private Const ERROR_NO_MORE_ITEMS                 As Long = 259
''Private Const ERROR_BADKEY                        As Long = 1010
''Private Const ERROR_CANTOPEN                      As Long = 1011
''Private Const ERROR_CANTREAD                      As Long = 1012
''Private Const ERROR_REGISTRY_CORRUPT              As Long = 1015
Type SECURITY_ATTRIBUTES
    nLength                                   As Long
    lpSecurityDescriptor                      As Long
    bInheritHandle                            As Boolean
End Type
Public Type FILETIME
    dwLowDateTime                             As Long
    dwHighDateTime                            As Long
End Type
Public Type KEYARRAY
    cnt                                       As Long
    key()                                     As String
    Data()                                    As Variant
    DataType()                                As Long
    DataSize()                                As Long
End Type
' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal ulOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                      ByVal lpValueName As String, _
                                                                                      ByVal lpReserved As Long, _
                                                                                      lpType As Long, _
                                                                                      lpData As Any, _
                                                                                      dwSize As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                                                                ByVal lpSubKey As String, _
                                                                                ByVal Reserved As Long, _
                                                                                ByVal lpClass As String, _
                                                                                ByVal dwOptions As Long, _
                                                                                ByVal samDesired As Long, _
                                                                                lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                                                                                phkResult As Long, _
                                                                                lpdwDisposition As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal dwReserved As Long, _
                                                                                  ByVal dwType As Long, _
                                                                                  lpValue As Any, _
                                                                                  ByVal dwSize As Long) As Long
''Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, _
                                                                                    ByVal lpValueName As String) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
''Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As FILETIME) As Long
''Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
''Private Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Long, phkResult As Long) As Long
''Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
''Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
''Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
''Private Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long
''Private Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
''Private Declare Function RegOpenKey Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
''Private Declare Function OSRegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
''Private Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Long
''Private Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, ByVal dwFlags As Long) As Long
''Private Declare Function RegSetValueExString Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Function DeleteRegValue(hKey As Long, _
                               SubKey As String, _
                               ValueName As String) As Long
Dim Result     As Long
Dim hKeyResult As Long
    If HKEYS.Count = 0 Then
        InitReg
    End If
    Result = RegOpenKeyEx(hKey, SubKey, 0, KEY_WRITE, hKeyResult)
    If Result <> ERROR_SUCCESS Then
        Exit Function
    End If
    Result = RegDeleteValue(hKeyResult, ValueName)
    DeleteRegValue = Result
    RegCloseKey hKeyResult
End Function
Public Sub InitReg()
    REGTYPES.Add 1, "REG_SZ"
    REGOUT.Add 1, ""
    REGTYPES.Add 2, "REG_EXPAND_SZ"
    REGOUT.Add 2, "hex(2):"
    REGTYPES.Add 3, "REG_BINARY"
    REGOUT.Add 3, "hex:"
    REGTYPES.Add 4, "REG_DWORD"
    REGOUT.Add 4, "dword:"
    With HKEYS
        .Add ROOT_HKEY.HKEY_CLASSES_ROOT, "HKEY_CLASSES_ROOT"
        .Add ROOT_HKEY.HKEY_CURRENT_USER, "HKEY_CURRENT_USER"
        .Add ROOT_HKEY.HKEY_USERS, "HKEY_USERS"
        .Add ROOT_HKEY.HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE"
        .Add ROOT_HKEY.HKEY_CURRENT_CONFIG, "HKEY_CURRENT_CONFIG"
        .Add ROOT_HKEY.HKEY_DYN_DATA, "HKEY_DYN_DATA"
        .Add ROOT_HKEY.HKEY_PERFORMANCE_DATA, "HKEY_PERFORMANCE_DATA"
    End With
End Sub
Public Function ReadReg(hKey As Long, _
                        SubKey As String, _
                        DataName As String, _
                        DefaultData As Variant) As Variant
Dim hKeyResult As Long
Dim lData      As Long
Dim sData      As String
Dim DataType   As Long
Dim DataSize   As Long
Dim Result     As Long
    If HKEYS.Count = 0 Then
        InitReg
    End If
    ReadReg = DefaultData
    Result = RegOpenKeyEx(hKey, SubKey, 0, KEY_QUERY_VALUE, hKeyResult)
    If Result <> ERROR_SUCCESS Then
        Exit Function
    End If
    Result = RegQueryValueEx(hKeyResult, DataName, 0&, DataType, ByVal 0, DataSize)
    If Result <> ERROR_SUCCESS Then
        RegCloseKey hKeyResult
        Exit Function
    End If
    Select Case DataType
    Case REG_SZ
        sData = Space$(DataSize + 1)
        Result = RegQueryValueEx(hKeyResult, DataName, 0&, DataType, ByVal sData, DataSize)
        If Result = ERROR_SUCCESS Then
            ReadReg = CVar(StripNulls(RTrim$(sData)))
        End If
    Case REG_DWORD
        Result = RegQueryValueEx(hKeyResult, DataName, 0&, DataType, lData, 4)
        If Result = ERROR_SUCCESS Then
            ReadReg = CVar(lData)
        End If
    End Select
    RegCloseKey hKeyResult
End Function
Public Function StripNulls(ByVal S As String) As String
Dim I As Integer
    I = InStr(S, vbNullChar) ' Find first Null byte
    If I > 0 Then
        StripNulls = Left$(S, I - 1)
    Else
        StripNulls = S
    End If
End Function
Public Function WriteRegString(hKey As Long, _
                               SubKey As String, _
                               DataName As String, _
                               DataValue As String) As Long
Dim SA           As SECURITY_ATTRIBUTES
Dim hKeyResult   As Long
Dim lDisposition As Long
Dim Result       As Long
    If HKEYS.Count = 0 Then
        InitReg
    End If
    With SA
        .nLength = Len(SA)
        .lpSecurityDescriptor = 0
        .bInheritHandle = False
    End With
    Result = RegCreateKeyEx(hKey, SubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, hKeyResult, lDisposition)
' bug fix?
    If DataValue <= "" Then
        DataValue = vbNullString
    End If
    If (Result = ERROR_SUCCESS) Or (Result = REG_CREATED_NEW_KEY) Or (Result = REG_OPENED_EXISTING_KEY) Then
        Result = RegSetValueEx(hKeyResult, DataName, 0&, REG_SZ, ByVal DataValue, Len(DataValue))
        RegCloseKey hKeyResult
    End If
    WriteRegString = Result
End Function
''
''''
''Public Function DeleteRegKey(hKey As Long, SubKey As String) As Long
''
''
''Dim Result As Long
''If HKEYS.Count = 0 Then
''InitReg
''End If
''Result = RegDeleteKey(hKey, SubKey)
''
''DeleteRegKey = Result
''End Function
''''
''''
''''Public Function GetRegKey(ByVal hKeyRoot As String)
''''
''''
''''
''''
''''
''''Select Case hKeyRoot
''''Case "HKEY_CLASSES_ROOT"
''''GetRegKey = HKEY_CLASSES_ROOT
''''Case "HKEY_CURRENT_USER"
''''GetRegKey = HKEY_CURRENT_USER
''''Case "HKEY_LOCAL_MACHINE"
''''GetRegKey = HKEY_LOCAL_MACHINE
''''Case "HKEY_USERS"
''''GetRegKey = HKEY_USERS
''''Case "HKEY_PERFORMANCE_DATA"
''''GetRegKey = HKEY_PERFORMANCE_DATA
''''Case "HKEY_CURRENT_CONFIG"
''''GetRegKey = HKEY_CURRENT_CONFIG
''''Case "HKEY_DYN_DATA"
''''GetRegKey = HKEY_DYN_DATA
''''End Select
''''End Function
''''
''''
''''Public Function GetSubKeys(hKey As Long, SubKey As String) As KEYARRAY
''''
''''
''''
'''''Dim s() As String
''''Dim hSubKey   As Long
''''Dim i         As Integer
''''Dim lResult   As Long
''''Dim ft        As FILETIME
''''Dim SubKeyCnt As Long
''''If HKEYS.Count = 0 Then
''''InitReg
''''End If
''''lResult = RegOpenKeyEx(hKey, SubKey, 0, KEY_ENUMERATE_SUB_KEYS Or KEY_READ, hSubKey)
''''If lResult <> ERROR_SUCCESS Then
''''GetSubKeys.cnt = 0
''''Exit Function
''''
''''
''''
''''
''''
''''
''''End If
''''lResult = RegQueryInfoKey(hSubKey, vbNullString, 0, 0, SubKeyCnt, 65, 0, 0, 0, 0, 0, ft)
''''If (lResult <> ERROR_SUCCESS) Or (SubKeyCnt <= 0) Then
''''GetSubKeys.cnt = 0
''''Exit Function
''''
''''
''''
''''
''''
''''
''''End If
''''GetSubKeys.cnt = SubKeyCnt
'''''ReDim GetSubKeys.key(SubKeyCnt)
''''For i = 0 To SubKeyCnt - 1
''''With GetSubKeys
''''.key(i) = String$(65, 0)
''''RegEnumKeyEx hSubKey, i, .key(i), 65, 0, vbNullString, 0, ft
''''.key(i) = StripNulls(.key(i))
''''End With 'GetSubKeys
''''Next i
''''RegCloseKey hSubKey
''''End Function
''''
''''
''''Public Function RegGetValuesLong(hKey As String, SubKey As String) As KEYARRAY
''''
''''
''''
''''Dim hSubKey         As Long
''''Dim lResult         As Long
''''Dim ValName         As String
''''Dim ValData         As Long
'''''Dim ValSize As Long
''''
''''Dim LastWriteTime   As FILETIME
''''Dim SubKeyCnt       As Long
''''Dim MaxSubKeyLen    As Long
''''Dim MaxClassLen     As Long
''''Dim ValueCnt        As Long
''''Dim MaxValueNameLen As Long
''''Dim MaxValueLen     As Long
''''Dim SecurityDesc    As Long
'''''Dim i As Integer
'''''Dim ValCnt As Long
'''''Dim TypeCode As Long
''''lResult = RegOpenKeyEx(hKey, SubKey, 0, KEY_READ, hSubKey)
''''If lResult <> ERROR_SUCCESS Then
''''RegGetValuesLong.cnt = 0
''''Else
''''lResult = RegQueryInfoKey(hSubKey, vbNull, 0, 0, SubKeyCnt, MaxSubKeyLen, MaxClassLen, ValueCnt, MaxValueNameLen, MaxValueLen, SecurityDesc, LastWriteTime)
''''RegGetValuesLong.cnt = 0
''''MaxValueNameLen = MaxValueNameLen + 2
''''ValName = String$(MaxValueNameLen + 1, 0)
'''''ValSize = MaxValueNameLen
''''
'''''lResult = RegEnumValue(hSubKey, RegGetValuesLong.cnt, ValName, ValSize, 0, TypeCode, ValData, 4)
''''Do While (lResult = ERROR_SUCCESS) Or (lResult = ERROR_MORE_DATA)
''''With RegGetValuesLong
''''.cnt = .cnt + 1
''''ReDim Preserve .key(.cnt + 1)
''''ReDim Preserve .Data(.cnt + 1)
''''.key(.cnt) = StripNulls(ValName)
''''.Data(.cnt) = ValData
''''End With 'RegGetValuesLong
''''ValName = String$(MaxValueNameLen + 1, 0)
'''''ValSize = MaxValueNameLen
''''ValData = 0
'''''lResult = RegEnumValue(hSubKey, RegGetValuesLong.cnt, ValName, ValSize, 0, TypeCode, ValData, 4)
''''Loop
''''RegCloseKey hSubKey
''''End If
''''End Function
''''
''''
''''Public Sub RegGetWndPos(frm As Form, id As String)
''''
''''
''''
''''Dim rkey As String
''''If HKEYS.Count = 0 Then
''''InitReg
''''End If
''''If LenB(id) = 0 Then
''''id = frm.Name
''''End If
''''rkey = PTSUBKEY & APPKEY & "\" & id
''''With frm
''''.Top = ReadReg(HKEY_CURRENT_USER, rkey, "Top", (Screen.Height - .Height) / 2)
''''.Left = ReadReg(HKEY_CURRENT_USER, rkey, "Left", (Screen.Width - .Width) / 2)
''''If (.Top < 0) Or (.Top > (Screen.Height - 1000)) Then
''''.Top = (Screen.Height - .Height) / 2
''''End If
''''End With 'frm
''''If (frm.Left < 0) Or (frm.Left > (Screen.Width - 1000)) Then
''''frm.Left = (Screen.Width = frm.Width) / 2
''''End If
''''End Sub
''''
''''
''''Public Sub RegGetWndSize(frm As Form, id As String)
''''
''''
''''
''''Dim rkey As String
''''Dim w    As Long
''''Dim h    As Long
''''If HKEYS.Count = 0 Then
''''InitReg
''''End If
''''If LenB(id) = 0 Then
''''id = frm.Name
''''End If
''''rkey = PTSUBKEY & APPKEY & "\" & id
''''w = ReadReg(HKEY_CURRENT_USER, rkey, "Width", frm.Width)
''''h = ReadReg(HKEY_CURRENT_USER, rkey, "Height", frm.Height)
''''If w > 0 Then
''''frm.Width = w
''''End If
''''If h > 0 Then
''''frm.Height = h
''''End If
''''End Sub
''''
''''
''''Public Sub RegSaveWndPos(frm As Form, id As String)
''''
''''
''''
''''Dim rkey As String
''''If HKEYS.Count = 0 Then
''''InitReg
''''End If
''''If frm.WindowState <> vbNormal Then
''''If LenB(id) = 0 Then
''''id = frm.Name
''''End If
''''rkey = PTSUBKEY & APPKEY & "\" & id
''''WriteRegLong HKEY_CURRENT_USER, rkey, "Top", frm.Top
''''WriteRegLong HKEY_CURRENT_USER, rkey, "Left", frm.Left
''''End If
''''End Sub
''''
''''
''''Public Sub RegSaveWndSize(frm As Form, id As String)
''''
''''
''''
''''Dim rkey As String
''''If HKEYS.Count = 0 Then
''''InitReg
''''End If
''''If frm.WindowState <> vbNormal Then
''''If LenB(id) = 0 Then
''''id = frm.Name
''''End If
''''rkey = PTSUBKEY & APPKEY & "\" & id
''''WriteRegLong HKEY_CURRENT_USER, rkey, "Width", frm.Width
''''WriteRegLong HKEY_CURRENT_USER, rkey, "Height", frm.Height
''''End If
''''End Sub
''''
''''
''''Public Function RegWriteStringValue(ByVal hKey, ByVal sValue, ByVal dwDataType, sNewValue) As Long
''''
''''
''''
''''
''''
'''''Dim success As Long
''''Dim dwNewValue As Long
''''dwNewValue = Len(sNewValue)
''''If dwNewValue > 0 Then
''''RegWriteStringValue = RegSetValueExString(hKey, sValue, 0&, dwDataType, sNewValue, dwNewValue)
''''End If
''''End Function
''''
''''
''''Public Function WriteRegBinary(hKey As Long, SubKey As String, DataName As String, DataValue As String) As Long
''''
''''
''''
''''Dim SA           As SECURITY_ATTRIBUTES
''''Dim hKeyResult   As Long
''''Dim lDisposition As Long
''''Dim Result       As Long
''''If HKEYS.Count = 0 Then
''''InitReg
''''End If
''''With SA
''''.nLength = Len(SA)
''''.lpSecurityDescriptor = 0
''''.bInheritHandle = False
''''End With 'SA
''''Result = RegCreateKeyEx(hKey, SubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, hKeyResult, lDisposition)
''''
''''' bug fix?
''''If DataValue <= "" Then
''''DataValue = vbNullString
''''End If
''''If (Result = ERROR_SUCCESS) Or (Result = REG_CREATED_NEW_KEY) Or (Result = REG_OPENED_EXISTING_KEY) Then
''''Result = RegSetValueEx(hKeyResult, DataName, 0&, REG_BINARY, ByVal DataValue, Len(DataValue))
''''
''''RegCloseKey hKeyResult
''''End If
''''WriteRegBinary = Result
''''End Function
''''
''
''Public Function RegGetValues(hKey As String, SubKey As String) As KEYARRAY
''
''
''
''Dim hSubKey         As Long
''Dim I               As Long
''
''Dim S               As String
''Dim lResult         As Long
''Dim ValName         As String
''Dim ValSize         As Long
''Dim LastWriteTime   As FILETIME
''Dim SubKeyCnt       As Long
''Dim MaxSubKeyLen    As Long
''Dim ValueCnt        As Long
''Dim MaxValueNameLen As Long
''Dim MaxValueLen     As Long
''Dim SecurityDesc    As Long
''Dim DataType        As Long
''Dim DataSize        As Long
''Dim ba()            As Byte
'''Dim ValData As String
'''Dim ValCnt As Long
'''Dim MaxClassLen As Long
'''Dim v As Variant
''If HKEYS.Count = 0 Then
''InitReg
''End If
''lResult = RegOpenKeyEx(hKey, SubKey, 0, KEY_ENUMERATE_SUB_KEYS Or KEY_READ, hSubKey)
''If lResult <> ERROR_SUCCESS Then
''RegGetValues.cnt = 0
''Exit Function
''
''
''
''
''
''
''End If
''lResult = RegQueryInfoKey(hSubKey, vbNull, 0, 0, SubKeyCnt, MaxSubKeyLen, vbNull, ValueCnt, MaxValueNameLen, MaxValueLen, SecurityDesc, LastWriteTime)
''RegGetValues.cnt = 0
''ValSize = MaxValueNameLen + 100
''ValName = String$(ValSize + 1, 0)
'''ValData = String$(ValSize + 1, 0)
''ReDim ba(MaxValueLen + 1) As Byte
''ValSize = MaxValueNameLen + 100
''ValName = String$(ValSize + 1, 0)
'''ValData = String$(ValSize + 1, 0)
''DataSize = UBound(ba) - 1
''lResult = RegEnumValue(hSubKey, RegGetValues.cnt, ValName, Len(ValName), 0, DataType, ba(0), DataSize)
''Do While lResult = ERROR_SUCCESS
''With RegGetValues
''.cnt = .cnt + 1
''ReDim Preserve .key(.cnt + 1)
''ReDim Preserve .DataType(.cnt + 1)
''ReDim Preserve .DataSize(.cnt + 1)
''ReDim Preserve .Data(.cnt + 1)
''.key(.cnt) = StripNulls(ValName)
''.DataType(.cnt) = DataType
''End With 'RegGetValues
''Select Case DataType
''Case REG_SZ ' SZ
''S = vbNullString
''I = 0
''Do While I < DataSize + 1
''S = S & Chr$(ba(I))
''I = I + 1
''Loop
''RegGetValues.Data(RegGetValues.cnt) = StripNulls(S)
''RegGetValues.DataSize(RegGetValues.cnt) = Len(S)
''Case REG_EXPAND_SZ
''S = vbNullString
''I = 0
''Do While I < (DataSize * 2) + 1
''S = S & Chr$(ba(I))
''I = I + 1
''Loop
''RegGetValues.Data(RegGetValues.cnt) = S
''RegGetValues.DataSize(RegGetValues.cnt) = (DataSize * 2)
''Case REG_MULTI_SZ
''S = vbNullString
''I = 0
''Do While I < (DataSize * 2) + 1
''S = S & Chr$(ba(I))
''I = I + 1
''Loop
''RegGetValues.Data(RegGetValues.cnt) = S
''RegGetValues.DataSize(RegGetValues.cnt) = (DataSize * 2)
''Case REG_BINARY ' binary
''RegGetValues.Data(RegGetValues.cnt) = ba
''RegGetValues.DataSize(RegGetValues.cnt) = DataSize
''Case REG_DWORD 'dword
''I = ba(0)
''I = I + (CLng(ba(1)) * 256)
''I = I + (CLng(ba(2)) * 256 * 256)
''S = Hex$(I)
''If Len(S) < 6 Then
''S = String$(6 - Len(S), "0") & S
''End If
'''i = (CLng(ba(3)) * 256 * 256 * 256)
''S = "&h" & Hex$(ba(3)) & S
''RegGetValues.Data(RegGetValues.cnt) = Val(S)
''RegGetValues.DataSize(RegGetValues.cnt) = 4
''Case Else
''RegGetValues.Data(RegGetValues.cnt) = ba
''RegGetValues.DataSize(RegGetValues.cnt) = DataSize
''End Select
''DataType = 0
''DataSize = UBound(ba) - 1
''ValSize = MaxValueNameLen + 100
''ValName = String$(ValSize + 1, 0)
'''ValData = String$(ValSize + 1, 0)
''lResult = RegEnumValue(hSubKey, RegGetValues.cnt, ValName, Len(ValName), 0, DataType, ba(0), DataSize)
''Loop
''RegCloseKey hSubKey
''End Function
''
''
''Public Function WriteRegLong(hKey As Long, SubKey As String, DataName As String, DataValue As Long) As Long
''
''
''Dim SA           As SECURITY_ATTRIBUTES
''Dim hKeyResult   As Long
''Dim lDisposition As Long
''Dim Result       As Long
''If HKEYS.Count = 0 Then
''InitReg
''End If
''With SA
''.nLength = Len(SA)
''.lpSecurityDescriptor = 0
''.bInheritHandle = False
''End With 'SA
''Result = RegCreateKeyEx(hKey, SubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SA, hKeyResult, lDisposition)
''
''If (Result = ERROR_SUCCESS) Or (Result = REG_CREATED_NEW_KEY) Or (Result = REG_OPENED_EXISTING_KEY) Then
''Result = RegSetValueEx(hKeyResult, DataName, 0&, REG_DWORD, DataValue, 4)
''
''RegCloseKey hKeyResult
''End If
''WriteRegLong = Result
''End Function
''
''
':)Code Fixer V3.0.9 (5/12/2009 7:27:46 PM) 98 + 572 = 670 Lines Thanks Ulli for inspiration and lots of code.


