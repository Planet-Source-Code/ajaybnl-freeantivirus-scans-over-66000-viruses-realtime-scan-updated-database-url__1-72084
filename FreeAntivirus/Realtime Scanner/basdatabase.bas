Attribute VB_Name = "basDataBase"
Option Explicit

'Used for storing virus signatures

Public DBF_NAME        As String
Private Const REC_1     As String = "Data"
Private Const FIELD_1   As String = "Field1"
Private Const FIELD_2   As String = "Field2"
Private DBName          As Database
Private mData           As Recordset
Type Entry
    Name                    As String
    Credit                  As String
    Relatives               As String
End Type
Public Function CountRecords() As String
On Error GoTo err
    Set mData = DBName.OpenRecordset(REC_1)
    If mData.RecordCount > 0 Then
        'debug.print mData.Fields(1).Value
        CountRecords = mData.RecordCount & ""
        Else
        CountRecords = "0"
    End If
Exit Function
err:
err.Clear
CountRecords = "0"
End Function
'Private Sub Createtables()
'Dim dbsNewDB As Database
'Dim tdfNew   As TableDef
'    Set dbsNewDB = OpenDatabase(DBF_NAME)
    
      
'    Set tdfNew = dbsNewDB.CreateTableDef(REC_1)
 '   With tdfNew
 '       .Fields.Append .CreateField(FIELD_1, dbText)
 '       .Fields.Append .CreateField(FIELD_2, dbText)
  '      dbsNewDB.TableDefs.Append tdfNew
  '  End With
  '  dbsNewDB.Close
'Exit Sub
'err:
'    MsgBox "ERROR CREATING DATABASE" & vbNewLine & vbNewLine & "Error : " & err.Description
'End Sub
Public Sub LoadRecordset()
On Error GoTo err
    DBF_NAME = App.Path & "\AV.mdb"
    'On Error GoTo err
If Dir(DBF_NAME) <> "" Then
Set DBName = OpenDatabase(DBF_NAME)
End If
Exit Sub
err:

If InStr(1, err.Description, "unrecognize", vbTextCompare) > 0 Then
'MsgBox "No AV.mdb Found! Please Start Antivirus and update it!"
Kill DBF_NAME
err.Clear
Exit Sub
End If
'MsgBox "DataBase -> LoadRecordset : Error: " & err.Description
err.Clear
End Sub
Public Function Match(ByVal D As String) As String
'MsgBox D
    Set mData = DBName.OpenRecordset("SELECT * FROM " & REC_1 & " WHERE " & FIELD_2 & "='" & D & "'")
    If mData.RecordCount > 0 Then
        Match = mData.Fields(FIELD_1)
    Else
        Match = vbNullString
    End If
End Function


'Used to generate Virus DB (IF YOU HAVE VIRUS COLLECTIONS LIKE ME)
''
''Public Function AddData(ByVal strName As String, ByVal Credit As String) As Boolean
''
''
''On Error GoTo err
''Set mData = DBName.OpenRecordset(REC_1)
''
''With mData
''.AddNew
''.Fields(FIELD_1) = strName
''.Fields(FIELD_2) = Credit
'''mData.Fields(FIELD_3) = Relatives
''.Update
''.Bookmark = .LastModified
''End With 'mData
''AddData = True
''Exit Function
''err:
''err1 = err1 + 1
'''MsgBox "ENTRY NOT ADDED DUE TO ERROR " & vbNewLine & vbNewLine & " Error : " & err.Description
''AddData = False
''End Function
''
':)Code Fixer V3.0.9 (5/12/2009 7:27:44 PM) 27 + 97 = 124 Lines Thanks Ulli for inspiration and lots of code.
