Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Const MAX_STRING_LEN As Long = 100
Private Const MAX_SECTION_LEN As Long = 70000
Public sGameCode As String * 4
Public Const INIFile As String = "\Bookmarks.ini"
Public Const CompressedPal As String = "<C>"

Private Sub SetAttrIfExists(sFile As String, vbFileAttribute As Byte)
On Error Resume Next
    SetAttr sFile, vbFileAttribute
End Sub
    
'Read a string from an ini file
Public Function sReadIniFileString(ByVal INIFile As String, ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = "") As String
    On Error GoTo sReadIniFileString_err
    Dim nLength As Long
    Dim sTemp As String
    sTemp = Space$(MAX_STRING_LEN)
    nLength = GetPrivateProfileString(Section, Key, Default, sTemp, MAX_STRING_LEN, INIFile)
    sReadIniFileString = Mid$(sTemp, 1, nLength)
    Exit Function
sReadIniFileString_err:
    sReadIniFileString = Default
End Function

'Read a long integer from an ini file
'Public Function lReadIniFileLong(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, Optional ByVal Default As Long = 0) As Long
'    On Error GoTo lReadIniFileLong_err
'    Dim sTemp As String
'    'Use existing function to get value back
'    '     as string
'    sTemp = sReadIniFileString(IniFile, Section, Key, "")
'
'    If Len(sTemp) = 0 Then
'        lReadIniFileLong = Default
'    Else
'        lReadIniFileLong = CLng(sTemp)
'    End If
'    Exit Function
'lReadIniFileLong_err:
'    lReadIniFileLong = Default
'End Function

'Read a double from an ini file
'Public Function dReadIniFileDouble(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, Optional ByVal Default As Double = 0) As Double
'    On Error GoTo dReadIniFileDouble_err
'    Dim sTemp As String
'    'Use existing function to get value back
'    '     as string
'    sTemp = sReadIniFileString(IniFile, Section, Key, "")
'
'    If Len(sTemp) = 0 Then
'        dReadIniFileDouble = Default
'    Else
'        dReadIniFileDouble = CDbl(sTemp)
'    End If
'    Exit Function
'dReadIniFileDouble_err:
'    dReadIniFileDouble = Default
'End Function

'Read a date from an ini file
'This will return a collection containing all entries for a given section

Public Function colReadIniFileSection(ByVal INIFile As String, ByVal Section As String) As Collection
    On Error GoTo colReadIniFileSection_err
    Dim sTemp As String
    Dim nPos As Long
    Dim nLength As Long
    Set colReadIniFileSection = New Collection
    sTemp = Space$(MAX_SECTION_LEN)
    nLength = GetPrivateProfileSection(Section, sTemp, MAX_SECTION_LEN, INIFile)
    sTemp = Mid$(sTemp, 1, nLength)
    nPos = InStr(1, sTemp, "=")
 
    Do While nPos > 0
        colReadIniFileSection.Add Mid$(sTemp, 1, nPos - 1)
        nPos = InStr(1, sTemp, vbNullChar)
        sTemp = Mid$(sTemp, nPos + 1)
        nPos = InStr(1, sTemp, "=")
 
        DoEvents
        Loop
 
        If Len(sTemp) > 0 Then
            colReadIniFileSection.Add sTemp
        End If
        Exit Function
colReadIniFileSection_err:
    End Function

'Write a string to an ini file
Public Function bWriteIniFileString(ByVal INIFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
    'On Error GoTo bWriteIniFileString_err
    Dim nRetVal As Long
    
    SetAttrIfExists INIFile, vbNormal
    
    'Clear existing entry first
    '(there is a problem with encrypted values otherwise)
    WritePrivateProfileString Section, Key, vbNullString, INIFile
    nRetVal = WritePrivateProfileString(Section, Key, Value, INIFile)
 
    If nRetVal > 0 Then
        bWriteIniFileString = True
    Else
        bWriteIniFileString = False
    End If
    Exit Function
bWriteIniFileString_err:
    bWriteIniFileString = False
End Function

'Write a long integer to an ini file
'Public Function bWriteIniFileLong(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As Long) As Boolean
'    bWriteIniFileLong = bWriteIniFileString(IniFile, Section, Key, CStr(Value))
'End Function

'Write a double to an ini file
'Public Function bWriteIniFileDouble(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As Double) As Boolean
'    bWriteIniFileDouble = bWriteIniFileString(IniFile, Section, Key, CStr(Value))
'End Function

'Write a date to an ini file
'Public Function bWriteIniFileDate(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As Date) As Boolean
'    bWriteIniFileDate = bWriteIniFileString(IniFile, Section, Key, Format$(Value, "dd mmm yyyy hh:nn:ss"))
'End Function

'This will return a collection containing all entries for a given section
'Public Function colReadIniFileSectionNames(ByVal INIFile As String) As Collection
'    On Error GoTo colReadIniFileSectionNames_err
'    Dim sTemp As String
'    Dim nPos As Long
'    Dim nLength As Long
'    Set colReadIniFileSectionNames = New Collection
'    sTemp = Space$(MAX_SECTION_LEN)
'    nLength = GetPrivateProfileSectionNames(sTemp, MAX_SECTION_LEN, INIFile)
'    sTemp = Mid$(sTemp, 1, nLength)
'    nPos = InStr(1, sTemp, vbNullChar)
'
'    Do While nPos > 0
'        colReadIniFileSectionNames.Add Mid$(sTemp, 1, nPos - 1)
'        sTemp = Mid$(sTemp, nPos + 1)
'        nPos = InStr(1, sTemp, vbNullChar)
'
'        DoEvents
'        Loop
'
'        If Len(sTemp) > 0 Then
'            colReadIniFileSectionNames.Add sTemp
'        End If
'        Exit Function
'colReadIniFileSectionNames_err:
'    End Function
    
'This will remove an entry
Public Function bRemoveIniFileEntry(ByVal INIFile As String, ByVal Section As String, ByVal Key As String) As Boolean
    On Error GoTo bRemoveIniFileEntry_err
    bRemoveIniFileEntry = False
    
    SetAttrIfExists INIFile, vbNormal
    
    If WritePrivateProfileString(Section, Key, vbNullString, INIFile) > 0 Then
        bRemoveIniFileEntry = True
    Else
        bRemoveIniFileEntry = False
    End If
bRemoveIniFileEntry_err:
End Function

'Rename a string in an INI file
Public Function bRenameIniFileString(ByVal INIFile As String, ByVal Section As String, ByVal Key As String, ByVal NewKey As String) As Boolean
    On Error GoTo bRenameIniFileEntry_err
    
    Dim tmpValue As String
    tmpValue = sReadIniFileString(INIFile, Section, Key) ' store the previous key value
    bRemoveIniFileEntry INIFile, Section, Key ' remove old key
    bWriteIniFileString INIFile, Section, NewKey, tmpValue ' create the new renamed one
    
bRenameIniFileEntry_err:
End Function
