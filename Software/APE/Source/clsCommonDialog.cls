VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private cdlg As OPENFILENAME
Private LastFileName As String, LastFilter As Long

Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpstrFilter       As String
    lpstrCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpstrFile         As String
    nMaxFile          As Long
    lpstrFileTitle    As String
    nMaxFileTitle     As Long
    lpstrInitialDir   As String
    lpstrTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpstrDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

Public Enum DialogFlags
    FILEMUSTEXIST = OFN_FILEMUSTEXIST
    HIDEREADONLY = OFN_HIDEREADONLY
    OVERWRITEPROMPT = OFN_OVERWRITEPROMPT
    PATHMUSTEXIST = OFN_PATHMUSTEXIST
End Enum

Const sPipe As String = "|"

Public Function ShowOpen(ByVal hwnd As Long, ByVal Title As String, Optional ByVal InitDir As String = "", Optional ByVal Filter As String = "All Files (*.*)|*.*|", Optional flags As DialogFlags = FILEMUSTEXIST Or HIDEREADONLY Or PATHMUSTEXIST) As String
    Filter = Replace(Filter, sPipe, vbNullChar)
    If Right$(Filter, 1) <> vbNullChar Then Filter = Filter & vbNullChar
    If LenB(InitDir) = 0 Then InitDir = LastFileName
    With cdlg
        .lStructSize = Len(cdlg)
        .hwndOwner = hwnd
        .hInstance = App.hInstance
        .lpstrFilter = Filter
        .lpstrFile = Space(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = InitDir
        .lpstrTitle = Title
        .flags = flags
    End With
    
    If GetOpenFileName(cdlg) Then
        ShowOpen = Trim$(cdlg.lpstrFile)
    Else
        ShowOpen = vbNullString
    End If
    
    If LenB(ShowOpen) > 0 Then LastFileName = ShowOpen: LastFilter = cdlg.nFilterIndex

End Function


Public Function ShowSave(ByVal hwnd As Long, ByVal Title As String, Optional ByVal FileName As String = "", Optional InitDir As String = "", Optional ByVal Filter As String, Optional flags As DialogFlags = FILEMUSTEXIST Or OVERWRITEPROMPT Or PATHMUSTEXIST) As String
    Dim Buffer As String
    Dim Result As Long

    Filter = Replace(Filter, sPipe, vbNullChar)
    If Right$(Filter, 1) <> vbNullChar Then Filter = Filter & vbNullChar
    Buffer = FileName & String$(128 - Len(FileName), 0)
    With cdlg
        .lStructSize = Len(cdlg)
        .hwndOwner = hwnd
        .hInstance = App.hInstance
        .flags = flags
        .nMaxFile = Len(Buffer)
        .lpstrFile = Buffer
        .lpstrFilter = Filter
        .lpstrInitialDir = InitDir
        .lpstrTitle = Title
        .lpstrFilter = Filter
        .lpstrDefExt = ".gpl"
    End With
    Result = GetSaveFileName(cdlg)
    If Result <> 0 Then
        ShowSave = Left$(cdlg.lpstrFile, InStr(cdlg.lpstrFile, vbNullChar) - 1)
        LastFilter = cdlg.nFilterIndex
    End If

End Function

'Public Property Get FileName() As String
'    FileName = LastFileName
'End Property

Public Property Get FilterIndex() As Long
    FilterIndex = LastFilter
End Property
