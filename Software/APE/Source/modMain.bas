Attribute VB_Name = "modMain"
Public Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Const LR_SHARED = &H8000&
Private Const IMAGE_ICON = 1
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

Public Sub Main()

    On Error Resume Next
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    On Error GoTo 0
    frmMain.Show

End Sub

Public Sub SetIcon(ByVal hwnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
    Dim lhWndTop As Long
    Dim lhWnd As Long
    Dim cx As Long
    Dim cy As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long

    If (bSetAsAppIcon) Then
        ' Find VB's hidden parent window:
        lhWnd = hwnd
        lhWndTop = lhWnd
        Do While Not (lhWnd = 0)
            lhWnd = GetWindow(lhWnd, GW_OWNER)
            If Not (lhWnd = 0) Then
                lhWndTop = lhWnd
            End If
        Loop
    End If
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    End If
    SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    End If
    SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall

End Sub


Sub LoadResStrings(frm As Form)

    On Error Resume Next
    Dim ctl As Control
    Dim sCtlType As String
    Dim nVal As Integer
    ' set the form's caption
    frm.Caption = LoadResString(CInt(frm.Tag))
    For Each ctl In frm.Controls
        sCtlType = TypeName(ctl)
        If sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.HelpContextID))
        Else
            nVal = 0
            nVal = Val(ctl.Tag)
            If nVal > 0 Then ctl.Caption = LoadResString(nVal)
            nVal = 0
            nVal = Val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(nVal)
        End If
    Next

End Sub


Public Function FileExists(sFileName As String) As Boolean

    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If

End Function


Public Sub KillFileIfExists(ByVal sFile As String)
    On Error Resume Next
    Kill sFile
End Sub

Function Dual(ByVal vValue As Long) As String

    vValue = CVar(vValue)
    If Not IsNumeric(vValue) Then
        Dual = "Value is non-numerical!"
        Exit Function
    ElseIf vValue > 999999999 Then
        Dual = "Number is too high!"
        Exit Function
    End If
    Do
        If vValue Mod 2 = 0 Then
            Dual = "0" & Dual
        Else
            Dual = "1" & Dual
        End If
        vValue = vValue \ 2
    Loop While vValue > 0
    Dual = Format(Dual, "000000000000000")

End Function

Function Round(ByVal Number As Double, Optional ByVal NumDigitsAfterDecimal As Integer = 0) As Double
    Round = Int(Number * 10 ^ NumDigitsAfterDecimal + 0.5) / 10 ^ NumDigitsAfterDecimal
End Function

Public Function BinToDecI(Bin As String) As Integer
    Dim i As Integer
    Dim nDec As Integer
    Dim nPos As Integer

    If Len(Bin) > 16 Then
        Err.Raise 6
    Else
        For i = Len(Bin) To 1 Step -1
            If Mid$(Bin, i, 1) = "1" Then
                SetBitI nDec, nPos
            End If
            nPos = nPos + 1
        Next
        BinToDecI = nDec
    End If

End Function


Public Sub SetBitI(Value As Integer, ByVal Position As Byte)

    Select Case Position
        Case 0 To 14
            Value = Value Or 2 ^ Position
        Case 15
            Value = Value Or &H8000
        Case Else
            Err.Raise 6
    End Select

End Sub


Function GB2RGB(GBPalette As String, Optional bSwitch As Boolean = False) As String

    Dim Red As String, Green As String, Blue As String
    Dim speicher As String, speicherL As Long
     
    speicherL = Val("&H" & Right$(GBPalette, 2) & Left$(GBPalette, 2))
    speicher = Dual(speicherL)
    Blue = Left$(speicher, 5)
    speicher = Right$(speicher, 10)
    Green = Left$(speicher, 5)
    Red = Right$(speicher, 5)
    Red = BinToDecI(Red)
    Red = Round((Red * 255) / 31)
    Green = BinToDecI(Green)
    Green = Round((Green * 255) / 31)
    Blue = BinToDecI(Blue)
    Blue = Round((Blue * 255) / 31)
    Red = Right$("0" & Hex$(Red), 2)
    Green = Right$("0" & Hex$(Green), 2)
    Blue = Right$("0" & Hex$(Blue), 2)
      
    If Not bSwitch Then
        GB2RGB = Red & Green & Blue
    Else
        GB2RGB = Blue & Green & Red
    End If
   
    
End Function


Function RGB2GBA(RGBPalette As String) As String
    Dim Red As String, Green As String, Blue As String, speicher As String

    Red = Val("&H" & Left$(RGBPalette, 2))
    RGBPalette = Right$(RGBPalette, 4)
    Green = Val("&H" & Left$(RGBPalette, 2))
    Blue = Val("&H" & Right$(RGBPalette, 2))
    speicher = Hex$((Blue \ 8) * 1024 + (Green \ 8) * 32 + (Red \ 8))
    speicher = Right$("000" & speicher, 4)
    RGB2GBA = Right$(speicher, 2) & Left$(speicher, 2)

End Function
