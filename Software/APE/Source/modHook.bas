Attribute VB_Name = "modHook"
Option Explicit

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cb As Long)

Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Const HC_ACTION = 0&
Private Const WH_KEYBOARD_LL = 13&
Private Const VK_LWIN = &H5B&
Private Const VK_RWIN = &H5C&
Private Const VK_TAB = &H9&
Private Const VK_ALT = &HA4&

Private hKeyb As Long

Private Function KeybCallback(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static udtHook As KBDLLHOOKSTRUCT
    
    If (Code = HC_ACTION) Then
        'Copy the keyboard data out of the lParam (which is a pointer)
        Call CopyMemory(udtHook, ByVal lParam, Len(udtHook))
        Select Case udtHook.vkCode
            Case VK_LWIN, VK_RWIN, VK_TAB, VK_ALT
                KeybCallback = 1
                Exit Function
        End Select
    End If
    KeybCallback = CallNextHookEx(hKeyb, Code, wParam, lParam)
End Function

Public Sub HookKeyboard()
    UnhookKeyboard
    hKeyb = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeybCallback, App.hInstance, 0&)
End Sub

Public Sub UnhookKeyboard()
    If hKeyb <> 0 Then
        Call UnhookWindowsHookEx(hKeyb)
        hKeyb = 0
    End If
End Sub
