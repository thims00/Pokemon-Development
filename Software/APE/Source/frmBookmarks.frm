VERSION 5.00
Begin VB.Form frmBookmarks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bookmarks"
   ClientHeight    =   4710
   ClientLeft      =   13785
   ClientTop       =   2790
   ClientWidth     =   2970
   DrawStyle       =   5  'Transparent
   Icon            =   "frmBookmarks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   Tag             =   "5000"
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Caption         =   "Edit Bookmark"
      Height          =   2175
      Left            =   3000
      TabIndex        =   21
      Tag             =   "5020"
      Top             =   1800
      Width           =   4455
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         ScaleHeight     =   375
         ScaleWidth      =   975
         TabIndex        =   25
         Top             =   360
         Width           =   975
         Begin VB.CommandButton cmdChange 
            Caption         =   "Change"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   7
            Tag             =   "5004"
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.TextBox txtBookmarkOffset 
         Height          =   285
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtBookmarkDescription 
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   9
         Top             =   885
         Width           =   2775
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   2775
         TabIndex        =   24
         Top             =   1440
         Width           =   2775
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Tag             =   "5010"
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton optRemoveType 
            Caption         =   "All"
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   11
            Tag             =   "5011"
            Top             =   60
            Width           =   1335
         End
         Begin VB.OptionButton optRemoveType 
            Caption         =   "Selected"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   12
            Tag             =   "5012"
            Top             =   300
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   240
         X2              =   4200
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Offset"
         Height          =   195
         Left            =   2520
         TabIndex        =   23
         Tag             =   "5003"
         Top             =   480
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Left            =   360
         TabIndex        =   22
         Tag             =   "5009"
         Top             =   920
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1530
      TabIndex        =   15
      Tag             =   "5019"
      Top             =   4200
      Width           =   1140
   End
   Begin VB.CommandButton cmdSwitch 
      Caption         =   "Editor"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Tag             =   "5002"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox chkOpen 
      Caption         =   "Keep open after loading bookmark"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Tag             =   "5013"
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Enabled         =   0   'False
      Height          =   375
      Left            =   300
      TabIndex        =   14
      Tag             =   "5014"
      Top             =   4200
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Add Bookmark"
      Height          =   1575
      Left            =   3000
      TabIndex        =   16
      Tag             =   "5005"
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtOtherOffset 
         Height          =   285
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   5
         Top             =   620
         Width           =   975
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   240
         ScaleHeight     =   375
         ScaleWidth      =   975
         TabIndex        =   20
         Top             =   440
         Width           =   975
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   375
            Left            =   0
            TabIndex        =   2
            Tag             =   "5006"
            Top             =   0
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   1440
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   19
         Top             =   260
         Width           =   2055
         Begin VB.OptionButton optAddType 
            Caption         =   "Other Offset"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   4
            Tag             =   "5008"
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optAddType 
            Caption         =   "Actual Offset"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Tag             =   "5007"
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   6
         Top             =   1040
         Width           =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   195
         Left            =   340
         TabIndex        =   17
         Tag             =   "5009"
         Top             =   1050
         Width           =   795
      End
   End
   Begin VB.ListBox lstBookmark 
      Height          =   2985
      ItemData        =   "frmBookmarks.frx":000C
      Left            =   120
      List            =   "frmBookmarks.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bookmarks"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Tag             =   "5001"
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmBookmarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LB_FINDSTRING = &H18F
Private Const LB_FINDSTRINGEXACT = &H1A2
Private Const LB_ADDSTRING = &H180
Private Const CB_ADDSTRING = &H143
 
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim i As Long

Public Sub FindItem(lst As ListBox, strText As String)
Dim iIndex As Integer
    
    iIndex = SendMessage(lst.hwnd, LB_FINDSTRING, -1, ByVal strText) 'LB_FINDSTRINGEXACT
    
    If iIndex >= 0 Then
        lst.ListIndex = iIndex
        lst.SetFocus
        frmBookmarkSearch.txtSearch.SetFocus
    End If
    
End Sub

Public Sub LoadINI()

lstBookmark.Clear
lstBookmark.Enabled = False

For i = 1 To colReadIniFileSection(App.Path & INIFile, sGameCode).Count
    Call SendMessageString(lstBookmark.hwnd, LB_ADDSTRING, 0, ByVal colReadIniFileSection(App.Path & INIFile, sGameCode).Item(i))
Next

lstBookmark.Enabled = True

If lstBookmark.ListCount > 0 Then
    lstBookmark.ListIndex = 0
    cmdLoad.Enabled = True
    cmdChange.Enabled = True
    cmdRemove.Enabled = True
Else
    txtDescription.SetFocus
    txtBookmarkOffset.Text = vbNullString
    Call cmdSwitch_Click
End If

End Sub

Public Sub LoadINIMain()

frmMain.cboLoadedBookmark.Clear
frmMain.cboLoadedBookmark.Enabled = False

For i = 1 To colReadIniFileSection(App.Path & INIFile, sGameCode).Count
    Call SendMessageString(frmMain.cboLoadedBookmark.hwnd, CB_ADDSTRING, 0, ByVal colReadIniFileSection(App.Path & INIFile, sGameCode).Item(i))
Next

frmMain.cboLoadedBookmark.Enabled = True

If frmMain.cboLoadedBookmark.ListCount > 1 Then
    frmMain.cboLoadedBookmark.ListIndex = 0
    frmMain.cboLoadedBookmark.Text = frmMain.cboLoadedBookmark.List(0)
    frmMain.cboLoadedBookmark.Tag = 0
    frmMain.cmdLoadBookmark.Enabled = True
    frmMain.cmdNextBookmark.Enabled = True
Else
    frmMain.cboLoadedBookmark.Text = vbNullString
    frmMain.cboLoadedBookmark.Enabled = False
    frmMain.cmdPrevBookmark.Enabled = False
    frmMain.cmdNextBookmark.Enabled = False
End If

End Sub

Private Sub cmdAdd_Click()
Dim lngRetVal As Long

lngRetVal = SendMessageString(lstBookmark.hwnd, LB_FINDSTRINGEXACT, -1&, txtDescription.Text)

If lngRetVal <= -1& Then
    ' It's not in the list so add it
    Call SendMessage(lstBookmark.hwnd, LB_ADDSTRING, 0, ByVal txtDescription.Text)
End If

If optAddType(1).Value = True Then
    txtOtherOffset.Text = Right$(String$(7, vbKey0) & txtOtherOffset.Text, txtOtherOffset.MaxLength)
    bWriteIniFileString App.Path & INIFile, sGameCode, txtDescription.Text, txtOtherOffset.Text
Else
    If LenB(frmMain.txtOffset.Text) > 0 Then
        bWriteIniFileString App.Path & INIFile, sGameCode, txtDescription.Text, frmMain.txtOffset.Text
    End If
End If

For i = 0 To lstBookmark.ListCount - 1
    If lstBookmark.List(i) = txtDescription.Text Then
        lstBookmark.ListIndex = i
        Exit For
    End If
Next

cmdRemove.Enabled = True
cmdChange.Enabled = True
cmdLoad.Enabled = True

End Sub

Public Sub cmdLoad_Click()

frmMain.optOffset.Value = True
frmMain.txtOffset.Text = sReadIniFileString(App.Path & INIFile, sGameCode, lstBookmark.List(lstBookmark.ListIndex))

frmMain.cboLoadedBookmark.Text = lstBookmark.List(lstBookmark.ListIndex)

If lstBookmark.ListCount > 1 Then
    Select Case lstBookmark.ListIndex
        Case 0
            frmMain.cmdPrevBookmark.Enabled = False
            frmMain.cmdNextBookmark.Enabled = True
        Case lstBookmark.ListCount - 1
            frmMain.cmdPrevBookmark.Enabled = True
            frmMain.cmdNextBookmark.Enabled = False
        Case Else
            frmMain.cmdPrevBookmark.Enabled = True
            frmMain.cmdNextBookmark.Enabled = True
    End Select
Else
    frmMain.cmdPrevBookmark.Enabled = False
    frmMain.cmdNextBookmark.Enabled = False
End If

If Left$(lstBookmark.List(lstBookmark.ListIndex), 3) = CompressedPal Then
    frmMain.chkCompressed.Value = vbChecked
Else
    frmMain.chkCompressed.Value = vbUnchecked
End If

Call frmMain.cmdLoadPalette_Click
If chkOpen.Value = vbUnchecked Then Unload Me: frmMain.SetFocus

End Sub

Private Sub cmdRefresh_Click()
    Call LoadINI
    Call LoadINIMain
End Sub

Private Sub cmdRemove_Click()
If optRemoveType(0).Value = True Then

    Dim Answer As Byte
    Me.SetFocus
    Answer = MsgBox(LoadResString(5016), vbYesNo + vbExclamation)
    
    If Answer = vbNo Then Exit Sub
    
    For i = 0 To lstBookmark.ListCount - 1
        bRemoveIniFileEntry App.Path & INIFile, sGameCode, lstBookmark.List(i)
    Next
    
    lstBookmark.Clear
    txtBookmarkOffset.Text = vbNullString
    cmdRemove.Enabled = False
    cmdChange.Enabled = False
    cmdLoad.Enabled = False
        
Else
    bRemoveIniFileEntry App.Path & INIFile, sGameCode, lstBookmark.List(lstBookmark.ListIndex)
    
    If lstBookmark.ListIndex > 0 Then
        lstBookmark.ListIndex = lstBookmark.ListIndex - 1
        lstBookmark.RemoveItem lstBookmark.ListIndex + 1
    Else
        lstBookmark.RemoveItem lstBookmark.ListIndex
        If lstBookmark.ListCount > 0 Then
            lstBookmark.ListIndex = 0
        Else
            txtBookmarkOffset.Text = vbNullString
            cmdRemove.Enabled = False
            cmdChange.Enabled = False
            cmdLoad.Enabled = False
        End If
    End If
    
End If

End Sub

Private Sub cmdChange_Click()
If LenB(txtBookmarkOffset.Text) > 0 And txtBookmarkDescription.Text = lstBookmark.List(lstBookmark.ListIndex) Then
    txtBookmarkOffset.Text = Right$(String$(7, vbKey0) & txtBookmarkOffset.Text, txtBookmarkOffset.MaxLength)
    bWriteIniFileString App.Path & INIFile, sGameCode, lstBookmark.List(lstBookmark.ListIndex), txtBookmarkOffset.Text
ElseIf LenB(txtBookmarkDescription.Text) > 0 And txtBookmarkDescription.Text <> lstBookmark.List(lstBookmark.ListIndex) Then
    bRenameIniFileString App.Path & INIFile, sGameCode, lstBookmark.List(lstBookmark.ListIndex), txtBookmarkDescription.Text
    lstBookmark.List(lstBookmark.ListIndex) = txtBookmarkDescription.Text
    frmMain.cboLoadedBookmark.List(lstBookmark.ListIndex) = txtBookmarkDescription.Text
End If
End Sub

Public Sub cmdSwitch_Click()
If cmdSwitch.Caption = LoadResString(5002) Then
    Me.Caption = LoadResString(5017)
    Me.Width = 7670
    Me.Left = CLng(Me.Left - 3060)
    cmdLoad.Left = 2655
    cmdRefresh.Left = 3885
    cmdSwitch.Caption = LoadResString(5018)
ElseIf cmdSwitch.Caption = LoadResString(5018) And lstBookmark.ListCount > 0 Then
    Me.Caption = LoadResString(5001)
    Me.Width = 3060
    Me.Left = CLng(Me.Left + 3060)
    cmdLoad.Left = 300
    cmdRefresh.Left = 1530
    cmdSwitch.Caption = LoadResString(5002)
End If
End Sub

Private Sub Form_Load()
    LoadResStrings Me
    Me.Left = CLng(frmMain.Left + 8310)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmBookmarkSearch
End Sub

Private Sub lstBookmark_Click()
    txtBookmarkOffset.Text = sReadIniFileString(App.Path & INIFile, sGameCode, lstBookmark.List(lstBookmark.ListIndex))
    txtBookmarkDescription.Text = lstBookmark.List(lstBookmark.ListIndex)
End Sub

Private Sub lstBookmark_DblClick()
    Call cmdLoad_Click
End Sub

Private Sub lstBookmark_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        frmBookmarkSearch.Show , Me
    End If
End Sub


Private Sub mnuSearch_Click()
    MsgBox "Search!!!"
End Sub

Private Sub optAddType_Click(Index As Integer)
    
    Call txtDescription_Change
    
    If Index = 0 Then
        txtOtherOffset.Locked = True
        txtOtherOffset.Text = frmMain.txtOffset.Text
    Else
        txtOtherOffset.Locked = False
    End If
    
End Sub

Private Sub txtBookmarkDescription_KeyPress(KeyCode As Integer)
    Call txtDescription_KeyPress(KeyCode)
End Sub

Private Sub txtBookmarkOffset_KeyPress(KeyCode As Integer)
If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
    If (KeyCode > 64 And KeyCode < vbKeyG) Then Exit Sub
    If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub
    If (KeyCode > 47 And KeyCode < 58) Then Exit Sub
    KeyCode = 0
End If
End Sub

Private Sub txtDescription_Change()
If optAddType(1).Value = True Then
    If Len(txtOtherOffset.Text) > 0 And Len(txtDescription.Text) > 0 Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
Else
    If Len(frmMain.txtOffset.Text) > 0 And Len(txtDescription.Text) > 0 Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
End If
End Sub

Private Sub txtDescription_KeyPress(KeyCode As Integer)
    If KeyCode = 59 Or KeyCode = 61 Or KeyCode = 91 Or KeyCode = 93 Then KeyCode = 0
End Sub

Private Sub txtOtherOffset_Change()
    Call txtDescription_Change
End Sub

Private Sub txtOtherOffset_KeyPress(KeyCode As Integer)
    Call txtBookmarkOffset_KeyPress(KeyCode)
End Sub
