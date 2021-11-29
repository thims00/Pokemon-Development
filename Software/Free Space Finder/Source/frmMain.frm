VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "FSF - Free Space Finder"
   ClientHeight    =   5505
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   7875
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":000C
   ScaleHeight     =   5505
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Search Options"
      Height          =   2655
      Left            =   240
      TabIndex        =   12
      Tag             =   "1"
      Top             =   1680
      Width           =   3855
      Begin VB.TextBox txtSearchInterval 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   17
         Text            =   "1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txtOffset 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "0"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   3375
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3375
         Begin VB.OptionButton optOffset 
            Caption         =   "Search from offset"
            Enabled         =   0   'False
            Height          =   435
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Tag             =   "5"
            Top             =   240
            Width           =   2055
         End
         Begin VB.OptionButton optOffset 
            Caption         =   "Search from the beginning of the ROM"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   6
            Tag             =   "4"
            Top             =   0
            Value           =   -1  'True
            Width           =   3135
         End
      End
      Begin VB.ComboBox cboByteType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMain.frx":37C5
         Left            =   1920
         List            =   "frmMain.frx":37CF
         TabIndex        =   4
         Text            =   "FF"
         Top             =   405
         Width           =   855
      End
      Begin VB.TextBox txtNeededBytes 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "64"
         Top             =   885
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search Interval"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Tag             =   "11"
         Top             =   1335
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Needed Bytes"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Tag             =   "3"
         Top             =   920
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Free-Space Byte"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Tag             =   "2"
         Top             =   450
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Tag             =   "8"
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   7695
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2007 HackMew"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2910
         TabIndex        =   10
         Tag             =   "1017"
         Top             =   195
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Tag             =   "7"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ListBox lstResult 
      Height          =   1620
      ItemData        =   "frmMain.frx":37DB
      Left            =   4920
      List            =   "frmMain.frx":37DD
      TabIndex        =   2
      Tag             =   "0"
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Tag             =   "6"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblGameCode 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROM: None"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   5400
      TabIndex        =   18
      Tag             =   "10"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "- NOTE: All offsets are in hexadecimal -"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2235
      TabIndex        =   11
      Tag             =   "9"
      Top             =   4560
      Width           =   3405
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   100
      Begin VB.Menu mnuOpenROM 
         Caption         =   "Open ROM..."
         HelpContextID   =   101
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         HelpContextID   =   102
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   103
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Enabled         =   0   'False
         HelpContextID   =   104
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         HelpContextID   =   105
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         HelpContextID   =   106
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   107
      Begin VB.Menu mnuReadme 
         Caption         =   "Readme"
         HelpContextID   =   108
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         HelpContextID   =   109
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sResult As String

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText lstResult.List(lstResult.ListIndex)
End Sub

Private Sub cmdFindNext_Click()
Dim FileNum As Integer
FileNum = FreeFile
    
    Dim lngIndex As Long
    Dim strContents As String, strSearch As String, strFound As Long
    
    Open sResult For Binary As #FileNum
        strContents = Space$(LOF(FileNum))
        Get #FileNum, , strContents
    Close #FileNum
    
    strSearch = String$(Val(txtNeededBytes.Text) * 2, Val("&H" & cboByteType.Text))
    lngIndex = InStr(lstResult.Tag + 1 + Val(txtSearchInterval.Text), strContents, strSearch, vbBinaryCompare)
    
    If Not lngIndex = 0 Then
      strFound = lngIndex - 1
      lstResult.Tag = strFound
      lstResult.AddItem Hex$(strFound)
      lstResult.ListIndex = lstResult.ListIndex + 1
    Else
      MsgBox LoadResString(1001), vbExclamation
End If
End Sub

Private Sub cmdSearch_Click()
Dim FileNum As Integer
FileNum = FreeFile
    
    Dim lngIndex As Long
    Dim strContents As String, strSearch As String, strFound As Long
    
    '  Reads the whole file into the String variable all at once
    Open sResult For Binary As #FileNum
        strContents = Space$(LOF(FileNum))
        Get #FileNum, , strContents
    Close #FileNum
    
    If cboByteType.ListIndex < 0 Then cboByteType.ListIndex = 1
    
    '  strContents now contains all of the text in the file
    strSearch = String$(Val(txtNeededBytes.Text) * 2, Val("&H" & cboByteType.List(cboByteType.ListIndex)))
    lngIndex = InStr(CLng("&H" & txtOffset.Text) + 1, strContents, strSearch, vbBinaryCompare)
    
    cmdFindNext.Enabled = False
    cmdCopy.Enabled = False
    mnuFindNext.Enabled = False
    mnuCopy.Enabled = False
    
    lstResult.Clear
    
    If Not lngIndex = 0 Then
        strFound = lngIndex - 1
        lstResult.Tag = strFound
        lstResult.AddItem Hex$(strFound)
        lstResult.ListIndex = 0
        cmdFindNext.Enabled = True
        mnuFindNext.Enabled = True
        cmdCopy.Enabled = True
        mnuCopy.Enabled = True
    Else
        MsgBox LoadResString(1000), vbExclamation
    End If
    
End Sub

Private Sub Form_Load()
    SetIcon Me.hWnd, "AAA"
    LoadResStrings Me
    lblGameCode.Caption = "ROM: " & lblGameCode.Caption
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, frmMain
End Sub

Private Sub mnuCopy_Click()
    Call cmdCopy_Click
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFindNext_Click()
    Call cmdFindNext_Click
End Sub

Private Sub mnuOpenROM_Click()

cmdSearch.Enabled = False
cmdFindNext.Enabled = False
cmdCopy.Enabled = False
mnuSearch.Enabled = False
mnuFindNext.Enabled = False
mnuCopy.Enabled = False
lblGameCode.Caption = "ROM: " & LoadResString(10)

lstResult.Clear

Dim oOpenDialog As clsCommonDialog
Set oOpenDialog = New clsCommonDialog

sResult = oOpenDialog.ShowOpen(Me.hWnd, LoadResString(1002), , "GameBoy Advance ROMs (*.gba, *.agb, *.bin)|*.gba;*.agb;*.bin|", FILEMUSTEXIST Or PATHMUSTEXIST Or HIDEREADONLY)

If LenB(sResult) > 0 Then
    
    cmdSearch.Enabled = True
    mnuSearch.Enabled = True
    cboByteType.Enabled = True
    txtNeededBytes.Enabled = True
    txtSearchInterval.Enabled = True
    optOffset(0).Enabled = True
    optOffset(1).Enabled = True
    cboByteType.ListIndex = 1
    
    Dim FileNum As Integer
    Dim sGameCode As String * 4
    
    FileNum = FreeFile
    Open sResult For Binary As #FileNum
        Get #FileNum, &HAD, sGameCode
        lblGameCode = "ROM: " & sGameCode
    Close #FileNum
    
End If

End Sub

Private Sub mnuReadme_Click()
    Shell "notepad.exe " & App.Path & "\Readme.txt", vbNormalFocus
End Sub

Private Sub mnuSearch_Click()
    Call cmdSearch_Click
End Sub

Private Sub optOffset_Click(Index As Integer)
    If Index = 0 Then txtOffset.Enabled = True Else txtOffset.Enabled = False: txtOffset.Text = 0
End Sub

Private Sub txtNeededBytes_KeyPress(KeyCode As Integer)
    If Not IsNumeric(Chr$(KeyCode)) Then
        KeyCode = 0
    End If
End Sub

Private Sub txtOffset_KeyPress(KeyCode As Integer)
If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
    If (KeyCode > 64 And KeyCode < 71) Then Exit Sub 'A-F are OK
    If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub 'a-f become A-F
    If (KeyCode > 47 And KeyCode < 58) Then Exit Sub 'Numbers are OK
KeyCode = 0 'All other letters are unwanted.
End If
End Sub

Private Sub txtSearchInterval_Change()
    If Val(txtSearchInterval.Text) < 1 Then txtSearchInterval.Text = 1
End Sub

Private Sub txtSearchInterval_KeyPress(KeyCode As Integer)
    Call txtNeededBytes_KeyPress(KeyCode)
End Sub
