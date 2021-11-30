VERSION 5.00
Begin VB.Form frmGradient 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gradient-o-matic"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6330
   Icon            =   "frmGradient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "7000"
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3730
      TabIndex        =   52
      Tag             =   "7013"
      Top             =   4200
      Width           =   1200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   360
      ScaleHeight     =   1335
      ScaleWidth      =   5655
      TabIndex        =   19
      Top             =   2640
      Width           =   5655
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   7
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   35
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   6
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   34
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   5
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   33
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   4
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   32
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   3
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   31
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   30
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   1
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   29
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   28
         Top             =   0
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   8
         Left            =   0
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   27
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   9
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   26
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   10
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   25
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   11
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   24
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   12
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   23
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   13
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   22
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   14
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   21
         Top             =   720
         Width           =   600
      End
      Begin VB.TextBox txtGradient 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   15
         Left            =   5040
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   20
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   0
         Left            =   0
         TabIndex        =   51
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   1
         Left            =   720
         TabIndex        =   50
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   2
         Left            =   1440
         TabIndex        =   49
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   3
         Left            =   2160
         TabIndex        =   48
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   4
         Left            =   2880
         TabIndex        =   47
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   5
         Left            =   3600
         TabIndex        =   46
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   6
         Left            =   4320
         TabIndex        =   45
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   7
         Left            =   5040
         TabIndex        =   44
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   8
         Left            =   0
         TabIndex        =   43
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   9
         Left            =   720
         TabIndex        =   42
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   10
         Left            =   1440
         TabIndex        =   41
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   11
         Left            =   2160
         TabIndex        =   40
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   12
         Left            =   2880
         TabIndex        =   39
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   13
         Left            =   3600
         TabIndex        =   38
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   14
         Left            =   4320
         TabIndex        =   37
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblGradient 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   15
         Left            =   5040
         TabIndex        =   36
         Top             =   1080
         Width           =   600
      End
   End
   Begin VB.CheckBox chkOpen 
      Caption         =   "Keep open after inserting gradient"
      Height          =   255
      Left            =   140
      TabIndex        =   8
      Tag             =   "7011"
      Top             =   4240
      Width           =   3615
   End
   Begin VB.TextBox txtStartingColor 
      Height          =   285
      Left            =   480
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "FF7F"
      Top             =   960
      Width           =   615
   End
   Begin VB.CheckBox chkInsert 
      Caption         =   "Don't insert the gradient"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Tag             =   "7007"
      Top             =   1260
      Width           =   3015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5015
      TabIndex        =   9
      Tag             =   "7012"
      Top             =   4200
      Width           =   1200
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Tag             =   "7009"
      Top             =   1800
      Width           =   1155
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Enabled         =   0   'False
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Tag             =   "7008"
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gradient colors"
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Tag             =   "7010"
      Top             =   2280
      Width           =   6100
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1575
      Left            =   2760
      TabIndex        =   12
      Tag             =   "7003"
      Top             =   120
      Width           =   3460
      Begin VB.TextBox txtTo 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2775
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "4"
         Top             =   760
         Width           =   375
      End
      Begin VB.TextBox txtFrom 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   10
         Text            =   "1"
         Top             =   760
         Width           =   375
      End
      Begin VB.TextBox txtColorNumber 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   2
         Text            =   "4"
         Top             =   320
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   195
         Left            =   2400
         TabIndex        =   16
         Tag             =   "7006"
         Top             =   780
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insert colors from"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Tag             =   "7005"
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colors number"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Tag             =   "7004"
         Top             =   345
         Width           =   1005
      End
   End
   Begin VB.TextBox txtEndingColor 
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   1
      Text            =   "FF7F"
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblEndingColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1680
      TabIndex        =   18
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblStartingColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   480
      TabIndex        =   17
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ending Color"
      Height          =   195
      Left            =   1560
      TabIndex        =   11
      Tag             =   "7002"
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Color"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Tag             =   "7001"
      Top             =   1440
      Width           =   945
   End
End
Attribute VB_Name = "frmGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IntMax = 32767
Private Const BytMax = 255

Private GradientCreated As Boolean


Private Sub chkInsert_Click()
    Label4.Enabled = Not CBool(chkInsert.Value)
    txtFrom.Enabled = Not CBool(chkInsert.Value)
    Label5.Enabled = Not CBool(chkInsert.Value)
    chkOpen.Enabled = Not CBool(chkInsert.Value)
    If GradientCreated Then
        cmdInsert.Enabled = Not CBool(chkInsert.Value)
    Else
        cmdInsert.Enabled = False
    End If
End Sub

Private Sub GradientInsert()
Dim X As Byte, i As Byte
    
    If chkInsert.Value = vbUnchecked Then
        If LenB(txtFrom.Text) > 0 And LenB(txtTo.Text) > 0 Then
            For i = txtFrom.Text - 1 To txtTo.Text - 1
                frmMain.txtChgPal(i).Text = txtGradient(X).Text
                X = X + 1
            Next
            If chkOpen.Value = vbUnchecked Then Unload Me: frmMain.SetFocus
        End If
    End If
    
End Sub

Private Sub cmdCreate_Click()

    Dim StartRGB As String, EndRGB As String
    Dim RedGradient() As Integer, GreenGradient() As Integer, BlueGradient() As Integer
    Dim MaxNumber As Byte
    Dim RedDiff As Double, GreenDiff As Double, BlueDiff As Double
    Dim i As Byte, tmpRGB As String
    
    GradientCreated = False
    
    StartRGB = GB2RGB(txtStartingColor.Text)
    EndRGB = GB2RGB(txtEndingColor.Text)
    
    MaxNumber = CByte(txtColorNumber.Text - 1)
    
    ReDim RedGradient(MaxNumber) As Integer
    ReDim GreenGradient(MaxNumber) As Integer
    ReDim BlueGradient(MaxNumber) As Integer
    
    RedGradient(0) = Val("&H" & Left$(StartRGB, 2))
    GreenGradient(0) = Val("&H" & Mid$(StartRGB, 3, 2))
    BlueGradient(0) = Val("&H" & Right$(StartRGB, 2))
       
    RedGradient(MaxNumber) = Val("&H" & Left$(EndRGB, 2))
    GreenGradient(MaxNumber) = Val("&H" & Mid$(EndRGB, 3, 2))
    BlueGradient(MaxNumber) = Val("&H" & Right$(EndRGB, 2))
    
    RedDiff = (RedGradient(0) - RedGradient(MaxNumber)) / MaxNumber
    GreenDiff = (GreenGradient(0) - GreenGradient(MaxNumber)) / MaxNumber
    BlueDiff = (BlueGradient(0) - BlueGradient(MaxNumber)) / MaxNumber
    
    For i = LBound(RedGradient()) + 1 To UBound(RedGradient()) - 1
        RedGradient(i) = RedGradient(i - 1) - RedDiff
        GreenGradient(i) = GreenGradient(i - 1) - GreenDiff
        BlueGradient(i) = BlueGradient(i - 1) - BlueDiff
        If RedGradient(i) > BytMax Then RedGradient(i) = BytMax
        If GreenGradient(i) > BytMax Then GreenGradient(i) = BytMax
        If BlueGradient(i) > BytMax Then BlueGradient(i) = BytMax
    Next
    
    For i = txtGradient.LBound To txtGradient.UBound
        txtGradient(i).Text = vbNullString
        lblGradient(i).BackColor = vbWhite
        If i < MaxNumber + 1 Then
            txtGradient(i).Visible = True: lblGradient(i).Visible = True
        Else
            txtGradient(i).Visible = False: lblGradient(i).Visible = False
        End If
    Next
    
    For i = 0 To MaxNumber
        tmpRGB = CStr(Right$("0" & Hex$(RedGradient(i)), 2) & Right$("0" & Hex$(GreenGradient(i)), 2) & Right$("0" & Hex$(BlueGradient(i)), 2))
        txtGradient(i).Text = RGB2GBA(tmpRGB)
        lblGradient(i).BackColor = "&H" & GB2RGB(txtGradient(i).Text, True)
    Next
    
    GradientCreated = True
    
    If chkInsert.Value = vbUnchecked Then
        cmdInsert.Enabled = True
    Else
        cmdInsert.Enabled = False
    End If
    
    Call GradientInsert
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    Call GradientInsert
End Sub

Private Sub cmdReset_Click()

    Dim i As Byte
    
    txtStartingColor.Text = vbNullString
    txtEndingColor.Text = vbNullString
    txtColorNumber.Text = vbNullString
    txtFrom.Text = vbNullString
    txtTo.Text = vbNullString
    
    For i = txtGradient.LBound To txtGradient.UBound
        txtGradient(i).Text = vbNullString
        txtGradient(i).Visible = True
        lblGradient(i).BackColor = vbWhite
        lblGradient(i).Visible = True
    Next
    
    cmdInsert.Enabled = False
    GradientCreated = False
    
End Sub

Private Sub CreateCheck()

    If Len(txtStartingColor.Text) = 4 And Len(txtEndingColor.Text) = 4 And LenB(txtColorNumber.Text) > 0 Then
        If txtStartingColor.Text <> txtEndingColor.Text Then
            cmdCreate.Enabled = True
        End If
    Else
        cmdCreate.Enabled = False
    End If
    
End Sub

Private Sub Form_Activate()

Dim strColor As String, strType As String
       
    If Len(Me.Tag) > 6 Then
        strColor = Mid$(Me.Tag, 1, 4)
        strType = Mid$(Me.Tag, 5, 3)
    Else
        Exit Sub
    End If
    
    Select Case strType
        Case "GrS"
            txtStartingColor.Text = strColor
        Case "GrE"
            txtEndingColor.Text = strColor
    End Select
    
    Me.Tag = vbNullString
    
End Sub

Private Sub Form_Load()
    LoadResStrings Me
End Sub

Private Sub lblEndingColor_DblClick()
    
    If Len(txtEndingColor.Text) = 4 Then
        frmColorPicker.picNewColor.BackColor = lblEndingColor.BackColor
        frmColorPicker.tmrTransfer.Enabled = True
        frmColorPicker.Tag = "GrE"
    End If

End Sub

Private Sub lblStartingColor_DblClick()
    
    If Len(txtStartingColor.Text) = 4 Then
        frmColorPicker.picNewColor.BackColor = lblStartingColor.BackColor
        frmColorPicker.tmrTransfer.Enabled = True
        frmColorPicker.Tag = "GrS"
    End If
    
End Sub

Private Sub txtColorNumber_Change()
    
    Call CreateCheck
    Call txtFrom_Change
    Call txtFrom_KeyUp(0, 0)

End Sub

Private Sub txtColorNumber_KeyPress(KeyCode As Integer)

    If KeyCode <> vbKeyBack Then
        If Not IsNumeric(Chr$(KeyCode)) Then KeyCode = 0
    End If

End Sub

Private Sub txtColorNumber_LostFocus()
    
    If Val(txtColorNumber.Text) > 16 Then
        txtColorNumber.Text = 16
    ElseIf Val(txtColorNumber.Text) < 3 And LenB(txtColorNumber.Text) > 0 Then
        txtColorNumber.Text = 3
    End If
    
End Sub

Private Sub txtEndingColor_Change()
    
    If Len(txtEndingColor.Text) = 4 Then
        If CLng("&H" & Right$(txtEndingColor, 2) & Left$(txtEndingColor.Text, 2)) > IntMax Then
            txtEndingColor.Text = "FF7F"
        End If
        lblEndingColor.BackColor = "&H" & GB2RGB(txtEndingColor.Text, True)
    ElseIf LenB(txtEndingColor.Text) = 0 Then
        lblEndingColor.BackColor = vbWhite
    End If
    
    Call CreateCheck
    
End Sub

Private Sub txtEndingColor_KeyPress(KeyCode As Integer)
    Call txtStartingColor_KeyPress(KeyCode)
End Sub

Private Sub txtFrom_Change()
    
    If LenB(txtFrom.Text) > 0 Then
        If Val(txtFrom.Text) = 0 Then txtFrom.Text = 1
        txtTo.Text = Val(txtFrom.Text) + (Val(txtColorNumber.Text) - 1)
    Else
        txtTo.Text = vbNullString
    End If
   
      
End Sub

Private Sub txtFrom_KeyPress(KeyCode As Integer)
    Call txtColorNumber_KeyPress(KeyCode)
End Sub

Private Sub txtFrom_KeyUp(KeyCode As Integer, Shift As Integer)

    If Val(txtFrom.Text) > 16 - (Val(txtColorNumber.Text) - 1) Then
        txtFrom.Text = 16 - (Val(txtColorNumber.Text) - 1)
    ElseIf Val(txtFrom.Text) < 1 Then
        txtFrom.Text = 1
    End If

End Sub

Private Sub txtStartingColor_Change()
    
    If Len(txtStartingColor.Text) = 4 Then
        If CLng("&H" & Right$(txtStartingColor.Text, 2) & Left$(txtStartingColor.Text, 2)) > IntMax Then
            txtStartingColor.Text = "FF7F"
        End If
        lblStartingColor.BackColor = "&H" & GB2RGB(txtStartingColor.Text, True)
    ElseIf LenB(txtStartingColor.Text) = 0 Then
        lblStartingColor.BackColor = vbWhite
    End If
    
    Call CreateCheck
    
End Sub

Private Sub txtStartingColor_KeyPress(KeyCode As Integer)
    
    If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
        Select Case KeyCode
            Case vbKey0 To vbKey9
            Case vbKeyA To vbKeyF
            Case 97 To 102
                KeyCode = KeyCode - 32
            Case Else
                KeyCode = 0: Exit Sub
        End Select
    End If
        
End Sub
