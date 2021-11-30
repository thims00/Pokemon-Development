VERSION 5.00
Begin VB.Form frmCalc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RGB/GBA Converter"
   ClientHeight    =   1710
   ClientLeft      =   3615
   ClientTop       =   3120
   ClientWidth     =   5670
   Icon            =   "frmCalc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Tag             =   "4000"
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   4120
      TabIndex        =   3
      Tag             =   "4005"
      Top             =   915
      Width           =   1155
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4120
      TabIndex        =   2
      Tag             =   "4004"
      Top             =   465
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Tag             =   "4001"
      Top             =   120
      Width           =   5415
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   360
         ScaleHeight     =   615
         ScaleWidth      =   1575
         TabIndex        =   9
         Top             =   480
         Width           =   1575
         Begin VB.OptionButton optRGB2GBA 
            Caption         =   "RGB to GBA"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Tag             =   "4002"
            Top             =   0
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optGBA2RGB 
            Caption         =   "GBA to RGB"
            Height          =   255
            Left            =   0
            TabIndex        =   5
            Tag             =   "4003"
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.TextBox txtGBAColor 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtRGBColor 
         Height          =   285
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GBA"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   885
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "RGB"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   525
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalc_Click()
    If optRGB2GBA Then
        txtGBAColor.Text = RGB2GBA(txtRGBColor.Text) 'Convert from RBG to GBA
    Else
        txtRGBColor.Text = GB2RGB(txtGBAColor.Text) 'Convert from GBA to RGB
    End If
    
    cmdCopy.SetFocus
    
End Sub

Private Sub cmdCopy_Click()
    If optRGB2GBA Then
        Clipboard.Clear
        Clipboard.SetText txtGBAColor.Text
    Else
        Clipboard.Clear
        Clipboard.SetText txtRGBColor.Text
    End If
End Sub

Private Sub Form_Load()
    LoadResStrings Me 'Load localized strings
End Sub

Private Sub optGBA2RGB_Click()
    txtGBAColor.SetFocus
    txtGBAColor.Locked = False
    txtRGBColor.Locked = True
    txtRGBColor.Text = vbNullString
End Sub

Private Sub optRGB2GBA_Click()
    txtRGBColor.SetFocus
    txtRGBColor.Locked = False
    txtGBAColor.Locked = True
    txtGBAColor.Text = vbNullString
End Sub

Private Sub txtGBAColor_Change()
    If optGBA2RGB And Len(Trim$(txtGBAColor.Text)) = 4 Then
        cmdCalc.Enabled = True
        cmdCalc.SetFocus
    ElseIf optGBA2RGB And Len(Trim$(txtGBAColor.Text)) < 4 Then
        cmdCalc.Enabled = False
    End If
End Sub

Private Sub txtGBAColor_KeyPress(KeyCode As Integer)
    If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
        If (KeyCode > 64 And KeyCode < vbKeyG) Then Exit Sub
        If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub
        If (KeyCode > 47 And KeyCode < 58) Then Exit Sub
        KeyCode = 0
    End If
End Sub

Private Sub txtRGBColor_Change()
    If optRGB2GBA And Len(Trim$(txtRGBColor.Text)) = 6 Then
        cmdCalc.Enabled = True
        cmdCalc.SetFocus
    ElseIf optRGB2GBA And Len(Trim$(txtRGBColor.Text)) < 6 Then
        cmdCalc.Enabled = False
    End If
End Sub

Private Sub txtRGBColor_KeyPress(KeyCode As Integer)
    Call txtGBAColor_KeyPress(KeyCode)
End Sub
