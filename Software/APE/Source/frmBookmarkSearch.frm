VERSION 5.00
Begin VB.Form frmBookmarkSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bookmark Search"
   ClientHeight    =   570
   ClientLeft      =   13605
   ClientTop       =   3930
   ClientWidth     =   2595
   Icon            =   "frmBookmarkSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "6000"
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   150
      MaxLength       =   25
      TabIndex        =   0
      Top             =   143
      Width           =   2295
   End
End
Attribute VB_Name = "frmBookmarkSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    LoadResStrings Me
End Sub

Private Sub txtSearch_Change()
    Call frmBookmarks.FindItem(frmBookmarks.lstBookmark, txtSearch.Text)
End Sub

Private Sub txtSearch_KeyPress(KeyCode As Integer)
    If KeyCode = 59 Or KeyCode = 61 Or KeyCode = 91 Or KeyCode = 93 Then KeyCode = 0
End Sub
