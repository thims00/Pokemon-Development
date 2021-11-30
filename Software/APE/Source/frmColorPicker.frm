VERSION 5.00
Begin VB.Form frmColorPicker 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "APE's Color Picker"
   ClientHeight    =   5175
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmColorPicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   623
   StartUpPosition =   2  'CenterScreen
   Tag             =   "3000"
   Begin VB.ComboBox cboPercent 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmColorPicker.frx":000C
      Left            =   6960
      List            =   "frmColorPicker.frx":0030
      TabIndex        =   16
      Text            =   "200%"
      Top             =   3210
      Width           =   975
   End
   Begin VB.TextBox txtCoords 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   3210
      Width           =   1290
   End
   Begin VB.Timer tmrTransfer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3120
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Color History"
      Height          =   615
      Left            =   180
      TabIndex        =   31
      Tag             =   "3008"
      Top             =   4440
      Width           =   3885
      Begin VB.Label lblHistoryColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   3180
         TabIndex        =   37
         Tag             =   "blank"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblHistoryColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   2580
         TabIndex        =   36
         Tag             =   "blank"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblHistoryColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   1980
         TabIndex        =   35
         Tag             =   "blank"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblHistoryColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1380
         TabIndex        =   34
         Tag             =   "blank"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblHistoryColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   780
         TabIndex        =   33
         Tag             =   "blank"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblHistoryColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   32
         Tag             =   "blank"
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Timer tmrThinBox 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   4200
      Top             =   4440
   End
   Begin VB.PictureBox picNewColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4785
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   30
      Top             =   495
      Width           =   765
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   2295
      Width           =   375
   End
   Begin VB.CommandButton cmdLess 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtGBAColor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   13
      Top             =   4320
      Width           =   615
   End
   Begin VB.PictureBox picPreView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   8400
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   27
      Top             =   480
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox picBigBox 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3870
      Left            =   180
      MouseIcon       =   "frmColorPicker.frx":0072
      MousePointer    =   99  'Custom
      ScaleHeight     =   256
      ScaleMode       =   0  'User
      ScaleWidth      =   256
      TabIndex        =   21
      Top             =   480
      Width           =   3885
      Begin VB.Image imgMarker 
         Height          =   165
         Left            =   1920
         Picture         =   "frmColorPicker.frx":01C4
         Top             =   1800
         Width           =   165
      End
   End
   Begin VB.PictureBox picThinBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      FillColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   3870
      Left            =   4260
      ScaleHeight     =   254
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   20
      Top             =   480
      Width           =   315
   End
   Begin VB.TextBox txtHexColor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      MaxLength       =   6
      TabIndex        =   12
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   11
      Top             =   3525
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   9
      Top             =   3165
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2805
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   5
      Top             =   2370
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   3
      Top             =   2010
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   5280
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1650
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      Caption         =   "B:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   10
      Top             =   3525
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      Caption         =   "G:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   4800
      TabIndex        =   8
      Top             =   3165
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      Caption         =   "R:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4800
      TabIndex        =   6
      Top             =   2805
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      Caption         =   "B:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   4
      Top             =   2370
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      Caption         =   "S:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   2
      Top             =   2010
      Width           =   495
   End
   Begin VB.OptionButton objOption 
      Caption         =   "H:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   0
      Top             =   1650
      Width           =   495
   End
   Begin VB.Timer tmrScreenColor 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8400
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8760
      Top             =   480
   End
   Begin VB.PictureBox picZoom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   6960
      ScaleHeight     =   148
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   148
      TabIndex        =   38
      Top             =   960
      Width           =   2250
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      Height          =   510
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   3705
      Width           =   510
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   510
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   3705
      Width           =   510
   End
   Begin VB.Image imgMouse3 
      Height          =   480
      Left            =   7320
      Picture         =   "frmColorPicker.frx":0210
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMouse2 
      Height          =   480
      Left            =   6840
      Picture         =   "frmColorPicker.frx":04C3
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMouse1 
      Height          =   480
      Left            =   7800
      Picture         =   "frmColorPicker.frx":0754
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWhiteMarker 
      Height          =   165
      Left            =   3840
      Picture         =   "frmColorPicker.frx":0A5E
      Top             =   240
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image imgBlackMarker 
      Height          =   165
      Left            =   3600
      Picture         =   "frmColorPicker.frx":0AAA
      Top             =   240
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Line linPoint 
      BorderColor     =   &H00808080&
      X1              =   305
      X2              =   306
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Line linTriang2Falling 
      BorderColor     =   &H00808080&
      X1              =   306
      X2              =   311
      Y1              =   256
      Y2              =   261
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   29
      Tag             =   "3007"
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "GBA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4860
      TabIndex        =   28
      Top             =   4365
      Width           =   300
   End
   Begin VB.Line linTriang2Rising 
      BorderColor     =   &H00808080&
      X1              =   306
      X2              =   311
      Y1              =   253
      Y2              =   248
   End
   Begin VB.Line linTriang2Vert 
      BorderColor     =   &H00808080&
      X1              =   310
      X2              =   310
      Y1              =   250
      Y2              =   260
   End
   Begin VB.Line linTriang1Falling 
      BorderColor     =   &H00808080&
      X1              =   277
      X2              =   283
      Y1              =   251
      Y2              =   256
   End
   Begin VB.Line linTriang1Rising 
      BorderColor     =   &H00808080&
      X1              =   277
      X2              =   283
      Y1              =   261
      Y2              =   256
   End
   Begin VB.Line linTriang1Vert 
      BorderColor     =   &H00808080&
      X1              =   276
      X2              =   276
      Y1              =   251
      Y2              =   261
   End
   Begin VB.Label lblComplementaryColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " C.C."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   26
      Top             =   840
      Width           =   405
   End
   Begin VB.Label lblSuffix 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5805
      TabIndex        =   24
      Top             =   2415
      Width           =   180
   End
   Begin VB.Label lblSuffix 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5805
      TabIndex        =   23
      Top             =   2055
      Width           =   180
   End
   Begin VB.Label lblSuffix 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "°"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   5820
      TabIndex        =   22
      Top             =   1650
      Width           =   105
   End
   Begin VB.Label lblOldColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   4785
      TabIndex        =   19
      Top             =   990
      Width           =   765
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   5595
      TabIndex        =   18
      Top             =   1725
      Width           =   15
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5040
      TabIndex        =   17
      Top             =   4005
      Width           =   135
   End
   Begin VB.Label lblContainer 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4770
      TabIndex        =   25
      Top             =   480
      Width           =   810
   End
   Begin VB.Image imgEyedropper 
      Height          =   480
      Left            =   7575
      Picture         =   "frmColorPicker.frx":0AF6
      Top             =   3720
      Width           =   480
   End
   Begin VB.Image imgMagnifier 
      Height          =   480
      Left            =   8175
      Picture         =   "frmColorPicker.frx":0D87
      Top             =   3720
      Width           =   480
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      HelpContextID   =   3001
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         HelpContextID   =   3002
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      HelpContextID   =   3003
      Begin VB.Menu mnuAutoCopy 
         Caption         =   "Auto Copy"
         Checked         =   -1  'True
         HelpContextID   =   3004
      End
      Begin VB.Menu mnuAutoPaste 
         Caption         =   "Auto Paste"
         Checked         =   -1  'True
         HelpContextID   =   3005
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearHistory 
         Caption         =   "Clear Color History"
         HelpContextID   =   3006
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Byte
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type


Private Type HSL
    Hue As Integer
    Saturation As Byte
    Luminance As Byte
End Type

Private Const HORZRES = 8
Private Const VERTRES = 10
Private Const SRCCOPY = &HCC0020
Dim poiMouse As POINTAPI
Const CircleRay As Byte = 5
Const bytMove As Byte = 28
Dim blnDrag As Boolean, intSystemColorAngleMax1530 As Integer, bteSaturationMax255 As Byte, bteBrightnessMax255 As Byte, lngcolor As Long
Dim mSngRValue As Single, mSngGValue As Single, mSngBValue As Single
Dim blnNotFirstTimeMarker As Boolean, mBteMarkerOldX As Integer, mBteMarkerOldY As Integer
Dim mBlnRecentThinBoxPress As Boolean, mBlnBigBoxReady As Boolean, i As Byte, X As Byte
Dim HistoryIndex As Byte, blnIgnoreHistory As Boolean, blnIgnoreColorChange As Boolean
Dim blnFirstTimeEyeDropper As Boolean

Private Sub TriangleMove(Y)

    Y = Round(Y, 0)
    ' First arrow
    linTriang1Vert.Y1 = Y + bytMove + 1: linTriang1Vert.Y2 = Y + bytMove + 10
    linTriang1Rising.Y1 = Y + bytMove + 10: linTriang1Rising.Y2 = Y + bytMove + 4
    linTriang1Falling.Y1 = Y + bytMove: linTriang1Falling.Y2 = Y + bytMove + 6
    ' Second arrow
    linTriang2Vert.Y1 = Y + bytMove + 1: linTriang2Vert.Y2 = Y + bytMove + 10
    linTriang2Rising.Y2 = Y + bytMove + 11: linTriang2Rising.Y1 = Y + bytMove + 6
    linTriang2Falling.Y2 = Y + bytMove - 1: linTriang2Falling.Y1 = Y + bytMove + 4
    linPoint.Y1 = Y + bytMove + 5: linPoint.Y2 = Y + bytMove + 5

End Sub


Public Sub MoveMarker(X, Y)

    imgMarker.Visible = True
    If bteBrightnessMax255 < 200 Then   ' White marker if the surroundings are grey.
        imgMarker.Picture = imgWhiteMarker.Picture
        imgMarker.Left = X - CircleRay: imgMarker.Top = Y - CircleRay
        Exit Sub
    End If
    If Text1(0) < 26 Or Text1(0) > 200 Then   ' Shades of blue.
        If bteSaturationMax255 > 70 Then      ' And bteSaturationMax255 < 150 Then 'White marker if the surroundings are grey..
            imgMarker.Picture = imgWhiteMarker.Picture
            imgMarker.Left = X - CircleRay: imgMarker.Top = Y - CircleRay
            Exit Sub
        End If
    End If
    imgMarker.Picture = imgBlackMarker.Picture
    imgMarker.Left = X - CircleRay: imgMarker.Top = Y - CircleRay

End Sub


Public Sub opt3RedPaintPicThinBox(ByVal G, b)
    Dim bteX As Byte, intCtr As Integer

    ' Painting picThinBox (19, 255)
    For bteX = 0 To 19
        For intCtr = 0 To 255
            SetPixelV picThinBox.hdc, bteX, intCtr, RGB(255 - intCtr, G, b)   ' Painting with API.
        Next intCtr
    Next bteX

End Sub


Public Sub opt4GreenPaintPicThinBox(ByVal R, b)
    Dim bteX As Byte, intCtr As Integer

    ' Painting picThinBox (19, 255)
    For bteX = 0 To 19
        For intCtr = 0 To 255
            SetPixelV picThinBox.hdc, bteX, intCtr, RGB(R, 255 - intCtr, b)   ' Painting by API.
        Next intCtr
    Next bteX

End Sub


Public Sub opt5BluePaintPicThinBox(ByVal R, G)
    Dim bteX As Byte, intCtr As Integer

    ' Painting picThinBox (19, 255)
    For bteX = 0 To 19
        For intCtr = 0 To 255
            SetPixelV picThinBox.hdc, bteX, intCtr, RGB(R, G, 255 - intCtr)   ' Painting by API.
        Next intCtr
    Next bteX

End Sub


Public Sub BigBoxOpt3Reaction(ByVal X, Y)
    Dim udtAngelSaturationBrightness As HSL

    If Y < 254 Then picNewColor.BackColor = picBigBox.Point(X, Y)
    Call SplitlblNewColorToRGBboxes   ' Updating the module global mSngRValue etc.
    Call opt3RedPaintPicThinBox(ByVal mSngGValue, mSngBValue)
    picThinBox.Refresh
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True)   ' Causes an update of all the constants.

End Sub


Public Sub BigBoxOpt4Reaction(ByVal X, Y)
    Dim udtAngelSaturationBrightness As HSL

    If Y < 254 Then picNewColor.BackColor = picBigBox.Point(X, Y)
    Call SplitlblNewColorToRGBboxes   ' Updating the module global mSngRValue etc.
    Call opt4GreenPaintPicThinBox(ByVal mSngRValue, mSngBValue)
    picThinBox.Refresh
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True)   ' Causes an update of all the constants. The letter of three stands for the RED-box.

End Sub


Public Sub BigBoxOpt5Reaction(ByVal X, Y)
    Dim udtAngelSaturationBrightness As HSL

    If Y < 254 Then picNewColor.BackColor = picBigBox.Point(X, Y)
    Call SplitlblNewColorToRGBboxes   ' Updating the module global mSngRValue etc.
    Call opt5BluePaintPicThinBox(ByVal mSngRValue, mSngGValue)
    picThinBox.Refresh
    udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True)   ' Causes an update of all the constants. The letter of three stands for the RED-box..

End Sub


Private Sub opt3RedPaintPicBigBox()
    Dim R As Single, G As Single, b As Single

    ' Paint the picBigBox
    R = Text1(3)  ' Red
    For b = 255 To 0 Step -1
        For G = 255 To 0 Step -1   ' Interesting if there is an error, thus a jump directly to EndSub.
            SetPixelV picBigBox.hdc, b, 255 - G, RGB(R, G, b)   ' Painting by API.
        Next G
    ' G = G - 1 'Because that G becomes too big when the loop has finished
    Next b

End Sub


Private Sub opt4GreenPaintPicBigBox()
    Dim R As Single, G As Single, b As Single

    ' Paint picBigBox
    G = Text1(4)  ' Green
    For b = 255 To 0 Step -1
        For R = 255 To 0 Step -1   ' Interesting if there is an error, thus a jump directly to EndSub.
            SetPixelV picBigBox.hdc, b, 255 - R, RGB(R, G, b)   ' Painting by API.
        Next R
        R = R - 1   ' Because that R becomes too big when the loop has finished
    Next b

End Sub


Private Sub opt5BluePaintPicBigBox()
    Dim R As Single, G As Single, b As Single

    ' Paint picBigBox
    b = Text1(5)   ' Blue
    For R = 255 To 0 Step -1
        For G = 255 To 0 Step -1  ' Interesting if there is an error, thus a jump directly to EndSub.
            SetPixelV picBigBox.hdc, R, 255 - G, RGB(R, G, b)   ' Ritar medelst API.
        Next G
        G = G - 1 ' Because that G becomes too big when the loop has finished
    Next R

End Sub


Public Sub LoadSettings()
    Dim FileNum As Integer, strWidth As String, strPercent As String, strAutoCopy As String
    Dim strColor As String, strAutoPaste As String, strHistoryIndex As String, arrColors(5) As String

    If Not FileExists(App.Path & "\Colorpicker.dat") Then Exit Sub
    
    ' Read from the dat file of the form.
    FileNum = FreeFile
    
    On Error GoTo continue:
    Open App.Path & "\Colorpicker.dat" For Input As #FileNum
        Line Input #FileNum, strColor
        If IsNumeric(Trim$(strColor)) Then lblOldColor.BackColor = strColor
        Line Input #FileNum, strWidth
        Line Input #FileNum, strPercent
        Line Input #FileNum, strAutoCopy
        Line Input #FileNum, strAutoPaste
        Line Input #FileNum, strHistoryIndex
        For X = lblHistoryColor.LBound To lblHistoryColor.UBound
            Line Input #FileNum, arrColors(X)
        Next
    Close #FileNum
    
    If IsNumeric(Trim$(strWidth)) Then Me.Width = Val(strWidth)
    If IsNumeric(Trim$(strPercent)) Then cboPercent.ListIndex = Val(strPercent)
    If LenB(strAutoCopy) > 0 Then mnuAutoCopy.Checked = strAutoCopy
    If LenB(strAutoPaste) > 0 Then mnuAutoPaste.Checked = strAutoPaste
    If LenB(strHistoryIndex) > 0 Then HistoryIndex = CByte(strHistoryIndex)
    For X = lblHistoryColor.LBound To lblHistoryColor.UBound
        If Mid$(arrColors(X), 2, 5) <> "blank" Then
            lblHistoryColor(X).BackColor = Val(arrColors(X))
            lblHistoryColor(X).Tag = Val(arrColors(X))
        Else
            lblHistoryColor(X).BackColor = vbWhite
            lblHistoryColor(X).Tag = "blank"
        End If
    Next
    GoTo finish
    
continue:
    Close #FileNum
    Me.Width = 9690
    cboPercent.ListIndex = 1
    mnuAutoCopy.Checked = vbChecked
    mnuAutoPaste.Checked = vbChecked
    HistoryIndex = 0
    
On Error Resume Next
    
finish:
    If FileExists(App.Path & "\scr.bmp") Then picZoom.Picture = LoadPicture(App.Path & "\scr.bmp")
End Sub


Private Sub ArrowsModeDepending()

    ' Adjusting imgArrows depending on the current mode.
    If objOption(0) Then
        TriangleMove (255 - (intSystemColorAngleMax1530 / 1530 * 255))
    ElseIf objOption(1) Then
        TriangleMove (255 - (Text1(1) * 2.55))
    ElseIf objOption(2) Then
        TriangleMove (255 - (Text1(2) * 2.55))
    Else
        TriangleMove (255 - Text1(3))
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, _
        Shift As Integer, _
        X As Single, _
        Y As Single)

    Call ReleaseCapture
    Call SendMessage(Me.hwnd, &HA1, 2, 0&)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim FileNum As Integer
    
    ' Opening to write to the ini-file of the form.
    FileNum = FreeFile
    If FileExists(App.Path & "\Colorpicker.dat") Then SetAttr App.Path & "\Colorpicker.dat", vbNormal
    Open App.Path & "\Colorpicker.dat" For Output As #FileNum
    ' Writing
    Write #FileNum, picNewColor.BackColor   ' Becames lblOld color the next time the program starts.
    If Me.Width < 6975 Then Write #FileNum, 6975 Else Write #FileNum, Me.Width
    Write #FileNum, cboPercent.ListIndex
    Write #FileNum, mnuAutoCopy.Checked
    Write #FileNum, mnuAutoPaste.Checked
    Write #FileNum, HistoryIndex
    
    For X = lblHistoryColor.LBound To lblHistoryColor.UBound
        If lblHistoryColor(X).Tag <> "blank" Then
            Write #FileNum, Val(lblHistoryColor(X).Tag)
        Else
            Write #FileNum, "blank"
        End If
    Next
    Close #FileNum

End Sub


Private Sub imgEyedropper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Me.Hide
    ' frmMain.WindowState = vbMinimized
    Me.MousePointer = vbCustom
    Me.MouseIcon = imgMouse1.Picture
    imgEyedropper.Picture = LoadPicture(vbNullString)
    ' Me.Left = (Screen.Width - 2500) / 2
    ' Me.Top = (Screen.Height - 4000) / 2
    ' Me.Height = 4000
    ' Me.Width = 2500
    ' Frame2.Left = 0
    ' Frame2.Top = 0
    ' picBigBox.Visible = False
    ' Frame1.Visible = False
    ' Label4.Visible = False
    ' mnufile.Visible = False
    ' mnuOptions.Visible = False
    ' cboPercent.Enabled = False
    ' Me.Show
    Call picBigBox_Colorize
    tmrScreenColor.Enabled = True
    Call HookKeyboard
End Sub

Private Sub imgEyedropper_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Me.Hide
    Me.MousePointer = vbDefault
    imgEyedropper.Picture = imgMouse2.Picture
    tmrScreenColor.Enabled = False
    ' Me.Left = (Screen.Width - 9690) / 2
    ' Me.Top = (Screen.Height - 5970) / 2
    ' Me.Width = 9690
    ' Me.Height = 5970
    ' Frame2.Left = 464
    ' Frame2.Top = 64.533
    ' picBigBox.Visible = True
    ' Frame1.Visible = True
    ' Label4.Visible = True
    ' mnufile.Visible = True
    ' mnuOptions.Visible = True
    ' Me.Show
    cboPercent.Enabled = True
    mBlnBigBoxReady = False
    'Call picBigBox_Colorize
    Call objOption_Click(0)
    Call UnhookKeyboard
End Sub

Private Sub imgMagnifier_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    cboPercent.Enabled = False
    Me.MousePointer = vbCustom
    Me.MouseIcon = imgMouse1.Picture
    imgMagnifier.Picture = LoadPicture(vbNullString)
    Call HookKeyboard
End Sub

Private Sub imgMagnifier_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = False
    cboPercent.Enabled = True
    SavePicture picZoom.Image, App.Path & "\scr.bmp"
    Me.MousePointer = vbDefault
    imgMagnifier.Picture = imgMouse3.Picture
    Call UnhookKeyboard
End Sub

Private Sub lblComplementaryColor_Click(Index As Integer)
    If Text1(0) < 180 Then Text1(0) = Text1(0) + 180 Else Text1(0) = Text1(0) - 180
    Call Text1_LostFocus(0)
End Sub

Private Sub PaintThinBox(Index As Integer)

    Select Case Index
        Case 0
            Call RainBowThinBox
        Case 1
            Call FadeThinBoxToGrey
        Case 2
            picThinBox.BackColor = HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal bteSaturationMax255, ByVal 255, False)   ' Delivers a lighter shade of the active color. 'Setting the whole square for easy fading.
            Call FadeThinBoxToBlack
    End Select
    picThinBox.Refresh

End Sub

Public Sub picBigBox_Colorize()
    Select Case objOption(i)
        Case i = 0: Call Bigbox3D
        Case i = 1: Call RainBowBigbox(False, True)
        Case i = 2: Call RainBowBigbox(True, False)
        Case i = 3: Call opt3RedPaintPicBigBox
        Case i = 4: Call opt4GreenPaintPicBigBox
        Case i = 5: Call opt5BluePaintPicBigBox
    End Select
    
    If blnNotFirstTimeMarker = True Then                  ' IN CASE THERE IS A marker-coordinate...
        Call MoveMarker(mBteMarkerOldX, mBteMarkerOldY)   ' REPAINT THE MARKER (if there is any).
        picNewColor.BackColor = picBigBox.Point(mBteMarkerOldX, mBteMarkerOldY)
    End If
    
    If mBlnBigBoxReady = False Then   ' PLACES A MARKER AT CORRECT LOCATION AT THE SETUP STAGE.
        blnNotFirstTimeMarker = True
        ' MODE DEPENDING NEW MARKER POSITION.
        If objOption(0) Then mBteMarkerOldX = bteSaturationMax255:   mBteMarkerOldY = 255 - bteBrightnessMax255            ' Transmitting logical values.
        If objOption(1) Then mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteBrightnessMax255   ' Transmitting logical values.
        If objOption(2) Then mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteSaturationMax255   ' Transmitting logical values.
        Call MoveMarker(mBteMarkerOldX, mBteMarkerOldY)                                                                    ' REPAINT THE MARKER (if there is any).
        mBlnBigBoxReady = True                                                                                             ' NOW AT LEAST THE FIRST SPONTANEOUS REDRAW HAS FINISHED.
    End If
End Sub


Public Function HSLToRGB(ByVal intLocalColorAngle As Integer, ByVal Saturation As Long, ByVal Luminance As Long, ByVal blnUpdateTextBoxes As Boolean) As Long
    Dim R As Long, G As Long, b As Long, lMax As Byte, lMid As Byte, lMin As Long, q As Single

    lMax = Luminance
    lMin = (255 - Saturation) * lMax / 255   ' 255 - (Saturation * lMax / 255)
    q = (lMax - lMin) / 255
    Select Case intLocalColorAngle
        Case 0 To 255
            lMid = (intLocalColorAngle - 0) * q + lMin
            R = lMax: G = lMid: b = lMin
        Case 256 To 510  ' This period surpasses the node border with one unit - over to gren color
            lMid = -(intLocalColorAngle - 255) * q + lMax   ' -(intLocalColorAngle - 256) * q + lMin
            R = lMid: G = lMax: b = lMin
        Case 511 To 765
            lMid = (intLocalColorAngle - 510) * q + lMin
            R = lMin: G = lMax: b = lMid
        Case 766 To 1020
            lMid = -(intLocalColorAngle - 765) * q + lMax
            R = lMin: G = lMid: b = lMax
        Case 1021 To 1275
            lMid = (intLocalColorAngle - 1020) * q + lMin
            R = lMid: G = lMin: b = lMax
        Case 1276 To 1530
            lMid = -(intLocalColorAngle - 1275) * q + lMax
            R = lMax: G = lMin: b = lMid
        Case Else
            MsgBox LoadResString(3009) & Str(intLocalColorAngle)
    End Select
    mSngRValue = R: mSngGValue = G: mSngBValue = b   ' Updating the sustem constants automatically. Perhaps must exclude this to give them protection.
    HSLToRGB = b * &H10000 + G * &H100& + R          ' Delivers lngColor in VB-format.
    If blnUpdateTextBoxes = True Then                ' Then the calling routine is not any of the complex automatic routines for fading etc.
        ' Since this is a single time called session I can safely update my system constants and convert my hifgh resolution system constants to textbox dito.
        Text1(0) = Round(intLocalColorAngle / 255 / 6 * 360)
        Text1(1) = Round(Saturation / 255 * 100)
        Text1(2) = Round(Luminance / 255 * 100)
        Text1(3) = mSngRValue
        Text1(4) = mSngGValue
        Text1(5) = mSngBValue
        ' txtHexColor = Hex$(mSngRValue * 65536 + mSngGValue * 256 + mSngBValue): txtHexColor.Refresh 'Applies to internetstandard<>VBstandard
        If mSngRValue < &H10 Then
            txtHexColor = Right$("00000" & Hex$(mSngRValue * 65536 + mSngGValue * 256 + mSngBValue), 6)   ' Padding with zeroletters to the left.
        Else
            txtHexColor = Hex$(mSngRValue * 65536 + mSngGValue * 256 + mSngBValue)
        End If
        picNewColor.BackColor = HSLToRGB
        intSystemColorAngleMax1530 = intLocalColorAngle   ' Sometims there is only a mouse Y coordinate
        bteSaturationMax255 = Saturation
        bteBrightnessMax255 = Luminance
    End If

End Function


Public Sub SplitlblNewColorToRGBboxes()   ' Updating the system constants and textboxes regarding to RGB.
    mSngRValue = picNewColor.BackColor And &HFF: Text1(3) = mSngRValue
    mSngGValue = (picNewColor.BackColor And &HFF00&) \ &H100&: Text1(4) = mSngGValue
    mSngBValue = (picNewColor.BackColor And &HFF0000) \ &H10000: Text1(5) = mSngBValue
End Sub

Private Function RGBToHSL201(ByVal RGBValue As Long, ByVal blnUpdateTextBoxes As Boolean) As HSL
    Dim R As Long, G As Long, b As Long
    Dim lMax As Long, lMin As Long, lDiff As Long, lSum As Long

    R = RGBValue And &HFF&
    G = (RGBValue And &HFF00&) \ &H100&
    b = (RGBValue And &HFF0000) \ &H10000
    
    If R > G Then lMax = R: lMin = G Else lMax = G: lMin = R   ' Finds the Superior and inferior components.
    If b > lMax Then lMax = b Else If b < lMin Then lMin = b
    
    lDiff = lMax - lMin
    lSum = lMax + lMin
    
    ' Luminance, thus brightness
    RGBToHSL201.Luminance = lMax / 255 * 100
    
    ' Saturation
    If lMax <> 0 Then   ' Protecting from the impossible operation of division by zero.
        RGBToHSL201.Saturation = 100 * lDiff / lMax   ' The logic of Adobe Photoshops is this simple.
    Else
        RGBToHSL201.Saturation = 0
    End If
    
    ' Hue
    Dim q As Single
    
    If lDiff = 0 Then q = 0 Else q = 60 / lDiff   ' Protecting from the impossible operation of division by zero.
    
    Select Case lMax
        Case R
            If G < b Then
                RGBToHSL201.Hue = 360& + q * (G - b)
                intSystemColorAngleMax1530 = (360& + q * (G - b)) * 4.25   ' Converting from degrees to my resolution of detail.
            Else
                RGBToHSL201.Hue = q * (G - b)
                intSystemColorAngleMax1530 = (q * (G - b)) * 4.25
            End If
        Case G
            RGBToHSL201.Hue = 120& + q * (b - R)
            intSystemColorAngleMax1530 = (120& + q * (b - R)) * 4.25
        Case b
            RGBToHSL201.Hue = 240& + q * (R - G)
            intSystemColorAngleMax1530 = (240& + q * (R - G)) * 4.25
    End Select
    
    If blnUpdateTextBoxes = True Then
        If R < &H10 Then
            txtHexColor = Right$("00000" & Hex$(R * 65536 + G * 256 + b), 6)   ' Adds letters of zero to the left which is a necessary so called padding.
        Else
            txtHexColor = Hex$(R * 65536 + G * 256 + b)
        End If
        
        Text1(0) = Int(intSystemColorAngleMax1530 / 1530 * 360)
        
        If lMax = 0 Then
            bteSaturationMax255 = 0   ' Protecting from the impossible operation of division by zero.
        Else
            bteSaturationMax255 = 255 * lDiff / lMax
            Text1(1) = RGBToHSL201.Saturation   ' = saturation both 0 To 255 and 0 To 100%.
        End If
        
        bteBrightnessMax255 = lMax: Text1(2) = RGBToHSL201.Luminance   ' =Brighness both 0 To 255 and 0 To 100%.
        txtGBAColor.Text = RGB2GBA(txtHexColor.Text)
    End If

End Function

Public Sub NudgeHueValue(ByVal intDirektion)
    ' 1530 levels. The triangels are moving every sixth step and are lying on the byte level of 1530/6.
    ' RGBtxtboxes tells the nudge level:
    Dim lngcolor As Long
    ' NudgeValue goes from ZERO to 1536.
    intSystemColorAngleMax1530 = intSystemColorAngleMax1530 + intDirektion     ' Calculating the new value of intSystemColorAngleMax1530, thus +1 or -1.
    If intSystemColorAngleMax1530 > 1530 Then intSystemColorAngleMax1530 = 1530  ' Limiter.
    If intSystemColorAngleMax1530 < 0 Then intSystemColorAngleMax1530 = 0
    lngcolor = HSLToRGB(ByVal intSystemColorAngleMax1530, bteSaturationMax255, bteBrightnessMax255, True)   ' lngColor as a function of HSLToRGB. System constants are being updated at the same time.
    Call TriangleMove(255 - (intSystemColorAngleMax1530 / 1530 * 255))   ' Moving the triangle
End Sub

Public Sub FadeThinBoxToGrey()
    Dim sng255saturation As Single, sngLokalBrightness As Single, X As Byte, Y As Integer   ' , YCtr As Integer

    sng255saturation = 255: sngLokalBrightness = bteBrightnessMax255
    For X = 0 To 19
        Y = 0  ' Sets YCtr for making a new countdown.
        Do     ' Interesting if there would raise an error, thus a leap directly to EndSub.
            SetPixelV picThinBox.hdc, X, Y, HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal Round(sng255saturation - sng255saturation * Y / 255), ByVal sngLokalBrightness, False)
            Y = Y + 1
        Loop While Y < 256  ' Because Y gets to big when the loop has finished.
    Next X

End Sub


Public Sub Bigbox3D()
    Dim sngLokalSaturation As Single, sngLokalBrightness As Single, YRADNOLL As Integer
    Dim sngR256delToBlack As Single, sngG256delToBlack As Single, sngB256delToBlack As Single
    Dim R As Single, G As Single, b As Single, lColor As Long, Y As Integer, X As Integer

    sngLokalSaturation = 255: sngLokalBrightness = 255   ' There is a need for intense start color.
    ' ********* Firstly a single fade from saturated to grey on the uppermost row.
    For X = 0 To 255
        SetPixelV picBigBox.hdc, X, YRADNOLL, HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal Round(sngLokalSaturation * X / 255), ByVal sngLokalBrightness, False)
    Next X                         ' Resets Y for a new row.
    ' ********* Here will be an FADE TO BLACK for all columns ********
    For X = 255 To 0 Step -1
        ' If blnVertical = True Then R = Ro: G = Go: B = Bo ' If line is vertical the reset for a new round.
        lColor = picBigBox.Point(X, 0)   ' Reading the uppermost pixel which is to be faded.
        R = lColor And &HFF
        G = (lColor And &HFF00&) \ &H100&
        b = (lColor And &HFF0000) \ &H10000
        sngR256delToBlack = R / 255   ' The fraction blocks which lead down to black.
        sngG256delToBlack = G / 255
        sngB256delToBlack = b / 255
        For Y = 0 To 255                                  ' Interesting if there would raise an error, thus a leap back to EndSub.
            SetPixelV picBigBox.hdc, X, Y, RGB(R, G, b)   ' Painting with API.
            R = R - sngR256delToBlack                     ' Darkening the shade one of a 256:th.
            G = G - sngG256delToBlack
            b = b - sngB256delToBlack
        Next Y
        Y = Y - 1  ' Because that Y gets too big when the loop is completed.
    Next X

End Sub


Public Sub FadeThinBoxToBlack()
    Dim sngR256delToBlack As Single, sngG256delToBlack As Single, sngB256delToBlack As Single
    Dim R As Single, G As Single, b As Single, lColor As Long, X As Byte, Y As Integer

    For X = 0 To 19
        lColor = picThinBox.Point(X, 0)   ' Reads the uppermost pixel MAX LIGHT which is to be faded.
        R = lColor And &HFF
        G = (lColor And &HFF00&) \ &H100&
        b = (lColor And &HFF0000) \ &H10000
        sngR256delToBlack = R / 255   ' Fractions which leads down to black.
        sngG256delToBlack = G / 255
        sngB256delToBlack = b / 255
        For Y = 0 To 255                                   ' Interesting if the is an error, thus a jump directly to EndSub.
            SetPixelV picThinBox.hdc, X, Y, RGB(R, G, b)   ' Painting with API.
            R = R - sngR256delToBlack                      ' Darkening the shade of one 256th.
            G = G - sngG256delToBlack
            b = b - sngB256delToBlack
        Next Y
        Y = Y - 1  ' Because Y gets too big when loop is complete.
    Next X

End Sub


Public Sub RainBowBigbox(blnFadeToGrey, blnFadeToBlack)  ' Is used by both radiobutton 1 & 2.
    Dim Ctr As Byte, blnUpdateTextBoxes As Boolean, bteK4243 As Byte
    Dim Saturation As Single, Luminance As Single
    Static intNODE As Integer, YCtr As Integer, XCtr As Integer, intRainbowAngle As Integer

    ' There is no risk for getting dull shades since I use the native principal by adding/subtracting values against at constant FF-component.
    intRainbowAngle = 0      ' Protects the systemcolorangel
    If blnFadeToGrey = vbTrue And blnFadeToBlack = vbFalse Then
        Saturation = 255
        Luminance = bteBrightnessMax255   ' Starting value fully saturated. Brightness is to be the same for the whole of bigbox.
    Else
        Saturation = bteSaturationMax255
        Luminance = 255   ' Fading from fully bright.
    End If
    ' For intRainbowAngle = 0 To 1529
    bteK4243 = 42   ' Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.
    ' For intNODE = 0 To 1275 Step 255
    XCtr = 0     ' To255
    For YCtr = 0 To 255
        Do
            For Ctr = 1 To bteK4243   ' Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.
                If blnFadeToBlack Then Luminance = 255 - YCtr  ' Round(bteBrightnessMax255 - (bteBrightnessMax255 / 255 * YCtr))
                If blnFadeToGrey Then Saturation = 255 - YCtr  ' Round(bteSaturationMax255 - (bteSaturationMax255 / 255 * YCtr))
                intRainbowAngle = intNODE + ((254 * (Ctr - 1)) / (bteK4243 - 1))   ' Wonderful solution: this logic about going from zero to the full value (here 254) I have been seeking for a long time.
                SetPixelV picBigBox.hdc, XCtr, YCtr, HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, False)
                XCtr = XCtr + 1
            Next Ctr
            If bteK4243 = 43 Then bteK4243 = 42 Else bteK4243 = 43
            intNODE = intNODE + 255   ' Bistabile switch.
        Loop While XCtr < 255
        intRainbowAngle = 0        ' Painting the last fully red which lies outside the logic.
        'picBigBox.PSet (XCtr, YCtr), HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdateTextBoxes)
    intNODE = 0: XCtr = 0: Next YCtr

End Sub


Public Sub RainBowThinBox()        ' By swapping the XY-values at the call you can paint either horisontal or vertical.
    Dim Ctr As Byte, blnUpdateTextBoxes As Boolean, bteK4243 As Byte
    Dim Saturation As Single, Luminance As Single
    Static intNODE As Integer, YCtr As Integer, XCtr As Integer, intRainbowAngle As Integer

    intRainbowAngle = 0                 ' Protecting systemcolorangel
    Saturation = 255: Luminance = 255   ' Fully shining colors.
    ' For intRainbowAngle = 0 To 1529
    bteK4243 = 42                  ' Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.
    ' For intNODE = 0 To 1275 Step 255
    ' Vertical
    For XCtr = 0 To 19
        Do
            For Ctr = 1 To bteK4243                                                ' Has to alternate between 42 and 43 pixels per colorfield to make even at 256 pixels.
                intRainbowAngle = intNODE + ((254 * (Ctr - 1)) / (bteK4243 - 1))   ' Wonderful solution: this logic about going from zero to the full value (here 254) I have been seeking for a long time.
                SetPixelV picThinBox.hdc, XCtr, YCtr, HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdateTextBoxes)
                YCtr = YCtr - 1
            Next Ctr               '
            If bteK4243 = 43 Then bteK4243 = 42 Else bteK4243 = 43
            intNODE = intNODE + 255   ' Bistabile switch.
        Loop While YCtr > 0
        intRainbowAngle = 0        ' Painting the last fully red which is outside the logic of the routine.
        SetPixelV picThinBox.hdc, XCtr, YCtr, HSLToRGB(ByVal intRainbowAngle, ByVal Saturation, ByVal Luminance, blnUpdateTextBoxes)
        intNODE = 0
        YCtr = 255
    Next XCtr

End Sub


Private Sub ColorHistory()

    If blnIgnoreHistory Then Exit Sub
    
    For X = lblHistoryColor.LBound To lblHistoryColor.UBound
        If lblHistoryColor(X).BackColor = picNewColor.BackColor Then
            Exit For: Exit Sub
        End If
    Next
    
    If lblHistoryColor(HistoryIndex).BackColor <> picNewColor.BackColor Then
        HistoryIndex = HistoryIndex + 1
        If HistoryIndex > lblHistoryColor.UBound Then HistoryIndex = 0
    End If
    
    If lblHistoryColor(0).Tag = "blank" Then HistoryIndex = 0
    lblHistoryColor(HistoryIndex).BackColor = picNewColor.BackColor
    lblHistoryColor(HistoryIndex).Tag = picNewColor.BackColor

End Sub


Private Sub ClearHistory()
    For X = lblHistoryColor.LBound To lblHistoryColor.UBound
        lblHistoryColor(X).BackColor = vbWhite
        lblHistoryColor(X).Tag = "blank"
    Next
    HistoryIndex = 0
End Sub


Private Sub cmdLess_Click()
    Me.Width = 6975
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub


Private Sub cmdMore_Click()
    Me.Width = 9435
    Me.Left = (Screen.Width - Me.Width) / 2
End Sub


Public Sub Form_Load()
    Dim udtAngelSaturationBrightness As HSL

    SetIcon Me.hwnd, "AAA"
    LoadResStrings Me
    
    linTriang2Vert.X1 = 311: linTriang2Vert.X2 = 311: linTriang2Vert.Y1 = 251: linTriang2Vert.Y2 = 261
    linTriang2Rising.X1 = 306: linTriang2Rising.X2 = 311: linTriang2Rising.Y2 = 261: linTriang2Rising.Y1 = 256
    linTriang2Falling.X1 = 306: linTriang2Falling.X2 = 311: linTriang2Falling.Y2 = 251: linTriang2Falling.Y1 = 256
    
    Me.Height = 5955
    
    Call LoadSettings     ' load latest settings
    picNewColor.BackColor = lblOldColor.BackColor
    Call PaintThinBox(0)   ' initialize picThinBox
    Call SplitlblNewColorToRGBboxes    ' ALSO THE SYSTEM CONSTANTS OF RGB GETS UPDATED.
    udtAngelSaturationBrightness = RGBToHSL201(picNewColor.BackColor, True)   ' TRUE MEANS THAT HSL IS UPDATING BOTH THE textboxes AND THE systemConstants.
    Call UnhookKeyboard

End Sub


Private Sub lblHistoryColor_Click(Index As Integer)
    Dim udtAngelSaturationBrightness As HSL
    If lblHistoryColor(Index).Tag = "blank" Then Exit Sub
    picNewColor.BackColor = lblHistoryColor(Index).BackColor
    picNewColor.Refresh
    mBlnBigBoxReady = False   ' Delivers fresh coordinates, but only in the HSL-model at this stage.
    Call SplitlblNewColorToRGBboxes   ' Also the system constants RGB are updated.
    udtAngelSaturationBrightness = RGBToHSL201(picNewColor.BackColor, True)
    blnIgnoreHistory = True
    Call objOption_Click(0)
    objOption(0).Value = True
    blnIgnoreHistory = False
End Sub

Private Sub mnuAutoCopy_Click()
    mnuAutoCopy.Checked = Not mnuAutoCopy.Checked
End Sub

Private Sub mnuAutoPaste_Click()
    mnuAutoPaste.Checked = Not mnuAutoPaste.Checked
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuClearHistory_Click()
    Call ClearHistory
End Sub


Private Sub objOption_Click(Index As Integer)

    Select Case Index
        Case 0
            Call PaintThinBox(Index)
            mBteMarkerOldX = bteSaturationMax255: mBteMarkerOldY = 255 - bteBrightnessMax255
        Case 1
            Call PaintThinBox(Index)
            mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteBrightnessMax255
        Case 2
            Call PaintThinBox(Index)
            mBteMarkerOldX = intSystemColorAngleMax1530 / 6: mBteMarkerOldY = 255 - bteSaturationMax255
        Case 3
            Call opt3RedPaintPicThinBox(ByVal Text1(4), Text1(5))
            mBteMarkerOldX = Text1(5): mBteMarkerOldY = 255 - Text1(4)
        Case 4
            Call opt4GreenPaintPicThinBox(ByVal Text1(3), Text1(5))
            mBteMarkerOldX = Text1(5): mBteMarkerOldY = 255 - Text1(3)
        Case 5
            Call opt5BluePaintPicThinBox(ByVal Text1(3), Text1(4))
            mBteMarkerOldX = Text1(3): mBteMarkerOldY = 255 - Text1(4)
    End Select
    
    Call picBigBox_Colorize
    Call ColorHistory
    Call ArrowsModeDepending
    picThinBox.Refresh

End Sub


Private Sub picBigBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If blnDrag = False Then Exit Sub
    If Button = 1 Then
        imgMarker.Visible = False
        picBigBox.Refresh
        If X > 255 Then X = 255    ' Limits
        If X < 0 Then X = 0
        If Y > 255 Then Y = 255
        If Y < 0 Then Y = 0
        If objOption(0) Then lngcolor = HSLToRGB(intSystemColorAngleMax1530, ByVal X, ByVal 255 - Y, True) ' CONVERT AND UPDATE TEXTBOXES.
        If objOption(1) Then lngcolor = HSLToRGB(ByVal X * 6, ByVal bteSaturationMax255, ByVal 255 - Y, True): Call PaintThinBox(1)   ' CONVERT AND UPDATE TEXTBOXES.
        If objOption(2) Then lngcolor = HSLToRGB(ByVal X * 6, ByVal 255 - Y, ByVal bteBrightnessMax255, True): Call PaintThinBox(2)   ' CONVERT AND UPDATE TEXTBOXES.
        If objOption(3) Then Call BigBoxOpt3Reaction(ByVal X, Y)      ' CONVERT AND UPDATE TEXTBOXES.
        If objOption(4) Then Call BigBoxOpt4Reaction(ByVal X, Y)      ' CONVERT AND UPDATE TEXTBOXES.
        If objOption(5) Then Call BigBoxOpt5Reaction(ByVal X, Y)      ' CONVERT AND UPDATE TEXTBOXES.
        mBlnRecentThinBoxPress = False
    End If

End Sub


Private Sub picBigBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngcolor As Long

    If mBlnBigBoxReady = False Then Exit Sub
    If Button = 1 Then
        blnDrag = True
        Select Case objOption(i)
            Case i = 0
                lngcolor = HSLToRGB(intSystemColorAngleMax1530, ByVal X, ByVal 255 - Y, True)   ' CONVERT AND UPDATE TEXTBOXES.
            Case i = 1
                lngcolor = HSLToRGB(ByVal X * 6, ByVal bteSaturationMax255, ByVal 255 - Y, True)   ' CONVERT AND UPDATE TEXTBOXES.
                Call FadeThinBoxToGrey                                                             ' REPAINT ThinBox - FADE SATURATED COLORS; THE SYSTEM CONSTANTS ARE ALREADY UPDATED.
                picThinBox.Refresh
            Case i = 2
                picThinBox.BackColor = HSLToRGB(ByVal X * 6, ByVal 255 - Y, 255, False)            ' SETTING THE BRIGHT COLOR THAT IS TO BE FADED. CONVERTING AND UPDATING TEXTBOXES.
                lngcolor = HSLToRGB(ByVal X * 6, ByVal 255 - Y, ByVal bteBrightnessMax255, True)   ' UPDATING THE REAL, NONSATURATED SYSTEM CONSTANTS AND picNewColor.
                Call FadeThinBoxToBlack                                                            ' REPAINTING ThinBox - FADE SATURATED COLORS ; THE SYSTEM CONSTANTS ARE ALREADY UPDATED.
                picThinBox.Refresh
            Case i = 3
                Call BigBoxOpt3Reaction(ByVal X, Y)
            Case i = 4
                Call BigBoxOpt4Reaction(ByVal X, Y)
            Case i = 5
                Call BigBoxOpt4Reaction(ByVal X, Y)
        End Select
        blnNotFirstTimeMarker = True
    End If

End Sub


Private Sub picBigBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mBlnBigBoxReady = False Then Exit Sub
    If Button = 1 Then
        blnDrag = False
        If X > 255 Then X = 255
        If X < 0 Then X = 0
        If Y > 255 Then Y = 255
        If Y < 0 Then Y = 0
        mBteMarkerOldX = X
        mBteMarkerOldY = Y
        Call MoveMarker(X, Y)
        Call ColorHistory
    End If

End Sub


Private Sub picThinBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' set flag to start drawing
    mBlnRecentThinBoxPress = True
    blnDrag = True: Call picThinBox_MouseMove(Button, Shift, X, Y)   ' REUSING THE UPDATE ROUTINES.
End Sub


Private Sub picThinBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngcolor As Long, udtAngelSaturationBrightness As HSL

    If blnDrag = False Then Exit Sub
    tmrThinBox.Enabled = True
    If Y < 0 Then Y = 0    ' Limits
    If Y > 255 Then Y = 255
    Call TriangleMove(Y)                                                                                                             ' Animation
    If objOption(0) Then lngcolor = HSLToRGB((255 - Y) * 6, ByVal bteSaturationMax255, ByVal bteBrightnessMax255, True): Exit Sub    ' Convert and update textboxes.
    If objOption(1) Then lngcolor = HSLToRGB(ByVal intSystemColorAngleMax1530, 255 - Y, ByVal bteBrightnessMax255, True): Exit Sub   ' Convert and update textboxes.
    If objOption(2) Then lngcolor = HSLToRGB(ByVal intSystemColorAngleMax1530, ByVal bteSaturationMax255, 255 - Y, True)             ' Convert and update textboxes.
    If objOption(3) Then
        Text1(3) = 255 - Y: udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True)   ' Convert and update textboxes.
        picNewColor.BackColor = RGB(Text1(3), Text1(4), Text1(5))
        picNewColor.Refresh
    End If
    If objOption(4) Then
        Text1(4) = 255 - Y: udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True)   ' Convert and update textboxes.
        picNewColor.BackColor = RGB(Text1(3), Text1(4), Text1(5))
        picNewColor.Refresh
    End If
    If objOption(5) Then
        Text1(5) = 255 - Y: udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True)   ' Convert and update textboxes.
        picNewColor.BackColor = RGB(Text1(3), Text1(4), Text1(5))
        picNewColor.Refresh
    End If

End Sub


Private Sub picThinBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' set flag to start drawing
    tmrThinBox.Enabled = False
    blnDrag = False
    Call ColorHistory
End Sub


Private Sub picThinBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intDirektion As Integer

    If objOption(0) Then
        If KeyCode = vbKeyUp Then
            intDirektion = 10
            Call NudgeHueValue(ByVal intDirektion)
            Call picBigBox_Colorize
        End If
        If KeyCode = vbKeyDown Then
            intDirektion = -10
            Call NudgeHueValue(ByVal intDirektion)
            Call picBigBox_Colorize
        End If
    End If

End Sub


Private Sub picThinBox_KeyUp(KeyCode As Integer, Shift As Integer)
    Call ColorHistory
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyCode As Integer)
    If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
        If (KeyCode < vbKey0 Or KeyCode > vbKey9) Then   ' Limiting the numerical textboxes (Text1[x]) to numerical characters
            KeyCode = 0
        End If
    End If
End Sub


Sub Text1_LostFocus(Index As Integer)
    Dim udtAngelSaturationBrightness As HSL, lngcolor As Long   ' Has to take care of intSystemColorAngleMax1530 0 To 1529.

    mBlnBigBoxReady = False    ' Gives me fresh coordinates, but only in the RBG-model at this stage.
    If Index = 0 Then    ' The user adjusted Hue so RGB will be aproximately calculated.
        If Text1(0) > 360 Then Text1(0) = 360
        lngcolor = HSLToRGB(Text1(0) / 360 * 255 * 6, bteSaturationMax255, bteBrightnessMax255, True)
    ElseIf Index = 1 Or Index = 2 Then   ' The user adjusted Saturation so RGB will be aproximately calculated.
        If Text1(Index) > 100 Then Text1(Index) = 100
        lngcolor = HSLToRGB(intSystemColorAngleMax1530, Text1(Index) / 100 * 255, bteBrightnessMax255, True)
    Else
        If Text1(Index) > 255 Then Text1(Index) = 255
        udtAngelSaturationBrightness = RGBToHSL201(RGB(Text1(3), Text1(4), Text1(5)), True)
    End If
    
    Call ArrowsModeDepending
    Call picBigBox_Colorize
    Call ColorHistory

End Sub


Private Sub Timer1_Timer()
    Dim cp As POINTAPI
    Dim dsDC As Long, dsHWND As Long
    Dim Percent As Long, ret As Long
    Dim hr As Integer, vr As Integer, lengthX As Integer, lengthY As Integer
    Dim offsetX As Integer, offsetY As Integer
    Dim blitAreaX As Integer, blitAreaY As Integer, lastcpx As Integer, lastcpy As Integer

    GetCursorPos cp   ' cp has cursor position assigned to it
    txtCoords.Text = " X: " & cp.X & Space$(2) & "Y: " & cp.Y
    
    ' get desktop device context to copy from
    dsDC = GetDC(0&)
    
    ' get screen width, height
    hr = GetDeviceCaps(dsDC, HORZRES)
    vr = GetDeviceCaps(dsDC, VERTRES)
    dsHWND = GetDesktopWindow()
    
    If cboPercent.ListIndex < 0 Then cboPercent.ListIndex = 1
    
    Percent = cboPercent.ItemData(cboPercent.ListIndex)
    lengthX = picZoom.ScaleWidth
    lengthY = picZoom.ScaleHeight
    
    ' center image about mouse
    offsetX = lengthX / (Percent * 2)
    offsetY = lengthY / (Percent * 2)
    
    ' actual area to blit to
    blitAreaX = picZoom.ScaleWidth * Percent
    blitAreaY = picZoom.ScaleHeight * Percent
    
    ' stop copying the screen off the edges
    ' Store the last cursor position that were valid
    If cp.X - offsetX >= -40 And cp.X + offsetX < hr + 40 Then
        lastcpx = cp.X
    End If
    
    If cp.Y - offsetY >= -40 And cp.Y + offsetY < vr + 40 Then
        lastcpy = cp.Y
    End If
    
    ret = StretchBlt(picZoom.hdc, 0, 0, blitAreaX, blitAreaY, dsDC, lastcpx - offsetX, lastcpy - offsetY, lengthX, lengthY, SRCCOPY)
    picZoom.Refresh
    ReleaseDC dsHWND, dsDC

End Sub


Private Sub tmrTransfer_Timer()
    Dim udtAngelSaturationBrightness As HSL

    mBlnBigBoxReady = False   ' Delivers fresh coordinates, but only in the HSL-model at this stage.
    Call SplitlblNewColorToRGBboxes  ' Also the system constants RGB are updated.
    udtAngelSaturationBrightness = RGBToHSL201(picNewColor.BackColor, True)   ' True means that HSL is updating both the textboxes and the system constants.
    blnIgnoreHistory = True
    Call objOption_Click(0)
    blnIgnoreHistory = False
    objOption(0).Value = True
    Me.Show
    tmrTransfer.Enabled = False

End Sub


Private Sub tmrScreenColor_Timer()
    Dim lngDC As Long, lngResult As Long, lnghWnd As Long

    GetCursorPos poiMouse
    lnghWnd = WindowFromPoint(poiMouse.X, poiMouse.Y)
    txtCoords.Text = " X: " & poiMouse.X & Space$(2) & "Y: " & poiMouse.Y
    lngDC = GetDC(lnghWnd)
    
    Call ScreenToClient(lnghWnd, poiMouse)
    
    lngResult = GetPixel(lngDC, poiMouse.X, poiMouse.Y)
    
    If lngResult = -1 Then
        Call BitBlt(picPreView.hdc, 0, 0, picNewColor.Width, picPreView.Height, lngDC, poiMouse.X, poiMouse.Y, vbSrcCopy)
        lngResult = picPreView.Point(0, 0)
    Else
        picPreView.PSet (0, 0), lngResult
    End If
    
    picNewColor.BackColor = lngResult
    
    Call SplitlblNewColorToRGBboxes  ' Also the system constants RGB are updated.
    Call RGBToHSL201(picNewColor.BackColor, True)   ' True means that HSL is updating both the textboxes and the system constants.

    blnIgnoreHistory = True

    Call PaintThinBox(0)
    picThinBox.Refresh
    Call ArrowsModeDepending
    
    blnIgnoreHistory = False
    objOption(0).Value = True

End Sub


Private Sub tmrThinBox_Timer()
    Call picBigBox_Colorize
End Sub

Private Sub txtGBAColor_KeyPress(KeyCode As Integer)
    Call txtHexColor_KeyPress(KeyCode)
End Sub

Private Sub txtGBAColor_LostFocus()
    If Len(txtGBAColor.Text) = 4 Then
        txtHexColor.Text = GB2RGB(txtGBAColor.Text)
        Call txtHexColor_LostFocus
    End If
End Sub

Private Sub txtHexColor_Change()
    txtGBAColor.Text = RGB2GBA(txtHexColor.Text)
     
    If mnuAutoCopy.Checked Then
        If Len(txtGBAColor) = 4 Then Clipboard.Clear: Clipboard.SetText txtGBAColor.Text
    End If
End Sub

Private Sub txtHexColor_KeyPress(KeyCode As Integer) ' Limits the textbox to numerics and A-F and capitals and to six pieces of letters.
    blnIgnoreColorChange = True
    If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
        If (KeyCode > 64 And KeyCode < 71) Then Exit Sub  ' A-F are OK.
        If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub   ' a-f becomes A-F. OK.
        If (KeyCode > 47 And KeyCode < 58) Then Exit Sub   ' Numerics are OK.
        KeyCode = 0 ' All other letters are unwanted.
    End If
    blnIgnoreColorChange = False
End Sub


Private Sub picNewColor_Click()
    Dim udtAngelSaturationBrightness As HSL

    mBlnBigBoxReady = False   ' Delivers fresh coordinates, but only in the HSL-model at this stage.
    Call SplitlblNewColorToRGBboxes  ' Also the system constants RGB are updated.
    udtAngelSaturationBrightness = RGBToHSL201(picNewColor.BackColor, True)   ' True means that HSL is updating both the textboxes and the system constants.
    Call TriangleMove(255 - (intSystemColorAngleMax1530 / 1530 * 255) + bytMove)   ' Animates the triangeln.
    Call picBigBox_Colorize   ' Redraw BigBox

End Sub


Private Sub lblOldColor_Click()
    Dim udtAngelSaturationBrightness As HSL

    picNewColor.BackColor = lblOldColor.BackColor
    mBlnBigBoxReady = False           ' Delivers fresh coordinates, but only in the HSL-model at this stage.
    Call SplitlblNewColorToRGBboxes   ' Also the system constants RGB are updated.
    udtAngelSaturationBrightness = RGBToHSL201(picNewColor.BackColor, True)
    Call TriangleMove(255 - (intSystemColorAngleMax1530 / 1530 * 255) + bytMove)   ' Animates the triangel.
    Call picBigBox_Colorize
    Call ColorHistory

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Len(Me.Tag) > 2 And mnuAutoPaste.Checked Then
        If Left$(Me.Tag, 2) = "Pr" Or Left$(Me.Tag, 2) = "Ch" Then
            frmMain.Tag = txtGBAColor & Me.Tag
            Me.Tag = vbNullString
        Else
            frmGradient.Tag = txtGBAColor & Me.Tag
            Me.Tag = vbNullString
        End If
    End If
End Sub

Private Sub txtHexColor_LostFocus()
    Dim sShift As String
    
    On Error Resume Next
    If Len(txtHexColor.Text) < 6 Then Exit Sub
    
    sShift = txtHexColor.Text: sShift = Right$(sShift, 2) & Mid$(sShift, 3, 2) & Left$(sShift, 2)   ' Shifting RGB to BGR.
    picNewColor.BackColor = ("&H" + sShift)
    Call SplitlblNewColorToRGBboxes   ' Automatic update of the RGB textboxes.
    Call Text1_LostFocus(3)           ' Simulating that the user adjusted the RGBtxtboxes
    Call picBigBox_Colorize

End Sub
