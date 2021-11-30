VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "APE - Advanced Palette Editor"
   ClientHeight    =   8640
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7950
   Icon            =   "frmPaletteInserter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmPaletteInserter.frx":000C
   ScaleHeight     =   8640
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboLoadedBookmark 
      Enabled         =   0   'False
      Height          =   315
      Left            =   795
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Top             =   3525
      Width           =   2175
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      Caption         =   "images"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6600
      TabIndex        =   91
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
      Begin VB.Image imgEnabled 
         Height          =   360
         Index           =   4
         Left            =   1800
         Picture         =   "frmPaletteInserter.frx":630C
         Top             =   720
         Width           =   360
      End
      Begin VB.Image imgDisabled 
         Height          =   360
         Index           =   4
         Left            =   2160
         Picture         =   "frmPaletteInserter.frx":6861
         Top             =   720
         Width           =   360
      End
      Begin VB.Image imgEnabled 
         Height          =   360
         Index           =   1
         Left            =   960
         Picture         =   "frmPaletteInserter.frx":6C0F
         Top             =   240
         Width           =   360
      End
      Begin VB.Image imgDisabled 
         Height          =   360
         Index           =   1
         Left            =   1320
         Picture         =   "frmPaletteInserter.frx":7178
         Top             =   240
         Width           =   360
      End
      Begin VB.Image imgDisabled 
         Height          =   360
         Index           =   3
         Left            =   1320
         Picture         =   "frmPaletteInserter.frx":76E1
         Top             =   720
         Width           =   360
      End
      Begin VB.Image imgEnabled 
         Height          =   360
         Index           =   3
         Left            =   960
         Picture         =   "frmPaletteInserter.frx":7C48
         Top             =   720
         Width           =   360
      End
      Begin VB.Image imgEnabled 
         Height          =   360
         Index           =   2
         Left            =   120
         Picture         =   "frmPaletteInserter.frx":819F
         Top             =   720
         Width           =   360
      End
      Begin VB.Image imgDisabled 
         Height          =   360
         Index           =   2
         Left            =   480
         Picture         =   "frmPaletteInserter.frx":86F7
         Top             =   720
         Width           =   360
      End
      Begin VB.Image imgEnabled 
         Height          =   360
         Index           =   0
         Left            =   120
         Picture         =   "frmPaletteInserter.frx":8C4A
         Top             =   240
         Width           =   360
      End
      Begin VB.Image imgDisabled 
         Height          =   360
         Index           =   0
         Left            =   480
         Picture         =   "frmPaletteInserter.frx":91A9
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdLoadBookmark 
      Caption         =   "Load"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3980
      TabIndex        =   8
      Tag             =   "5014"
      Top             =   3525
      Width           =   1455
   End
   Begin VB.CommandButton cmdPrevBookmark 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   255
      Left            =   450
      TabIndex        =   6
      Top             =   3555
      Width           =   255
   End
   Begin VB.CommandButton cmdNextBookmark 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3060
      TabIndex        =   50
      Top             =   3555
      Width           =   255
   End
   Begin VB.CheckBox chkCompressed 
      Caption         =   "Compressed Palette (LZ77)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2600
      TabIndex        =   5
      Tag             =   "1004"
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CommandButton cmdPrevPal 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3870
      TabIndex        =   2
      Top             =   2050
      Width           =   255
   End
   Begin VB.CommandButton cmdNextPal 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5250
      TabIndex        =   4
      Top             =   2050
      Width           =   255
   End
   Begin VB.CommandButton cmdBookmarks 
      Caption         =   "Bookmarks..."
      Enabled         =   0   'False
      Height          =   450
      Left            =   6190
      TabIndex        =   9
      Tag             =   "31"
      Top             =   3420
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Caption         =   "Changed Palette"
      Height          =   1935
      Left            =   195
      TabIndex        =   56
      Tag             =   "1010"
      Top             =   6030
      Width           =   7600
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1600
         Left            =   6120
         ScaleHeight     =   1605
         ScaleWidth      =   1335
         TabIndex        =   94
         Top             =   200
         Width           =   1335
         Begin VB.CommandButton cmdNextChPal 
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   255
            Left            =   920
            TabIndex        =   47
            Top             =   250
            Width           =   255
         End
         Begin VB.CommandButton cmdPrevChPal 
            Caption         =   "<"
            Enabled         =   0   'False
            Height          =   255
            Left            =   140
            TabIndex        =   45
            Top             =   250
            Width           =   255
         End
         Begin VB.TextBox txtChgPalIndex 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   460
            TabIndex        =   46
            Top             =   240
            Width           =   375
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C0C0C0&
            X1              =   180
            X2              =   1140
            Y1              =   640
            Y2              =   640
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index"
            Height          =   195
            Left            =   465
            TabIndex        =   95
            Tag             =   "1011"
            Top             =   0
            Width           =   390
         End
         Begin VB.Image imgExportChgPal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   675
            Picture         =   "frmPaletteInserter.frx":9565
            ToolTipText     =   "13"
            Top             =   760
            Width           =   360
         End
         Begin VB.Image imgImportChgPal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   240
            Picture         =   "frmPaletteInserter.frx":9ACC
            ToolTipText     =   "12"
            Top             =   760
            Width           =   360
         End
         Begin VB.Image imgClearPal2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   435
            Picture         =   "frmPaletteInserter.frx":A01F
            ToolTipText     =   "11"
            Top             =   1240
            Width           =   360
         End
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   15
         Left            =   5280
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   89
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   14
         Left            =   4560
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   88
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   13
         Left            =   3840
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   87
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   12
         Left            =   3120
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   86
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   11
         Left            =   2400
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   85
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   10
         Left            =   1680
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   84
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   9
         Left            =   960
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   83
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   8
         Left            =   240
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   82
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   7
         Left            =   5280
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   81
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   6
         Left            =   4560
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   80
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   5
         Left            =   3840
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   79
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   4
         Left            =   3120
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   78
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   3
         Left            =   2400
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   77
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   2
         Left            =   1680
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   76
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   1
         Left            =   960
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   75
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picChgPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   0
         Left            =   240
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   74
         Top             =   780
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   240
         MaxLength       =   4
         TabIndex        =   29
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   960
         MaxLength       =   4
         TabIndex        =   30
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   31
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   32
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   33
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   3840
         MaxLength       =   4
         TabIndex        =   34
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   240
         MaxLength       =   4
         TabIndex        =   37
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   960
         MaxLength       =   4
         TabIndex        =   38
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   39
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   40
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   3120
         MaxLength       =   4
         TabIndex        =   41
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   3840
         MaxLength       =   4
         TabIndex        =   42
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   4560
         MaxLength       =   4
         TabIndex        =   35
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   36
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   14
         Left            =   4560
         MaxLength       =   4
         TabIndex        =   43
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtChgPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   15
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   44
         Top             =   1080
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdReplacePalette 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   450
      Left            =   6190
      TabIndex        =   49
      Tag             =   "1006"
      Top             =   2480
      Width           =   1500
   End
   Begin VB.CommandButton cmdLoadPalette 
      Caption         =   "Load"
      Enabled         =   0   'False
      Height          =   450
      Left            =   6190
      TabIndex        =   48
      Tag             =   "1005"
      Top             =   2000
      Width           =   1500
   End
   Begin VB.Frame Frame4 
      Height          =   495
      Left            =   195
      TabIndex        =   54
      Top             =   8025
      Width           =   7600
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2007 HackMew"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2760
         TabIndex        =   55
         Top             =   195
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Palette Loading Options"
      Height          =   1500
      Left            =   195
      TabIndex        =   51
      Tag             =   "1000"
      Top             =   1680
      Width           =   5870
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         ScaleHeight     =   495
         ScaleWidth      =   5415
         TabIndex        =   96
         Top             =   960
         Width           =   5415
         Begin VB.Label lblROMName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROM: None"
            ForeColor       =   &H00C0C0C0&
            Height          =   195
            Left            =   120
            TabIndex        =   97
            Tag             =   "1027"
            Top             =   210
            Width           =   855
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   5280
            Y1              =   120
            Y2              =   120
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   57
         Top             =   360
         Width           =   2055
         Begin VB.OptionButton optOffset 
            Caption         =   "Load from offset"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   1
            Tag             =   "1002"
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optSearch 
            Caption         =   "Load by searching"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Tag             =   "1001"
            Top             =   0
            Width           =   2175
         End
      End
      Begin VB.TextBox txtOffset 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4000
         MaxLength       =   8
         TabIndex        =   3
         Tag             =   "0"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Offset (Hex)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2400
         TabIndex        =   52
         Tag             =   "1003"
         Top             =   405
         Width           =   840
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Bookmarks"
      Height          =   735
      Left            =   195
      TabIndex        =   90
      Tag             =   "5000"
      Top             =   3240
      Width           =   7600
   End
   Begin VB.Frame Frame2 
      Caption         =   "Actual Palette"
      Height          =   1935
      Left            =   195
      TabIndex        =   53
      Tag             =   "1007"
      Top             =   4035
      Width           =   7600
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1600
         Left            =   6120
         ScaleHeight     =   1605
         ScaleWidth      =   1335
         TabIndex        =   92
         Top             =   200
         Width           =   1335
         Begin VB.TextBox txtPrevPalIndex 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   285
            Left            =   460
            TabIndex        =   27
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdPrevPrPal 
            Caption         =   "<"
            Enabled         =   0   'False
            Height          =   255
            Left            =   140
            TabIndex        =   26
            Top             =   250
            Width           =   255
         End
         Begin VB.CommandButton cmdNextPrPal 
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   255
            Left            =   920
            TabIndex        =   28
            Top             =   250
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Index"
            Height          =   195
            Left            =   465
            TabIndex        =   93
            Tag             =   "1008"
            Top             =   0
            Width           =   390
         End
         Begin VB.Image imgClearPal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   675
            Picture         =   "frmPaletteInserter.frx":A588
            ToolTipText     =   "7"
            Top             =   1240
            Width           =   360
         End
         Begin VB.Image imgAddBookmarks 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   915
            Picture         =   "frmPaletteInserter.frx":AAF1
            ToolTipText     =   "32"
            Top             =   760
            Width           =   360
         End
         Begin VB.Image imgCopyPal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   240
            Picture         =   "frmPaletteInserter.frx":AE9F
            ToolTipText     =   "6"
            Top             =   1240
            Width           =   360
         End
         Begin VB.Image imgImportPrevPal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   0
            Picture         =   "frmPaletteInserter.frx":B25B
            ToolTipText     =   "8"
            Top             =   760
            Width           =   360
         End
         Begin VB.Image imgExportPrevPal 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   465
            Picture         =   "frmPaletteInserter.frx":B7AE
            ToolTipText     =   "9"
            Top             =   760
            Width           =   360
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   180
            X2              =   1140
            Y1              =   640
            Y2              =   640
         End
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   15
         Left            =   5280
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   73
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   14
         Left            =   4560
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   72
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   13
         Left            =   3840
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   71
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   12
         Left            =   3120
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   70
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   11
         Left            =   2400
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   69
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   10
         Left            =   1680
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   68
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   9
         Left            =   960
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   67
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   8
         Left            =   240
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   66
         Top             =   1440
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   7
         Left            =   5280
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   65
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   6
         Left            =   4560
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   64
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   5
         Left            =   3840
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   63
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   4
         Left            =   3120
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   62
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   3
         Left            =   2400
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   61
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   2
         Left            =   1680
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   60
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   1
         Left            =   960
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   59
         Top             =   780
         Width           =   600
      End
      Begin VB.PictureBox picPrevPreview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   145
         Index           =   0
         Left            =   240
         ScaleHeight     =   120
         ScaleWidth      =   570
         TabIndex        =   58
         Top             =   780
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   15
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   25
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   14
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   24
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   7
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   17
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   6
         Left            =   4560
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   16
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   13
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   23
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   12
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   22
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   11
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   21
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   10
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   20
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   9
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   19
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   8
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1080
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   5
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   15
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   4
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   14
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   13
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   12
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   11
         Top             =   420
         Width           =   600
      End
      Begin VB.TextBox txtPrevPal 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00808080&
         Height          =   300
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   10
         Top             =   420
         Width           =   600
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   1
      Begin VB.Menu mnuOpenROM 
         Caption         =   "Open ROM..."
         HelpContextID   =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         HelpContextID   =   3
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   4
      Begin VB.Menu mnuActualPalette 
         Caption         =   "Actual Palette"
         Enabled         =   0   'False
         HelpContextID   =   5
         Begin VB.Menu mnuCopyPal 
            Caption         =   "Copy"
            Enabled         =   0   'False
            HelpContextID   =   6
         End
         Begin VB.Menu mnuClearPal 
            Caption         =   "Clear"
            Enabled         =   0   'False
            HelpContextID   =   7
         End
         Begin VB.Menu mnuSeparator3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuImportPrevPal 
            Caption         =   "Import..."
            Enabled         =   0   'False
            HelpContextID   =   8
         End
         Begin VB.Menu mnuExportPrevPal 
            Caption         =   "Export..."
            Enabled         =   0   'False
            HelpContextID   =   9
         End
         Begin VB.Menu mnuSepp 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAddBookmarks 
            Caption         =   "Add to Bookmarks"
            Enabled         =   0   'False
            HelpContextID   =   32
         End
      End
      Begin VB.Menu mnuChangedPalette 
         Caption         =   "Changed Palette"
         Enabled         =   0   'False
         HelpContextID   =   10
         Begin VB.Menu mnuClearPal2 
            Caption         =   "Clear"
            Enabled         =   0   'False
            HelpContextID   =   11
         End
         Begin VB.Menu mnuSeparator5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuImportChgPal 
            Caption         =   "Import..."
            Enabled         =   0   'False
            HelpContextID   =   12
         End
         Begin VB.Menu mnuExportChgPal 
            Caption         =   "Export..."
            Enabled         =   0   'False
            HelpContextID   =   13
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      HelpContextID   =   14
      Begin VB.Menu mnuNextPal 
         Caption         =   "Load Next Palette"
         Enabled         =   0   'False
         HelpContextID   =   15
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuPrevPal 
         Caption         =   "Load Previous Palette"
         Enabled         =   0   'False
         HelpContextID   =   16
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSeparator6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActualPal 
         Caption         =   "Actual Palette"
         Enabled         =   0   'False
         HelpContextID   =   17
         Begin VB.Menu mnuNextPrPal 
            Caption         =   "Next Index"
            Enabled         =   0   'False
            HelpContextID   =   18
         End
         Begin VB.Menu mnuPrevPrPal 
            Caption         =   "Previous Index"
            Enabled         =   0   'False
            HelpContextID   =   19
         End
         Begin VB.Menu mnuSeparator7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGotoPrev 
            Caption         =   "Goto Index..."
            Enabled         =   0   'False
            HelpContextID   =   20
         End
      End
      Begin VB.Menu mnuChangedPal 
         Caption         =   "Changed Palette"
         Enabled         =   0   'False
         HelpContextID   =   21
         Begin VB.Menu mnuNextChPal 
            Caption         =   "Next Index"
            Enabled         =   0   'False
            HelpContextID   =   22
         End
         Begin VB.Menu mnuPrevChPal 
            Caption         =   "Previous Index"
            Enabled         =   0   'False
            HelpContextID   =   23
         End
         Begin VB.Menu mnuSeparator8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuGotoChg 
            Caption         =   "Goto Index..."
            Enabled         =   0   'False
            HelpContextID   =   24
         End
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookmarks 
         Caption         =   "Bookmarks..."
         Enabled         =   0   'False
         HelpContextID   =   31
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      HelpContextID   =   25
      Begin VB.Menu mnuColorPicker 
         Caption         =   "Color picker"
         HelpContextID   =   26
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuCalc 
         Caption         =   "RGB/GBA Converter"
         HelpContextID   =   27
      End
      Begin VB.Menu mnuGradient 
         Caption         =   "Gradient-o-matic"
         Enabled         =   0   'False
         HelpContextID   =   33
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   28
      Begin VB.Menu mnuReadme 
         Caption         =   "Readme"
         HelpContextID   =   29
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         HelpContextID   =   30
      End
   End
   Begin VB.Menu mnuClearOptions 
      Caption         =   "ClearOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuAllPalettes 
         Caption         =   "All Palettes"
         HelpContextID   =   34
         Index           =   0
      End
      Begin VB.Menu mnuAllPalettes 
         Caption         =   "Only Active Palette"
         Checked         =   -1  'True
         HelpContextID   =   35
         Index           =   1
      End
   End
   Begin VB.Menu mnuClearOptions2 
      Caption         =   "ClearOptions2"
      Visible         =   0   'False
      Begin VB.Menu mnuAllPalettes2 
         Caption         =   "All Palettes"
         HelpContextID   =   34
         Index           =   0
      End
      Begin VB.Menu mnuAllPalettes2 
         Caption         =   "Only Active Palette"
         Checked         =   -1  'True
         HelpContextID   =   35
         Index           =   1
      End
   End
   Begin VB.Menu mnuCopyOptions 
      Caption         =   "CopyOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyAllPalettes 
         Caption         =   "All Palettes"
         HelpContextID   =   34
         Index           =   0
      End
      Begin VB.Menu mnuCopyAllPalettes 
         Caption         =   "Only Active Palette"
         Checked         =   -1  'True
         HelpContextID   =   35
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Offset As Long
Dim i As Long, X As Long, FileNum As Integer, sResult As String, FileLength As Long
Dim LastFilter As Long, blnCompressed As Boolean
Dim blnIgnoreChange As Boolean, blnIgnoreChange2 As Boolean, blnCopying As Boolean
Dim blnClearing As Boolean, blnClearing2 As Boolean
Dim arrPalette(255) As String, arrPalette2(255) As String, arrPalOffset(15) As String, arrPalCompressed(15) As Byte

Private Const vbUnsafeColor As Long = &HC0C0FF
Private Const CB_FINDSTRING = &H14C
Private blnDelete As Boolean
Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Dim nRet As Long

Private Function ByteSwap(Data As Byte, Data2 As Byte) As String
    ByteSwap = Right$("0" & Hex$(Data), 2) & Right$("0" & Hex$(Data2), 2)
End Function

Private Function Char(TextBox As String) As String
    Char = Chr$(Val("&H" & Left$(TextBox, 2))) & Chr$(Val("&H" & Right$(TextBox, 2)))
End Function

Private Function UnCompress(strChar As String, lngIndex As Long) As String
    UnCompress = Right$("0" & Hex$(Asc(Mid$(strChar, lngIndex, 2))), 2) & Right$("0" & Hex$(Asc(Mid$(strChar, (lngIndex + 1), 2))), 2)
End Function

Private Function IsHex(ByRef sInput As String) As Boolean
Dim i As Integer
Dim sValue As String
Dim iLength As Integer
Dim sChars As String

sChars = "0123456789abcdefABCDEF"
iLength = Len(sInput)

For i = 1 To iLength
    sValue = Mid$(sInput, i, 1)

    IsHex = True
    If InStrB(1, sChars, sValue, vbBinaryCompare) = 0 Then
        IsHex = False
        Exit Function
    End If
Next

End Function

Private Function SafeCheck(sString As String) As Boolean
    
    Const IntMax As Integer = 32767 '&H7FFF
    
    If Len(sString) < 4 Then SafeCheck = True: Exit Function
    
    sString = Right$(sString, 2) & Left$(sString, 2)
    
    If CLng("&H" & sString) <= IntMax And CLng("&H" & sString) >= 0 Then
        SafeCheck = True
    Else
        SafeCheck = False
    End If

End Function

Private Sub ReadData()
    X = 1
    Dim pal(1) As Byte
    Dim molt As Integer
    molt = 16 * Val(txtPrevPalIndex.Tag) - 16
    If molt < 0 Then molt = 0
    
    FileNum = FreeFile
    Open sResult For Binary As #FileNum
        For i = txtPrevPal.LBound To txtPrevPal.UBound
            Get #FileNum, Offset + X, pal
            txtPrevPal(i).Text = ByteSwap(pal(0), pal(1))
            arrPalette(molt + i) = txtPrevPal(i).Text
            X = X + 2
        Next
    Close #FileNum
    
    txtOffset.Text = Right$(String$(7, vbKey0) & txtOffset.Text, txtOffset.MaxLength)
    arrPalOffset(Val(txtPrevPalIndex.Text) - 1) = txtOffset.Text
    arrPalCompressed(Val(txtPrevPalIndex.Text) - 1) = chkCompressed.Value
       
    If Offset = 0 Then
        mnuPrevPal.Enabled = False
        cmdPrevPal.Enabled = False
    ElseIf Offset > 31 And Offset <= FileLength - 32 Then
        mnuPrevPal.Enabled = True
        cmdPrevPal.Enabled = True
    End If
    
End Sub

Private Sub DecompressData()

    Dim CmpData(1 To 40) As Byte, CompressionCheck As String, CmpString As String
    Dim z As Long

    'CompressionCheck = Chr$(16) & Space$(1) & String$(2, vbNullChar)
    CompressionCheck = String$(2, vbNullChar)
    
    FileNum = FreeFile
    Open sResult For Binary As #FileNum
        Get #FileNum, Offset + 1, CmpData
    Close #FileNum

    For i = LBound(CmpData) To UBound(CmpData)
        CmpString = CmpString & Chr$(CmpData(i))
    Next
    
    Dim molt As Long
    molt = 16 * CLng(txtPrevPalIndex.Tag) - 16

    If molt < 0 Then molt = 0

    If Left$(CmpString, 1) = Chr$(16) And InStr(1, CmpString, CompressionCheck, vbBinaryCompare) = 3 Then
        CmpString = Mid$(CmpString, 6)
        z = 1
        For i = txtPrevPal.LBound To txtPrevPal.UBound
            txtPrevPal(i).Text = UnCompress(CmpString, z)
            arrPalette(molt + i) = txtPrevPal(i).Text
            If z = 7 Or z = 16 Or z = 25 Then z = z + 1
            z = z + 2
        Next
        arrPalOffset(Val(txtPrevPalIndex.Text) - 1) = txtOffset.Text
        arrPalCompressed(Val(txtPrevPalIndex.Text) - 1) = chkCompressed.Value
    Else
        MsgBox LoadResString(1013), vbExclamation
    End If
    
    If Offset = 0 Then
        mnuPrevPal.Enabled = False
        cmdPrevPal.Enabled = False
    ElseIf Offset > 31 Then
        mnuPrevPal.Enabled = True
        cmdPrevPal.Enabled = True
    End If
    
End Sub


Private Sub RecompressData()

    'Dim CmpHeader As String,
    Dim Data As Byte, Data2 As Byte ', ReplaceData As Byte, ReplaceData2 As Byte
'    Dim strData As String * 40
    Dim z As Long
        
    'CmpHeader = Chr$(16) & Space$(1) & String$(3, vbNullChar)
    z = 1
    
    FileNum = FreeFile
    Open sResult For Binary As #FileNum
        'Put #FileNum, Offset + 1, CmpHeader
        Offset = Offset + 5
'        Get #FileNum, Offset, strData
        For i = txtChgPal.LBound To txtChgPal.UBound
'            Data = Val("&H" & Left$(txtPrevPal(i).Text, 2))
'            Data2 = Val("&H" & Right$(txtPrevPal(i).Text, 2))
'
'            ReplaceData = Val("&H" & Left$(txtChgPal(i).Text, 2))
'            ReplaceData2 = Val("&H" & Right$(txtChgPal(i).Text, 2))
'
'            z = InStr(1, strData, Data & Data2, vbBinaryCompare)
'
'            Put #FileNum, Offset + z + i, ReplaceData
'            Put #FileNum, Offset + z + i + 1, ReplaceData
        
            Data = Val("&H" & Left$(txtChgPal(i).Text, 2))
            Data2 = Val("&H" & Right$(txtChgPal(i).Text, 2))
            Put #FileNum, Offset + z + i, Data
            Put #FileNum, Offset + z + i + 1, Data2
            If z = 4 Or z = 9 Or z = 14 Then z = z + 1 'Put #FileNum, Offset + z + i + 2, vbNullChar: z = z + 1
            z = z + 1
        Next
    Close #FileNum

End Sub


Private Sub NavigateCompressed(blnNext As Boolean)

    Dim lngIndex As Long, lngFound As Long, strSearch As String, strContents As String
    
    FileNum = FreeFile
    Open sResult For Binary As #FileNum
        strContents = Space$(LOF(FileNum))
        Get #FileNum, , strContents
    Close #FileNum
    
    strSearch = Chr$(16) & Space$(1) & String$(2, vbNullChar)
    
    If blnNext Then
        lngIndex = InStr(CLng("&H" & txtOffset.Text) + 4, strContents, strSearch, vbBinaryCompare)
    Else
        lngIndex = InStrRev(strContents, strSearch, CLng("&H" & txtOffset.Text) - 4, vbBinaryCompare)
    End If
    If Not lngIndex = 0 Then
        lngFound = lngIndex - 1
        txtOffset.Text = Right$(String$(7, vbKey0) & Hex$(lngFound), txtOffset.MaxLength)
        Offset = lngFound
        Call DecompressData
    Else
        MsgBox LoadResString(1014), vbExclamation
    End If

End Sub

Private Sub SearchCompressed()
  
    Dim lngIndex As Long, lngFound As Long, strSearch As String, strContents As String
    
    FileNum = FreeFile
    Open sResult For Binary As #FileNum
        strContents = Space$(LOF(FileNum))
        Get #FileNum, , strContents
    Close #FileNum
    
    For i = txtPrevPal.LBound To txtPrevPal.UBound
        If Len(txtPrevPal(i).Text) = 4 Then
            strSearch = strSearch & Char(txtPrevPal(i).Text)
            If i = 3 Or i = 7 Or i = 11 Then strSearch = strSearch & vbNullChar
        Else
            Exit For
        End If
    Next
    
    lngIndex = InStr(CLng("&H" & txtOffset.Tag) + 5 + 1, strContents, strSearch, vbBinaryCompare)
    
    If Not lngIndex = 0 Then
        If Mid$(strContents, lngIndex - 5, 1) = Chr$(16) Then
            If Mid$(strContents, lngIndex - 3, 2) = String$(2, vbNullChar) Then
                lngFound = lngIndex - 6
                txtOffset.Text = Right$(String$(7, vbKey0) & Hex$(lngFound), txtOffset.MaxLength)
                Offset = lngFound
                txtOffset.Tag = Hex$(lngIndex - 5)
                optOffset.Tag = txtOffset.Text
            Else
                txtOffset.Tag = "0"
                txtOffset.Text = String$(8, vbKey0)
                MsgBox LoadResString(1015), vbExclamation
            End If
        'Else
        '    txtOffset.Tag = "0"
        '    txtOffset.Text = String$(8, vbKey0)
        '    MsgBox LoadResString(1015), vbExclamation
        End If
    Else
        If txtOffset.Tag = "0" Then
            MsgBox LoadResString(1015), vbExclamation
        Else
            MsgBox LoadResString(1016), vbExclamation
        End If
    End If
End Sub

Private Sub SearchPalette()

    FileNum = FreeFile
    Dim lngIndex As Long, lngFound As Long, strSearch As String, strContents As String
    Open sResult For Binary As #FileNum
        strContents = Space$(LOF(FileNum))
        Get #FileNum, , strContents
    Close #FileNum
       
    For i = txtPrevPal.LBound To txtPrevPal.UBound
        If Len(txtPrevPal(i)) = 4 Then
            strSearch = strSearch & Char(txtPrevPal(i))
        Else
            Exit For
        End If
    Next
       
    lngIndex = InStr(CLng("&H" & txtOffset.Tag) + 1, strContents, strSearch, vbBinaryCompare)
    
    If Not lngIndex = 0 Then
        lngFound = lngIndex - 1
        txtOffset.Text = Right$(String$(7, vbKey0) & Hex$(lngFound), txtOffset.MaxLength)
        Offset = lngFound
        txtOffset.Tag = Hex$(lngIndex)
        optOffset.Tag = txtOffset.Text
    Else
        If txtOffset.Tag = "0" Then
            MsgBox LoadResString(1015), vbExclamation
        Else
            MsgBox LoadResString(1016), vbExclamation
        End If
    End If

End Sub

Private Sub ExportPalette(arrName() As String)

    Dim sExport As String, tmp As String
      
    FileNum = FreeFile
    
    Dim oOpenDialog As clsCommonDialog
    Set oOpenDialog = New clsCommonDialog
    
    sExport = oOpenDialog.ShowSave(Me.hwnd, LoadResString(1021), "palette_" & txtOffset.Text, , "APE Palette (*.gpl)|*.gpl|Adobe Color Table (*.act)|*.act|PaintShop Palette (*.pal)|*.pal|Tile Layer Pro Palette (*.tpl)|*.tpl|", FILEMUSTEXIST Or PATHMUSTEXIST Or OVERWRITEPROMPT)
    LastFilter = oOpenDialog.FilterIndex
    
    If LenB(sExport) > 0 Then
        Select Case LastFilter
            Case 1
                Open sExport For Output As #FileNum
                    Print #FileNum, "[APE Palette]"
                    For i = LBound(arrName) To UBound(arrName)
                        Print #FileNum, Val("&H" & arrName(i))
                    Next
                Close FileNum
            Case 2
                Call KillFileIfExists(sExport)
                Open sExport For Binary As #FileNum
                    For i = LBound(arrName) To UBound(arrName)
                        tmp = GB2RGB(arrName(i))
                        Put #FileNum, , CByte(Val("&H" & Mid$(tmp, 1, 2)))
                        Put #FileNum, , CByte(Val("&H" & Mid$(tmp, 3, 2)))
                        Put #FileNum, , CByte(Val("&H" & Mid$(tmp, 5, 2)))
                    Next
                Close FileNum
            Case 3
                Open sExport For Output As #FileNum
                    Print #FileNum, "JASC-PAL"
                    Print #FileNum, "0100"
                    Print #FileNum, "256"
                    For i = LBound(arrName) To UBound(arrName)
                        tmp = GB2RGB(arrName(i))
                        Print #FileNum, Val("&H" & Mid$(tmp, 1, 2)) & " " & Val("&H" & Mid$(tmp, 3, 2)) & " " & Val("&H" & Mid$(tmp, 5, 2))
                    Next
                Close FileNum
            Case 4
                Call KillFileIfExists(sExport)
                Open sExport For Binary As #FileNum
                    Put #FileNum, , "TPL" & vbNullChar
                    For i = LBound(arrName) To UBound(arrName)
                        tmp = GB2RGB(arrName(i))
                        Put #FileNum, , CByte(Val("&H" & Mid$(tmp, 1, 2)))
                        Put #FileNum, , CByte(Val("&H" & Mid$(tmp, 3, 2)))
                        Put #FileNum, , CByte(Val("&H" & Mid$(tmp, 5, 2)))
                    Next
                Close FileNum
        End Select
    End If

End Sub

Private Sub ToggleControl(bEnabled As Boolean)

    optSearch.Enabled = bEnabled
    optOffset.Enabled = bEnabled
    Label3.Enabled = bEnabled
    txtOffset.Enabled = bEnabled
    chkCompressed.Enabled = bEnabled
   
    mnuImportChgPal.Enabled = bEnabled
    If mnuImportChgPal.Enabled = True Then
        imgImportChgPal.Enabled = True: imgImportChgPal.Picture = imgEnabled(2).Picture
    Else
        imgImportChgPal.Enabled = False: imgImportChgPal.Picture = imgDisabled(2).Picture
    End If
    
    mnuBookmarks.Enabled = bEnabled
    cmdBookmarks.Enabled = bEnabled
    mnuActualPalette.Enabled = bEnabled
    mnuChangedPalette.Enabled = bEnabled
    mnuActualPal.Enabled = bEnabled
    mnuChangedPal.Enabled = bEnabled
    mnuGradient.Enabled = bEnabled
    
    For i = txtPrevPal.LBound To txtPrevPal.UBound
        txtPrevPal(i).Enabled = bEnabled
        txtChgPal(i).Enabled = bEnabled
    Next
    
    

End Sub

Private Sub cboLoadedBookmark_Change()
Dim nSelStart As Long

On Error Resume Next
  
    If blnDelete Then
        If cboLoadedBookmark.Text = vbNullString Then cmdLoadBookmark.Enabled = False
        Exit Sub
    End If
  
    nRet = SendMessageString(cboLoadedBookmark.hwnd, CB_FINDSTRING, 0, cboLoadedBookmark.Text)
    nSelStart = cboLoadedBookmark.SelStart
  
    If nRet >= 0 Then
        cboLoadedBookmark.ListIndex = nRet
        cboLoadedBookmark.SelStart = nSelStart
        cboLoadedBookmark.SelLength = Len(cboLoadedBookmark.Text)
        cmdLoadBookmark.Enabled = True
        cboLoadedBookmark.Tag = nRet
        frmBookmarks.lstBookmark.ListIndex = Val(cboLoadedBookmark.Tag)
    Else
        cmdLoadBookmark.Enabled = False
    End If
    
End Sub

Private Sub cboLoadedBookmark_GotFocus()
    cboLoadedBookmark.SelStart = 0
    cboLoadedBookmark.SelLength = Len(cboLoadedBookmark.Text)
End Sub

Private Sub cboLoadedBookmark_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyDelete, vbKeyBack
            blnDelete = True
        Case Else
            blnDelete = False
    End Select
   
End Sub

Private Sub cboLoadedBookmark_KeyPress(KeyCode As Integer)
    If KeyCode <> vbKeyBack Then
        If Len(cboLoadedBookmark) > 25 Then
            KeyCode = 0
        End If
    End If
End Sub

Private Sub cboLoadedBookmark_LostFocus()
    If cboLoadedBookmark.List(nRet) = cboLoadedBookmark.Text Then cboLoadedBookmark.ListIndex = nRet
End Sub

Private Sub chkCompressed_Click()
    blnCompressed = CBool(chkCompressed.Value)
    txtOffset.Tag = "0"
End Sub

Private Sub cmdBookmarks_Click()
    Call mnuBookmarks_Click
End Sub

Private Sub cmdLoadBookmark_Click()

If cboLoadedBookmark.ListIndex < 0 Then Exit Sub

    optOffset.Value = True
    txtOffset.Text = sReadIniFileString(App.Path & INIFile, sGameCode, cboLoadedBookmark.List(cboLoadedBookmark.ListIndex))
    
    If cboLoadedBookmark.ListCount > 1 Then
        Select Case cboLoadedBookmark.ListIndex
            Case 0
                cmdPrevBookmark.Enabled = False
                cmdNextBookmark.Enabled = True
            Case cboLoadedBookmark.ListCount - 1
                cmdPrevBookmark.Enabled = True
                cmdNextBookmark.Enabled = False
            Case Else
                cmdPrevBookmark.Enabled = True
                cmdNextBookmark.Enabled = True
        End Select
    Else
        cmdPrevBookmark.Enabled = False
        cmdNextBookmark.Enabled = False
    End If

    If Left$(cboLoadedBookmark.List(cboLoadedBookmark.ListIndex), 3) = CompressedPal Then
        chkCompressed.Value = vbChecked
    Else
        chkCompressed.Value = vbUnchecked
    End If

    Call cmdLoadPalette_Click
    
End Sub

Private Sub cmdNextBookmark_Click()
On Error Resume Next

    cmdPrevBookmark.Enabled = True
    
    cboLoadedBookmark.ListIndex = Val(cboLoadedBookmark.Tag) + 1
    cboLoadedBookmark.Tag = cboLoadedBookmark.ListIndex
    frmBookmarks.lstBookmark.ListIndex = Val(cboLoadedBookmark.Tag)
    
    cmdLoadBookmark.Enabled = True
    
    If cboLoadedBookmark.ListIndex < cboLoadedBookmark.ListCount - 1 Then
        cmdNextBookmark.Enabled = True
    Else
        cmdNextBookmark.Enabled = False
    End If

End Sub

Private Sub cmdNextChPal_Click()
    
    Dim molt As Long
    
    If Val(txtChgPalIndex.Text) < txtChgPal.UBound + 1 Then
        txtChgPalIndex.Text = Val(txtChgPalIndex.Text) + 1
    End If
    
    molt = 16 * CLng(txtChgPalIndex.Text) - 1
    
    blnIgnoreChange2 = True
    
    X = 0
    
    For i = molt - 15 To molt
        txtChgPal(X).Text = arrPalette2(i)
        X = X + 1
    Next
       
    If Val(txtChgPalIndex.Text) = 16 Then cmdNextChPal.Enabled = False: mnuNextChPal.Enabled = False
    cmdPrevChPal.Enabled = True
    mnuPrevChPal.Enabled = True
    blnIgnoreChange2 = False
    
    Call ChgIndexChange

End Sub

Private Sub cmdNextPal_Click()
    Call mnuNextPal_Click
End Sub

Private Sub cmdNextPrPal_Click()
    
    Dim molt As Long
    
    If Val(txtPrevPalIndex.Text) < txtPrevPal.UBound + 1 Then
        txtPrevPalIndex.Text = Val(txtPrevPalIndex.Text) + 1
    End If
    
    molt = 16 * CLng(txtPrevPalIndex.Tag) - 1
    
    blnIgnoreChange = True
    
    X = 0
    
    For i = molt - 15 To molt
        txtPrevPal(X).Text = arrPalette(i)
        X = X + 1
    Next
    
    If LenB(arrPalOffset(Val(txtPrevPalIndex.Text) - 1)) > 0 Then
        txtOffset.Text = Right$(String$(7, vbKey0) & arrPalOffset(Val(txtPrevPalIndex.Text) - 1), txtOffset.MaxLength)
        chkCompressed.Value = arrPalCompressed(Val(txtPrevPalIndex.Text) - 1)
    End If
       
    cmdPrevPrPal.Enabled = True
    mnuPrevPrPal.Enabled = True
    blnIgnoreChange = False
    
    cboLoadedBookmark.Text = vbNullString
    Call PrevIndexChange
    
    txtOffset.Tag = "0"

End Sub


Private Sub cmdPrevBookmark_Click()
On Error Resume Next

    cmdNextBookmark.Enabled = True
    
    cboLoadedBookmark.ListIndex = Val(cboLoadedBookmark.Tag) - 1
    cboLoadedBookmark.Tag = cboLoadedBookmark.ListIndex
    frmBookmarks.lstBookmark.ListIndex = Val(cboLoadedBookmark.Tag)
    
    cmdLoadBookmark.Enabled = True
    
    If cboLoadedBookmark.ListIndex > 0 Then
        cmdPrevBookmark.Enabled = True
    Else
        cmdPrevBookmark.Enabled = False
    End If

End Sub

Private Sub ChgIndexChange()
    
    For i = txtChgPal.LBound To txtChgPal.UBound
        If Len(txtChgPal(i).Text) = 4 Then
            cmdReplacePalette.Enabled = True
            mnuClearPal2.Enabled = True
            imgClearPal2.Enabled = True: imgClearPal2.Picture = imgEnabled(1).Picture
            mnuExportChgPal.Enabled = True
            imgExportChgPal.Enabled = True: imgExportChgPal.Picture = imgEnabled(3).Picture
        Else
            cmdReplacePalette.Enabled = False
            mnuClearPal2.Enabled = False
            imgClearPal2.Enabled = False: imgClearPal2.Picture = imgDisabled(1).Picture
            mnuExportChgPal.Enabled = False
            imgExportChgPal.Enabled = False: imgExportChgPal.Picture = imgDisabled(3).Picture
        End If
        
    Next
    
End Sub

Private Sub cmdPrevChPal_Click()
    
    Dim molt As Long
    
    If Val(txtChgPalIndex.Text) > txtChgPal.LBound + 1 Then
        txtChgPalIndex.Text = Val(txtChgPalIndex.Text) - 1
    End If
    
    molt = 16 * CLng(txtChgPalIndex.Text) - 1
    
    blnIgnoreChange2 = True
    
    X = 0
    
    For i = molt - 15 To molt
        txtChgPal(X).Text = arrPalette2(i)
        X = X + 1
    Next
    
    If Val(txtChgPalIndex.Text) = 1 Then cmdPrevChPal.Enabled = False: mnuPrevChPal.Enabled = False
    cmdNextChPal.Enabled = True
    mnuNextChPal.Enabled = True
    blnIgnoreChange2 = False
    
    Call ChgIndexChange

End Sub

Private Sub cmdPrevPal_Click()
    Call mnuPrevPal_Click
End Sub

Private Sub PrevIndexChange()

    For i = txtPrevPal.LBound To txtPrevPal.UBound
        If Len(txtPrevPal(i).Text) = 4 Then
            mnuCopyPal.Enabled = True
            imgCopyPal.Enabled = True: imgCopyPal.Picture = imgEnabled(0).Picture
            If optSearch.Value Then
                mnuClearPal.Enabled = True
                imgClearPal.Enabled = True: imgClearPal.Picture = imgEnabled(1).Picture
            End If
            mnuExportPrevPal.Enabled = True
            imgExportPrevPal.Enabled = True: imgExportPrevPal.Picture = imgEnabled(3).Picture
        Else
            mnuCopyPal.Enabled = False
            imgCopyPal.Enabled = False: imgCopyPal.Picture = imgDisabled(0).Picture
            If optSearch.Value Then
                mnuClearPal.Enabled = False
                imgClearPal.Enabled = False: imgClearPal.Picture = imgDisabled(1).Picture
            End If
            mnuExportPrevPal.Enabled = False
            imgExportPrevPal.Enabled = False: imgExportPrevPal.Picture = imgDisabled(3).Picture
        End If
    Next
    
End Sub

Private Sub cmdPrevPrPal_Click()
    
    Dim molt As Long
    
    If Val(txtPrevPalIndex.Text) > txtPrevPal.LBound + 1 Then
        txtPrevPalIndex.Text = Val(txtPrevPalIndex.Text) - 1
    End If
    
    molt = 16 * CLng(txtPrevPalIndex.Tag) - 1
    
    blnIgnoreChange = True
    
    X = 0
    
    For i = molt - 15 To molt
        txtPrevPal(X).Text = arrPalette(i)
        X = X + 1
    Next
    
    If LenB(arrPalOffset(Val(txtPrevPalIndex.Text) - 1)) > 0 Then
        txtOffset.Text = Right$(String$(7, vbKey0) & arrPalOffset(Val(txtPrevPalIndex.Text) - 1), txtOffset.MaxLength)
        chkCompressed.Value = arrPalCompressed(Val(txtPrevPalIndex.Text) - 1)
    End If
    
    cmdNextPrPal.Enabled = True
    mnuNextPrPal.Enabled = True
    blnIgnoreChange = False
    
    cboLoadedBookmark.Text = vbNullString
    Call PrevIndexChange
    
    txtOffset.Tag = "0"
    
End Sub

Private Sub Form_Activate()
    Dim strColor As String, strType As String, intIndex As Integer
       
    If Len(Me.Tag) > 6 Then
        strColor = Mid$(Me.Tag, 1, 4)
        strType = Mid$(Me.Tag, 5, 2)
        intIndex = CInt(Mid$(Me.Tag, 7))
    Else
        Exit Sub
    End If
    
    Select Case strType
        Case "Pr"
            txtPrevPal(intIndex).Text = strColor
        Case "Ch"
            txtChgPal(intIndex).Text = strColor
    End Select
    
    Me.Tag = vbNullString
    
End Sub

Private Sub Form_Load()
    SetIcon Me.hwnd, "AAA"
    LoadResStrings Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmColorPicker
    Unload frmCalc
    Unload frmBookmarks
    Unload frmGradient
End Sub

Public Sub cmdLoadPalette_Click()

    If LenB(sResult) > 0 Then
        mnuCopyPal.Enabled = True
        imgCopyPal.Enabled = True: imgCopyPal.Picture = imgEnabled(0).Picture
        mnuExportPrevPal.Enabled = True
        imgExportPrevPal.Enabled = True: imgExportPrevPal.Picture = imgEnabled(3).Picture

        If LenB(txtOffset.Text) = 0 Then txtOffset.Text = String$(8, vbKey0)
        If optSearch.Value = True And Not blnCompressed Then
            Call SearchPalette
            Call txtChgPal_Change(0)
        ElseIf optOffset.Value And chkCompressed = vbUnchecked Then
            Offset = CLng("&H" & txtOffset.Text)
            Call ReadData
            mnuAddBookmarks.Enabled = True
            imgAddBookmarks.Enabled = True: imgAddBookmarks.Picture = imgEnabled(4).Picture
        ElseIf optSearch.Value = True And blnCompressed Then
            Call SearchCompressed
            Call txtChgPal_Change(0)
        Else
            Offset = CLng("&H" & txtOffset.Text)
            Call DecompressData
            mnuAddBookmarks.Enabled = True
            imgAddBookmarks.Enabled = True: imgAddBookmarks.Picture = imgEnabled(4).Picture
        End If
    End If

End Sub

Private Sub cmdReplacePalette_Click()

Dim Data As Byte, Data2 As Byte
Dim Answer As Byte

For i = txtPrevPal.LBound To txtPrevPal.UBound
    If txtPrevPal(i).BackColor = vbUnsafeColor Then
        Answer = MsgBox(LoadResString(1026), vbYesNo + vbExclamation)
        If Answer = vbNo Then
            Exit Sub
        Else
            Exit For
        End If
    End If
Next

    FileNum = FreeFile
    If LenB(sResult) > 0 Then
        Offset = CLng("&H" & txtOffset.Text)
        If Not blnCompressed Then
            X = 1
            SetAttr sResult, vbNormal
            Open sResult For Binary As #FileNum
            For i = txtChgPal.LBound To txtChgPal.UBound
                Data = Val("&H" & Left$(txtChgPal(i).Text, 2))
                Data2 = Val("&H" & Right$(txtChgPal(i).Text, 2))
                Put #FileNum, Offset + X + i, Data
                Put #FileNum, Offset + X + i + 1, Data2
                X = X + 1
            Next
            Close #FileNum
        Else
            Call RecompressData
        End If
        txtOffset.Tag = "0"
    End If

End Sub


Private Sub imgAddBookmarks_Click()
    Call mnuAddBookmarks_Click
End Sub

Private Sub imgClearPal_Click()
    Call mnuClearPal_Click
End Sub

Private Sub imgClearPal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuClearOptions
    End If
End Sub

Private Sub imgClearPal2_Click()
    Call mnuClearPal2_Click
End Sub

Private Sub imgClearPal2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuClearOptions2
    End If
End Sub

Private Sub imgCopyPal_Click()
    Call mnuCopyPal_Click
End Sub

Private Sub imgCopyPal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuCopyOptions
    End If
End Sub

Private Sub imgExportChgPal_Click()
    Call mnuExportChgPal_Click
End Sub

Private Sub imgExportPrevPal_Click()
    Call mnuExportPrevPal_Click
End Sub

Private Sub imgImportChgPal_Click()
    Call mnuImportChgPal_Click
End Sub

Private Sub imgImportPrevPal_Click()
    Call mnuImportPrevPal_Click
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, frmMain
End Sub

Private Sub mnuAddBookmarks_Click()
    frmBookmarks.Show
    Call frmBookmarks.LoadINI
    Call frmBookmarks.cmdSwitch_Click
    frmBookmarks.optAddType(0).Value = True
    If chkCompressed.Value = vbChecked Then frmBookmarks.txtDescription = CompressedPal
    frmBookmarks.txtDescription.SetFocus
    frmBookmarks.txtDescription.SelStart = Len(frmBookmarks.txtDescription)
End Sub

Private Sub mnuAllPalettes_Click(Index As Integer)
    mnuAllPalettes(0).Checked = Not mnuAllPalettes(0).Checked
    mnuAllPalettes(1).Checked = Not mnuAllPalettes(1).Checked
End Sub

Private Sub mnuAllPalettes2_Click(Index As Integer)
    mnuAllPalettes2(0).Checked = Not mnuAllPalettes2(0).Checked
    mnuAllPalettes2(1).Checked = Not mnuAllPalettes2(1).Checked
End Sub

Private Sub mnuBookmarks_Click()
    frmBookmarks.Show , Me
    Call frmBookmarks.LoadINI
    If LenB(cboLoadedBookmark.Tag) > 0 Then frmBookmarks.lstBookmark.ListIndex = Val(cboLoadedBookmark.Tag)
End Sub

Private Sub mnuCalc_Click()
    frmCalc.Show , Me
End Sub

Private Sub mnuClearPal_Click()
    
    If mnuAllPalettes(0).Checked Then
        
        Erase arrPalette
        Erase arrPalOffset
        Erase arrPalCompressed
    
    Else
        
        Dim molt As Long
        molt = 16 * CLng(txtPrevPalIndex.Text - 1)
        
        For i = molt To molt + 15
            arrPalette(i) = vbNullString
        Next
        
        arrPalOffset(txtPrevPalIndex - 1) = vbNullString
        arrPalCompressed(txtPrevPalIndex - 1) = 0
            
    End If
    
    blnClearing = True
    
    For i = txtPrevPal.LBound To txtPrevPal.UBound
        txtPrevPal(i).Text = vbNullString
        txtPrevPal(i).BackColor = vbWhite
    Next
    
    blnClearing = False
    
    mnuCopyPal.Enabled = False
    imgCopyPal.Enabled = False: imgCopyPal.Picture = imgDisabled(0).Picture
    mnuClearPal.Enabled = False
    imgClearPal.Enabled = False: imgClearPal.Picture = imgDisabled(1).Picture
    mnuExportPrevPal.Enabled = False
    imgExportPrevPal.Enabled = False: imgExportPrevPal.Picture = imgDisabled(3).Picture
    
End Sub

Private Sub mnuClearPal2_Click()

    If mnuAllPalettes2(0).Checked Then
        Erase arrPalette2
    Else
        
        Dim molt As Long
        molt = 16 * CLng(txtChgPalIndex.Text - 1)
        
        For i = molt To molt + 15
            arrPalette2(i) = vbNullString
        Next
        
    End If
    
    blnClearing2 = True
    
    For i = txtChgPal.LBound To txtChgPal.UBound
        txtChgPal(i).Text = vbNullString
        txtChgPal(i).BackColor = vbWhite
    Next
    
    blnClearing2 = False
    
    mnuClearPal2.Enabled = False
    imgClearPal2.Enabled = False: imgClearPal2.Picture = imgDisabled(1).Picture
    mnuExportChgPal.Enabled = False
    imgExportChgPal.Enabled = False: imgExportChgPal.Picture = imgDisabled(3).Picture

End Sub

Private Sub mnuColorPicker_Click()
    frmColorPicker.Show
End Sub

Private Sub mnuCopyAllPalettes_Click(Index As Integer)
    mnuCopyAllPalettes(0).Checked = Not mnuCopyAllPalettes(0).Checked
    mnuCopyAllPalettes(1).Checked = Not mnuCopyAllPalettes(1).Checked
End Sub

Private Sub mnuCopyPal_Click()
    
    blnCopying = True
    
    For i = txtPrevPal.LBound To txtPrevPal.UBound
        txtChgPal(i).Text = txtPrevPal(i).Text
    Next
    
    If mnuCopyAllPalettes(0).Checked Then
        For i = LBound(arrPalette) To UBound(arrPalette)
            arrPalette2(i) = arrPalette(i)
        Next
    End If
    
    blnCopying = False
    
    mnuClearPal2.Enabled = True
    imgClearPal2.Enabled = True: imgClearPal2.Picture = imgEnabled(1).Picture
    mnuImportChgPal.Enabled = True
    imgImportChgPal.Enabled = True: imgImportChgPal.Picture = imgEnabled(2).Picture
    mnuExportChgPal.Enabled = True
    imgExportChgPal.Enabled = True: imgExportChgPal.Picture = imgEnabled(3).Picture
    
    If optSearch.Value = True And txtOffset.Tag <> "0" Then
        cmdReplacePalette.Enabled = True
    ElseIf optOffset.Value = True And LenB(txtOffset.Text) > 0 Then
        cmdReplacePalette.Enabled = True
    Else
        cmdReplacePalette.Enabled = False
    End If
    
    'Call ChgIndexChange

End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExportChgPal_Click()
    Call ExportPalette(arrPalette2)
End Sub

Private Sub mnuExportPrevPal_Click()
    Call ExportPalette(arrPalette)
End Sub

Private Sub mnuGotoChg_Click()
Dim Number As String, molt As Long

    Number = InputBox(LoadResString(1022))

    If Not IsNumeric(Number) Or Val(Number) > 16 Or Val(Number) < 1 Then
        MsgBox LoadResString(1017), vbExclamation
        Exit Sub
    End If
    txtChgPalIndex.Text = Number
    
    molt = 16 * CLng(txtChgPalIndex.Text) - 1
    
    blnIgnoreChange2 = True
    X = 0
    For i = molt - 15 To molt
        txtChgPal(X).Text = arrPalette2(i)
        X = X + 1
    Next
    blnIgnoreChange2 = False
    
    If Val(txtChgPalIndex.Text) = 16 Then cmdNextChPal.Enabled = False: mnuNextChPal.Enabled = False
    cmdPrevChPal.Enabled = True
    mnuPrevChPal.Enabled = True
    blnIgnoreChange2 = False

End Sub

Private Sub mnuGotoPrev_Click()
Dim Number As String, molt As Long

    Number = InputBox(LoadResString(1022))
    
    If Not IsNumeric(Number) Or Val(Number) > 16 Or Val(Number) < 1 Then
        MsgBox LoadResString(1017), vbExclamation
        Exit Sub
    End If
    txtPrevPalIndex.Text = Number
    
    molt = 16 * CLng(txtPrevPalIndex.Tag) - 1
    blnIgnoreChange = True
    X = 0
    
    For i = molt - 15 To molt
        txtPrevPal(X).Text = arrPalette(i)
        X = X + 1
    Next
    
    blnIgnoreChange = False
    
    If LenB(arrPalOffset(Val(txtPrevPalIndex.Text) - 1)) > 0 Then
        txtOffset.Text = Right$(String$(7, vbKey0) & arrPalOffset(Val(txtPrevPalIndex.Text) - 1), txtOffset.MaxLength)
    End If
    
    If Val(txtPrevPalIndex.Text) = 16 Then cmdNextPrPal.Enabled = False: mnuNextPrPal.Enabled = False
    cmdPrevPrPal.Enabled = True
    mnuPrevPrPal.Enabled = True

End Sub

Private Sub mnuGradient_Click()
    frmGradient.Show , Me
End Sub

Private Sub mnuImportChgPal_Click()
    Dim InBuff As String, sImport As String, tmp As String, tmp2 As String, temp() As String

    FileNum = FreeFile
    
    Dim oOpenDialog As clsCommonDialog
    Set oOpenDialog = New clsCommonDialog
    
    sImport = oOpenDialog.ShowOpen(Me.hwnd, LoadResString(1020), , "APE Palette (*.gpl)|*.gpl|Adobe Color Table (*.act)|*.act|PaintShop Palette (*.pal)|*.pal|Tile Layer Pro Palette (*.tpl)|*.tpl|", FILEMUSTEXIST Or PATHMUSTEXIST Or HIDEREADONLY)

    LastFilter = oOpenDialog.FilterIndex

    If LenB(sImport) > 0 Then
    
        'Erase arrPalette2

        Select Case LastFilter
            Case 1
                On Error GoTo errorHandler
                Open sImport For Input As #FileNum
                    blnIgnoreChange2 = True
                    Input #FileNum, InBuff
                    For i = LBound(arrPalette2) To UBound(arrPalette2)
                        Input #FileNum, InBuff: arrPalette2(i) = Right$("000" & Hex$(InBuff), 4)
                        If i < txtChgPal.UBound + 1 Then txtChgPal(i).Text = arrPalette2(i)
                    Next
                Close FileNum

                txtChgPalIndex.Text = 1
                blnIgnoreChange2 = False

            Case 2
                On Error GoTo errorHandler
                Open sImport For Binary As #FileNum
                    tmp = Space$(LOF(FileNum))
                    Get #FileNum, , tmp

                    X = 1
                    blnIgnoreChange2 = True

                    For i = LBound(arrPalette2) To UBound(arrPalette2)
                        tmp2 = Mid$(tmp, X, 6)
                        arrPalette2(i) = RGB2GBA(Right$("0" & Hex$(Asc(Mid$(tmp2, 1, 2))), 2) & Right$("0" & Hex$(Asc(Mid$(tmp2, 2, 2))), 2) & Right("0" & Hex$(Asc(Mid$(tmp2, 3, 2))), 2))
                        If i < txtChgPal.UBound + 1 Then txtChgPal(i).Text = arrPalette2(i)
                        X = X + 3
                    Next

                    txtChgPalIndex.Text = 1
                    blnIgnoreChange2 = False

                Close FileNum
            Case 3
                On Error GoTo errorHandler
                Open sImport For Input As #FileNum
                    Input #FileNum, InBuff
                    Input #FileNum, InBuff
                    Input #FileNum, InBuff

                    blnIgnoreChange2 = True

                    For i = LBound(arrPalette2) To UBound(arrPalette2)
                        Input #FileNum, InBuff
                        temp = Split(InBuff, Space$(1), -1, vbBinaryCompare)
                        arrPalette2(i) = RGB2GBA(Right$("0" & Hex$(temp(0)), 2) & Right$("0" & Hex$(temp(1)), 2) & Right$("0" & Hex$(temp(2)), 2))
                        If i < txtChgPal.UBound + 1 Then txtChgPal(i).Text = arrPalette2(i)
                    Next

                    txtChgPalIndex.Text = 1
                    blnIgnoreChange2 = False

                Close FileNum
            Case 4
                On Error GoTo errorHandler
                Open sImport For Binary As #FileNum
                    blnIgnoreChange2 = True

                    tmp = Space$(LOF(FileNum))
                    Get #FileNum, , tmp

                    tmp = Mid$(tmp, 5)

                    X = 1
                    For i = LBound(arrPalette2) To UBound(arrPalette2)
                        tmp2 = Mid$(tmp, X, 6)
                        arrPalette2(i) = RGB2GBA(Right$("0" & Hex$(Asc(Mid$(tmp2, 1, 2))), 2) & Right$("0" & Hex$(Asc(Mid$(tmp2, 2, 2))), 2) & Right("0" & Hex$(Asc(Mid$(tmp2, 3, 2))), 2))
                        If i < txtChgPal.UBound + 1 Then txtChgPal(i).Text = arrPalette2(i)
                        X = X + 3
                    Next

                    txtChgPalIndex.Text = 1
                    blnIgnoreChange2 = False

                Close FileNum
        End Select

        If LenB(txtOffset.Text) > 0 Then cmdReplacePalette.Enabled = True

    End If

cmdNextChPal.Enabled = True
mnuNextChPal.Enabled = True

Exit Sub

errorHandler:
Call mnuClearPal2_Click
txtChgPalIndex.Text = vbNullString
cmdPrevChPal.Enabled = False
cmdNextChPal.Enabled = False
mnuPrevChPal.Enabled = False
mnuNextChPal.Enabled = False
MsgBox LoadResString(1018), vbCritical

End Sub

Private Sub mnuImportPrevPal_Click()
    Dim InBuff As String, sImport As String, tmp As String, tmp2 As String, temp() As String

    FileNum = FreeFile
    
    Dim oOpenDialog As clsCommonDialog
    Set oOpenDialog = New clsCommonDialog
    
    sImport = oOpenDialog.ShowOpen(Me.hwnd, LoadResString(1020), , "APE Palette (*.gpl)|*.gpl|Adobe Color Table (*.act)|*.act|PaintShop Palette (*.pal)|*.pal|Tile Layer Pro Palette (*.tpl)|*.tpl|", FILEMUSTEXIST Or PATHMUSTEXIST Or HIDEREADONLY)
    
    LastFilter = oOpenDialog.FilterIndex
    
    If LenB(sImport) > 0 Then
        
'        Erase arrPalette
        Erase arrPalOffset
        Erase arrPalCompressed
        txtOffset.Text = String$(8, vbKey0)
'        Offset = 0
                       
        Select Case LastFilter
            Case 1
                On Error GoTo errorHandler
                Open sImport For Input As #FileNum
                    blnIgnoreChange = True
                    
                    Input #FileNum, InBuff
                    For i = txtPrevPal.LBound To txtPrevPal.UBound
                        Input #FileNum, InBuff: arrPalette(i) = Right$("000" & Hex$(InBuff), 4)
                        If i < txtPrevPal.UBound + 1 Then txtPrevPal(i).Text = arrPalette(i)
                    Next
                    
                    txtPrevPalIndex.Text = 1
                    blnIgnoreChange = False
                Close FileNum
            Case 2
                On Error GoTo errorHandler
                Open sImport For Binary As #FileNum
                    tmp = Space$(LOF(FileNum))
                    Get #FileNum, , tmp
                    
                    X = 1
                    blnIgnoreChange = True
                    
                    For i = LBound(arrPalette) To UBound(arrPalette)
                        tmp2 = Mid$(tmp, X, 6)
                        arrPalette(i) = RGB2GBA(Right$("0" & Hex$(Asc(Mid$(tmp2, 1, 2))), 2) & Right$("0" & Hex$(Asc(Mid$(tmp2, 2, 2))), 2) & Right("0" & Hex$(Asc(Mid$(tmp2, 3, 2))), 2))
                        If i < txtPrevPal.UBound + 1 Then txtPrevPal(i).Text = arrPalette(i)
                        X = X + 3
                    Next
                    
                    txtPrevPalIndex.Text = 1
                    blnIgnoreChange = False
                    
                Close FileNum
            Case 3
                On Error GoTo errorHandler
                Open sImport For Input As #FileNum
                    Input #FileNum, InBuff
                    Input #FileNum, InBuff
                    Input #FileNum, InBuff
                    
                    blnIgnoreChange = True
                    
                    For i = LBound(arrPalette) To UBound(arrPalette)
                        Input #FileNum, InBuff
                        temp = Split(InBuff, Space$(1), -1, vbBinaryCompare)
                        arrPalette(i) = RGB2GBA(Right$("0" & Hex$(temp(0)), 2) & Right$("0" & Hex$(temp(1)), 2) & Right$("0" & Hex$(temp(2)), 2))
                        If i < txtPrevPal.UBound + 1 Then txtPrevPal(i).Text = arrPalette(i)
                    Next
                    
                    txtPrevPalIndex.Text = 1
                    blnIgnoreChange = False
                    
                Close FileNum
            Case 4
                On Error GoTo errorHandler
                Open sImport For Binary As #FileNum
                    blnIgnoreChange = True
                    
                    tmp = Space$(LOF(FileNum))
                    Get #FileNum, , tmp
                    
                    tmp = Mid$(tmp, 5)
                    
                    X = 1
                    For i = LBound(arrPalette) To UBound(arrPalette)
                        tmp2 = Mid$(tmp, X, 6)
                        arrPalette(i) = RGB2GBA(Right$("0" & Hex$(Asc(Mid$(tmp2, 1, 2))), 2) & Right$("0" & Hex$(Asc(Mid$(tmp2, 2, 2))), 2) & Right("0" & Hex$(Asc(Mid$(tmp2, 3, 2))), 2))
                        If i < txtPrevPal.UBound + 1 Then txtPrevPal(i).Text = arrPalette(i)
                        X = X + 3
                    Next
                    
                    txtPrevPalIndex.Text = 1
                    blnIgnoreChange = False
                Close FileNum
        End Select
        
        mnuCopyPal.Enabled = True
        imgCopyPal.Enabled = True: imgCopyPal.Picture = imgEnabled(0).Picture
        mnuClearPal.Enabled = True
        imgClearPal.Enabled = True: imgClearPal.Picture = imgEnabled(1).Picture
        mnuExportPrevPal.Enabled = True
        imgExportPrevPal.Enabled = True: imgExportPrevPal.Picture = imgEnabled(3).Picture
        cmdLoadPalette.Enabled = True
        
    End If

Exit Sub

errorHandler:
Call mnuClearPal_Click
txtPrevPalIndex.Text = vbNullString
cmdPrevPrPal.Enabled = False
cmdNextPrPal.Enabled = False
mnuPrevPrPal.Enabled = False
mnuNextPrPal.Enabled = False
MsgBox LoadResString(1018), vbCritical

End Sub


Private Sub mnuNextChPal_Click()
    Call cmdNextChPal_Click
End Sub

Private Sub mnuNextPal_Click()

    If LenB(sResult) > 0 Then
        If CLng("&H" & txtOffset.Text) > 31 Then mnuPrevPal.Enabled = True: cmdPrevPal.Enabled = True
        If CLng("&H" & txtOffset.Text) >= FileLength - 32 Then mnuPrevPal.Enabled = False: cmdPrevPal.Enabled = True
        If Not blnCompressed Then
            txtOffset.Text = Right$(String$(7, vbKey0) & Hex$(CLng("&H" & txtOffset.Text) + 32), txtOffset.MaxLength)
            Offset = CLng("&H" & txtOffset.Text)
            Call ReadData
        Else
            Call NavigateCompressed(True)
        End If
        
        cboLoadedBookmark.Text = vbNullString
        Call PrevIndexChange
        
    End If

End Sub


Private Sub mnuNextPrPal_Click()
    Call cmdNextPrPal_Click
End Sub

Private Sub mnuOpenROM_Click()
    Dim oOpenDialog As clsCommonDialog
    Set oOpenDialog = New clsCommonDialog

    sResult = oOpenDialog.ShowOpen(Me.hwnd, LoadResString(1024), , "GameBoy Advance ROMs (*.gba, *.agb, *.bin)|*.gba;*.agb;*.bin|", FILEMUSTEXIST Or PATHMUSTEXIST Or HIDEREADONLY)
    cmdLoadPalette.Enabled = False
    cmdReplacePalette.Enabled = False
    mnuImportPrevPal.Enabled = False
    imgImportPrevPal.Enabled = False: imgImportPrevPal.Picture = imgDisabled(2).Picture
    txtOffset.Text = vbNullString
    txtPrevPalIndex.Text = 1
    txtChgPalIndex.Text = 1
    cmdPrevPrPal.Enabled = False
    cmdNextPrPal.Enabled = False
    cmdPrevChPal.Enabled = False
    cmdNextChPal.Enabled = False
    mnuPrevPrPal.Enabled = False
    mnuNextPrPal.Enabled = False
    mnuPrevChPal.Enabled = False
    mnuNextChPal.Enabled = False
    mnuGotoPrev.Enabled = False
    mnuGotoChg.Enabled = False
    mnuAddBookmarks.Enabled = False
    imgAddBookmarks.Enabled = False: imgAddBookmarks.Picture = imgDisabled(4).Picture
    mnuCopyPal.Enabled = False
    imgCopyPal.Enabled = False: imgCopyPal.Picture = imgDisabled(0).Picture
    mnuExportPrevPal.Enabled = False
    imgExportPrevPal.Enabled = False: imgExportPrevPal.Picture = imgDisabled(3).Picture
    mnuClearPal.Enabled = False
    imgClearPal.Enabled = False: imgClearPal.Picture = imgDisabled(1).Picture
    mnuClearPal2.Enabled = False
    imgClearPal2.Enabled = False: imgClearPal2.Picture = imgDisabled(1).Picture
    ToggleControl (False)
    optOffset.Value = True
    cboLoadedBookmark.Tag = vbNullString
    cboLoadedBookmark.Text = vbNullString
    cboLoadedBookmark.Enabled = False
    cmdPrevBookmark.Enabled = False
    cmdNextBookmark.Enabled = False
    cmdLoadBookmark.Enabled = False
    lblROMName.Caption = LoadResString(1027)
    
    For i = txtPrevPal.LBound To txtPrevPal.UBound
        txtPrevPal(i).Text = vbNullString
        txtPrevPal(i).BackColor = vbWhite
        txtChgPal(i).Text = vbNullString
        txtChgPal(i).BackColor = vbWhite
    Next
    
    If LenB(sResult) > 0 Then
        ToggleControl (True)
        txtPrevPalIndex.Text = 1
        txtChgPalIndex.Text = 1
        cmdNextPrPal.Enabled = True
        cmdNextChPal.Enabled = True
        mnuNextPrPal.Enabled = True
        mnuNextChPal.Enabled = True
        mnuGotoPrev.Enabled = True
        mnuGotoChg.Enabled = True
        
        FileNum = FreeFile
        Open sResult For Binary As #FileNum
            FileLength = LOF(FileNum)
            Get #FileNum, &HAD, sGameCode
        Close #FileNum
        
        lblROMName.Caption = "ROM: [" & sGameCode & "] " & StrReverse(Mid$(StrReverse(sResult), 1, InStr(1, StrReverse(sResult), "\", vbBinaryCompare) - 1))
        
        Call frmBookmarks.LoadINIMain
        
    Else
        Unload frmBookmarks
    End If
    

End Sub


Private Sub LockDown(bLocked As Boolean, Color As Long, Color2 As Long)

    mnuImportPrevPal.Enabled = Not (bLocked)
    imgImportPrevPal.Enabled = Not (bLocked): If imgImportPrevPal.Enabled Then _
    imgImportPrevPal.Picture = imgEnabled(2).Picture Else imgImportPrevPal.Picture = imgDisabled(2).Picture
    txtOffset.Locked = Not (bLocked)
    txtOffset.ForeColor = Color2
    
    For i = txtPrevPal.LBound To txtPrevPal.UBound
        txtPrevPal(i).Locked = bLocked
        txtPrevPal(i).ForeColor = Color
    Next

End Sub

Private Sub mnuPrevChPal_Click()
    Call cmdPrevChPal_Click
End Sub

Private Sub mnuPrevPal_Click()

    If LenB(sResult) > 0 Then
        mnuNextPal.Enabled = True
        cmdNextPal.Enabled = True
        If CLng("&H" & txtOffset.Text) > 0 Then
            mnuPrevPal.Enabled = True
            cmdPrevPal.Enabled = True
        ElseIf CLng("&H" & txtOffset.Text) < 32 Then
            mnuPrevPal.Enabled = False
            cmdPrevPal.Enabled = False
        End If
        If Not blnCompressed Then
            txtOffset.Text = Right$(String$(7, vbKey0) & Hex$(CLng("&H" & txtOffset.Text) - 32), txtOffset.MaxLength)
            Offset = CLng("&H" & txtOffset.Text)
            Call ReadData
        Else
            Call NavigateCompressed(False)
        End If
        
        Call PrevIndexChange
        
    End If

End Sub


Private Sub mnuPrevPrPal_Click()
    Call cmdPrevPrPal_Click
End Sub

Private Sub mnuReadme_Click()
Shell "notepad.exe " & App.Path & "\Readme.txt", vbNormalFocus
End Sub

Private Sub optOffset_Click()

    Call LockDown(True, vbGrayText, 0)
    txtOffset.Tag = "0"
    txtOffset.Text = optOffset.Tag
    cmdLoadPalette.Caption = LoadResString(1005)
    
    If LenB(sResult) > 0 Then
        'Call mnuClearPal_Click
        mnuNextPal.Enabled = True
        cmdNextPal.Enabled = True
        mnuClearPal.Enabled = False
        imgClearPal.Enabled = False: imgClearPal.Picture = imgDisabled(1).Picture
        If LenB(txtOffset.Text) > 0 Then
            Call cmdLoadPalette_Click
            Call txtPrevPalIndex_Change
            Call txtChgPal_Change(0)
        End If
    End If

End Sub


Private Sub optSearch_Click()

    Call LockDown(False, 0, vbGrayText)
    cmdLoadPalette.Caption = LoadResString(1023)
    optOffset.Tag = txtOffset.Text
    txtOffset.Text = String$(8, vbKey0)
        
    If LenB(sResult) > 0 Then
        mnuPrevPal.Enabled = False
        cmdPrevPal.Enabled = False
        mnuNextPal.Enabled = False
        cmdNextPal.Enabled = False
    End If
    
    For i = txtPrevPal.LBound To txtPrevPal.UBound
        If txtPrevPal(i).Text = vbNullString Then
            mnuClearPal.Enabled = False
            imgClearPal.Enabled = False: imgClearPal.Picture = imgDisabled(1).Picture
        Else
            mnuClearPal.Enabled = True
            imgClearPal.Enabled = True: imgClearPal.Picture = imgEnabled(1).Picture
            Exit For
        End If
    Next

End Sub


Private Sub picChgPreview_DblClick(Index As Integer)

Dim Answer As Byte

    If LenB(txtChgPal(Index).Text) > 0 Then
        If txtChgPal(Index).BackColor = vbUnsafeColor Then
            Answer = MsgBox(LoadResString(1025), vbYesNo + vbExclamation)
            If Answer = vbYes Then
                frmColorPicker.picNewColor.BackColor = picChgPreview(Index).BackColor
                frmColorPicker.tmrTransfer.Enabled = True
                frmColorPicker.Tag = "Ch" & Index
            End If
        Else
            frmColorPicker.picNewColor.BackColor = picChgPreview(Index).BackColor
            frmColorPicker.tmrTransfer.Enabled = True
            frmColorPicker.Tag = "Ch" & Index
        End If
    End If

End Sub

Private Sub picPrevPreview_DblClick(Index As Integer)
    
    If LenB(txtPrevPal(Index).Text) > 0 And Not optOffset Then
        frmColorPicker.picNewColor.BackColor = picPrevPreview(Index).BackColor
        frmColorPicker.tmrTransfer.Enabled = True
        frmColorPicker.Tag = "Pr" & Index
    End If
    
End Sub

Private Sub txtChgPal_Change(Index As Integer)
    
    If Not IsHex(txtChgPal(Index).Text) And Not blnClearing2 Then
        txtChgPal(Index).Text = vbNullString
        txtChgPal(Index).BackColor = vbWhite
        picChgPreview(Index).BackColor = vbWhite
        Exit Sub
    End If
    
    Dim tmp As Long

    If Len(txtChgPal(Index).Text) = 4 Then
        picChgPreview(Index).BackColor = "&H" & GB2RGB(txtChgPal(Index).Text, True)
    Else
        picChgPreview(Index).BackColor = vbWhite
    End If
    
    If Not SafeCheck(txtChgPal(Index).Text) Then
        txtChgPal(Index).BackColor = vbUnsafeColor
    Else
        txtChgPal(Index).BackColor = vbWhite
    End If
    
    mnuClearPal2.Enabled = True
    imgClearPal2.Enabled = True: imgClearPal2.Picture = imgEnabled(1).Picture
    
    If Not blnIgnoreChange2 And LenB(txtChgPalIndex.Text) > 0 Then
        tmp = 16 * CLng(txtChgPalIndex.Text) - 16
        arrPalette2(tmp + Index) = txtChgPal(Index).Text
    End If
    
    If blnCopying Then Exit Sub
    If blnIgnoreChange2 Then Exit Sub
    
    For i = txtChgPal.LBound To txtChgPal.UBound
        If Len(txtChgPal(i).Text) < 4 Then
            cmdReplacePalette.Enabled = False
            mnuExportChgPal.Enabled = False
            imgExportChgPal.Enabled = False: imgExportChgPal.Picture = imgDisabled(3).Picture
            Exit For
        Else
            If optSearch.Value = True And txtOffset.Tag <> "0" Then
                cmdReplacePalette.Enabled = True
            ElseIf optOffset.Value = True And LenB(txtOffset.Text) > 0 Then
                cmdReplacePalette.Enabled = True
            Else
                cmdReplacePalette.Enabled = False
            End If
            mnuExportChgPal.Enabled = True
            imgExportChgPal.Enabled = True: imgExportChgPal.Picture = imgEnabled(3).Picture
        End If
    Next

End Sub

Private Sub txtChgPal_KeyPress(Index As Integer, KeyCode As Integer)
        If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
        Select Case KeyCode
            Case vbKey0 To vbKey9
            Case vbKeyA To vbKeyF
            Case 97 To 102
                KeyCode = KeyCode - 32
            Case Else
                KeyCode = 0: Exit Sub
        End Select
        If Len(txtChgPal(Index).Text) = 4 And Index < txtChgPal.UBound Then
                If LenB(txtChgPal(Index + 1).Text) = 0 Then SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub txtOffset_Change()

    On Error GoTo errHandler
       
    If LenB(txtOffset.Text) > 0 And optOffset.Value = True Then
        cmdLoadPalette.Enabled = True
    ElseIf LenB(txtOffset.Text) = 0 And optOffset = True Then
        cmdLoadPalette.Enabled = False
    End If
    
    mnuPrevPal.Enabled = False
    cmdPrevPal.Enabled = False
    mnuNextPal.Enabled = False
    cmdNextPal.Enabled = False
    
    If txtOffset.Text = String$(8, vbKey0) Then
        mnuAddBookmarks.Enabled = False
        imgAddBookmarks.Enabled = False: imgAddBookmarks.Picture = imgDisabled(4).Picture
    Else
        mnuAddBookmarks.Enabled = True
        imgAddBookmarks.Enabled = True: imgAddBookmarks.Picture = imgEnabled(4).Picture
    End If
    
    If LenB(txtOffset.Text) > 0 And optOffset.Value = True Then
        mnuNextPal.Enabled = True
        cmdNextPal.Enabled = True
        If CLng("&H" & txtOffset.Text) > 31 Then mnuPrevPal.Enabled = True: cmdPrevPal.Enabled = True
    End If
      
    If CLng("&H" & txtOffset.Text) > FileLength - 32 Then
        mnuNextPal.Enabled = False
        cmdNextPal.Enabled = False
        txtOffset.Text = Right$(String$(7, vbKey0) & Hex$(FileLength - 32), txtOffset.MaxLength)
        Me.SetFocus
        txtOffset.Enabled = False
        txtOffset.Enabled = True
        mnuNextPal.Enabled = False
        cmdNextPal.Enabled = False
    End If
    
    Exit Sub

errHandler:
If Err.Number = 13 And LenB(txtOffset.Text) > 0 Then
    MsgBox LoadResString(1019), vbExclamation
    txtOffset.Text = vbNullString
End If

End Sub

Private Sub txtOffset_KeyPress(KeyCode As Integer)
If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
    If (KeyCode > 64 And KeyCode < vbKeyG) Then Exit Sub
    If (KeyCode > 96 And KeyCode < 103) Then KeyCode = KeyCode - 32: Exit Sub
    If (KeyCode > 47 And KeyCode < 58) Then Exit Sub
    KeyCode = 0
End If
End Sub

Private Sub txtPrevPal_Change(Index As Integer)
    
    If Not IsHex(txtPrevPal(Index).Text) And Not blnClearing Then
        txtPrevPal(Index).Text = vbNullString
        txtPrevPal(Index).BackColor = vbWhite
        picPrevPreview(Index).BackColor = vbWhite
        Exit Sub
    End If
    
    Dim tmp As Long
    
    If Len(txtPrevPal(Index).Text) = 4 Then
        picPrevPreview(Index).BackColor = "&H" & GB2RGB(txtPrevPal(Index).Text, True)
    Else
        picPrevPreview(Index).BackColor = vbWhite
    End If
    
    If Not SafeCheck(txtPrevPal(Index).Text) Then
        txtPrevPal(Index).BackColor = vbUnsafeColor
    Else
        txtPrevPal(Index).BackColor = vbWhite
    End If
    
'    If Not blnIgnoreChange And LenB(txtPrevPalIndex.Text) > 0 Then
'        tmp = 16 * CLng(txtPrevPalIndex.Tag) - 16
'        arrPalette(tmp + Index) = txtPrevPal(Index).Text
'    End If
    
    If optSearch.Value Then
        mnuClearPal.Enabled = True
        imgClearPal.Enabled = True: imgClearPal.Picture = imgEnabled(1).Picture
        mnuCopyPal.Enabled = True
        imgCopyPal.Enabled = True: imgCopyPal.Picture = imgEnabled(0).Picture
        If blnIgnoreChange Then Exit Sub
        If Len(txtPrevPal(0).Text) = 4 And Not blnIgnoreChange Then
            cmdLoadPalette.Enabled = True
            mnuExportPrevPal.Enabled = False
            imgExportPrevPal.Enabled = False: imgExportPrevPal.Picture = imgDisabled(3).Picture
        ElseIf Not blnIgnoreChange Then
            cmdLoadPalette.Enabled = False
        End If
    End If
    
    If optSearch Then
        mnuClearPal.Enabled = True
        imgClearPal.Enabled = True: imgClearPal.Picture = imgEnabled(1).Picture
        mnuCopyPal.Enabled = True
        imgCopyPal.Enabled = True: imgCopyPal.Picture = imgEnabled(0).Picture
        If blnIgnoreChange Then Exit Sub

        For i = txtPrevPal.LBound To txtPrevPal.UBound
            If Len(txtPrevPal(i).Text) < 4 Then
                cmdLoadPalette.Enabled = False
                mnuExportPrevPal.Enabled = False
                imgExportPrevPal.Enabled = False: imgExportPrevPal.Picture = imgDisabled(3).Picture
                Exit For
            Else
                cmdLoadPalette.Enabled = True
                mnuExportPrevPal.Enabled = True
                imgExportPrevPal.Enabled = True: imgExportPrevPal.Picture = imgEnabled(3).Picture
            End If
        Next
        
        If Len(txtPrevPal(0).Text) = 4 Then
            cmdLoadPalette.Enabled = True
        Else
            cmdLoadPalette.Enabled = False
        End If
        
    End If

End Sub

Private Sub txtPrevPal_KeyPress(Index As Integer, KeyCode As Integer)
    If KeyCode <> vbKeyBack And KeyCode <> 22 And KeyCode <> 3 And KeyCode <> 24 Then
        Select Case KeyCode
            Case vbKey0 To vbKey9
            Case vbKeyA To vbKeyF
            Case 97 To 102
                KeyCode = KeyCode - 32
            Case Else
                KeyCode = 0: Exit Sub
        End Select
        If Len(txtPrevPal(Index).Text) = 4 And Index < txtPrevPal.UBound Then
                If LenB(txtPrevPal(Index + 1).Text) = 0 Then SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub txtPrevPalIndex_Change()
    
    If Val(txtPrevPalIndex.Text) = 1 Then
        cmdPrevPrPal.Enabled = False
        mnuPrevPrPal.Enabled = False
        cmdNextPrPal.Enabled = True
        mnuNextPrPal.Enabled = True
    ElseIf Val(txtPrevPalIndex.Text) = 16 Then
        cmdPrevPrPal.Enabled = True
        mnuPrevPrPal.Enabled = True
        cmdNextPrPal.Enabled = False
        mnuNextPrPal.Enabled = False
    Else
        cmdPrevPrPal.Enabled = True
        mnuPrevPrPal.Enabled = True
        cmdNextPrPal.Enabled = True
        mnuNextPrPal.Enabled = True
    End If
    
    If LenB(txtPrevPalIndex.Text) > 0 Then
        txtPrevPalIndex.Tag = txtPrevPalIndex.Text
    Else
        txtPrevPalIndex.Tag = vbNullString
    End If
    
    
End Sub
