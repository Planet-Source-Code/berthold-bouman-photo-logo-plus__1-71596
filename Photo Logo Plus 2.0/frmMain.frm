VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Photo Logo Plus 2.0"
   ClientHeight    =   10575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15045
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   705
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1003
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picNormalBlend 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11745
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   79
      Top             =   12210
      Width           =   675
   End
   Begin MSComDlg.CommonDialog DialogMain 
      Left            =   11115
      Top             =   12210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrBlink 
      Interval        =   250
      Left            =   3345
      Top             =   10890
   End
   Begin VB.CheckBox chkLogoIndicator 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3540
      TabIndex        =   69
      Top             =   9690
      Value           =   1  'Checked
      Width           =   195
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   8595
      Top             =   12165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrCurTime 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2895
      Top             =   10890
   End
   Begin VB.PictureBox picPort 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   3210
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   12
      Top             =   1185
      Width           =   7710
      Begin VB.PictureBox picSource 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   9570
         Left            =   405
         MousePointer    =   99  'Custom
         ScaleHeight     =   638
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   634
         TabIndex        =   13
         Top             =   360
         Width           =   9510
         Begin VB.PictureBox picMaskedLogo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   135
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   185
            TabIndex        =   47
            Top             =   165
            Width           =   2775
         End
         Begin VB.PictureBox picNormalLogo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   2985
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   185
            TabIndex        =   46
            Top             =   165
            Width           =   2775
         End
      End
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9195
      TabIndex        =   11
      Top             =   12210
      Width           =   1710
   End
   Begin VB.PictureBox picThumbFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2475
      Left            =   5460
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   214
      TabIndex        =   10
      Top             =   7395
      Width           =   3210
      Begin VB.PictureBox picThumb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   240
         ScaleHeight     =   97
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   106
         TabIndex        =   68
         Top             =   225
         Width           =   1590
         Begin VB.Shape thumbViewPort 
            BorderColor     =   &H0000FFFF&
            FillColor       =   &H00FFFFFF&
            Height          =   690
            Left            =   270
            Top             =   300
            Width           =   690
         End
         Begin VB.Shape shpLogoPos 
            BorderColor     =   &H0000FF00&
            FillColor       =   &H00FFFFFF&
            Height          =   135
            Left            =   0
            Top             =   0
            Width           =   285
         End
      End
   End
   Begin VB.PictureBox picCentre 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3150
      Picture         =   "frmMain.frx":1A820
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   522
      TabIndex        =   7
      Top             =   6930
      Width           =   7830
      Begin VB.Label lblLogoText 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   105
         TabIndex        =   38
         Top             =   30
         Width           =   7680
      End
   End
   Begin VB.PictureBox picSide 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8925
      Left            =   11055
      ScaleHeight     =   595
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   3
      Top             =   1110
      Width           =   3915
      Begin VB.PictureBox picHide1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   3135
         ScaleHeight     =   60
         ScaleWidth      =   690
         TabIndex        =   78
         Top             =   885
         Width           =   690
      End
      Begin VB.PictureBox picHide2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3810
         ScaleHeight     =   375
         ScaleWidth      =   75
         TabIndex        =   77
         Top             =   525
         Width           =   75
      End
      Begin VB.OptionButton optFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "[00000]"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2265
         TabIndex        =   67
         ToolTipText     =   " Serial Number Format "
         Top             =   4605
         Width           =   990
      End
      Begin VB.OptionButton optFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "[0000]"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1350
         TabIndex        =   66
         ToolTipText     =   " Serial Number Format "
         Top             =   4605
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optFormat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "[000]"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   65
         ToolTipText     =   " Serial Number Format "
         Top             =   4605
         Width           =   795
      End
      Begin VB.PictureBox portMaskedLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   45
         ScaleHeight     =   58
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   252
         TabIndex        =   44
         Top             =   7950
         Width           =   3810
         Begin MSComCtl2.FlatScrollBar maskLogoVerScroll 
            Height          =   765
            Left            =   3645
            TabIndex        =   53
            Top             =   0
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1349
            _Version        =   393216
            Orientation     =   1245184
         End
         Begin MSComCtl2.FlatScrollBar maskLogoHorScroll 
            Height          =   135
            Left            =   0
            TabIndex        =   52
            Top             =   735
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   238
            _Version        =   393216
            Arrows          =   65536
            LargeChange     =   10
            Orientation     =   1245185
         End
         Begin VB.PictureBox picFillMask 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   3630
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   57
            Top             =   750
            Width           =   135
         End
         Begin VB.PictureBox picMaskedSource 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   185
            TabIndex        =   45
            ToolTipText     =   " Double-Click to view Logo Fullsize "
            Top             =   0
            Width           =   2775
         End
      End
      Begin VB.PictureBox portNormalLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   45
         ScaleHeight     =   58
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   252
         TabIndex        =   42
         Top             =   5940
         Width           =   3810
         Begin MSComCtl2.FlatScrollBar normLogoHorScroll 
            Height          =   135
            Left            =   0
            TabIndex        =   54
            Top             =   735
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   238
            _Version        =   393216
            Arrows          =   65536
            LargeChange     =   10
            Orientation     =   1245185
         End
         Begin MSComCtl2.FlatScrollBar normLogoVerScroll 
            Height          =   750
            Left            =   3645
            TabIndex        =   55
            Top             =   0
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   1323
            _Version        =   393216
            LargeChange     =   10
            Orientation     =   1245184
         End
         Begin VB.PictureBox picFillNorm 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   120
            Left            =   3645
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   56
            Top             =   750
            Width           =   120
         End
         Begin VB.PictureBox picNormalSource 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   0
            ScaleHeight     =   41
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   185
            TabIndex        =   43
            ToolTipText     =   " Double-Click to view Logo Fullsize "
            Top             =   0
            Width           =   2775
         End
      End
      Begin VB.TextBox txtAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   555
         TabIndex        =   33
         Text            =   "©"
         Top             =   2895
         Width           =   1500
      End
      Begin VB.CheckBox chkAddText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   32
         Top             =   2970
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkAddText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   31
         Top             =   990
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkAddText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   30
         Top             =   3405
         Width           =   195
      End
      Begin VB.CheckBox chkAddText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   29
         Top             =   3840
         Width           =   195
      End
      Begin VB.CheckBox chkAddText 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "   Add Serial"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   28
         Top             =   4260
         Width           =   195
      End
      Begin VB.CheckBox chkCurDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   2265
         TabIndex        =   27
         Top             =   3330
         Width           =   195
      End
      Begin VB.CheckBox chkCurTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Current Time"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   2265
         TabIndex        =   26
         Top             =   3765
         Width           =   195
      End
      Begin VB.CheckBox chkAutoIncrement 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Auto Increment"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   2265
         TabIndex        =   25
         Top             =   4185
         Width           =   195
      End
      Begin VB.TextBox txtAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   1
         Left            =   540
         TabIndex        =   24
         Text            =   "Photo Logo Plus"
         Top             =   975
         Width           =   3300
      End
      Begin VB.TextBox txtAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   555
         TabIndex        =   23
         Text            =   "My Date"
         Top             =   3330
         Width           =   1500
      End
      Begin VB.TextBox txtAdd 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   4
         Left            =   540
         TabIndex        =   21
         Text            =   "My Serial"
         Top             =   4200
         Width           =   1500
      End
      Begin VB.CheckBox chkShadow 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "  Shadow On/Off"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   180
         TabIndex        =   20
         Top             =   2460
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.ComboBox cboFontSize 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3150
         TabIndex        =   19
         ToolTipText     =   " Change Fontsize"
         Top             =   555
         Width           =   675
      End
      Begin VB.TextBox txtFontName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   540
         TabIndex        =   18
         Text            =   "Trebuchet MS"
         Top             =   555
         Width           =   1725
      End
      Begin VB.PictureBox picShadowColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1365
         MousePointer    =   99  'Custom
         ScaleHeight     =   390
         ScaleWidth      =   390
         TabIndex        =   17
         ToolTipText     =   " Change Shadow Colour "
         Top             =   1965
         Width           =   420
      End
      Begin VB.PictureBox picTextColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   495
         MousePointer    =   99  'Custom
         ScaleHeight     =   390
         ScaleWidth      =   390
         TabIndex        =   16
         ToolTipText     =   " Change Text Colour "
         Top             =   1965
         Width           =   420
      End
      Begin VB.PictureBox picSideBar3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   0
         Picture         =   "frmMain.frx":222E2
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   6
         Top             =   6915
         Width           =   3915
         Begin VB.OptionButton optLogo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   50
            Top             =   165
            Width           =   195
         End
         Begin VB.Label lblMaskedLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "Add Masked Image Logo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   975
            TabIndex        =   41
            Top             =   135
            Width           =   1995
         End
      End
      Begin VB.PictureBox picSideBar2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   0
         Picture         =   "frmMain.frx":28214
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   5
         Top             =   4905
         Width           =   3915
         Begin VB.OptionButton optLogo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   49
            Top             =   165
            Width           =   195
         End
         Begin VB.Label lblNormalLogo 
            BackStyle       =   0  'Transparent
            Caption         =   "Add Normal Image Logo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1020
            TabIndex        =   40
            Top             =   135
            Width           =   1965
         End
      End
      Begin VB.PictureBox picSideBar1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   0
         Picture         =   "frmMain.frx":2E146
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   261
         TabIndex        =   4
         Top             =   0
         Width           =   3915
         Begin VB.OptionButton optLogo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   48
            Top             =   165
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.Label lblPlainText 
            BackStyle       =   0  'Transparent
            Caption         =   "Add Plain Text Logo"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   1140
            TabIndex        =   39
            Top             =   135
            Width           =   1665
         End
      End
      Begin MSComCtl2.UpDown udcShadowOffset 
         Height          =   255
         Left            =   1830
         TabIndex        =   34
         Top             =   2520
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         _Version        =   393216
         Value           =   1
         Max             =   4
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtOffset 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2085
         TabIndex        =   35
         Text            =   "1"
         Top             =   2505
         Width           =   405
      End
      Begin VB.TextBox txtAdd 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   555
         TabIndex        =   22
         Text            =   "My Time"
         Top             =   3765
         Width           =   1500
      End
      Begin VB.Image imgNormLogo 
         Height          =   420
         Index           =   0
         Left            =   600
         ToolTipText     =   " Open Normal Image Logo "
         Top             =   5445
         Width           =   420
      End
      Begin VB.Label lblShadow 
         BackStyle       =   0  'Transparent
         Caption         =   "Shadow On/Off"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   510
         TabIndex        =   74
         Top             =   2535
         Width           =   1320
      End
      Begin VB.Label lblAutoInrement 
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Inrement"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2565
         TabIndex        =   73
         Top             =   4260
         Width           =   1185
      End
      Begin VB.Label lblCurTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Time"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2565
         TabIndex        =   72
         Top             =   3825
         Width           =   1185
      End
      Begin VB.Label lblCurDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Date"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2565
         TabIndex        =   71
         Top             =   3390
         Width           =   1185
      End
      Begin VB.Image imgMaskedLogo 
         Height          =   420
         Index           =   6
         Left            =   3435
         ToolTipText     =   " Invert Masked Image Logo "
         Top             =   7455
         Width           =   420
      End
      Begin VB.Image imgMaskedLogo 
         Height          =   420
         Index           =   5
         Left            =   3030
         ToolTipText     =   " Flip Masked Image Logo "
         Top             =   7455
         Width           =   420
      End
      Begin VB.Image imgMaskedLogo 
         Height          =   420
         Index           =   4
         Left            =   2625
         ToolTipText     =   " View Masked Image Logo "
         Top             =   7455
         Width           =   420
      End
      Begin VB.Image imgMaskedLogo 
         Height          =   420
         Index           =   3
         Left            =   2220
         ToolTipText     =   " Save Masked Image Logo "
         Top             =   7455
         Width           =   420
      End
      Begin VB.Image imgMaskedLogo 
         Height          =   420
         Index           =   2
         Left            =   1815
         ToolTipText     =   " Paste Masked Image Logo from Clipboard "
         Top             =   7455
         Width           =   420
      End
      Begin VB.Image imgMaskedLogo 
         Height          =   420
         Index           =   1
         Left            =   1410
         ToolTipText     =   " Restore Masked Image Logo "
         Top             =   7455
         Width           =   420
      End
      Begin VB.Image imgMaskedLogo 
         Height          =   420
         Index           =   0
         Left            =   1005
         ToolTipText     =   " Open Masked Image Logo "
         Top             =   7455
         Width           =   420
      End
      Begin VB.Image imgNormLogo 
         Height          =   420
         Index           =   7
         Left            =   3435
         ToolTipText     =   " Greyscale Normal Image Logo "
         Top             =   5445
         Width           =   420
      End
      Begin VB.Image imgNormLogo 
         Height          =   420
         Index           =   6
         Left            =   3030
         ToolTipText     =   " Invert Normal Image Logo "
         Top             =   5445
         Width           =   420
      End
      Begin VB.Image imgNormLogo 
         Height          =   420
         Index           =   5
         Left            =   2625
         ToolTipText     =   " Flip Normal Image Logo "
         Top             =   5445
         Width           =   420
      End
      Begin VB.Image imgNormLogo 
         Height          =   420
         Index           =   4
         Left            =   2220
         ToolTipText     =   " View Normal Image Logo "
         Top             =   5445
         Width           =   420
      End
      Begin VB.Image imgNormLogo 
         Height          =   420
         Index           =   3
         Left            =   1815
         ToolTipText     =   " Save Normal Image Logo "
         Top             =   5445
         Width           =   420
      End
      Begin VB.Image imgNormLogo 
         Height          =   420
         Index           =   2
         Left            =   1410
         ToolTipText     =   " Paste Normal Image Logo from Clipboard "
         Top             =   5445
         Width           =   420
      End
      Begin VB.Image imgNormLogo 
         Height          =   420
         Index           =   1
         Left            =   1005
         ToolTipText     =   " Restore Normal Image Logo "
         Top             =   5445
         Width           =   420
      End
      Begin VB.Image imgFontButton 
         Height          =   420
         Index           =   7
         Left            =   60
         MousePointer    =   99  'Custom
         ToolTipText     =   " Change Font "
         Top             =   510
         Width           =   420
      End
      Begin VB.Image imgFontButton 
         Height          =   420
         Index           =   6
         Left            =   3420
         MousePointer    =   99  'Custom
         ToolTipText     =   " Font Striked "
         Top             =   1965
         Width           =   420
      End
      Begin VB.Image imgFontButton 
         Height          =   420
         Index           =   5
         Left            =   3015
         MousePointer    =   99  'Custom
         ToolTipText     =   " Font Underlined "
         Top             =   1965
         Width           =   420
      End
      Begin VB.Image imgFontButton 
         Height          =   420
         Index           =   4
         Left            =   2610
         MousePointer    =   99  'Custom
         ToolTipText     =   " Font Italic "
         Top             =   1965
         Width           =   420
      End
      Begin VB.Image imgFontButton 
         Height          =   420
         Index           =   3
         Left            =   2205
         MousePointer    =   99  'Custom
         ToolTipText     =   " Font Bold "
         Top             =   1965
         Width           =   420
      End
      Begin VB.Image imgFontButton 
         Height          =   420
         Index           =   2
         Left            =   1800
         MousePointer    =   99  'Custom
         ToolTipText     =   " Change Shadow Colour "
         Top             =   1965
         Width           =   420
      End
      Begin VB.Image imgFontButton 
         Height          =   420
         Index           =   1
         Left            =   930
         MousePointer    =   99  'Custom
         ToolTipText     =   " Swap Font / Shadow colour "
         Top             =   1965
         Width           =   420
      End
      Begin VB.Label lblOffSet 
         BackStyle       =   0  'Transparent
         Caption         =   "Shadow Offset"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2640
         TabIndex        =   37
         Top             =   2535
         Width           =   1320
      End
      Begin VB.Image imgFontButton 
         Height          =   420
         Index           =   0
         Left            =   60
         MousePointer    =   99  'Custom
         ToolTipText     =   " Change Text Colour "
         Top             =   1965
         Width           =   420
      End
      Begin VB.Label lblAddSymbol 
         BackStyle       =   0  'Transparent
         Caption         =   "Add © Symbol"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2430
         TabIndex        =   36
         Top             =   2970
         Width           =   1185
      End
   End
   Begin VB.PictureBox picMenuBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      Picture         =   "frmMain.frx":34078
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1003
      TabIndex        =   2
      Top             =   465
      Width           =   15045
      Begin VB.Image imgMenu 
         Height          =   555
         Index           =   6
         Left            =   6075
         ToolTipText     =   " Open Help "
         Top             =   45
         Width           =   960
      End
      Begin VB.Image imgMenu 
         Height          =   555
         Index           =   5
         Left            =   5085
         ToolTipText     =   " Batch Convert "
         Top             =   45
         Width           =   960
      End
      Begin VB.Image imgMenu 
         Height          =   555
         Index           =   4
         Left            =   4095
         ToolTipText     =   " Batch Process  "
         Top             =   45
         Width           =   960
      End
      Begin VB.Image imgMenu 
         Height          =   555
         Index           =   3
         Left            =   3105
         ToolTipText     =   " Show Full Size Photo "
         Top             =   45
         Width           =   960
      End
      Begin VB.Image imgMenu 
         Height          =   555
         Index           =   2
         Left            =   2115
         ToolTipText     =   " Save Photo "
         Top             =   45
         Width           =   960
      End
      Begin VB.Image imgMenu 
         Height          =   555
         Index           =   1
         Left            =   1125
         ToolTipText     =   " Paste Image from Clipboard "
         Top             =   45
         Width           =   960
      End
      Begin VB.Image imgMenu 
         Height          =   555
         Index           =   0
         Left            =   135
         ToolTipText     =   " Open Folder "
         Top             =   45
         Width           =   960
      End
   End
   Begin VB.PictureBox picBot 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      Picture         =   "frmMain.frx":5466A
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1003
      TabIndex        =   1
      Top             =   10035
      Width           =   15045
      Begin VB.Image imgBotSplit4 
         Height          =   465
         Left            =   13665
         Picture         =   "frmMain.frx":6EE3C
         Top             =   0
         Width           =   90
      End
      Begin VB.Image imgBotSplit3 
         Height          =   465
         Left            =   12405
         Picture         =   "frmMain.frx":6F0EA
         Top             =   0
         Width           =   90
      End
      Begin VB.Image imgBotSplit2 
         Height          =   465
         Left            =   10980
         Picture         =   "frmMain.frx":6F398
         Top             =   0
         Width           =   90
      End
      Begin VB.Image imgBotSplit1 
         Height          =   465
         Left            =   3075
         Picture         =   "frmMain.frx":6F646
         Top             =   0
         Width           =   90
      End
      Begin VB.Image imgFile 
         Height          =   240
         Left            =   3255
         Picture         =   "frmMain.frx":6F8F4
         Top             =   135
         Width           =   240
      End
      Begin VB.Image imgFolder 
         Height          =   240
         Left            =   180
         Picture         =   "frmMain.frx":6FCF7
         Top             =   150
         Width           =   240
      End
      Begin VB.Label lblSource 
         BackStyle       =   0  'Transparent
         Caption         =   "Images in this folder:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   480
         TabIndex        =   64
         Top             =   135
         Width           =   2490
      End
      Begin VB.Label lblYpos 
         BackStyle       =   0  'Transparent
         Caption         =   "Y Position:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   13785
         TabIndex        =   61
         Top             =   135
         Width           =   1155
      End
      Begin VB.Label lblXpos 
         BackStyle       =   0  'Transparent
         Caption         =   "X Position:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   12540
         TabIndex        =   60
         Top             =   135
         Width           =   1155
      End
      Begin VB.Label lblLogoType 
         BackStyle       =   0  'Transparent
         Caption         =   "Masked Image Logo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   11130
         TabIndex        =   51
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "File:"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3525
         TabIndex        =   15
         Top             =   135
         Width           =   7440
      End
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      Picture         =   "frmMain.frx":700B8
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1003
      TabIndex        =   0
      Top             =   0
      Width           =   15045
      Begin VB.Image imgClose 
         Height          =   285
         Left            =   14640
         Picture         =   "frmMain.frx":86DB6
         Top             =   105
         Width           =   285
      End
      Begin VB.Image imgMin 
         Height          =   285
         Left            =   14295
         Picture         =   "frmMain.frx":8726C
         Top             =   105
         Width           =   285
      End
   End
   Begin MSComCtl2.FlatScrollBar scrSourceVer 
      Height          =   1290
      Left            =   4965
      TabIndex        =   8
      Top             =   7860
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   2275
      _Version        =   393216
      Orientation     =   1245184
   End
   Begin MSComCtl2.FlatScrollBar scrSourceHor 
      Height          =   90
      Left            =   3555
      TabIndex        =   9
      Top             =   9255
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   159
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin MSComCtl2.FlatScrollBar scrLogoVer 
      Height          =   1275
      Left            =   10500
      TabIndex        =   58
      Top             =   7860
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   2249
      _Version        =   393216
      Orientation     =   1245184
   End
   Begin MSComCtl2.FlatScrollBar scrLogoHor 
      Height          =   90
      Left            =   9090
      TabIndex        =   59
      Top             =   9255
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   159
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7905
      Top             =   12105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":87722
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":87B4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":87F63
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":88387
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":887A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvFiles 
      Height          =   8430
      Left            =   135
      TabIndex        =   75
      Top             =   1545
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   14870
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16761024
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox txtFolder 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   135
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1170
      Width           =   2880
   End
   Begin MSComctlLib.Slider scrBlend 
      Height          =   210
      Left            =   9090
      TabIndex        =   80
      ToolTipText     =   " Blend with Background (Normal Logo Only) "
      Top             =   9495
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   370
      _Version        =   393216
      Max             =   255
      TickStyle       =   3
      TickFrequency   =   15
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Blend Background"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9180
      TabIndex        =   81
      Top             =   9675
      Width           =   1185
   End
   Begin VB.Label lblImageControl 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Image Control"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3555
      TabIndex        =   76
      Top             =   7395
      Width           =   1290
   End
   Begin VB.Label lblLogoIndicator 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Logo Indicator"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3780
      TabIndex        =   70
      Top             =   9675
      Width           =   1485
   End
   Begin VB.Label lblLogoControl 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Logo Control"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9090
      TabIndex        =   63
      Top             =   7395
      Width           =   1290
   End
   Begin VB.Image menu_down 
      Height          =   555
      Index           =   6
      Left            =   13920
      Picture         =   "frmMain.frx":88B67
      Top             =   11385
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_norm 
      Height          =   555
      Index           =   6
      Left            =   13920
      Picture         =   "frmMain.frx":8A769
      Top             =   10770
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_down 
      Height          =   555
      Index           =   5
      Left            =   12900
      Picture         =   "frmMain.frx":8C36B
      Top             =   11385
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_down 
      Height          =   555
      Index           =   4
      Left            =   11880
      Picture         =   "frmMain.frx":8DF6D
      Top             =   11385
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_down 
      Height          =   555
      Index           =   3
      Left            =   10875
      Picture         =   "frmMain.frx":8FB6F
      Top             =   11385
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_down 
      Height          =   555
      Index           =   2
      Left            =   9855
      Picture         =   "frmMain.frx":91771
      Top             =   11385
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_down 
      Height          =   555
      Index           =   1
      Left            =   8835
      Picture         =   "frmMain.frx":93373
      Top             =   11385
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_down 
      Height          =   555
      Index           =   0
      Left            =   7815
      Picture         =   "frmMain.frx":94F75
      Top             =   11385
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_norm 
      Height          =   555
      Index           =   5
      Left            =   12900
      Picture         =   "frmMain.frx":96B77
      Top             =   10770
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_norm 
      Height          =   555
      Index           =   4
      Left            =   11880
      Picture         =   "frmMain.frx":98779
      Top             =   10770
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_norm 
      Height          =   555
      Index           =   3
      Left            =   10875
      Picture         =   "frmMain.frx":9A37B
      Top             =   10770
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_norm 
      Height          =   555
      Index           =   2
      Left            =   9855
      Picture         =   "frmMain.frx":9BF7D
      Top             =   10770
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_norm 
      Height          =   555
      Index           =   1
      Left            =   8835
      Picture         =   "frmMain.frx":9DB7F
      Top             =   10770
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image menu_norm 
      Height          =   555
      Index           =   0
      Left            =   7815
      Picture         =   "frmMain.frx":9F781
      Top             =   10770
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblPrintText 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "lblText"
      Height          =   270
      Left            =   1275
      TabIndex        =   62
      Top             =   11295
      Width           =   600
   End
   Begin VB.Image navLogo 
      Height          =   420
      Index           =   0
      Left            =   9525
      MouseIcon       =   "frmMain.frx":A1383
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A14D5
      Top             =   7860
      Width           =   420
   End
   Begin VB.Image navLogo 
      Height          =   420
      Index           =   1
      Left            =   9525
      MouseIcon       =   "frmMain.frx":A1E47
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A1F99
      Top             =   8730
      Width           =   420
   End
   Begin VB.Image navLogo 
      Height          =   420
      Index           =   2
      Left            =   9090
      MouseIcon       =   "frmMain.frx":A290B
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A2A5D
      Top             =   8295
      Width           =   420
   End
   Begin VB.Image navLogo 
      Height          =   420
      Index           =   3
      Left            =   9960
      MouseIcon       =   "frmMain.frx":A33CF
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A3521
      Top             =   8295
      Width           =   420
   End
   Begin VB.Image navLogo 
      Height          =   420
      Index           =   4
      Left            =   9525
      MouseIcon       =   "frmMain.frx":A3E93
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A3FE5
      Top             =   8295
      Width           =   420
   End
   Begin VB.Image navLogo 
      Height          =   420
      Index           =   5
      Left            =   9090
      MouseIcon       =   "frmMain.frx":A4957
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A4AA9
      Top             =   7860
      Width           =   420
   End
   Begin VB.Image navLogo 
      Height          =   420
      Index           =   6
      Left            =   9960
      MouseIcon       =   "frmMain.frx":A541B
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A556D
      Top             =   7860
      Width           =   420
   End
   Begin VB.Image navLogo 
      Height          =   420
      Index           =   7
      Left            =   9090
      MouseIcon       =   "frmMain.frx":A5EDF
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A6031
      Top             =   8730
      Width           =   420
   End
   Begin VB.Image navLogo 
      Height          =   420
      Index           =   8
      Left            =   9960
      MouseIcon       =   "frmMain.frx":A69A3
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":A6AF5
      Top             =   8730
      Width           =   420
   End
   Begin VB.Image logo_down 
      Height          =   420
      Index           =   7
      Left            =   7320
      Picture         =   "frmMain.frx":A7467
      Top             =   12180
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_down 
      Height          =   420
      Index           =   6
      Left            =   6870
      Picture         =   "frmMain.frx":A78DA
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_down 
      Height          =   420
      Index           =   5
      Left            =   6435
      Picture         =   "frmMain.frx":A7D14
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_down 
      Height          =   420
      Index           =   4
      Left            =   6000
      Picture         =   "frmMain.frx":A813F
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_down 
      Height          =   420
      Index           =   3
      Left            =   5550
      Picture         =   "frmMain.frx":A85C7
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_down 
      Height          =   420
      Index           =   2
      Left            =   5100
      Picture         =   "frmMain.frx":A8A13
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_down 
      Height          =   420
      Index           =   1
      Left            =   4650
      Picture         =   "frmMain.frx":A8E81
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_down 
      Height          =   420
      Index           =   0
      Left            =   4200
      Picture         =   "frmMain.frx":A92D7
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_norm 
      Height          =   420
      Index           =   7
      Left            =   7320
      Picture         =   "frmMain.frx":A9732
      Top             =   11730
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_norm 
      Height          =   420
      Index           =   6
      Left            =   6870
      Picture         =   "frmMain.frx":A9BA7
      Top             =   11730
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_norm 
      Height          =   420
      Index           =   5
      Left            =   6435
      Picture         =   "frmMain.frx":A9FEB
      Top             =   11730
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_norm 
      Height          =   420
      Index           =   4
      Left            =   6000
      Picture         =   "frmMain.frx":AA420
      Top             =   11730
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_norm 
      Height          =   420
      Index           =   3
      Left            =   5550
      Picture         =   "frmMain.frx":AA8B1
      Top             =   11730
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_norm 
      Height          =   420
      Index           =   2
      Left            =   5100
      Picture         =   "frmMain.frx":AACFB
      Top             =   11730
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_norm 
      Height          =   420
      Index           =   1
      Left            =   4650
      Picture         =   "frmMain.frx":AB176
      Top             =   11730
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image logo_norm 
      Height          =   420
      Index           =   0
      Left            =   4200
      Picture         =   "frmMain.frx":AB5D6
      Top             =   11730
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_down 
      Height          =   420
      Index           =   7
      Left            =   7305
      Picture         =   "frmMain.frx":ABA3A
      Top             =   10770
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_down 
      Height          =   420
      Index           =   6
      Left            =   6855
      Picture         =   "frmMain.frx":ABE8C
      Top             =   10770
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_down 
      Height          =   420
      Index           =   5
      Left            =   6420
      Picture         =   "frmMain.frx":AC29D
      Top             =   10770
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_down 
      Height          =   420
      Index           =   4
      Left            =   5970
      Picture         =   "frmMain.frx":AC698
      Top             =   10770
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_down 
      Height          =   420
      Index           =   3
      Left            =   5520
      Picture         =   "frmMain.frx":ACA88
      Top             =   10770
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_down 
      Height          =   420
      Index           =   2
      Left            =   5070
      Picture         =   "frmMain.frx":ACE7B
      Top             =   10770
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_down 
      Height          =   420
      Index           =   1
      Left            =   4620
      Picture         =   "frmMain.frx":AD272
      Top             =   10770
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_down 
      Height          =   420
      Index           =   0
      Left            =   4185
      Picture         =   "frmMain.frx":AD685
      Top             =   10770
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_norm 
      Height          =   420
      Index           =   7
      Left            =   7305
      Picture         =   "frmMain.frx":ADA8C
      Top             =   11235
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_norm 
      Height          =   420
      Index           =   6
      Left            =   6870
      Picture         =   "frmMain.frx":ADEE9
      Top             =   11235
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_norm 
      Height          =   420
      Index           =   5
      Left            =   6420
      Picture         =   "frmMain.frx":AE303
      Top             =   11235
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_norm 
      Height          =   420
      Index           =   4
      Left            =   5970
      Picture         =   "frmMain.frx":AE70E
      Top             =   11235
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_norm 
      Height          =   420
      Index           =   3
      Left            =   5520
      Picture         =   "frmMain.frx":AEB0A
      Top             =   11235
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_norm 
      Height          =   420
      Index           =   2
      Left            =   5070
      Picture         =   "frmMain.frx":AEF0E
      Top             =   11235
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_norm 
      Height          =   420
      Index           =   1
      Left            =   4620
      Picture         =   "frmMain.frx":AF30C
      Top             =   11235
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image font_norm 
      Height          =   420
      Index           =   0
      Left            =   4185
      Picture         =   "frmMain.frx":AF726
      Top             =   11235
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavNorm 
      Height          =   420
      Index           =   8
      Left            =   3630
      Picture         =   "frmMain.frx":AFB39
      Top             =   11715
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavNorm 
      Height          =   420
      Index           =   7
      Left            =   3180
      Picture         =   "frmMain.frx":B04AB
      Top             =   11715
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavNorm 
      Height          =   420
      Index           =   6
      Left            =   2730
      Picture         =   "frmMain.frx":B0E1D
      Top             =   11715
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavNorm 
      Height          =   420
      Index           =   5
      Left            =   2280
      Picture         =   "frmMain.frx":B178F
      Top             =   11715
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavNorm 
      Height          =   420
      Index           =   4
      Left            =   1830
      Picture         =   "frmMain.frx":B2101
      Top             =   11715
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavNorm 
      Height          =   420
      Index           =   3
      Left            =   1380
      Picture         =   "frmMain.frx":B2A73
      Top             =   11715
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavNorm 
      Height          =   420
      Index           =   2
      Left            =   930
      Picture         =   "frmMain.frx":B33E5
      Top             =   11715
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavNorm 
      Height          =   420
      Index           =   1
      Left            =   480
      Picture         =   "frmMain.frx":B3D57
      Top             =   11715
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavNorm 
      Height          =   420
      Index           =   0
      Left            =   30
      Picture         =   "frmMain.frx":B46C9
      Top             =   11715
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavDown 
      Height          =   420
      Index           =   8
      Left            =   3615
      Picture         =   "frmMain.frx":B503B
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavDown 
      Height          =   420
      Index           =   7
      Left            =   3165
      Picture         =   "frmMain.frx":B59AD
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavDown 
      Height          =   420
      Index           =   6
      Left            =   2730
      Picture         =   "frmMain.frx":B631F
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavDown 
      Height          =   420
      Index           =   5
      Left            =   2280
      Picture         =   "frmMain.frx":B6C91
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavDown 
      Height          =   420
      Index           =   4
      Left            =   1830
      Picture         =   "frmMain.frx":B7603
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavDown 
      Height          =   420
      Index           =   3
      Left            =   1380
      Picture         =   "frmMain.frx":B7F75
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavDown 
      Height          =   420
      Index           =   2
      Left            =   930
      Picture         =   "frmMain.frx":B88E7
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavDown 
      Height          =   420
      Index           =   1
      Left            =   480
      Picture         =   "frmMain.frx":B9259
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image imgNavDown 
      Height          =   420
      Index           =   0
      Left            =   30
      Picture         =   "frmMain.frx":B9BCB
      Top             =   12195
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Image navSource 
      Height          =   420
      Index           =   8
      Left            =   4425
      MouseIcon       =   "frmMain.frx":BA53D
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BA68F
      Top             =   8730
      Width           =   420
   End
   Begin VB.Image navSource 
      Height          =   420
      Index           =   7
      Left            =   3555
      MouseIcon       =   "frmMain.frx":BB001
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BB153
      Top             =   8730
      Width           =   420
   End
   Begin VB.Image navSource 
      Height          =   420
      Index           =   6
      Left            =   4425
      MouseIcon       =   "frmMain.frx":BBAC5
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BBC17
      Top             =   7860
      Width           =   420
   End
   Begin VB.Image navSource 
      Height          =   420
      Index           =   5
      Left            =   3555
      MouseIcon       =   "frmMain.frx":BC589
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BC6DB
      Top             =   7860
      Width           =   420
   End
   Begin VB.Image navSource 
      Height          =   420
      Index           =   4
      Left            =   3990
      MouseIcon       =   "frmMain.frx":BD04D
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BD19F
      Top             =   8295
      Width           =   420
   End
   Begin VB.Image navSource 
      Height          =   420
      Index           =   3
      Left            =   4425
      MouseIcon       =   "frmMain.frx":BDB11
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BDC63
      Top             =   8295
      Width           =   420
   End
   Begin VB.Image navSource 
      Height          =   420
      Index           =   2
      Left            =   3555
      MouseIcon       =   "frmMain.frx":BE5D5
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BE727
      Top             =   8295
      Width           =   420
   End
   Begin VB.Image navSource 
      Height          =   420
      Index           =   1
      Left            =   3990
      MouseIcon       =   "frmMain.frx":BF099
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BF1EB
      Top             =   8730
      Width           =   420
   End
   Begin VB.Image navSource 
      Height          =   420
      Index           =   0
      Left            =   3990
      MouseIcon       =   "frmMain.frx":BFB5D
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":BFCAF
      Top             =   7860
      Width           =   420
   End
   Begin VB.Image curDropper 
      Height          =   480
      Left            =   2505
      Picture         =   "frmMain.frx":C0621
      Top             =   10890
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image curGrab 
      Height          =   480
      Left            =   1560
      Picture         =   "frmMain.frx":C0EEB
      Top             =   10785
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image curRelease 
      Height          =   480
      Left            =   1170
      Picture         =   "frmMain.frx":C11F5
      Top             =   10785
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image curHand 
      Height          =   480
      Left            =   2085
      Picture         =   "frmMain.frx":C14FF
      Top             =   10860
      Width           =   480
   End
   Begin VB.Image imgSplit2 
      Height          =   9600
      Left            =   10980
      Picture         =   "frmMain.frx":C1651
      Stretch         =   -1  'True
      Top             =   465
      Width           =   75
   End
   Begin VB.Image imgSplit1 
      Height          =   9600
      Left            =   3075
      Picture         =   "frmMain.frx":C18A3
      Stretch         =   -1  'True
      Top             =   450
      Width           =   75
   End
   Begin VB.Image close_down 
      Height          =   285
      Left            =   765
      Picture         =   "frmMain.frx":C1AF5
      Top             =   11250
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_hot 
      Height          =   285
      Left            =   420
      Picture         =   "frmMain.frx":C1FAB
      Top             =   11250
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_norm 
      Height          =   285
      Left            =   60
      Picture         =   "frmMain.frx":C2461
      Top             =   11250
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_down 
      Height          =   285
      Left            =   765
      Picture         =   "frmMain.frx":C2917
      Top             =   10875
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_hot 
      Height          =   285
      Left            =   420
      Picture         =   "frmMain.frx":C2DCD
      Top             =   10875
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_norm 
      Height          =   285
      Left            =   60
      Picture         =   "frmMain.frx":C3283
      Top             =   10875
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgRight 
      Height          =   9660
      Left            =   14970
      Picture         =   "frmMain.frx":C3739
      Stretch         =   -1  'True
      Top             =   465
      Width           =   75
   End
   Begin VB.Image imgLeft 
      Height          =   9630
      Left            =   0
      Picture         =   "frmMain.frx":C398B
      Stretch         =   -1  'True
      Top             =   450
      Width           =   75
   End
   Begin VB.Image imgBack 
      Height          =   5835
      Left            =   3150
      Picture         =   "frmMain.frx":C3BDD
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   7830
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'                       (C) Photo Logo Plus - Author Berthold Bouman                            '
'                                        December 2008                                          '
'                                     All Rights reserved                                       '
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private strNormalLogo           As String       'normal logo file path
Private strMaskedLogo           As String       'masked logo file path
Private strPhotoName            As String       'file and path of image to be saved

Private InputX                  As Long         'drag source image
Private InputY                  As Long         'drag source image
Private MaxMoveX                As Long         'drag source image
Private MaxMoveY                As Long         'drag source image

Private blnDragSource           As Boolean      'flags we drag source image
Private blnScrollSource         As Boolean      'flags we scroll source image
Private blnScrollLogo           As Boolean      'flags we scroll logo image

Private flipNormalCounter       As Integer      'flip normal image counter
Private flipMaskedCounter       As Integer      'flip masked image counter
Private scrollSpeed             As Integer      'sets scrollspeed for text and logo

Private blnFontBold             As Boolean      'flags bold text
Private blnFontItalic           As Boolean      'flags italic text
Private blnFontUnderline        As Boolean      'flags underline text
Private blnFontStrikeThru       As Boolean      'flags striked text

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'       Note:   Some of the control events are declared Public and                              '
'               not Private, so we can call them from Batch Process                             '
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
Private Sub Form_Load()
    
    On Error GoTo ErrHandler
    
    If App.PrevInstance Then End
    
    Dim mTop        As Single       'form top position
    Dim mLeft       As Single       'form left position
    Dim msg         As String       'message box and error logging
    
    'error logfile
    strErrLog = App.Path & "\photologo.log"
    
    'initialize and check the size of the logfile (or create one if there isn't one)
    Call chkErrorLog(strErrLog, ByteSize)
    
    'create report directory
    If DirExists(App.Path & "\Reports") = False Then
        MkDir App.Path & "\Reports"
    End If
    
    'retrieve top and left coordinates from register
    mTop = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", "mainTop")
    mLeft = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", "mainLeft")
    
    'if not found or coordinates are beyond screen boundery
    '(if not found, the query returns 0 and is therefore smaller then 300)
    If mTop <= 300 Or mTop > Screen.Height - Me.Height _
                   Or mLeft <= 300 Or mLeft > Screen.Width - Me.Width Then
        'center form
        mTop = (Screen.Height - Me.Height) / 2
        mLeft = (Screen.Width - Me.Width) / 2
    End If
        
    'set position
    Me.Top = mTop
    Me.Left = mLeft
    
    'set font flags
    blnFontBold = False
    blnFontItalic = False
    blnFontUnderline = False
    blnFontStrikeThru = False
    
    'set other flags
    blnLeftTop = True
    blnBatch = False
    
    'set counters
    flipNormalCounter = 0
    flipMaskedCounter = 0
    
    Call setControls
    Call onImageLoad
    Call setLogoScrollBars
    Call makeThumb
    Call setLogoText
    
    'default logo is text logo
    Call optLogo_Click(0)
    
    'start with image in center position
    Call navSource_MouseDown(4, 0, 0, 0, 0)
    Call navSource_MouseUp(4, 0, 0, 0, 0)
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main Form_Load - Error " & Err.Number & ": " & Err.Description
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo ErrHandler
    
    Dim msg             As String       'message box
    
    'write to register - first we create a key - in case there isn't one,
    'e.g. when the app runs for the first time or when register is messed up
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Main"
    
    'now we can write the values to the register
    'save top and left position to register
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", _
                                   "mainTop", Me.Top, REG_SZ
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", _
                                   "mainLeft", Me.Left, REG_SZ
    
    'we have to properly unload the forms otherwise
    'the screencoordinates are not saved
    Unload frmFull
    Unload frmLogo
    Unload frmBatch
    Unload frmConvert
    Unload frmHelp
    Unload frmReport
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main Sub Form_Unload - Error " & Err.Number & ": " & _
              Err.Description
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

'++++++++++++++++++++++++++++++++++++++++++ CONTROL EVENTS ++++++++++++++++++++++++++++++++++++++

Private Sub imgMenu_Click(Index As Integer)
    
    'menubar
    On Error GoTo ErrHandler
    
    Dim strResFolder    As String       'open folder
    Dim strExt          As String       'file extention
    Dim msg             As String       'message box
    Dim submsg          As String       'message box
    
    Select Case Index
    
        Case 0          'open folder
            
            strResFolder = BrowseForFolder(hwnd, "Photo Logo Plus - Select a folder")
            'if cancel was selected
            If strResFolder = "" Then Exit Sub
            File1.Path = strResFolder
            txtFolder.Text = strResFolder
            'put in listview
            Call getFiles
            'save last used folder path to register
            CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Main"
            SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", _
                                           "Last Used Folder", txtFolder.Text, REG_SZ
            submsg = strResFolder
        
        Case 1          'paste image from clipboard
            
            If Clipboard.GetFormat(vbCFBitmap) Then
                picSource.Picture = Clipboard.GetData()
                Call onImageLoad
                Call setNavScrollBars
                Call makeThumb
                lblFileName.Caption = "File: Pasted from Clipboard " & "  -  Size: " & _
                          picSource.Width & " x " & picSource.Height
                strPhotoName = "Pasted from Clipboard.bmp"
                submsg = "Pasted from Clipboard"
            End If
        
        Case 2          'save image + logo
            
            'create filename string - get file extention first, we need
            'to know the length of the file extention to get rid of it
            strExt = GetFileExtention(strPhotoName)
            strPhotoName = Mid(strPhotoName, 1, Len(strPhotoName) - Len(strExt) - 1)
            submsg = strPhotoName
            
            Call savePhotoAs
            
            'increment serial (we do this after(!) the image has been saved)
            If chkAutoIncrement.Value = 1 And blnTextLogo = True Then
                valSerial = Val(txtAdd(4).Text) + 1
                If optFormat(0) = True Then strDigits = "000"
                If optFormat(1) = True Then strDigits = "0000"
                If optFormat(2) = True Then strDigits = "00000"
                txtAdd(4).Text = Format(valSerial, strDigits)
            End If
            
        Case 3          'show full size image + logo
        
            Unload frmFull
            frmFull.Show
        
        Case 4          'batch process
        
            frmBatch.Show
            blnBatch = True
            
        Case 5          'batch convert
        
            frmConvert.Show
        
        Case 6          'open help
        
            frmHelp.Show
        
    End Select

ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main imgMenu_Click - Error " & Err.Number & ": " & _
              Err.Description & " - " & submsg
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
        
End Sub

Private Sub imgMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'menu bar
    imgMenu(Index).Picture = menu_down(Index).Picture
    
End Sub

Private Sub imgMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'menu bar
    imgMenu(Index).Picture = menu_norm(Index).Picture
    
End Sub

Private Sub lsvFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    On Error GoTo ErrHandler
    
    Dim msg     As String   'message box
    Dim token   As Long     'GDI+
    
    'load new source image
    token = InitGDIPlus
    picSource.Picture = LoadPictureGDIPlus(fixPath(File1.Path, Item.Text))
    
    'if picture could not be loaded
    If Err.Number = 999 Then
        msg = Now & " Main Sub lsvFiles_ItemClick GDI+ failure 999 - " & _
        "File: " & fixPath(File1.Path, Item.Text)
        Call writeErrorLog(strErrLog, msg)
        msg = "Error loading picture, not a valid bitmap." & Chr(13) & _
        "File: " & fixPath(File1.Path, Item.Text) & "      "
        MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
    End If
     
    'free GDI+
    FreeGDIPlus token
    
    lblFileName.Caption = "File: " & Item.Text & "  -  Size: " & _
                          picSource.Width & " x " & picSource.Height
    strPhotoName = Item.Text
    strThumbName = fixPath(File1.Path, Item.Text)
    
    'implement changes
    Call onImageLoad
    Call setNavScrollBars
    Call makeThumb
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main Sub lsvFiles_ItemClick - Error " & Err.Number & ": " & _
              Err.Description & " - Path: " & fixPath(File1.Path, Item.Text)
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If

End Sub

Private Sub optLogo_Click(Index As Integer)
    
    'set logo type
    
    Dim WindowRegion As Long        'transparancy
    
    'reset flags
    blnTextLogo = False
    blnNormalLogo = False
    blnMaskedLogo = False
            
    Select Case Index
        
        Case 0          'text logo
            optLogo(0).Value = 1
            optLogo(1).Value = 0
            optLogo(2).Value = 0
            'images
            picSideBar1.Picture = LoadPicture(App.Path & "\Skin\gui_sidebar_hot.bmp")
            picSideBar2.Picture = LoadPicture(App.Path & "\Skin\gui_sidebar.bmp")
            picSideBar3.Picture = LoadPicture(App.Path & "\Skin\gui_sidebar.bmp")
            'set flag
            blnTextLogo = True
            Call enableTextControls
            lblLogoType.Caption = "Plain Text Logo"
            picNormalLogo.Visible = False
            picMaskedLogo.Visible = False
            Call setNavScrollBars
            If blnBatch = True Then frmBatch.optLogo(0).Value = 1
        
        Case 1          'normal logo image
            optLogo(0).Value = 0
            optLogo(1).Value = 1
            optLogo(2).Value = 0
            'images
            picSideBar1.Picture = LoadPicture(App.Path & "\Skin\gui_sidebar.bmp")
            picSideBar2.Picture = LoadPicture(App.Path & "\Skin\gui_sidebar_hot.bmp")
            picSideBar3.Picture = LoadPicture(App.Path & "\Skin\gui_sidebar.bmp")
            'set flag
            blnNormalLogo = True
            Call disableTextControls
            picSource.Cls
            lblLogoType.Caption = "Normal Image Logo"
            picNormalLogo.Visible = True
            picMaskedLogo.Visible = False
            picNormalLogo.Move 0, 0
            Call setNavScrollBars
            If blnBatch = True Then frmBatch.optLogo(1).Value = 1
            'do our alphablending thing
            scrBlend.Value = 0
            
        Case 2          'masked logo image
            optLogo(0).Value = 0
            optLogo(1).Value = 0
            optLogo(2).Value = 1
            'images
            picSideBar1.Picture = LoadPicture(App.Path & "\Skin\gui_sidebar.bmp")
            picSideBar2.Picture = LoadPicture(App.Path & "\Skin\gui_sidebar.bmp")
            picSideBar3.Picture = LoadPicture(App.Path & "\Skin\gui_sidebar_hot.bmp")
            'set flag
            blnMaskedLogo = True
            Call disableTextControls
            picSource.Cls
            lblLogoType.Caption = "Masked Image Logo"
            picNormalLogo.Visible = False
            picMaskedLogo.Visible = True
            picMaskedLogo.Move 0, 0
            'make transparant
            WindowRegion = MakeRegionTransparent(picMaskedLogo)
            SetWindowRgn picMaskedLogo.hwnd, WindowRegion, True
            Call setNavScrollBars
            If blnBatch = True Then frmBatch.optLogo(2).Value = 1
            
    End Select
    
    'we start at the top left corner - we have to do this to
    'implement the new scrollbar values again, otherwise the
    'logo ends up in the wrong place of the source image
    Call navLogo_MouseDown(5, 0, 0, 0, 0)
    Call navLogo_MouseUp(5, 0, 0, 0, 0)
    
End Sub

Public Sub navSource_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'navigate source image
    
    On Error GoTo ErrHandler
    
    blnScrollSource = True
    
    Select Case Index
    
        Case 0  'scroll source image up
            
            Do
                If scrSourceVer.Value < 0 Then Exit Sub
                scrSourceVer.Value = scrSourceVer.Value - scrollSpeed
                navSource(Index).Picture = imgNavDown(Index).Picture
                DoEvents
            Loop Until blnScrollSource = False
        
        Case 1  'scroll source image down
            
            Do
                If scrSourceVer.Value > scrSourceVer.Max Then Exit Sub
                scrSourceVer.Value = scrSourceVer.Value + scrollSpeed
                navSource(Index).Picture = imgNavDown(Index).Picture
                DoEvents
            Loop Until blnScrollSource = False
        
        Case 2  'scroll source image left
            
            Do
                If scrSourceHor.Value < 0 Then Exit Sub
                scrSourceHor.Value = scrSourceHor.Value - scrollSpeed
                navSource(Index).Picture = imgNavDown(Index).Picture
                DoEvents
            Loop Until blnScrollSource = False
            
        Case 3  'scroll source image right
            
            Do
                If scrSourceHor.Value > scrSourceHor.Max Then Exit Sub
                scrSourceHor.Value = scrSourceHor.Value + scrollSpeed
                navSource(Index).Picture = imgNavDown(Index).Picture
                DoEvents
            Loop Until blnScrollSource = False
        
        Case 4  'goto centre of source image
            
            scrSourceHor.Value = scrSourceHor.Max / 2
            scrSourceVer.Value = scrSourceVer.Max / 2
            navSource(Index).Picture = imgNavDown(Index).Picture
                                    
        Case 5  'goto top left of source image
            
            scrSourceHor.Value = 0
            scrSourceVer.Value = 0
            navSource(Index).Picture = imgNavDown(Index).Picture
                        
        Case 6  'goto top right of source image
            
            scrSourceHor.Value = scrSourceHor.Max
            scrSourceVer.Value = 0
            navSource(Index).Picture = imgNavDown(Index).Picture
                        
        Case 7  'goto bottom left of source image
            
            scrSourceHor.Value = 0
            scrSourceVer.Value = scrSourceVer.Max
            navSource(Index).Picture = imgNavDown(Index).Picture
                        
        Case 8  'goto bottom right of source image
            
            scrSourceHor.Value = scrSourceHor.Max
            scrSourceVer.Value = scrSourceVer.Max
            navSource(Index).Picture = imgNavDown(Index).Picture
                        
    End Select
            
ErrHandler:

    Exit Sub
            
End Sub

Public Sub navSource_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'navigate source image
    blnScrollSource = False
    navSource(Index).Picture = imgNavNorm(Index).Picture
    
End Sub

Public Sub navLogo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo ErrHandler
    
    'navigate logo image
    
    blnScrollLogo = True

    Select Case Index
        
        Case 0  'scroll logo up
            
            Do
                If scrLogoVer.Value < 0 Then Exit Sub
                scrLogoVer.Value = scrLogoVer.Value - scrollSpeed
                scrSourceVer.Value = (scrSourceVer.Max / scrLogoVer.Max) * scrLogoVer.Value
                navLogo(Index).Picture = imgNavDown(Index).Picture
                DoEvents
            Loop Until blnScrollLogo = False
        
        Case 1  'scroll logo down
            
            Do
                If scrLogoVer.Value > scrLogoVer.Max Then Exit Sub
                scrLogoVer.Value = scrLogoVer.Value + scrollSpeed
                scrSourceVer.Value = (scrSourceVer.Max / scrLogoVer.Max) * scrLogoVer.Value
                navLogo(Index).Picture = imgNavDown(Index).Picture
                DoEvents
            Loop Until blnScrollLogo = False
        
        Case 2  'scroll logo left
            
            Do
                If scrLogoHor.Value < 0 Then Exit Sub
                scrLogoHor.Value = scrLogoHor.Value - scrollSpeed
                scrSourceHor.Value = (scrSourceHor.Max / scrLogoHor.Max) * scrLogoHor.Value
                navLogo(Index).Picture = imgNavDown(Index).Picture
                DoEvents
            Loop Until blnScrollLogo = False
            
        Case 3  'scroll logo right
            
            Do
                If scrLogoHor.Value > scrLogoHor.Max Then Exit Sub
                scrLogoHor.Value = scrLogoHor.Value + scrollSpeed
                scrSourceHor.Value = (scrSourceHor.Max / scrLogoHor.Max) * scrLogoHor.Value
                navLogo(Index).Picture = imgNavDown(Index).Picture
                DoEvents
            Loop Until blnScrollLogo = False
            
        Case 4  'move logo to the centre
            
            scrSourceHor.Value = scrSourceHor.Max / 2
            scrSourceVer.Value = scrSourceVer.Max / 2
            scrLogoHor.Value = scrLogoHor.Max / 2
            scrLogoVer.Value = scrLogoVer.Max / 2
            navLogo(Index).Picture = imgNavDown(Index).Picture
            Call resetPostionFlags
            blnCenter = True
            If blnBatch = True Then Call frmBatch.setPosition(2)
            
        Case 5  'move logo to the left top
            
            scrLogoHor.Value = 0
            scrLogoVer.Value = 0
            scrSourceHor.Value = 0
            scrSourceVer.Value = 0
            navLogo(Index).Picture = imgNavDown(Index).Picture
            Call resetPostionFlags
            blnLeftTop = True
            If blnBatch = True Then Call frmBatch.setPosition(0)
            
        Case 6  'move logo to right top
            
            scrLogoHor.Value = scrLogoHor.Max
            scrLogoVer.Value = 0
            scrSourceHor.Value = scrSourceHor.Max
            scrSourceVer.Value = 0
            navLogo(Index).Picture = imgNavDown(Index).Picture
            Call resetPostionFlags
            blnRightTop = True
            If blnBatch = True Then Call frmBatch.setPosition(1)
                        
        Case 7  'move logo to the left bottom
            
            scrLogoHor.Value = 0
            scrLogoVer.Value = scrLogoVer.Max
            scrSourceHor.Value = 0
            scrSourceVer.Value = scrSourceVer.Max
            navLogo(Index).Picture = imgNavDown(Index).Picture
            Call resetPostionFlags
            blnLeftBot = True
            If blnBatch = True Then Call frmBatch.setPosition(3)
                                                           
        Case 8  'move logo to right bottom
            
            scrLogoHor.Value = scrLogoHor.Max
            scrLogoVer.Value = scrLogoVer.Max
            scrSourceHor.Value = scrSourceHor.Max
            scrSourceVer.Value = scrSourceVer.Max
            navLogo(Index).Picture = imgNavDown(Index).Picture
            Call resetPostionFlags
            blnRightBot = True
            If blnBatch = True Then Call frmBatch.setPosition(4)
                        
    End Select
    
ErrHandler:

    Exit Sub
    
End Sub

Public Sub navLogo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'navigate logo image
    blnScrollLogo = False
    navLogo(Index).Picture = imgNavNorm(Index).Picture
    
End Sub

Private Sub imgFontButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'plain text font buttons
    imgFontButton(Index).Picture = font_down(Index).Picture
    
End Sub

Private Sub imgFontButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'plain text font buttons
    imgFontButton(Index).Picture = font_norm(Index).Picture
    
End Sub

Private Sub imgFontButton_Click(Index As Integer)
    
    'plain text font buttons
    
    On Error GoTo ErrHandler
    
    Dim oldCol      As Long         'exchange text/shadow colors
    Dim msg         As String       'message box
    
    Call optLogo_Click(0)           'set the correct logo option
    
    Select Case Index
        
        Case 0      'change text color
            Dialog.CancelError = True
            Dialog.flags = cdlCCFullOpen + cdlCCRGBInit
            Dialog.Color = picTextColor.BackColor
            Dialog.ShowColor
            picTextColor.BackColor = Dialog.Color
        
        Case 1      'exchange text and shadow colors
            oldCol = picTextColor.BackColor
            picTextColor.BackColor = picShadowColor.BackColor
            picShadowColor.BackColor = oldCol
        
        Case 2      'change shadow color
            Dialog.CancelError = True
            Dialog.flags = cdlCCFullOpen + cdlCCRGBInit
            Dialog.Color = picShadowColor.BackColor
            Dialog.ShowColor
            picShadowColor.BackColor = Dialog.Color
        
        Case 3      'font bold
            If txtAdd(1).FontBold = True Then
                blnFontBold = False
                txtAdd(1).FontBold = False
            ElseIf txtAdd(1).FontBold = False Then
                blnFontBold = True
                txtAdd(1).FontBold = True
            End If
            
        Case 4      'font italic
            If txtAdd(1).FontItalic = True Then
                blnFontItalic = False
                txtAdd(1).FontItalic = False
            ElseIf txtAdd(1).FontItalic = False Then
                blnFontItalic = True
                txtAdd(1).FontItalic = True
            End If
        
        Case 5      'font underline
            If txtAdd(1).FontUnderline = True Then
                blnFontUnderline = False
                txtAdd(1).FontUnderline = False
            ElseIf txtAdd(1).FontUnderline = False Then
                blnFontUnderline = True
                txtAdd(1).FontUnderline = True
            End If
        
        Case 6      'font strike thru
            If txtAdd(1).FontStrikethru = True Then
                blnFontStrikeThru = False
                txtAdd(1).FontStrikethru = False
            ElseIf txtAdd(1).FontStrikethru = False Then
                blnFontStrikeThru = True
                txtAdd(1).FontStrikethru = True
            End If
        
        Case 7      'change font
            
            With Dialog
                'preset current font properties in dialog box
                .CancelError = True
                .flags = cdlCFScreenFonts + cdlCFEffects
                .FontName = txtFontName.Text
                .FontSize = Val(cboFontSize.Text)
                .FontBold = blnFontBold
                .FontItalic = blnFontItalic
                .FontUnderline = blnFontUnderline
                .FontStrikethru = blnFontStrikeThru
                .ShowFont
            End With

            'implement changes to user textbox
            txtFontName.Text = Dialog.FontName
            cboFontSize.Text = Dialog.FontSize
            
            txtAdd(1).Font = Dialog.FontName
            
            If Dialog.FontBold = True Then
                txtAdd(1).FontBold = True
                blnFontBold = True
            ElseIf Dialog.FontBold = False Then
                txtAdd(1).FontBold = False
                blnFontBold = False
            End If
            
            If Dialog.FontItalic = True Then
                txtAdd(1).FontItalic = True
                blnFontItalic = True
            ElseIf Dialog.FontItalic = False Then
                txtAdd(1).FontItalic = False
                blnFontItalic = False
            End If
            
            If Dialog.FontUnderline = True Then
                txtAdd(1).FontUnderline = True
                blnFontUnderline = True
            ElseIf Dialog.FontUnderline = False Then
                txtAdd(1).FontUnderline = False
                blnFontUnderline = False
            End If
            
            If Dialog.FontStrikethru = True Then
                txtAdd(1).FontStrikethru = True
                blnFontStrikeThru = True
            ElseIf Dialog.FontStrikethru = False Then
                txtAdd(1).FontStrikethru = False
                blnFontStrikeThru = False
            End If
            
    End Select
    
    'implement changes
    Call setLogoText
    Call setNavScrollBars
    
ErrHandler:
    If Err.Number = 32755 Then Exit Sub     'cancel error
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main Sub imgFontButton_Click " & Index & " - Error " & Err.Number & ": " & _
              Err.Description & " - Font: " & Dialog.FontName
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub imgNormLogo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'normal logo buttons
    imgNormLogo(Index).Picture = logo_down(Index).Picture
    
End Sub

Private Sub imgNormLogo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'normal logo buttons
    imgNormLogo(Index).Picture = logo_norm(Index).Picture
    
End Sub

Private Sub imgNormLogo_Click(Index As Integer)
    
    'normal logo buttons
    
    On Error GoTo ErrHandler
    
    Dim msg             As String       'message box
    Dim token           As Long         'GDI+
    
    Call optLogo_Click(1)               'set the correct option
    
    Select Case Index
    
        Case 0          'load normal logo
    
            Dialog.CancelError = True
            Dialog.DialogTitle = "Select a Logo"
            Dialog.Filter = "All Files (*.*)|*.*|" & _
                            "Windows Bitmap (*.bmp)|*.bmp|" & _
                            "JP(E)G - JFIF Compliant (*.jpg*.jif*.jpeg)|*.jpg;*.jif;*.jpeg|" & _
                            "CompuServe Graphic Interchange (*.gif)|*.gif|" & _
                            "Windows Meta File (*.wmf)|*.wmf|" & _
                            "Portable Networks Graphics (*.png)|*.png|" & _
                            "Tagged Image Format (*.tif.*tiff)|*.tif;*.tiff"
            Dialog.FilterIndex = 1
            Dialog.InitDir = App.Path & "\LogoNorm"
            Dialog.ShowOpen
            scrBlend.Value = 0
            
            'initialise GDI+
            token = InitGDIPlus
            strNormalLogo = Dialog.FileName
            picNormalSource.Picture = LoadPictureGDIPlus(Dialog.FileName)
            picNormalLogo.Picture = LoadPictureGDIPlus(Dialog.FileName)
            
            'if picture could not be loaded
            If Err.Number = 999 Then
                msg = Now & " Main Sub imgNormLogo_Click GDI+ failure 999 - " & _
                "File: " & Dialog.FileName
                Call writeErrorLog(strErrLog, msg)
                msg = "Error loading picture, not a valid bitmap." & Chr(13) & _
                "File: " & Dialog.FileName & "     "
                MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
            End If
            
            'free GDI+
            FreeGDIPlus token
            
            picNormalLogo.Move 0, 0
            
            Call setLogoScrollBars
            Call setNavScrollBars
        
        Case 1          'restore normal logo
            
            If strNormalLogo = "" Then Exit Sub
            'initialise GDI+
            token = InitGDIPlus
            picNormalSource.Picture = LoadPictureGDIPlus(strNormalLogo)
            picNormalLogo.Picture = LoadPictureGDIPlus(strNormalLogo)
            picNormalBlend.Picture = LoadPictureGDIPlus(strNormalLogo)
            'free GDI+
            FreeGDIPlus token
            picNormalLogo.Move 0, 0
            flipNormalCounter = 0
            scrBlend.Value = 0
        
        Case 2          'paste normal logo from clipboard
            
            If Clipboard.GetFormat(vbCFBitmap) Then
                picNormalSource.Picture = Clipboard.GetData()
                picNormalLogo.Picture = Clipboard.GetData()
                strNormalLogo = "Pasted from Clipboard.bmp"
            End If
        
        Case 3          'save normal logo
        
            'prepare dialog box
            Dialog.CancelError = True
            Dialog.flags = cdlOFNOverwritePrompt
            Dialog.Filter = "Windows Bitmap (*.bmp)|*.bmp|" & _
                            "JPEG - JFIF Compliant (*.jpg*.jif*.jpeg)|*.jpg|" & _
                            "CompuServe Graphic Interchange (*.gif)|*.gif|" & _
                            "Portable Networks Graphics (*.png)|*.png|" & _
                            "Tagged Image Format (*.tif)|*.tif"
            Dialog.InitDir = App.Path & "\LogoNorm"
            
            'we remove the file extention
            strNormalLogo = Mid(strNormalLogo, 1, Len(strNormalLogo) - Len(GetFileExtention(strNormalLogo)) - 1)
            Dialog.FileName = strNormalLogo
            
            Dialog.ShowSave
            
            'string with user selected file extention
            strNormalLogo = Dialog.FileName
            
            'initialise GDI+
            token = InitGDIPlus
    
            'now we can save the logo
            If SavePictureFromHDC(picNormalSource.Picture, strNormalLogo) = False Then
                msg = Now & " Main Sub imgNormLogo_Click - GDI+ Failure - " & _
                      "Picture could NOT be saved. " & _
                      "File: " & strNormalLogo
                Call writeErrorLog(strErrLog, msg)
                msg = "Picture could NOT be saved, try again or        " & _
                  Chr(13) & "select another file format. (GDI+ Failure)"
                MsgBox msg, vbOKOnly + vbCritical, "Photo Logo Plus"
            End If
    
            'free GDI+
            FreeGDIPlus token
                    
        Case 4          'view fullsize normal logo
            
            Unload frmLogo
            frmLogo.Show
        
        Case 5          'flip normal logo
            
            'increment counter
            flipNormalCounter = flipNormalCounter + 1
            
            If flipNormalCounter > 4 Then flipNormalCounter = 1
            
            If flipNormalCounter = 1 Then   'flip horizontal
                picNormalLogo.PaintPicture picNormalSource.Picture, -1, 0, _
                picNormalSource.Width, picNormalSource.Height, picNormalSource.Width, _
                0, -picNormalSource.Width, picNormalSource.Height, vbSrcCopy
            End If
            
            If flipNormalCounter = 2 Then   'flip vertical
                picNormalLogo.PaintPicture picNormalSource.Picture, 0, -1, _
                picNormalSource.Width, picNormalSource.Height, 0, picNormalSource.Height, _
                picNormalSource.Width, -picNormalSource.Height, vbSrcCopy
            End If
            
            If flipNormalCounter = 3 Then   'flip both horizontal and vertical
                picNormalLogo.PaintPicture picNormalSource.Picture, -1, -1, _
                picNormalSource.Width, picNormalSource.Height, picNormalSource.Width, _
                picNormalSource.Height, -picNormalSource.Width, -picNormalSource.Height, vbSrcCopy
            End If
            
            If flipNormalCounter = 4 Then   'restore original image
                picNormalLogo.Picture = picNormalSource.Picture
                flipNormalCounter = 0
            End If
            
        Case 6          'invert normal logo
            
            BitBlt picNormalLogo.hDC, _
                    0, _
                    0, _
                    picNormalSource.ScaleWidth, _
                    picNormalSource.ScaleHeight, _
                    picNormalSource.hDC, _
                    picNormalSource.Left, _
                    picNormalSource.Top, _
                    vbNotSrcCopy
           
            picNormalSource.Picture = picNormalLogo.image
            flipNormalCounter = 0
            
        Case 7          'greyscale normal logo
            
            Call GreyScale(picNormalLogo)
            picNormalSource.Picture = picNormalLogo.image
            
    End Select
    
    'do our alphablending thing
    picNormalBlend.Picture = picNormalLogo.image
    Call scrBlend_Change
    
    'implement changes
    Call normalLogo
    
ErrHandler:
    If Err.Number = 32755 Then Exit Sub     'cancel error
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main Sub imgNormLogo_Click " & Index & " - Error " & Err.Number & ": " & _
              Err.Description & " File: " & strNormalLogo
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub imgMaskedLogo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'masked logo buttons
    imgMaskedLogo(Index).Picture = logo_down(Index).Picture
    
End Sub

Private Sub imgMaskedLogo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'normal logo buttons
    imgMaskedLogo(Index).Picture = logo_norm(Index).Picture
    
End Sub

Private Sub imgMaskedLogo_Click(Index As Integer)
    
    'masked logo buttons
    
    On Error GoTo ErrHandler
    
    Dim WindowRegion    As Long         'sets transparancy
    Dim msg             As String       'message box
    Dim token           As Long         'GDI+
    
    Call optLogo_Click(2)               'set the correct logo option
     
    Select Case Index
    
        Case 0          'load masked logo
        
            Dialog.CancelError = True
            Dialog.DialogTitle = "Select a Logo"
            Dialog.Filter = "All Files (*.*)|*.*|" & _
                            "Windows Bitmap (*.bmp)|*.bmp|" & _
                            "JP(E)G - JFIF Compliant (*.jpg*.jif*.jpeg)|*.jpg;*.jif;*.jpeg|" & _
                            "CompuServe Graphic Interchange (*.gif)|*.gif|" & _
                            "Windows Meta File (*.wmf)|*.wmf|" & _
                            "Portable Networks Graphics (*.png)|*.png|" & _
                            "Tagged Image Format (*.tif.*tiff)|*.tif;*.tiff"
            Dialog.InitDir = App.Path & "\LogoTrans"
            Dialog.ShowOpen
            
            strMaskedLogo = Dialog.FileName
            
            'initialise GDI+
            token = InitGDIPlus
            strMaskedLogo = Dialog.FileName
            picMaskedSource.Picture = LoadPictureGDIPlus(Dialog.FileName)
            picMaskedLogo.Picture = LoadPictureGDIPlus(Dialog.FileName)
            
            'if picture could not be loaded
            If Err.Number = 999 Then
                msg = Now & " Main Sub imgMaskedLogo_Click GDI+ failure 999 - " & _
                "File: " & Dialog.FileName
                Call writeErrorLog(strErrLog, msg)
                msg = "Error loading picture, not a valid bitmap." & Chr(13) & _
                "File: " & Dialog.FileName & "     "
                MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
            End If
            
            'free GDI+
            FreeGDIPlus token
            
            picMaskedLogo.Move 0, 0
            
            Call setLogoScrollBars
            Call setNavScrollBars
            
            'make transparant
            WindowRegion = MakeRegionTransparent(picMaskedLogo)
            SetWindowRgn picMaskedLogo.hwnd, WindowRegion, True
        
        Case 1          'restore masked logo
            
            If strMaskedLogo = "" Then Exit Sub
            'initialise GDI+
            token = InitGDIPlus
            picMaskedSource.Picture = LoadPictureGDIPlus(strMaskedLogo)
            picMaskedLogo.Picture = LoadPictureGDIPlus(strMaskedLogo)
            'free GDI+
            FreeGDIPlus token
            'make transparant
            WindowRegion = MakeRegionTransparent(picMaskedLogo)
            SetWindowRgn picMaskedLogo.hwnd, WindowRegion, True
            picMaskedLogo.Move 0, 0
            flipMaskedCounter = 0
            
        Case 2          'paste masked logo from clipboard
        
            If Clipboard.GetFormat(vbCFBitmap) Then
                picMaskedSource.Picture = Clipboard.GetData()
                picMaskedLogo.Picture = Clipboard.GetData()
                'make transparant
                WindowRegion = MakeRegionTransparent(picMaskedLogo)
                SetWindowRgn picMaskedLogo.hwnd, WindowRegion, True
                strMaskedLogo = "Pasted from Clipboard.bmp"
            End If
        
        Case 3          'save masked logo
        
            'prepare dialog box
            Dialog.CancelError = True
            Dialog.flags = cdlOFNOverwritePrompt
            Dialog.Filter = "Windows Bitmap (*.bmp)|*.bmp|" & _
                            "JPEG - JFIF Compliant (*.jpg*.jif*.jpeg)|*.jpg|" & _
                            "CompuServe Graphic Interchange (*.gif)|*.gif|" & _
                            "Portable Networks Graphics (*.png)|*.png|" & _
                            "Tagged Image Format (*.tif)|*.tif"
            Dialog.InitDir = App.Path & "\LogoTrans"
            
            'we remove the file extention
            strMaskedLogo = Mid(strMaskedLogo, 1, Len(strMaskedLogo) - Len(GetFileExtention(strMaskedLogo)) - 1)
            Dialog.FileName = strMaskedLogo
            
            Dialog.ShowSave
            
            'string with user selected file extention
            strMaskedLogo = Dialog.FileName
            
            'initialise GDI+
            token = InitGDIPlus
    
            'now we can save the logo
            If SavePictureFromHDC(picMaskedSource.Picture, strMaskedLogo) = False Then
                msg = Now & " Main Sub imgMaskedLogo_Click - GDI+ Failure - " & _
                      "Picture could NOT be saved. " & _
                      "File: " & strMaskedLogo
                Call writeErrorLog(strErrLog, msg)
                 msg = "Picture could NOT be saved, try again or        " & _
                  Chr(13) & "select another file format. (GDI+ Failure)"
                MsgBox msg, vbOKOnly + vbCritical, "Photo Logo Plus"
            End If
    
            'free GDI+
            FreeGDIPlus token
        
        Case 4          'view fullsize masked logo
        
            Unload frmLogo
            frmLogo.Show
        
        Case 5          'flip masked logo
        
            'increment counter
            flipMaskedCounter = flipMaskedCounter + 1
            
            If flipMaskedCounter > 4 Then flipMaskedCounter = 1
            
            If flipMaskedCounter = 1 Then   'flip horizontal
                picMaskedLogo.PaintPicture picMaskedSource.Picture, -1, 0, _
                picMaskedSource.Width, picMaskedSource.Height, picMaskedSource.Width, _
                0, -picMaskedSource.Width, picMaskedSource.Height, vbSrcCopy
            End If
            
            If flipMaskedCounter = 2 Then   'flip vertical
                picMaskedLogo.PaintPicture picMaskedSource.Picture, 0, -1, _
                picMaskedSource.Width, picMaskedSource.Height, 0, picMaskedSource.Height, _
                picMaskedSource.Width, -picMaskedSource.Height, vbSrcCopy
            End If
            
            If flipMaskedCounter = 3 Then   'flip both horizontal and vertical
                picMaskedLogo.PaintPicture picMaskedSource.Picture, -1, -1, _
                picMaskedSource.Width, picMaskedSource.Height, picMaskedSource.Width, _
                picMaskedSource.Height, -picMaskedSource.Width, -picMaskedSource.Height, vbSrcCopy
            End If
            
            If flipMaskedCounter = 4 Then   'restore original image
                picMaskedLogo.Picture = picMaskedSource.Picture
                flipMaskedCounter = 0
            End If
            
            'make transparant
            WindowRegion = MakeRegionTransparent(picMaskedLogo)
            SetWindowRgn picMaskedLogo.hwnd, WindowRegion, True
                
        Case 6          'invert masked logo
            
            BitBlt picMaskedLogo.hDC, _
                    0, _
                    0, _
                    picMaskedSource.ScaleWidth, _
                    picMaskedSource.ScaleHeight, _
                    picMaskedSource.hDC, _
                    picMaskedSource.Left, _
                    picMaskedSource.Top, _
                    vbNotSrcCopy
           
            picMaskedSource.Picture = picMaskedLogo.image
            'make transparant
            WindowRegion = MakeRegionTransparent(picMaskedLogo)
            SetWindowRgn picMaskedLogo.hwnd, WindowRegion, True
            flipMaskedCounter = 0
            
    End Select
    
    'implement changes
    Call maskedLogo
    
ErrHandler:
    If Err.Number = 32755 Then Exit Sub     'cancel error
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main Sub imgMaskedLogo_Click " & Index & " - Error " & Err.Number & ": " & _
              Err.Description & " File: " & strMaskedLogo
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub cboFontSize_Change()
    
    'change fontsize
    txtAdd(1).FontSize = Val(cboFontSize.Text)
    
    'implement changes
    Call setLogoText
    Call setNavScrollBars
    
End Sub

Private Sub cboFontSize_Click()
    
    'change fontsize
    txtAdd(1).FontSize = Val(cboFontSize.Text)
    
    'implement changes
    Call setLogoText
    Call setNavScrollBars
    
End Sub

Private Sub picTextColor_Click()
    
    'change text color
    
    On Error GoTo ErrHandler
    
    Dialog.CancelError = True
    Dialog.flags = cdlCCFullOpen + cdlCCRGBInit
    Dialog.Color = picTextColor.BackColor
    Dialog.ShowColor
    picTextColor.BackColor = Dialog.Color
    
    'implement changes
    Call setLogoText
    Call setNavScrollBars
    
ErrHandler:

End Sub

Private Sub picShadowColor_Click()
    
    'change shadow color
    
    On Error GoTo ErrHandler
    
    Dialog.CancelError = True
    Dialog.flags = cdlCCFullOpen + cdlCCRGBInit
    Dialog.Color = picShadowColor.BackColor
    Dialog.ShowColor
    picShadowColor.BackColor = Dialog.Color
    
    'implement changes
    Call setLogoText
    Call setNavScrollBars
    
ErrHandler:

End Sub

Private Sub picSource_DblClick()
    
    'show full size image
    frmFull.Show
    'in case the form was already loaded, we refresh the image
    Call frmFull.showImage
    
End Sub

Private Sub picMaskedLogo_DblClick()
    
    'show full size image
    frmFull.Show
    'in case the form was already loaded, we refresh the image
    Call frmFull.showImage
    
End Sub

Private Sub picNormalLogo_DblClick()
    
    'show full size image
    frmFull.Show
    'in case the form was already loaded, we refresh the image
    Call frmFull.showImage
    
End Sub

Private Sub picNormalSource_DblClick()
    
    'show fullsize logo
    Call optLogo_Click(1)
    Unload frmLogo
    frmLogo.Show
    
End Sub

Private Sub picMaskedSource_DblClick()
    
    'show fullsize logo
    Call optLogo_Click(2)
    Unload frmLogo
    frmLogo.Show
    
End Sub

Private Sub picSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'drag source picture
    picSource.MouseIcon = curGrab.Picture
    
    InputX = x
    InputY = y
    
    blnDragSource = True
    
End Sub

Private Sub picSource_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'drag source picture
    
    Dim CurrX As Long
    Dim CurrY As Long
    
    If blnDragSource = False Then Exit Sub
    
    picSource.MouseIcon = curGrab.Picture
    
    CurrX = picSource.Left + (x - InputX)
    CurrY = picSource.Top + (y - InputY)

    If CurrX > 0 Then
        CurrX = 0
        InputX = x
    ElseIf CurrX < MaxMoveX Then
        CurrX = MaxMoveX
        InputX = x
    End If

    If CurrY > 0 Then
        CurrY = 0
        InputY = y
    ElseIf CurrY < MaxMoveY Then
        CurrY = MaxMoveY
        InputY = y
    End If

    picSource.Move CurrX, CurrY
    
    'update scrollbar values
    scrSourceHor.Value = -CurrX
    scrSourceVer.Value = -CurrY
   
    DoEvents
    
End Sub

Private Sub picSource_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'drag source picture
    picSource.MouseIcon = curRelease.Picture
    blnDragSource = False
    
End Sub

Private Sub chkLogoIndicator_Click()
    
    'show/don't show logo indicator
    If chkLogoIndicator.Value = 1 Then
        shpLogoPos.Visible = True
        tmrBlink.Enabled = True
    ElseIf chkLogoIndicator.Value = 0 Then
        shpLogoPos.Visible = False
        tmrBlink.Enabled = False
    End If
    
End Sub

Private Sub tmrBlink_Timer()
    
    'blink logo indicator
    If shpLogoPos.BorderStyle = 3 Then
        shpLogoPos.BorderStyle = 1
    ElseIf shpLogoPos.BorderStyle = 1 Then
        shpLogoPos.BorderStyle = 3
    End If
    
End Sub

'+++++++++++++++++++++++++++++++++++ NAVIGATION SCROLL BARS +++++++++++++++++++++++++++++++++++++

Private Sub scrSourceHor_Change()
    
    'scroll source picture horizontal
    picSource.Left = -scrSourceHor.Value
    thumbViewPort.Left = (scrSourceHor.Value * (picThumb.Width / picSource.Width))
    
End Sub

Private Sub scrSourceHor_Scroll()
    
    'scroll source picture horizontal
    Call scrSourceHor_Change
    
End Sub

Private Sub scrSourceVer_Change()
    
    'scroll source picture vertical
    picSource.Top = -scrSourceVer.Value
    thumbViewPort.Top = (scrSourceVer.Value * (picThumb.Height / picSource.Height))
    
End Sub

Private Sub scrSourceVer_Scroll()
    
    'scroll source picture vertical
    Call scrSourceVer_Change
    
End Sub

Private Sub scrLogoVer_Change()
    
    'scroll logo image vertical
    If blnTextLogo = True Then Call textLogo
    If blnNormalLogo = True Then
        Call normalLogo
        'do our alphablending thing
        Call scrBlend_Change
    End If
    If blnMaskedLogo = True Then Call maskedLogo
    
    scrSourceVer.Value = (scrSourceVer.Max / scrLogoVer.Max) * scrLogoVer.Value
    shpLogoPos.Top = (scrLogoVer.Value * (picThumb.Width / picSource.Width))
    
End Sub

Private Sub scrLogoVer_Scroll()
    
    'scroll logo image vertical
    Call scrLogoVer_Change
    
End Sub

Private Sub scrLogoHor_Change()
    
    'scroll logo image horizontal
    If blnTextLogo = True Then Call textLogo
    If blnNormalLogo = True Then
        Call normalLogo
        'do our alphablending thing
        Call scrBlend_Change
    End If
    If blnMaskedLogo = True Then Call maskedLogo
        
    scrSourceHor.Value = (scrSourceHor.Max / scrLogoHor.Max) * scrLogoHor.Value
    shpLogoPos.Left = (scrLogoHor.Value * (picThumb.Height / picSource.Height))
    
End Sub

Private Sub scrLogoHor_Scroll()
    
    'scroll logo image horizontal
    Call scrLogoHor_Change
    
End Sub

Public Sub scrBlend_Change()
    
    'alphablend normal logo
    
    Dim tProperties     As typeBlendProperties     'set type structure
    Dim lngBlend        As Long                    'CopyMemory
    
    If blnNormalLogo = True Then
        
        'clear the destination picture
        picNormalLogo.Cls
        
        'copy part of the image that is covered
        'by the logo to the logo background
        BitBlt picNormalLogo.hDC, _
            0, _
            0, _
            picNormalLogo.ScaleWidth, _
            picNormalLogo.ScaleHeight, _
            picSource.hDC, _
            picNormalLogo.Left, _
            picNormalLogo.Top, _
            vbSrcCopy
        
        'set blend value
        tProperties.tBlendAmount = 255 - scrBlend.Value
        'call the CopyMemory with the specified parameters
        CopyMemory lngBlend, tProperties, 4
    
        'blend the images
        AlphaBlend picNormalLogo.hDC, _
            0, _
            0, _
            picNormalSource.ScaleWidth, _
            picNormalSource.ScaleHeight, _
            picNormalBlend.hDC, _
            0, _
            0, _
            picNormalSource.ScaleWidth, _
            picNormalSource.ScaleHeight, _
            lngBlend
    
        'refresh the picture box with the new image
        'picNormalLogo.Refresh
        
    End If

End Sub

Private Sub scrBlend_Scroll()
    
    'set transparency
    Call scrBlend_Change
    
End Sub

'++++++++++++++++++++++++++++++++++++++ LOGO IMAGE SCROLL BARS ++++++++++++++++++++++++++++++++++

Private Sub normLogoHorScroll_Change()
    'normal logo horizontal scroll
    picNormalSource.Left = -normLogoHorScroll.Value

End Sub

Private Sub normLogoHorScroll_Scroll()
    'normal logo horizontal scroll
    picNormalSource.Left = -normLogoHorScroll.Value

End Sub

Private Sub normLogoVerScroll_Change()
    'normal logo vertical scroll
    picNormalSource.Top = -normLogoVerScroll.Value

End Sub

Private Sub normLogoVerScroll_Scroll()
    'normal logo vertical scroll
    picNormalSource.Top = -normLogoVerScroll.Value

End Sub

Private Sub maskLogoHorScroll_Change()
    'masked logo horizontal scroll
    picMaskedSource.Left = -maskLogoHorScroll.Value

End Sub

Private Sub maskLogoHorScroll_Scroll()
    'masked logo horizontal scroll
    picMaskedSource.Left = -maskLogoHorScroll.Value

End Sub

Private Sub maskLogoVerScroll_Change()
    'masked logo vertical scroll
    picMaskedSource.Top = -maskLogoVerScroll.Value

End Sub

Private Sub maskLogoVerScroll_Scroll()
    'masked logo vertical scroll
    picMaskedSource.Top = -maskLogoVerScroll.Value

End Sub

'+++++++++++++++++++++++++++++++++++++++++++++ TEXT LOGO ++++++++++++++++++++++++++++++++++++++++

Private Sub txtAdd_Change(Index As Integer)
        
    If blnTextLogo = False Then Exit Sub
    
    'refesh string and implement changes
    Call setLogoText
    Call setNavScrollBars
        
End Sub

Private Sub chkAddText_Click(Index As Integer)
    
    If chkAddText(3).Value = 0 Then chkCurTime.Value = 0
    
    'refesh string and implement changes
    Call setLogoText
    Call setNavScrollBars
        
End Sub

Private Sub chkCurDate_Click()
    
    'add current date
    If chkCurDate.Value = 1 Then
        txtAdd(2).Text = Format(Now, "DD-MM-YYYY")
    Else
        txtAdd(2).Text = "My Date"
    End If
    
    'implement changes
    Call setLogoText
    Call setNavScrollBars
    
End Sub

Private Sub chkCurTime_Click()
    
    'add current time
    If chkCurTime.Value = 1 Then
        tmrCurTime.Enabled = True
    Else
        tmrCurTime.Enabled = False
        txtAdd(3).Text = "My Time"
    End If
    
    'implement changes
    Call setLogoText
    Call setNavScrollBars
    
End Sub

Private Sub udcShadowOffset_Change()
    
    'shadow offset
    txtOffset.Text = udcShadowOffset.Value
    
    'implement changes
    Call setLogoText
    Call setNavScrollBars
    
End Sub

Private Sub chkShadow_Click()
    
    'do or don't add shadow
    
    'implement changes
    Call setLogoText
    Call setNavScrollBars
    
End Sub

Private Sub chkAutoIncrement_Click()
    
    'auto increment serial number
    If chkAutoIncrement.Value = 1 Then
        If optFormat(0) = True Then
            txtAdd(4).Text = "001"
            strDigits = "000"
        End If
        If optFormat(1) = True Then
            txtAdd(4).Text = "0001"
            strDigits = "0000"
        End If
        If optFormat(2) = True Then
            txtAdd(4).Text = "00001"
            strDigits = "00000"
        End If
    ElseIf chkAutoIncrement.Value = 0 Then
        txtAdd(4).Text = "My Serial"
    End If
    
End Sub

Private Sub optFormat_Click(Index As Integer)
    
    'serial number format
    If chkAutoIncrement.Value = 1 Then
        If Index = 0 Then
            txtAdd(4).Text = "001"
            strDigits = "000"
        End If
        If Index = 1 Then
            txtAdd(4).Text = "0001"
            strDigits = "0000"
        End If
        If Index = 2 Then
            txtAdd(4).Text = "00001"
            strDigits = "00000"
        End If
    End If
        
End Sub

Private Sub tmrCurTime_Timer()
    
    'display current time in textbox
    txtAdd(3).Text = Format(Now, "HH:MM:SS")
        
End Sub

'++++++++++++++++++++++++++++++++++++++++++++ COMMON SUBS +++++++++++++++++++++++++++++++++++++++

Private Sub setControls()
    
    On Error GoTo ErrHandler
    
    Dim n           As Integer      'counter
    Dim strFolder   As String       'folder path
    Dim msg         As String       'message box
    Dim clmX        As ColumnHeader 'listview column
    
    'colors
    picSide.BackColor = RGB(76, 131, 214)
    
    For n = 0 To Me.Height Step 3
        'paint background
        Me.Line (0, n)-(Me.Width, n), RGB(190, 212, 255)
        Me.Line (0, n + 1)-(Me.Width, n + 1), RGB(204, 224, 255)
        Me.Line (0, n + 2)-(Me.Width, n + 2), RGB(255, 255, 255)
    Next n
    
    'Me.BackColor = RGB(76, 131, 214)
    chkCurDate.BackColor = RGB(76, 131, 214)
    chkCurTime.BackColor = RGB(76, 131, 214)
    chkAutoIncrement.BackColor = RGB(76, 131, 214)
    chkShadow.BackColor = RGB(76, 131, 214)
    chkAddText(4).BackColor = RGB(76, 131, 214)
    
    For n = 0 To 4
        txtAdd(n).BackColor = RGB(203, 218, 237)
    Next n
    
    For n = 0 To 2
        optFormat(n).BackColor = RGB(76, 131, 214)
        optFormat(n).MousePointer = 99
        optFormat(n).MouseIcon = curHand
    Next n
    
    cboFontSize.BackColor = RGB(203, 218, 237)
    txtFontName.BackColor = RGB(203, 218, 237)
    portNormalLogo.BackColor = RGB(203, 218, 237)
    portMaskedLogo.BackColor = RGB(203, 218, 237)
    picFillNorm.BackColor = RGB(203, 218, 237)
    picFillMask.BackColor = RGB(203, 218, 237)
    
    For n = 0 To 2
        optLogo(n).BackColor = RGB(9, 41, 155)
        optLogo(n).MousePointer = 99
        optLogo(n).MouseIcon = curHand
    Next n
    
    For n = 0 To 6
        imgMenu(n).MousePointer = 99
        imgMenu(n).MouseIcon = curHand
        imgMenu(n).Picture = menu_norm(n).Picture
    Next n
    
    'cursors
    imgMin.MousePointer = 99
    imgMin.MouseIcon = curHand
    imgClose.MousePointer = 99
    imgClose.MouseIcon = curHand
    
    picShadowColor.MouseIcon = curDropper
    picTextColor.MouseIcon = curDropper
    
    For n = 0 To 7
        imgFontButton(n).MousePointer = 99
        imgFontButton(n).MouseIcon = curHand
        imgFontButton(n).Picture = font_norm(n).Picture
    Next n
    
    For n = 0 To 7
        imgNormLogo(n).MousePointer = 99
        imgNormLogo(n).MouseIcon = curHand
        imgNormLogo(n).Picture = logo_norm(n).Picture
    Next n
    
    For n = 0 To 6
        imgMaskedLogo(n).MousePointer = 99
        imgMaskedLogo(n).MouseIcon = curHand
        imgMaskedLogo(n).Picture = logo_norm(n).Picture
    Next n
    
    For n = 0 To 8
        navSource(n).Picture = imgNavNorm(n)
        navLogo(n).Picture = imgNavNorm(n)
    Next n
    
    For n = 0 To 4
        chkAddText(n).MousePointer = 99
        chkAddText(n).MouseIcon = curHand
    Next n
    
    chkShadow.MousePointer = 99
    chkShadow.MouseIcon = curHand
    
    chkAutoIncrement.MousePointer = 99
    chkAutoIncrement.MouseIcon = curHand
    
    chkCurDate.MousePointer = 99
    chkCurDate.MouseIcon = curHand
    
    chkCurTime.MousePointer = 99
    chkCurTime.MouseIcon = curHand
    
    chkLogoIndicator.MousePointer = 99
    chkLogoIndicator.MouseIcon = curHand
    
    scrBlend.MousePointer = 99
    scrBlend.MouseIcon = curHand
    
    'fill combo font size
    For n = 6 To 16
        cboFontSize.AddItem n
    Next n
    For n = 18 To 36 Step 2
        cboFontSize.AddItem n
    Next n
    
    'default = size 8
    cboFontSize.ListIndex = 2
    
    'file box
    File1.Pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.wmf;*.png;*.tif;*.tiff"
    
    'retrieve last used folder from register
    strFolder = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", "Last Used Folder")
    If strFolder = "" Then strFolder = App.Path & "\Samples"
    
    File1.Path = strFolder
    txtFolder.Text = File1.Path
    lblFileName.Caption = "File: Default" & "  -  Size: " & picSource.Width & " x " & picSource.Height
    
    'listview
    Set clmX = lsvFiles.ColumnHeaders.Add(1)
    clmX.Text = "File Name"
    clmX.Width = 400
    clmX.Alignment = lvwColumnLeft
    
    lsvFiles.HideColumnHeaders = True
    lsvFiles.Sorted = True
    
    'put files in listview
    Call getFiles
        
    'load default logo images
    picNormalSource.Picture = LoadPicture(App.Path & "\LogoNorm\Logo_03.bmp")
    picNormalLogo.Picture = LoadPicture(App.Path & "\LogoNorm\Logo_03.bmp")
    picNormalBlend = LoadPicture(App.Path & "\LogoNorm\Logo_03.bmp")
    strNormalLogo = App.Path & "\LogoNorm\Logo_03.bmp"
    
    picMaskedSource.Picture = LoadPicture(App.Path & "\LogoTrans\Logo_01.bmp")
    picMaskedLogo.Picture = LoadPicture(App.Path & "\LogoTrans\Logo_01.bmp")
    strMaskedLogo = App.Path & "\LogoTrans\Logo_01.bmp"
    
    'load default source image
    picSource.Picture = LoadPicture(App.Path & "\Samples\default.jpg")
    'default source and thumb image name
    strPhotoName = App.Path & "\Samples\default.jpg"
    strThumbName = App.Path & "\Samples\default.jpg"
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main Sub SetControls - Error " & Err.Number & ": " & Err.Description
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
        
End Sub

Private Sub getFiles()
    
    'get all files
    
    On Error GoTo ErrHandler
    
    Dim msg         As String           'message box
    Dim itmX        As ListItem         'sets our listitem
    Dim n           As Integer          'counter
    Dim strExt      As String           'extention string
    
    lsvFiles.ListItems.Clear
    
    'add all items to list
    For n = 1 To File1.ListCount
    
        File1.ListIndex = n - 1
        
        Set itmX = lsvFiles.ListItems.Add()
        itmX.Text = File1.FileName
        
        'get correct icon
        strExt = GetFileExtention(File1.FileName)
        
        Select Case strExt
            Case "bmp"
                itmX.SmallIcon = 1
            Case "gif"
                itmX.SmallIcon = 2
            Case "wmf"
                itmX.SmallIcon = 2
            Case "jpg"
                itmX.SmallIcon = 3
            Case "jpeg"
                itmX.SmallIcon = 3
            Case "png"
                itmX.SmallIcon = 4
            Case "tif"
                itmX.SmallIcon = 4
            Case "tiff"
                itmX.SmallIcon = 4
            Case Else
                itmX.SmallIcon = 5
        End Select
                                                        
        itmX.Selected = False
    Next n
    
    'update label
    lblSource.Caption = lsvFiles.ListItems.Count & " Images in this Folder"
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main sub getFiles - Error " & Err.Number & ": " & Err.Description & _
              " - " & fixPath(File1.Path, File1.FileName)
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Public Sub setLogoText()
    
    Dim strSymbol       As String
    Dim strText         As String
    Dim strDate         As String
    Dim strTime         As String
    Dim strSerial       As String
    Dim strSeperation   As String
    
    strSeperation = " - "
    
    'add copyright symbol
    If chkAddText(0) = 1 Then
        strSymbol = txtAdd(0).Text
           If chkAddText(1).Value = 1 Or _
           chkAddText(2).Value = 1 Or _
           chkAddText(3).Value = 1 Or _
           chkAddText(4).Value = 1 Then strSymbol = strSymbol & " "
    Else:
        strSymbol = ""
    End If
    
    'add user text
    If chkAddText(1).Value = 1 Then
        strText = txtAdd(1).Text
        If chkAddText(2).Value = 1 Or _
           chkAddText(3).Value = 1 Or _
           chkAddText(4).Value = 1 Then strText = strText & strSeperation
    Else:
        strText = ""
    End If
    
    'add date
    If chkAddText(2).Value = 1 Then
        strDate = txtAdd(2).Text
        If chkAddText(3).Value = 1 Or _
           chkAddText(4).Value = 1 Then strDate = strDate & strSeperation
    Else:
        strDate = ""
    End If
    
    'add time
    If chkAddText(3).Value = 1 Then
        strTime = txtAdd(3).Text
        If chkAddText(4).Value = 1 Then strTime = strTime & strSeperation
    Else:
        strTime = ""
    End If
    
    'add serial number
    If chkAddText(4).Value = 1 Then
        strSerial = txtAdd(4).Text
    Else:
        strSerial = ""
    End If
        
    'update labels
    lblLogoText.Caption = "Logo Text:  " & strSymbol & strText & strDate & strTime & strSerial
    lblPrintText.Caption = strSymbol & strText & strDate & strTime & strSerial
    
End Sub

Public Sub textLogo()
    
    'print text logo on source image
    picSource.Cls
    
    'set font properties
    picSource.Font = txtAdd(1).Font
    picSource.FontSize = Val(cboFontSize.Text)
    
    If txtAdd(1).FontBold = True Then
        picSource.FontBold = True
    ElseIf txtAdd(1).FontBold = False Then
        picSource.FontBold = False
    End If
    
    If txtAdd(1).FontItalic = True Then
        picSource.FontItalic = True
    ElseIf txtAdd(1).FontItalic = False Then
        picSource.FontItalic = False
    End If
    
    If txtAdd(1).FontUnderline = True Then
        picSource.FontUnderline = True
    ElseIf txtAdd(1).FontUnderline = False Then
        picSource.FontUnderline = False
    End If
    
    If txtAdd(1).FontStrikethru = True Then
        picSource.FontStrikethru = True
    ElseIf txtAdd(1).FontStrikethru = False Then
        picSource.FontStrikethru = False
    End If
                                
    'no shadow selected
    If chkShadow.Value = 0 Then GoTo Skip
    
    'print shadow first
    picSource.ForeColor = picShadowColor.BackColor
    picSource.CurrentX = scrLogoHor.Value + 2 + Val(txtOffset.Text)
    picSource.CurrentY = scrLogoVer.Value + Val(txtOffset.Text)
    picSource.Print lblPrintText.Caption
    
Skip:

    'now we can print the text
    picSource.ForeColor = picTextColor.BackColor
    picSource.CurrentX = scrLogoHor.Value + 2
    picSource.CurrentY = scrLogoVer.Value
    picSource.Print lblPrintText.Caption
    
    'update status bar
    lblXpos.Caption = " X position: " & scrLogoHor.Value + 3
    lblYpos.Caption = " Y position: " & scrLogoVer.Value + 1
    
End Sub

Public Sub normalLogo()
    
    'Note: we don't blit the logos to the source image here,
    'we just move the pictureboxes, logo is only blitted
    'to the source image when we save the image
    
    'normal logo
    picNormalLogo.Move scrLogoHor.Value, scrLogoVer.Value
    lblXpos.Caption = " X position: " & picNormalLogo.Left
    lblYpos.Caption = " Y position: " & picNormalLogo.Top
    
End Sub

Public Sub maskedLogo()
    
    'masked logo
    picMaskedLogo.Move scrLogoHor.Value, scrLogoVer.Value
    lblXpos.Caption = " X position: " & scrLogoHor.Value
    lblYpos.Caption = " Y position: " & scrLogoVer.Value
        
End Sub

Private Sub savePhotoAs()
    
    'save current image + logo
    On Error GoTo ErrHandler
    
    Dim msg             As String       'message box
    Dim token           As Long         'GDI+
    
    'prepare dialog box
    'retrieve last used folder from register
    DialogMain.InitDir = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", "Last Used Save")
    'string without file extention
    DialogMain.FileName = strPhotoName
    DialogMain.CancelError = True
    DialogMain.flags = cdlOFNOverwritePrompt
    DialogMain.Filter = "Windows Bitmap (*.bmp)|*.bmp|" & _
                    "JPEG - JFIF Compliant (*.jpg*.jif*.jpeg)|*.jpg|" & _
                    "CompuServe Graphic Interchange (*.gif)|*.gif|" & _
                    "Portable Networks Graphics (*.png)|*.png|" & _
                    "Tagged Image Format (*.tif)|*.tif"
    
    DialogMain.ShowSave
    
    'string with user selected file extention
    strPhotoName = DialogMain.FileName
    
    'save last used folder to register
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Main"
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", _
            "Last Used Save", Mid(DialogMain.FileName, 1, Len(DialogMain.FileName) - _
            Len(GetFileName(DialogMain.FileName)) - 1), REG_SZ
    
    'text logo
    If blnTextLogo = True Then
        'nothing to do here, text is already on source image
        'just to let you know :-)
    End If
    
    'normal image logo
    If blnNormalLogo = True Then
        'blit normal logo to picSource
        BitBlt picSource.hDC, _
            scrLogoHor.Value, _
            scrLogoVer.Value, _
            picNormalLogo.ScaleWidth, _
            picNormalLogo.ScaleHeight, _
            picNormalLogo.hDC, _
            0, _
            0, _
            vbSrcCopy
    End If
    
    'masked image logo
    If blnMaskedLogo = True Then
        'blit masked logo to picSource
        TransparentBlt picSource.hDC, _
            scrLogoHor.Value, _
            scrLogoVer.Value, _
            picMaskedLogo.ScaleWidth, _
            picMaskedLogo.ScaleHeight, _
            picMaskedLogo.hDC, _
            0, _
            0, _
            picMaskedLogo.ScaleWidth, _
            picMaskedLogo.ScaleHeight, _
            GetPixel(picMaskedSource.hDC, 0, 0)
    End If
    
    'initialise GDI+
    token = InitGDIPlus
        
    'now we can save the image + logo
    If SavePictureFromHDC(picSource.image, strPhotoName) = False Then
        msg = Now & " Sub savePhotoAs - GDI+ Failure - " & _
              "Picture could NOT be saved. " & _
              "File: " & strPhotoName
        Call writeErrorLog(strErrLog, msg)
        msg = "Picture could NOT be saved, try again or        " & _
                  "select another file format. (GDI+ Failure)"
        MsgBox msg, vbOKOnly + vbCritical, "Photo Logo Plus"
    End If
    
    'free GDI+
    FreeGDIPlus token
   
    'clear blitted logo from source image again
    If blnNormalLogo = True Or blnMaskedLogo = True Then picSource.Cls
   
ErrHandler:
    If Err.Number = 32755 Then Exit Sub     'cancel error
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main Sub savePhotoAs - Error " & Err.Number & ": " & _
              Err.Description & " File: " & strPhotoName
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Public Sub makeThumb()
    
    'make thumbnail
    On Error GoTo ErrHandler
    
    'Note: this way of making a thumbnail is only possible when the
    'max height and max width of the thumb are about the same, if they
    'are different we need to add some more calculations to compare
    'the height and width of the thumb with the source height and width,
    'now we only have to calculate the height or the width.
    
    Dim sideHor         As Integer
    Dim sideVer         As Integer
    Dim maxWidth        As Integer
    Dim maxHeight       As Integer
    Dim iRatio          As Double
    Dim token           As Long         'GDI+
    
    Dim msg             As String       'message box
    Dim submsg          As String       'message box
    
    'set max thumbnail size
    maxWidth = 140
    maxHeight = 140
    
    sideHor = picSource.Width
    sideVer = picSource.Height
    
    If sideHor > sideVer Then
        'horizontal side is longest
        picThumb.Width = maxWidth
        iRatio = picSource.Height / picSource.Width
        picThumb.Height = maxHeight * iRatio
    End If
    
    If sideHor < sideVer Then
        'vertical side is longest
        picThumb.Height = maxHeight
        iRatio = picSource.Width / picSource.Height
        picThumb.Width = maxWidth * iRatio
    End If
    
    If sideHor = sideVer Then
        'both same length
        picThumb.Height = maxHeight
        picThumb.Width = maxWidth
    End If
    
    'initialise GDI+
    submsg = " GDI+ failure"
    token = InitGDIPlus
    
    'load thumbnail image and centre in thumb frame
    picThumb = LoadPictureGDIPlus(strThumbName, picThumb.ScaleWidth, picThumb.ScaleHeight)
     
    'free GDI+
    FreeGDIPlus token
    submsg = ""
    
    picThumb.Left = (picThumbFrame.Width - picThumb.Width) / 2
    picThumb.Top = (picThumbFrame.Height - picThumb.Height) / 2
    
    'set thumb viewport
    thumbViewPort.Width = (picPort.Width / picSource.Width) * picThumb.Width
    thumbViewPort.Height = (picPort.Height / picSource.Height) * picThumb.Height
    
    'flags we are batch processing, we leave before we
    'reset the logo position
    If blnProcessing = True Then Exit Sub
    
    'we start with logo left top position
    Call navLogo_MouseDown(5, 0, 0, 0, 0)
    Call navLogo_MouseUp(5, 0, 0, 0, 0)
    
    'we start with image centre position
    Call navSource_MouseDown(4, 0, 0, 0, 0)
    Call navSource_MouseUp(4, 0, 0, 0, 0)
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Main Sub makeThumb - Error " & Err.Number & ": " & _
              Err.Description & strThumbName & " - " & submsg
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Public Sub onImageLoad()

    'if image width is larger then port width
    If picSource.Width >= 514 Then
        picPort.Width = 514
    Else
        'if smaller we resize the port width
        picPort.Width = picSource.Width
    End If
    
    'if image height is larger then port height
    If picSource.Height >= 379 Then
        picPort.Height = 379
    Else
        'if smaller we resize the port height
        picPort.Height = picSource.Height
    End If
    
    'set source picture box params
    With picSource
        
        If .Height > picPort.Height Then
            MaxMoveY = picPort.Height - .ScaleHeight
        Else
            MaxMoveY = 0
        End If
  
        If .Width > picPort.Width Then
            MaxMoveX = picPort.Width - .ScaleWidth
        Else:
            MaxMoveX = 0
        End If
   
    End With
    
    picSource.Top = 0
    picSource.Left = 0
    
End Sub

Public Sub setNavScrollBars()
    
    'set source picture scrollbars
    
    'set navigation source image scrollbars
    With scrSourceHor
        .SmallChange = 2
        .LargeChange = 10
        .Max = picSource.Width - picPort.Width
    End With
    
    With scrSourceVer
        .SmallChange = 2
        .LargeChange = 10
        .Max = picSource.Height - picPort.Height
    End With
    
    'set navigation logo scroll bars
    scrLogoVer.SmallChange = 2
    scrLogoVer.LargeChange = 10
    scrLogoHor.SmallChange = 2
    scrLogoHor.LargeChange = 10
    
    'text logo ----------------------------------------------------------------------------------
    If blnTextLogo = True Then
    
        'set text label - we use its length and height properties
        'to determine the scrollbars maximum values
        'Note:  This can also be done with the TextWidth and TextHeight
        '       properties, but the result is about the same.
        lblPrintText.Font = txtAdd(1).Font
        lblPrintText.FontSize = txtAdd(1).FontSize
        If txtAdd(1).FontBold = True Then lblPrintText.FontBold = True _
                            Else lblPrintText.FontBold = False
        If txtAdd(1).FontItalic = True Then lblPrintText.FontItalic = True _
                            Else lblPrintText.FontItalic = False
        If txtAdd(1).FontStrikethru = True Then lblPrintText.FontStrikethru = True _
                            Else lblPrintText.FontStrikethru = False
        If txtAdd(1).FontUnderline = True Then lblPrintText.FontUnderline = True _
                            Else lblPrintText.FontUnderline = False
                
        scrollSpeed = 2
        
        'set navigation logo scroll bars
        scrLogoVer.Max = picSource.Height - lblPrintText.Height - 4
        scrLogoHor.Max = picSource.Width - lblPrintText.Width - 8
        
        'logo position indicator
        shpLogoPos.Height = lblPrintText.Height * (picThumb.Height / picSource.Height)
        shpLogoPos.Width = lblPrintText.Width * (picThumb.Width / picSource.Width)
        
        Call textLogo
        
        'the changing time in the textbox triggers the txtAdd_Change event
        'and repositions the logo every time the time changes, we don't
        'want that so in that case we leave here without repositioning the logo
        If chkCurTime.Value = 1 Then Exit Sub
        
        'reposition logo (only for text logo)
        If blnLeftTop = True Then
            Call navLogo_MouseDown(5, 0, 0, 0, 0)
            Call navLogo_MouseUp(5, 0, 0, 0, 0)
        End If
        
        If blnRightTop = True Then
            Call navLogo_MouseDown(6, 0, 0, 0, 0)
            Call navLogo_MouseUp(6, 0, 0, 0, 0)
        End If
        
        If blnLeftBot = True Then
            Call navLogo_MouseDown(7, 0, 0, 0, 0)
            Call navLogo_MouseUp(7, 0, 0, 0, 0)
        End If
        
        If blnRightBot = True Then
            Call navLogo_MouseDown(8, 0, 0, 0, 0)
            Call navLogo_MouseUp(8, 0, 0, 0, 0)
        End If
        
        If blnCenter = True Then
            Call navLogo_MouseDown(4, 0, 0, 0, 0)
            Call navLogo_MouseUp(4, 0, 0, 0, 0)
        End If
        
    End If
    
    'normal logo --------------------------------------------------------------------------------
    If blnNormalLogo = True Then
        scrollSpeed = 3
        scrLogoHor.Max = picSource.Width - picNormalSource.Width
        scrLogoVer.Max = picSource.Height - picNormalSource.Height
        'logo position indicator
        shpLogoPos.Height = picNormalSource.Height * (picThumb.Height / picSource.Height)
        shpLogoPos.Width = picNormalSource.Width * (picThumb.Width / picSource.Width)
        Call normalLogo
    End If
    
    'masked logo --------------------------------------------------------------------------------
    If blnMaskedLogo = True Then
        scrollSpeed = 3
        scrLogoHor.Max = picSource.Width - picMaskedSource.Width
        scrLogoVer.Max = picSource.Height - picMaskedSource.Height
        'logo position indicator
        shpLogoPos.Height = picMaskedSource.Height * (picThumb.Height / picSource.Height)
        shpLogoPos.Width = picMaskedSource.Width * (picThumb.Width / picSource.Width)
        Call maskedLogo
    End If
    
End Sub

Private Sub setLogoScrollBars()

    'normal logo
    If picNormalSource.Width <= portNormalLogo.Width Then normLogoHorScroll.Max = 0
    If picNormalSource.Height <= portNormalLogo.Height Then normLogoVerScroll.Max = 0
    
    'horizontal
    If picNormalSource.Width >= portNormalLogo.Width Then
        normLogoHorScroll.Max = (picNormalLogo.Width - portNormalLogo.Width) + 13
        normLogoHorScroll.Value = 0
    End If
    
    'vertical
    If picNormalSource.Height >= portNormalLogo.Height Then
        normLogoVerScroll.Max = (picNormalLogo.Height - portNormalLogo.Height) + 13
        normLogoVerScroll.Value = 0
    End If
    
    'masked logo
    If picMaskedSource.Width <= portMaskedLogo.Width Then maskLogoHorScroll.Max = 0
    If picMaskedSource.Height <= portMaskedLogo.Height Then maskLogoVerScroll.Max = 0
    
    'horizontal
    If picMaskedSource.Width >= portMaskedLogo.Width Then
        maskLogoHorScroll.Max = (picMaskedLogo.Width - portMaskedLogo.Width) + 13
        maskLogoHorScroll.Value = 0
    End If
    
    'vertical
    If picMaskedSource.Height >= portMaskedLogo.Height Then
        maskLogoVerScroll.Max = (picMaskedLogo.Height - portMaskedLogo.Height) + 13
        maskLogoVerScroll.Value = 0
    End If
    
    If blnNormalLogo = True Then
        Call normLogoHorScroll_Scroll
        Call normLogoVerScroll_Scroll
    End If
    
    If blnMaskedLogo = True Then
        Call maskLogoHorScroll_Scroll
        Call maskLogoVerScroll_Scroll
    End If
    
End Sub

Private Sub disableTextControls()
    
    Dim n As Integer
    
    For n = 0 To 4
        txtAdd(n).Enabled = False
        chkAddText(n).Enabled = False
    Next n
    
    For n = 0 To 7
        imgFontButton(n).Enabled = False
    Next n
    
    picShadowColor.Enabled = False
    picTextColor.Enabled = False
    
    chkShadow.Enabled = False
    chkCurDate.Enabled = False
    chkCurTime.Enabled = False
    chkAutoIncrement.Enabled = False
    
    udcShadowOffset.Enabled = False
    txtOffset.Enabled = False
    cboFontSize.Enabled = False
    txtFontName.Enabled = False
    
End Sub

Private Sub enableTextControls()
    
    Dim n As Integer
    
    For n = 0 To 4
        txtAdd(n).Enabled = True
        chkAddText(n).Enabled = True
    Next n
    
    For n = 0 To 7
        imgFontButton(n).Enabled = True
    Next n
    
    picShadowColor.Enabled = True
    picTextColor.Enabled = True
    
    chkShadow.Enabled = True
    chkCurDate.Enabled = True
    chkCurTime.Enabled = True
    chkAutoIncrement.Enabled = True
    
    udcShadowOffset.Enabled = True
    txtOffset.Enabled = True
    cboFontSize.Enabled = True
    txtFontName.Enabled = True
    
End Sub

'+++++++++++++++++++++++++++++++++++++++++ FORM BUTTONS +++++++++++++++++++++++++++++++++++++++++

Private Sub resetButtons()
    
    'reset buttons
    imgClose.Picture = close_norm
    imgMin.Picture = min_norm
        
End Sub

Private Sub imgMin_Click()
    
    'minimize button
    Me.WindowState = 1
    
End Sub

Private Sub imgMin_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'minimize button
    imgMin.Picture = min_down
    
End Sub

Private Sub imgMin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'minimize button
    imgMin.Picture = min_hot
    
End Sub

Private Sub imgMin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'minimize button
    imgMin.Picture = min_norm
    
End Sub

Private Sub imgClose_Click()
    
    'exit button
    
    Dim msg     As String       'message box
    
    'no exit if Batch Process is running
    If blnProcessing = True Then
        msg = "Batch Process is in progress, cancel the          " & Chr(13) & _
              "process before exiting Photo Logo Plus. "
        MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
        Exit Sub
    End If
    
    'no exit if Batch Convert is running
    If blnConverting = True Then
        msg = "Batch Convert is in progress, cancel the          " & Chr(13) & _
              "conversion before exiting Photo LOgo Plus. "
        MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
        Exit Sub
    End If
    
    Unload Me
    
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'exit button
    imgClose.Picture = close_down
    
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'exit button
    imgClose.Picture = close_hot
    
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'exit button
    imgClose.Picture = close_norm
    
End Sub

Private Sub picTop_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 1 Then
        'move form without caption
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If

End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Call resetButtons
    
End Sub







    








