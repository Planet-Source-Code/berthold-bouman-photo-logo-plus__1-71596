VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Batch Process"
   ClientHeight    =   9450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   Icon            =   "frmBatch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   630
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3495
      TabIndex        =   37
      Top             =   10365
      Width           =   1980
   End
   Begin VB.PictureBox picBot 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      Picture         =   "frmBatch.frx":08CA
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   17
      Top             =   8925
      Width           =   6465
      Begin MSComctlLib.ProgressBar Prog 
         Height          =   165
         Left            =   165
         TabIndex        =   18
         Top             =   195
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.FileListBox File3 
      Height          =   285
      Left            =   3510
      TabIndex        =   23
      Top             =   9960
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.PictureBox PicTop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   0
      Picture         =   "frmBatch.frx":BA3C
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   19
      Top             =   0
      Width           =   6465
      Begin VB.Image imgClose 
         Height          =   285
         Left            =   6045
         Picture         =   "frmBatch.frx":1D5FE
         Top             =   105
         Width           =   285
      End
      Begin VB.Image imgMin 
         Height          =   285
         Left            =   5700
         Picture         =   "frmBatch.frx":1DAB4
         Top             =   105
         Width           =   285
      End
      Begin VB.Label lblSelSource 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Source Directory:"
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
         Left            =   180
         TabIndex        =   21
         Top             =   510
         Width           =   2685
      End
      Begin VB.Label lblSelOutput 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Output Directory:"
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
         Left            =   3300
         TabIndex        =   20
         Top             =   510
         Width           =   2685
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
      Height          =   4590
      Left            =   135
      TabIndex        =   16
      Top             =   1275
      Width           =   3060
   End
   Begin VB.FileListBox File2 
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
      Height          =   4590
      Left            =   3270
      TabIndex        =   15
      Top             =   1275
      Width           =   3060
   End
   Begin VB.TextBox txtSource 
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
      Top             =   900
      Width           =   2475
   End
   Begin VB.TextBox txtOutput 
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
      Left            =   3270
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   900
      Width           =   2475
   End
   Begin VB.PictureBox picSelLogoframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   135
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   202
      TabIndex        =   6
      Top             =   6495
      Width           =   3060
      Begin VB.CheckBox chkPromptErrors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   35
         Top             =   1665
         Width           =   195
      End
      Begin VB.CheckBox chkPromptOverwrite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         TabIndex        =   25
         Top             =   1290
         Width           =   195
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   1605
         Picture         =   "frmBatch.frx":1DF6A
         ScaleHeight     =   71
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   24
         Top             =   105
         Width           =   1335
         Begin VB.Image imgPosition 
            Height          =   180
            Index           =   4
            Left            =   1065
            Picture         =   "frmBatch.frx":1F51A
            Top             =   795
            Width           =   180
         End
         Begin VB.Image imgPosition 
            Height          =   180
            Index           =   3
            Left            =   90
            Picture         =   "frmBatch.frx":1F894
            Top             =   795
            Width           =   180
         End
         Begin VB.Image imgPosition 
            Height          =   180
            Index           =   2
            Left            =   585
            Picture         =   "frmBatch.frx":1FC0E
            Top             =   435
            Width           =   180
         End
         Begin VB.Image imgPosition 
            Height          =   180
            Index           =   1
            Left            =   1065
            Picture         =   "frmBatch.frx":1FF88
            Top             =   90
            Width           =   180
         End
         Begin VB.Image imgPosition 
            Height          =   180
            Index           =   0
            Left            =   90
            Picture         =   "frmBatch.frx":20302
            Top             =   90
            Width           =   180
         End
      End
      Begin VB.OptionButton optLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Masked Logo"
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
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   9
         Top             =   900
         Width           =   195
      End
      Begin VB.OptionButton optLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Normal Logo"
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
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   510
         Width           =   195
      End
      Begin VB.OptionButton optLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Plain Text"
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
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   135
         Width           =   195
      End
      Begin VB.Label lblPromptErrors 
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt for Errors"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   420
         TabIndex        =   36
         Top             =   1665
         Width           =   1395
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblPromptOverwrite 
         BackStyle       =   0  'Transparent
         Caption         =   "Prompt to Overwrite"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   420
         TabIndex        =   26
         Top             =   1290
         Width           =   1860
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblText 
         BackStyle       =   0  'Transparent
         Caption         =   "Text Logo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   390
         TabIndex        =   12
         Top             =   105
         Width           =   900
      End
      Begin VB.Label lblNormal 
         BackStyle       =   0  'Transparent
         Caption         =   "Normal Logo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   390
         TabIndex        =   11
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label lblMasked 
         BackStyle       =   0  'Transparent
         Caption         =   "Masked Logo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   390
         TabIndex        =   10
         Top             =   870
         Width           =   1140
      End
   End
   Begin VB.PictureBox picSeFormatFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   3270
      ScaleHeight     =   135
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   202
      TabIndex        =   3
      Top             =   6495
      Width           =   3060
      Begin VB.OptionButton optFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   38
         Top             =   165
         Width           =   195
      End
      Begin VB.OptionButton optFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   31
         Top             =   465
         Width           =   195
      End
      Begin VB.OptionButton optFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   29
         Top             =   1065
         Width           =   195
      End
      Begin VB.OptionButton optFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   27
         Top             =   1665
         Width           =   195
      End
      Begin VB.OptionButton optFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   28
         Top             =   1365
         Width           =   195
      End
      Begin VB.OptionButton optFileType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   30
         Top             =   765
         Width           =   195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No Changes"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   465
         TabIndex        =   39
         Top             =   150
         Width           =   1020
      End
      Begin VB.Label lblTIF 
         BackStyle       =   0  'Transparent
         Caption         =   "Save as TIF"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   450
         TabIndex        =   34
         Top             =   1635
         Width           =   1020
      End
      Begin VB.Label lblPNG 
         BackStyle       =   0  'Transparent
         Caption         =   "Save as PNG"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   450
         TabIndex        =   33
         Top             =   1335
         Width           =   1020
      End
      Begin VB.Label lblGIF 
         BackStyle       =   0  'Transparent
         Caption         =   "Save as GIF"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   450
         TabIndex        =   32
         Top             =   1035
         Width           =   1020
      End
      Begin VB.Image imgStart 
         Height          =   375
         Left            =   1965
         Picture         =   "frmBatch.frx":2067C
         Top             =   585
         Width           =   960
      End
      Begin VB.Image imgCancel 
         Height          =   375
         Left            =   1965
         Picture         =   "frmBatch.frx":20B0A
         Top             =   1065
         Width           =   960
      End
      Begin VB.Image imgExit 
         Height          =   375
         Left            =   1965
         Picture         =   "frmBatch.frx":20FD0
         Top             =   1545
         Width           =   960
      End
      Begin VB.Label lblBMP 
         BackStyle       =   0  'Transparent
         Caption         =   "Save as BMP"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   450
         TabIndex        =   5
         Top             =   435
         Width           =   1020
      End
      Begin VB.Label lblJPG 
         BackStyle       =   0  'Transparent
         Caption         =   "Save as JPG"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   450
         TabIndex        =   4
         Top             =   735
         Width           =   1020
      End
   End
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   75
      Picture         =   "frmBatch.frx":2143E
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   0
      Top             =   5955
      Width           =   6315
      Begin VB.Label lblFileFormat 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Output File Format:"
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
         Left            =   3360
         TabIndex        =   2
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label lbLogolSelection 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Logo and Position:"
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
         Left            =   210
         TabIndex        =   1
         Top             =   120
         Width           =   2355
      End
   End
   Begin VB.Image opt_down 
      Height          =   180
      Left            =   1905
      Picture         =   "frmBatch.frx":2AD90
      Top             =   9975
      Width           =   180
   End
   Begin VB.Image opt_norm 
      Height          =   180
      Left            =   1695
      Picture         =   "frmBatch.frx":2B0F5
      Top             =   9975
      Width           =   180
   End
   Begin VB.Image but_start_norm 
      Height          =   375
      Left            =   465
      Picture         =   "frmBatch.frx":2B46F
      Top             =   10830
      Width           =   960
   End
   Begin VB.Image but_cancel_norm 
      Height          =   375
      Left            =   1485
      Picture         =   "frmBatch.frx":2B8FD
      Top             =   10830
      Width           =   960
   End
   Begin VB.Image but_start_down 
      Height          =   375
      Left            =   465
      Picture         =   "frmBatch.frx":2BDC3
      Top             =   11250
      Width           =   960
   End
   Begin VB.Image but_cancel_down 
      Height          =   375
      Left            =   1485
      Picture         =   "frmBatch.frx":2C237
      Top             =   11250
      Width           =   960
   End
   Begin VB.Image but_exit_norm 
      Height          =   375
      Left            =   2490
      Picture         =   "frmBatch.frx":2C6DA
      Top             =   10830
      Width           =   960
   End
   Begin VB.Image but_exit_down 
      Height          =   375
      Left            =   2490
      Picture         =   "frmBatch.frx":2CB48
      Top             =   11250
      Width           =   960
   End
   Begin VB.Image imgOpenOutput 
      Height          =   315
      Left            =   5790
      Picture         =   "frmBatch.frx":2CF99
      Top             =   900
      Width           =   540
   End
   Begin VB.Image imgOpenSource 
      Height          =   315
      Left            =   2655
      Picture         =   "frmBatch.frx":2D8B7
      Top             =   900
      Width           =   540
   End
   Begin VB.Image folder_down 
      Height          =   315
      Left            =   2415
      Picture         =   "frmBatch.frx":2E1D5
      Top             =   10320
      Width           =   540
   End
   Begin VB.Image folder_norm 
      Height          =   315
      Left            =   2415
      Picture         =   "frmBatch.frx":2E5DF
      Top             =   9975
      Width           =   540
   End
   Begin VB.Image curHand 
      Height          =   480
      Left            =   1815
      Picture         =   "frmBatch.frx":2E9F5
      Top             =   10305
      Width           =   480
   End
   Begin VB.Image close_down 
      Height          =   285
      Left            =   1200
      Picture         =   "frmBatch.frx":2EB47
      Top             =   10350
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_hot 
      Height          =   285
      Left            =   855
      Picture         =   "frmBatch.frx":2EFFD
      Top             =   10350
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_norm 
      Height          =   285
      Left            =   495
      Picture         =   "frmBatch.frx":2F4B3
      Top             =   10350
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_down 
      Height          =   285
      Left            =   1200
      Picture         =   "frmBatch.frx":2F969
      Top             =   9975
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_hot 
      Height          =   285
      Left            =   855
      Picture         =   "frmBatch.frx":2FE1F
      Top             =   9975
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_norm 
      Height          =   285
      Left            =   495
      Picture         =   "frmBatch.frx":302D5
      Top             =   9975
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgLeft 
      Height          =   8130
      Left            =   0
      Picture         =   "frmBatch.frx":3078B
      Stretch         =   -1  'True
      Top             =   825
      Width           =   75
   End
   Begin VB.Image imgRight 
      Height          =   8115
      Left            =   6390
      Picture         =   "frmBatch.frx":309DD
      Stretch         =   -1  'True
      Top             =   840
      Width           =   75
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Processing:"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   150
      TabIndex        =   22
      Top             =   8640
      Width           =   6165
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBatch"
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

Private blnCancel           As Boolean          'cancel process
Private strExtention        As String           'file extention we convert to
Private errCount            As Integer          'does what it says
Private m                   As Form             'main form

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Form_Load()
    
    On Error GoTo ErrHandler
    
    Dim mTop        As Single       'form top position
    Dim mLeft       As Single       'form left position
    Dim msg         As String       'message box
    
    'saves me a lot of typing
    Set m = frmMain
    
    'retrieve top and left coordinates from register
    mTop = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch", "batchTop")
    mLeft = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch", "batchLeft")
    
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
    
    'set default output format
    optFileType(5).Value = 1
    
    Call setControls
    
    Call resetSerialNumber
        
    'retrieve current logo type from frmMain
    If blnTextLogo = True Then optLogo(0).Value = 1
    If blnNormalLogo = True Then optLogo(1).Value = 1
    If blnMaskedLogo = True Then optLogo(2).Value = 1
    
    'retrieve current logo position from frmMain
    If blnLeftTop = True Then imgPosition(0).Picture = opt_down.Picture
    If blnRightTop = True Then imgPosition(1).Picture = opt_down.Picture
    If blnCenter = True Then imgPosition(2).Picture = opt_down.Picture
    If blnLeftBot = True Then imgPosition(3).Picture = opt_down.Picture
    If blnRightBot = True Then imgPosition(4).Picture = opt_down.Picture
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Process Form_Load - Error " & Err.Number & ": " & Err.Description
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    blnBatch = False
    
    'write to register - first we create a key - in case there isn't one,
    'e.g. when the app runs for the first time
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch"
    
    'now we can write the values to the register
    'save top and left position to register
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch", _
                                   "batchTop", Me.Top, REG_SZ
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch", _
                                   "batchLeft", Me.Left, REG_SZ
                                   
End Sub

'+++++++++++++++++++++++++++++++++++++++++ CONTROL EVENTS +++++++++++++++++++++++++++++++++++++++

Private Sub imgOpenSource_Click()
    
    'open source directory
    
    On Error GoTo ErrHandler
    
    Dim strResFolder    As String       'open folder
    Dim strOldFolder    As String       'old folder path
    Dim msg             As String       'messagebox
    
    strOldFolder = txtSource.Text
    strResFolder = BrowseForFolder(hwnd, "Photo Logo Plus - Select a folder")
    'if cancel was selected
    If strResFolder = "" Then strResFolder = strOldFolder
    
     'if source directory is the same as output directory
    If strResFolder = txtOutput.Text Then
        
        msg = "Source Directory cannot be the same as the     " & Chr(13) & _
            "Output Directory, select another Directory.    "
        MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
        Exit Sub
        
    End If
    
    File1.Path = strResFolder
    File3.Path = strResFolder
    txtSource.Text = strResFolder
    'save last used folder path to register
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch"
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch", _
                                           "Last Source Folder", txtSource.Text, REG_SZ
                                           
    'update label
    lblStatus.Caption = File3.ListCount & " Images in Source Directory"
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Process imgOpenSource_Click - Error " & Err.Number & ": " & _
              Err.Description & " " & strResFolder
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub imgOpenSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'open source directory
    imgOpenSource.Picture = folder_down.Picture
    
End Sub

Private Sub imgOpenSource_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'open source directory
    imgOpenSource.Picture = folder_norm.Picture
    
End Sub

Private Sub imgOpenOutput_Click()
    
    'open output directory
    
    On Error GoTo ErrHandler
    
    Dim strResFolder    As String       'open folder
    Dim strOldFolder    As String       'old folder path
    Dim msg             As String       'messagebox
    
    strOldFolder = txtOutput.Text
    strResFolder = BrowseForFolder(hwnd, "Photo Logo Plus - Select a folder")
    'if cancel was selected
    If strResFolder = "" Then strResFolder = strOldFolder
    
    'if source directory is the same as output directory
    If strResFolder = txtSource.Text Then
        
        msg = "Output Directory cannot be the same as the     " & Chr(13) & _
            "Source Directory, select another Directory.    "
        MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
        Exit Sub
        
    End If
    
    File2.Path = strResFolder
    txtOutput.Text = strResFolder
    'save last used folder path to register
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch"
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch", _
                                           "Last Output Folder", txtOutput.Text, REG_SZ
                                           
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Process imgOpenOutput_Click - Error " & Err.Number & ": " & _
              Err.Description & " " & strResFolder
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub imgOpenOutput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'open Output directory
    imgOpenOutput.Picture = folder_down.Picture
    
End Sub

Private Sub imgOpenOutput_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'open source directory
    imgOpenOutput.Picture = folder_norm.Picture
    
End Sub

Private Sub imgStart_Click()
    
    'start processing
    
    Dim msg     As String       'message box
    
    'first check if there are files to process
    If File1.ListCount = 0 Then
        msg = "There are no files to process. Select another directory.       "
        MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
        Exit Sub
    End If
    
    'check if the number of files in the selected folder doesn't
    'exceed the string format of the serial number
    If m.chkAutoIncrement.Value = 1 Then
        If Len(strDigits) = 3 And File1.ListCount > 999 _
        Or Len(strDigits) = 4 And File1.ListCount > 9999 _
        Or Len(strDigits) = 5 And File1.ListCount > 99999 Then
            msg = "The number of files in this directory exceeds the" & Chr(13) & _
                  "format of the serial number you have selected." & Chr(13) & _
                  "Select another serial number format in the Main        " & Chr(13) & _
                  "window. Batch Process will be cancelled."
            MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
            Exit Sub
        End If
    End If
    
    'set flags
    blnProcessing = True
    blnCancel = False
    
    List1.Clear
    
    Call startProcess
    
End Sub

Private Sub imgStart_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'start processing
    imgStart.Picture = but_start_down.Picture
    
End Sub

Private Sub imgStart_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'start processing
    imgStart.Picture = but_start_norm.Picture
    
End Sub

Private Sub imgCancel_Click()
    
    'cancel process
    
    Dim msg     As String       'message box
    Dim retVal                  'message box
    
    'do or don't cancel process
    If blnProcessing = True Then
        msg = "Are you sure you want to cancel Batch Process?        "
        retVal = MsgBox(msg, vbExclamation + vbYesNo, "Photo Logo Plus")
        If retVal = vbNo Then
            Exit Sub
        ElseIf retVal = vbYes Then
            'write to error log
            msg = Now & " Process imgCancel_Click - User cancelled process"
            Call writeErrorLog(strErrLog, msg)
            blnProcessing = False
            blnCancel = True
            Exit Sub
        End If
    End If
    
End Sub

Private Sub imgCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'start process
    imgCancel.Picture = but_cancel_down.Picture
    
End Sub

Private Sub imgCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'start process
    imgCancel.Picture = but_cancel_norm.Picture
    
End Sub

Private Sub imgExit_Click()
    
    'exit batch process
    
    Dim msg     As String       'message box
    
    'no exit if still processing
    If blnProcessing = True Then
        msg = "Batch Process is in progress, cancel the          " & Chr(13) & _
              "process before exiting Batch Process. "
        MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
        Exit Sub
    End If
    
    Unload Me
    
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'exit batch process
    imgExit.Picture = but_exit_down.Picture
    
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'exit batch process
    imgExit.Picture = but_exit_norm.Picture
    
End Sub

Private Sub optLogo_Click(Index As Integer)
    
    'transfer to main form
    If blnProcessing = True Then Exit Sub
    
    Dim n As Integer
    
    m.optLogo(Index).Value = optLogo(Index).Value
    
    'reset buttons
    For n = 0 To 4
        imgPosition(n).Picture = opt_norm.Picture
    Next n
    
    'we always start at the top left corner
    imgPosition(0).Picture = opt_down.Picture
    
End Sub

Private Sub optFileType_Click(Index As Integer)
        
    'select file format (and extention) for saved images
    If blnProcessing = True Then Exit Sub
    
    Select Case Index
        
        Case 0          'save as BMP
            
            strExtention = "bmp"
        
        Case 1          'save as JPG
            
            strExtention = "jpg"
            
        Case 2          'save as GIF
            
            strExtention = "gif"
            
        Case 3          'save as PNG
            
            strExtention = "png"
            
        Case 4          'save as TIF
            
            strExtention = "tif"
        
        Case 5          'don't change file format
            
            strExtention = "No Changes"
            
    End Select
            
End Sub

Public Sub imgPosition_Click(Index As Integer)
    
    'set logo position
    
    Dim n As Integer
    
    'reset images
    For n = 0 To 4
        imgPosition(n).Picture = opt_norm.Picture
    Next n
    
    Call resetPostionFlags
    
    Select Case Index
        
        Case 0          'left top
            blnLeftTop = True
            imgPosition(0).Picture = opt_down.Picture
            Call m.navLogo_MouseDown(5, 0, 0, 0, 0)
            Call m.navLogo_MouseUp(5, 0, 0, 0, 0)
        Case 1          'right top
            blnRightTop = True
            imgPosition(1).Picture = opt_down.Picture
            Call m.navLogo_MouseDown(6, 0, 0, 0, 0)
            Call m.navLogo_MouseUp(6, 0, 0, 0, 0)
        Case 2          'centre image
            blnCenter = True
            imgPosition(2).Picture = opt_down.Picture
            Call m.navLogo_MouseDown(4, 0, 0, 0, 0)
            Call m.navLogo_MouseUp(4, 0, 0, 0, 0)
        Case 3          'left bottom
            blnLeftTop = True
            imgPosition(3).Picture = opt_down.Picture
            Call m.navLogo_MouseDown(7, 0, 0, 0, 0)
            Call m.navLogo_MouseUp(7, 0, 0, 0, 0)
        Case 4          'right bottom
            blnRightBot = True
            imgPosition(4).Picture = opt_down.Picture
            Call m.navLogo_MouseDown(8, 0, 0, 0, 0)
            Call m.navLogo_MouseUp(8, 0, 0, 0, 0)
        
    End Select
    
End Sub

'++++++++++++++++++++++++++++++++++++++++++++ COMMON SUBS +++++++++++++++++++++++++++++++++++++++

Private Sub setControls()
    
    On Error GoTo ErrHandler
    
    Dim msg         As String       'message box
    Dim n           As Integer      'counter
    Dim strFolder   As String       'folder path
    
    'colors
    For n = 0 To Me.Height Step 3
        'paint background
        Me.Line (0, n)-(Me.Width, n), RGB(190, 212, 255)
        Me.Line (0, n + 1)-(Me.Width, n + 1), RGB(204, 224, 255)
        Me.Line (0, n + 2)-(Me.Width, n + 2), RGB(255, 255, 255)
    Next n
  
    For n = 0 To 2
        optLogo(n).MousePointer = 99
        optLogo(n).MouseIcon = curHand
        'retrieve logo format from form main
        optLogo(n).Value = m.optLogo(n).Value
    Next n
    
    'cursors
    imgMin.MousePointer = 99
    imgMin.MouseIcon = curHand
    imgClose.MousePointer = 99
    imgClose.MouseIcon = curHand
    
    imgOpenSource.Picture = folder_norm.Picture
    imgOpenSource.MousePointer = 99
    imgOpenSource.MouseIcon = curHand
    imgOpenOutput.Picture = folder_norm.Picture
    imgOpenOutput.MousePointer = 99
    imgOpenOutput.MouseIcon = curHand
    
    imgStart.MousePointer = 99
    imgStart.MouseIcon = curHand
    imgCancel.MousePointer = 99
    imgCancel.MouseIcon = curHand
    imgExit.MousePointer = 99
    imgExit.MouseIcon = curHand
    
    chkPromptOverwrite.MousePointer = 99
    chkPromptOverwrite.MouseIcon = curHand
    
    chkPromptErrors.MousePointer = 99
    chkPromptErrors.MouseIcon = curHand
    
    For n = 0 To 4
        imgPosition(n).MousePointer = 99
        imgPosition(n).MouseIcon = curHand
        imgPosition(n).Picture = opt_norm.Picture
    Next n
    
    For n = 0 To 5
        optFileType(n).MousePointer = 99
        optFileType(n).MouseIcon = curHand
    Next n
    
    'filelistboxes
    File1.Pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.wmf;*.png;*.tif;*.tiff"
    File2.Pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.wmf;*.png;*.tif;*.tiff"
    File3.Pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.wmf;*.png;*.tif;*.tiff"
    
    'retrieve last used folders from register
    strFolder = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch", "Last Source Folder")
    If strFolder = "" Then strFolder = App.Path
    File1.Path = strFolder
    File3.Path = strFolder
    txtSource.Text = File1.Path
    
    'update label
    lblStatus.Caption = File3.ListCount & " Images in Source Directory"
    
    strFolder = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Batch", "Last Output Folder")
    If strFolder = "" Then strFolder = App.Path
    File2.Path = strFolder
    txtOutput.Text = File2.Path
    
    'change progbar colors
    Call SendMessage(Prog.hwnd, PBM_SETBKCOLOR, 0&, ByVal &HFFC0C0)
    Call SendMessage(Prog.hwnd, PBM_SETBARCOLOR, 0&, ByVal vbBlue)
    Prog.Value = 100
    Prog.Visible = False
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Process Sub SetControls - Error " & Err.Number & ": " & Err.Description
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub disableControls()
    
    'disable controls while processing
    
    Dim n As Integer
    
    imgStart.Enabled = False
    imgOpenSource.Enabled = False
    imgOpenOutput.Enabled = False
    
    For n = 0 To 4
        imgPosition(n).Enabled = False
    Next n
        
    For n = 0 To 2
        optLogo(n).MousePointer = 12
    Next n
    
    For n = 0 To 5
        optFileType(n).MousePointer = 12
    Next n
    
End Sub

Private Sub enableControls()
    
    'disable controls while processing
    
    Dim n As Integer
    
    imgStart.Enabled = True
    imgOpenSource.Enabled = True
    imgOpenOutput.Enabled = True
    
    For n = 0 To 4
        imgPosition(n).Enabled = True
    Next n
    
    For n = 0 To 2
        optLogo(n).MousePointer = 99
    Next n
    
    For n = 0 To 5
        optFileType(n).MousePointer = 99
    Next n
    
End Sub

Private Sub startProcess()
    
    'batch process images
    
    On Error GoTo ErrHandler
    
    Dim token           As Long         'GDI+
    Dim n               As Integer      'counter
    Dim strFilename     As String       'filename string, including jpg or bmp extention
    Dim strOldExt       As String       'flename extention
    Dim msg             As String       'message box
    Dim retVal                          'message box
    
    'start process report
    strReport = App.Path & "\Reports\" & Format(Now, "ddmmyyhhmmss") & ".tmp"
    Call startReport(File1.Path, File2.Path, "ALL", UCase(strExtention))
    
    errCount = 0
    
    'set progbar
    Prog.Visible = True
    Prog.Max = File3.ListCount
    Prog.Value = 0
    
    'reset serial number
    Call resetSerialNumber
    
    Call disableControls
    
    'hide thumbnail viewport and logo indicator main form
    m.thumbViewPort.Visible = False
    m.shpLogoPos.Visible = False
    
    For n = 0 To File3.ListCount - 1
        
        'we leave if processing is cancelled ----------------------------------------------------
        If blnCancel = True Then
            lblStatus.Caption = "Processed: " & n + 1 & " Images to " & _
                        File2.Path & " - Batch Process was cancelled"
            m.lblFileName.Caption = lblStatus.Caption
            Call addReport("")
            Call addReport("Batch Process was cancelled by user.")
            Call exitProcess
            Call enableControls
            Exit Sub
        End If '---------------------------------------------------------------------------------
        
        File3.ListIndex = n
        Prog.Value = n
        
        'load new source image
        token = InitGDIPlus
        m.picSource.Picture = LoadPictureGDIPlus(fixPath(File3.Path, File3.FileName))
        
        'if picture could not be loaded ---------------------------------------------------------
        If Err.Number = 999 Then
            errCount = errCount + 1
            'write to error log
            msg = Now & " Process Sub startProcess GDI+ failure 999 - " & _
            "File: " & fixPath(File3.Path, File3.FileName)
            Call writeErrorLog(strErrLog, msg)
            msg = "Could not process File: " & fixPath(File3.Path, File3.FileName)
            Call addReport(msg)
            If chkPromptErrors.Value = 1 Then
                'show message
                msg = "Error loading picture, not a valid bitmap." & Chr(13) & _
                "File: " & fixPath(File3.Path, File3.FileName) & Chr(13) & Chr(13) & _
                        "Do you want to continue Batch Conversion?        "
                retVal = MsgBox(msg, vbYesNo + vbExclamation, "Photo Logo Plus")
                'when too many errors occur, we must offer the user a way out
                If retVal = vbNo Then
                    lblStatus.Caption = "Processed: " & n + 1 & " Images to " & _
                    File2.Path & " - Process was cancelled"
                    'write to error log
                    msg = Now & " Process sub startProcess - User cancelled process " & strFilename
                    Call writeErrorLog(strErrLog, msg)
                    Call addReport("")
                    Call addReport("Batch Process was cancelled by user.")
                    'free GDI+
                    FreeGDIPlus token
                    Call exitProcess
                    Call enableControls
                    Exit Sub
                End If
            End If
        Else
            List1.AddItem "Processed: " & File3.FileName
        End If '---------------------------------------------------------------------------------
        
        'free GDI+
        FreeGDIPlus token
                
        'update status label
        If Len("Processing: " & n + 1 & "/" & File3.ListCount & _
                            " - " & fixPath(File3.Path, File3.FileName)) >= 62 Then
            'get rid of excessive long path names
            lblStatus.Caption = Mid("Processing: " & n + 1 & "/" & File3.ListCount & _
                            " - " & fixPath(File3.Path, File3.FileName), 1, 62) & "..."
        Else
            lblStatus.Caption = "Processing: " & n + 1 & "/" & File3.ListCount & _
                            " - " & fixPath(File3.Path, File3.FileName)
        End If
        m.lblFileName.Caption = lblStatus.Caption
        
        'set our thumbnail file string
        strThumbName = fixPath(File3.Path, File3.FileName)
                
        'set our parameters
        Call m.onImageLoad
        Call m.makeThumb
        Call m.setNavScrollBars
        
        'set logo position
        If blnCenter = True Then Call m.navLogo_MouseDown(4, 0, 0, 0, 0)
        If blnLeftTop = True Then Call m.navLogo_MouseDown(5, 0, 0, 0, 0)
        If blnRightTop = True Then Call m.navLogo_MouseDown(6, 0, 0, 0, 0)
        If blnLeftBot = True Then Call m.navLogo_MouseDown(7, 0, 0, 0, 0)
        If blnRightBot = True Then Call m.navLogo_MouseDown(8, 0, 0, 0, 0)
        
        'reset button images
        If blnCenter = True Then Call m.navLogo_MouseUp(4, 0, 0, 0, 0)
        If blnLeftTop = True Then Call m.navLogo_MouseUp(5, 0, 0, 0, 0)
        If blnRightTop = True Then Call m.navLogo_MouseUp(6, 0, 0, 0, 0)
        If blnLeftBot = True Then Call m.navLogo_MouseUp(7, 0, 0, 0, 0)
        If blnRightBot = True Then Call m.navLogo_MouseUp(8, 0, 0, 0, 0)
                
        'select logo type
        If blnTextLogo = True Then Call m.textLogo
        If blnNormalLogo = True Then Call m.normalLogo
        If blnMaskedLogo = True Then Call m.maskedLogo
    
        'normal image logo
        If blnNormalLogo = True Then
            'alphablending
            Call frmMain.scrBlend_Change
            'blit normal logo to picSource
            BitBlt m.picSource.hDC, _
                m.scrLogoHor.Value, _
                m.scrLogoVer.Value, _
                m.picNormalLogo.ScaleWidth, _
                m.picNormalLogo.ScaleHeight, _
                m.picNormalLogo.hDC, _
                0, _
                0, _
                vbSrcCopy
        
        End If
        
        'masked image logo
        If blnMaskedLogo = True Then
            'blit masked logo to picSource
            TransparentBlt m.picSource.hDC, _
                m.scrLogoHor.Value, _
                m.scrLogoVer.Value, _
                m.picMaskedLogo.ScaleWidth, _
                m.picMaskedLogo.ScaleHeight, _
                m.picMaskedLogo.hDC, _
                0, _
                0, _
                m.picMaskedLogo.ScaleWidth, _
                m.picMaskedLogo.ScaleHeight, _
                GetPixel(m.picMaskedSource.hDC, 0, 0)
        End If
        
        DoEvents 'we need this to refresh the screen
        
        If optFileType(5) = True Then
            'don't change file format
            strFilename = fixPath(File2.Path, File3.FileName)
        Else
            'create filename string - get file extention first, we need
            'to know the length of the file extention to get rid of it
            strOldExt = GetFileExtention(File3.FileName)
            'now we remove the old extention and add the new extention
            strFilename = fixPath(File2.Path, _
                            Mid(File3.FileName, 1, Len(File3.FileName) - _
                            Len(strOldExt)) & strExtention)
        End If
        
        'check if file aleady exists ------------------------------------------------------------
        If chkPromptOverwrite.Value = 1 Then
            If FileExists(strFilename) = True Then
                msg = "Filename already exists, do you want to       " & Chr(13) & _
                      "replace the existing file with the new file?" & Chr(13) & Chr(13) & _
                      "Select Cancel to stop Batch Process.      "
                retVal = MsgBox(msg, vbExclamation + vbYesNoCancel, "Photo Logo Plus")
                'don't replace
                If retVal = vbNo Then GoTo Skip
                'cancel the conversion
                If retVal = vbCancel Then
                    lblStatus.Caption = "Processed: " & n + 1 & " Images to " & _
                    File2.Path & " - Batch Process was cancelled"
                    'write to error log
                    msg = Now & " Process sub startProcess chkPromptOverwrite = True" & _
                                " - User cancelled process"
                    Call writeErrorLog(strErrLog, msg)
                    Call addReport("")
                    Call addReport("Batch Process was cancelled by user.")
                    Call exitProcess
                    Call enableControls
                    Exit Sub
                End If
            End If
        End If '---------------------------------------------------------------------------------
        
        'initialise GDI+
        token = InitGDIPlus
        
        'now we can save the image --------------------------------------------------------------
        If SavePictureFromHDC(m.picSource.image, strFilename) = False Then
            If chkPromptErrors.Value = 1 Then
                'show message
                msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Err.Description & Chr(13) & _
                      "Do you want to continue Batch Process?"
                retVal = MsgBox(msg, vbYesNo + vbExclamation, "Photo Logo Plus")
                'when too many errors occur, we must offer the user a way out
                If retVal = vbNo Then
                    lblStatus.Caption = "Processed: " & n + 1 & " Images to " & _
                    File2.Path & " - Process was cancelled"
                    'write to error log
                    msg = Now & " Process sub startProcess - User cancelled process " & strFilename
                    Call writeErrorLog(strErrLog, msg)
                    Call addReport("")
                    Call addReport("Batch Process was cancelled by user.")
                    'free GDI+
                    FreeGDIPlus token
                    Call exitProcess
                    Call enableControls
                    Exit Sub
                End If
            End If
            
        End If '---------------------------------------------------------------------------------
        
        'free GDI+
        FreeGDIPlus token
        
        File2.Refresh
        
Skip:
    
        'increment serial
        If blnTextLogo = True Then
            If m.chkAutoIncrement.Value = 1 Then
                valSerial = Val(m.txtAdd(4).Text) + 1
                'set format
                If m.optFormat(0) = True Then strDigits = "000"
                If m.optFormat(1) = True Then strDigits = "0000"
                If m.optFormat(2) = True Then strDigits = "00000"
                m.txtAdd(4).Text = Format(valSerial, strDigits)
            End If
        End If
        
    Next n
    
    'update label
    lblStatus.Caption = "Processed: " & File3.ListCount & " Images to Folder " & _
                        File2.Path
    m.lblFileName.Caption = lblStatus.Caption
    
    Call exitProcess
    Call enableControls
    
ErrHandler:
    
    'normal error routine -----------------------------------------------------------------------
    If Err.Number <> 0 Then
        
        errCount = errCount + 1
        
        'write to error log
        msg = Now & " Process sub startProcess - Error " & Err.Number & ": " & _
                Err.Description & " " & fixPath(File3.Path, File3.FileName) & " - " & strExtention
        Call writeErrorLog(strErrLog, msg)
        
        'add error to report
        msg = "Could not process File: " & fixPath(File3.Path, File3.FileName)
        Call addReport(msg)
        
        If chkPromptErrors.Value = 1 Then
            'show message
            msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Err.Description & Chr(13) & _
                  "Do you want to continue Batch Process?"
            retVal = MsgBox(msg, vbYesNo + vbExclamation, "Photo Logo Plus")
            'when too many errors occur, we must offer the user a way out
            If retVal = vbYes Then
                'we continue
                Resume Next
            ElseIf retVal = vbNo Then
                'we leave
                lblStatus.Caption = "Processed: " & n + 1 & " Images to " & _
                    File2.Path & " - Process was cancelled"
                'write to error log
                msg = Now & " Process sub startProcess - User cancelled conversion " & strFilename
                Call writeErrorLog(strErrLog, msg)
                Call addReport("")
                Call addReport("Batch Process was cancelled by user.")
                Call exitProcess
                Call enableControls
                Exit Sub
            End If
        End If
    
        Resume Next
    
    End If
    
End Sub

Private Sub exitProcess()
    
    Dim n As Integer
    
    'show report
    If errCount <> 0 Then
        Call addReport("")
        Call addReport(Format(Now, "dd-mm-yyyy - hh:mm:ss") & _
                    "  Successfully processed: " & List1.ListCount & " files")
        'add processed filenames to report
        If List1.ListCount <> 0 Then
            Call addReport("")
            For n = 0 To List1.ListCount - 1
                List1.ListIndex = n
                Call addReport(List1.Text)
            Next n
        End If
        frmReport.Show
    End If
    
    'delete our temporarily error report if no errors occurred
    If errCount = 0 Then Call SafeKill(strReport)
    errCount = 0
    
    'show thumbnail viewport and logo indicator main form
    m.thumbViewPort.Visible = True
    m.shpLogoPos.Visible = True
    
    'centre source image
    Call m.navSource_MouseDown(4, 0, 0, 0, 0)
    Call m.navSource_MouseUp(4, 0, 0, 0, 0)
    
    'progbar
    Prog.Value = 0
    Prog.Visible = False
    
    'reset flags
    blnCancel = False
    blnProcessing = False
    
End Sub

Private Sub resetSerialNumber()
    
    'reset serial number
    If blnTextLogo = True Then
        If m.chkAutoIncrement.Value = 1 Then
            valSerial = 1
            'set format
            If m.optFormat(0) = True Then strDigits = "000"
            If m.optFormat(1) = True Then strDigits = "0000"
            If m.optFormat(2) = True Then strDigits = "00000"
            m.txtAdd(4).Text = Format(valSerial, strDigits)
        End If
    End If
    
End Sub

Public Sub setPosition(mIndex As Integer)
    
    'set logo position
    
    Dim n As Integer
    
    'reset images
    For n = 0 To 4
        imgPosition(n).Picture = opt_norm.Picture
    Next n
    
    Select Case mIndex
        
        Case 0          'left top
            imgPosition(0).Picture = opt_down.Picture
        Case 1          'right top
            imgPosition(1).Picture = opt_down.Picture
        Case 2          'centre image
            imgPosition(2).Picture = opt_down.Picture
        Case 3          'left bottom
            imgPosition(3).Picture = opt_down.Picture
        Case 4          'right bottom
            imgPosition(4).Picture = opt_down.Picture
        
    End Select
    
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
    
    'no exit if still processing
    If blnProcessing = True Then
        msg = "Batch Process is in progress, cancel the          " & Chr(13) & _
              "process before exiting Batch Process. "
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


