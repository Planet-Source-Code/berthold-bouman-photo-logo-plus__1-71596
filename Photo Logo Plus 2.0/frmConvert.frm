VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConvert 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Batch Convert"
   ClientHeight    =   9465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   631
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   3480
      TabIndex        =   46
      Top             =   9570
      Width           =   1410
   End
   Begin VB.PictureBox picBot 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      Picture         =   "frmConvert.frx":08CA
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   3
      Top             =   8940
      Width           =   6465
      Begin MSComctlLib.ProgressBar Prog 
         Height          =   165
         Left            =   165
         TabIndex        =   4
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
      Left            =   3870
      TabIndex        =   20
      Top             =   10095
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.PictureBox picSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Left            =   3870
      ScaleHeight     =   72
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   25
      Top             =   10485
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox picConvertFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   2265
      ScaleHeight     =   136
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   12
      Top             =   6480
      Width           =   4065
      Begin VB.PictureBox picConvertTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2070
         Left            =   1845
         ScaleHeight     =   138
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   147
         TabIndex        =   27
         Top             =   0
         Width           =   2205
         Begin VB.CheckBox chkPromptOverwrite 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   180
            TabIndex        =   33
            Top             =   1695
            Width           =   195
         End
         Begin VB.OptionButton optFileType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   165
            TabIndex        =   39
            Top             =   1395
            Width           =   195
         End
         Begin VB.OptionButton optFileType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   165
            TabIndex        =   38
            Top             =   1095
            Width           =   195
         End
         Begin VB.OptionButton optFileType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   165
            TabIndex        =   37
            Top             =   795
            Width           =   195
         End
         Begin VB.OptionButton optFileType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   36
            Top             =   495
            Width           =   195
         End
         Begin VB.OptionButton optFileType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   35
            Top             =   195
            Width           =   195
         End
         Begin VB.Label lblPrompt 
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
            Height          =   285
            Left            =   450
            TabIndex        =   34
            Top             =   1695
            Width           =   1695
            WordWrap        =   -1  'True
         End
         Begin VB.Image imgStart 
            Height          =   375
            Left            =   1140
            Picture         =   "frmConvert.frx":BA3C
            Top             =   105
            Width           =   960
         End
         Begin VB.Image imgCancel 
            Height          =   375
            Left            =   1140
            Picture         =   "frmConvert.frx":BECA
            Top             =   585
            Width           =   960
         End
         Begin VB.Image imgExit 
            Height          =   375
            Left            =   1140
            Picture         =   "frmConvert.frx":C390
            Top             =   1065
            Width           =   960
         End
         Begin VB.Label lbl2Bmp 
            BackStyle       =   0  'Transparent
            Caption         =   "BMP"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   465
            TabIndex        =   32
            Top             =   180
            Width           =   450
         End
         Begin VB.Label lbl2JPG 
            BackStyle       =   0  'Transparent
            Caption         =   "JPG"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   31
            Top             =   465
            Width           =   450
         End
         Begin VB.Label lbl2GIF 
            BackStyle       =   0  'Transparent
            Caption         =   "GIF"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   480
            TabIndex        =   30
            Top             =   765
            Width           =   255
         End
         Begin VB.Label lbl2PNG 
            BackStyle       =   0  'Transparent
            Caption         =   "PNG"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   29
            Top             =   1065
            Width           =   450
         End
         Begin VB.Label lbl2TIF 
            BackStyle       =   0  'Transparent
            Caption         =   "TIF"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   28
            Top             =   1380
            Width           =   450
         End
      End
      Begin VB.OptionButton optConvert 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "All2BMP"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   21
         Top             =   195
         Width           =   195
      End
      Begin VB.OptionButton optConvert 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "All2JPG"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   495
         Width           =   195
      End
      Begin VB.OptionButton optConvert 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "BMP2JPG"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   23
         Top             =   795
         Width           =   195
      End
      Begin VB.OptionButton optConvert 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "JPG2BMP"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   24
         Top             =   1095
         Width           =   195
      End
      Begin VB.OptionButton optConvert 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "JPG2BMP"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   42
         Top             =   1395
         Width           =   195
      End
      Begin VB.OptionButton optConvert 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "JPG2BMP"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   43
         Top             =   1695
         Width           =   195
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TIF"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   45
         Top             =   1665
         Width           =   450
      End
      Begin VB.Label lblPNG 
         BackStyle       =   0  'Transparent
         Caption         =   "PNG"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   44
         Top             =   1365
         Width           =   450
      End
      Begin VB.Image imgArrow 
         Height          =   555
         Left            =   1065
         Picture         =   "frmConvert.frx":C7FE
         Top             =   630
         Width           =   660
      End
      Begin VB.Label lblAll 
         BackStyle       =   0  'Transparent
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   16
         Top             =   165
         Width           =   450
      End
      Begin VB.Label lblJPG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "JPG"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   480
         TabIndex        =   15
         Top             =   765
         Width           =   450
      End
      Begin VB.Label lblGIF 
         BackStyle       =   0  'Transparent
         Caption         =   "GIF"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   14
         Top             =   1065
         Width           =   450
      End
      Begin VB.Label lblBMP 
         BackStyle       =   0  'Transparent
         Caption         =   "BMP"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   13
         Top             =   465
         Width           =   450
      End
   End
   Begin VB.PictureBox picThumbFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2070
      Left            =   135
      Picture         =   "frmConvert.frx":DB54
      ScaleHeight     =   138
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   138
      TabIndex        =   11
      Top             =   6480
      Width           =   2070
      Begin VB.PictureBox picThumb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   135
         ScaleHeight     =   54
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   54
         TabIndex        =   26
         Top             =   135
         Width           =   810
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
      Picture         =   "frmConvert.frx":1BBD6
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   421
      TabIndex        =   9
      Top             =   5955
      Width           =   6315
      Begin VB.CheckBox chkPromptErrors 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Check1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4485
         TabIndex        =   40
         Top             =   150
         Width           =   195
      End
      Begin VB.CheckBox chkThumb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Thumbnail"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   150
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Label lblErrors 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4770
         TabIndex        =   41
         Top             =   135
         Width           =   1410
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblShowThumb 
         BackStyle       =   0  'Transparent
         Caption         =   "Show Thumbnail"
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
         Height          =   270
         Left            =   315
         TabIndex        =   18
         Top             =   135
         Width           =   1440
      End
      Begin VB.Label lblFileFormat 
         BackStyle       =   0  'Transparent
         Caption         =   "Select File Conversion:"
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
         Left            =   2280
         TabIndex        =   10
         Top             =   135
         Width           =   2025
      End
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
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   900
      Width           =   2475
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
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   900
      Width           =   2475
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
      TabIndex        =   6
      Top             =   1275
      Width           =   3060
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
      TabIndex        =   5
      Top             =   1275
      Width           =   3060
   End
   Begin VB.PictureBox PicTop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   0
      Picture         =   "frmConvert.frx":25528
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   0
      Top             =   0
      Width           =   6465
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
         TabIndex        =   2
         Top             =   510
         Width           =   2685
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
         TabIndex        =   1
         Top             =   510
         Width           =   2685
      End
      Begin VB.Image imgMin 
         Height          =   285
         Left            =   5700
         Picture         =   "frmConvert.frx":370EA
         Top             =   105
         Width           =   285
      End
      Begin VB.Image imgClose 
         Height          =   285
         Left            =   6045
         Picture         =   "frmConvert.frx":375A0
         Top             =   105
         Width           =   285
      End
   End
   Begin VB.Image but_exit_down 
      Height          =   375
      Left            =   2370
      Picture         =   "frmConvert.frx":37A56
      Top             =   11220
      Width           =   960
   End
   Begin VB.Image but_exit_norm 
      Height          =   375
      Left            =   2370
      Picture         =   "frmConvert.frx":37EA7
      Top             =   10800
      Width           =   960
   End
   Begin VB.Image but_cancel_down 
      Height          =   375
      Left            =   1365
      Picture         =   "frmConvert.frx":38315
      Top             =   11220
      Width           =   960
   End
   Begin VB.Image but_start_down 
      Height          =   375
      Left            =   360
      Picture         =   "frmConvert.frx":387B8
      Top             =   11220
      Width           =   960
   End
   Begin VB.Image but_cancel_norm 
      Height          =   375
      Left            =   1365
      Picture         =   "frmConvert.frx":38C2C
      Top             =   10800
      Width           =   960
   End
   Begin VB.Image but_start_norm 
      Height          =   375
      Left            =   360
      Picture         =   "frmConvert.frx":390F2
      Top             =   10800
      Width           =   960
   End
   Begin VB.Image imgOpenOutput 
      Height          =   315
      Left            =   5790
      Picture         =   "frmConvert.frx":39580
      Top             =   900
      Width           =   540
   End
   Begin VB.Image imgOpenSource 
      Height          =   315
      Left            =   2655
      Picture         =   "frmConvert.frx":39E9E
      Top             =   900
      Width           =   540
   End
   Begin VB.Image folder_down 
      Height          =   315
      Left            =   2250
      Picture         =   "frmConvert.frx":3A7BC
      Top             =   10305
      Width           =   540
   End
   Begin VB.Image folder_norm 
      Height          =   315
      Left            =   2250
      Picture         =   "frmConvert.frx":3ABC6
      Top             =   9960
      Width           =   540
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
      Height          =   330
      Left            =   150
      TabIndex        =   19
      Top             =   8625
      Width           =   6165
   End
   Begin VB.Image imgRight 
      Height          =   8175
      Left            =   6390
      Picture         =   "frmConvert.frx":3AFDC
      Stretch         =   -1  'True
      Top             =   795
      Width           =   75
   End
   Begin VB.Image min_norm 
      Height          =   285
      Left            =   330
      Picture         =   "frmConvert.frx":3B22E
      Top             =   9960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_hot 
      Height          =   285
      Left            =   690
      Picture         =   "frmConvert.frx":3B6E4
      Top             =   9960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_down 
      Height          =   285
      Left            =   1035
      Picture         =   "frmConvert.frx":3BB9A
      Top             =   9960
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_norm 
      Height          =   285
      Left            =   330
      Picture         =   "frmConvert.frx":3C050
      Top             =   10335
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_hot 
      Height          =   285
      Left            =   690
      Picture         =   "frmConvert.frx":3C506
      Top             =   10335
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_down 
      Height          =   285
      Left            =   1035
      Picture         =   "frmConvert.frx":3C9BC
      Top             =   10335
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image curHand 
      Height          =   480
      Left            =   1830
      Picture         =   "frmConvert.frx":3CE72
      Top             =   10230
      Width           =   480
   End
   Begin VB.Image imgLeft 
      Height          =   8205
      Left            =   0
      Picture         =   "frmConvert.frx":3CFC4
      Stretch         =   -1  'True
      Top             =   825
      Width           =   75
   End
End
Attribute VB_Name = "frmConvert"
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

Private blnCancel           As Boolean          'cancel conversion
Private strExtention        As String           'file extention we convert to
Private strConvertFrom      As String           'contains which files we convert
Private errCount            As Integer          'does what it says

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Form_Load()
    
    On Error GoTo ErrHandler
    
    Dim msg         As String       'message box
    Dim mTop        As Single       'form top position
    Dim mLeft       As Single       'form left position
    
    'retrieve top and left coordinates from register
    mTop = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert", "convertTop")
    mLeft = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert", "convertLeft")
    
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
    
    Call setControls
    
    'set default thumbnail image in case if File1 is not populated
    If File1.ListCount = 0 Then
        picSource.Picture = frmMain.picSource.Picture
        Call makeThumb
    End If
    
    'set default conversion to ALL
    optConvert(0) = True
    Call optConvert_Click(0)
    
    'set default file type to BMP
    optFileType(0) = True
    Call optFileType_Click(0)
    
    'load first thumbnail if File1 is populated
    If File1.ListCount > 0 Then
        File1.ListIndex = 0
        Call File1_Click
    End If
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Convert Form_Load - Error " & Err.Number & ": " & Err.Description
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    'write to register - first we create a key - in case there isn't one,
    'e.g. when the app runs for the first time
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert"
    
    'now we can write the values to the register
    'save top and left position to register
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert", _
                                   "convertTop", Me.Top, REG_SZ
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert", _
                                   "convertLeft", Me.Left, REG_SZ
                                   
End Sub

'+++++++++++++++++++++++++++++++++++++++++ CONTROL EVENTS +++++++++++++++++++++++++++++++++++++++

Private Sub imgOpenSource_Click()
    
    On Error GoTo ErrHandler
    
    'open source directory
    Dim strResFolder    As String       'open folder
    Dim strOldFolder    As String       'old folder path
    Dim msg             As String       'message box
    
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
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert"
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert", _
                                           "Last Source Folder", txtSource.Text, REG_SZ
    'reset conversion option
    optConvert(0).Value = 1
    
    'update label
    lblStatus.Caption = File3.ListCount & " Images in Source Directory"
    
    'load first thumbnail
    If File1.ListCount <> 0 Then
        File1.ListIndex = 0
        Call File1_Click
    End If
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Convert imgOpenSource_Click - Error " & Err.Number & ": " & _
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
    
    On Error GoTo ErrHandler
    
    'open output directory
    Dim strResFolder    As String       'open folder
    Dim strOldFolder    As String       'old folder path
    Dim msg             As String       'message box
    
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
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert"
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert", _
                                           "Last Output Folder", txtOutput.Text, REG_SZ
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Convert imgOpenOutput_Click - Error " & Err.Number & ": " & _
              Err.Description & " " & strResFolder
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Private Sub imgOpenOutput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'open output directory
    imgOpenOutput.Picture = folder_down.Picture
    
End Sub

Private Sub imgOpenOutput_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'open source directory
    imgOpenOutput.Picture = folder_norm.Picture
    
End Sub

Private Sub imgStart_Click()
    
    'start conversion
    
    Dim msg     As String       'message box
    
    'first check if there are files to process
    If File1.ListCount = 0 Then
        msg = "There are no files to process. Select another       " & Chr(13) & _
              "directory or select another conversion option.      "
        MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
        Exit Sub
    End If
    
    'set flags
    blnConverting = True
    blnCancel = False
    
    'clear report listbox
    List1.Clear
    
    Call startConversion
    
End Sub

Private Sub imgStart_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'start conversion
    imgStart.Picture = but_start_down.Picture
    
End Sub

Private Sub imgStart_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'start conversion
    imgStart.Picture = but_start_norm.Picture
    
End Sub

Private Sub imgCancel_Click()
    
    'cancel conversion
    
    Dim msg     As String       'message box
    Dim retVal                  'message box
    
    'do or don't cancel conversion
    If blnConverting = True Then
        msg = "Are you sure you want to cancel Batch Conversion?        "
        retVal = MsgBox(msg, vbExclamation + vbYesNo, "Photo Logo Plus")
        If retVal = vbNo Then
            Exit Sub
        ElseIf retVal = vbYes Then
            'write to error log
            msg = Now & " Convert imgCancel_Click - User cancelled conversion"
            Call writeErrorLog(strErrLog, msg)
            blnConverting = False
            blnCancel = True
            Exit Sub
        End If
    End If
    
End Sub

Private Sub imgCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'start conversion
    imgCancel.Picture = but_cancel_down.Picture
    
End Sub

Private Sub imgCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'start conversion
    imgCancel.Picture = but_cancel_norm.Picture
    
End Sub

Private Sub imgExit_Click()
    
    'exit batch conversion
    
    Dim msg     As String       'message box
    
    'no exit if still converting
    If blnConverting = True Then
        msg = "File Conversion is in progress, cancel the file         " & Chr(13) & _
              "conversion first before exiting Batch Convertion. "
        MsgBox msg, vbExclamation + vbOKOnly, "Photo Logo Plus"
        Exit Sub
    End If
    
    Unload Me
    
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'exit batch conversion
    imgExit.Picture = but_exit_down.Picture
    
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'exit batch conversion
    imgExit.Picture = but_exit_norm.Picture
    
End Sub

Private Sub File1_Click()
    
    On Error Resume Next
    
    'no other function then showing the thumbnail of the selected file
    If chkThumb.Value = 1 Then
        strThumbName = fixPath(File1.Path, File1.FileName)
        Call makeThumb
    End If
    
End Sub

Private Sub File2_Click()
    
    On Error Resume Next
    
    'no other function then showing the thumbnail of the selected file
    If chkThumb.Value = 1 Then
        strThumbName = fixPath(File2.Path, File2.FileName)
        Call makeThumb
    End If
    
End Sub

Private Sub optConvert_Click(Index As Integer)
    
    'set convert option
    
    If blnConverting = True Then Exit Sub
    
    Select Case Index
    
        Case 0          'ALL files
            
            File1.Pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.png;*.tif;*,tiff"
            File3.Pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.png;*.tif;*.tiff"
            strConvertFrom = "ALL"
            
        Case 1          'BMP only
            
            File1.Pattern = "*.bmp"
            File3.Pattern = "*.bmp"
            strConvertFrom = "BMP"
            
        Case 2          'JPG only
            
            File1.Pattern = "*.jpg;*.jpeg"
            File3.Pattern = "*.jpg;*.jpeg"
            strConvertFrom = "JPG"
            
        Case 3          'GIF only
            
            File1.Pattern = "*.gif"
            File3.Pattern = "*.gif"
            strConvertFrom = "GIF"
        
        Case 4          'PNG only
            
            File1.Pattern = "*.png"
            File3.Pattern = "*.png"
            strConvertFrom = "PNG"
            
        Case 5          'TIF only
            
            File1.Pattern = "*.tif"
            File3.Pattern = "*.tif"
            strConvertFrom = "TIF"
            
    End Select
    
    'load first thumbnail if File1 is populated
    If File1.ListCount > 0 Then
        File1.ListIndex = 0
        Call File1_Click
        'update label
        lblStatus.Caption = File3.ListCount & " Images in Source Directory"
    End If
    
    If File1.ListCount = 0 Then
        'update label
        lblStatus.Caption = "No files to convert in this Directory"
    End If
    
End Sub

Private Sub optFileType_Click(Index As Integer)
        
    'set file format to convert to
    If blnConverting = True Then Exit Sub
    
    Select Case Index
        
        Case 0          'convert to BMP
            
            strExtention = "bmp"
        
        Case 1          'convert to JPG
            
            strExtention = "jpg"
            
        Case 2          'convert to GIF
            
            strExtention = "gif"
            
        Case 3          'convert to PNG
            
            strExtention = "png"
            
        Case 4          'convert to TIF
            
            strExtention = "tif"
            
    End Select
            
End Sub

'++++++++++++++++++++++++++++++++++++++++++++ COMMON SUBS +++++++++++++++++++++++++++++++++++++++

Private Sub setControls()
    
    On Error GoTo ErrHandler
    
    Dim msg         As String       'message box
    Dim n           As Integer      'counter
    Dim strFolder   As String       'folder name
    
    'colors
    For n = 0 To Me.Height Step 3
        'paint background
        Me.Line (0, n)-(Me.Width, n), RGB(190, 212, 255)
        Me.Line (0, n + 1)-(Me.Width, n + 1), RGB(204, 224, 255)
        Me.Line (0, n + 2)-(Me.Width, n + 2), RGB(255, 255, 255)
    Next n
    
    'cursors
    imgMin.MousePointer = 99
    imgMin.MouseIcon = curHand
    imgClose.MousePointer = 99
    imgClose.MouseIcon = curHand
    
    For n = 0 To 5
        optConvert(n).MousePointer = 99
        optConvert(n).MouseIcon = curHand
    Next n
    
    For n = 0 To 4
        optFileType(n).MousePointer = 99
        optFileType(n).MouseIcon = curHand
    Next n
    
    chkPromptOverwrite.MousePointer = 99
    chkPromptOverwrite.MouseIcon = curHand
    
    chkPromptErrors.MousePointer = 99
    chkPromptErrors.MouseIcon = curHand
    
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
    
    'file boxes
    File1.Pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.wmf;*.png;*.tif;*.tiff"
    File2.Pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.wmf;*.png;*.tif;*.tiff"
    File3.Pattern = "*.jpg;*.jpeg;*.bmp;*.gif;*.wmf;*.png;*.tif;*.tiff"
    
    'retrieve last used folders from register
    strFolder = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert", "Last Source Folder")
    If strFolder = "" Then strFolder = App.Path
    File1.Path = strFolder
    File3.Path = strFolder
    txtSource.Text = File1.Path
    lblStatus.Caption = File3.ListCount & " Images in Source Directory"
    
    strFolder = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Convert", "Last Output Folder")
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
        msg = Now & " Convert Sub SetControls - Error " & Err.Number & ": " & Err.Description
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Resume Next
                
    End If
    
End Sub

Public Sub makeThumb()
    
    On Error GoTo ErrHandler
    
    'make thumbnail
    Dim token           As Long         'GDI+
    Dim sideHor         As Integer
    Dim sideVer         As Integer
    Dim maxWidth        As Integer
    Dim maxHeight       As Integer
    Dim iRatio          As Double
    Dim msg             As String       'messagebox
    Dim submsg          As String       'messagebox
    
    'set max thumbnail size
    maxWidth = 120
    maxHeight = 120
    
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
    
    'center the thumbnail
    picThumb.Left = (picThumbFrame.Width - picThumb.Width) / 2
    picThumb.Top = (picThumbFrame.Height - picThumb.Height) / 2
    
ErrHandler:
    
    If Err.Number <> 0 Then
        
        'write to error log
        msg = Now & " Convert sub makeThumb - Error " & Err.Number & ": " & _
              Err.Description & " " & strThumbName & " - " & submsg
        Call writeErrorLog(strErrLog, msg)
        
        'show message
        msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Chr(13) & Err.Description
        MsgBox msg, vbOKOnly + vbExclamation, "Photo Logo Plus"
        
        Exit Sub
                
    End If
        
End Sub

Private Sub disableControls()
    
    'disable controls while converting
    
    Dim n As Integer
    
    imgStart.Enabled = False
    imgOpenSource.Enabled = False
    imgOpenOutput.Enabled = False
    
    For n = 0 To 5
        optConvert(n).MousePointer = 12
    Next n
    
    For n = 0 To 4
        optFileType(n).MousePointer = 12
    Next n
    
End Sub

Private Sub enableControls()
    
    'disable controls while converting
    
    Dim n As Integer
    
    imgStart.Enabled = True
    imgOpenSource.Enabled = True
    imgOpenOutput.Enabled = True
    
    For n = 0 To 5
        optConvert(n).MousePointer = 99
    Next n
    
    For n = 0 To 4
        optFileType(n).MousePointer = 99
    Next n
    
End Sub

Private Sub startConversion()

    On Error GoTo ErrHandler
    
    'convert images
    Dim token           As Long         'GDI+
    
    Dim n               As Integer      'counter
    Dim strFilename     As String       'filename string, without extention
    Dim strOldExt       As String       'old file extention
    Dim msg             As String       'message box
    Dim retVal                          'message box
    
    'start conversion report
    strReport = App.Path & "\Reports\" & Format(Now, "ddmmyyhhmmss") & ".tmp"
    Call startReport(File1.Path, File2.Path, strConvertFrom, UCase(strExtention))
    
    errCount = 0
    
    Call disableControls
    
    'set progbar
    Prog.Visible = True
    Prog.Max = File3.ListCount
    Prog.Value = 0
        
    For n = 0 To File3.ListCount - 1
        
        'we leave if conversion is cancelled ----------------------------------------------------
        If blnCancel = True Then
            lblStatus.Caption = "Converted: " & n + 1 & " Images to " & _
                        File2.Path & " - Conversion was cancelled"
            Call addReport("")
            Call addReport("Batch Conversion was cancelled by user.")
            Call exitConversion
            Call enableControls
            Exit Sub
        End If '---------------------------------------------------------------------------------
        
        File3.ListIndex = n
        
        'load new source image
        token = InitGDIPlus
        picSource.Picture = LoadPictureGDIPlus(fixPath(File3.Path, File3.FileName))
        'free GDI+
        FreeGDIPlus token
        
        'thumbnail
        strThumbName = fixPath(File3.Path, File3.FileName)
        If chkThumb.Value = 1 Then Call makeThumb
        
        Prog.Value = n
        
        'update status label
        If Len("Connverting: " & n + 1 & "/" & File3.ListCount & _
                            " - " & fixPath(File3.Path, File3.FileName)) >= 62 Then
            'get rid of excessive long path names
            lblStatus.Caption = Mid("Converting: " & n + 1 & "/" & File3.ListCount & _
                            " - " & fixPath(File3.Path, File3.FileName), 1, 62) & "..."
        Else
            lblStatus.Caption = "Converting: " & n + 1 & "/" & File3.ListCount & _
                            " - " & fixPath(File3.Path, File3.FileName)
        End If
        
        DoEvents
            
        'create filename string - get file extention first, we need
        'to know the length of the file extention to get rid of it
        strOldExt = GetFileExtention(File3.FileName)
        'now we remove the old extention and add the new extention
        strFilename = fixPath(File2.Path, _
                    Mid(File3.FileName, 1, Len(File3.FileName) - _
                    Len(strOldExt)) & strExtention)
            
        'check if file aleady exists ------------------------------------------------------------
        If chkPromptOverwrite.Value = 1 Then
            If FileExists(strFilename) = True Then
                msg = "Filename already exists, do you want to          " & Chr(13) & _
                      "replace the existing file with the new file?" & Chr(13) & Chr(13) & _
                      "Select Cancel to stop Batch Convert."
                retVal = MsgBox(msg, vbExclamation + vbYesNoCancel, "Photo Logo Plus")
                'don't replace
                If retVal = vbNo Then GoTo Skip
                'cancel the conversion - our way out
                If retVal = vbCancel Then
                    lblStatus.Caption = "Converted: " & n + 1 & " Images to " & _
                    File2.Path & " - Conversion was cancelled"
                    'write to error log
                    msg = Now & " Convert sub startConversion chkPromptOverwrite = True" & _
                                "- User cancelled conversion"
                    Call writeErrorLog(strErrLog, msg)
                    Call addReport("")
                    Call addReport("Batch Conversion was cancelled by user.")
                    Call exitConversion
                    Call enableControls
                    Exit Sub
                End If
            End If
        End If '---------------------------------------------------------------------------------
        
        'initialise GDI+
        token = InitGDIPlus
        
        'now we can save the image --------------------------------------------------------------
        If SavePictureFromHDC(picSource.Picture, strFilename) = False Then
            'save file info
            msg = "Could not convert File: " & fixPath(File3.Path, File3.FileName)
            Call addReport(msg)
            errCount = errCount + 1
            msg = Now & " Convert sub startConversion GDI+ Save Picture Error" & " - " & fixPath(File3.Path, _
                        File3.FileName) & " - Convert to: " & strExtention
            Call writeErrorLog(strErrLog, msg)
            If chkPromptErrors.Value = 1 Then
                'show message
                msg = "Photo Logo Plus - GDI+ Error - Filename: " & Chr(13) _
                      & fixPath(File3.Path, File3.FileName) & Chr(13) & Chr(13) & _
                      "Do you want to continue Batch Conversion?        "
                retVal = MsgBox(msg, vbYesNo + vbExclamation, "Photo Logo Plus")
                'when too many errors occur, we must offer the user a way out
                If retVal = vbNo Then
                    lblStatus.Caption = "Converted: " & n + 1 & " Images to " & _
                    File2.Path & " - Conversion was cancelled"
                    'write to error log
                    msg = Now & " Convert sub startConversion - User cancelled conversion " & strFilename
                    Call writeErrorLog(strErrLog, msg)
                    Call addReport("")
                    Call addReport("Batch Conversion was cancelled by user.")
                    'free GDI+
                    FreeGDIPlus token
                    Call exitConversion
                    Call enableControls
                    Exit Sub
                End If
            End If
        Else
            List1.AddItem "Converted: " & File3.FileName
        End If '---------------------------------------------------------------------------------
        
        'free GDI+
        FreeGDIPlus token
        
        File2.Refresh
Skip:
        
    Next n
    
    'update label
    lblStatus.Caption = "Converted: " & File3.ListCount & " Images to Folder " & _
                        File2.Path
   
    Call exitConversion
    Call enableControls
    
ErrHandler:
    
    'normal error routine -----------------------------------------------------------------------
    If Err.Number <> 0 Then
        
        errCount = errCount + 1
        
        'write to error log
        msg = Now & " Convert sub startConversion - Error " & Err.Number & ": " & _
                Err.Description & " " & fixPath(File3.Path, File3.FileName) & " - " & strExtention
        Call writeErrorLog(strErrLog, msg)
        
        'add error to report
        msg = "Could not convert File: " & fixPath(File3.Path, File3.FileName)
        Call addReport(msg)
        
        If chkPromptErrors.Value = 1 Then
            'show message
            msg = "Photo Logo Plus - Error: " & Err.Number & ":" & Err.Description & Chr(13) & _
                  "Do you want to continue Batch Conversion?"
            retVal = MsgBox(msg, vbYesNo + vbExclamation, "Photo Logo Plus")
            'when too many errors occur, we must offer the user a way out
            If retVal = vbYes Then
                'we continue
                Resume Next
            ElseIf retVal = vbNo Then
                'we leave
                lblStatus.Caption = "Converted: " & n + 1 & " Images to " & _
                    File2.Path & " - Conversion was cancelled"
                'write to error log
                msg = Now & " Convert sub startConversion - User cancelled conversion " & strFilename
                Call writeErrorLog(strErrLog, msg)
                Call addReport("")
                Call addReport("Batch Conversion was cancelled by user.")
                Call enableControls
                Exit Sub
            End If
        End If
    
        Resume Next
    
    End If
        
End Sub

Private Sub exitConversion()
    
    Dim n As Integer
    
    'show report
    If errCount <> 0 Then
        Call addReport("")
        Call addReport(Format(Now, "dd-mm-yyyy - hh:mm:ss") & _
                    "  Successfully converted: " & List1.ListCount & " files")
        'add converted filenames to report
        If List1.ListCount <> 0 Then
            Call addReport("")
            For n = 0 To List1.ListCount - 1
                List1.ListIndex = n
                Call addReport(List1.Text)
            Next n
        End If
        frmReport.Show
    End If
    
    'delete error report if no errors occurred
    If errCount = 0 Then Call SafeKill(strReport)
    errCount = 0
    
    'progbar
    Prog.Value = 0
    Prog.Visible = False
    
     'reset flags
    blnCancel = False
    blnConverting = False
    
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
    
    'no exit if still converting
    If blnConverting = True Then
        msg = "File Conversion is in progress, cancel the file         " & Chr(13) & _
              "conversion first before exiting Batch Convertion. "
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




