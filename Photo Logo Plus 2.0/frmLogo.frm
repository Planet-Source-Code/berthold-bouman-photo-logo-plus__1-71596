VERSION 5.00
Begin VB.Form frmLogo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Photo Logo Plus"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3915
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
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
      Left            =   45
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   158
      TabIndex        =   0
      Top             =   45
      Width           =   2370
   End
   Begin VB.Image curRelease 
      Height          =   480
      Left            =   2970
      Picture         =   "frmLogo.frx":08CA
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image curGrab 
      Height          =   480
      Left            =   2970
      Picture         =   "frmLogo.frx":0BD4
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmLogo"
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

Private m As Form

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Form_Load()
    
    On Error Resume Next
    
    Set m = frmMain
    
    'set controls
    If blnNormalLogo = True Then
        
        Me.Width = ((m.picNormalSource.ScaleWidth) * Screen.TwipsPerPixelX) + 90
        Me.Height = ((m.picNormalSource.ScaleHeight) * Screen.TwipsPerPixelY) + 390
        picLogo.Picture = m.picNormalSource
        picLogo.Move 0, 0, m.picNormalSource.ScaleWidth, m.picNormalSource.ScaleHeight
        
    ElseIf blnMaskedLogo = True Then
        
        Me.Width = ((m.picMaskedSource.ScaleWidth) * Screen.TwipsPerPixelX) + 90
        Me.Height = ((m.picMaskedSource.ScaleHeight) * Screen.TwipsPerPixelY) + 390
        picLogo.Picture = m.picMaskedSource
        picLogo.Move 0, 0, m.picMaskedSource.ScaleWidth, m.picMaskedSource.ScaleHeight
        
    End If
       
    'set cursor
    picLogo.MousePointer = 99
    picLogo.MouseIcon = curRelease
    
End Sub

Private Sub picLogo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error Resume Next
    
    If Button = 1 Then
        'move form without caption
        picLogo.MouseIcon = curGrab
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
        picLogo.MouseIcon = curRelease
        
    End If

End Sub

Private Sub picLogo_DblClick()
    
    Unload Me
    
End Sub
