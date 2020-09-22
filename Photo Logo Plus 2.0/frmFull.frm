VERSION 5.00
Begin VB.Form frmFull 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Photo Logo Plus: Full Size Image"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3915
   Icon            =   "frmFull.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   156
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   261
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   45
      MousePointer    =   99  'Custom
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   158
      TabIndex        =   0
      ToolTipText     =   " Left Mouse Down to Drag Image, Double-Click to Close Window "
      Top             =   45
      Width           =   2370
   End
   Begin VB.Image curGrab 
      Height          =   480
      Left            =   2970
      Picture         =   "frmFull.frx":08CA
      Top             =   585
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image curRelease 
      Height          =   480
      Left            =   2970
      Picture         =   "frmFull.frx":0BD4
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmFull"
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
    
    Call showImage
        
End Sub

Private Sub picShow_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error Resume Next
    
    If Button = 1 Then
        'move form without caption
        picShow.MouseIcon = curGrab
        ReleaseCapture
        SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
        picShow.MouseIcon = curRelease
        
    End If

End Sub

Private Sub picShow_DblClick()
    
    Unload Me
    
End Sub

Public Sub showImage()
    
    On Error Resume Next
    
    'set controls
    Me.Width = ((m.picSource.ScaleWidth) * Screen.TwipsPerPixelX) + 90
    Me.Height = ((m.picSource.ScaleHeight) * Screen.TwipsPerPixelY) + 345
    picShow.Move 0, 0, m.picSource.ScaleWidth, m.picSource.ScaleHeight
    picShow.MouseIcon = curRelease

    'copy source picture to fullsize picture (for all three cases)
    picShow.Cls
    BitBlt picShow.hDC, _
           0, _
           0, _
           m.picSource.ScaleWidth, _
           m.picSource.ScaleHeight, _
           m.picSource.hDC, _
           0, _
           0, _
           vbSrcCopy
    
    'plain text is already printed on the source image,
    'so we do nothing with text here
    
    'add normal logo
    If blnNormalLogo = True Then
    
        BitBlt picShow.hDC, _
            m.scrLogoHor.Value, _
            m.scrLogoVer.Value, _
            m.picNormalLogo.ScaleWidth, _
            m.picNormalLogo.ScaleHeight, _
            m.picNormalLogo.hDC, _
            0, _
            0, _
            vbSrcCopy
           
    End If
    
    'add masked logo
    If blnMaskedLogo = True Then
    
        TransparentBlt picShow.hDC, _
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

End Sub

