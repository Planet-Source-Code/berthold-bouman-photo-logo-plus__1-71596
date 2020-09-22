VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmHelp 
   BorderStyle     =   0  'None
   Caption         =   "Help"
   ClientHeight    =   11850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   790
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.PictureBox picBot 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      Picture         =   "frmHelp.frx":08CA
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   2
      Top             =   11310
      Width           =   12000
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      Picture         =   "frmHelp.frx":15A8C
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   1
      Top             =   0
      Width           =   12000
      Begin VB.Image imgMin 
         Height          =   285
         Left            =   11250
         Picture         =   "frmHelp.frx":27D6E
         Top             =   105
         Width           =   285
      End
      Begin VB.Image imgClose 
         Height          =   285
         Left            =   11595
         Picture         =   "frmHelp.frx":28224
         Top             =   105
         Width           =   285
      End
   End
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   10905
      Left            =   60
      TabIndex        =   0
      Top             =   465
      Width           =   11865
      ExtentX         =   20929
      ExtentY         =   19235
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Image min_norm 
      Height          =   285
      Left            =   240
      Picture         =   "frmHelp.frx":286DA
      Top             =   12030
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_hot 
      Height          =   285
      Left            =   600
      Picture         =   "frmHelp.frx":28B90
      Top             =   12030
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image min_down 
      Height          =   285
      Left            =   945
      Picture         =   "frmHelp.frx":29046
      Top             =   12030
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_norm 
      Height          =   285
      Left            =   240
      Picture         =   "frmHelp.frx":294FC
      Top             =   12405
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_hot 
      Height          =   285
      Left            =   600
      Picture         =   "frmHelp.frx":299B2
      Top             =   12405
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image close_down 
      Height          =   285
      Left            =   945
      Picture         =   "frmHelp.frx":29E68
      Top             =   12405
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image curHand 
      Height          =   480
      Left            =   1740
      Picture         =   "frmHelp.frx":2A31E
      Top             =   12300
      Width           =   480
   End
   Begin VB.Image imgRight 
      Height          =   11280
      Left            =   11925
      Picture         =   "frmHelp.frx":2A470
      Stretch         =   -1  'True
      Top             =   255
      Width           =   75
   End
   Begin VB.Image imgLeft 
      Height          =   11355
      Left            =   0
      Picture         =   "frmHelp.frx":2A6C2
      Stretch         =   -1  'True
      Top             =   240
      Width           =   75
   End
End
Attribute VB_Name = "frmHelp"
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

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Form_Load()
    
    On Error Resume Next
    
    Dim mTop        As Single       'form top position
    Dim mLeft       As Single       'form left position
    
    'retrieve top and left coordinates from register
    mTop = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Help", "helpTop")
    mLeft = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Help", "helpLeft")
    
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
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    'write to register - first we create a key - in case there isn't one,
    'e.g. when the app runs for the first time
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Help"
    
    'now we can write the values to the register
    'save top and left position to register
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Help", _
                                   "helpTop", Me.Top, REG_SZ
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Help", _
                                   "helpLeft", Me.Left, REG_SZ
                                   
End Sub

'++++++++++++++++++++++++++++++++++++++++++++ COMMON SUBS +++++++++++++++++++++++++++++++++++++++

Private Sub setControls()
    
    On Error Resume Next
    
    'cursors
    imgMin.MousePointer = 99
    imgMin.MouseIcon = curHand
    imgClose.MousePointer = 99
    imgClose.MouseIcon = curHand
    
    Browser.Navigate App.Path & "\Help\index.htm"
    
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
