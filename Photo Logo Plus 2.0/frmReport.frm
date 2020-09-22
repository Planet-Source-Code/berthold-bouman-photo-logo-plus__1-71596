VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Photo Logo Report"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8325
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   -15
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   5190
      Width           =   8355
      Begin VB.CheckBox chkPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Prompt to save Report"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   225
         TabIndex        =   2
         Top             =   225
         Width           =   2310
      End
      Begin VB.Image imgExit 
         Height          =   375
         Left            =   7275
         Picture         =   "frmReport.frx":08CA
         Top             =   210
         Width           =   960
      End
      Begin VB.Image imgSave 
         Height          =   375
         Left            =   6225
         Picture         =   "frmReport.frx":0D38
         Top             =   210
         Width           =   960
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   300
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtReport 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5190
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   8340
   End
   Begin VB.Image curHand 
      Height          =   480
      Left            =   390
      Picture         =   "frmReport.frx":11DA
      Top             =   6285
      Width           =   480
   End
   Begin VB.Image but_save_norm 
      Height          =   375
      Left            =   2310
      Picture         =   "frmReport.frx":132C
      Top             =   6435
      Width           =   960
   End
   Begin VB.Image but_save_down 
      Height          =   375
      Left            =   2310
      Picture         =   "frmReport.frx":17CE
      Top             =   6855
      Width           =   960
   End
   Begin VB.Image but_exit_norm 
      Height          =   375
      Left            =   1245
      Picture         =   "frmReport.frx":1C5B
      Top             =   6435
      Width           =   960
   End
   Begin VB.Image but_exit_down 
      Height          =   375
      Left            =   1245
      Picture         =   "frmReport.frx":20C9
      Top             =   6855
      Width           =   960
   End
End
Attribute VB_Name = "frmReport"
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

Private blnSaved        As Boolean      'flags we saved the report or not
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub Form_Load()
    
    On Error Resume Next
    
    'set colors
    txtReport.BackColor = RGB(203, 218, 237)
    chkPrompt.BackColor = RGB(203, 218, 237)
    picBar.BackColor = RGB(203, 218, 237)
    
    'cursors
    chkPrompt.MousePointer = 99
    chkPrompt.MouseIcon = curHand
    imgSave.MousePointer = 99
    imgSave.MouseIcon = curHand
    imgExit.MousePointer = 99
    imgExit.MouseIcon = curHand
    
    'retrieve prompt to save flag
    chkPrompt.Value = QueryValue(HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", "reportPrompt")
    
    'If doPrompt = 1 Then chkPrompt.Value = 1
    'If doPrompt = 0 Then chkPrompt.Value = 0
    
    blnSaved = False
    
    Call showReport
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    Dim msg         As String       'messagebox
    Dim retVal                      'messagebox
    
    'do/don't save report
    If blnSaved = False And chkPrompt.Value = 1 Then
        msg = "Photo Logo Report was not saved,      " & Chr(13) & _
              "do you want save the report?"
        retVal = MsgBox(msg, vbInformation + vbYesNo, "Photo Logo Plus")
        If retVal = vbYes Then Call saveReport
    End If
    
    'delete temp file
    If FileExists(strReport) = True Then
        Kill strReport
    End If
    
End Sub

Private Sub showReport()
    
    On Error Resume Next
    
    Dim FF As Integer
    FF = FreeFile
    
    txtReport.Text = ""
    
    Open strReport For Input As #FF
        Do Until EOF(FF)
            txtReport.Text = Input(LOF(FF), FF)
        Loop
    Close #FF
  
End Sub

Private Sub saveReport()

    'save report
    On Error GoTo ErrHandler
    
    Dialog.CancelError = True
    Dialog.InitDir = App.Path & "\Reports"
    Dialog.DialogTitle = "Save Report"
    Dialog.flags = cdlOFNOverwritePrompt
    Dialog.Filter = "Windows Text File (*.txt)|*.txt|"
    Dialog.FileName = "*.txt"
    Dialog.ShowSave
    
    'keep it simple
    FileCopy strReport, Dialog.FileName
    blnSaved = True
    
ErrHandler:

End Sub

Private Sub chkPrompt_Click()
    
    On Error Resume Next
    
    'write to register - first we create a key - in case there isn't one,
    'e.g. when the app runs for the first time or when register is messed up
    CreateNewKey HKEY_CURRENT_USER, "Software\PhotoLogo20\Main"
    
    'now we can write the value to the register
    SetKeyValue HKEY_CURRENT_USER, "Software\PhotoLogo20\Main", _
                                   "reportPrompt", chkPrompt.Value, REG_SZ
        
End Sub

Private Sub imgSave_Click()
    
    'save report
    Call saveReport
    
End Sub

Private Sub imgSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'exit batch conversion
    imgSave.Picture = but_save_down.Picture
    
End Sub

Private Sub imgSave_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'exit batch conversion
    imgSave.Picture = but_save_norm.Picture
    
End Sub

Private Sub imgExit_Click()
    
    'exit report
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



