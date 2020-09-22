Attribute VB_Name = "modCommon"
Option Explicit

'logo type flags
Public blnTextLogo          As Boolean      'text logo flag
Public blnNormalLogo        As Boolean      'image logo flag
Public blnMaskedLogo        As Boolean      'masked logo flag

'position flags
Public blnLeftTop           As Boolean      'position of logo
Public blnRightTop          As Boolean      'position of logo
Public blnLeftBot           As Boolean      'position of logo
Public blnRightBot          As Boolean      'position of logo
Public blnCenter            As Boolean      'position of logo

Public blnBatch             As Boolean      'flags if Batch Process active

'batch process and batch convert
Public blnProcessing        As Boolean      'flags we are batch processing files
Public blnConverting        As Boolean      'flags we are batch converting files

'serial number
Public strDigits            As String       'serial number format e.g. "000"
Public valSerial            As Integer      'serial number value

Public strThumbName         As String       'file and path for our thumbnail

'error logging
Public strErrLog            As String       'path to log file
Public strReport            As String       'contains path to (temp) report file
Public Const ByteSize = 50000               'sets maximum size error log file

'+++++++++++++++++++++++++++++++++++++ BROWSE FOR FOLDER ++++++++++++++++++++++++++++++++++++++++

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
     
     'open browse for folder dialog
     On Error Resume Next
     
     Dim iNull      As Integer
     Dim lpIDList   As Long
     Dim lResult    As Long
     Dim sPath      As String
     Dim udtBI      As BrowseInfo

    'initialise variables
     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

    'call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
     
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If

    'note: if cancel was selected, sPath returns ""
     BrowseForFolder = sPath

End Function

'+++++++++++++++++++++++++++++++++++++++++ GRAPHIC ++++++++++++++++++++++++++++++++++++++++++++++

Public Function MakeRegionTransparent(picSource As PictureBox) As Long
    
    'makes part of a picturebox transparent
    On Error Resume Next
    
    Dim x                   As Long
    Dim y                   As Long
    Dim StartLineX          As Long
    Dim FullRegion          As Long
    Dim LineRegion          As Long
    Dim TransparentColor    As Long
    Dim hDC                 As Long
    Dim picWidth            As Long
    Dim picHeight           As Long
    Dim InFirstRegion       As Boolean
    Dim InLine              As Boolean    'flags whether we are in a non-tranparent pixel sequence
    
    hDC = picSource.hDC
    picWidth = picSource.ScaleWidth
    picHeight = picSource.ScaleHeight
    
    InFirstRegion = True
    InLine = False
    x = y = StartLineX = 0
    
    'the transparent color is always the color of the
    'top-left pixel in the picture
    TransparentColor = GetPixel(hDC, 0, 0)
    
    'start pixel read out
    For y = 0 To picHeight - 1
        For x = 0 To picWidth - 1
           If GetPixel(hDC, x, y) = TransparentColor Or x = picWidth Then
                'we reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, y, x, y + 1)
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        'always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                'we reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
    
    MakeRegionTransparent = FullRegion
    
End Function

Public Sub GreyScale(picSource As PictureBox)
       
    'greyscale image
    On Error Resume Next
    
    Dim x       As Integer
    Dim y       As Integer
    Dim R       As Integer
    Dim G       As Integer
    Dim B       As Integer
    Dim GC      As Integer
    Dim pix     As Long

    For x = 0 To picSource.ScaleWidth

        For y = 0 To picSource.ScaleHeight
            pix = picSource.Point(x, y)
            RGBtoGrey pix, R, G, B
            GC = (R + G + B) / 3
            picSource.PSet (x, y), RGB(GC, GC, GC)
        Next y
        
    Next x
    
End Sub

Public Sub RGBtoGrey(ByVal Color As OLE_COLOR, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)
    
    'supports function greyscale
    On Error Resume Next
    
    B = Color \ 65536
    G = (Color \ 256) Mod 256
    R = Color Mod 256
    
End Sub

'++++++++++++++++++++++++++++++++++++++ FILES AND STRINGS +++++++++++++++++++++++++++++++++++++++

Public Function fixPath(strPath As String, strFilename As String) As String
    
    'fix "\"
    If Right(strPath, 1) = "\" Then
        fixPath = strPath & strFilename
    Else
        fixPath = strPath & "\" & strFilename
    End If
    
End Function

Public Function GetFileName(sFullPath As String) As String
    
    'get filename from path
    On Error Resume Next

    If InStr(1, sFullPath$, "/") > 0 Then
        GetFileName = Mid(sFullPath, InStrRev(sFullPath, "/") + 1)
    ElseIf InStr(1, sFullPath, "\") > 0 Then
        GetFileName$ = Mid(sFullPath, InStrRev(sFullPath, "\") + 1)
    Else
        GetFileName = sFullPath
    End If
    
    '--------------------------------------------------------------------------------------------
    'we can use this function also to determine the file path (without filename:
    'Debug.Print Mid(strFileName, 1, Len(strFileName) - Len(GetFileName(strFileName))-1)
    '--------------------------------------------------------------------------------------------
    
End Function

Public Function GetFileExtention(ByVal sStr As String) As String
    
    'get file extention
    On Error Resume Next
    
    Dim i As Integer

    i = InStrRev(sStr, ".")
    If i = 0 Then Exit Function
    GetFileExtention = LCase(Mid$(sStr, i + 1, 12))
    
End Function

Public Function FileExists(Path As String) As Boolean
  
  On Error Resume Next
  
  'returns true if file exists
  Const NotFile = vbDirectory + vbVolume
  FileExists = (GetAttr(Path) And NotFile) = 0
  
End Function

Public Function DirExists(ByVal strDirName As String) As Boolean

    Dim strDummy As String

    On Error Resume Next

    If Right$(strDirName, 1) <> "\" Then
        strDirName = strDirName & "\"
    End If

    strDummy = Dir$(strDirName & "*.*", vbDirectory)
    DirExists = Not (strDummy = vbNullString)

    Err = 0
    
End Function

Public Sub SafeKill(ByVal strFilename As String)
    
    If FileExists(strFilename) Then
        Kill strFilename
    End If
    
End Sub

'++++++++++++++++++++++++++++++++++++ E R R O R  L O G G I N G ++++++++++++++++++++++++++++++++++

Public Sub chkErrorLog(errFile As String, bitVal As Long)

    'check if the logfile isn't larger then (XX)* bytes, if so we delete the old .bck
    'file (if there is one), rename the current file to .bck and start a new logfile
    '*(XX) is the value of Constant Bytsize
    
    On Error Resume Next
    
    Dim curSize         As Long         'length of file in bytes
    Dim myString        As String       'find old log file
    Dim FF              As Integer      'free file
        
    If FileExists(errFile) = True Then
        
        FF = FreeFile
        
        Open errFile For Append As #FF
            curSize = LOF(FF)
        Close #FF
    
        If curSize > bitVal Then
            
            'remove extention from filename
            myString = Left(errFile, Len(errFile) - 4)
            
            'delete the old bck file (if exists)
            If FileExists(myString & ".bck") = True Then
                Kill myString & ".bck"
            End If
            
            'rename the old log file to .bck
            Name errFile As myString & ".bck"
            
            'write date & time to the new log file
            Open errFile For Append As #FF
                Print #FF, "Log started: " & Now
                Print #FF, "--------------------------------"
                Print #FF, ""
            Close #FF
           
        End If
        
    End If
    
    If FileExists(errFile) = False Then
        'create a (new) logfile
        
        FF = FreeFile
        'write date & time to the new log file
        Open errFile For Append As #FF
            Print #FF, "Log started: " & Now
            Print #FF, "--------------------------------"
            Print #FF, ""
        Close #FF
        
    End If
        
End Sub

Public Function writeErrorLog(errFile As String, strMsg As String)
    
    'write to error log
    On Error Resume Next
    
    Dim FF As Integer
    
    FF = FreeFile
    
    Open errFile For Append As #FF
        Print #FF, strMsg
    Close #FF
    
End Function

Public Function startReport(strSource As String, strDest As String, strFrom As String, strTo As String)
    
    Dim FF          As Integer      'counter
    Dim strBatch    As String       'convert/process
    
    FF = FreeFile
    
    If blnConverting = True Then strBatch = "Convert"
    If blnProcessing = True Then strBatch = "Process"
    
    Open strReport For Output As #FF
        
        Print #FF, "----------------------------------------------------------------------------"
        Print #FF, "                    Photo Logo Batch " & strBatch & " Report"
        Print #FF, "----------------------------------------------------------------------------"
        Print #FF, ""
        Print #FF, Format(Now, "dd-mm-yyyy - hh:mm:ss")
        Print #FF, ""
        Print #FF, "Source Directory:          " & strSource
        Print #FF, "Destination Directory:     " & strDest
        Print #FF, strBatch & ":                   " & strFrom & " - " & strTo
        Print #FF, ""
        Print #FF, "Errors: "
        Print #FF, ""
        
    Close #FF
        
End Function

Public Function addReport(str As String)
    
    'add a line to the report
    Dim FF As Integer
    FF = FreeFile
    
    Open strReport For Append As #FF
        
        Print #FF, str
        
    Close #FF
    
End Function

'++++++++++++++++++++++++++++++++++++++++++++++ MISC ++++++++++++++++++++++++++++++++++++++++++++

Public Sub resetPostionFlags()
    
    'reset logo position flags
    On Error Resume Next
    
    blnLeftTop = False
    blnRightTop = False
    blnLeftBot = False
    blnRightBot = False
    blnCenter = False
    
End Sub
