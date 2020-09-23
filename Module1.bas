Attribute VB_Name = "Module1"
Option Explicit
'Used for Undo/redo and loading the initial message from me
Private Declare Function GetTempFilename Lib "Kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFilename As String _
    ) As Long
Private Declare Function GetTempPath Lib "Kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Used to "Float" the Find/Replace window
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const FLOAT = 1, SINK = 0
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'This just stops things jumping around and updating before we're ready
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
'This stops the Find window unloading till Program close when we save settings
Public FinalClose As Boolean
'Set a form to On Top
Sub FloatWindow(x As Long, action As Integer)
Dim wFlags As Integer, result As Integer
wFlags = SWP_NOMOVE Or SWP_NOSIZE
If action <> 0 Then
    Call SetWindowPos(x, HWND_TOPMOST, 0, 0, 0, 0, wFlags)
Else
    Call SetWindowPos(x, HWND_NOTOPMOST, 0, 0, 0, 0, wFlags)
End If
End Sub
'The next 4 functions make handling files easier by parsing
Public Function PathOnly(ByVal filepath As String) As String
Dim temp As String
    temp = Mid$(filepath, 1, InStrRev(filepath, "\"))
    If Right(temp, 1) = "\" Then temp = Left(temp, Len(temp) - 1)
    PathOnly = temp
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = Mid$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function ExtOnly(ByVal filepath As String, Optional dot As Boolean) As String
    ExtOnly = Mid$(filepath, InStrRev(filepath, ".") + 1)
If dot = True Then ExtOnly = "." + ExtOnly
End Function
Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
Dim temp As String
temp = Mid$(filepath, 1, InStrRev(filepath, "."))
temp = Left(temp, Len(temp) - 1)
If newext <> "" Then newext = "." + newext
ChangeExt = temp + newext
End Function
'Save a plain text file without the RTF gobbledygook
Public Sub FileSave(Text As String, filepath As String)
On Error Resume Next
Dim Directory As String
              Directory$ = filepath
              Open Directory$ For Output As #1
           Print #1, Text
       Close #1
Exit Sub
End Sub
'Stops us attempting to do something to a file that isn't there
Function FileExists(ByVal FileName As String) As Integer
Dim temp$, MB_OK
    FileExists = True
On Error Resume Next
    temp$ = FileDateTime(FileName)
    Select Case Err
        Case 53, 76, 68
            FileExists = False
            Err = 0
        Case Else
            If Err <> 0 Then
                MsgBox "Error Number: " & Err & Chr$(10) & Chr$(13) & " " & Error, MB_OK, "Error"
            End If
    End Select
End Function
'API often returns a string with a null character at the end
Public Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
'Where's the temp folder ?
Public Function temppath() As String
    Dim sBuffer As String
    Dim lRet As Long
    sBuffer = String$(255, vbNullChar)
    lRet = GetTempPath(255, sBuffer)
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If
    temppath = sBuffer
    If Right(temppath, 1) = "\" Then temppath = Left(temppath, Len(temppath) - 1)
End Function
'Get a unique name for a temp file
Public Function GetTempFile(lpTempFilename As String) As Boolean
    lpTempFilename = String(255, vbNullChar)
    GetTempFile = GetTempFilename(temppath, "bb", 0, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function
'Shrink a path to fit a label or text box
Public Function LabelEdit(Path As String, Length As Integer) As String
Dim temp As String, temp1 As String, temp2 As String
If Len(Path) < Length Then
    LabelEdit = Path
    Exit Function
End If
temp = Mid$(Path, InStrRev(Path, "\") + 1)
If Len(temp) + 7 < Length Then
    temp1 = Mid$(Path, 1, InStrRev(Path, "\") - 1)
    temp2 = Mid$(temp1, InStrRev(temp1, "\") + 1)
    If Len(temp) + 7 + Len(temp2) < Length Then temp = temp2 + "\" + temp
End If
LabelEdit = Left(Path, 3) + "...\" + temp
End Function
'Remove empty lines from a richtextbox
Public Sub RemoveRTFblanks(RTF As RichTextBox)
Dim textfound As Long
Dim Position As Long
Dim St As Long
Dim Lng As Long
Dim temp As String
St = 0
If RTF.Text = "" Then Exit Sub
Screen.MousePointer = 11
LockWindowUpdate RTF.hWnd
Do Until Position >= Len(RTF.Text)
    textfound = RTF.Find(vbCrLf, St)
    If textfound = -1 Then GoTo fin
    Position = RTF.SelStart + RTF.SelLength
    RTF.SelStart = St
    RTF.SelLength = Position - St - 1
    If RTF.SelText <> String(Len(RTF.SelText), " ") Then temp = temp + RTF.SelText + vbCrLf
    St = Position
Loop
fin:
RTF.Text = temp
LockWindowUpdate 0
Screen.MousePointer = 0

End Sub
'Used to load my initial blurb
Public Sub LoadRTFres(rtftext As RichTextBox, mynum As Integer)
Dim sFileName As String
If GetTempFile2("", "~rs", 0, sFileName) Then
    If Not SaveResItemToDisk(mynum, "Custom", sFileName) Then
        rtftext.LoadFile sFileName
        Kill sFileName
    Else
        MsgBox "Unable to save resource item to disk!", vbCritical
    End If
Else
    MsgBox "Unable to get temp file name!", vbCritical
End If

End Sub
'Used to load my initial blurb
Public Function GetTempFile2( _
    ByVal strDestPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Integer, _
    lpTempFilename As String _
    ) As Boolean
   If strDestPath = "" Then
        strDestPath = String(255, vbNullChar)
        If GetTempPath(255, strDestPath) = 0 Then
            GetTempFile2 = False
            Exit Function
        End If
    End If
    lpTempFilename = String(255, vbNullChar)
    GetTempFile2 = GetTempFilename(strDestPath, lpPrefixString, wUnique, lpTempFilename) > 0
    lpTempFilename = StripTerminator(lpTempFilename)
End Function
'Used to load my initial blurb
Public Function SaveResItemToDisk( _
            ByVal iResourceNum As Integer, _
            ByVal sResourceType As String, _
            ByVal sDestFileName As String _
            ) As Long
    Dim bytResourceData()   As Byte
    Dim iFileNumOut         As Integer
    On Error GoTo SaveResItemToDisk_err
    bytResourceData = LoadResData(iResourceNum, sResourceType)
    iFileNumOut = FreeFile
    Open sDestFileName For Binary Access Write As #iFileNumOut
        Put #iFileNumOut, , bytResourceData
    Close #iFileNumOut
    SaveResItemToDisk = 0
    Exit Function
SaveResItemToDisk_err:
    SaveResItemToDisk = Err.Number
End Function

