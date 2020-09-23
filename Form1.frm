VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "MrBobos' Basic Text Editor"
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7920
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0442
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0556
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0672
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":078E
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08AA
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09C6
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0AE2
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0BFE
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D1A
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E36
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F52
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":106E
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1186
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":129E
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13B6
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14CE
            Key             =   "Font"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C22
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Undo"
            ImageKey        =   "Undo"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Redo"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            ImageKey        =   "Left"
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            ImageKey        =   "Right"
            Style           =   2
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            ImageKey        =   "Font"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Red"
                  Text            =   "Red"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Blue"
                  Text            =   "Blue"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Green"
                  Text            =   "Green"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Black"
                  Text            =   "Black"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "spac"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Lower"
                  Text            =   "Lower case"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Upper"
                  Text            =   "Upper Case"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            ImageKey        =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   5
   End
   Begin RichTextLib.RichTextBox rtftext 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":1D36
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuFileMRUSpace 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   "MRU"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEditSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFormat 
         Caption         =   "Format"
      End
      Begin VB.Menu mnuEditRemBlanks 
         Caption         =   "Remove Blank Lines"
      End
      Begin VB.Menu mnuEditSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "Redo"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the best way to perform edit functions
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const WM_UNDO = &H304
Public curfile As String 'Full path of file in the rich textbox
Dim curEXT As String 'Example "txt"
Dim lastopened As String 'Path to last opened from dir
Dim lastopenedEXT As String 'Extension of last opened file
Dim lastsaved As String 'Path to last saved to dir
Dim MRU(0 To 4) As String 'menu for mru
Dim savedas As String ' path to undo files
Dim Undobuffer() As String 'array of undo paths
Dim UndobufferCur() As Long 'array of cursor positions at undo
Dim UndoCount As Integer 'how many undos
Dim UndoPosition As Integer 'where we are in the undo/redo sequence
Dim NewBoy As Boolean 'it's a new empty file
Dim AlreadySaved As Boolean 'dont prompt to save



Private Sub Form_Load()
Dim x As Integer, temp As String
'used to make CommonDialog appear intelligent
lastsaved = GetSetting(App.Title, "Paths", "Lastsaved", "C:\")
lastopened = GetSetting(App.Title, "Paths", "LastOpened", "C:\")
lastopenedEXT = GetSetting(App.Title, "Paths", "LastOpenedEXT", "txt")
'Fill our menu with MRUs
For x = 0 To 4
    temp = GetSetting(App.Title, "Paths", "MRU" + Str(x), "")
    If temp <> "" Then
        mnuFileMRUSpace.Visible = True
        mnuFileMRU(x).Tag = temp
        mnuFileMRU(x).Caption = FileOnly(temp)
        mnuFileMRU(x).Visible = True
        MRU(x) = temp
    End If
Next x
'Fill the Find/Replace combos with thre last searches made
For x = 0 To 9
    temp = GetSetting(App.Title, "Search", "Find" + Str(x), "")
    If temp <> "" Then frmFind.Findme.AddItem temp
    temp = GetSetting(App.Title, "Search", "Replace" + Str(x), "")
    If temp <> "" Then frmFind.Replaceme.AddItem temp
Next x
'Show my opening blurb
'Delete this along with the RES file if you like
LoadRTFres rtftext, 101
NewBoy = True
'Initialise the Undo system
Backup
mnuEditUndo.Enabled = False
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = False
TB.Buttons(11).Enabled = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim x As Integer
'Remember everything for next time
SaveSetting App.Title, "Paths", "Lastsaved", lastsaved
SaveSetting App.Title, "Paths", "LastOpened", lastopened
SaveSetting App.Title, "Paths", "LastOpenedEXT", lastopenedEXT
For x = 0 To 4
    If mnuFileMRU(x).Visible Then
        SaveSetting App.Title, "Paths", "MRU" + Str(x), mnuFileMRU(x).Tag
    Else
        SaveSetting App.Title, "Paths", "MRU" + Str(x), ""
    End If
Next x
For x = 0 To 9
    SaveSetting App.Title, "Search", "Find" + Str(x), frmFind.Findme.List(x)
    SaveSetting App.Title, "Search", "Replace" + Str(x), frmFind.Replaceme.List(x)
Next x
'Get rid of the tmp files used for the Undo system
For x = 1 To UndoCount
    If FileExists(Undobuffer(x)) Then Kill Undobuffer(x)
Next x
Checksafe
End Sub

Private Sub Form_Resize()
On Error Resume Next
rtftext.Top = TB.Top + TB.Height
rtftext.Height = Me.Height - rtftext.Top - 680
rtftext.Width = Me.Width - 90
End Sub

Private Sub Form_Unload(Cancel As Integer)
'We're really closing - better tell the others this is it !
FinalClose = True
Unload frmFind
Unload frmFormat
End Sub

Private Sub mnuEditCopy_Click()
SendMessage rtftext.hWnd, WM_COPY, 0, 0

End Sub

Private Sub mnuEditCut_Click()
SendMessage rtftext.hWnd, WM_CUT, 0, 0
Backup
End Sub

Private Sub mnuEditDelete_Click()
SendMessage rtftext.hWnd, WM_CLEAR, 0, 0
Backup

End Sub

Private Sub mnuEditFormat_Click()
frmFormat.Show vbModal
End Sub

Private Sub mnuEditPaste_Click()
SendMessage rtftext.hWnd, WM_PASTE, 0, 0
Backup

End Sub

Private Sub mnuEditRedo_Click()
'Look - I dont know how to explain this - let's see
'We make temp files each time they change things.
'We keep track of things using an array
'When they want to undo or redo something
'the array tells us which temp file to load
LockWindowUpdate rtftext.hWnd
mnuEditUndo.Enabled = True
TB.Buttons(10).Enabled = True
rtftext.LoadFile Undobuffer(UndoPosition + 1)
rtftext.SelStart = UndobufferCur(UndoPosition + 1)
UndoPosition = UndoPosition + 1
If UndoPosition = UndoCount Then
    mnuEditRedo.Enabled = False
    TB.Buttons(11).Enabled = False
End If
LockWindowUpdate 0

End Sub

Private Sub mnuEditRemBlanks_Click()
'Call my remove blanks function
RemoveRTFblanks rtftext
Backup
End Sub

Private Sub mnuEditUndo_Click()
'Look - I dont know how to explain this - let's see
'We make temp files each time they change things.
'We keep track of things using an array
'When they want to undo or redo something
'the array tells us which temp file to load
LockWindowUpdate rtftext.hWnd
rtftext.LoadFile Undobuffer(UndoPosition - 1)
rtftext.SelStart = UndobufferCur(UndoPosition - 1)
If UndoPosition > 1 Then
    UndoPosition = UndoPosition - 1
End If
If UndoPosition < 2 Then
    mnuEditUndo.Enabled = False
    TB.Buttons(10).Enabled = False
End If
mnuEditRedo.Enabled = True
TB.Buttons(11).Enabled = True
LockWindowUpdate 0
End Sub

Private Sub mnuFileAbout_Click()
'I wrote this - yes ME !!
MsgBox "RichTextBox Example - Bobo Enterprises 2001"
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileMRU_Click(Index As Integer)
Dim x As Integer
Dim isVis As Boolean
Checksafe
If FileExists(mnuFileMRU(Index).Tag) Then
    ClearUndo
    rtftext.LoadFile mnuFileMRU(Index).Tag
    curfile = mnuFileMRU(Index).Tag
    curEXT = ExtOnly(curfile)
    Backup
    mnuEditUndo.Enabled = False
    mnuEditRedo.Enabled = False
    TB.Buttons(10).Enabled = False
    TB.Buttons(11).Enabled = False
Else
    MsgBox "This file has been moved or deleted."
    MRU(Index) = ""
    mnuFileMRU(Index).Visible = False
    For x = 0 To mnuFileMRU.count - 1
        If mnuFileMRU(x).Visible = True Then isVis = True
    Next x
    mnuFileMRUSpace.Visible = isVis
End If

End Sub

Private Sub mnuFileNew_Click()
Checksafe 'Do they want to save changes before we move on
ClearUndo
NewBoy = True
curfile = ""
rtftext.SelStart = 0
rtftext.SelLength = Len(rtftext.Text)
SendMessage rtftext.hWnd, WM_CLEAR, 0, 0
Me.Caption = "Untitled"
Backup
mnuEditUndo.Enabled = False
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = False
TB.Buttons(11).Enabled = False
AlreadySaved = False
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo woops
Checksafe
NewBoy = False
 With CommonDialog1
 'use our variables to open the commondialog appropriately
    If lastopenedEXT = "" Then lastopenedEXT = "txt"
    If lastopened = "" Then lastopened = "C:\"
    Select Case LCase(lastopenedEXT)
        Case "txt"
            .FilterIndex = 1
        Case "doc"
            .FilterIndex = 2
        Case Else
            .FilterIndex = 3
    End Select
    .InitDir = lastopened
    .DialogTitle = "Open Text Files"
    .CancelError = True
    .Filter = "Text files (*.txt)|*.txt|Document files (*.doc)|*.doc|All files (*.*)|*.*"
    .ShowOpen
    If Len(.FileName) = 0 Then Exit Sub
    curfile = .FileName
    lastopened = PathOnly(.FileName)
    lastopenedEXT = ExtOnly(.FileName)
    ClearUndo
End With
curEXT = ExtOnly(curfile)
rtftext.LoadFile curfile
Me.Caption = FileOnly(curfile)
'reset the MRU menu
fixMRUs curfile
'Reset the Undo system
Backup
mnuEditUndo.Enabled = False
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = False
TB.Buttons(11).Enabled = False
'reset the warning about changed files
AlreadySaved = False
woops: Exit Sub

End Sub

Private Sub mnuFileSave_Click()
NewBoy = False
Select Case LCase(curEXT)
    Case "txt"
        FileSave rtftext.Text, curfile
    Case "doc"
        rtftext.SaveFile curfile
    Case Else
        FileSave rtftext.Text, curfile
End Select
lastsaved = PathOnly(curfile)
AlreadySaved = True
End Sub

Private Sub mnuFileSaveAs_Click()
On Error GoTo woops
If lastsaved = "" Then lastsaved = "C:\"
NewBoy = False
With CommonDialog1
    .InitDir = lastsaved
    .DialogTitle = "Save Text Files"
    .CancelError = True
    .Filter = "Text files (*.txt)|*.txt|Document files (*.doc)|*.doc|All files (*.*)|*.*"
    Select Case LCase(curEXT)
        Case "txt"
            .FilterIndex = 1
        Case "doc"
            .FilterIndex = 2
        Case Else
            .FilterIndex = 3
    End Select
    .ShowSave
    If Len(.FileName) = 0 Then Exit Sub
    curfile = .FileName
    curEXT = ExtOnly(curfile)
    lastsaved = PathOnly(curfile)
End With
Select Case LCase(curEXT)
    Case "txt"
        FileSave rtftext.Text, curfile
    Case "doc"
        rtftext.SaveFile curfile
    Case Else
        FileSave rtftext.Text, curfile
End Select
Me.Caption = FileOnly(curfile)
AlreadySaved = True
woops: Exit Sub

End Sub


Private Sub rtftext_KeyUp(KeyCode As Integer, Shift As Integer)
'Something's changed tell the undo system
Backup

End Sub

Private Sub rtftext_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dont show the default context menu - we've got a
'better undo system
If Button = 2 Then Me.PopupMenu mnuEdit
End Sub

Private Sub rtftext_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'makes the buttons seem intelligent
If rtftext.SelBold = True Then
  TB.Buttons(13).Value = tbrPressed
Else
  TB.Buttons(13).Value = tbrUnpressed
End If

If rtftext.SelItalic = True Then
  TB.Buttons(14).Value = tbrPressed
Else
  TB.Buttons(14).Value = tbrUnpressed
End If

If rtftext.SelUnderline = True Then
  TB.Buttons(15).Value = tbrPressed
Else
  TB.Buttons(15).Value = tbrUnpressed
End If
If rtftext.SelAlignment = rtfLeft Then
    TB.Buttons(17).Value = tbrPressed
Else
    TB.Buttons(17).Value = tbrUnpressed
End If

If rtftext.SelAlignment = rtfCenter Then
    TB.Buttons(18).Value = tbrPressed
Else
    TB.Buttons(18).Value = tbrUnpressed
End If

If rtftext.SelAlignment = rtfRight Then
    TB.Buttons(19).Value = tbrPressed
Else
    TB.Buttons(19).Value = tbrUnpressed
End If
If rtftext.SelLength > 0 Then
    frmFind.Findme.Text = rtftext.SelText
    frmFormat.rtfHighLightString.Text = rtftext.SelText
End If
End Sub

Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)
'Standard Toolbar stuff
Select Case Button.Key
    Case "New"
        mnuFileNew_Click
    Case "Open"
        mnuFileOpen_Click
    Case "Save"
        mnuFileSave_Click
    Case "Cut"
        mnuEditCut_Click
    Case "Copy"
        mnuEditCopy_Click
    Case "Paste"
        mnuEditPaste_Click
    Case "Delete"
        mnuEditDelete_Click
    Case "Undo"
        mnuEditUndo_Click
    Case "Redo"
        mnuEditRedo_Click
    Case "Find"
        frmFind.Show
        FloatWindow frmFind.hWnd, FLOAT
    Case "Bold"
       If rtftext.SelBold = True Then
        rtftext.SelBold = False
        TB.Buttons(13).Value = tbrUnpressed
       Else
        rtftext.SelBold = True
        TB.Buttons(13).Value = tbrPressed
       End If
       Backup
    Case "Italic"
       If rtftext.SelItalic = True Then
        rtftext.SelItalic = False
         TB.Buttons(14).Value = tbrUnpressed
       Else
        rtftext.SelItalic = True
       TB.Buttons(14).Value = tbrPressed
        End If
       Backup
    Case "Underline"
       If rtftext.SelUnderline = True Then
        rtftext.SelUnderline = False
         TB.Buttons(15).Value = tbrUnpressed
       Else
        rtftext.SelUnderline = True
        TB.Buttons(15).Value = tbrPressed
       End If
       Backup
    Case "Right"
        rtftext.SelAlignment = rtfRight
       Backup
    Case "Center"
        rtftext.SelAlignment = rtfCenter
       Backup
    Case "Left"
        rtftext.SelAlignment = rtfLeft
       Backup
    Case "Font"
        On Error GoTo woops
        CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects
        With rtftext
            CommonDialog1.FontName = .SelFontName
            CommonDialog1.FontSize = .SelFontSize
            CommonDialog1.FontBold = .SelBold
            CommonDialog1.FontItalic = .SelItalic
            CommonDialog1.FontStrikethru = .SelStrikeThru
            CommonDialog1.FontUnderline = .SelUnderline
            CommonDialog1.Color = .SelColor
        End With
        CommonDialog1.ShowFont
        With rtftext
            .SelFontName = CommonDialog1.FontName
            .SelFontSize = CommonDialog1.FontSize
            .SelBold = CommonDialog1.FontBold
            .SelItalic = CommonDialog1.FontItalic
            .SelStrikeThru = CommonDialog1.FontStrikethru
            .SelUnderline = CommonDialog1.FontUnderline
            .SelColor = CommonDialog1.Color
        End With
       Backup
End Select
woops:
End Sub


Public Sub fixMRUs(filepath As String)
'Used to reset the MRUs' when another file has been opened
Dim x As Integer, count As Integer
mnuFileMRU(0).Caption = FileOnly(filepath)
mnuFileMRU(0).Tag = filepath
mnuFileMRU(0).Visible = True
mnuFileMRUSpace.Visible = True
For x = 0 To 3
    If MRU(x) <> "" Then
        count = count + 1
        mnuFileMRU(count).Caption = FileOnly(MRU(x))
        mnuFileMRU(count).Tag = MRU(x)
        mnuFileMRU(count).Visible = True
    Else
        mnuFileMRU(x + 1).Visible = False
    End If
Next x
For x = 0 To 4
    If mnuFileMRU(x).Visible Then
        MRU(x) = mnuFileMRU(x).Tag
    End If
Next x

End Sub


Public Sub Backup()
'Look - I dont know how to explain this - let's see
'We make temp files each time they change things.
'We keep track of things using an array
'When they want to undo or redo something
'the array tells us which temp file to load
Dim x As Integer
If UndoPosition <> UndoCount Then
    For x = UndoPosition + 1 To UndoCount
        If FileExists(Undobuffer(x)) Then Kill Undobuffer(x)
    Next x
    ReDim Preserve Undobuffer(UndoPosition)
    ReDim Preserve Undobuffer(UndoPosition + 10)
    UndoCount = UndoPosition
End If
If (UndoCount Mod 10) = 0 Then
    ReDim Preserve Undobuffer(UndoCount + 10)
    ReDim Preserve UndobufferCur(UndoCount + 10)
End If
UndoCount = UndoCount + 1
If GetTempFile(savedas) Then
    Undobuffer(UndoCount) = savedas
    UndobufferCur(UndoCount) = rtftext.SelStart
    rtftext.SaveFile savedas
End If
UndoPosition = UndoCount
mnuEditUndo.Enabled = True
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = mnuEditUndo.Enabled
TB.Buttons(11).Enabled = mnuEditRedo.Enabled
If UndoCount > 1 Then AlreadySaved = False
End Sub

Public Sub ClearUndo()
'Clear the Undos'- they've either saved or opened a new file
mnuEditUndo.Enabled = False
mnuEditRedo.Enabled = False
TB.Buttons(10).Enabled = False
TB.Buttons(11).Enabled = False
Dim x As Integer
For x = 1 To UndoCount
    If FileExists(Undobuffer(x)) Then Kill Undobuffer(x)
Next x
ReDim Undobuffer(0 To 10)
UndoCount = 0
UndoPosition = 0

End Sub

Public Sub Checksafe()
'Has the file changed - better tell someone
If UndoCount < 2 Or AlreadySaved = True Then
    GoTo endum
Else
    If MsgBox("File has changed. Do you wish to save changes ?", vbQuestion + vbYesNo, "Bobo Enterprises") = vbNo Then
        GoTo endum
    Else
        If NewBoy Then
            mnuFileSaveAs_Click
            GoTo endum
        Else
            mnuFileSave_Click
            GoTo endum
        End If
    End If
End If
endum:
    Exit Sub
End Sub

Private Sub TB_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
'The dropdown menus enable quick access foe editing
'color and case
Select Case ButtonMenu.Key
    Case "Black"
    rtftext.SelColor = vbBlack
    Case "Red"
    rtftext.SelColor = vbRed
    Case "Blue"
    rtftext.SelColor = vbBlue
    Case "Green"
    rtftext.SelColor = vbGreen
    Case "Upper"
    rtftext.SelText = UCase(rtftext.SelText)
    Case "Lower"
    rtftext.SelText = LCase(rtftext.SelText)

End Select

End Sub
