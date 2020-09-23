VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFormat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Format Wizard"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmFormat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   390
      Left            =   240
      ScaleHeight     =   330
      ScaleWidth      =   360
      TabIndex        =   60
      Top             =   4440
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Flags           =   5
   End
   Begin VB.Frame FrFormat 
      Height          =   3015
      Index           =   2
      Left            =   5500
      TabIndex        =   10
      Top             =   360
      Width           =   3615
      Begin VB.CheckBox Check2 
         Caption         =   "Strikethrough"
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Underline"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmFormat.frx":0442
         Left            =   2160
         List            =   "frmFormat.frx":044F
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Bold/Italic"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Italic"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2040
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Bold"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Regular"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1560
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdColor 
         BackColor       =   &H80000008&
         Height          =   255
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2640
         Width           =   255
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Alignment"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Color"
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Size"
         Height          =   255
         Left            =   2280
         TabIndex        =   17
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Font"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Select the style of formatting you wish to apply"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame FrFormat 
      Caption         =   "Summary"
      Height          =   3015
      Index           =   8
      Left            =   5500
      TabIndex        =   51
      Top             =   360
      Width           =   3615
      Begin VB.Label LblAction 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   59
         Top             =   2640
         Width           =   3255
      End
      Begin VB.Label LblAction 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   57
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label LblAction 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   56
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label LblAction 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   55
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label LblAction 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   54
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label LblAction 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label LblAction 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   52
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.Frame FrFormat 
      Height          =   3015
      Index           =   7
      Left            =   5500
      TabIndex        =   48
      Top             =   360
      Width           =   3615
      Begin RichTextLib.RichTextBox rtfInsertText 
         Height          =   2175
         Left            =   240
         TabIndex        =   49
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3836
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmFormat.frx":0468
      End
      Begin VB.Label Label14 
         Caption         =   "Enter the text you wish to insert"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FrFormat 
      Height          =   3015
      Index           =   6
      Left            =   5500
      TabIndex        =   39
      Top             =   360
      Width           =   3615
      Begin VB.CommandButton Command2 
         Caption         =   "Browse"
         Height          =   255
         Left            =   2520
         TabIndex        =   41
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   840
         Width           =   3135
      End
      Begin RichTextLib.RichTextBox rtfThumb 
         Height          =   1215
         Left            =   240
         TabIndex        =   44
         Top             =   1560
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2143
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmFormat.frx":0531
      End
      Begin VB.Label Label12 
         Caption         =   "Select a text file"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label11 
         Caption         =   "Path to your text file"
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame FrFormat 
      Height          =   3015
      Index           =   5
      Left            =   5500
      TabIndex        =   33
      Top             =   360
      Width           =   3615
      Begin VB.OptionButton Option12 
         Caption         =   "Insert text from a file"
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Enter text for insertion "
         Height          =   255
         Left            =   840
         TabIndex        =   45
         Top             =   1080
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.Label Label13 
         Caption         =   "Do you wish to ..."
         Height          =   375
         Left            =   480
         TabIndex        =   47
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Frame FrFormat 
      Height          =   3015
      Index           =   4
      Left            =   5500
      TabIndex        =   32
      Top             =   360
      Width           =   3615
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frmFormat.frx":05FA
         Left            =   360
         List            =   "frmFormat.frx":0607
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Browse"
         Height          =   255
         Left            =   2520
         TabIndex        =   36
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   35
         Top             =   840
         Width           =   3135
      End
      Begin VB.PictureBox PicThumb 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   2280
         ScaleHeight     =   1095
         ScaleWidth      =   1095
         TabIndex        =   34
         Top             =   1680
         Width           =   1095
         Begin VB.Image ImgThumb 
            Height          =   855
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Select an image"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label9 
         Caption         =   "Path to your Image file"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.Frame FrFormat 
      Height          =   3015
      Index           =   3
      Left            =   5500
      TabIndex        =   11
      Top             =   360
      Width           =   3615
      Begin VB.OptionButton Option10 
         Caption         =   "Image"
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Text"
         Height          =   255
         Left            =   960
         TabIndex        =   29
         Top             =   1320
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "What do you wish to insert ?"
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame FrFormat 
      Height          =   3015
      Index           =   1
      Left            =   5500
      TabIndex        =   9
      Top             =   360
      Width           =   3615
      Begin RichTextLib.RichTextBox rtfHighLightString 
         Height          =   855
         Left            =   720
         TabIndex        =   13
         Top             =   1560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmFormat.frx":0620
      End
      Begin VB.Label Label2 
         Caption         =   "Enter the word or group of words you wish to highlight"
         Height          =   495
         Left            =   840
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      Height          =   300
      Left            =   3720
      TabIndex        =   7
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< Back"
      Enabled         =   0   'False
      Height          =   300
      Left            =   2640
      TabIndex        =   5
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame FrFormat 
      Height          =   3015
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3615
      Begin VB.OptionButton Option4 
         Caption         =   "Insert at Cursor"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Insert Footer"
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Insert Header"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "HighLight Words"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "How would you like to format the current document ?"
         Height          =   495
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Just to insert a picture in a Richtextbox
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302
'What stage is the wizard up to ?
Dim level As Integer
Private Sub cmdBack_Click()
'Woops I want to go back and change something
level = level - 1
If level = 1 Then cmdBack.Enabled = False
RetroFixLevels
End Sub
Private Sub cmdCancel_Click()
'Bail out
Unload Me
End Sub
Private Sub cmdColor_Click()
'Select color using CommonDialog
On Error GoTo woops
CommonDialog1.CancelError = True
CommonDialog1.Flags = 0
CommonDialog1.action = 3
cmdColor.BackColor = CommonDialog1.Color
woops: Exit Sub

End Sub
Private Sub cmdNext_Click()
'OK I'm done, next Wizard page please
If cmdNext.Caption = "Finish" Then
    WindUp
    Exit Sub
End If
cmdBack.Enabled = True
level = level + 1
FixLevels
End Sub
Private Sub Command1_Click()
'Choose an image to insert
On Error GoTo woops
 With CommonDialog1
  .DialogTitle = "Open Image Files"
  .CancelError = True
  .Filter = "Picture files |*.bmp;*.jpg;*.gif;*.ico;*.cur|All files (*.*)|*.*"
  .ShowOpen
 If Len(.FileName) = 0 Then Exit Sub
 Text1.Text = .FileName
End With
Picture1.Picture = LoadPicture(Text1.Text)
SizeThumb
woops: Exit Sub

End Sub

Private Sub Command2_Click()
'Choose a text file to insert
On Error GoTo woops
 With CommonDialog1
  .DialogTitle = "Open Image Files"
  .CancelError = True
  .Filter = "Text files |*.doc;*.txt;*.ini;*.log|All files (*.*)|*.*"
  .ShowOpen
 If Len(.FileName) = 0 Then Exit Sub
 Text2.Text = .FileName
End With
rtfThumb.LoadFile Text2.Text
SizeThumb
woops: Exit Sub

End Sub

Private Sub Form_Load()
'Fill our combo boxes with fonts etc.
Dim x As Integer
Dim fsz As Integer
Me.Width = 4965
level = 1
For x = 0 To Screen.FontCount - 1
    Combo1.AddItem Screen.Fonts(x)
Next x
For x = 6 To 12
    Combo2.AddItem Str(x)
Next x
fsz = 14
For x = 0 To 30
    Combo2.AddItem Str(fsz)
    fsz = fsz + 2
Next x
For x = 0 To Combo1.ListCount - 1
    If Combo1.List(x) = "MS Sans Serif" Then
        Combo1.ListIndex = x
        Exit For
    End If
Next x
Combo2.ListIndex = 2
Combo3.ListIndex = 0
Combo4.ListIndex = 0
MoveFrames 0
End Sub

Public Sub FixLevels()
'Make sure the right thing shows up in the wizards' next page
If Option1.Value = True Then
    Select Case level
    Case 2
        MoveFrames 1
    Case 3
        If Len(rtfHighLightString.Text) = 0 Then
            MsgBox "No text entered for highlighting"
            level = 2
            Exit Sub
        End If
        Label7.Visible = False
        Combo3.Visible = False
        MoveFrames 2
    Case 4
        DrawSummary
        MoveFrames 8
        cmdNext.Caption = "Finish"
    End Select
Else
    Select Case level
    Case 2
        MoveFrames 3
    Case 3
        If Option9.Value = True Then
            MoveFrames 5
        Else
            MoveFrames 4
        End If
    Case 4
        If Option9.Value = True Then
            If Option11.Value = True Then
                MoveFrames 7
            Else
                MoveFrames 6
            End If
        Else
            If Len(Text1.Text) = 0 Then
                MsgBox "Please enter a path to your file"
                level = 3
                Exit Sub
            End If
            If Not FileExists(Text1.Text) Then
                MsgBox "Cant find this file." + vbCrLf + "Please enter a valid path to your file"
                level = 3
                Exit Sub
            End If
            DrawSummary
            MoveFrames 8
            cmdNext.Caption = "Finish"
        End If
    Case 5
        If Option12.Value = True Then
        If Len(Text2.Text) = 0 Then
            MsgBox "Please enter a path to your file"
            level = 4
            Exit Sub
        End If
        If Not FileExists(Text2.Text) Then
            MsgBox "Cant find this file." + vbCrLf + "Please enter a valid path to your file"
            level = 4
            Exit Sub
        End If
        Else
        If Len(rtfInsertText.Text) = 0 Then
            MsgBox "No text entered for inserting"
            level = 4
            Exit Sub
        End If
        End If
        Label7.Visible = True
        Combo3.Visible = True
        MoveFrames 2
    Case 6
        DrawSummary
        MoveFrames 8
        cmdNext.Caption = "Finish"
    End Select
End If
End Sub
Public Sub MoveFrames(Fr As Integer)
'This is how a wizard shows a new page
'It just sets the left property to an unseeable location
Dim x As Integer
For x = 0 To FrFormat.count - 1
    FrFormat(x).Top = 360
    If x <> Fr Then
        FrFormat(x).Left = 5500
    Else
        FrFormat(x).Left = 600
    End If
Next x
End Sub

Public Sub RetroFixLevels()
'The user wants to go back one. how inconvenient !
cmdNext.Caption = "Next >>"

Select Case level
Case 1
    MoveFrames 0
Case 2
    If Option1.Value = True Then
        MoveFrames 1
    Else
        MoveFrames 3
    End If
Case 3
    If Option1.Value = True Then
        Label7.Visible = False
        Combo3.Visible = False
        MoveFrames 2
    Else
        If Option9.Value = True Then
            MoveFrames 5
        Else
            MoveFrames 4
        End If
    End If
Case 4
        If Option11.Value = True Then
            MoveFrames 7
        Else
            MoveFrames 6
        End If
Case 5
        If Option9.Value = True Then MoveFrames 2
        Label7.Visible = True
        Combo3.Visible = True

End Select
End Sub

Public Sub DrawSummary()
'Ok they think they've finished
'Let's show them what they've settled on
Dim temp As String, temp1 As String, x As Integer
For x = 0 To 6
    LblAction(x).Caption = ""
Next x
LblAction(6).ForeColor = vbBlack
If Option1.Value = True Then
    LblAction(0) = "Action : Highlight"
    If Len(rtfHighLightString.Text) > 20 Then
        LblAction(1) = "Text : " + Left(rtfHighLightString.Text, 20) + "..."
    Else
        LblAction(1) = "Text : " + rtfHighLightString.Text
    End If
    LblAction(2) = "Font : " + Combo1.Text
    LblAction(3) = "Size : " + Combo2.Text
    If Option5.Value = True Then temp = "Regular"
    If Option6.Value = True Then temp = "Bold"
    If Option7.Value = True Then temp = "Italic"
    If Option8.Value = True Then temp = "Bold/Italic"
    temp1 = ""
    If Check1.Value = 1 Then temp1 = " Underline"
    If Check2.Value = 1 Then temp1 = " Strikethrough"
    If Check2.Value = 1 And Check1.Value = 1 Then
        temp1 = " Underline and Strikethrough"
    End If
    LblAction(4) = "Style : " + temp + temp1
    LblAction(5).ForeColor = cmdColor.BackColor
    LblAction(5) = "This Color"
Else
    If Option2.Value = True Then
        LblAction(0) = "Action : Insert Header"
    ElseIf Option3.Value = True Then
        LblAction(0) = "Action : Insert Footer"
    ElseIf Option4.Value = True Then
        LblAction(0) = "Action : Insert at Cursor"
    End If
    If Option9.Value = True Then
        LblAction(1) = "Insertion Object : Text"
        If Option11.Value = True Then
            If Len(rtfInsertText.Text) > 20 Then
                LblAction(2) = "Text : " + Left(rtfInsertText.Text, 20) + "..."
            Else
                LblAction(2) = "Text : " + rtfInsertText.Text
            End If
        Else
            LblAction(2) = "From file : " + LabelEdit(Text2.Text, 50)
        End If
        LblAction(3) = "Font : " + Combo1.Text
        LblAction(4) = "Size : " + Combo2.Text
        LblAction(5) = "Alignment : " + Combo3.Text
        If Option5.Value = True Then temp = "Regular"
        If Option6.Value = True Then temp = "Bold"
        If Option7.Value = True Then temp = "Italic"
        If Option8.Value = True Then temp = "Bold/Italic"
        temp1 = ""
        If Check1.Value = 1 Then temp1 = " Underline"
        If Check2.Value = 1 Then temp1 = " Strikethrough"
        If Check2.Value = 1 And Check1.Value = 1 Then
            temp1 = " Underline and Strikethrough"
        End If
        LblAction(6) = "Style : " + temp + temp1
    Else
        LblAction(1) = "Insertion Object : Image"
        LblAction(2) = "From file : " + LabelEdit(Text1.Text, 50)
        LblAction(3) = "Alignment : " + Combo4.Text
    End If
End If
End Sub

Public Sub WindUp()
'Now to put their choices into action
Dim textfound As Long
Dim Position As Long
Dim curX As Long
If Option1.Value = True Then
'Highlighting text
    curX = Form1.rtftext.SelStart 'remember where the cursor is
    LockWindowUpdate Form1.rtftext.hWnd
    textfound = Form1.rtftext.Find(rtfHighLightString.Text, 0)
    If textfound <> -1 Then 'OK we've found it - make it pretty
        If Option8.Value = False Then
            Form1.rtftext.SelBold = Option6.Value
            Form1.rtftext.SelItalic = Option7.Value
        Else
            Form1.rtftext.SelBold = Option8.Value
            Form1.rtftext.SelItalic = Option8.Value
        End If
        Form1.rtftext.SelUnderline = Check1.Value
        Form1.rtftext.SelStrikeThru = Check2.Value
        Form1.rtftext.SelFontName = Combo1.Text
        Form1.rtftext.SelFontSize = Val(Combo2.Text)
        Form1.rtftext.SelColor = cmdColor.BackColor
        Position = Form1.rtftext.SelStart + Form1.rtftext.SelLength + 1
        Do Until Position >= Len(Form1.rtftext.Text) - Len(rtfHighLightString.Text) + 1
            DoEvents
            If Check1.Value Then
                textfound = Form1.rtftext.Find(rtfHighLightString.Text, Position, , rtfMatchCase)
            Else
                textfound = Form1.rtftext.Find(rtfHighLightString.Text, Position)
            End If
            If textfound <> -1 Then 'Found some more - keep going
                Position = Form1.rtftext.SelStart + Form1.rtftext.SelLength + 1
                If Option8.Value = False Then
                    Form1.rtftext.SelBold = Option6.Value
                    Form1.rtftext.SelItalic = Option7.Value
                Else
                    Form1.rtftext.SelBold = Option8.Value
                    Form1.rtftext.SelItalic = Option8.Value
                End If
                Form1.rtftext.SelUnderline = Check1.Value
                Form1.rtftext.SelStrikeThru = Check2.Value
                Form1.rtftext.SelFontName = Combo1.Text
                Form1.rtftext.SelFontSize = Val(Combo2.Text)
                Form1.rtftext.SelColor = cmdColor.BackColor
            Else
                Exit Do
            End If
        Loop
    Else
        MsgBox "Cant find it !!", vbExclamation, "Bobo Enterprises"
        'Idiots - they didn't even have the text they told us to find !
    End If
Else
    'The user wants to insert something
    If Option9.Value = True Then
        'text
        If Option12.Value = True Then rtfInsertText.LoadFile Text2.Text
        If Option2.Value = True Then 'as header
            Form1.rtftext.Text = rtfInsertText.Text + vbCrLf + Form1.rtftext.Text
            Form1.rtftext.SelStart = 0
            Form1.rtftext.SelLength = Len(rtfInsertText.Text)
        ElseIf Option3.Value = True Then 'as footer
            Form1.rtftext.Text = Form1.rtftext.Text + vbCrLf + rtfInsertText.Text
            Form1.rtftext.SelStart = Len(Form1.rtftext.Text) - Len(rtfInsertText.Text)
            Form1.rtftext.SelLength = Len(rtfInsertText.Text)
        ElseIf Option4.Value = True Then 'at the cursor
            curX = Form1.rtftext.SelStart
            Form1.rtftext.SelStart = 0
            Form1.rtftext.SelLength = curX
            rtfInsertText.Text = Form1.rtftext.SelText + vbCrLf + rtfInsertText.Text
            Form1.rtftext.SelText = rtfInsertText.Text
            Form1.rtftext.SelStart = curX
            Form1.rtftext.SelLength = Len(rtfInsertText.Text) - curX
        End If
        If Option8.Value = False Then
            Form1.rtftext.SelBold = Option6.Value
            Form1.rtftext.SelItalic = Option7.Value
        Else
            Form1.rtftext.SelBold = Option8.Value
            Form1.rtftext.SelItalic = Option8.Value
        End If
        Form1.rtftext.SelUnderline = Check1.Value
        Form1.rtftext.SelStrikeThru = Check2.Value
        Form1.rtftext.SelFontName = Combo1.Text
        Form1.rtftext.SelFontSize = Val(Combo2.Text)
        Form1.rtftext.SelColor = cmdColor.BackColor
        Select Case Combo3.ListIndex
            Case 0
                Form1.rtftext.SelAlignment = rtfLeft
            Case 1
                Form1.rtftext.SelAlignment = rtfCenter
            Case 2
                Form1.rtftext.SelAlignment = rtfRight
        End Select

    ElseIf Option10.Value = True Then ' an Image file
        Clipboard.Clear
        Clipboard.SetData Picture1.Image, vbCFBitmap
        If Option2.Value = True Then 'as header
            Form1.rtftext.Text = vbCrLf + Form1.rtftext.Text
            Form1.rtftext.SelStart = 0
            SendMessage Form1.rtftext.hWnd, WM_PASTE, 0, 0
            Form1.rtftext.SelLength = 1
        ElseIf Option3.Value = True Then 'as footer
            Form1.rtftext.Text = Form1.rtftext.Text + vbCrLf
            Form1.rtftext.SelStart = Len(Form1.rtftext.Text)
            SendMessage Form1.rtftext.hWnd, WM_PASTE, 0, 0
            Form1.rtftext.SelLength = 1
        Else 'at the cursor
            SendMessage Form1.rtftext.hWnd, WM_PASTE, 0, 0
            Form1.rtftext.SelStart = Form1.rtftext.SelStart - 1
            Form1.rtftext.SelLength = 1
        End If
        Select Case Combo4.ListIndex
            Case 0
                Form1.rtftext.SelAlignment = rtfLeft
            Case 1
                Form1.rtftext.SelAlignment = rtfCenter
            Case 2
                Form1.rtftext.SelAlignment = rtfRight
        End Select
    End If
End If
Form1.rtftext.SelStart = curX
Form1.rtftext.SelLength = 0
LockWindowUpdate 0
Form1.Backup
Unload Me
End Sub


Public Sub SizeThumb()
'Just gives them a thumbnail of their selected picture file
Dim imgRatio As Double
If Picture1.Width > Picture1.Height Then
    imgRatio = PicThumb.Width / Picture1.Width
Else
    imgRatio = PicThumb.Height / Picture1.Height
End If
ImgThumb.Height = Picture1.Height * imgRatio
ImgThumb.Width = Picture1.Width * imgRatio
ImgThumb.Picture = Picture1.Image
End Sub
