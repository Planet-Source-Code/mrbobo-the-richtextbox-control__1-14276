VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find and Replace"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "All"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Replace"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox Replaceme 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.ComboBox Findme 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'Find
Dim h As Integer
Dim x As Integer
Dim textfound As Long
Dim Position As Long
'Update our combobox and avoid duplication
If Findme.Text <> "" Then
    h = 0
    For x = 0 To Findme.ListCount - 1
        If Findme.Text = Findme.List(x) Then h = 1
    Next x
    If h = 0 Then Findme.AddItem Findme.Text, (0)
Else
    MsgBox "Search for what ??"
    Exit Sub
End If
    If Check1.Value Then 'match case
        textfound = Form1.rtftext.Find(Findme.Text, 0, , rtfMatchCase)
    Else
        textfound = Form1.rtftext.Find(Findme.Text, 0)
    End If
    If textfound <> -1 Then 'Found it !
        Form1.rtftext.SetFocus
    Else 'No I didn't
        MsgBox "Cant find it !!", vbExclamation, "Bobo Enterprises"
    End If

End Sub

Private Sub Command2_Click()
Dim textfound As Long
Dim Position As Long
'Same as Find except we start at the cursor instead of
'the top of the page
Position = Form1.rtftext.SelStart + Form1.rtftext.SelLength + 1
If Check1.Value Then
    textfound = Form1.rtftext.Find(Findme.Text, Position, , rtfMatchCase)
Else
    textfound = Form1.rtftext.Find(Findme.Text, Position)
End If
If textfound <> -1 Then
    Form1.rtftext.SetFocus
Else
    MsgBox "Cant find it !!", vbExclamation, "Bobo Enterprises"
End If
End Sub

Private Sub Command3_Click()
Dim h As Integer
Dim x As Integer
'Same as Find except we replace it with Replaceme text
'if we find it
If Findme.Text <> "" Then
    h = 0
    For x = 0 To Findme.ListCount - 1
        If Findme.Text = Findme.List(x) Then h = 1
    Next x
    If h = 0 Then Findme.AddItem Findme.Text, (0)
    If Replaceme.Text <> "" Then
      h = 0
      For x = 0 To Replaceme.ListCount - 1
        If Replaceme.Text = Replaceme.List(x) Then h = 1
      Next x
      If h = 0 Then Replaceme.AddItem Replaceme.Text, (0)
      h = 0
      For x = 0 To Findme.ListCount - 1
          If Findme.Text = Findme.List(x) Then h = 1
      Next x
      If h = 0 Then Findme.AddItem Findme.Text, (0)
    End If
Else
    MsgBox "Nothing to replace.", vbExclamation, "Bobo Enterpises"
End If
If Form1.rtftext.SelText = "" Then Exit Sub
Form1.rtftext.SelText = Replaceme.Text

End Sub

Private Sub Command4_Click()
Dim textfound As Long
Dim Occurences As Long
Dim Position As Long
Dim curX As Integer
Dim x As Integer
Dim h As Integer
'Same as find except if we find it we replace it and continue
'doing so till we get to the bottom of the page
If Findme.Text <> "" Then
    h = 0
    For x = 0 To Findme.ListCount - 1
        If Findme.Text = Findme.List(x) Then h = 1
    Next x
    If h = 0 Then Findme.AddItem Findme.Text, (0)
      If Replaceme.Text <> "" Then
        h = 0
        For x = 0 To Replaceme.ListCount - 1
          If Replaceme.Text = Replaceme.List(x) Then h = 1
        Next x
        If h = 0 Then
          Replaceme.AddItem Replaceme.Text, (0)
        End If
      End If
Else
    MsgBox "Nothing to replace.", vbExclamation, "Bobo Enterpises"
End If
        curX = Form1.rtftext.SelStart
        Occurences = 0
        LockWindowUpdate Form1.rtftext.hWnd
        If Check1.Value Then
            textfound = Form1.rtftext.Find(Findme.Text, 0, , rtfMatchCase)
        Else
            textfound = Form1.rtftext.Find(Findme.Text, 0)
        End If
        If textfound <> -1 Then
            Occurences = Occurences + 1
            Position = Form1.rtftext.SelStart + Form1.rtftext.SelLength + 1
            Form1.rtftext.SelText = Replaceme.Text
            Do Until Position >= Len(Form1.rtftext.Text) - Len(Findme.Text) + 1
                DoEvents
                If Check1.Value Then
                    textfound = Form1.rtftext.Find(Findme.Text, Position, , rtfMatchCase)
                Else
                    textfound = Form1.rtftext.Find(Findme.Text, Position)
                End If
                If textfound <> -1 Then
                    Occurences = Occurences + 1
                    Position = Form1.rtftext.SelStart + Form1.rtftext.SelLength + 1
                    Form1.rtftext.SelText = Replaceme.Text
                Else
                    Exit Do
                End If
            Loop
        Else
            MsgBox "Cant find it !!", vbExclamation, "Bobo Enterprises"
        End If
Form1.Backup
Me.Caption = "Replaced " + Str(Occurences) + " occurences"
Form1.rtftext.SelStart = curX
LockWindowUpdate 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Dont unload unless the whole program ends to make the
'job of saving settings easier
If FinalClose = False Then
    Me.Visible = False
    Cancel = 1
End If

End Sub
