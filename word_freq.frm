VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form mainwindow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "word frequency analyser"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "word_freq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8010
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Search 
      Caption         =   "Search"
      Height          =   255
      Left            =   3720
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   855
   End
   Begin ComctlLib.ProgressBar prog 
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   4200
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Resetbutton 
      Caption         =   "Reset"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton lengthbutton 
      Caption         =   "Sort by Length"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton freqbutton 
      Caption         =   "Sort by Freq"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton parsebutton 
      Caption         =   "Parse"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox main 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      HideSelection   =   0   'False
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "word_freq.frx":0442
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "mainwindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const reset_Text = "Double click in this textbox to paste contents of clipboard to this window."
Dim words(1 To 1500) As String      '1500 possible unique words
Dim wordcount(1 To 1500) As Integer
Dim numwords As Integer

Private Sub Form_Load()             'hide the progress bar
    prog.Visible = False
    main.Text = reset_Text
End Sub

Private Sub main_DblClick()
    main.Text = Clipboard.GetText
End Sub

Private Sub parsebutton_Click()     'reveal the hidden controls
    freqbutton.Visible = True
    lengthbutton.Visible = True
    Search.Visible = True
    parsebutton.Visible = False
    Resetbutton.Visible = True
    Call parse(main.Text)
End Sub
Private Sub freqbutton_Click()
    Call sort_by_frequency
    Call showcount
End Sub
Private Sub lengthbutton_Click()
    Call sort_by_length
    Call showcount
End Sub

Private Sub Resetbutton_Click()
    Call reset
End Sub

Private Sub parse(inText As String)
'This code parses a string into individual words,
'and then calls another function that either
'adds that word to the words() array if it hasn't been
'added yet, or increases the frequency counter wordcount()
'for that word if it has.  I wasn't aware of
'the split() function at the time i wrote this; i use chr()
'to filter out punctuation marks and the like,
'hyphenated words ["-" = chr(45)] are considered two words at this
'time...apostrophes are considered a part of a word
'due to contractions like "don't" and "I'm"

Dim inWord As Boolean   'are we "in a word"?
Dim b As String * 1     'single character string
Dim Word As String
Dim t As Integer        'counter

    inText = inText + vbCrLf  'add a carriage return to the end
    numwords = 1
    inWord = False
    For t = 1 To Len(inText)
        b = Mid(inText, t, 1)
        Select Case b
            Case Chr(10)
            Case Chr(32) To Chr(38), Chr(40) To Chr(47), Chr(58) To Chr(64), _
            Chr(91) To Chr(96), Chr(123) To Chr(127), Chr(13)
                If inWord = True Then
                    inWord = False
                    Call tally(Word)
                End If
            Case Chr(48) To Chr(57), Chr(65) To Chr(90), Chr(39), _
            Chr(97) To Chr(122)
                If inWord = False Then
                    Word = ""
                    inWord = True
                End If
                Word = Word + b
            Case Else
                If inWord = True Then
                    inWord = False
                    Call tally(Word)
                End If
        End Select
    Next
    numwords = numwords - 1
    Call showcount
End Sub

Private Sub tally(Word As String)
'this function determines whether or not the word it's given
'has been entered into the array yet.  if it has, it increases
'the frequency array, otherwise it adds it and sets the frequency
'to one
Dim t As Integer        'counter
    
    For t = 1 To numwords - 1
        If UCase(Word) = UCase(words(t)) Then
            wordcount(t) = wordcount(t) + 1
            Exit Sub
        End If
    Next
    words(numwords) = Word
    wordcount(numwords) = 1
    numwords = numwords + 1
End Sub

Private Sub showcount()
'this sub generates a report in the textbox and hides
'some of the controls from the user
Dim t As Integer
Dim a As String     'all purpose dummy variable
main.Visible = False
main.Visible = False
prog.Visible = True
prog.Value = 0
main.Text = "      # Word                                 Frequency    Length " + vbCrLf + _
            "=======================================================================" + vbCrLf
    
    For t = 1 To numwords
        prog.Value = Int((t / numwords) * 100)
        a = Trim(Str(t))
        main = main + String(7 - Len(a), 32) + a + " "
        a = words(t)
        main = main + a + String(30 - Len(a), 32)
        a = Trim(Str(wordcount(t)))
        main = main + String(16 - Len(a), 32) + a + " "
        a = Trim(Str(Len(words(t))))
        main = main + String(9 - Len(a), 32) + a + vbCrLf
    Next
    prog.Visible = False
    prog.Value = 0
    main.Visible = True
End Sub

Private Sub sort_by_frequency()
'this sub uses a bubble sort to sort by frequency.
Dim t As Integer    'counter
Dim sorted As Boolean 'if sorted=true at the end of the for..loop, it's sorted
    Do
        sorted = True
        For t = numwords To 2 Step -1
            If wordcount(t) > wordcount(t - 1) Then
                swap t, t - 1
                sorted = False
            End If
        Next
    Loop Until sorted = True
End Sub

Private Sub swap(a As Integer, b As Integer)
'this sub is used by the bubble sort algorithms to
'swap values in the words and wordcount arrays

Dim swap_word As String  'the obligatory third variable
Dim swap_freq As String  'necessary for swapping in VB

    swap_word = words(a)
    swap_freq = wordcount(a)
    words(a) = words(b)
    wordcount(a) = wordcount(b)
    words(b) = swap_word
    wordcount(b) = swap_freq
End Sub


Private Sub sort_by_length()
'this sub uses a bubble sort to sort by highest
'word lengths first.
Dim t As Integer    'counter
Dim sorted As Boolean

    Do
    sorted = True
        For t = numwords To 2 Step -1
            If Len(words(t)) > Len(words(t - 1)) Then
                swap t, t - 1
                sorted = False
            End If
        Next
    Loop Until sorted = True
End Sub



Private Sub reset()
'resets all variables and hides controls
Dim t As Integer
    
    main.Text = reset_Text
    Resetbutton.Visible = False
    numwords = 1
    For t = 1 To 1500
        words(t) = ""
        wordcount(t) = 0
    Next
    parsebutton.Visible = True
    Search.Visible = False
    freqbutton.Visible = False
    lengthbutton.Visible = False
    End Sub

Private Sub Search_Click()
'lets you search for a string in the textbox and selects it
Dim a As String
a = InputBox("Enter search string:")
If InStr(UCase(main.Text), UCase(a)) > 0 Then
    main.SelStart = InStr(UCase(main.Text), UCase(a)) - 1
    main.SelLength = Len(a)
    Else
    MsgBox "Not found."
End If
End Sub
