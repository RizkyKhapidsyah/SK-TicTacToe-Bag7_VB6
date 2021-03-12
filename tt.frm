VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   5370
   ClientTop       =   2040
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4890
   Begin VB.OptionButton Option2 
      Caption         =   "Advanced"
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   3480
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Beginner"
      Height          =   255
      Left            =   960
      TabIndex        =   22
      Top             =   3480
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000004&
      Height          =   285
      Left            =   4440
      TabIndex        =   16
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   4440
      TabIndex        =   15
      Text            =   "0"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   4440
      TabIndex        =   14
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   4440
      TabIndex        =   13
      Text            =   "1"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reset all"
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   615
      Left            =   3960
      TabIndex        =   11
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Again?"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   9
      Left            =   1920
      TabIndex        =   9
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   7
      Left            =   720
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Tic Tac Toe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   960
      TabIndex        =   21
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Games Tied"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Wins by O"
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Wins by X"
      Height          =   255
      Left            =   3600
      TabIndex        =   18
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Game #"
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   1560
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim owin As Integer
Dim xwin As Integer
Dim gnum As Integer
Dim tiedg As Integer
Dim moves As Integer
Dim xoro As String
Dim x As String
Dim comp As Integer
Dim kswitch As Integer
Dim diff As Integer
Dim s As Boolean
Dim done As Integer
Dim square As Integer
Dim y As Boolean
Dim yn As String
Dim dblscore As Integer
Private Function cdone()
        done = rowcheck()
        If (done <> -1) Then
            dblscore = dblscore + 1
            Call printmsg("row", done)
        End If
        done = colcheck()
        If (done <> -1) Then
            dblscore = dblscore + 1
            Call printmsg("col", done)
        End If
        done = diagcheck()
        If (done <> -1) Then
            dblscore = dblscore + 1
            Call printmsg("diag", done)
        End If
End Function
Private Function gquit()
Dim ii As Integer
      For ii = 0 To 9
        Text1(ii).Enabled = False
      Next ii
        Call win
End Function
Private Function movesf()
    If (moves = 10) Then
        tiedg = tiedg + 1
        MsgBox "The game is a tie."
        Text4.Text = Str(tiedg)
        Call gquit
    End If
End Function
Private Function win()
moves = 0
If (dblscore = 1) Then
    If (xoro = "X") Then
        xwin = xwin + 1
        Text2.Text = Str(xwin)
    Else
        owin = owin + 1
        Text3.Text = Str(owin)
    End If
Else
    If (xoro = "X") Then
        xwin = xwin
        Text2.Text = Str(xwin)
    Else
        owin = owin
        Text3.Text = Str(owin)
    End If
End If
    kswitch = 1
End Function

Private Function rowcheck()
  Dim a As Boolean
  Dim ii As Integer
  rowcheck = -1
  For ii = 0 To 6 Step 3
    a = (Text1(ii).Text = Text1(ii + 1).Text) And (Text1(ii + 1).Text = Text1(ii + 2).Text)
    If (a = True And (Text1(ii).Text <> "")) Then
      Text1(ii).BackColor = vbYellow
      Text1(ii + 1).BackColor = vbYellow
      Text1(ii + 2).BackColor = vbYellow
      rowcheck = ii
    End If
  Next ii
  a = (Text1(7).Text = Text1(8).Text) And (Text1(8).Text = Text1(9).Text)
  If (a And Text1(7).Text <> "") Then
    Text1(7).BackColor = vbYellow
    Text1(8).BackColor = vbYellow
    Text1(9).BackColor = vbYellow
    rowcheck = 12
  End If
End Function
Private Function colcheck()
  Dim a As Boolean
  Dim ii As Integer
  colcheck = -1
  For ii = 0 To 2 Step 1
    a = (Text1(ii).Text = Text1(ii + 3).Text) And (Text1(ii + 3).Text = Text1(ii + 6).Text)
    If (a = True And (Text1(ii).Text <> "")) Then
      Text1(ii).BackColor = vbYellow
      Text1(ii + 3).BackColor = vbYellow
      Text1(ii + 6).BackColor = vbYellow
      colcheck = ii
    End If
  Next ii
End Function
Private Function diagcheck()
  Dim a As Boolean
  diagcheck = -1
  a = (Text1(0).Text = Text1(4).Text) And (Text1(4).Text = Text1(8).Text)
  If (a = True And (Text1(0).Text <> "")) Then
    Text1(0).BackColor = vbYellow
    Text1(4).BackColor = vbYellow
    Text1(8).BackColor = vbYellow
    diagcheck = 0
  End If
  a = (Text1(2).Text = Text1(4).Text) And (Text1(4).Text = Text1(6).Text)
  If (a = True And (Text1(2).Text <> "")) Then
    Text1(2).BackColor = vbYellow
    Text1(4).BackColor = vbYellow
    Text1(6).BackColor = vbYellow
    diagcheck = 2
  End If
  a = (Text1(1).Text = Text1(5).Text) And (Text1(5).Text = Text1(9).Text)
  If (a = True And (Text1(1).Text <> "")) Then
    Text1(1).BackColor = vbYellow
    Text1(5).BackColor = vbYellow
    Text1(9).BackColor = vbYellow
    diagcheck = 1
  End If
End Function
Private Function printmsg(x As String, done As Integer)
  Select Case x
    Case "row"
      If (done < 12) Then
        xoro = Text1(done).Text
      Else
        xoro = Text1(7).Text
      End If
      MsgBox xoro & " has won by row " & Str(done / 3)
      Call gquit
    Case "col"
      xoro = Text1(done).Text
      MsgBox xoro & " has won by col " & Str(done)
      Call gquit
    Case "diag"
      xoro = Text1(done).Text
      MsgBox xoro & " has won by diagonal " & Str(done)
      Call gquit
   End Select
End Function

Private Sub Command1_Click()
Dim ii As Integer
dblscore = 0
kswitch = 0
moves = 0
gnum = gnum + 1
Text5.Text = gnum
For ii = 0 To 9
    Text1(ii).Text = ""
    Text1(ii).BackColor = vbGreen
    Text1(ii).Enabled = True
Next ii
  xoro = InputBox("who goes first X or O?", "tic tac toe", "X")
  xoro = UCase(xoro)
  Do While (xoro <> "X") And (xoro <> "O")
    MsgBox "sorry, your choices are X or O"
    xoro = InputBox("who goes first X or O?", , "X")
  Loop
    y = (yn = 1) And (xoro = "O")
    If (y = True) And (moves = 0) Then
        comp = 1
        moves = moves + 1
        Text1(4).Text = "O"
        Text1(4).BackColor = vbRed
        xoro = "X"
    Else
        If (yn = 1) Then
          comp = 1
          Option1.Enabled = True
          Option2.Enabled = True
        Else
          comp = 0
          Option1.Enabled = False
          Option2.Enabled = False
        End If
    End If
End Sub

Private Sub Command2_Click()
Dim ii As Integer
Text2.Text = "0"
Text3.Text = "0"
Text4.Text = "0"
Text5.Text = "0"
For ii = 0 To 9
    Text1(ii).Text = ""
    Text1(ii).BackColor = vbGreen
Next ii
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
'fix that damn "X" bug
Form1.Caption = "Tic Tac Toe"
Form1.Show
  dblscore = 0
  diff = 1
  comp = 0
  kswitch = 0
  gnum = 1
  moves = 0
  xwin = 0
  owin = 0
  tiedg = 0
  yn = InputBox("Select number of players.", "Player Number", "1")
  Do While (yn <> "1") And (yn <> "2")
    MsgBox "Sorry, must be 1 or 2."
    yn = InputBox("Number or players?", "Player Number", "1")
  Loop
  xoro = InputBox("who goes first X or O?", "tic tac toe", "X")
  xoro = UCase(xoro)
  Do While (xoro <> "X") And (xoro <> "O")
    MsgBox "Sorry, your choices are X or O"
    xoro = InputBox("Who goes first X or O?", "tic tac toe", "X")
  Loop
    y = (yn = 1) And (xoro = "O")
    If (y = True) And (moves = 0) Then
        comp = 1
        moves = moves + 1
        Text1(4).Text = "O"
        Text1(4).BackColor = vbRed
        xoro = "X"
    Else
        If (yn = 1) Then
          comp = 1
          Option1.Enabled = True
          Option2.Enabled = True
        Else
          comp = 0
          Option1.Enabled = False
          Option2.Enabled = False
        End If
    End If
End Sub

Private Sub Option1_Click()
diff = 1
End Sub

Private Sub Option2_Click()
diff = 2
End Sub

Private Sub Text1_Click(Index As Integer)
Dim ii As Integer
If (comp = 0) Then
  moves = moves + 1
  If (Text1(Index).Text = "") Then
    Text1(Index).Text = xoro
    Text1(Index).BackColor = vbRed
    Call cdone
    If xoro = "X" Then
      xoro = "O"
    Else
      xoro = "X"
    End If
  Else
    MsgBox "Sorry, already selected."
    moves = moves - 1
  End If
End If
If (comp = 1) Then
    moves = moves + 1
    If xoro = "X" And Text1(Index).Text = "" Then
        If (moves = 10) Then
            Text1(Index).Text = "X"
            Text1(Index).BackColor = vbRed
            kswitch = 1
        Else
            Text1(Index).Text = "X"
            Text1(Index).BackColor = vbRed
        End If
        Call cdone
        If kswitch = 1 Then
            Call movesf
            xoro = "X"
            Exit Sub
        Else
            xoro = "O"
        End If
    Else
        MsgBox "Sorry, already selected."
        moves = moves - 1
        xoro = "X"
        Exit Sub
    End If
End If
s = (comp = 1) And (xoro = "O")
If (s = True) And (diff >= 1) Then
Dim i As Integer
Dim a As Boolean
If (diff = 2) Then
    For i = 0 To 6 Step 3
        a = (Text1(i).Text = "") And (Text1(i + 1).Text = Text1(i + 2).Text)
        If (a = True) And (Text1(i + 1).Text = "X" Or Text1(i + 1).Text = "O") Then
            Text1(i).Text = "O"
            Text1(i).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
        End If
    Next i
    For i = 0 To 6 Step 3
        a = (Text1(i + 1).Text = "") And (Text1(i).Text = Text1(i + 2).Text)
        If (a = True) And (Text1(i).Text = "X" Or Text1(i).Text = "O") Then
            Text1(i + 1).Text = "O"
            Text1(i + 1).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
        End If
    Next i
    For i = 0 To 6 Step 3
        a = (Text1(i + 2).Text = "") And (Text1(i).Text = Text1(i + 1).Text)
        If (a = True) And (Text1(i).Text = "X" Or Text1(i).Text = "O") Then
            Text1(i + 2).Text = "O"
            Text1(i + 2).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
        End If
    Next i
    For i = 0 To 1 Step 1
        a = (Text1(i).Text = "") And (Text1(i + 4).Text = Text1(i + 8).Text)
        If (a = True) And (Text1(i + 4).Text = "X" Or Text1(i + 4).Text = "O") Then
            Text1(i).Text = "O"
            Text1(i).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
        End If
    Next i
    For i = 0 To 1 Step 1
        a = (Text1(i + 8).Text = "") And (Text1(i).Text = Text1(i + 4).Text)
        If (a = True) And (Text1(i).Text = "X" Or Text1(i).Text = "O") Then
            Text1(i + 8).Text = "O"
            Text1(i + 8).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
        End If
    Next i
    For i = 0 To 1 Step 1
        a = (Text1(i + 4).Text = "") And (Text1(i).Text = Text1(i + 8).Text)
        If (a = True) And (Text1(i).Text = "X" Or Text1(i).Text = "O") Then
            Text1(i + 4).Text = "O"
            Text1(i + 4).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
        End If
    Next i
    For i = 0 To 2 Step 1
        a = (Text1(i).Text = "") And (Text1(i + 3).Text = Text1(i + 6).Text)
        If (a = True) And (Text1(i + 3).Text = "X" Or Text1(i + 3).Text = "O") Then
            Text1(i).Text = "O"
            Text1(i).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
        End If
    Next i
    For i = 0 To 2 Step 1
        a = (Text1(i + 3).Text = "") And (Text1(i).Text = Text1(i + 6).Text)
        If (a = True) And (Text1(i).Text = "X" Or Text1(i).Text = "O") Then
            Text1(i + 3).Text = "O"
            Text1(i + 3).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
        End If
    Next i
    For i = 0 To 2 Step 1
        a = (Text1(i + 6).Text = "") And (Text1(i).Text = Text1(i + 3).Text)
        If (a = True) And (Text1(i).Text = "X" Or Text1(i).Text = "O") Then
            Text1(i + 6).Text = "O"
            Text1(i + 6).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
        End If
    Next i
    a = (Text1(2).Text = "") And (Text1(4).Text = Text1(6).Text)
    If (a = True) And (Text1(4).Text = "X" Or Text1(4).Text = "O") Then
            Text1(2).Text = "O"
            Text1(2).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
    End If
    a = (Text1(4).Text = "") And (Text1(2).Text = Text1(6).Text)
    If (a = True) And (Text1(2).Text = "X" Or Text1(2).Text = "O") Then
            Text1(4).Text = "O"
            Text1(4).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
    End If
    a = (Text1(6).Text = "") And (Text1(2).Text = Text1(4).Text)
    If (a = True) And (Text1(2).Text = "X" Or Text1(2).Text = "O") Then
          Text1(6).Text = "O"
            Text1(6).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
    End If
    a = (Text1(7).Text = "") And (Text1(8).Text = Text1(9).Text)
    If (a = True) And (Text1(8).Text = "X" Or Text1(8).Text = "O") Then
          Text1(7).Text = "O"
            Text1(7).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
    End If
    a = (Text1(8).Text = "") And (Text1(7).Text = Text1(9).Text)
    If (a = True) And (Text1(7).Text = "X" Or Text1(7).Text = "O") Then
          Text1(8).Text = "O"
            Text1(8).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
    End If
    a = (Text1(9).Text = "") And (Text1(7).Text = Text1(8).Text)
    If (a = True) And (Text1(7).Text = "X" Or Text1(7).Text = "O") Then
          Text1(9).Text = "O"
            Text1(9).BackColor = vbRed
            moves = moves + 1
            xoro = "X"
            Call cdone
            Call movesf
            Exit Sub
    End If
squ:
Randomize
square = Int(Rnd * 10)
    If s = True And (Text1(square).Text = "") Then
        moves = moves + 1
        Text1(square).Text = "O"
        Text1(square).BackColor = vbRed
        Call cdone
        Call movesf
        xoro = "X"
    Else
        GoTo squ
    End If
    Exit Sub
Else
sq:
Randomize
square = Int(Rnd * 10)
    If s = True And (Text1(square).Text = "") Then
        moves = moves + 1
        Text1(square).Text = "O"
        Text1(square).BackColor = vbRed
        Call cdone
        xoro = "X"
    Else
        GoTo sq
    End If
    xoro = "X"
    Call movesf
  End If
End If
End Sub
