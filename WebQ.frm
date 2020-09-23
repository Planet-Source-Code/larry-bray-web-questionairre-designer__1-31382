VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "WebQ"
   ClientHeight    =   6585
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8025
   Icon            =   "WebQ.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   8025
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Height          =   495
      Left            =   2760
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4920
      TabIndex        =   17
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Write question"
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtNumOpt 
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   3840
      Width           =   615
   End
   Begin VB.OptionButton OptType 
      Caption         =   "Text area"
      Height          =   495
      Index           =   3
      Left            =   600
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton OptType 
      Caption         =   "Text box"
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton OptType 
      Caption         =   "Check box"
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   10
      Top             =   3600
      Width           =   1215
   End
   Begin VB.OptionButton OptType 
      Caption         =   "Radio button"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   9
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtQtext 
      Height          =   1095
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtQlab 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtHead 
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label labLast 
      Height          =   375
      Left            =   6600
      TabIndex        =   21
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Last question"
      Height          =   255
      Left            =   6600
      TabIndex        =   20
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "# Of Options"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Question Type"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Question Text"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Q Label"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Questionairre Heading"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Window Title"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Web Questionairre Design"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutItem 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
  Close #1
  Unload Form1
  End
End Sub

Private Sub cmdFinish_Click()
  Print #1, "<A NAME=term></A>"
  Print #1, "Thank you very much for your participation.</BODY><P>"
  Print #1, "<INPUT TYPE=""submit"" VALUE=""Submit Survey"">"
  Print #1, ""
  Print #1, ""
  Print #1, "</FORM>"
  Print #1, "<BR>"
  Print #1, "</BODY>"
  Print #1, "</HTML>"
End Sub

Private Sub cmdNext_Click()
 Qlab$ = txtQlab.Text
  Print #1, ""
  Print #1, "<B><A NAME=" & txtQlab.Text & ">" & txtQlab.Text & "</A>. </B>"
  Print #1, txtQtext.Text
  Print #1, " <P>"
  Print #1, "<BLOCKQUOTE>"
If OptType.Item(0).Value = True Then
  For i% = 1 To txtNumOpt.Text
    opt$ = ""
    optif$ = ""
    opt$ = InputBox("Enter option" & i% & " text", "Option string")
    optif$ = InputBox("Option" & i% & " goto label", "goto")
    If optif$ <> "" Then
      Print #1, "<INPUT TYPE=""radio"" NAME=" & txtQlab.Text & " value=" & i% & " Onclick=AutoJump('#" & optif$ & "')>" & opt$
    Else
      Print #1, "<INPUT TYPE=""radio"" NAME=" & txtQlab.Text & " value=" & i% & " >" & opt$
    End If
  Print #1, "<BR>"
  Next i%
End If
If OptType.Item(1).Value = True Then
  For i% = 1 To txtNumOpt.Text
    opt$ = ""
    optif$ = ""
    opt$ = InputBox("Enter option" & i% & " text", "Option string")
    Print #1, "<INPUT TYPE=""checkbox"" NAME=" & txtQlab.Text & " value=" & i% & " >" & opt$
  Print #1, "<BR>"
  Next i%
End If
If OptType.Item(2).Value = True Then
  Print #1, "<TABLE><TR><TD><INPUT TYPE=""Text"" NAME=" & txtQlab.Text & " SIZE=""34""></TD></TR>"
  Print #1, "</TABLE>"
End If
If OptType.Item(3).Value = True Then
  Print #1, "<TEXTAREA name=" & txtQlab.Text & " Rows=""3"" Cols=""50""></TEXTAREA>"
End If
  Print #1, "</BLOCKQUOTE>"
 If OptType.Item(1).Value = True Or OptType.Item(2).Value = True Or OptType.Item(3).Value = True Then
   nxt$ = InputBox("Enter next question label", "Next question")
   Print #1, ""
   Print #1, "<CENTER><INPUT name=btn" & txtQlab.Text & " type=button value=Next></CENTER>"
   Print #1, "<SCRIPT LANGUAGE=""VBScript"">"
   Print #1, "Sub btn" & txtQlab.Text & "_OnClick()"
   Print #1, ""
   Print #1, "location.href = ""#" & nxt$ & " "" "
   Print #1, ""
   Print #1, "End Sub"
   Print #1, "</SCRIPT>"
 End If
  Print #1, ""
  Print #1, "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"
  Print #1, "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"

txtQlab.Text = ""
txtQtext.Text = ""
OptType.Item(0).Value = False
OptType.Item(1).Value = False
OptType.Item(2).Value = False
OptType.Item(3).Value = False
txtNumOpt.Text = ""
labLast.Caption = Qlab$
Label7.Visible = False
txtNumOpt.Visible = False

End Sub

Private Sub cmdStart_Click()
  Title$ = txtTitle.Text
  Head$ = txtHead.Text
    
  Print #1, "<HTML><HEAD>"
  Print #1, "<TITLE>" & Title$ & "</TITLE>"
  Print #1, "<script><!--"
  Print #1, ""
  Print #1, "function AutoJump(destination) {"
  Print #1, " window.location = destination"
  Print #1, "}"
  Print #1, ""
  Print #1, "// --></script>"
  Print #1, "</HEAD>"
  Print #1, "<BODY bgcolor=""white"">"
  Print #1, ""
  Print #1, "<FORM action=mailto:aaaaa@bbbbb.com encType=text/plain"
  Print #1, "          method=post name=LayoutRegion1FORM>"
  Print #1, ""
  Print #1, "<CENTER><H1>" & Head$ & "</H1></CENTER>"
End Sub

Private Sub Form_Load()
Label7.Visible = False
txtNumOpt.Visible = False
z% = 0
Open "Webq.htm" For Append As #1
End Sub

Private Sub mnuAboutItem_Click()
  frmAbout.Visible = True
End Sub

Private Sub OptType_Click(Index As Integer)
If OptType.Item(0).Value = True Or OptType.Item(1).Value = True Then
  Label7.Visible = True
  txtNumOpt.Visible = True
End If
End Sub

