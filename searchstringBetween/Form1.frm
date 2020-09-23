VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Search string between string"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtstart 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   5055
   End
   Begin VB.TextBox txtEnd 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   2880
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtstring 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search from "
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Result :"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   3960
      Width           =   540
   End
   Begin VB.Label lblresult 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End at"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start from"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function GetStringBetween(ByVal Str As String, ByVal str1 As String, ByVal str2 As String, Optional ByVal st As Long = 0) As String
    On Error Resume Next
    Dim s1, s2, s, l As Long
    Dim foundstr As String
    
    s1 = InStr(st + 1, Str, str1, vbTextCompare)
    s2 = InStr(s1 + 1, Str, str2, vbTextCompare)
    
    If s1 = 0 Or s2 = 0 Or IsNull(s1) Or IsNull(s2) Then
        foundstr = Str
    Else
        s = s1 + Len(str1)
        l = s2 - s
        foundstr = Mid(Str, s, l)
    End If
    
    GetStringBetween = foundstr
End Function

Private Sub Command1_Click()
Dim strtest As String

strtest = GetStringBetween(txtstring.Text, txtstart.Text, txtEnd.Text, 0)
lblresult.Caption = strtest

End Sub

Private Sub Form_Activate()
lblresult.ForeColor = RGB(250, 250, 210)
Me.BackColor = RGB(255, 140, 0)
txtEnd.BackColor = RGB(255, 160, 122)
txtstart.BackColor = RGB(255, 160, 122)
txtstring.BackColor = RGB(255, 160, 122)
End Sub

Private Sub Form_Load()
txtstring.Text = "<p><center><b>Welcome Sahir kazi!.<br>Your A/C type =  Startup Rs. 100 Validity 90 Days Pack.<br><br>Your A/c Balance = Rs.97.75<br>Package Expiry Date = 2003-07-21<br><br>Thanks for logging.</b></center>"
txtstart.Text = "<b>"
txtEnd.Text = "<br>"
lblresult.Caption = ""
End Sub
