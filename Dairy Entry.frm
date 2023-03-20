VERSION 5.00
Begin VB.Form DairyEntry 
   BackColor       =   &H00FF8080&
   Caption         =   "Dairy Entry"
   ClientHeight    =   6345
   ClientLeft      =   6870
   ClientTop       =   3735
   ClientWidth     =   9480
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   9480
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3960
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   3960
      TabIndex        =   10
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton exit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton save 
      Caption         =   "Save"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DAIRY ENTRY"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Licence No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dairy Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dairy Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1560
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1560
      TabIndex        =   3
      Top             =   4320
      Width           =   1455
   End
End
Attribute VB_Name = "DairyEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub save_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "All Fields Required"
Command2.SetFocus
Else
conn
sql = "insert into Dairy_Entry values('" + Text1.Text + "','" + Text2.Text + "'," + Text3.Text + ",'" + Text4.Text + "','" + Text5.Text + "')"
Set r = c.Execute(sql)
sql = "insert into Dairy_dues values('" + Text1.Text + "'," + Label6.Caption + ")"
Set r = c.Execute(sql)

MsgBox "Record Saved"
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub
Private Sub Exit_Click()
Unload Me
home.Show
End Sub

Private Sub Form_Load()
Text1.Enabled = False
save.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox ("Enter Character Only")
End If

If KeyAscii = 13 Then
Text2.Text = UCase(Text2.Text)
Text3.SetFocus
End If
End Sub

Private Sub text3_keypress(KeyAscii As Integer)
Text3.MaxLength = 10
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox ("Enter Number Only")
End If
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
Text1.Text = LCase(Text2.Text)
End Sub

Private Sub Text3_LostFocus()
If Len(Text3.Text) <> 10 Then
MsgBox "Enter 10 Digit"
Text5.SetFocus
End If
End Sub

Private Sub Text4_Click()
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub text4_keypress(KeyAscii As Integer)

If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
save.Enabled = True
If KeyAscii = 13 Then
save.SetFocus
End If
End Sub
