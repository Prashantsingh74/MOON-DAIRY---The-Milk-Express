VERSION 5.00
Begin VB.Form PasswordReset 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Reset"
   ClientHeight    =   3270
   ClientLeft      =   8115
   ClientTop       =   2130
   ClientWidth     =   7560
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "RESET PASSWORD"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ph No."
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Id"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "PasswordReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As String, B As String
Dim X As String
Dim Y As String
Private Sub CancelButton_Click()
Unload Me
Form1.Show
End Sub

Private Sub Form_Load()
conn
End Sub
Private Sub OKButton_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "ENTER ALL FEILDS"
Text1.SetFocus
Exit Sub
End If
If a = Text1.Text And B = Text2.Text Then
X = InputBox("Enter New Password", "For x sub")
Y = InputBox("Enter Confirm Password", "For y sub")
Else
MsgBox "NOT MATCH"
Text2.SetFocus
Exit Sub
End If
If X = "" Or Y = "" Then
MsgBox "Enter value"
Exit Sub
End If
If X = Y Then

sql = "update login set password='" + UCase(Y) + "' where USER_ID = '" + a + "'"
Set r = c.Execute(sql)
MsgBox "Password Reset Successfully"
Else
MsgBox "New Password & Confirm Password Unmatched"
End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo ccc
If KeyAscii = 13 Then
Text1.Text = UCase(Text1.Text)
sql = "select * from login where USER_ID='" + Text1.Text + "'"
Set r = c.Execute(sql)
a = r.Fields(0)
B = r.Fields(3)
Text2.SetFocus
Exit Sub
ccc:
MsgBox "NO DATA FOUND"
Text1.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
OKButton.SetFocus
End If
End Sub
