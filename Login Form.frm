VERSION 5.00
Begin VB.Form LoginForm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Login  Form"
   ClientHeight    =   7455
   ClientLeft      =   9315
   ClientTop       =   2145
   ClientWidth     =   7515
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CreateUser 
      Caption         =   "Create User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "IF YOU ARE NEW USER  THEN CLICK HERE"
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton Forget 
      Caption         =   "Forget"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5640
      TabIndex        =   6
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4680
      TabIndex        =   5
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Submit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   4
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   3
      ToolTipText     =   "Enter password"
      Top             =   4080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   2
      ToolTipText     =   "Enter userId"
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   3000
      Picture         =   "Login Form.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2730
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Id"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3240
      Width           =   1050
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim u As String, p As String

Private Sub CreateUser_Click()
USERENTRY.Show
End Sub

Private Sub Exit_Click()
Unload Me
End
End Sub
Private Sub Forget_Click()
PasswordReset.Show
End Sub

Private Sub Form_Load()
conn
End Sub


Private Sub NewUsers_Click()

End Sub

Private Sub Submit_Click()
If u = Text1.Text And p = Text2.Text Then
Unload Me
frmSplash.Show
Else
MsgBox "wrong user or passw"
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error GoTo err
If KeyAscii = 13 Then
Text1.Text = UCase(Text1.Text)
sql = "select * from login where user_id='" + Text1.Text + "'"
Set r = c.Execute(sql)
u = r.Fields(0)
p = r.Fields(2)
Text2.SetFocus
Exit Sub
err:
MsgBox "NO DATA FOUND"
Text1.Text = ""
Text1.SetFocus
End If
End Sub



Private Sub Text2_GotFocus()
sql = "select * from login where user_id='" + Text1.Text + "'"
Set r = c.Execute(sql)
u = r.Fields(0)
p = r.Fields(2)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.Text = UCase(Text2.Text)
Submit.SetFocus
End If
End Sub

