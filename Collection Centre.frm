VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CollectionCentre 
   BackColor       =   &H00FF8080&
   Caption         =   "Collection Centre"
   ClientHeight    =   5625
   ClientLeft      =   4605
   ClientTop       =   3735
   ClientWidth     =   12720
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
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   12720
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   3015
      Left            =   10800
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton clear 
         Caption         =   "Clear"
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Center Details"
      Height          =   4215
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   9255
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   855
         Left            =   4320
         TabIndex        =   5
         Top             =   3000
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1508
         _Version        =   393217
         TextRTF         =   $"Collection Centre.frx":0000
      End
      Begin VB.TextBox Text3 
         Height          =   480
         Left            =   4320
         TabIndex        =   4
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   4320
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   6960
         TabIndex        =   1
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   139984897
         CurrentDate     =   44973
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Address"
         Height          =   375
         Left            =   2400
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Phone No"
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Center Name"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Registration Date"
         Height          =   375
         Left            =   6960
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Registration No."
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COLLECTION CENTRE"
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   3000
      Width           =   375
   End
End
Attribute VB_Name = "CollectionCentre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Exit_Click()
Unload Me
home.Show
End Sub

Private Sub save_Click()
'If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or richtexbox1.Text = "" Then
'MsgBox " All Fields Required"
'Else
sql = "insert into collection_centre values(" + Text1.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "','" + Text2.Text + "'," + Text3.Text + ",'" + RichTextBox1.Text + "')"
Set r = c.Execute(sql)

sql = "INSERT INTO CENTRE_DUES VALUES('" + Text1.Text + "'," + Label6.Caption + ")"
Set r = c.Execute(sql)
MsgBox "Record Saved"
Text1.Text = ""
Text1.Text = ""
Text1.Text = ""
RichTextBox1.Text = ""
'End If
'save.Enabled = False
End Sub


Private Sub clear_Click()
Text1.Text = ""
Text1.Text = ""
Text1.Text = ""
RichTextBox1.Text = ""
End Sub

Private Sub Form_Load()
conn

End Sub

Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
RichTextBox1.Text = UCase(RichTextBox1.Text)
End If
End Sub

Private Sub RichTextBox1_LostFocus()
RichTextBox1.Text = UCase(RichTextBox1.Text)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If
If KeyAscii = 13 Then
DTPicker1.SetFocus
End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.Text = UCase(Text2.Text)
Text3.SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
Text2.Text = UCase(Text2.Text)
End Sub

Private Sub text3_keypress(KeyAscii As Integer)
Text3.MaxLength = 10
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or KeyAscii = 8 Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox (" enter numeric value only")
End If
If KeyAscii = 13 Then
RichTextBox1.SetFocus
End If
End Sub
