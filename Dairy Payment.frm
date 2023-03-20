VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DairyPayment 
   BackColor       =   &H00FF8080&
   Caption         =   "DairyPayment"
   ClientHeight    =   11835
   ClientLeft      =   4275
   ClientTop       =   1065
   ClientWidth     =   15315
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
   ScaleHeight     =   11835
   ScaleWidth      =   15315
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFC0C0&
      Height          =   3495
      Left            =   12720
      TabIndex        =   47
      Top             =   3360
      Width           =   1695
      Begin VB.CommandButton new 
         Caption         =   "New"
         Height          =   495
         Left            =   240
         TabIndex        =   57
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   240
         TabIndex        =   50
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton clear 
         Caption         =   "Clear"
         Height          =   495
         Left            =   240
         TabIndex        =   49
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   48
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Collection Details"
      Height          =   1695
      Left            =   360
      TabIndex        =   32
      Top             =   8160
      Width           =   12015
      Begin VB.TextBox Text6 
         Height          =   480
         Left            =   9960
         TabIndex        =   44
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   9960
         TabIndex        =   42
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   480
         Left            =   5400
         TabIndex        =   40
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   465
         Left            =   5400
         TabIndex        =   38
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   480
         Left            =   1920
         TabIndex        =   36
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1920
         TabIndex        =   34
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   375
         Left            =   7080
         TabIndex        =   43
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Milk Collection "
         Height          =   375
         Left            =   7080
         TabIndex        =   41
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Evening Amt"
         Height          =   375
         Left            =   3600
         TabIndex        =   39
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Evening Qty "
         Height          =   375
         Left            =   3600
         TabIndex        =   37
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Morning Amt"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   " Morning Qty"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Height          =   1575
      Left            =   360
      TabIndex        =   23
      Top             =   10080
      Width           =   8775
      Begin VB.TextBox Text10 
         Height          =   480
         Left            =   6360
         TabIndex        =   46
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   4080
         TabIndex        =   31
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   480
         Left            =   2280
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   480
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Dues"
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Dues"
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Left"
         Height          =   375
         Left            =   6360
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Recived"
         Height          =   375
         Left            =   4080
         TabIndex        =   25
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Evening"
      Height          =   4695
      Left            =   6480
      TabIndex        =   14
      Top             =   3240
      Width           =   5895
      Begin VB.ListBox List5 
         Height          =   3660
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.ListBox List6 
         Height          =   3660
         Left            =   2040
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.ListBox List7 
         Height          =   3660
         Left            =   3240
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List8 
         Height          =   3660
         Left            =   4440
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fat %"
         Height          =   375
         Left            =   3360
         TabIndex        =   20
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   375
         Left            =   4440
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Morning"
      Height          =   4695
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   6015
      Begin VB.ListBox List4 
         Height          =   3660
         Left            =   4560
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.ListBox List3 
         Height          =   3660
         Left            =   3360
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox List2 
         Height          =   3660
         Left            =   2040
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   3660
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   375
         Left            =   4560
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fat %"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2055
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   9855
      Begin VB.TextBox Text13 
         Height          =   480
         Left            =   8040
         TabIndex        =   56
         Text            =   "Text13"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   4800
         TabIndex        =   54
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Height          =   480
         Left            =   4800
         TabIndex        =   53
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton submit 
         Caption         =   "Submit"
         Height          =   495
         Left            =   7560
         TabIndex        =   24
         Top             =   1320
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   139919361
         CurrentDate     =   44982
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   139919361
         CurrentDate     =   44982
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No"
         Height          =   375
         Left            =   7080
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Dairy Name"
         Height          =   375
         Left            =   3000
         TabIndex        =   52
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Dairy Id"
         Height          =   375
         Left            =   3000
         TabIndex        =   51
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "to"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date Range"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DAIRY PAYMENT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   5640
      TabIndex        =   29
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "DairyPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_Click()
Unload Me
home.Show
End Sub

Private Sub Submit_Click()

sql = "select bill_date,qty,fat,total_amt from morning_dairysale where bill_date between '" + Format(DTPicker1.Value, "dd mmm yyyy") + "' and '" + Format(DTPicker2.Value, "dd mmm yyyy") + "'" ' and f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
List1.AddItem r.Fields(0)
List2.AddItem r.Fields(1)
List3.AddItem r.Fields(2)
List4.AddItem r.Fields(3)
r.MoveNext
Loop
sql = "select bill_date,qty,fat,total_amt from evening_dairysale where bill_date between '" + Format(DTPicker1.Value, "dd mmm yyyy") + "' and '" + Format(DTPicker2.Value, "dd mmm yyyy") + "'" 'and f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
List5.AddItem r.Fields(0)
List6.AddItem r.Fields(1)
List7.AddItem r.Fields(2)
List8.AddItem r.Fields(3)
r.MoveNext
Loop
For i = 0 To List2.ListCount Step 1
total = total + Val(List2.List(i))
Text1.Text = total
Next
For i = 0 To List4.ListCount Step 1
Sum = Sum + Val(List4.List(i))
Text2.Text = Sum
Next
For i = 0 To List6.ListCount Step 1
add = add + Val(List6.List(i))
Text3.Text = add
Next

For i = 0 To List8.ListCount Step 1
Addition = Addition + Val(List8.List(i))
Text4.Text = Addition
Next
sql = "select dues from dairy_dues"
Set r = c.Execute(sql)
Text7.Text = r.Fields(0)
Text5.Text = Val(Text1.Text) + Val(Text3.Text)
Text6.Text = Val(Text2.Text) + Val(Text4.Text)
Text8.Text = Val(Text6.Text) + Val(Text7.Text)
Text9.Locked = False
save.Enabled = True
clear.Enabled = True
End Sub

Private Sub save_Click()
If Text6.Text = "" Or Text9.Text = "" Or Text11.Text = "" Then
MsgBox " Enter All Fields"
Else
sql = "insert into dairy_payment values(" + Text13.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "','" + Format(DTPicker2.Value, "dd mmm yyyy") + "','" + Text11.Text + "'," + Text1.Text + "," + Text3.Text + "," + Text5.Text + "," + Text6.Text + "," + Text8.Text + "," + Text9.Text + "," + Text10.Text + ")"
Set r = c.Execute(sql)
MsgBox "record saved"
sql = " update dairy_dues set dues=" + Text10.Text + " where dairy_id='" + Text11.Text + "'"
Set r = c.Execute(sql)
MsgBox "dues Updated"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
DTPicker1.Enabled = False
DTPicker2.Enabled = False
Submit.Enabled = False
save.Enabled = False
clear.Enabled = False
End If
End Sub

Private Sub clear_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
End Sub

Private Sub new_Click()
sql = "select count(bill_no) from dairy_payment"
Set r = c.Execute(sql)
Text13.Text = r.Fields(0) + 1
sql = "select dairy_id from dairy_entry"
Set r = c.Execute(sql)
Text11.Text = r.Fields(0)
sql = "select dairy_nm from dairy_entry"
Set r = c.Execute(sql)
Text12.Text = r.Fields(0)
DTPicker1.Enabled = True
DTPicker2.Enabled = True
Submit.Enabled = True
Text13.Locked = True
End Sub

Private Sub Form_Load()
conn
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text10.Locked = True
Text11.Locked = True
Text12.Locked = True
DTPicker1.Enabled = False
DTPicker2.Enabled = False
Submit.Enabled = False
save.Enabled = False
clear.Enabled = False
End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox ("Enter Number Only")
End If
If KeyAscii = 13 Then
save.SetFocus
End If
End Sub

Private Sub Text9_LostFocus()
Text10.Text = Val(Text8.Text) - Val(Text9.Text)
End Sub
