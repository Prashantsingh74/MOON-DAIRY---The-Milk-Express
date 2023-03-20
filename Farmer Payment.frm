VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FarmerPayment 
   BackColor       =   &H00FF8080&
   Caption         =   " Farmer Payment"
   ClientHeight    =   9705
   ClientLeft      =   4935
   ClientTop       =   2070
   ClientWidth     =   14130
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9705
   ScaleWidth      =   14130
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   4815
      Left            =   11640
      TabIndex        =   46
      Top             =   2400
      Width           =   1815
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   615
         Left            =   240
         TabIndex        =   50
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   615
         Left            =   240
         TabIndex        =   49
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   615
         Left            =   240
         TabIndex        =   48
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View"
         Height          =   615
         Left            =   240
         TabIndex        =   47
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   975
      Left            =   240
      TabIndex        =   39
      Top             =   8400
      Width           =   11175
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9000
         TabIndex        =   45
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   5760
         TabIndex        =   43
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   2400
         TabIndex        =   42
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues"
         Height          =   375
         Left            =   7680
         TabIndex        =   44
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid"
         Height          =   375
         Left            =   4680
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Dues"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Farmer Payment Details"
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   11175
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   8880
         TabIndex        =   34
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   8880
         TabIndex        =   33
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   4080
         TabIndex        =   32
         Top             =   5640
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   4080
         TabIndex        =   31
         Top             =   5160
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   8160
         TabIndex        =   30
         Top             =   6720
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   8160
         TabIndex        =   29
         Top             =   6240
         Width           =   1815
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   8880
         TabIndex        =   26
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   7920
         TabIndex        =   24
         Top             =   2880
         Width           =   855
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   6960
         TabIndex        =   22
         Top             =   2880
         Width           =   855
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   5520
         TabIndex        =   20
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   4080
         TabIndex        =   19
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   3120
         TabIndex        =   17
         Top             =   2880
         Width           =   855
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   2160
         TabIndex        =   14
         Top             =   2880
         Width           =   855
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   600
         TabIndex        =   13
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   8040
         TabIndex        =   10
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   2640
         TabIndex        =   8
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   2640
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   6360
         TabIndex        =   3
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy"
         Format          =   225705987
         CurrentDate     =   44983
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd MMM yyyy"
         Format          =   225705987
         CurrentDate     =   44983
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   300
         Left            =   4920
         TabIndex        =   51
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Evening Amount"
         Height          =   375
         Left            =   5520
         TabIndex        =   38
         Top             =   5640
         Width           =   2775
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Evening Quantity"
         Height          =   375
         Left            =   5520
         TabIndex        =   37
         Top             =   5160
         Width           =   2775
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Total morning Amount"
         Height          =   375
         Left            =   600
         TabIndex        =   36
         Top             =   5640
         Width           =   2775
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Total morning Quantity"
         Height          =   375
         Left            =   600
         TabIndex        =   35
         Top             =   5160
         Width           =   2775
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   375
         Left            =   6240
         TabIndex        =   28
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Qty"
         Height          =   375
         Left            =   6240
         TabIndex        =   27
         Top             =   6360
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   375
         Left            =   8880
         TabIndex        =   25
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Fat"
         Height          =   375
         Left            =   7920
         TabIndex        =   23
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   375
         Left            =   6960
         TabIndex        =   21
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Fat"
         Height          =   375
         Left            =   3240
         TabIndex        =   16
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Evening"
         Height          =   375
         Left            =   5520
         TabIndex        =   12
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Morning"
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         Height          =   375
         Left            =   6480
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Famer Name"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Farmer ID"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   300
         Left            =   5880
         TabIndex        =   4
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Range"
         Height          =   375
         Left            =   5280
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "FARMER PAYMENT"
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
      Left            =   4440
      TabIndex        =   52
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "FarmerPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Single
Private Sub Combo1_Click()
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
Label23.Caption = Combo1.Text
sql = "select f_name from farmer_entry where f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text1.Text = r.Fields(0)
sql = "select ph_no from farmer_entry where f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)
sql = "select dues from farmer_dues where f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text9.Text = r.Fields(0)
sql = "select curr_date,qty,fat,amount from morning_coll where curr_date between '" + Format(DTPicker1.Value, "dd mmm yyyy") + "' and '" + Format(DTPicker2.Value, "dd mmm yyyy") + "' and f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
List1.AddItem r.Fields(0)
List2.AddItem r.Fields(1)
List3.AddItem r.Fields(2)
List4.AddItem r.Fields(3)
r.MoveNext
Loop
sql = "select curr_date,qty,fat,amount from evening_coll where curr_date between '" + Format(DTPicker1.Value, "dd mmm yyyy") + "' and '" + Format(DTPicker2.Value, "dd mmm yyyy") + "' and f_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
List5.AddItem r.Fields(0)
List6.AddItem r.Fields(1)
List7.AddItem r.Fields(2)
List8.AddItem r.Fields(3)
r.MoveNext
Loop

For i = 0 To List6.ListCount Step 1
total = total + Val(List6.List(i))
Text7.Text = total
Next
For i = 0 To List2.ListCount Step 1
Sum = Sum + Val(List2.List(i))
Text5.Text = Sum
Next
For i = 0 To List8.ListCount Step 1
add = add + Val(List8.List(i))
Text8.Text = add
Next

For i = 0 To List4.ListCount Step 1
Addition = Addition + Val(List4.List(i))
Text6.Text = Addition
Next
Text4.Text = Val(Text6.Text) + Val(Text8.Text)
Text3.Text = Val(Text5.Text) + Val(Text7.Text)
Text10.SetFocus

If Text3.Text = "0" Then
MsgBox "No Record found"
End If

Command3.Enabled = True
End Sub

Private Sub Command1_Click()
sql = "select f_id from farmer_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Combo1.Enabled = True
Command4.Enabled = True
End Sub

Private Sub Command2_Click()
Unload Me
home.Show
End Sub

Private Sub Command3_Click()
If Label23.Caption = "" Or Text10.Text = "" Then
MsgBox "Farmer Paid And ID required"
Else
sql = "insert into Farmer_payment values('" + Format(DTPicker1.Value, "dd mmm yyyy") + "','" + Format(DTPicker2.Value, "dd mmm yyyy") + "','" + Label23.Caption + "'," + Text5.Text + "," + Text7.Text + "," + Text3.Text + "," + Text4.Text + "," + Text9.Text + "," + Text10.Text + "," + Text11.Text + ")"
Set r = c.Execute(sql)
sql = " update farmer_dues set dues=" + Text11.Text + " where f_id='" + Label23.Caption + "'"
Set r = c.Execute(sql)
MsgBox "Record Saved"
Combo1.SetFocus
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
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
End If
End Sub

Private Sub Command4_Click()
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
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
End Sub
Private Sub Form_Load()
conn
'Text9.Text = "0"
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Combo1.Enabled = False
Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text11.Locked = True
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Command3.SetFocus
End If
End Sub

Private Sub Text10_LostFocus()
Text11.Text = Val(Text4.Text) + Val(Text9.Text) - Val(Text10.Text)
End Sub
