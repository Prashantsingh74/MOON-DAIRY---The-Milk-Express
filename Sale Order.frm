VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CustomerOrder 
   BackColor       =   &H00FF8080&
   Caption         =   "Customer Order"
   ClientHeight    =   8730
   ClientLeft      =   3960
   ClientTop       =   2400
   ClientWidth     =   14835
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   14835
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   4335
      Left            =   12960
      TabIndex        =   43
      Top             =   3840
      Width           =   1695
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         Height          =   540
         Left            =   240
         TabIndex        =   47
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   615
         Left            =   240
         TabIndex        =   46
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   615
         Left            =   240
         TabIndex        =   45
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New"
         Height          =   615
         Left            =   240
         TabIndex        =   44
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      Height          =   1830
      Left            =   5880
      TabIndex        =   30
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Order Details"
      Height          =   4335
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   12375
      Begin VB.CommandButton Command6 
         Caption         =   "Remove"
         Height          =   615
         Left            =   11040
         TabIndex        =   41
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add"
         Height          =   615
         Left            =   11040
         TabIndex        =   40
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   4320
         TabIndex        =   39
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   2640
         TabIndex        =   38
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   600
         TabIndex        =   35
         Top             =   3720
         Width           =   1575
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   9720
         TabIndex        =   33
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   8280
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   7080
         TabIndex        =   31
         Top             =   1320
         Width           =   975
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   4200
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   2280
         TabIndex        =   28
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   480
         TabIndex        =   27
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9720
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   8280
         TabIndex        =   25
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   7080
         TabIndex        =   24
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   4200
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   420
         Left            =   5640
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   480
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   2280
         TabIndex        =   18
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
         Height          =   255
         Left            =   2640
         TabIndex        =   37
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues"
         Height          =   255
         Left            =   4320
         TabIndex        =   36
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   375
         Left            =   600
         TabIndex        =   34
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9720
         TabIndex        =   20
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   375
         Left            =   8280
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   375
         Left            =   7080
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   375
         Left            =   5640
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "weight"
         Height          =   375
         Left            =   4200
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Product name"
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Id"
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Customer Details"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   12375
      Begin VB.ComboBox Combo3 
         Height          =   420
         Left            =   3360
         TabIndex        =   42
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9600
         TabIndex        =   12
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   3360
         TabIndex        =   11
         Top             =   1680
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   9600
         TabIndex        =   10
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   225640449
         CurrentDate     =   44958
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   9600
         TabIndex        =   9
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   225640449
         CurrentDate     =   44958
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone no"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
         Height          =   375
         Left            =   7320
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
         Height          =   375
         Left            =   7320
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ORDER"
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
      Left            =   5280
      TabIndex        =   48
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "CustomerOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Single, Y As Single, z As Single
Dim k As Single

Private Sub Combo1_Click()
sql = "select pr_name from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text5.Text = r.Fields(0)
sql = "select weight from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text6.Text = r.Fields(0)
sql = "select unit from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Combo2.Text = r.Fields(0)
sql = "select MRP from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text8.Text = r.Fields(0)
'Combo1.Enabled = True
'Text4.Enabled = True
Text5.Enabled = True
Combo2.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
'Text10.Enabled = True
'Text11.Enabled = True
'Text12.Enabled = True
'Text13.Enabled = True
'Text14.Enabled = True
Text5.Locked = True
Text6.Locked = True
Combo2.Locked = True
Text8.Locked = True
Text9.Locked = True
Text7.SetFocus
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub


Private Sub Combo3_Click()
sql = "select name from customer_entry where cust_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
sql = "select ph_no from customer_entry where cust_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text4.Text = r.Fields(0)
Text3.Enabled = True
Text4.Enabled = True
Combo1.Enabled = True
Text3.Locked = True
Text4.Locked = True
Combo1.SetFocus
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo1.SetFocus
End If
End Sub

Private Sub Command1_Click()
sql = "select count(order_no) from SALEORDER_details "
Set r = c.Execute(sql)
Text1.Text = r.Fields(0) + 1
sql = "select cust_id from customer_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo3.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select pr_id from product_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
DTPicker1.Enabled = True
DTPicker2.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Text1.Enabled = True
Combo3.Enabled = True
'Text3.Enabled = True
'Combo1.Enabled = True
'Text4.Enabled = True
'Text5.Enabled = True
'Combo2.Enabled = True
'Text6.Enabled = True
'Text7.Enabled = True
'Text8.Enabled = True
'Text9.Enabled = True
'Text10.Enabled = True
'Text11.Enabled = True
'Text12.Enabled = True
'Text13.Enabled = True
'Text14.Enabled = True
Combo3.Text = ""
Text3.Text = ""
Combo1.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo2.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
'Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
'Text14.Text = ""
Combo3.SetFocus
End Sub

Private Sub Command2_Click()
conn
If Text12.Text = "" Then
MsgBox "Enter All Fields"
Else
sql = "insert into SALEORDER_details values(" + Text1.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "','" + Format(DTPicker2.Value, "dd mmm yyyy") + "','" + Combo3.Text + "'," + Text10.Text + "," + Text12.Text + "," + Text13.Text + ")"
Set r = c.Execute(sql)
For k = 0 To List1.ListCount - 1
sql = "insert into saleorder_pr values(" + Text1.Text + ",'" + List1.List(k) + "','" + List5.List(k) + "','" + List6.List(k) + "','" + List7.List(k) + "')"
Set r = c.Execute(sql)
Next
MsgBox "record saved"
Text1.Text = ""
Combo3.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo2.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
'Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
'Text14.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
Command1.SetFocus
End If
End Sub
Private Sub Command3_Click()
Text1.Text = ""
Combo3.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo2.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
Command1.SetFocus
End Sub

Private Sub Command4_Click()
Unload Me
home.Show
End Sub

Private Sub Command5_Click()
If Combo1.Text = "" Or Text7.Text = "" Then
MsgBox " Enter All Fields"
Else
List1.AddItem Combo1.Text
List2.AddItem Text5.Text
List3.AddItem Text6.Text
List4.AddItem Combo2.Text
List5.AddItem Text7.Text
List6.AddItem Text8.Text
List7.AddItem Text9.Text
For i = 0 To List7.ListCount Step 1
total = total + Val(List7.List(i))
Text10.Text = total
Next
Text10.Enabled = True
'Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
'Text14.Enabled = True
Text10.Locked = True
Text13.Locked = True
'Text14.Locked = True
Combo1.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo2.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
End If
End Sub
Private Sub Command6_Click()
List1.RemoveItem ListIndex
List2.RemoveItem ListIndex
List3.RemoveItem ListIndex
List4.RemoveItem ListIndex
List5.RemoveItem ListIndex
List6.RemoveItem ListIndex
List7.RemoveItem ListIndex
For i = 0 To List7.ListCount Step 1
total = total + Val(List7.List(i))
Text10.Text = total
Next
End Sub

Private Sub Form_Load()
conn
'Combo1.AddItem "ghee"
Combo2.AddItem "kg"
Combo2.AddItem "lit"
Combo2.AddItem "gm"
Combo2.AddItem "ml"
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Text1.Enabled = False
Combo3.Enabled = False
Text3.Enabled = False
Combo1.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Combo2.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
'Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
'Text14.Enabled = False
End Sub

'Private Sub Text10_KeyPress(KeyAscii As Integer)
'If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
'KeyAscii = KeyAscii
'Else
'KeyAscii = 0
'MsgBox "ENTER NUMBER ONLY"
'End If
'End Sub

'Private Sub Text11_KeyPress(KeyAscii As Integer)
'If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
'KeyAscii = KeyAscii
'Else
'KeyAscii = 0
'MsgBox "ENTER NUMBER ONLY"
'End If
'If KeyAscii = 13 Then
'Text12.SetFocus
'End If

'End Sub

'Private Sub Text11_LostFocus()
'Text14.Text = Val(Text10.Text) - Val(Text10.Text) * Val(Text11.Text) / 100

'End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Text13.Text = Val(Text10.Text) - Val(Text12.Text)
Command2.SetFocus
End If
End Sub



Private Sub Text12_LostFocus()
Text13.Text = Val(Text10.Text) - Val(Text12.Text)
End Sub

'Private Sub Text13_KeyPress(KeyAscii As Integer)
'If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
'KeyAscii = KeyAscii
'Else
'KeyAscii = 0
'MsgBox "ENTER NUMBER ONLY"
'End If
'End Sub




'Private Sub Text14_KeyPress(KeyAscii As Integer)
'If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
'KeyAscii = KeyAscii
'Else
'KeyAscii = 0
'MsgBox "ENTER NUMBER ONLY"
'End If
'End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub text3_keypress(KeyAscii As Integer)
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER ALPHABET"
End If
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text4_LostFocus()
If Len(Text4.Text) <> 10 Then
MsgBox "Enter 10 Digit"
Text4.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Combo2.SetFocus
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Command5.SetFocus
End If
End Sub

Private Sub Text7_LostFocus()
X = Val(Text7.Text)
Y = Val(Text8.Text)
z = X * Y
Text9.Text = z
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Command5.SetFocus
End If
End Sub

Private Sub Text8_LostFocus()
X = Val(Text7.Text)
Y = Val(Text8.Text)
z = X * Y
Text9.Text = z
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
End Sub
