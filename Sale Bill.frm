VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SaleBill 
   BackColor       =   &H00FF8080&
   Caption         =   "Sale Bill"
   ClientHeight    =   10215
   ClientLeft      =   3795
   ClientTop       =   2565
   ClientWidth     =   15945
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10215
   ScaleWidth      =   15945
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Height          =   3855
      Left            =   12720
      TabIndex        =   50
      Top             =   3840
      Width           =   1935
      Begin VB.CommandButton Command6 
         Caption         =   "CLEAR"
         Height          =   495
         Left            =   240
         TabIndex        =   54
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "EXIT"
         Height          =   420
         Left            =   240
         TabIndex        =   53
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SAVE"
         Height          =   540
         Left            =   240
         TabIndex        =   52
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ADD NEW"
         Height          =   540
         Left            =   240
         TabIndex        =   51
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Choose Billing Type"
      Height          =   975
      Left            =   4200
      TabIndex        =   47
      Top             =   960
      Width           =   5175
      Begin VB.OptionButton Option2 
         Caption         =   "Bill By Order"
         Height          =   300
         Left            =   2880
         TabIndex        =   49
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Direct Bill"
         Height          =   375
         Left            =   480
         TabIndex        =   48
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Product Details"
      Height          =   5295
      Left            =   360
      TabIndex        =   9
      Top             =   4800
      Width           =   12255
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   6360
         TabIndex        =   74
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9480
         TabIndex        =   69
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   6360
         TabIndex        =   68
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   5880
         TabIndex        =   67
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   1920
         TabIndex        =   66
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9480
         TabIndex        =   42
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   7800
         TabIndex        =   41
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   4800
         TabIndex        =   40
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   2040
         TabIndex        =   39
         Top             =   3960
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   240
         TabIndex        =   38
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "REMOVE"
         Enabled         =   0   'False
         Height          =   495
         Left            =   10800
         TabIndex        =   32
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         Height          =   495
         Left            =   11040
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   9600
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   8520
         TabIndex        =   29
         Top             =   1320
         Width           =   975
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   7200
         TabIndex        =   28
         Top             =   1320
         Width           =   975
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   5160
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   3840
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   2040
         TabIndex        =   25
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   2130
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9600
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   8520
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   7200
         TabIndex        =   21
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   420
         Left            =   5160
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   3840
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   420
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   2040
         TabIndex        =   17
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Dues"
         Height          =   375
         Left            =   6360
         TabIndex        =   73
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label30 
         Caption         =   "Label30"
         Height          =   375
         Left            =   2280
         TabIndex        =   71
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Order Quantity"
         Height          =   375
         Left            =   7320
         TabIndex        =   65
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
         Height          =   375
         Left            =   6360
         TabIndex        =   64
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Previous Dues"
         Height          =   375
         Left            =   3960
         TabIndex        =   63
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
         Height          =   375
         Left            =   360
         TabIndex        =   62
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues Left"
         Height          =   255
         Left            =   9480
         TabIndex        =   37
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid"
         Height          =   255
         Left            =   7800
         TabIndex        =   36
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Amount"
         Height          =   255
         Left            =   4800
         TabIndex        =   35
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount %"
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   375
         Left            =   9600
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   255
         Left            =   8520
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   375
         Left            =   7200
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   375
         Left            =   5160
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Product name"
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "Label29"
         Height          =   375
         Left            =   2640
         TabIndex        =   70
         Top             =   1920
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bill Details"
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   12255
      Begin VB.ComboBox Combo4 
         Height          =   420
         Left            =   5160
         TabIndex        =   45
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   9000
         TabIndex        =   44
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   140443651
         UpDown          =   -1  'True
         CurrentDate     =   44967
      End
      Begin VB.ComboBox Combo3 
         Height          =   420
         Left            =   5160
         TabIndex        =   43
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Billno 
         Appearance      =   0  'Flat
         DataField       =   "BILL_NO"
         DataSource      =   "Adodc1"
         Height          =   390
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   5160
         TabIndex        =   2
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   5160
         TabIndex        =   1
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No."
         Height          =   255
         Left            =   3000
         TabIndex        =   46
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Id"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Mob. No."
         Height          =   255
         Left            =   3000
         TabIndex        =   6
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Date"
         Height          =   255
         Left            =   7800
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No."
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.ListBox List8 
      Height          =   960
      Left            =   10200
      TabIndex        =   56
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "SALE BILL"
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
      Height          =   495
      Left            =   6000
      TabIndex        =   75
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label31 
      Caption         =   "Label31"
      Height          =   375
      Left            =   10800
      TabIndex        =   72
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label24 
      Caption         =   "Label24"
      Height          =   255
      Left            =   10680
      TabIndex        =   61
      Top             =   9240
      Width           =   1215
   End
   Begin VB.Label Label23 
      Caption         =   "Label23"
      Height          =   255
      Left            =   10680
      TabIndex        =   60
      Top             =   8880
      Width           =   1095
   End
   Begin VB.Label Label22 
      Caption         =   "Label22"
      Height          =   375
      Left            =   10680
      TabIndex        =   59
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "Label21"
      Height          =   255
      Left            =   10680
      TabIndex        =   58
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "Label20"
      Height          =   255
      Left            =   10680
      TabIndex        =   57
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   375
      Left            =   10440
      TabIndex        =   55
      Top             =   2040
      Width           =   1575
   End
End
Attribute VB_Name = "Salebill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer

Private Sub Combo1_Click()
sql = "select * from STOCK where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
'Label21.Caption = R.Fields(0)
Label23.Caption = r.Fields(4)
Text6.Text = 1
If Option1.Value = True Then
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Combo2.Enabled = True
Command1.Enabled = True
sql = "select pr_name from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text4.Text = r.Fields(0)
sql = "select weight from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text5.Text = r.Fields(0)
sql = "select unit from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Combo2.Text = r.Fields(0)
sql = "select mrp from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text7.Text = r.Fields(0)
Text6.SetFocus
sql = "select balance from stock where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text15.Text = r.Fields(0)
End If
If Option2.Value = True Then
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Combo2.Enabled = True
Command1.Enabled = True
sql = "select pr_name from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text4.Text = r.Fields(0)
sql = "select weight from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text5.Text = r.Fields(0)
sql = "select unit from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Combo2.Text = r.Fields(0)
sql = "select qty from saleorder_pr where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text6.Text = r.Fields(0)
sql = "select rate from saleorder_pr where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text7.Text = r.Fields(0)
Text6.SetFocus
sql = "select balance from stock where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text15.Text = r.Fields(0)

End If
End Sub

Private Sub Combo1_LostFocus()
'Text6.SetFocus
End Sub

Private Sub Combo3_Click()
sql = "select name from customer_Entry where cust_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)
sql = "select ph_no from customer_Entry where cust_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
Text1.Text = 0
Label31.Caption = Combo3.Text
sql = "select dues from customer_dues where cust_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text14.Text = r.Fields(0)
End Sub

Private Sub Combo4_Click()
Combo1.clear
Label19.Caption = Combo4.Text
sql = "select cust_id from SALEORDER_details where order_no='" + Combo4.Text + "'"
Set r = c.Execute(sql)
Combo3.Text = r.Fields(0)
Label31.Caption = r.Fields(0)
sql = "select name from customer_entry where cust_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)
sql = "select ph_no from customer_entry where cust_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
sql = "select pr_id from saleorder_pr where order_no='" + Combo4.Text + "'"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select paid from saleorder_details where order_no=" + Combo4.Text + ""
Set r = c.Execute(sql)
Text1.Text = r.Fields(0)

End Sub

Private Sub Combo4_LostFocus()
sql = "select dues from customer_dues where cust_id='" + Label31.Caption + "'"
Set r = c.Execute(sql)
Text14.Text = r.Fields(0)
End Sub

Private Sub Command1_Click()
Label24.Caption = Val(Label23.Caption) - Label22.Caption
Label20.Caption = Combo1.Text
Label21.Caption = Text4.Text
'SQL = "select total_dues from customer_dues where cust_id='" + Label31.Caption + "'"
'Set R = C.Execute(SQL)
'Combo3.Text = R.Fields(0)

If Combo1.Text = "" Or Text6.Text = "" Or Text8.Text = "" Or Text7.Text = "" Or Text2.Text = "" Or Combo3.Text = "" Then
MsgBox "All Fields Required"
Command1.SetFocus
Else
List1.AddItem Combo1.Text
List2.AddItem Text4.Text
List3.AddItem Text5.Text
List4.AddItem Combo2.Text
List5.AddItem Text6.Text
List6.AddItem Text7.Text
List7.AddItem Text8.Text
List8.AddItem Val(Label24.Caption)

End If

Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text9.Locked = True
Text11.Locked = True
Text13.Locked = True
Combo1.Text = ""
Combo2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""

Dim tot As Single
For i = 0 To List7.ListCount - 1
tot = tot + Val(List7.List(i))
Next
Text9.Text = tot
Text10.Text = ""
Combo1.SetFocus
End Sub

Private Sub Command2_Click()
If l1 > List1.ListIndex And l1 > List1.ListIndex And l1 > List1.ListIndex And l1 > List1.ListIndex And l1 > List1.ListIndex And l1 > List1.ListIndex And l1 > List1.listindexn Then
Command2.Enabled = True
End If
If List1.ListIndex And List2.ListIndex And List3.ListIndex And List4.ListIndex And List5.ListIndex And List6.ListIndex And List7.ListIndex Then
List1.RemoveItem List1.ListIndex
List2.RemoveItem List2.ListIndex
List3.RemoveItem List3.ListIndex
List4.RemoveItem List4.ListIndex
List5.RemoveItem List5.ListIndex
List6.RemoveItem List6.ListIndex
List7.RemoveItem List7.ListIndex
For i = 0 To List7.ListCount - 1
tot = tot + Val(List7.List(i))
Next
Text9.Text = tot
Command2.Enabled = False
Else
List1.RemoveItem List1.ListIndex
List2.RemoveItem List2.ListIndex
List3.RemoveItem List3.ListIndex
List4.RemoveItem List4.ListIndex
List5.RemoveItem List5.ListIndex
List6.RemoveItem List6.ListIndex
List7.RemoveItem List7.ListIndex
For i = 0 To List7.ListCount - 1
tot = tot + Val(List7.List(i))
Next
Text9.Text = tot
End If
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
If Text10.Text = "" Or Text12.Text = "" Then
MsgBox "Enter All Fields"
'If Option1.Value = True Then
'SQL = "insert into salebill_details values(" + Billno.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "'," + null + ",'" + Combo3.Text + "','" + Text2.Text + "'," + Text3.Text + "," + Text9.Text + "," + Text10.Text + "," + Text11.Text + "," + Text12.Text + "," + Text13.Text + ")"
'Set R = C.Execute(SQL)
'Else
Else
sql = "insert into salebill_details values(" + Billno.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "'," + Combo4.Text + ",'" + Combo3.Text + "'," + Text9.Text + "," + Text10.Text + "," + Text11.Text + "," + Text12.Text + "," + Text13.Text + ")"
Set r = c.Execute(sql)
'End If
For k = 0 To List1.ListCount - 1
sql = "insert into salebill_pr values (" + Billno.Text + ",'" + List1.List(k) + "'," + List5.List(k) + "," + List6.List(k) + "," + List7.List(k) + ")"
Set r = c.Execute(sql)
Next

For k = 0 To List1.ListCount - 1
sql = "UPDATE stock SET balance =" + List8.List(k) + " WHERE pr_id = '" + List1.List(k) + "'"
'SQL = "update stock set BALANCE="+balance + Label20.Caption + ",AMOUNT=" + Label19.Caption + " WHERE PR_ID='" + Label21.Caption + "'"
Set r = c.Execute(sql)
'MsgBox " 1 data saved"
Next

sql = "update customer_dues set dues=" + Text13.Text + " where cust_id='" + Combo3.Text + "'"
Set r = c.Execute(sql)
MsgBox "dues updated"
Billno.Text = ""
Combo4.Text = ""
Combo3.Text = ""
Text2.Text = ""
Text3.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text1.Text = ""
Text14.Text = ""
Text17.Text = ""

End If
End Sub

Private Sub Command4_Click()
Unload Me
home.Show
End Sub

Private Sub Command5_Click()
Combo4.clear
Combo3.clear
'List8.Clear
If Option1.Value = True Then
Combo4.Visible = False
Combo4.Text = "Null"
Label18.Visible = False
Label25.Visible = False
Label28.Visible = False
Text1.Visible = False
Text16.Visible = False
'Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo3.Enabled = True
Combo1.Enabled = True
Combo1.clear
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
End If
Text9.Locked = True
Text11.Locked = True
Text13.Locked = True
Text8.Locked = True

If Option2.Value = True Then
Combo4.Visible = True
Label18.Visible = True
'lable1.Visible = True
Combo3.Visible = True
Label25.Visible = True
Label28.Visible = True
Text1.Visible = True
Text16.Visible = True
'Text1.Enabled = False
Billno.Enabled = True
Combo4.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Combo1.clear

sql = "select order_no from saleorder_details"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo4.AddItem r.Fields(0)
r.MoveNext
Loop

Text9.Locked = True
Text11.Locked = True
Text13.Locked = True
Text8.Locked = True
End If
sql = "select count(bill_no) from salebill_details"
Set r = c.Execute(sql)
Billno = r.Fields(0) + 1
Command6.Enabled = True
End Sub

Private Sub Command6_Click()
Billno.Text = ""
Combo4.Text = ""
Combo3.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List8.clear
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""

End Sub

Private Sub Form_Load()
n = 1
conn
Billno.Locked = True
Billno.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
'Command4.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Combo2.Locked = True
Text7.Locked = True
End Sub




Private Sub List1_Click()
a = List1.ListIndex
If a >= 0 Then

End If
End Sub

Private Sub List7_Click()
a = List1.ListIndex
If a >= 0 Then
Command2.Enabled = True
End If
End Sub

Private Sub Option3_Click()

End Sub

Private Sub Option1_Click()
Command5.Enabled = True
Command5.SetFocus
End Sub

Private Sub Option2_Click()
Label19.Caption = "Null"
Command5.Enabled = True
Command5.SetFocus
End Sub

Private Sub Text1_Change()
'Text1.Text = 0
End Sub

Private Sub Text10_LostFocus()
Label30.Caption = Text14.Text - Text1.Text
Label29.Caption = (Text9.Text * Text10.Text) / 100
Text11.Text = Text9.Text - Val(Label29.Caption)
Text17.Text = Text11.Text + Val(Label30.Caption)
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox ("Enter Number Only")
End If
If KeyAscii = 13 Then
Command3.Enabled = True
Command3.SetFocus
End If
End Sub

Private Sub Text12_LostFocus()
Text13.Text = Val(Text17.Text) - Val(Text12.Text)
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.SetFocus
End If
End Sub

Private Sub Text6_LostFocus()
Label22.Caption = Text6.Text
If Val(Text6.Text) > Val(Text15.Text) Then
MsgBox "Not Enough Stock"
Combo1.SetFocus

Else
Text8.Text = Val(Text6.Text) * Val(Text7.Text)
End If
If Val(Text6.Text) <= 0 Then
MsgBox "Minimum Quantity 1"
Combo1.SetFocus
End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox ("Enter Number Only")
End If
If KeyAscii = 13 Then
Text12.SetFocus
End If
End Sub

