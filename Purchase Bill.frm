VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PurchaseBill 
   BackColor       =   &H00FF8080&
   Caption         =   " Purchase Bill"
   ClientHeight    =   10185
   ClientLeft      =   2985
   ClientTop       =   1740
   ClientWidth     =   16290
   ClipControls    =   0   'False
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
   ScaleHeight     =   10185
   ScaleWidth      =   16290
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   3135
      Left            =   13800
      TabIndex        =   58
      Top             =   720
      Width           =   1695
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   240
         TabIndex        =   62
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton clear 
         Caption         =   "Clear"
         Height          =   495
         Left            =   240
         TabIndex        =   61
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   60
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton addnew 
         Caption         =   "New"
         Height          =   495
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   15135
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   8760
         TabIndex        =   68
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   140640257
         CurrentDate     =   44985
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   7080
         TabIndex        =   67
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   140640257
         CurrentDate     =   44985
      End
      Begin VB.TextBox Text17 
         Height          =   420
         Left            =   7560
         TabIndex        =   65
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   10560
         TabIndex        =   57
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   4800
         TabIndex        =   55
         Top             =   3960
         Width           =   1215
      End
      Begin VB.ListBox List12 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   10440
         TabIndex        =   54
         Top             =   1440
         Width           =   975
      End
      Begin VB.ListBox List11 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   8760
         TabIndex        =   53
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ListBox List10 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   7080
         TabIndex        =   52
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   10440
         TabIndex        =   51
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   6120
         TabIndex        =   42
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9000
         TabIndex        =   41
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   3240
         TabIndex        =   38
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   1920
         TabIndex        =   36
         Top             =   3960
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   240
         TabIndex        =   34
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton remove 
         Caption         =   "Remove"
         Height          =   420
         Left            =   13920
         TabIndex        =   32
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton Add 
         Caption         =   "Add"
         Height          =   495
         Left            =   14040
         TabIndex        =   31
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   12600
         TabIndex        =   30
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   420
         Left            =   4800
         TabIndex        =   29
         Text            =   "Unit"
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   240
         TabIndex        =   28
         Text            =   "Product ID"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   11520
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   5880
         TabIndex        =   26
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   3480
         TabIndex        =   25
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   1800
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   12600
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   11520
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   5880
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   4800
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   3480
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   1800
         TabIndex        =   10
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ListBox List8 
         Height          =   1260
         Left            =   12720
         TabIndex        =   47
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
         Height          =   375
         Left            =   4800
         TabIndex        =   70
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Amount"
         Height          =   375
         Left            =   7560
         TabIndex        =   64
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Dues Left"
         Height          =   375
         Left            =   10560
         TabIndex        =   56
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch"
         Height          =   375
         Left            =   10440
         TabIndex        =   50
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Exp. Date"
         Height          =   375
         Left            =   8760
         TabIndex        =   49
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "MFD"
         Height          =   375
         Left            =   7200
         TabIndex        =   48
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Amount"
         Height          =   375
         Left            =   3240
         TabIndex        =   40
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount%"
         Height          =   375
         Left            =   1920
         TabIndex        =   39
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "PreviousDues"
         Height          =   375
         Left            =   6120
         TabIndex        =   37
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid"
         Height          =   375
         Left            =   9480
         TabIndex        =   35
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   375
         Left            =   12600
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   375
         Left            =   11520
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   375
         Left            =   6000
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Id"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   3255
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   10095
      Begin VB.TextBox Text18 
         Height          =   420
         Left            =   4080
         TabIndex        =   73
         Top             =   840
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   7320
         TabIndex        =   66
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   140574721
         CurrentDate     =   44985
      End
      Begin VB.ComboBox Combo3 
         Height          =   420
         Left            =   4080
         TabIndex        =   43
         Text            =   "Order No"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   4080
         TabIndex        =   22
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   4080
         TabIndex        =   20
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   1200
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Dairy Id"
         Height          =   375
         Left            =   2760
         TabIndex        =   72
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   300
         Left            =   6600
         TabIndex        =   23
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         Height          =   300
         Left            =   2760
         TabIndex        =   21
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         Height          =   300
         Left            =   2760
         TabIndex        =   19
         Top             =   2640
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Dairy Name"
         Height          =   300
         Left            =   2640
         TabIndex        =   18
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE BILL"
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
      Left            =   6480
      TabIndex        =   74
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label29 
      Caption         =   "Label29"
      Height          =   375
      Left            =   8880
      TabIndex        =   71
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Advance"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   13320
      TabIndex        =   69
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Label26"
      Height          =   375
      Left            =   13320
      TabIndex        =   63
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
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
      Left            =   13200
      TabIndex        =   46
      Top             =   7680
      Width           =   1410
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label21"
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
      Left            =   13320
      TabIndex        =   45
      Top             =   7680
      Width           =   1410
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label22"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   13200
      TabIndex        =   44
      Top             =   7680
      Width           =   1410
   End
End
Attribute VB_Name = "PurchaseBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Single, X As Single, Y As Single
Dim p As Integer
Private Sub add_Click()

'Label21.Caption = Combo1.Text
Label22.Caption = Text7.Text
Label18.Caption = Val(Label20.Caption) + Val(Label22.Caption)

If Combo1.Text = "" Or Text7.Text = "" Or Text15.Text = "" Or Text8.Text = "" Then
MsgBox "Enter All Data"
Else
List1.AddItem Combo1.Text
List2.AddItem Text5.Text
List3.AddItem Text6.Text
List4.AddItem Combo2.Text
List5.AddItem Text7.Text
List6.AddItem Text8.Text
List7.AddItem Text9.Text
List8.AddItem Label18.Caption
List10.AddItem DTPicker2.Value
List11.AddItem DTPicker3.Value
List12.AddItem Text15.Text
End If
For i = 0 To List7.ListCount Step 1
total = total + Val(List7.List(i))
Text10.Text = total
Next
p = Combo1.ListIndex
If p > -1 Then
Combo1.RemoveItem p
End If
Combo1.Text = ""
Text5.Text = ""
Combo2.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text15.Text = ""
remove.Enabled = True
End Sub

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
sql = "select Qty From PurchaseOrder_pr where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text7.Text = r.Fields(0)
sql = "select MRP from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text8.Text = r.Fields(0)
Text5.Locked = True
Text6.Locked = True
Combo2.Locked = True
'Text8.Locked = True
Text7.SetFocus
'sql = "select * from STOCK where pr_id='" + Combo1.Text + "'"
'Set r = c.Execute(sql)
'Label21.Caption = r.Fields(0)
'Label20.Caption = r.Fields(4)
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
End Sub


Private Sub Combo3_Click()
sql = "select pr_id from PURchaseorder_Pr where order_no=" + Combo3.Text + ""
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select advance from purchaseORDER_details where order_no=" + Combo3.Text + ""
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)
Combo1.Enabled = True
addnew.Enabled = False
'Sql = "select dues from purchaseORDER_details where order_no=" + Combo3.Text + ""
'Set R = C.Execute(Sql)
'Text14.Text = R.Fields(0)
'Sql = "select tot_amt from purchaseORDER_details where order_no=" + Combo3.Text + ""
'Set R = C.Execute(Sql)
'Text10.Text = R.Fields(0)
End Sub

Private Sub remove_Click()
Combo1.AddItem List1.Text
List1.RemoveItem ListIndex
List2.RemoveItem ListIndex
List3.RemoveItem ListIndex
List4.RemoveItem ListIndex
List5.RemoveItem ListIndex
List6.RemoveItem ListIndex
List7.RemoveItem ListIndex
List10.RemoveItem ListIndex
List11.RemoveItem ListIndex
List12.RemoveItem ListIndex
For i = 0 To List7.ListCount Step 1
total = total + Val(List7.List(i))
Text10.Text = total
Next
End Sub

Private Sub addnew_Click()
Combo3.clear
conn
sql = "select count(Bill_no)from purchasebill_details"
Set r = c.Execute(sql)
Text1.Text = r.Fields(0) + 1
sql = "select dairy_id from dairy_entry"
Set r = c.Execute(sql)
Text18.Text = r.Fields(0)
sql = "select dairy_nm from dairy_entry"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
sql = "select phone_no from dairy_entry"
Set r = c.Execute(sql)
Text4.Text = r.Fields(0)
sql = "select centre_id from collection_centre"
Set r = c.Execute(sql)
Label29.Caption = r.Fields(0)
Text14.Text = "0"
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Combo2.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text14.Enabled = True
Text15.Enabled = True
Text16.Enabled = True
save.Enabled = True
clear.Enabled = True
Text9.Locked = True
Text10.Locked = True
Text2.Locked = True
Text12.Locked = True
Text14.Locked = True
sql = "select order_no from PurchaseORDER_DETAILS"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo3.AddItem r.Fields(0)
r.MoveNext
Loop
Combo3.SetFocus
Text16.Locked = True
End Sub
Private Sub save_Click()
conn
If Text11.Text = "" Or Text13.Text = "" Then
MsgBox "Enter All Fields"
Else
sql = "insert into purchasebill_details values (" + Text1.Text + "," + Combo3.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "','" + Text18.Text + "'," + Text10.Text + "," + Text11.Text + "," + Text17.Text + "," + Text13.Text + "," + Text16.Text + ")"
Set r = c.Execute(sql)
For k = 0 To List1.ListCount - 1
sql = "insert into purchasebill_pr values(" + Text1.Text + ",'" + List1.List(k) + "','" + List5.List(k) + "','" + Format(List10.List(k), "dd mmm yyyy") + "','" + Format(List11.List(k), "dd mmm yyyy") + "','" + List6.List(k) + "','" + List7.List(k) + "')"
Set r = c.Execute(sql)
Next
For k = 0 To List1.ListCount - 1
balance = balance + Val(List5.List(k))
Label18.Caption = balance
sql = "UPDATE stock SET balance =" + List8.List(k) + " WHERE pr_id = '" + List1.List(k) + "'"
'sql = "update stock set BALANCE=" + balance + Label20.Caption + ",AMOUNT=" + Label19.Caption + " WHERE PR_ID='" + Label21.Caption + "'"
Set r = c.Execute(sql)
'MsgBox " 1 data saved"
sql = "UPDATE stock SET MFD ='" + Format(List10.List(k), "dd mmm yyyy") + "' WHERE pr_id = '" + List1.List(k) + "'"
Set r = c.Execute(sql)
sql = "UPDATE stock SET EXP ='" + Format(List11.List(k), "dd mmm yyyy") + "' WHERE pr_id = '" + List1.List(k) + "'"
Set r = c.Execute(sql)

Next
sql = "update CENTRE_dues set total_dues=" + Text16.Text + "" ' where centre_id=" + lable29.Caption + ""
Set r = c.Execute(sql)

For i = 0 To List7.ListCount - 1
tot = tot + Val(List7.List(i))
Next
Text9.Text = tot
'sql = "UPDATE STOCK SET BALANCE=" + Label20.Caption + ", AMOUNT=" + Label19.Caption + " WHERE PR_ID='" + Label21.Caption + "'"
'Set r = c.Execute(sql)
MsgBox "Record Saved"
Combo3.Text = ""
Text2.Text = ""
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
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Text14.Text = ""
Text16.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
List10.clear
List11.clear
List12.clear
addnew.Enabled = True
End If
End Sub
Private Sub clear_Click()
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
Text13.Text = ""
Text12.Text = ""
Text14.Text = ""
Combo1.Text = ""
Combo2.Text = ""
End Sub
Private Sub Exit_Click()
Unload Me
home.Show
End Sub

Private Sub Form_Load()
conn

List8.Visible = True
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo1.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Combo2.Enabled = False
Text7.Enabled = False
Text8.Enabled = False
Text9.Enabled = False
Text10.Enabled = False
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
Text14.Enabled = False
Text15.Enabled = False
Text16.Enabled = False
save.Enabled = False
clear.Enabled = False
remove.Enabled = False
add.Enabled = False
'List8.Visible = False
End Sub



Private Sub Text11_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Text13.SetFocus
End If
End Sub

Private Sub Text11_LostFocus()
Label26.Caption = (Val(Text10.Text) * Val(Text11.Text)) / 100
Text12.Text = Val(Text10.Text) - Val(Label26.Caption)
Text17.Text = Text12.Text - (Text2.Text - Text14.Text)
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
save.SetFocus
End If
End Sub

Private Sub Text13_LostFocus()
Text16.Text = (Val(Text17.Text) - Val(Text13.Text))
Label18.Caption = Val(Label22.Caption) + Val(Label20.Caption)
'Label19.Caption = Val(Label20.Caption) * Val(Label19.Caption)
'Text14.Text = Text12.Text - Text13.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Text8.SetFocus
End If
End Sub

Private Sub text3_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub text4_keypress(KeyAscii As Integer)
Text4.MaxLength = 10
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Combo1.SetFocus
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
'x = Text7.Text
'y = Text8.Text
If KeyAscii = 13 Then
'Text9.Text = x * y
DTPicker2.SetFocus
End If
End Sub
Private Sub Text7_LostFocus()
a = Val(Text7.Text)
B = Val(Text8.Text)
cc = a * B
Text9.Text = cc
'Label20.Caption = Val(Label18.Caption) + Val(Text7.Text)
'Label19.Caption = a * b
Label22.Caption = Text7.Text
End Sub
'Set DataEnvironment1 = Nothing

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
add.Enabled = True
If KeyAscii = 13 Then
add.SetFocus
End If

End Sub

