VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PurchaseOrder 
   BackColor       =   &H00FF8080&
   Caption         =   "Product Order"
   ClientHeight    =   8745
   ClientLeft      =   4275
   ClientTop       =   2400
   ClientWidth     =   14430
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
   ScaleHeight     =   8745
   ScaleWidth      =   14430
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   3375
      Left            =   12720
      TabIndex        =   45
      Top             =   3240
      Width           =   1575
      Begin VB.CommandButton new 
         Caption         =   "New"
         Height          =   615
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         Height          =   615
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton clear 
         Caption         =   "Clear"
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         Height          =   615
         Left            =   120
         TabIndex        =   46
         Top             =   2640
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Product order"
      Height          =   4935
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   12255
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   9720
         TabIndex        =   40
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9720
         TabIndex        =   39
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9720
         TabIndex        =   37
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   7920
         TabIndex        =   36
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   5880
         TabIndex        =   35
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton remove 
         Caption         =   "Remove"
         Height          =   495
         Left            =   11040
         TabIndex        =   31
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton add 
         Caption         =   "Add "
         Height          =   495
         Left            =   11160
         TabIndex        =   30
         Top             =   840
         Width           =   975
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   8160
         TabIndex        =   29
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   6720
         TabIndex        =   28
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   5160
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   3600
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   1560
         TabIndex        =   25
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   8160
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   6720
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   5160
         TabIndex        =   20
         Text            =   "Combo2"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   3600
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   1560
         TabIndex        =   18
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   120
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   300
         Left            =   9720
         TabIndex        =   38
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dues"
         Height          =   300
         Left            =   9720
         TabIndex        =   34
         Top             =   3360
         Width           =   570
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance"
         Height          =   300
         Left            =   7920
         TabIndex        =   33
         Top             =   3360
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   300
         Left            =   5880
         TabIndex        =   32
         Top             =   3360
         Width           =   1425
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   300
         Left            =   8160
         TabIndex        =   22
         Top             =   360
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   300
         Left            =   6720
         TabIndex        =   11
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit "
         Height          =   300
         Left            =   5160
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Weight"
         Height          =   300
         Left            =   3600
         TabIndex        =   9
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   300
         Left            =   1560
         TabIndex        =   8
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Product Details"
      Height          =   2055
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   12255
      Begin VB.TextBox Text13 
         Height          =   420
         Left            =   6000
         TabIndex        =   44
         Text            =   "Text13"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Height          =   420
         Left            =   6000
         TabIndex        =   42
         Text            =   "Text9"
         Top             =   960
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   225050625
         CurrentDate     =   44957
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   225050625
         CurrentDate     =   44957
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   6000
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   9960
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   2040
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No"
         Height          =   300
         Left            =   4560
         TabIndex        =   43
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dairy Name"
         Height          =   300
         Left            =   4560
         TabIndex        =   41
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dairy ID"
         Height          =   300
         Left            =   4560
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Date"
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Centre ID"
         Height          =   300
         Left            =   8640
         TabIndex        =   3
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         Height          =   300
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE ORDER"
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
      Left            =   5160
      TabIndex        =   50
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "PurchaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Single
Dim a As Integer, B As Integer, cc As Integer, d As Single
Private Sub Combo1_Click()
sql = "select pr_name from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text4.Text = r.Fields(0)
sql = "select weight from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text5.Text = r.Fields(0)
sql = "select unit from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Combo2.Text = r.Fields(0)
sql = "select MRP from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text7.Text = r.Fields(0)
Text4.Locked = True
Text5.Locked = True
Combo2.Locked = True
Text7.Locked = True
Text6.SetFocus
End Sub
Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
End Sub


Private Sub new_Click()
conn
sql = "select count(order_no)from PurchaseORDER_details"
Set r = c.Execute(sql)
Text1.Text = r.Fields(0) + 1
sql = "select centre_id from collection_centre"
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)

sql = "select dairy_id from dairy_entry"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
sql = "select dairy_nm from dairy_entry"
Set r = c.Execute(sql)
Text9.Text = r.Fields(0)
sql = "select phone_no from dairy_entry"
Set r = c.Execute(sql)
Text13.Text = r.Fields(0)
Combo1.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo2.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
DTPicker1.Enabled = True
DTPicker2.Enabled = True
save.Enabled = True
clear.Enabled = True
add.Enabled = True
remove.Enabled = True
Text1.Locked = True
Text2.Enabled = True
Text3.Enabled = True
Combo1.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Combo2.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
Text13.Enabled = True
Text2.Locked = True
Text3.Locked = True
End Sub

Private Sub save_Click()
conn
sql = "insert into purchaseORDER_details values (" + Text1.Text + ",'" + Text2.Text + "','" + Format(DTPicker1.Value, "dd mmm yyyy") + "','" + Format(DTPicker2.Value, "dd mmm yyyy") + "','" + Text3.Text + "'," + Text8.Text + "," + Text10.Text + "," + Text11.Text + ")"
Set r = c.Execute(sql)
For k = 0 To List1.ListCount - 1
sql = "insert into PurchaseOrder_pr values(" + Text1.Text + ",'" + List1.List(k) + "','" + List5.List(k) + "','" + List6.List(k) + "','" + List7.List(k) + "')"
Set r = c.Execute(sql)
Next
MsgBox "record saved"
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
List1.clear
List2.clear
List3.clear
List4.clear
List5.clear
List6.clear
List7.clear
End Sub

Private Sub clear_Click()
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
End Sub

Private Sub Exit_Click()
Unload Me
home.Show
End Sub

Private Sub add_Click()
If Combo1.Text = "" Or Text6.Text = "" Or Text12.Text = "" Then
MsgBox "All Fields Required"
Else
List1.AddItem Combo1.Text
List2.AddItem Text4.Text
List3.AddItem Text5.Text
List4.AddItem Combo2.Text
List5.AddItem Text6.Text
List6.AddItem Text7.Text
List7.AddItem Text12.Text
For i = 0 To List7.ListCount Step 1
total = total + Val(List7.List(i))
Text8.Text = total
Next
Text7.Text = ""
Text6.Text = ""
Text5.Text = ""
Text4.Text = ""
Text12.Text = ""
Combo1.Text = ""
Combo2.Text = ""
End If
End Sub

Private Sub remove_Click()
List1.RemoveItem ListIndex
List2.RemoveItem ListIndex
List3.RemoveItem ListIndex
List4.RemoveItem ListIndex
List5.RemoveItem ListIndex
List6.RemoveItem ListIndex
List7.RemoveItem ListIndex
For i = 0 To List7.ListCount Step 1
total = total + Val(List7.List(i))
Text8.Text = total
Next
End Sub

Private Sub Form_Load()
conn
sql = "select pr_id from product_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0) + ""
r.MoveNext
Loop
save.Enabled = False
clear.Enabled = False
add.Enabled = False
remove.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
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
Text11.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
Text11.SetFocus
End If
End Sub

Private Sub Text10_LostFocus()
Text11.Text = Text8.Text - Text10.Text
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
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


Private Sub text4_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
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

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then

Text7.SetFocus
End If

End Sub

Private Sub Text6_LostFocus()
a = Val(Text6.Text)
B = Val(Text7.Text)
cc = a * B
Text12.Text = cc
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
If KeyAscii = 13 Then
add.SetFocus
End If
End Sub

Private Sub Text7_LostFocus()
a = Text6.Text
B = Text7.Text
cc = a * B
Text12.Text = cc
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER NUMBER ONLY"
End If
End Sub


