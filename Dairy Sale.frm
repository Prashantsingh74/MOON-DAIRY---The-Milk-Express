VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DairySale 
   BackColor       =   &H00FF8080&
   Caption         =   "Dairy Sale"
   ClientHeight    =   8730
   ClientLeft      =   4605
   ClientTop       =   2070
   ClientWidth     =   15210
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
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   15210
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   1935
      Left            =   4440
      TabIndex        =   31
      Top             =   960
      Width           =   4095
      Begin VB.OptionButton Option2 
         Caption         =   "Evening"
         Height          =   495
         Left            =   2160
         TabIndex        =   33
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Morning"
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         Height          =   375
         Left            =   1560
         TabIndex        =   36
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   3735
      Left            =   12720
      TabIndex        =   14
      Top             =   3360
      Width           =   1935
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton clear 
         Caption         =   "Clear"
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton addnew 
         Caption         =   "Add New"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dairy Sale Product"
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   12255
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   10320
         TabIndex        =   38
         Top             =   4200
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   480
         Left            =   8160
         TabIndex        =   37
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   5760
         TabIndex        =   30
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         Height          =   480
         Left            =   4320
         TabIndex        =   28
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Height          =   480
         Left            =   5520
         TabIndex        =   27
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         Height          =   480
         Left            =   5520
         TabIndex        =   26
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Height          =   480
         Left            =   5520
         TabIndex        =   25
         Top             =   960
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   10200
         TabIndex        =   20
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   139853825
         CurrentDate     =   44985
      End
      Begin VB.TextBox Text6 
         Height          =   480
         Left            =   7320
         TabIndex        =   13
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   9240
         TabIndex        =   11
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   480
         Left            =   2280
         TabIndex        =   10
         Top             =   4200
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   480
         Left            =   360
         TabIndex        =   9
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   5520
         TabIndex        =   8
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   480
         Left            =   1440
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Bal"
         Height          =   375
         Left            =   4320
         TabIndex        =   35
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   375
         Left            =   10200
         TabIndex        =   29
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   375
         Left            =   5760
         TabIndex        =   24
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   375
         Left            =   2280
         TabIndex        =   23
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Center Name"
         Height          =   375
         Left            =   3360
         TabIndex        =   22
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Center Id"
         Height          =   375
         Left            =   3360
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   375
         Left            =   9360
         TabIndex        =   19
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   375
         Left            =   9240
         TabIndex        =   5
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fat"
         Height          =   375
         Left            =   8280
         TabIndex        =   4
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Id"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dairy Id"
         Height          =   375
         Left            =   3360
         TabIndex        =   2
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Label17"
      Height          =   495
      Left            =   7800
      TabIndex        =   39
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "DAIRY SALE"
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
      Left            =   5880
      TabIndex        =   34
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "DairySale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
sql = "select pr_name from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
sql = "select balance from stock where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text10.Text = r.Fields(0)
sql = "select unit from product_entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text6.Text = r.Fields(0)
Text3.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
'Text4.Enabled = True
'Text5.Enabled = True
Text6.Enabled = True
Combo2.Enabled = True
Text3.Locked = True
Text10.Locked = True
Text6.Locked = True
End Sub
Private Sub Combo2_Click()
sql = "select rate from fatlist where fat=" + Combo2.Text + ""
Set r = c.Execute(sql)
Text5.Text = r.Fields(0)
Text4.Enabled = True
Text5.Enabled = True
Text4.Locked = True
Text5.Locked = True
Text4 = Val(Text5.Text) * Val(Text11.Text)
save.SetFocus
End Sub

Private Sub addnew_Click()
Text1.Enabled = True
Text2.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
DTPicker1.Enabled = True
Combo1.Enabled = True
save.Enabled = True
clear.Enabled = True
Text1.Locked = True
Text2.Locked = True
Text7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text2.Locked = True
sql = "select dairy_id from dairy_entry"
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)
sql = "select dairy_nm from dairy_entry"
Set r = c.Execute(sql)
Text9.Text = r.Fields(0)

If Option1.Value = True Then
sql = "select count(bill_no) from morning_dairysale"
Set r = c.Execute(sql)
Text1.Text = r.Fields(0) + 1
Text1.Locked = True
End If

If Option2.Value = True Then
sql = "select count(bill_no) from evening_dairysale"
Set r = c.Execute(sql)
Text1.Text = r.Fields(0) + 1
Text1.Locked = True
End If
sql = "select centre_id from collection_centre"
Set r = c.Execute(sql)
Text7.Text = r.Fields(0)
sql = "select centre_name from collection_centre"
Set r = c.Execute(sql)
Text8.Text = r.Fields(0)
End Sub



Private Sub save_Click()
sql = "update stock set balance=" + Label17.Caption + " where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
MsgBox "stock updated"
If Option1.Value = True Then
sql = "insert into morning_dairysale values(" + Text1.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "', '" + Text7.Text + "','" + Text2.Text + "','" + Combo1.Text + "'," + Text11.Text + ",'" + Text6.Text + "'," + Combo2.Text + "," + Text5.Text + "," + Text4.Text + ")"
Set r = c.Execute(sql)
MsgBox "Morning record saved"
End If
If Option2.Value = True Then
sql = "insert into evening_dairysale values(" + Text1.Text + ",'" + Format(DTPicker1.Value, "dd mmm yyyy") + "', '" + Text7.Text + "','" + Text2.Text + "','" + Combo1.Text + "'," + Text11.Text + ",'" + Text6.Text + "'," + Combo2.Text + "," + Text5.Text + "," + Text4.Text + ")"
Set r = c.Execute(sql)
MsgBox "Evening record saved"
End If
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo2.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
End Sub

Private Sub clear_Click()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo2.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""

End Sub

Private Sub Exit_Click()
Unload Me
home.Show
End Sub

Private Sub Form_Load()
conn
Text1.Enabled = False
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
addnew.Enabled = False
Combo2.Enabled = False
Combo1.Enabled = False
save.Enabled = False
clear.Enabled = False
DTPicker1.Enabled = False
sql = "select pr_id from product_entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select fat from fatlist"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo2.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub

Private Sub Option1_Click()
Text1.Enabled = False
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
addnew.Enabled = False
Combo2.Enabled = False
Combo1.Enabled = False
save.Enabled = False
clear.Enabled = False
DTPicker1.Enabled = False
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo2.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
addnew.Enabled = True
addnew.SetFocus
End Sub
Private Sub Option2_Click()
Text1.Enabled = False
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
addnew.Enabled = False
Combo2.Enabled = False
Combo1.Enabled = False
save.Enabled = False
clear.Enabled = False
DTPicker1.Enabled = False
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Combo2.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
addnew.Enabled = True
addnew.SetFocus
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If (KeyAscii > 47 And KeyAscii < 58) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox ("Enter Number Only")
End If
If KeyAscii = 13 Then
Combo2.SetFocus
End If
End Sub

Private Sub Text11_LostFocus()
Label17.Caption = Val(Text10.Text) - Val(Text11.Text)
End Sub
