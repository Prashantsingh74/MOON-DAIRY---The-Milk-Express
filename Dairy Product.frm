VERSION 5.00
Begin VB.Form DairyProduct 
   BackColor       =   &H00FF8080&
   Caption         =   "Dairy Product"
   ClientHeight    =   8715
   ClientLeft      =   5895
   ClientTop       =   2070
   ClientWidth     =   12270
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
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   12270
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   3615
      Left            =   9720
      TabIndex        =   23
      Top             =   2760
      Width           =   1695
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton clear 
         Caption         =   "Clear"
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton addnew 
         Caption         =   "Add New"
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dairy Product Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   7935
      Begin VB.TextBox Text5 
         Height          =   420
         Left            =   2280
         TabIndex        =   22
         Top             =   960
         Width           =   2055
      End
      Begin VB.ListBox List4 
         Height          =   3660
         Left            =   5160
         TabIndex        =   19
         Top             =   3120
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   420
         Left            =   5160
         TabIndex        =   17
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   420
         Left            =   3120
         TabIndex        =   15
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ListBox List3 
         Height          =   3660
         Left            =   3120
         TabIndex        =   13
         Top             =   3120
         Width           =   1815
      End
      Begin VB.ListBox List2 
         Height          =   3660
         Left            =   1560
         TabIndex        =   10
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton removelist 
         Caption         =   "Remove"
         Height          =   735
         Left            =   6600
         TabIndex        =   9
         Top             =   6000
         Width           =   1095
      End
      Begin VB.CommandButton addlist 
         Caption         =   "Add"
         Height          =   615
         Left            =   6600
         TabIndex        =   8
         Top             =   3120
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   3660
         Left            =   360
         TabIndex        =   7
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   420
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   420
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   420
         Left            =   1560
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Dairy Name"
         Height          =   375
         Left            =   600
         TabIndex        =   21
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial no."
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   375
         Left            =   5280
         TabIndex        =   18
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Product name"
         Height          =   375
         Left            =   3120
         TabIndex        =   12
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Id"
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Product id"
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dairy ID"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Serial no."
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "DAIRY PRODUCT"
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
      Left            =   3480
      TabIndex        =   28
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "DairyProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim k As Integer
Private Sub Combo1_Click()
Text4.Enabled = True
sql = "select pr_name from Product_Entry where pr_id='" + Combo1.Text + "'"
Set r = c.Execute(sql)
Text3.Text = r.Fields(0)
Text3.Enabled = True
Text3.Locked = True
Text4.SetFocus
End Sub
Private Sub addlist_Click()
removelist.Enabled = True
save.Enabled = True
clear.Enabled = True

List1.AddItem Text1.Text
List2.AddItem Combo1.Text
List3.AddItem Text3.Text
List4.AddItem Text4.Text
If Text1.Text = "" Then
sql = "select count(serial_no) from dairy_pr "
Set r = c.Execute(sql)
Text1.Text = r.Fields(0) + 1
Else
Text1.Text = Text1.Text + 1
End If
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
addnew.Enabled = True
addlist.Enabled = False
End Sub

Private Sub removelist_Click()
List1.RemoveItem ListIndex
List2.RemoveItem ListIndex
List3.RemoveItem ListIndex
List4.RemoveItem ListIndex
End Sub

Private Sub addnew_Click()
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Text1.Locked = True
Text2.Locked = True
If Text1.Text = "" Then
sql = "select count(serial_no) from dairy_pr "
Set r = c.Execute(sql)
Text1.Text = r.Fields(0) + 1
Else
Text1.Text = Text1.Text + 1
End If
sql = "select Dairy_Id from Dairy_Entry"
Set r = c.Execute(sql)
Text2.Text = r.Fields(0)
sql = "select Dairy_nm from Dairy_Entry"
Set r = c.Execute(sql)
Text5.Text = r.Fields(0)
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
removelist.Enabled = False

End Sub

Private Sub save_Click()
For k = 0 To List1.ListCount - 1
sql = "insert into dairy_pr(serial_no,dairy_id,pr_id,rate) values(" + List1.List(k) + ",'" + Text2.Text + "','" + List2.List(k) + "'," + List4.List(k) + ")"
Set r = c.Execute(sql)
Next
MsgBox "record saved"
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear

End Sub



Private Sub clear_Click()
Text1.Text = ""
Text2.Text = ""
Combo1.Text = ""
Text3.Text = ""
Text4.Text = ""
List1.clear
List2.clear
List3.clear
List4.clear
End Sub

Private Sub Exit_Click()
Unload Me
home.Show
End Sub

Private Sub Form_Load()
addnew.TabIndex = 0
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Combo1.Enabled = False
addlist.Enabled = False
removelist.Enabled = False
save.Enabled = False
clear.Enabled = False
conn
sql = "select pr_id from Product_Entry"
Set r = c.Execute(sql)
Do While Not r.EOF
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
Text1.Locked = True
Text2.TabIndex = 0
End Sub



Private Sub text4_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
addlist.Enabled = True
addlist.SetFocus
End If
End Sub

Private Sub Text4_LostFocus()
addlist.Enabled = True
addlist.SetFocus
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Not IsNumeric(Text4.Text) Then
cancle = True
MsgBox "Enter Numeric Value"
Text4.Text = ""
Text4.SetFocus
Else
cancle = False
End If
End Sub
