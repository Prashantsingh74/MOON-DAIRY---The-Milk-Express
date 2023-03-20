VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CustomerEntry 
   BackColor       =   &H00FF8080&
   Caption         =   "Customer Entry"
   ClientHeight    =   8295
   ClientLeft      =   4290
   ClientTop       =   2430
   ClientWidth     =   12615
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
   ScaleHeight     =   8295
   ScaleWidth      =   12615
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Customer Entry.frx":0000
      Height          =   2415
      Left            =   720
      TabIndex        =   43
      Top             =   5640
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CUSTOMER DETAILS"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "CUST_ID"
         Caption         =   "CUSTOMER ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NAME"
         Caption         =   "NAME"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "GENDER"
         Caption         =   "GENDER"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ADDRESS"
         Caption         =   "ADDRESS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PH_NO"
         Caption         =   "PHONE NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2069.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1844.787
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2684.977
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2055.118
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8281
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Entry"
      TabPicture(0)   =   "Customer Entry.frx":0015
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Update"
      TabPicture(1)   =   "Customer Entry.frx":0031
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label15"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Text11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Delete"
      TabPicture(2)   =   "Customer Entry.frx":004D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label23"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "Text17"
      Tab(2).Control(3)=   "Command7"
      Tab(2).Control(4)=   "Command8"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command8 
         Caption         =   "Delete"
         Height          =   615
         Left            =   -66960
         TabIndex        =   42
         Top             =   3720
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Search"
         Height          =   615
         Left            =   -66960
         TabIndex        =   41
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   -67080
         TabIndex        =   40
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Customer Delete"
         Height          =   4095
         Left            =   -74640
         TabIndex        =   29
         Top             =   480
         Width           =   6855
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   51
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox Text16 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   38
            Top             =   3480
            Width           =   2055
         End
         Begin RichTextLib.RichTextBox RichTextBox3 
            Height          =   1455
            Left            =   3120
            TabIndex        =   37
            Top             =   1920
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2566
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"Customer Entry.frx":0069
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   36
            Top             =   840
            Width           =   3375
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   35
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No"
            Height          =   300
            Left            =   480
            TabIndex        =   34
            Top             =   3600
            Width           =   1050
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   300
            Left            =   480
            TabIndex        =   33
            Top             =   2160
            Width           =   885
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   300
            Left            =   480
            TabIndex        =   32
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
            Height          =   300
            Left            =   480
            TabIndex        =   31
            Top             =   1080
            Width           =   1725
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer ID"
            Height          =   300
            Left            =   480
            TabIndex        =   30
            Top             =   480
            Width           =   1350
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Update"
         Height          =   615
         Left            =   7920
         TabIndex        =   28
         Top             =   3720
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Search"
         Height          =   615
         Left            =   7920
         TabIndex        =   27
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   7920
         TabIndex        =   25
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Customer Update"
         Height          =   4095
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   6855
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   50
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox Text8 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   24
            Top             =   3480
            Width           =   2055
         End
         Begin RichTextLib.RichTextBox RichTextBox2 
            Height          =   1455
            Left            =   3120
            TabIndex        =   23
            Top             =   1920
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2566
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"Customer Entry.frx":00EB
         End
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   22
            Top             =   840
            Width           =   3375
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   21
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No"
            Height          =   300
            Left            =   480
            TabIndex        =   20
            Top             =   3600
            Width           =   1050
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   300
            Left            =   480
            TabIndex        =   19
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   300
            Left            =   480
            TabIndex        =   18
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
            Height          =   300
            Left            =   480
            TabIndex        =   17
            Top             =   960
            Width           =   1725
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer ID"
            Height          =   300
            Left            =   480
            TabIndex        =   16
            Top             =   480
            Width           =   1350
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3495
         Left            =   -67080
         TabIndex        =   10
         Top             =   960
         Width           =   2175
         Begin VB.CommandButton Command4 
            Caption         =   "Exit"
            Height          =   615
            Left            =   360
            TabIndex        =   14
            Top             =   2760
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Clear"
            Height          =   615
            Left            =   360
            TabIndex        =   13
            Top             =   1920
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Save"
            Height          =   615
            Left            =   360
            TabIndex        =   12
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CommandButton Command1 
            Caption         =   "New"
            Height          =   615
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Customer Entry"
         Height          =   4095
         Left            =   -74640
         TabIndex        =   1
         Top             =   480
         Width           =   6855
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   1455
            Left            =   3120
            TabIndex        =   47
            Top             =   1920
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2566
            _Version        =   393217
            Enabled         =   -1  'True
            Appearance      =   0
            TextRTF         =   $"Customer Entry.frx":016D
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Female"
            Height          =   300
            Left            =   4560
            TabIndex        =   46
            Top             =   1440
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Male"
            Height          =   300
            Left            =   3120
            TabIndex        =   45
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   9
            Top             =   3480
            Width           =   2175
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   8
            Top             =   840
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   420
            Left            =   3120
            TabIndex        =   7
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label24 
            Height          =   375
            Left            =   3600
            TabIndex        =   44
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No"
            Height          =   300
            Left            =   480
            TabIndex        =   6
            Top             =   3600
            Width           =   1050
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   300
            Left            =   480
            TabIndex        =   5
            Top             =   1920
            Width           =   885
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
            Height          =   300
            Left            =   480
            TabIndex        =   4
            Top             =   1440
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer Name"
            Height          =   300
            Left            =   480
            TabIndex        =   3
            Top             =   840
            Width           =   1725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer ID"
            Height          =   300
            Left            =   480
            TabIndex        =   2
            Top             =   360
            Width           =   1350
         End
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Search by Customer Id"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -67320
         TabIndex        =   39
         Top             =   720
         Width           =   2430
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Search By Customer ID"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7560
         TabIndex        =   26
         Top             =   840
         Width           =   2505
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   9000
      Top             =   6960
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=moon/admin;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=moon/admin;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from customer_entry "
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ENTRY"
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
      TabIndex        =   49
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      Height          =   375
      Left            =   9360
      TabIndex        =   48
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "CustomerEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = "C00"
sql = "select count(cust_id) from customer_entry"
Set r = c.Execute(sql)
Text1.Text = a & r.Fields(0) + 1
Text1.Enabled = True
Text2.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
RichTextBox1.Enabled = True
Text4.Enabled = True
End Sub

Private Sub Command2_Click()

sql = "insert into customer_entry values ('" + Text1.Text + "','" + Text2.Text + "','" + Label24.Caption + "','" + RichTextBox1.Text + "'," + Text4.Text + ")"
Set r = c.Execute(sql)
'SQL = "insert into customer_dues values ('" + Text1.Text + "','" + Text2.Text + "'," + Label5.Caption + "," + Label7.Caption + "," + Label12.Caption + "," + lablel14.Caption + ")"
'Set R = C.Execute(SQL)
sql = "insert into customer_dues values ('" + Text1.Text + "','" + Text2.Text + "'," + Label5.Caption + ")"
Set r = c.Execute(sql)
MsgBox "Record saved"
Adodc1.Refresh
Text2.Text = ""
Label24.Caption = ""
RichTextBox1.Text = ""
Text4.Text = ""
Option1.Value = False
Option2.Value = False
End Sub

Private Sub Command3_Click()
Text2.Text = ""
RichTextBox1.Text = ""
Text4.Text = ""

End Sub

Private Sub Command4_Click()
Unload Me
home.Show
End Sub

Private Sub Command5_Click()
On Error GoTo ABC
sql = "select * from Customer_entry where cust_id='" + Text11.Text + "'"
Set r = c.Execute(sql)
Text6.Text = r.Fields(0)
Text7.Text = r.Fields(1)
Text3.Text = r.Fields(2)
RichTextBox2.Text = r.Fields(3)
Text8.Text = r.Fields(4)
Text6.Enabled = True
Text7.Enabled = True
'Option3.Enabled = true
'Option4.Enabled = true
RichTextBox2.Enabled = True
Text8.Enabled = True
Exit Sub
ABC:
MsgBox "NO DATA FOUND"
End Sub

Private Sub Command6_Click()
If Text11.Text = "" Then
MsgBox "Enter Value"
Else
sql = "update customer_entry set name='" + Text7.Text + "',gender='" + Text3.Text + "',address='" + RichTextBox2.Text + "',ph_no=" + Text8.Text + "where cust_id='" + Text6.Text + "'"
Set r = c.Execute(sql)
MsgBox "Record Updated"
Adodc1.Refresh
Text6.Locked = True
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text3.Text = ""
RichTextBox2.Text = ""
End If
End Sub

Private Sub Command7_Click()
On Error GoTo ABC
sql = "select * from customer_entry WHERE cust_id='" + Text17.Text + "'"
Set r = c.Execute(sql)
Text12.Text = r.Fields(0)
Text13.Text = r.Fields(1)
Text5.Text = r.Fields(2)
RichTextBox3.Text = r.Fields(3)
Text16.Text = r.Fields(4)
Exit Sub
ABC:
MsgBox "NO DATA FOUND"
End Sub

Private Sub Command8_Click()
On Error GoTo ABC
sql = "DELETE FROM CUSTomer_ENTRY WHERE CUST_ID='" + Text17.Text + "'"
Set r = c.Execute(sql)
MsgBox "RECORD DELETED"
Adodc1.Refresh
Text12.Text = ""
Text13.Text = ""
RichTextBox3.Text = ""
Text16.Text = ""
Exit Sub
ABC:
MsgBox "Data Can't be deleted"
End Sub

Private Sub Form_Load()
conn
Adodc1.Refresh
Text1.Enabled = False
Text2.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
RichTextBox1.Enabled = False
Text4.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
'Option3.Enabled = False
'Option4.Enabled = False
RichTextBox2.Enabled = False
Text8.Enabled = False
Text12.Enabled = False
Text13.Enabled = False
RichTextBox3.Enabled = False
Text16.Enabled = False
Text5.Enabled = False

End Sub



Private Sub Option1_Click()
If Option1.Value = True Then
Label24.Caption = Option1.Caption
Option2.Value = False
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Label24.Caption = Option2.Caption
Option1.Value = False
End If
End Sub

Private Sub Option6_Click()

End Sub

Private Sub RichTextBox1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub RichTextBox1_LostFocus()
RichTextBox1.Text = UCase(RichTextBox1.Text)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command5.SetFocus
End If
End Sub

Private Sub Text11_LostFocus()
Text11.Text = UCase(Text11.Text)
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command7.SetFocus
End If

End Sub

Private Sub Text17_LostFocus()
Text17.Text = UCase(Text17.Text)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 96 And KeyAscii < 123) Or (KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii = 32) Or (KeyAscii = 13) Or (KeyAscii = 8) Then
KeyAscii = KeyAscii
Else
KeyAscii = 0
MsgBox "ENTER ALPHABET"
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
Command2.SetFocus
End If
End Sub
