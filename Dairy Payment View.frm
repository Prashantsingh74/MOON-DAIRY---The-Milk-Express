VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form DairyPaymentView 
   BackColor       =   &H00FF8080&
   Caption         =   "Dairy Payment View"
   ClientHeight    =   9720
   ClientLeft      =   3960
   ClientTop       =   1395
   ClientWidth     =   15780
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
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   15780
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   540
      Left            =   14760
      TabIndex        =   7
      Top             =   7800
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   3015
      Left            =   5160
      TabIndex        =   2
      Top             =   960
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   480
         Left            =   2280
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Submit"
         Height          =   615
         Left            =   2280
         TabIndex        =   5
         Top             =   2040
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         Caption         =   "By Bill No"
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All Payment"
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Dairy Payment View.frx":0000
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   4320
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   7435
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   29
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
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "BILL_NO"
         Caption         =   "BILL NO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "START_DATE"
         Caption         =   "START DATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "END_DATE"
         Caption         =   "END DATE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "D_ID"
         Caption         =   "DAIRY ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "MORNING_QTY"
         Caption         =   "MORNING QUANTITY"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "EVENING_QTY"
         Caption         =   "EVENING QUANTITY"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "TOTAL_QTY"
         Caption         =   "TOTAL QUANTITY"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "TOTAL_AMT"
         Caption         =   "TOTAL AMOUNT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "TOTAL_DUES"
         Caption         =   "TOTAL DUES"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "PAYMENT"
         Caption         =   "PAYMENT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "DUES_LEFT"
         Caption         =   "DUES LEFT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2220.094
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1860.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2039.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   3075.024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2940.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2520
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2399.811
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1874.835
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1695.118
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   10320
      Top             =   5400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1085
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
      RecordSource    =   "select * from dairy_payment where 1=2"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Dairy Payment View"
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
      Left            =   6120
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "DairyPaymentView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value = True Then
Adodc1.RecordSource = "select * from dairy_payment "
Adodc1.Refresh
End If
If Option2.Value = True Then
Text1.Visible = True
If Text1.Text = "" Then
MsgBox "Enter Bill No"
Text1.SetFocus
Else
Adodc1.RecordSource = "select * from dairy_payment where bill_no=" + Text1.Text + ""
Adodc1.Refresh
End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
home.Show
End Sub

Private Sub Form_Load()
Text1.Visible = False
End Sub

Private Sub Option1_Click()
Text1.Text = ""
Text1.Visible = False
Adodc1.RecordSource = "select * from dairy_payment  where 1>2"
Adodc1.Refresh
End Sub
Private Sub Option2_Click()
Text1.Text = ""
Text1.Visible = True
Adodc1.RecordSource = "select * from dairy_payment  where 1>2"
Adodc1.Refresh
End Sub
