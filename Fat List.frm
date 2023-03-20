VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FatList 
   BackColor       =   &H00FF8080&
   Caption         =   "Fat List"
   ClientHeight    =   6045
   ClientLeft      =   6870
   ClientTop       =   3060
   ClientWidth     =   10560
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
   ScaleHeight     =   6045
   ScaleWidth      =   10560
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   2775
      Left            =   7920
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
      Begin VB.CommandButton exit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton save 
         Caption         =   "Save"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton update 
         Caption         =   "Update"
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fat Details"
      Height          =   4695
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   480
         Left            =   2040
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   480
         Left            =   3720
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2295
         Left            =   360
         TabIndex        =   1
         Top             =   1920
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   4048
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   27
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
         Caption         =   "View "
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "FAT"
            Caption         =   "FAT"
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
            DataField       =   "RATE"
            Caption         =   "RATE"
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
               ColumnWidth     =   1635.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1904.882
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   1560
         Top             =   3600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
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
         Connect         =   "Provider=MSDAORA.1;User ID=moondairy/moon;Persist Security Info=False"
         OLEDBString     =   "Provider=MSDAORA.1;User ID=moondairy/moon;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from fatlist"
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fat%"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FAT LIST"
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
      Left            =   4440
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "FatList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Single
Dim i As Single



Private Sub Combo1_Click()
n = Combo1.Text
Select Case n
Case 3
Text2.Text = 29.75
Case 3.1
Text2.Text = 30.11
Case 3.2
Text2.Text = 30.46
Case 3.3
Text2.Text = 30.82
Case 3.4
Text2.Text = 31.18
Case 3.5
Text2.Text = 31.54
Case 3.6
Text2.Text = 31.89
Case 3.7
Text2.Text = 32.25
Case 3.8
Text2.Text = 32.61
Case 3.9
Text2.Text = 32.96
Case 4
Text2.Text = 33.32
Case 4.1
Text2.Text = 33.68
Case 4.2
Text2.Text = 34.03
Case 4.3
Text2.Text = 34.39
Case 4.4
Text2.Text = 34.75
Case 4.5
Text2.Text = 35.6
Case 4.6
Text2.Text = 35.96
Case 4.7
Text2.Text = 36.32
Case 4.8
Text2.Text = 36.68
Case 4.9
Text2.Text = 37.04
Case 5
Text2.Text = 37.41
Case 5.1
Text2.Text = 37.77
Case 5.2
Text2.Text = 38.13
Case 5.3
Text2.Text = 38.49
Case 5.4
Text2.Text = 38.85
Case 5.5
Text2.Text = 40.4
Case 5.6
Text2.Text = 40.77
Case 5.7
Text2.Text = 41.13
Case 5.8
Text2.Text = 41.5
Case 5.9
Text2.Text = 41.86
Case 6
Text2.Text = 40.74
Case 6.1
Text2.Text = 41.42
Case 6.2
Text2.Text = 42.1
Case 6.3
Text2.Text = 42.78
Case 6.4
Text2.Text = 43.46
Case 6.5
Text2.Text = 44.14
Case 6.6
Text2.Text = 44.81
Case 6.7
Text2.Text = 45.49
Case 6.8
Text2.Text = 46.17
Case 6.9
Text2.Text = 46.85
Case 7
Text2.Text = 47.53
Case 7.1
Text2.Text = 48.21
Case 7.2
Text2.Text = 48.89
Case 7.3
Text2.Text = 49.57
Case 7.4
Text2.Text = 50.25
Case 7.5
Text2.Text = 50.93
Case 7.6
Text2.Text = 51.6
Case 7.7
Text2.Text = 52.28
Case 7.8
Text2.Text = 52.96
Case 7.9
Text2.Text = 53.64
Case 8
Text2.Text = 54.32
Case 8.1
Text2.Text = 55
Case 8.2
Text2.Text = 55.68
Case 8.3
Text2.Text = 56.36
Case 8.4
Text2.Text = 57.04
Case 8.5
Text2.Text = 57.72
Case 8.6
Text2.Text = 58.39
Case 8.7
Text2.Text = 59.07
Case 8.8
Text2.Text = 59.75
Case 8.9
Text2.Text = 60.43
Case 9
Text2.Text = 61.11
Case 9.1
Text2.Text = 61.79
Case 9.2
Text2.Text = 62.47
Case 9.3
Text2.Text = 63.15
Case 9.4
Text2.Text = 63.83
Case 9.5
Text2.Text = 64.51
Case 9.6
Text2.Text = 65.18
Case 9.7
Text2.Text = 65.86
Case 9.8
Text2.Text = 66.54
Case 9.9
Text2.Text = 67.22
Case 10
Text2.Text = 67.9
End Select

End Sub

Private Sub save_Click()
sql = "insert into fatlist values(" + Combo1.Text + "," + Text2.Text + ")"
Set r = c.Execute(sql)
MsgBox "record saved"
Adodc1.Refresh
End Sub

Private Sub update_Click()
Text2.Locked = False
sql = " update fatlist set rate=" + Text2.Text + " where fat=" + Combo1.Text + ""
Set r = c.Execute(sql)
MsgBox "record updated"
Adodc1.Refresh
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text2.Locked = False
'Combo1.Locked = True
Text1.Locked = True
update.Enabled = False
End Sub

Private Sub Exit_Click()
Unload Me
home.Show
End Sub

Private Sub Form_Load()
conn
For i = 3 To 11 Step 0.1
Combo1.AddItem Round(i, 1)
If i >= 10 Then
Exit For
End If
Next
Text1.Text = "1 liter"
'Text1.Locked = True
Text2.Locked = True
End Sub



Private Sub Text2_Change()
update.Enabled = True
End Sub
